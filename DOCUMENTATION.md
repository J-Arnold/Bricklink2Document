# Bricklink2Document — Documentation

## Table of Contents

1. [Overview](#overview)
2. [Architecture](#architecture)
3. [Data Model](#data-model)
4. [XML Parsing](#xml-parsing)
5. [Image & Description Fetching](#image--description-fetching)
6. [GUI Components](#gui-components)
7. [Export Formats](#export-formats)
8. [Configuration & Persistence](#configuration--persistence)
9. [Caching](#caching)
10. [Keyboard Shortcuts](#keyboard-shortcuts)
11. [Dependency Reference](#dependency-reference)

---

## Overview

Bricklink2Document is a single-file Python desktop application (`Bricklink2Document.py`). It reads Bricklink wanted-list XML files, enriches each item with a live thumbnail and description fetched from the Bricklink website, and lets the user export the result as a formatted Excel spreadsheet, PDF, or Word document.

---

## Architecture

```
Bricklink2Document.py
│
├── Data layer
│   ├── BricklinkItem         — domain model (dataclass)
│   ├── parse_xml()           — XML → list[BricklinkItem]
│   ├── BRICKLINK_COLORS      — color ID → name lookup table
│   ├── ITEM_TYPES / CONDITIONS — code → label maps
│   └── _item_values()        — BricklinkItem → column dict for exporters
│
├── Background worker
│   └── ImageDownloadThread   — QThread: downloads images + descriptions
│
├── GUI
│   ├── MainWindow            — main application window
│   ├── DraggableTable        — QTableWidget with drag-drop row reordering
│   ├── ColumnConfigDialog    — show/hide columns dialog
│   ├── ReadOnlyDelegate      — copy-friendly cell editor
│   └── NumericItem           — QTableWidgetItem with numeric sort
│
├── Exporters
│   ├── export_excel()        — openpyxl .xlsx export
│   ├── export_pdf()          — reportlab A4 PDF export
│   └── export_word()         — python-docx .docx export
│
└── Persistence
    ├── load_config() / save_config()           — column visibility and order
    ├── _load_desc_cache() / _save_desc_cache() — description cache
    └── Image cache                             — per-session temp directory
```

The application runs entirely on the main thread except for the background `ImageDownloadThread`, which communicates back to the GUI via Qt signals.

---

## Data Model

### `BricklinkItem` (dataclass)

| Field | Type | Source |
|---|---|---|
| `item_type` | `str` | XML `<ITEMTYPE>` — `P`, `S`, `M`, `G`, … |
| `item_id` | `str` | XML `<ITEMID>` |
| `color_id` | `int` | XML `<COLOR>` |
| `max_price` | `float` | XML `<MAXPRICE>` — `-1` means "Any" |
| `min_qty` | `int` | XML `<MINQTY>` |
| `condition` | `str` | XML `<CONDITION>` — `N`, `U`, `X` |
| `notify` | `str` | XML `<NOTIFY>` |
| `source_file` | `str` | Stem of the XML filename |
| `description` | `str` | Fetched from Bricklink catalog page |
| `image_data` | `bytes \| None` | Downloaded from Bricklink CDN |

Computed properties:

| Property | Returns |
|---|---|
| `color_name` | Human-readable color from `BRICKLINK_COLORS` |
| `type_label` | Human-readable type from `ITEM_TYPES` |
| `condition_label` | Human-readable condition from `CONDITIONS` |
| `image_urls` | Ordered list of CDN URLs to try for this item (item ID is URL-encoded) |
| `catalog_url` | Bricklink catalog page URL — used to scrape the description (item ID is URL-encoded) |
| `price_label` | Formatted price string, e.g. `$0.0500` or `Any` |

---

## XML Parsing

`parse_xml(path)` uses Python's built-in `xml.etree.ElementTree`. It expects the standard Bricklink wanted-list format:

```xml
<INVENTORY>
  <ITEM>
    <ITEMTYPE>P</ITEMTYPE>
    <ITEMID>3001</ITEMID>
    <COLOR>11</COLOR>
    <MAXPRICE>0.05</MAXPRICE>
    <MINQTY>4</MINQTY>
    <CONDITION>N</CONDITION>
    <NOTIFY>N</NOTIFY>
  </ITEM>
  ...
</INVENTORY>
```

Each `<ITEM>` is parsed inside its own `try/except (ValueError, TypeError)` block. A malformed numeric field (e.g. non-numeric `<COLOR>`) causes that individual item to be skipped silently; the rest of the file continues loading.

Multiple XML files can be loaded simultaneously; they are combined into a single list and distinguished by the `source_file` column.

---

## Image & Description Fetching

### `ImageDownloadThread`

Runs in a separate `QThread` to keep the GUI responsive. For each item it:

1. **Image** — checks the session image cache directory first. On a miss, tries each URL in `item.image_urls` in order until it gets a valid `image/*` response no larger than 5 MB. Saves the raw bytes to the cache directory and emits `image_ready(index, bytes)`.

2. **Description** — checks `descriptions.json` first. On a miss, fetches the Bricklink catalog page and extracts the item name via `_parse_description()`. The result is written back to `descriptions.json` immediately and emitted via `desc_ready(index, str)`.

### Image URL strategy

`item_id` is percent-encoded via `urllib.parse.quote` before being inserted into any URL.

| Item type | URLs tried |
|---|---|
| Set (`S`) | `ItemImage/ST/0/{id}.t2.png` → `ItemImage/SN/0/{id}-1.jpg` |
| Minifig (`M`) | `ItemImage/ST/0/{id}.t2.png` → `ItemImage/MN/0/{id}.png` |
| Gear (`G`) | `ItemImage/GN/0/{id}.png` |
| Part (`P`) | `ItemImage/PN/{colorId}/{id}.png` |

### Image safety limits

- `PILImage.MAX_IMAGE_PIXELS = 10_000_000` is set at module level to guard against decompression-bomb images.
- HTTP responses larger than **5 MB** are rejected before the bytes are stored.
- The cache filename uses a sanitized copy of `item_id` (`re.sub(r"[^\w\-]", "_", ...)`) to prevent path traversal.

### Description extraction (`_parse_description`)

Tries two strategies on the HTML of the catalog page:

1. `meta name="description"` content — pattern `ItemName: LEGO {name}, ItemType: …`
2. `<title>` tag — takes the text before ` : ` and strips the `| BrickLink` suffix.

---

## GUI Components

### `MainWindow`

The top-level `QMainWindow`. Owns:

- The toolbar with **Open XML**, **Export Excel**, **Export PDF**, **Export Word**, **Columns** buttons and a filename label.
- A progress bar shown during image downloading.
- The `DraggableTable`.
- A status bar for download progress messages.

On close, the background download thread is stopped and the session image cache directory is deleted.

### `DraggableTable`

Extends `QTableWidget` with proper whole-row drag-and-drop. Configured with `setDragEnabled(True)`, `setAcceptDrops(True)`, `setDragDropMode(InternalMove)`, and `setDragDropOverwriteMode(False)`.

The default Qt `dropEvent` moves only cell items, not cell widgets (images). The override saves all cell items with `takeItem`, removes the source row, inserts a new row at the target position, and restores the items. After the move it emits `rowOrderChanged` so the main window refreshes image widgets.

### `ColumnConfigDialog`

A `QDialog` with one `QCheckBox` per column. Opens from the **Columns …** button. On accept, updates `_col_cfg` and saves it to `config.json`.

### `ReadOnlyDelegate`

A `QStyledItemDelegate` that opens a read-only `QLineEdit` on double-click, allowing the user to select and copy cell text without accidentally editing it.

### `NumericItem`

A `QTableWidgetItem` that overrides `__lt__` to compare by `float(text)` instead of string, enabling correct numeric sorting on columns like `#`, `Color ID`, `Qty`.

---

## Export Formats

All three exporters share the same column metadata:

- **Visible columns** — determined by `col_cfg` (show/hide) and `col_order` (sequence).
- **Column widths** — defined in `_EXCEL_COL_WIDTHS` (characters) and `_PDF_COL_WIDTHS_MM` (millimetres, also reused for Word).
- **Alignment** — `Description`, `Source File`, and `Color Name` are left-aligned; all other columns are centred.
- **Image thumbnails** — generated with Pillow (`_to_thumbnail`), RGBA mode, `LANCZOS` resampling, max 80 px (100 px for PDF).
- **Temp files** — image PNGs are written to a `tempfile.mkdtemp` directory during export and the entire directory is removed with `shutil.rmtree` after the document is saved.

### Excel (`.xlsx`) — `export_excel`

Uses **openpyxl**.

| Element | Style |
|---|---|
| Header row | Bold white text, dark blue fill (`#1F3864`), centred, 28 pt height |
| Data rows | 10 pt Calibri; alternating white / light blue (`#DCE6F1`) |
| Image column | 72 pt row height; `OneCellAnchor` positions image centred within the cell |
| Frozen pane | Row 1 (header) stays visible on scroll |
| Borders | Thin border on all cells |

### PDF — `export_pdf`

Uses **reportlab** with `SimpleDocTemplate` on A4 paper (15 mm margins).

- Column widths are scaled proportionally to fill the available width.
- Text cells are `Paragraph` objects with `ParagraphStyle` (avoids the `FONT`/`ALIGN` per-cell conflict in `TableStyle`).
- Header paragraphs use `Helvetica-Bold` size 9, white; data paragraphs use `Helvetica` size 8.
- `repeatRows=1` keeps the header on every page.

### Word (`.docx`) — `export_word`

Uses **python-docx**.

- Page margins 1.5 cm on all sides.
- Header cells use `w:shd` XML element for background colour and `RGBColor` white bold font.
- Column widths are proportional to `_PDF_COL_WIDTHS_MM`.
- Images are inserted with `add_picture`; width is clamped to 90 % of the column width.
- Vertical alignment set via `w:vAlign` XML element.

---

## Configuration & Persistence

### `config.json`

Written next to the script. Stores column visibility and visual column order:

```json
{
  "columns": {
    "#": true,
    "Image": true,
    "Source File": true,
    "Part ID": true,
    "Description": true,
    "Color ID": false,
    "Color Name": true,
    "Type": true,
    "Qty": true,
    "Condition": true,
    "Max Price": true
  },
  "col_order": ["#", "Image", "Source File", "Part ID", "Description", "Color Name", "Type", "Qty", "Condition", "Max Price"]
}
```

`col_order` stores the visual order of column names and is restored on startup via `horizontalHeader().moveSection()`. Write errors (e.g. permissions) are silently ignored so the app continues to function.

### `descriptions.json`

A flat JSON object mapping `{item_type}_{item_id}` → description string. Grows incrementally as new items are encountered; never cleared automatically.

---

## Caching

| Cache | Location | Lifetime |
|---|---|---|
| Image cache | `tempfile.mkdtemp(prefix="bl_cache_")` | Session — deleted when the application closes |
| Description cache | `descriptions.json` next to the script | Persistent across sessions |

Images are not stored permanently because Bricklink CDN images are stable by item+colour and download fast. Descriptions are cached persistently because scraping the catalog page takes ~200–500 ms per item.

---

## Keyboard Shortcuts

| Shortcut | Action |
|---|---|
| Ctrl+C | Copy selected cells as tab-separated text to clipboard |

Column headers support click-to-sort (ascending/descending). Rows can be reordered by dragging. Column headers can be reordered by dragging.

---

## Dependency Reference

| Package | Purpose |
|---|---|
| `PyQt6` | GUI framework |
| `requests` | HTTP downloads (images, catalog pages) |
| `Pillow` | Image decoding, decompression-bomb protection, thumbnail generation |
| `openpyxl` | Excel .xlsx creation |
| `reportlab` | PDF generation |
| `python-docx` | Word .docx creation |

Standard library modules used: `sys`, `io`, `json`, `re`, `shutil`, `tempfile`, `urllib.parse`, `xml.etree.ElementTree`, `pathlib`, `dataclasses`, `typing`.
