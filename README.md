# Bricklink2Document

A PyQt6 desktop app that loads Bricklink wanted-list XML exports, previews parts with live images, and exports the inventory to Excel, PDF, or Word.

## Features

- Open one or multiple Bricklink XML wanted-list files at once
- Part images and descriptions fetched automatically from Bricklink (cached locally)
- Sortable, reorderable table with drag-and-drop row and column support
- Show/hide individual columns via the Columns dialog
- Export to **Excel** (.xlsx), **PDF** (A4), or **Word** (.docx) — all with embedded images and styled headers
- Column visibility and column order are persisted across sessions in `config.json`
- Copy selected cells to clipboard with Ctrl+C

## Requirements

- Python 3.10+
- PyQt6
- requests
- Pillow
- openpyxl
- reportlab
- python-docx

## Installation

```bash
pip install PyQt6 requests Pillow openpyxl reportlab python-docx
```

## Usage

```bash
python Bricklink2Document.py
```

1. Click **Open XML …** and select one or more Bricklink wanted-list XML files.
2. The app populates the table and downloads part images and descriptions in the background.
3. Reorder rows by dragging them; reorder columns by dragging the header.
4. Use **Columns …** to toggle column visibility.
5. Click **Export Excel**, **Export PDF**, or **Export Word** to save the inventory.

## Exported columns

| Column | Description |
|---|---|
| # | Row number |
| Image | Part thumbnail |
| Source File | Name of the source XML file |
| Part ID | Bricklink item ID |
| Description | Item name fetched from Bricklink |
| Color ID | Numeric Bricklink color ID |
| Color Name | Human-readable color name |
| Type | Part / Set / Minifig / Gear / … |
| Qty | Minimum quantity |
| Condition | New / Used / Any |
| Max Price | Maximum price (USD) |

## Files

| File | Purpose |
|---|---|
| `Bricklink2Document.py` | Main application |
| `config.json` | Saved column visibility and column order (auto-generated) |
| `descriptions.json` | Cached item descriptions fetched from Bricklink (auto-generated) |

## License

See [LICENSE](LICENSE).
