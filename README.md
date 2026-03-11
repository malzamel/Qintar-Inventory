# Qintar Inventory

**Version 1.0**

A jewelry store product management desktop application with Arabic interface. Manage product data, assign images from folders, and import/export to Excel while preserving file structure and barcode formatting.

إدارة منتجات متجر المجوهرات - تطبيق سطح مكتب لإدارة المنتجات مع واجهة عربية

## Features

- **Product Management** – Full product catalog with 19 fields (name, category, price, weight, barcode, karat, supplier, etc.)
- **Image Assignment** – Select folder with product images; click to assign images to products
- **Excel Integration** – Import from existing Excel files and export while preserving structure
- **Barcode Support** – Leading zeros preserved (e.g., 001234)
- **Undo/Redo** – Full history support
- **Auto-save** – Automatic backups every 5 minutes
- **Search & Filter** – Quick product search
- **Column Customization** – Show/hide columns; preferences saved

## Requirements

- Python 3.8 or higher
- Windows (primary target; may work on Linux/macOS with tkinter)

## Installation

```bash
# Clone the repository
git clone https://github.com/YOUR_USERNAME/Qintar-Inventory.git
cd Qintar-Inventory

# Create virtual environment (recommended)
python -m venv .venv
.venv\Scripts\activate

# Install dependencies
pip install -r requirements.txt
```

## Usage

```bash
python main.py
```

1. **Choose Folder** – Click "اختيار مجلد الصور" to select a folder containing product images and optionally an existing Excel file
2. **Assign Images** – Click images in the left panel to assign them to the selected product
3. **Edit Products** – Right-click a row for Edit, Copy, Paste, or Delete
4. **Save** – Use Ctrl+S or the Save button; data exports to Excel in the selected folder

## Keyboard Shortcuts

| Shortcut | Action |
|----------|--------|
| Ctrl+S | Save |
| Ctrl+Z | Undo |
| Ctrl+Y | Redo |
| Ctrl+F | Advanced search |
| F3 / Shift+F3 | Find next / previous |
| Ctrl+N | Add new product |
| Delete | Remove selected image |
| Ctrl+A | Select all images |
| Escape | Clear selection |

## Configuration

- `jewelry_config.json` – Default values (supplier, invoice, karat), window geometry
- `column_preferences.json` – Column visibility settings

## License

Proprietary – All rights reserved
