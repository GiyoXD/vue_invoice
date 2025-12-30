# Config Manager Documentation

## Overview
The **Config Manager** is responsible for automating the creation of invoice generation configurations. It takes a raw Excel invoice/packing list (populated with data) and transforms it into a reusable **Bundle**.

## The Blueprint Concept
A "Blueprint" (formerly Bundle) is a self-contained directory that contains everything needed to generate invoices for a specific customer or template type.

**Structure:**
```
database/blueprints/registry/
  └── {CUSTOMER_CODE}/
      ├── {CUSTOMER_CODE}_config.json       # Processing rules (mappings, layout, styles)
      ├── {CUSTOMER_CODE}.xlsx              # Cleaned, blank Excel template
      └── {CUSTOMER_CODE}_template.json     # Extra layout metadata
```

## Core Components

### 1. BlueprintGenerator (`blueprint_generator/blueprint_generator.py`)
The main engine. It orchestrates the entire process:
- **Scans** the raw Excel file using `ExcelLayoutScanner`.
- **Builds** the configuration JSON using `ConfigBuilder`.
- **Sanitizes** the Excel file (removes data, keeps formatting) using `ExcelTemplateSanitizer`.
- **Saves** the results into a structured blueprint folder.

### 2. Excel Layout Scanner (`blueprint_generator/excel_scanner.py`)
Scans the Excel file to detect:
- Header rows (by looking for keywords like "Description", "QTY").
- Data columns (mapping headers to system IDs).
- Styling (fonts, borders, merges).

### 3. Excel Template Sanitizer (`blueprint_generator/excel_sanitizer.py`)
Creates the `{CUSTOMER_CODE}_template.xlsx`.
- Removes distinct data rows.
- Preserves headers and footers.
- Maintains column widths and row heights.
- **Robustness**: If sanitization fails (e.g., due to complex images), the system falls back to using the original file as the template to ensure a config is still generated.

## CLI Usage

The entry point is `core/blueprint_generator/main.py`.

### Basic Usage
Generate a bundle from an Excel file:
```bash
python core/blueprint_generator/main.py path/to/invoice.xlsx
```
*Output*: Creates a bundle in `database/blueprints/registry/`.

### Custom Output Directory
```bash
python core/blueprint_generator/main.py path/to/invoice.xlsx -o path/to/output_dir
```

### Verbose Logging
See detailed analysis steps:
```bash
python core/blueprint_generator/main.py path/to/invoice.xlsx -v
```

## Integration with Invoice Generation
When an invoice is generated:
1. The **Asset Resolver** looks for a bundle matching the input filename (e.g., `JF25059.xlsx` -> `JF25059` or `JF`).
2. It loads the `config.json` and `template.xlsx` from that bundle.
3. **Strict Mode**: If no specific bundle or config is found, the generation **fails**. It does NOT fall back to defaults.

## Troubleshooting

### "Shipping List as Template"
If the generated invoice looks like the original shipping list (data written over existing data):
- This usually means **Template Cleaning Failed**.
- Check logs for "Failed to generate/save cleaned template".
- The system fell back to using the original file as the template.
- **Solution**: Manually create a clean template or ensure the input file is simple enough for the cleaner.

### "No Valid Sheets"
- If the generator fails to detect sheets to process.
- The system now **falls back to the first sheet** automatically.
- Check `config.json` `processing.sheets` to see what was detected.