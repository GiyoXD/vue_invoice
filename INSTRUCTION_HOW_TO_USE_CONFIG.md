# Invoice Generator Configuration Manual

This document explains the structure and usage of the bundled configuration file (e.g., `JF_config.json`) for the Invoice Generator.

## 1. Overview
The configuration file acts as the central brain for the invoice generation process. It tells the system **what** to build (structure), **how** to fill it (data mappings), and **how** it should look (styling).

The file is divided into 5 main sections:
1.  `_meta` (Identity)
2.  `processing` (Execution Plan)
3.  `layout_bundle` (The Blueprint Core)
4.  `styling_bundle` (The Visuals)
5.  `defaults` (Global Fail-safes)

---

## 2. Detailed Breakdown

### 2.1 `_meta` (Metadata)
**Purpose:** Defines the identity and version of the configuration.
**Who Consumes It:** `config_loader.py` checks this to ensure compatibility.

```json
"_meta": {
    "created_at": "...",
    "customer": "JF",           // Determines the output folder name
    "config_version": "2.1"     // MUST be 2.1+ for this bundled format
}
```

### 2.2 `processing` (Execution Plan)
**Purpose:** This is the "Traffic Controller". It tells the generator which sheets to process and which processor to use.
**Who Consumes It:** `generate_invoice.py`

```json
"processing": {
    "sheets": [
        "Invoice",          // Only sheets listed here are processed
        "Packing List"
    ],
    "data_sources": {
        "Invoice": "aggregation",              // Uses SingleTableProcessor
        "Packing List": "processed_tables_multi" // Uses MultiTableProcessor
    }
}
```

### 2.3 `layout_bundle` (The Construction Blueprint)
**Purpose:** The most critical section. It defines the grid, data flow, and footer logic per sheet.
**Who Consumes It:** `LayoutBuilder` and its sub-builders (`HeaderBuilder`, `DataTableBuilder`, `FooterBuilder`).

#### A. `structure` (The Grid)
*   **Used By:** `HeaderBuilder`
*   **Job:** Defines the visual columns on the Excel sheet.
*   **Key Fields:** `id` (System ID), `header` (Display Name), `width`, `children` (for merged headers).
*   **Filtering:** Flags like `"skip_in_daf": true` or `"skip_in_custom": true` are checked here by strict mode filters.

#### B. `data_flow` (The Wiring)
*   **Used By:** `DataTableBuilder`
*   **Auto-Mapping:** The system acts smart! If your Column ID is `col_qty` and your data has a key `col_qty`, it maps automatically. **You DO NOT need to write a mapping rule.**
*   **When to map?** Only when:
    1.  Keys don't match (e.g., `col_new_qty` needs data from `col_qty_pcs`).
    2.  You need a **Fallback** (e.g., for descriptions).

```json
"mappings": {
    // Standard columns like 'col_po', 'col_qty_pcs' are AUTO-MAPPED. No entry needed!

    // Special Case: Description with Fallback
    "col_desc": { 
        "fallback_on_none": "LEATHER"  // If description is empty, use "LEATHER"
    }
}
```

#### C. `footer` (The Closer)
*   **Used By:** `FooterBuilder`
*   **Job:** Automatically builds the "TOTAL" row.
*   **Key Fields:**
    *   `sum_column_ids`: List of column IDs to calculate totals (e.g., `["col_qty", "col_amount"]`).
    *   `merge_rules`: Defines how to merge cells for the "TOTAL:" label.

### 2.4 `styling_bundle` (The Paint)
**Purpose:** Decouples visual styling from structure.
**Who Consumes It:** `style_applier.py` & Builders.

*   **`row_contexts`**: Sets global height and font for 'header', 'data', and 'footer' rows.
*   **`columns`**: Sets specific alignment (Center/Right/Left) and Number Format for each column ID.
    *   *Example:* To change `col_amount` to use commas, set `number_format` here.

### 2.5 `defaults` (Global Rules)
**Purpose:** Global toggles and settings.
**Who Consumes It:** Mostly `FooterBuilder`.
*   Example: `"show_pallet_count": true`

---

## 3. Common Tasks Checklist

### Task 1: Adding a New Column
1.  **Structure**: Add the column definition to `layout_bundle > [Sheet] > structure > columns`. assign it a unique `id` (e.g., `col_new`).
2.  **Mapping**: Add the data mapping to `layout_bundle > [Sheet] > data_flow > mappings` connecting `col_new` to the JSON data field.
3.  **Styling**: Add a styling entry to `styling_bundle > [Sheet] > columns > col_new` (e.g., define it as `center` aligned).

### Task 2: Modifying the Total Row
1.  Go to `layout_bundle > [Sheet] > footer`.
2.  To sum a new column, add its ID to `sum_column_ids`.
3.  To change the "TOTAL:" label, edit `total_text`.

### Task 3: Excluding a Column for DAF (Custom Mode)
1.  Go to `layout_bundle > [Sheet] > structure > columns`.
2.  Find the column you want to hide.
3.  Add `"skip_in_daf": true`.

---

## 4. Troubleshooting
*   **Column shows up empty?** Check `data_flow > mappings`. You likely forgot to map it.
*   **Invlalid Format?** Check `styling_bundle > columns > [col_id] > number_format`.
*   **Sheet not generating?** Check `processing > sheets`. Is the sheet name listed exactly?
