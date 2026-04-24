# Data Mapping Exec

A lightweight Python script that exports a source-to-target-mapping table into four formats: **XLSX, CSV, JSON, and YAML**.
Done by few LLM commands, while sipping Heineken on Friday afternoon for Data Mapping Registry, or whatever need to share or integrate mapping definitions in various formats.

---

## Project Structure

```
mapping_exec/
├── transform_mapping.py   # Main script
├── requirements.txt       # Python dependencies
└── README.md              # This file
```

### Output files (generated on run)

```
mapping_exec/
├── data_mapping.xlsx
├── data_mapping.csv
├── data_mapping.json
└── data_mapping.yaml
```

---

## Data Schema

Each mapping record follows a three-layer source-to-target structure:

| Column | Description |
|---|---|
| `Source Entity` | Origin table or object in the source system |
| `Source Column` | Field name in the source system |
| `EDM Entity` | Canonical entity in the Enterprise Data Model |
| `EDM Column` | Canonical field name in the EDM |
| `Target Entity` | Destination table in the data warehouse |
| `Target Column` | Field name in the target DWH table |

---

## Source Systems Covered

| Source System | EDM Entity | Target (DWH) |
|---|---|---|
| `CRM_Customers` | `Party` | `DWH_Dim_Customer` |
| `ERP_Orders` | `SalesOrder` | `DWH_Fact_Sales` |
| `ERP_Products` | `Product` | `DWH_Dim_Product` |
| `HR_Employees` | `Person` | `DWH_Dim_Employee` |
| `FIN_Invoices` | `Invoice` | `DWH_Fact_Finance` |

---

## Requirements

- Python 3.8+
- pip packages listed in `requirements.txt`

---

## Installation

```bash
pip install -r requirements.txt
```

If `pip` and `python` point to different environments, use:

```bash
python -m pip install -r requirements.txt
```

---

## Usage

```bash
python transform_mapping.py
```

All four output files are written to the **same directory as the script**.

---

## Output Formats

### XLSX
Styled spreadsheet with a forest green header, alternating row fills, frozen header row, and auto-filter enabled. Built with `openpyxl`.

### CSV
Plain UTF-8 comma-separated file. Compatible with Excel, Google Sheets, and most data tools.

### JSON
Array of objects, one record per mapping row. Indented for readability.

```json
[
  {
    "Source Entity": "CRM_Customers",
    "Source Column": "cust_id",
    "EDM Entity": "Party",
    "EDM Column": "party_id",
    "Target Entity": "DWH_Dim_Customer",
    "Target Column": "customer_key"
  }
]
```

### YAML
Human-readable list of mapping records. Useful for version-controlled config files or pipeline definitions.

```yaml
- Source Entity: CRM_Customers
  Source Column: cust_id
  EDM Entity: Party
  EDM Column: party_id
  Target Entity: DWH_Dim_Customer
  Target Column: customer_key
```

---

## Extending the Mapping

To add new mappings, append rows to the `DATA` list in `transform_mapping.py`:

```python
DATA = [
    # existing rows...
    ["NEW_Source", "source_col", "EDM_Entity", "edm_col", "DWH_Target", "target_col"],
]
```

Re-run the script to regenerate all four output files.

---

## Troubleshooting

| Error | Fix |
|---|---|
| `ModuleNotFoundError: No module named 'openpyxl'` | Run `python -m pip install -r requirements.txt` |
| `ModuleNotFoundError: No module named 'yaml'` | Run `python -m pip install pyyaml` |
| `OSError: Read-only file system` | Ensure the script's directory is writable |
