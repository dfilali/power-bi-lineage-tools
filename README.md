# Power BI Lineage Tools

Python utilities (stdlib only — no pip required) for extracting and documenting Power BI semantic models and reports.

---

## Tools

### `pbit_extract.py` — `.pbit` → CSV extractor

Reads a **Power BI Template** (`.pbit`) file and exports **every object** (tables, columns, measures, partitions, shared expressions, relationships, RLS roles, pages, visuals, visual fields) into a **single long-format CSV**.

```bash
python pbit_extract.py mon_rapport.pbit
python pbit_extract.py mon_rapport.pbit --output extraction.csv
python pbit_extract.py mon_rapport.pbit --delimiter , --no-sanitize
```

**CLI options**

| Option | Default | Description |
|--------|---------|-------------|
| `fichier.pbit` | *(required)* | Path to the `.pbit` file |
| `--output` / `-o` | `<stem>_extraction.csv` | Output CSV path |
| `--delimiter` / `-d` | `;` | CSV column delimiter |
| `--no-sanitize` | *(off)* | Disable masking of org-specific connection strings |

**CSV columns (25)**

```
object_type, table, name, parent, data_type, is_calculated, is_hidden,
expression, expression_language, source, format_string, description,
display_folder, page, visual_type, visual_title, field, role,
from_table, from_column, to_table, to_column, cardinality, cross_filter, extra
```

**`object_type` values**

| Value | Description |
|-------|-------------|
| `table` | Regular (M-based) table |
| `calculated_table` | DAX calculated table |
| `partition` | Partition with M or DAX expression |
| `column` | Regular or calculated column |
| `measure` | DAX measure |
| `shared_expression` | Shared M function (parameters excluded) |
| `relationship` | Model relationship |
| `role` | RLS role (one row for the role + one per table permission) |
| `page` | Report page |
| `visual` | Visual on a report page |
| `visual_field` | Field binding inside a visual |

---

### `pbip_extract.py` — `.pbip` → Markdown / Copilot KB generator

Reads a **Power BI Project** folder (`.pbip`) and generates either:
- **Markdown documentation** (human-readable reference)
- **Copilot knowledge base** (plain-text for Copilot / ChatGPT context injection)

Supports both TMSL (`.bim`) and TMDL (`.tmdl`) semantic model formats.

```bash
python pbip_extract.py myproject.pbip --output docs.md
python pbip_extract.py myproject.pbip --copilot --output kb.txt
```

---

## Requirements

- **Python 3.8+**
- **No external dependencies** — stdlib only (`zipfile`, `json`, `csv`, `re`, `argparse`, …)

---

## Privacy / Sanitizer

By default, `pbit_extract.py` masks organisation-specific values before writing them to the CSV:

| What is masked | Replaced by |
|----------------|-------------|
| SharePoint site URLs | `https://[TENANT].sharepoint.com/sites/[SITE]/...` |
| Fabric SQL endpoint hostnames | `[FABRIC_SQL_ENDPOINT]` |
| `workspaceId`, `lakehouseId`, `driveId`, … | `[WORKSPACE_ID]`, … |
| `schema = "brons_lh"` (line-level assignments) | `schema = [SCHEMA]` |

Use `--no-sanitize` to disable masking (not recommended for shared/public files).

---

## Adding connector patterns

To recognise a new M connector, add a tuple to `_CONNECTOR_PATTERNS` in `pbit_extract.py` (line ~129). **Order matters** — more specific patterns must come before generic ones.

```python
_CONNECTOR_PATTERNS = [
    # existing entries ...

    # Example: Snowflake
    (re.compile(r"\bSnowflake\.Databases\s*\(", re.IGNORECASE), "Snowflake"),

    # Example: Azure Data Lake Storage Gen2
    (re.compile(r"\bAzureStorage\.DataLake\s*\(", re.IGNORECASE), "ADLS Gen2"),
]
```

> **Tip:** Find the function name in Power Query's Advanced Editor (`Source = SomeConnector.Function(...)`).

---

## Known limits

1. `Report/Layout` only covers the legacy `sections / visualContainers` format; new-format report definitions inside a `.pbit` are rare but not supported.
2. Multi-step M queries cannot always pinpoint the primary connector without execution.
3. RLS filter expressions embedded as lists are joined with newlines.
4. Visual fields from `prototypeQuery/dataTransforms` may miss roles when binding metadata is absent.
5. Shared expressions marked `IsParameterQuery = true` are excluded.

---

## License

MIT
