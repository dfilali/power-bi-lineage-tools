#!/usr/bin/env python3
"""
pbit_extract.py — Power BI Template (.pbit) → single CSV extractor

Reads a .pbit archive (ZIP), parses the TMSL semantic model (DataModelSchema)
and the legacy report layout (Report/Layout), and writes every object as one
row in a "long-format" CSV file.

Usage:
    python pbit_extract.py mon_rapport.pbit --output extraction.csv
    python pbit_extract.py mon_rapport.pbit --delimiter , --no-sanitize

Requirements: Python 3.8+, stdlib only (no pip installs needed).

Known limits (≤5):
1. Report/Layout visual config only covers the legacy "sections/visualContainers"
   format; new-format (.pbip-style) report definitions inside a .pbit are rare but
   not supported here.
2. Partitions with multiple M steps cannot always pinpoint the primary connector
   function without executing the query.
3. RLS filter expressions embedded as lists are joined with newlines; complex
   multi-table permission sets may lose some structure.
4. Visual fields extracted from prototypeQuery/dataTransforms may miss roles when
   the binding metadata is absent from the layout JSON.
5. Shared expressions marked IsParameterQuery=true are excluded (they often carry
   organisation-specific default values).
"""

import argparse
import csv
import io
import json
import re
import sys
import urllib.parse
import zipfile
from pathlib import Path
from typing import Optional


# ---------------------------------------------------------------------------
# Sanitizer — masks org-specific connection strings before output
# (identical to pbip_extract.py, kept verbatim as requested)
# ---------------------------------------------------------------------------

_SHAREPOINT_URL_RE = re.compile(
    r'https?://[a-zA-Z0-9\-]+\.sharepoint\.com/sites/[^"\')\s,\]]+',
    re.IGNORECASE,
)

_FABRIC_SQL_ENDPOINT_RE = re.compile(
    r'\b[a-z0-9]{6,}\.datawarehouse\.fabric\.microsoft\.com\b',
    re.IGNORECASE,
)

_INFRA_KEY_PLACEHOLDERS = {
    "sqlendpoint": "FABRIC_SQL_ENDPOINT",
    "sql endpoint": "FABRIC_SQL_ENDPOINT",
    "workspaceid": "WORKSPACE_ID",
    "lakehouseid": "LAKEHOUSE_ID",
    "warehouseid": "WAREHOUSE_ID",
    "datasetid": "DATASET_ID",
    "groupid": "WORKSPACE_ID",
    "driveid": "SHAREPOINT_DRIVE_ID",
    "itemid": "SHAREPOINT_ITEM_ID",
    "siteid": "SHAREPOINT_SITE_ID",
}

_INFRA_KV_RE = re.compile(
    r'\b(' + "|".join(re.escape(k) for k in _INFRA_KEY_PLACEHOLDERS) + r')'
    r'(\s*[:=]\s*)'
    r'"?([A-Za-z0-9_.\-]+)"?',
    re.IGNORECASE,
)

_SCHEMA_ASSIGN_RE = re.compile(
    r'(?im)^(\s*)schema(\s*[:=]\s*)"?([A-Za-z0-9_.\-]+)"?',
)


def sanitize(text: str) -> str:
    """
    Masks org-specific infrastructure details in Power Query / DAX expressions
    before output: SharePoint URLs, Fabric SQL endpoints, and known Fabric /
    SharePoint object IDs (workspaceId, lakehouseId, driveId, ...).

    The query structure itself (function names, transformation steps, column
    references) is always preserved — only org-specific literal values are
    replaced by a placeholder so the model remains understandable without
    leaking which tenant/workspace/resource it runs on.
    """
    if not text:
        return text

    def _replace_sharepoint(m: re.Match) -> str:
        url = m.group(0)
        try:
            parsed = urllib.parse.urlparse(url)
            path_decoded = urllib.parse.unquote(parsed.path)
            parts = path_decoded.strip("/").split("/")
            if len(parts) >= 2 and parts[0].lower() == "sites":
                rest = "/" + "/".join(parts[2:]) if len(parts) > 2 else ""
                masked = f"https://[TENANT].sharepoint.com/sites/[SITE]{rest}"
            else:
                masked = "https://[TENANT].sharepoint.com/[PAD]"
        except Exception:
            masked = "[SHAREPOINT-URL]"
        return masked

    def _replace_infra_kv(m: re.Match) -> str:
        key, sep = m.group(1), m.group(2)
        placeholder = _INFRA_KEY_PLACEHOLDERS[key.lower()]
        return f"{key}{sep}[{placeholder}]"

    def _replace_schema(m: re.Match) -> str:
        return f"{m.group(1)}schema{m.group(2)}[SCHEMA]"

    text = _SHAREPOINT_URL_RE.sub(_replace_sharepoint, text)
    text = _FABRIC_SQL_ENDPOINT_RE.sub("[FABRIC_SQL_ENDPOINT]", text)
    text = _INFRA_KV_RE.sub(_replace_infra_kv, text)
    text = _SCHEMA_ASSIGN_RE.sub(_replace_schema, text)
    return text


# ---------------------------------------------------------------------------
# Source-connector classifier (reused from pbip_extract.py)
# ---------------------------------------------------------------------------

_CONNECTOR_PATTERNS = [
    (re.compile(r"\.datawarehouse\.fabric\.microsoft\.com", re.IGNORECASE),
     "Fabric SQL Endpoint"),
    (re.compile(r"\bLakehouse\.Contents\s*\(", re.IGNORECASE),
     "Fabric Lakehouse"),
    (re.compile(r"\bDataflow\b|\bPowerPlatform\.Dataflows\s*\(", re.IGNORECASE),
     "Power BI Dataflow"),
    (re.compile(r"sharepoint\.com|\bSharePoint\.(Files|Contents)\s*\(", re.IGNORECASE),
     "SharePoint"),
    (re.compile(r"\bOneDrive\.", re.IGNORECASE),
     "OneDrive"),
    (re.compile(r"\bSql\.Database\s*\(", re.IGNORECASE),
     "Sql.Database"),
    (re.compile(r"\bPostgreSQL\.Database\s*\(", re.IGNORECASE),
     "PostgreSQL.Database"),
    (re.compile(r"\bMySQL\.Database\s*\(", re.IGNORECASE),
     "MySQL.Database"),
    (re.compile(r"\bOdbc\.(Query|DataSource)\s*\(", re.IGNORECASE),
     "ODBC"),
    (re.compile(r"\bOData\.Feed\s*\(", re.IGNORECASE),
     "OData.Feed"),
    (re.compile(r"\bWeb\.Contents\s*\(", re.IGNORECASE),
     "Web.Contents"),
    (re.compile(r"\bExcel\.Workbook\s*\(", re.IGNORECASE),
     "Excel.Workbook"),
    (re.compile(r"\bCsv\.Document\s*\(", re.IGNORECASE),
     "Csv.Document"),
    (re.compile(r"\bJson\.Document\s*\(", re.IGNORECASE),
     "Json.Document"),
    (re.compile(r"\bFolder\.(Files|Contents)\s*\(", re.IGNORECASE),
     "Folder"),
]


def _detect_connector(m_expression: str) -> str:
    """Return the primary connector name found in a Power Query M expression."""
    if not m_expression:
        return ""
    for pattern, label in _CONNECTOR_PATTERNS:
        if pattern.search(m_expression):
            return label
    return "Other/custom M"


# ---------------------------------------------------------------------------
# TMSL parser — reads the DataModelSchema JSON from inside the .pbit ZIP
# (adapted from pbip_extract.py: reads from a dict instead of a file path)
# ---------------------------------------------------------------------------

class TMSLParser:
    """Parses a TMSL model dict (already decoded from DataModelSchema)."""

    def __init__(self, data: dict):
        self.data = data

    def _model(self) -> dict:
        return self.data.get("model", self.data)

    def tables(self) -> list:
        """All tables — hidden ones included (exported with is_hidden=true)."""
        return self._model().get("tables", [])

    def relationships(self) -> list:
        return self._model().get("relationships", [])

    def roles(self) -> list:
        return self._model().get("roles", [])

    def shared_expressions(self) -> list:
        """Shared M functions — parameters (IsParameterQuery=true) excluded."""
        result = []
        for expr in self._model().get("expressions", []):
            raw = expr.get("expression", "")
            text = "\n".join(raw) if isinstance(raw, list) else str(raw)
            if re.search(r"IsParameterQuery\s*=\s*true", text, re.IGNORECASE):
                continue
            result.append({
                "name": expr.get("name", ""),
                "kind": expr.get("kind", ""),
                "expression": text.strip(),
            })
        return result

    def get_measures(self, table: dict) -> list:
        return table.get("measures", [])

    def get_columns(self, table: dict) -> list:
        result = []
        for c in table.get("columns", []):
            if c.get("type") == "rowNumber":
                continue
            col = dict(c)
            raw_expr = c.get("expression", "")
            expr = "\n".join(raw_expr) if isinstance(raw_expr, list) else str(raw_expr or "")
            col["expression"] = expr.strip()
            col["isCalculated"] = c.get("type") == "calculated" or bool(col["expression"])
            result.append(col)
        return result

    def get_partitions(self, table: dict) -> list:
        return table.get("partitions", [])


# ---------------------------------------------------------------------------
# Report parser — reads Report/Layout (legacy sections/visualContainers JSON)
# (adapted from pbip_extract.py: reads from a dict instead of a file path)
# ---------------------------------------------------------------------------

def _extract_field_name(field: dict) -> Optional[str]:
    if not isinstance(field, dict):
        return None

    if "Measure" in field:
        m = field["Measure"]
        entity = m.get("Expression", {}).get("SourceRef", {}).get("Entity", "")
        prop = m.get("Property", "")
        return f"{entity}[{prop}]" if entity else f"[{prop}]"

    if "Column" in field:
        c = field["Column"]
        entity = c.get("Expression", {}).get("SourceRef", {}).get("Entity", "")
        prop = c.get("Property", "")
        return f"{entity}[{prop}]" if entity else f"[{prop}]"

    if "Aggregation" in field:
        agg = field["Aggregation"]
        func_map = {0: "SUM", 1: "AVG", 2: "COUNT", 3: "MIN", 4: "MAX",
                    5: "COUNTROWS", 6: "COUNTDISTINCT"}
        func = func_map.get(agg.get("Function", -1), "AGG")
        inner = _extract_field_name(agg.get("Expression", {}))
        return f"{func}({inner})" if inner else func

    if "HierarchyLevel" in field:
        hl = field["HierarchyLevel"]
        hier_expr = hl.get("Expression", {}).get("Hierarchy", {})
        entity = hier_expr.get("Expression", {}).get("SourceRef", {}).get("Entity", "")
        hier_name = hier_expr.get("Hierarchy", "")
        level = hl.get("Level", "")
        parts = [p for p in [entity, hier_name, level] if p]
        return ".".join(parts)

    if "SourceRef" in field:
        return field["SourceRef"].get("Entity", "")

    return None


def _parse_visual_config(config: dict) -> dict:
    """Extract visual type, title, and field bindings from a visual config dict."""
    sv = config.get("visual", config.get("singleVisual", {}))
    visual_type = sv.get("visualType", "unknown")

    title = ""
    for obj_key in ("objects", "vcObjects"):
        title_obj = sv.get(obj_key, {}).get("title", [])
        if isinstance(title_obj, list) and title_obj:
            props = title_obj[0].get("properties", {})
            title_text = props.get("text", {})
            if isinstance(title_text, dict):
                val = title_text.get("expr", {}).get("Literal", {}).get("Value", "")
                title = val.strip("'\"")
            elif isinstance(title_text, str):
                title = title_text
            if title:
                break

    fields: list = []

    query_state = sv.get("query", {}).get("queryState", {})
    if query_state:
        for role_name, role_data in query_state.items():
            for proj in role_data.get("projections", []):
                field_ref = proj.get("field", {})
                native_ref = proj.get("nativeQueryRef", proj.get("queryRef", ""))
                field_name = _extract_field_name(field_ref) or native_ref
                if field_name:
                    fields.append({
                        "field": field_name,
                        "role": role_name,
                        "displayName": native_ref,
                    })

    if not fields:
        proto_select = sv.get("prototypeQuery", {}).get("Select", [])
        dt_selects = sv.get("dataTransforms", {}).get("selects", [])
        dt_lookup: dict = {}
        for dt in dt_selects:
            qname = dt.get("queryName", "")
            roles_dict = dt.get("roles", {})
            role_name = next(iter(roles_dict.keys()), "") if roles_dict else ""
            dt_lookup[qname] = {"role": role_name, "displayName": dt.get("displayName", "")}

        for i, item in enumerate(proto_select):
            field_name = _extract_field_name(item)
            if not field_name:
                continue
            q_name = item.get("Name", "")
            meta = dt_lookup.get(q_name, dt_selects[i] if i < len(dt_selects) else {})
            roles_dict = meta.get("roles", {}) if isinstance(meta, dict) else {}
            fields.append({
                "field": field_name,
                "role": next(iter(roles_dict.keys()), meta.get("role", "")),
                "displayName": meta.get("displayName", ""),
            })

    return {"visualType": visual_type, "title": title, "fields": fields}


def parse_report(layout_data: dict) -> dict:
    """Parse a Report/Layout dict (sections/visualContainers format)."""
    result: dict = {"pages": []}

    for section in layout_data.get("sections", []):
        page = {
            "name": section.get("displayName", section.get("name", "Unknown")),
            "visuals": [],
        }

        for vc in section.get("visualContainers", []):
            config_str = vc.get("config", "{}")
            try:
                config = json.loads(config_str) if isinstance(config_str, str) else config_str
            except json.JSONDecodeError:
                config = {}

            if "singleVisual" not in config:
                continue

            visual_info = _parse_visual_config(config)

            if visual_info["visualType"] in ("", "unknown", "shape", "image", "textbox"):
                if not visual_info["fields"]:
                    continue

            page["visuals"].append(visual_info)

        result["pages"].append(page)

    return result


# ---------------------------------------------------------------------------
# .pbit reader — opens the ZIP and decodes the two key files
# ---------------------------------------------------------------------------

def _read_pbit_member(zf: zipfile.ZipFile, member: str) -> Optional[str]:
    """
    Read a member from the ZIP and decode it.  .pbit files use UTF-16 LE (with
    BOM) for both DataModelSchema and Report/Layout; fall back to utf-8-sig so
    the function also works with .pbit files produced by older tools that used
    UTF-8.
    """
    try:
        raw = zf.read(member)
    except KeyError:
        return None
    for enc in ("utf-16", "utf-8-sig", "utf-8"):
        try:
            return raw.decode(enc)
        except (UnicodeDecodeError, Exception):
            continue
    # Last resort: replace undecodable bytes
    return raw.decode("utf-16", errors="replace")


def open_pbit(pbit_path: Path) -> tuple:
    """
    Open a .pbit file and return (tmsl_parser, report_data).

    tmsl_parser is always a TMSLParser instance (exits on error).
    report_data is a dict with "pages" key, or None if Report/Layout is absent.
    """
    if not pbit_path.exists():
        print(f"ERROR: File not found: {pbit_path}", file=sys.stderr)
        sys.exit(1)
    if not zipfile.is_zipfile(pbit_path):
        print(f"ERROR: Not a valid ZIP/pbit archive: {pbit_path}", file=sys.stderr)
        sys.exit(1)

    with zipfile.ZipFile(pbit_path, "r") as zf:
        # -- DataModelSchema (required) --
        schema_text = _read_pbit_member(zf, "DataModelSchema")
        if schema_text is None:
            print(
                "ERROR: DataModelSchema not found inside the .pbit archive.\n"
                "       The file may be corrupted or use an unsupported format.",
                file=sys.stderr,
            )
            sys.exit(1)
        try:
            schema_data = json.loads(schema_text)
        except json.JSONDecodeError as exc:
            print(f"ERROR: Could not parse DataModelSchema as JSON: {exc}", file=sys.stderr)
            sys.exit(1)

        # -- Report/Layout (optional) --
        layout_text = _read_pbit_member(zf, "Report/Layout")
        report_data: Optional[dict] = None
        if layout_text and layout_text.strip():
            try:
                layout_json = json.loads(layout_text)
                report_data = parse_report(layout_json)
            except json.JSONDecodeError as exc:
                print(
                    f"WARNING: Could not parse Report/Layout as JSON: {exc}\n"
                    "         Report pages/visuals will be skipped.",
                    file=sys.stderr,
                )
        else:
            print("INFO: Report/Layout absent or empty — continuing with model only.")

    parser = TMSLParser(schema_data)
    return parser, report_data


# ---------------------------------------------------------------------------
# DAX dependency analyser
# ---------------------------------------------------------------------------

# Matches 'TableName'[ColumnOrMeasure] or TableName[ColumnOrMeasure]
_DAX_TABLE_REF_RE = re.compile(
    r"'?([A-Za-z0-9_ ]+)'?\[[^\]]+\]",
)


def _normalize_expr(raw) -> str:
    """Flatten a list-of-lines expression to a single string.

    Fix #4: TMSL sometimes stores expressions as ["line1", "line2", ...]
    instead of a plain string.  All dependency analysis and connector
    detection must work on the flattened string, not the raw list.
    """
    if isinstance(raw, list):
        return "\n".join(raw)
    return str(raw or "")


def _extract_dax_table_refs(dax: str, all_table_names: list) -> list:
    """Return sorted list of table names referenced in a DAX expression.

    Two strategies:
    1. Pattern 'TableName'[Field] / TableName[Field] — catches most cases.
    2. Whole-word match of every known table name — catches SUMMARIZECOLUMNS
       and other functions that reference tables without a column qualifier.
    """
    found: set = set()

    # Strategy 1 — bracketed references
    for m in _DAX_TABLE_REF_RE.finditer(dax):
        t = m.group(1).strip()
        if t:
            found.add(t)

    # Strategy 2 — whole-word table name scan
    for t in all_table_names:
        if not t:
            continue
        if re.search(r"\b" + re.escape(t) + r"\b", dax, re.IGNORECASE):
            found.add(t)

    return sorted(found)


# ---------------------------------------------------------------------------
# CSV row builders
# ---------------------------------------------------------------------------

# Fixed column order (26 columns — added "depends_on" for table_dependency)
CSV_COLUMNS = [
    "object_type", "table", "name", "parent",
    "data_type", "is_calculated", "is_hidden",
    "expression", "expression_language", "source",
    "format_string", "description", "display_folder",
    "page", "visual_type", "visual_title",
    "field", "role",
    "from_table", "from_column", "to_table", "to_column",
    "cardinality", "cross_filter",
    "depends_on",
    "extra",
]


def _blank_row() -> dict:
    return {col: "" for col in CSV_COLUMNS}


# Fix #2 — _is_calc_table: normalise list before regex check, and
# Fix #1 — detect SUMMARIZECOLUMNS() in addition to SUMMARIZE()
_CALC_TABLE_DAX_RE = re.compile(
    r"\b(SUMMARIZE|SUMMARIZECOLUMNS|DATATABLE|GENERATESERIES|CALENDAR|CALENDARAUTO"
    r"|UNION|INTERSECT|EXCEPT|CROSSJOIN|GENERATE|GENERATEALL"
    r"|ADDCOLUMNS|SELECTCOLUMNS|ROW|TOPN|SAMPLE|FILTERS|VALUES|ALL|ALLNOBLANKROW)"
    r"\s*\(",
    re.IGNORECASE,
)


def _rows_for_table(
    table: dict,
    parser: TMSLParser,
    do_sanitize: bool,
    all_table_names: Optional[list] = None,
) -> list:
    """Yield CSV row dicts for a single table and all its children."""
    rows = []
    table_name = table.get("name", "")
    partitions = parser.get_partitions(table)

    # Fix #4 — normalise the partition source expression to str FIRST so that
    # every downstream check (is_calc_table, _detect_connector, dep analysis)
    # works on a plain string, not a raw list.
    def _pt_expr(pt: dict) -> str:
        src = pt.get("source", {})
        if isinstance(src, str):
            return ""
        raw = src.get("expression", src.get("query", ""))
        return _normalize_expr(raw).strip()

    def _pt_type(pt: dict) -> str:
        src = pt.get("source", {})
        if isinstance(src, str):
            return ""
        return src.get("type", "")

    # Fix #1 + #2 — is_calc_table now uses normalised expr and SUMMARIZECOLUMNS
    is_calc_table = bool(
        table.get("calculationGroup")
        or any(_pt_type(p) == "calculated" for p in partitions)
        or any(
            _pt_type(p) != "calculated"
            and _CALC_TABLE_DAX_RE.search(_pt_expr(p))
            for p in partitions
        )
    )

    # ---- table row ----
    row = _blank_row()
    row["object_type"] = "calculated_table" if is_calc_table else "table"
    row["table"] = table_name
    row["name"] = table_name
    row["is_hidden"] = str(table.get("isHidden", False)).lower()
    row["description"] = table.get("description", "")
    rows.append(row)

    # ---- partition rows + Fix #3 table_dependency ----
    for pt in partitions:
        pt_expr_str = _pt_expr(pt)
        pt_type_str = _pt_type(pt)

        expr_lang = "DAX" if pt_type_str == "calculated" else "M"
        connector = ""
        if expr_lang == "M" and pt_expr_str:
            connector = _detect_connector(pt_expr_str)
        elif expr_lang == "DAX":
            connector = "DAX"

        expr_out = (sanitize(pt_expr_str) if do_sanitize else pt_expr_str)

        row = _blank_row()
        row["object_type"] = "partition"
        row["table"] = table_name
        row["name"] = pt.get("name", "")
        row["parent"] = table_name
        row["expression"] = expr_out
        row["expression_language"] = expr_lang
        row["source"] = sanitize(connector) if do_sanitize else connector
        row["extra"] = pt.get("mode", "")
        rows.append(row)

        # Fix #3 — table_dependency rows for calculated-table DAX expressions
        if expr_lang == "DAX" and pt_expr_str and all_table_names:
            deps = _extract_dax_table_refs(pt_expr_str, all_table_names)
            for dep_table in deps:
                if dep_table == table_name:
                    continue  # skip self-reference
                dep_row = _blank_row()
                dep_row["object_type"] = "table_dependency"
                dep_row["table"] = table_name
                dep_row["name"] = table_name
                dep_row["parent"] = dep_table
                dep_row["depends_on"] = dep_table
                dep_row["expression_language"] = "DAX"
                dep_row["extra"] = pt.get("name", "")
                rows.append(dep_row)

    # ---- column rows ----
    for col in parser.get_columns(table):
        col_name = col.get("name", "")
        is_calc = col.get("isCalculated", False)
        raw_expr = col.get("expression", "")
        expr_out = (sanitize(raw_expr) if do_sanitize else raw_expr) if is_calc else ""
        source_col = col.get("sourceColumn", "")

        row = _blank_row()
        row["object_type"] = "column"
        row["table"] = table_name
        row["name"] = col_name
        row["parent"] = table_name
        row["data_type"] = col.get("dataType", col.get("type", ""))
        row["is_calculated"] = str(is_calc).lower()
        row["is_hidden"] = str(col.get("isHidden", False)).lower()
        row["expression"] = expr_out
        row["expression_language"] = "DAX" if is_calc and raw_expr else ""
        row["source"] = source_col
        row["format_string"] = col.get("formatString", "")
        row["description"] = col.get("description", "")
        row["display_folder"] = col.get("displayFolder", "")
        rows.append(row)

    # ---- measure rows ----
    for measure in parser.get_measures(table):
        m_name = measure.get("name", "")
        raw_expr = measure.get("expression", "")
        if isinstance(raw_expr, list):
            raw_expr = "\n".join(raw_expr)
        raw_expr = raw_expr.strip()
        expr_out = sanitize(raw_expr) if do_sanitize else raw_expr

        row = _blank_row()
        row["object_type"] = "measure"
        row["table"] = table_name
        row["name"] = m_name
        row["parent"] = table_name
        row["is_hidden"] = str(measure.get("isHidden", False)).lower()
        row["expression"] = expr_out
        row["expression_language"] = "DAX" if raw_expr else ""
        row["format_string"] = measure.get("formatString", "")
        row["description"] = measure.get("description", "")
        row["display_folder"] = measure.get("displayFolder", "")
        rows.append(row)

    return rows


def _rows_for_shared_expressions(parser: TMSLParser, do_sanitize: bool) -> list:
    rows = []
    for expr in parser.shared_expressions():
        raw_expr = expr.get("expression", "")
        expr_out = sanitize(raw_expr) if do_sanitize else raw_expr

        row = _blank_row()
        row["object_type"] = "shared_expression"
        row["name"] = expr.get("name", "")
        row["expression"] = expr_out
        row["expression_language"] = "M"
        row["extra"] = expr.get("kind", "")
        rows.append(row)
    return rows


def _rows_for_relationships(parser: TMSLParser) -> list:
    rows = []
    for rel in parser.relationships():
        from_t = rel.get("fromTable", rel.get("fromTableId", ""))
        from_c = rel.get("fromColumn", rel.get("fromColumnId", ""))
        to_t = rel.get("toTable", rel.get("toTableId", ""))
        to_c = rel.get("toColumn", rel.get("toColumnId", ""))
        card = rel.get("cardinality", rel.get("fromCardinality", ""))
        cf = rel.get("crossFilteringBehavior", rel.get("crossFilter", ""))

        # Build a readable name for the relationship
        rel_name = rel.get("name", f"{from_t}[{from_c}]→{to_t}[{to_c}]")

        row = _blank_row()
        row["object_type"] = "relationship"
        row["name"] = rel_name
        row["from_table"] = from_t
        row["from_column"] = from_c
        row["to_table"] = to_t
        row["to_column"] = to_c
        row["cardinality"] = card
        row["cross_filter"] = cf
        row["is_hidden"] = str(rel.get("isActive", True) is False).lower()
        row["extra"] = "inactive" if rel.get("isActive") is False else ""
        rows.append(row)
    return rows


def _rows_for_roles(parser: TMSLParser) -> list:
    rows = []
    for role in parser.roles():
        role_name = role.get("name", "")
        model_perm = role.get("modelPermission", "")

        # One row for the role itself
        row = _blank_row()
        row["object_type"] = "role"
        row["name"] = role_name
        row["extra"] = model_perm
        rows.append(row)

        # One row per table permission inside the role
        for perm in role.get("tablePermissions", []):
            t = perm.get("table", perm.get("name", ""))
            raw_filter = perm.get("filterExpression", perm.get("expression", ""))
            if isinstance(raw_filter, list):
                raw_filter = "\n".join(raw_filter)

            row = _blank_row()
            row["object_type"] = "role"
            row["name"] = role_name
            row["parent"] = role_name
            row["table"] = t
            row["expression"] = (raw_filter or "").strip()
            row["expression_language"] = "DAX" if raw_filter else ""
            rows.append(row)

    return rows


def _rows_for_report(report_data: dict) -> list:
    """Emit page, visual, and visual_field rows from a parsed report dict."""
    rows = []
    for page in report_data.get("pages", []):
        page_name = page.get("name", "")

        # page row
        row = _blank_row()
        row["object_type"] = "page"
        row["name"] = page_name
        row["page"] = page_name
        rows.append(row)

        for visual in page.get("visuals", []):
            v_type = visual.get("visualType", "")
            v_title = visual.get("title", "")
            v_label = v_title or v_type

            # visual row
            row = _blank_row()
            row["object_type"] = "visual"
            row["name"] = v_label
            row["parent"] = page_name
            row["page"] = page_name
            row["visual_type"] = v_type
            row["visual_title"] = v_title
            rows.append(row)

            # visual_field rows
            for f in visual.get("fields", []):
                field_str = f.get("field", "")
                role_str = f.get("role", "")

                row = _blank_row()
                row["object_type"] = "visual_field"
                row["name"] = field_str
                row["parent"] = v_label
                row["page"] = page_name
                row["visual_type"] = v_type
                row["visual_title"] = v_title
                row["field"] = field_str
                row["role"] = role_str
                rows.append(row)

    return rows


# ---------------------------------------------------------------------------
# CSV writer
# ---------------------------------------------------------------------------

def write_csv(rows: list, output_path: Path, delimiter: str) -> None:
    """Write all rows to a UTF-8-BOM CSV with QUOTE_ALL so multiline cells are safe."""
    with open(output_path, "w", newline="", encoding="utf-8-sig") as fh:
        writer = csv.DictWriter(
            fh,
            fieldnames=CSV_COLUMNS,
            delimiter=delimiter,
            quoting=csv.QUOTE_ALL,
        )
        writer.writeheader()
        writer.writerows(rows)


# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------

def main() -> None:
    ap = argparse.ArgumentParser(
        description="Extract everything from a .pbit file into a single long-format CSV.",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog=(
            "Examples:\n"
            "  python pbit_extract.py mon_rapport.pbit --output extraction.csv\n"
            "  python pbit_extract.py mon_rapport.pbit --delimiter , --no-sanitize\n"
        ),
    )
    ap.add_argument("pbit_file", help="Path to the .pbit file")
    ap.add_argument(
        "--output", "-o",
        default="",
        help="Output CSV path (default: <stem>_extraction.csv next to the input file)",
    )
    ap.add_argument(
        "--delimiter", "-d",
        default=";",
        help="CSV column delimiter (default: ;)",
    )
    ap.add_argument(
        "--no-sanitize",
        action="store_true",
        help=(
            "Disable the sanitizer: write raw connection strings, SharePoint URLs, "
            "and Fabric endpoint IDs as-is (not recommended for shared files)."
        ),
    )
    args = ap.parse_args()

    input_path = Path(args.pbit_file)
    if not input_path.is_absolute():
        input_path = Path.cwd() / args.pbit_file.lstrip("/\\")

    if input_path.suffix.lower() != ".pbit":
        print(f"ERROR: Expected a .pbit file, got: {input_path}", file=sys.stderr)
        sys.exit(1)

    do_sanitize = not args.no_sanitize

    print(f"Opening: {input_path}")
    parser, report_data = open_pbit(input_path)

    tables = parser.tables()
    rels = parser.relationships()
    roles = parser.roles()
    shared = parser.shared_expressions()

    n_measures = sum(len(parser.get_measures(t)) for t in tables)
    n_columns = sum(len(parser.get_columns(t)) for t in tables)

    print(f"  Tables:        {len(tables)}")
    print(f"  Columns:       {n_columns}")
    print(f"  Measures:      {n_measures}")
    print(f"  Relationships: {len(rels)}")
    print(f"  RLS Roles:     {len(roles)}")
    print(f"  Shared exprs:  {len(shared)}")
    if report_data:
        n_pages = len(report_data.get("pages", []))
        n_visuals = sum(len(p.get("visuals", [])) for p in report_data.get("pages", []))
        print(f"  Pages:         {n_pages}")
        print(f"  Visuals:       {n_visuals}")

    # Build all CSV rows
    all_rows: list = []

    # Pass all_table_names for Fix #3 DAX dependency analysis
    all_table_names = [t.get("name", "") for t in tables]

    for table in tables:
        all_rows.extend(_rows_for_table(table, parser, do_sanitize, all_table_names))

    all_rows.extend(_rows_for_shared_expressions(parser, do_sanitize))
    all_rows.extend(_rows_for_relationships(parser))
    all_rows.extend(_rows_for_roles(parser))

    if report_data:
        all_rows.extend(_rows_for_report(report_data))

    # Determine output path
    if args.output:
        output_path = Path(args.output)
    else:
        output_path = input_path.with_name(input_path.stem + "_extraction.csv")

    write_csv(all_rows, output_path, args.delimiter)

    print(f"\nDone. {len(all_rows)} rows written to: {output_path.resolve()}")
    if do_sanitize:
        print("  (Sanitizer active: org-specific connection details have been masked.)")
    else:
        print("  (Sanitizer disabled: raw connection strings are included.)")


if __name__ == "__main__":
    main()
