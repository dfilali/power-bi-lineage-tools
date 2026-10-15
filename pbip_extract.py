#!/usr/bin/env python3
"""
pbip_document.py — Power BI Project (.pbip) Documentation Generator

Supports both TMSL (.bim) and TMDL (.tmdl) semantic model formats.

Output modes:
  Default:   Markdown documentation (human-readable reference)
  --copilot: Plain-text knowledge base optimised for Copilot / ChatGPT context injection.
             Describes every field, measure, and relationship in prose so an LLM can
             answer questions like "how is measure X built?" or "what does column Y mean?"

Usage:
    python pbip_document.py <path-to-project>.pbip [--output docs.md]
    python pbip_document.py <path-to-project>.pbip --copilot [--output knowledge.txt]

Requirements: Python 3.8+ (stdlib only, no pip installs needed)
"""

import argparse
import json
import os
import re
import sys
import urllib.parse
from collections import Counter
from pathlib import Path
from typing import Optional


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------

def read_json(path: Path) -> dict:
    with open(path, encoding="utf-8-sig") as f:
        return json.load(f)


def slugify(name: str) -> str:
    return re.sub(r"[^\w\-]", "-", name).lower()


# ---------------------------------------------------------------------------
# Sanitizer — maskeert organisatie-specifieke connectiestrings vóór output
# ---------------------------------------------------------------------------

_SHAREPOINT_URL_RE = re.compile(
    r'https?://[a-zA-Z0-9\-]+\.sharepoint\.com/sites/[^"\')\s,\]]+',
    re.IGNORECASE,
)

# Bestandsnamen die intern zijn maar niet gevoelig (bijv. swt_producten.xlsx) laten we staan.
# Alleen de volledige URL wordt gemaskeerd.

# Fabric SQL analytics endpoint hostnames, e.g. abc123xyz456.datawarehouse.fabric.microsoft.com
# — masked regardless of context (named parameter, or a bare literal passed straight into
# Sql.Database(...)), since the hostname itself already identifies the tenant/workspace.
_FABRIC_SQL_ENDPOINT_RE = re.compile(
    r'\b[a-z0-9]{6,}\.datawarehouse\.fabric\.microsoft\.com\b',
    re.IGNORECASE,
)

# key = value / key: value / key = "value" assignments for well-known Fabric- and
# SharePoint/OneDrive-connector identifiers that show up in M-query "let" steps, TMDL
# parameter definitions, or navigation records (e.g. Source{[workspaceId="...", ...]}).
# Extend this dict if other infra-identifying key names turn up.
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


# The value itself is restricted to identifier-ish characters (letters, digits,
# underscore, dot, hyphen — covers hostnames and GUIDs) with NO spaces or quotes
# allowed. This matters: a greedy/lazy "anything but a few delimiters" value class
# would, for a key name that also happens to be an ordinary business/column name
# (e.g. WarehouseId inside a passthrough SQL string like
# [Query="SELECT * FROM T WHERE WarehouseId = 5 AND Status = 'Open'"]), keep
# consuming through spaces and text until the next delimiter — swallowing the rest
# of the SQL and corrupting the output. Real infra values never contain spaces, so
# stopping at the first non-identifier character is both safe and sufficient.
_INFRA_KV_RE = re.compile(
    r'\b(' + "|".join(re.escape(k) for k in _INFRA_KEY_PLACEHOLDERS) + r')'
    r'(\s*[:=]\s*)'
    r'"?([A-Za-z0-9_.\-]+)"?',
    re.IGNORECASE,
)

# "schema" is masked separately and more narrowly: it's a common, legitimate, benign
# record field in classic SQL-connector navigation (Source{[Schema="dbo",Item="X"]}),
# where masking it would remove useful, non-sensitive documentation. It's only masked
# as a standalone parameter/let-step assignment that starts its own line
# (e.g. "    schema = "brons_lh","), never as a field inside a "[...]" record literal —
# those are always written inline (Schema="dbo",Item=...) rather than starting a line.
# Same identifier-only value class as _INFRA_KV_RE, for the same over-matching reason.
_SCHEMA_ASSIGN_RE = re.compile(
    r'(?im)^(\s*)schema(\s*[:=]\s*)"?([A-Za-z0-9_.\-]+)"?',
)


def sanitize(text: str) -> str:
    """
    Maskeert organisatie-specifieke infrastructuurdetails in Power Query / DAX
    expressies vóór output: SharePoint-URLs, Fabric SQL-endpoints, en bekende
    Fabric/SharePoint object-ID's (workspaceId, lakehouseId, driveId, ...).

    De query-structuur zelf (functienamen, transformatiestappen, kolomverwijzingen)
    blijft altijd intact — alleen de organisatie-specifieke waarden worden vervangen
    door een placeholder, zodat het model te begrijpen blijft zonder te lekken op
    welke tenant/workspace/resource het draait.

    Wat er o.a. gebeurt:
      https://organisatie.sharepoint.com/sites/SITENAAM/Gedeelde%20documenten/map/bestand.xlsx
      → https://[TENANT].sharepoint.com/sites/[SITE]/Gedeelde documenten/map/bestand.xlsx
      abc123xyz456.datawarehouse.fabric.microsoft.com → [FABRIC_SQL_ENDPOINT]
      workspaceId = "12345678-...."                   → workspaceId = [WORKSPACE_ID]
      schema = "brons_lh"                              → schema = [SCHEMA]
    """
    if not text:
        return text

    def _replace_sharepoint(m: re.Match) -> str:
        url = m.group(0)
        try:
            parsed = urllib.parse.urlparse(url)
            # Decodeer %20 etc. zodat het pad leesbaar is
            path_decoded = urllib.parse.unquote(parsed.path)
            parts = path_decoded.strip("/").split("/")
            # parts[0] = "sites", parts[1] = sitenaam, parts[2:] = rest van het pad
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
# TMSL parser (.bim format)
# ---------------------------------------------------------------------------

class TMSLParser:
    def __init__(self, bim_path: Path):
        self.data = read_json(bim_path)

    def _model(self) -> dict:
        return self.data.get("model", self.data)

    def tables(self) -> list[dict]:
        return [t for t in self._model().get("tables", []) if not t.get("isHidden", False)]

    def relationships(self) -> list[dict]:
        return self._model().get("relationships", [])

    def roles(self) -> list[dict]:
        return self._model().get("roles", [])

    def get_measures(self, table: dict) -> list[dict]:
        return table.get("measures", [])

    def get_columns(self, table: dict) -> list[dict]:
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

    def get_partitions(self, table: dict) -> list[dict]:
        return table.get("partitions", [])

    def shared_expressions(self) -> list[dict]:
        """Shared M functions from the model's top-level `expressions` array — parameters
        are excluded on purpose, since they often carry organisation-specific default
        values (servers, environments)."""
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


# ---------------------------------------------------------------------------
# TMDL parser (.tmdl format — Tabular Model Definition Language)
# ---------------------------------------------------------------------------

class TMLDParser:
    def __init__(self, definition_dir: Path):
        self.definition_dir = definition_dir
        self._tables: list[dict] = []
        self._relationships: list[dict] = []
        self._roles: list[dict] = []
        self._parse_all()

    def _parse_all(self):
        tables_dir = self.definition_dir / "tables"
        if tables_dir.exists():
            for tmdl_file in sorted(tables_dir.glob("*.tmdl")):
                table = self._parse_table_file(tmdl_file)
                if table and not table.get("isHidden", False):
                    self._tables.append(table)

        rel_file = self.definition_dir / "relationships.tmdl"
        if rel_file.exists():
            self._relationships = self._parse_relationships(rel_file)

        roles_dir = self.definition_dir / "roles"
        if roles_dir.exists():
            for role_file in sorted(roles_dir.glob("*.tmdl")):
                role = self._parse_role_file(role_file)
                if role:
                    self._roles.append(role)

        expr_file = self.definition_dir / "expressions.tmdl"
        if expr_file.exists():
            self._shared_expressions = self._parse_expressions_tmdl(expr_file)
        else:
            self._shared_expressions = []

    def _lines(self, path: Path) -> list[str]:
        return path.read_text(encoding="utf-8-sig").splitlines()

    def _parse_table_file(self, path: Path) -> Optional[dict]:
        lines = self._lines(path)
        if not lines:
            return None

        table: dict = {"name": "", "isHidden": False, "columns": [], "measures": [], "partitions": []}
        i = 0

        while i < len(lines) and not lines[i].strip():
            i += 1
        if i >= len(lines):
            return None

        header = lines[i].strip()
        if header.startswith("table "):
            # TMDL quotes names that contain spaces/special characters, e.g.
            # table 'Beoordelingen - laatste 4 jaar' — strip that quoting so
            # the logical name doesn't end up with literal quote characters in it.
            table["name"] = header[6:].strip().strip("'")
        else:
            table["name"] = path.stem
        i += 1

        current_block: list[str] = []
        current_type: str = ""
        current_name: str = ""

        def flush_block():
            if current_type == "measure":
                table["measures"].append(self._parse_measure_block(current_name, current_block))
            elif current_type == "column":
                col = self._parse_column_block(current_name, current_block)
                if col:
                    table["columns"].append(col)
            elif current_type == "partition":
                table["partitions"].append(self._parse_partition_block(current_name, current_block))

        while i < len(lines):
            line = lines[i]
            stripped = line.strip()

            if stripped.startswith("measure ") and line.startswith("\t") and not line.startswith("\t\t"):
                flush_block()
                current_type = "measure"
                raw_name = stripped[8:].split("=")[0].strip().strip("'")
                current_name = raw_name
                current_block = [line]
            elif stripped.startswith("column ") and line.startswith("\t") and not line.startswith("\t\t"):
                flush_block()
                current_type = "column"
                # A calculated column's header is "column Name = <DAX expression>" —
                # split off the "= ..." part so the name doesn't include the formula.
                current_name = stripped[7:].split("=")[0].strip().strip("'")
                current_block = [line]
            elif stripped.startswith("partition ") and line.startswith("\t") and not line.startswith("\t\t"):
                flush_block()
                current_type = "partition"
                current_name = stripped[10:].split("=")[0].strip().strip("'")
                current_block = [line]
            elif stripped.startswith("isHidden:") and line.startswith("\t") and not line.startswith("\t\t"):
                table["isHidden"] = "true" in stripped.lower()
            elif current_type:
                current_block.append(line)
            i += 1

        flush_block()
        return table

    def _parse_measure_block(self, name: str, block: list[str]) -> dict:
        dax_lines: list[str] = []
        description = ""
        format_string = ""

        SIBLING_PROPS = (
            "formatString:", "lineageTag:", "description:", "displayFolder:",
            "annotation ", "isHidden:", "kpiStatusType:", "dataCategory:",
        )
        # TMDL's bare-boolean shorthand: a lone property name on its own line means
        # "= true" (isHidden ⇔ isHidden: true) — also a sibling-property boundary.
        SIBLING_PROPS_BARE = {"isHidden"}

        in_dax = False

        for line in block:
            stripped = line.strip()

            if stripped.startswith("measure ") and "=" in stripped:
                parts = stripped.split("=", 1)
                inline = parts[1].strip() if len(parts) == 2 else ""
                if inline:
                    dax_lines.append(inline)
                in_dax = True
                continue

            if stripped and (any(stripped.startswith(p) for p in SIBLING_PROPS) or stripped in SIBLING_PROPS_BARE):
                in_dax = False
                if stripped.startswith("formatString:"):
                    format_string = stripped.split(":", 1)[1].strip().strip("'\"")
                elif stripped.startswith("description:"):
                    description = stripped.split(":", 1)[1].strip().strip("'\"")
                continue

            if in_dax:
                clean = stripped.strip("`").strip()
                dax_lines.append(clean if clean else "")

        while dax_lines and not dax_lines[0]:
            dax_lines.pop(0)
        while dax_lines and not dax_lines[-1]:
            dax_lines.pop()

        return {
            "name": name,
            "expression": "\n".join(dax_lines).strip(),
            "description": description,
            "formatString": format_string,
        }

    def _parse_column_block(self, name: str, block: list[str]) -> Optional[dict]:
        col: dict = {
            "name": name, "dataType": "unknown", "isHidden": False,
            "description": "", "expression": "", "isCalculated": False,
        }

        # Sibling properties that can follow a column header — anything else while
        # in_dax is still open is treated as (part of) a calculated column's DAX formula.
        SIBLING_PROPS = (
            "dataType:", "lineageTag:", "description:", "displayFolder:",
            "annotation ", "isHidden:", "formatString:", "summarizeBy:",
            "sourceColumn:", "isDataTypeInferred:", "sortByColumn:",
            "dataCategory:", "isNameInferred:", "isKey:", "isUnique:",
            "isAvailableInMdx:", "keepUniqueRows:", "isNullable:",
            "encoding:", "state:", "changedProperty:", "sourceProviderType:",
            "alignment:", "formatStringDefinition:",
        )
        # TMDL's bare-boolean shorthand: a lone property name on its own line means
        # "= true" (isHidden ⇔ isHidden: true) — also a sibling-property boundary.
        SIBLING_PROPS_BARE = {
            "isHidden", "isKey", "isUnique", "isDataTypeInferred", "isNameInferred",
            "isAvailableInMdx", "keepUniqueRows", "isNullable",
        }

        dax_lines: list[str] = []
        in_dax = False

        for line in block:
            stripped = line.strip()

            if stripped.startswith("column ") and "=" in stripped:
                inline = stripped.split("=", 1)[1].strip()
                if inline:
                    dax_lines.append(inline)
                in_dax = True
                continue

            if stripped and (any(stripped.startswith(p) for p in SIBLING_PROPS) or stripped in SIBLING_PROPS_BARE):
                in_dax = False
                if stripped.startswith("dataType:"):
                    col["dataType"] = stripped.split(":", 1)[1].strip()
                elif stripped.startswith("isHidden:"):
                    col["isHidden"] = "true" in stripped.lower()
                elif stripped == "isHidden":
                    col["isHidden"] = True
                elif stripped.startswith("description:"):
                    col["description"] = stripped.split(":", 1)[1].strip().strip("'\"")
                continue

            if in_dax:
                clean = stripped.strip("`").strip()
                dax_lines.append(clean if clean else "")

        while dax_lines and not dax_lines[0]:
            dax_lines.pop(0)
        while dax_lines and not dax_lines[-1]:
            dax_lines.pop()

        if dax_lines:
            col["expression"] = "\n".join(dax_lines).strip()
            col["isCalculated"] = True

        return col

    def _parse_partition_block(self, name: str, block: list[str]) -> dict:
        partition: dict = {"name": name, "mode": "import", "source": {}}
        m_lines: list[str] = []
        in_source = False
        # None = no fence seen yet (or none used); True = opening ``` has been seen
        # and we're waiting for the matching closing ```. TMDL allows the opening
        # fence either attached to "source =" on the same line, or on its own line
        # right after — both must be recognised as *opening*, not closing.
        fenced = False

        for line in block:
            stripped = line.strip()
            if stripped.startswith("mode:"):
                partition["mode"] = stripped.split(":", 1)[1].strip()
            elif stripped == "source =" or stripped.startswith("source ="):
                in_source = True
                fenced = False
                rest = stripped[len("source ="):].strip()
                if rest == "```":
                    # opening fence attached to the "source =" line itself
                    fenced = True
                    rest = ""
                else:
                    rest = rest.strip("`")
                if rest:
                    m_lines.append(rest)
            elif in_source:
                if stripped == "```":
                    if fenced:
                        in_source = False  # closing fence
                    else:
                        fenced = True  # opening fence on its own line
                    continue
                if not fenced:
                    sibling_keywords = {"mode:", "annotation", "description:", "sourceColumn:", "isHidden:"}
                    if any(stripped.startswith(k) for k in sibling_keywords):
                        in_source = False
                        continue
                # Keep blank lines so multi-step M queries keep their original
                # spacing instead of being collapsed onto consecutive lines.
                m_lines.append(stripped)

        if m_lines:
            # Trim leading/trailing blank lines only, keep interior spacing intact.
            while m_lines and not m_lines[0]:
                m_lines.pop(0)
            while m_lines and not m_lines[-1]:
                m_lines.pop()
            partition["source"] = {"expression": "\n".join(m_lines).strip()}

        return partition

    def _parse_expressions_tmdl(self, path: Path) -> list[dict]:
        lines = self._lines(path)
        expressions: list[dict] = []
        current_name = ""
        current_kind = ""
        m_lines: list[str] = []
        in_expr = False

        def flush():
            if current_name and m_lines:
                while m_lines and not m_lines[0]:
                    m_lines.pop(0)
                while m_lines and not m_lines[-1]:
                    m_lines.pop()
                expr_text = "\n".join(m_lines).strip()
                expressions.append({
                    "name": current_name,
                    "kind": current_kind,
                    "expression": expr_text,
                    "is_parameter": bool(re.search(r"IsParameterQuery\s*=\s*true", expr_text, re.IGNORECASE)),
                })

        for line in lines:
            stripped = line.strip()

            if stripped.startswith("expression "):
                flush()
                name_part = stripped[len("expression "):].split("=")[0].strip().strip("'\"")
                current_name = name_part
                current_kind = "expression"
                m_lines = []
                in_expr = True
                if "=" in stripped:
                    rest = stripped.split("=", 1)[1].strip().strip("`")
                    if rest:
                        m_lines.append(rest)
            elif stripped.startswith("kind:"):
                current_kind = stripped.split(":", 1)[1].strip()
            elif in_expr and not stripped.startswith("kind:") and not stripped.startswith("lineageTag:") and not stripped.startswith("annotation"):
                m_lines.append(line.rstrip().strip("`"))
            elif stripped.startswith("lineageTag:") or stripped.startswith("annotation"):
                in_expr = False

        flush()

        return expressions

    def shared_expressions(self) -> list[dict]:
        """Shared M functions from expressions.tmdl — parameters are excluded on purpose,
        since they often carry organisation-specific default values (servers, environments)."""
        return [e for e in getattr(self, "_shared_expressions", []) if not e.get("is_parameter")]

    def _parse_relationships(self, path: Path) -> list[dict]:
        lines = self._lines(path)
        rels = []
        current: dict = {}

        def split_table_column(value: str):
            value = value.strip()
            m = re.match(r"^'([^']+)'\.'([^']+)'$", value)
            if m:
                return m.group(1), m.group(2)
            m = re.match(r"^'([^']+)'\.([^\s]+)$", value)
            if m:
                return m.group(1), m.group(2)
            m = re.match(r"^([^'.]+)\.'([^']+)'$", value)
            if m:
                return m.group(1), m.group(2)
            if "." in value:
                parts = value.split(".", 1)
                return parts[0], parts[1]
            return value, ""

        for line in lines:
            stripped = line.strip()
            if stripped.startswith("relationship "):
                if current:
                    rels.append(current)
                current = {"name": stripped[13:].strip()}
            elif stripped.startswith("fromTable:"):
                current["fromTable"] = stripped.split(":", 1)[1].strip()
            elif stripped.startswith("fromColumn:"):
                raw = stripped.split(":", 1)[1].strip()
                if "." in raw:
                    current["fromTable"], current["fromColumn"] = split_table_column(raw)
                else:
                    current["fromColumn"] = raw
            elif stripped.startswith("toTable:"):
                current["toTable"] = stripped.split(":", 1)[1].strip()
            elif stripped.startswith("toColumn:"):
                raw = stripped.split(":", 1)[1].strip()
                if "." in raw:
                    current["toTable"], current["toColumn"] = split_table_column(raw)
                else:
                    current["toColumn"] = raw
            elif stripped.startswith("crossFilteringBehavior:"):
                current["crossFilter"] = stripped.split(":", 1)[1].strip()
            elif stripped.startswith("cardinality:") or stripped.startswith("fromCardinality:"):
                current["cardinality"] = stripped.split(":", 1)[1].strip()

        if current:
            rels.append(current)
        return rels

    def _parse_role_file(self, path: Path) -> Optional[dict]:
        lines = self._lines(path)
        role: dict = {"name": path.stem, "tablePermissions": []}

        for line in lines:
            stripped = line.strip()
            if stripped.startswith("role "):
                role["name"] = stripped[5:].strip()
            elif stripped.startswith("tablePermission "):
                table_name = stripped[16:].strip()
                role["tablePermissions"].append({"table": table_name, "filterExpression": ""})

        return role

    def tables(self) -> list[dict]:
        return self._tables

    def relationships(self) -> list[dict]:
        return self._relationships

    def roles(self) -> list[dict]:
        return self._roles

    def get_measures(self, table: dict) -> list[dict]:
        return table.get("measures", [])

    def get_columns(self, table: dict) -> list[dict]:
        return table.get("columns", [])

    def get_partitions(self, table: dict) -> list[dict]:
        return table.get("partitions", [])


# ---------------------------------------------------------------------------
# Report parser (report.json)
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
        expr = agg.get("Expression", {})
        inner = _extract_field_name(expr)
        return f"{func}({inner})" if inner else func

    if "HierarchyLevel" in field:
        hl = field["HierarchyLevel"]
        expr = hl.get("Expression", {})
        hier_expr = expr.get("Hierarchy", {})
        entity = hier_expr.get("Expression", {}).get("SourceRef", {}).get("Entity", "")
        hier_name = hier_expr.get("Hierarchy", "")
        level = hl.get("Level", "")
        parts = [p for p in [entity, hier_name, level] if p]
        return ".".join(parts)

    if "SourceRef" in field:
        return field["SourceRef"].get("Entity", "")

    return None


def _parse_visual_config(config: dict) -> dict:
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

    fields: list[dict] = []

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
        dt_lookup: dict[str, dict] = {}
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

    return {
        "visualType": visual_type,
        "title": title,
        "fields": fields,
    }


def parse_report(report_json_path: Path) -> dict:
    data = read_json(report_json_path)
    result = {"pages": []}

    for section in data.get("sections", []):
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


def parse_report_definition(definition_dir: Path) -> dict:
    result = {"pages": []}
    pages_dir = definition_dir / "pages"

    if not pages_dir.exists():
        return result

    page_order: list[str] = []
    pages_meta_file = pages_dir / "pages.json"
    if pages_meta_file.exists():
        try:
            pages_meta = read_json(pages_meta_file)
            page_order = pages_meta.get("pageOrder", [])
        except Exception:
            pass

    all_page_dirs = {d.name: d for d in pages_dir.iterdir() if d.is_dir()}
    ordered_hashes = page_order + [h for h in sorted(all_page_dirs) if h not in page_order]

    for page_hash in ordered_hashes:
        page_dir = all_page_dirs.get(page_hash)
        if not page_dir:
            continue

        page_json = page_dir / "page.json"
        display_name = page_hash
        if page_json.exists():
            try:
                page_meta = read_json(page_json)
                display_name = page_meta.get("displayName", page_hash)
            except Exception:
                pass

        page = {"name": display_name, "visuals": []}
        visuals_dir = page_dir / "visuals"

        if visuals_dir.exists():
            for visual_dir in sorted(visuals_dir.iterdir()):
                if not visual_dir.is_dir():
                    continue
                visual_json = visual_dir / "visual.json"
                if not visual_json.exists():
                    continue

                try:
                    visual_data = read_json(visual_json)
                except Exception:
                    continue

                visual_info = _parse_visual_config(visual_data)

                vtype = visual_info["visualType"]
                if vtype in ("", "unknown", "shape", "image", "textbox"):
                    if not visual_info["fields"]:
                        continue

                page["visuals"].append(visual_info)

        result["pages"].append(page)

    return result


# ---------------------------------------------------------------------------
# Report structure (pages & visuals) — shared between the full renderers and
# the report-only renderers used when a report has no local semantic model
# (live connection to a remote/published model).
# ---------------------------------------------------------------------------

def _report_structure_markdown_lines(report_data: dict) -> list[str]:
    lines: list[str] = []
    def p(t=""): lines.append(t)
    def h3(t): lines.extend([f"### {t}", ""])

    pages = report_data.get("pages", [])
    p(f"Total pages: **{len(pages)}**")
    p()
    for page in pages:
        h3(page["name"])
        visuals = page.get("visuals", [])
        if not visuals:
            p("_No visual information extracted._")
            p()
            continue

        type_counts = Counter(v["visualType"] for v in visuals)
        p(f"**{len(visuals)} visual(s):** " + ", ".join(f"{cnt}× {vt}" for vt, cnt in sorted(type_counts.items())))
        p()

        for i, visual in enumerate(visuals, 1):
            vtype = visual["visualType"]
            title = visual.get("title", "")
            fields = visual.get("fields", [])

            header = f"**Visual {i}: `{vtype}`**"
            if title:
                header += f" — _{title}_"
            p(header)

            if fields:
                p("| Role | Field | Display Name |")
                p("|------|-------|--------------|")
                for f in fields:
                    role = f.get("role", "")
                    field = f.get("field", "")
                    display = f.get("displayName", "")
                    p(f"| {role} | `{field}` | {display} |")
            else:
                p("_No field bindings extracted._")
            p()

    return lines


def _report_structure_copilot_lines(report_data: dict) -> list[str]:
    lines: list[str] = []
    def w(s=""): lines.append(s)

    pages = report_data.get("pages", [])
    w(f"The report has {len(pages)} page(s).")
    w()
    for page in pages:
        page_name = page["name"]
        visuals = page.get("visuals", [])
        w(f"Page: {page_name} ({len(visuals)} visual(s))")
        w("-" * 60)

        if not visuals:
            w("  No visual information available.")
            w()
            continue

        for i, visual in enumerate(visuals, 1):
            vtype = visual["visualType"]
            title = visual.get("title", "")
            fields = visual.get("fields", [])

            label = f'Visual {i}: {vtype}'
            if title:
                label += f' (title: "{title}")'
            w(f"  {label}")

            if fields:
                for f in fields:
                    role = f.get("role", "")
                    field = f.get("field", "")
                    display = f.get("displayName", "")
                    role_str = f" [{role}]" if role else ""
                    display_str = f' (shown as "{display}")' if display and display != field else ""
                    w(f"    - {field}{role_str}{display_str}")
            else:
                w("    No field bindings extracted.")
            w()

    return lines


def render_report_only_markdown(project_name: str, report_data: Optional[dict]) -> str:
    """Report-only documentation for a report bound to a live connection to a
    remote/published semantic model: no local model to document, so this is
    just the pages/visuals overview."""
    lines: list[str] = []
    def h1(t): lines.extend([f"# {t}", ""])
    def h2(t): lines.extend([f"## {t}", ""])
    def p(t=""): lines.append(t)

    h1(f"Power BI Documentation: {project_name}")
    p("_This report uses a live connection to a remote/published semantic model — "
      "there is no local semantic model in this project to document. Only the "
      "report structure (pages and visuals) below was extracted._")
    p()

    h2("Report Structure")
    if report_data and report_data.get("pages"):
        lines.extend(_report_structure_markdown_lines(report_data))
    else:
        p("_No report structure could be extracted._")
        p()

    return "\n".join(lines)


def render_report_only_copilot_kb(project_name: str, report_data: Optional[dict]) -> str:
    """Copilot-mode counterpart to render_report_only_markdown."""
    lines: list[str] = []
    def w(s=""): lines.append(s)

    w("=" * 80)
    w(f"POWER BI REPORT OVERVIEW: {project_name.upper()}")
    w("=" * 80)
    w()
    w("This report uses a live connection to a remote/published semantic model.")
    w("There is no local semantic model in this project to document, so this file")
    w("only lists the report's pages and visuals (and the fields each visual uses).")
    w()

    if report_data and report_data.get("pages"):
        lines.extend(_report_structure_copilot_lines(report_data))
    else:
        w("No report structure could be extracted.")
        w()

    return "\n".join(lines)


# ---------------------------------------------------------------------------
# Markdown renderer
# ---------------------------------------------------------------------------

def render_markdown(
    project_name: str,
    parser,
    report_data: Optional[dict],
    model_format: str,
    row_counts: Optional[dict[str, int]] = None,
) -> str:
    row_counts = row_counts or {}
    lines: list[str] = []

    def h1(t): lines.extend([f"# {t}", ""])
    def h2(t): lines.extend([f"## {t}", ""])
    def h3(t): lines.extend([f"### {t}", ""])
    def p(t=""): lines.append(t)

    h1(f"Power BI Documentation: {project_name}")
    p(f"_Semantic model format: **{model_format}**_")
    p()

    h2("Table of Contents")
    toc_entries = [
        "[Data Model Overview](#data-model-overview)",
        "[Tables & Columns](#tables--columns)",
        "[Measure Catalogue](#measure-catalogue)",
        "[Relationships](#relationships)",
    ]
    if parser.roles():
        toc_entries.append("[Row-Level Security](#row-level-security)")
    if report_data:
        toc_entries.append("[Report Structure](#report-structure)")
    toc_entries.append("[Data Lineage & Impact Analysis](#data-lineage--impact-analysis)")
    for i, entry in enumerate(toc_entries, 1):
        p(f"{i}. {entry}")
    p()

    h2("Data Model Overview")
    tables = parser.tables()
    all_table_names = [t.get("name", "") for t in tables]
    all_measure_names = [m.get("name", "") for t in tables for m in parser.get_measures(t)]
    total_measures = sum(len(parser.get_measures(t)) for t in tables)
    total_columns = sum(len(parser.get_columns(t)) for t in tables)
    total_rels = len(parser.relationships())

    p(f"| Metric | Count |")
    p(f"|--------|-------|")
    p(f"| Tables | {len(tables)} |")
    p(f"| Columns | {total_columns} |")
    p(f"| Measures | {total_measures} |")
    p(f"| Relationships | {total_rels} |")
    p(f"| RLS Roles | {len(parser.roles())} |")
    if row_counts:
        known_total = sum(row_counts.get(t.get("name", ""), 0) for t in tables if t.get("name", "") in row_counts)
        p(f"| Total rows (known tables) | {known_total:,} |")
    p()

    if row_counts:
        p("**Table sizes** _(pasted from a manually run DAX row-count query — see "
          "`--rowcounts`; not necessarily current)_:")
        p()
        p("| Table | Rows |")
        p("|-------|------|")
        for table in tables:
            name = table.get("name", "")
            count = row_counts.get(name)
            p(f"| {name} | {count:,} |" if count is not None else f"| {name} | _unknown_ |")
        p()

    h2("Tables & Columns")
    for table in tables:
        table_name = table.get("name", "Unknown")
        h3(table_name)

        columns = parser.get_columns(table)
        measures = parser.get_measures(table)
        partitions = parser.get_partitions(table)

        if table_name in row_counts:
            p(f"**Rows:** {row_counts[table_name]:,}")
            p()

        if partitions:
            for pt in partitions:
                source = pt.get("source", {})
                if isinstance(source, dict) and source.get("query", source.get("expression", "")):
                    mode = pt.get("mode", "import")
                    p(f"**Source:** Power Query (`{mode}`) — full expression under "
                      f"[Power Query (M) Sources](#power-query-m-sources).")
                    p()
                    break

        if columns:
            p("| Column | Data Type | Calculated | Hidden | Description |")
            p("|--------|-----------|------------|--------|-------------|")
            for col in columns:
                name = col.get("name", "")
                dtype = col.get("dataType", col.get("type", "unknown"))
                calc = "✓" if col.get("isCalculated") else ""
                hidden = "✓" if col.get("isHidden") else ""
                desc = col.get("description", "")
                p(f"| {name} | {dtype} | {calc} | {hidden} | {desc} |")
            p()

            calc_columns = [c for c in columns if c.get("isCalculated") and c.get("expression")]
            if calc_columns:
                p("**Calculated column formulas:**")
                p()
                for col in calc_columns:
                    p(f"#### `{col.get('name', '')}` (calculated column)")
                    p("```dax")
                    p(sanitize(col["expression"]))
                    p("```")
                    p()
        else:
            p("_No columns defined (may be a calculated or virtual table)._")
            p()

        if measures:
            p(f"**{len(measures)} measure(s)** — see Measure Catalogue below.")
            p()

    h2("Measure Catalogue")
    has_measures = False
    for table in tables:
        measures = parser.get_measures(table)
        if not measures:
            continue
        has_measures = True
        table_name = table.get("name", "Unknown")
        h3(f"Table: {table_name}")

        for measure in measures:
            name = measure.get("name", "")
            dax = measure.get("expression", "").strip()
            desc = measure.get("description", "")
            fmt = measure.get("formatString", "")

            p(f"#### `{name}`")
            if desc:
                p(f"_{desc}_")
                p()
            if fmt:
                p(f"**Format:** `{fmt}`")
                p()
            if dax:
                p("```dax")
                # DAX bevat normaal geen URLs, maar voor de zekerheid
                p(sanitize(dax))
                p("```")
            else:
                p("_No DAX expression found._")
            p()

    if not has_measures:
        p("_No measures found in this model._")
        p()

    h2("Relationships")
    rels = parser.relationships()
    if rels:
        p("| From Table | From Column | To Table | To Column | Cardinality | Cross Filter |")
        p("|------------|-------------|----------|-----------|-------------|--------------|")
        for rel in rels:
            from_t = rel.get("fromTable", rel.get("fromTableId", ""))
            from_c = rel.get("fromColumn", rel.get("fromColumnId", ""))
            to_t = rel.get("toTable", rel.get("toTableId", ""))
            to_c = rel.get("toColumn", rel.get("toColumnId", ""))
            card = rel.get("cardinality", "many-to-one")
            cf = rel.get("crossFilter", rel.get("crossFilteringBehavior", ""))
            p(f"| {from_t} | {from_c} | {to_t} | {to_c} | {card} | {cf} |")
        p()
    else:
        p("_No relationships defined._")
        p()

    roles = parser.roles()
    if roles:
        h2("Row-Level Security")
        for role in roles:
            role_name = role.get("name", "")
            h3(role_name)
            members = role.get("members", [])
            if members:
                p(f"**Members:** {', '.join(str(m) for m in members)}")
                p()
            perms = role.get("tablePermissions", [])
            model_perms = role.get("modelPermission", "")
            if model_perms:
                p(f"**Model permission:** `{model_perms}`")
                p()
            if perms:
                p("| Table | Filter Expression |")
                p("|-------|------------------|")
                for perm in perms:
                    t = perm.get("table", perm.get("name", ""))
                    expr = perm.get("filterExpression", perm.get("expression", "_none_"))
                    p(f"| {t} | `{expr}` |")
                p()

    if report_data:
        h2("Report Structure")
        lines.extend(_report_structure_markdown_lines(report_data))

    h2("Power Query (M) Sources")
    has_pq = False
    for table in tables:
        partitions = parser.get_partitions(table)
        for pt in partitions:
            source = pt.get("source", {})
            expr = source.get("expression", "") if isinstance(source, dict) else source
            if not expr:
                continue
            has_pq = True
            table_name = table.get("name", "Unknown")
            h3(table_name)
            p(f"**Mode:** `{pt.get('mode', 'import')}`")
            p()
            p("```powerquery")
            # ↓ sanitize vóór output
            p(sanitize(expr.strip()))
            p("```")
            p()

    if not has_pq:
        p("_No Power Query expressions found (model may use DirectQuery or live connection)._")
        p()

    shared = parser.shared_expressions()
    if shared:
        h2("Shared Power Query Functions")
        p("These are reusable queries/functions defined in `expressions.tmdl`, "
          "typically used as helper functions called from table queries above. "
          "Query parameters are intentionally excluded from this documentation.")
        p()
        for expr in shared:
            name = expr.get("name", "")
            kind = expr.get("kind", "")
            body = expr.get("expression", "").strip()
            h3(name)
            if kind:
                p(f"**Kind:** `{kind}`")
                p()
            if body:
                p("```powerquery")
                # ↓ sanitize vóór output
                p(sanitize(body))
                p("```")
            p()

    h2("Data Lineage & Impact Analysis")
    p("For each table: where its data comes from (upstream), and what depends on it "
      "within the model and report (downstream) — useful for assessing the impact of "
      "a change before making it.")
    p()

    lineage = build_lineage_index(parser, tables, all_table_names, all_measure_names, report_data)
    for table in tables:
        name = table.get("name", "Unknown")
        info = lineage.get(name, {})
        h3(name)

        source = info.get("source")
        if source:
            detail = f"**Upstream (source):** {source['connector']}"
            bits = []
            if source.get("schema"):
                bits.append(f"schema: `{source['schema']}`")
            if source.get("object"):
                bits.append(f"object: `{source['object']}`")
            if bits:
                detail += " — " + ", ".join(bits)
            p(detail)
        else:
            p("**Upstream (source):** no Power Query source found for this table "
              "(may be a calculated table, or DirectQuery/live connection).")
        p()

        ref_columns = info.get("ref_columns", [])
        ref_measures = info.get("ref_measures", [])
        related = info.get("related", [])
        used_in_visuals = info.get("used_in_visuals", [])

        if ref_columns or ref_measures or related:
            p("**Downstream — referenced by (within the model):**")
            if ref_columns:
                p(f"- Calculated columns ({len(ref_columns)}): "
                  + ", ".join(f"`{t}[{c}]`" for t, c in ref_columns))
            if ref_measures:
                p(f"- Measures ({len(ref_measures)}): "
                  + ", ".join(f"`{m}`" for _, m in ref_measures))
            if related:
                p(f"- Relationships ({len(related)}): " + "; ".join(f"`{r}`" for r in related))
        else:
            p("**Downstream — referenced by (within the model):** none found — no other "
              "calculated column, measure, or relationship references this table.")
        p()

        if used_in_visuals:
            p(f"**Downstream — used in report ({len(used_in_visuals)} visual(s)):**")
            for usage in used_in_visuals:
                p(f"- {usage}")
        else:
            p("**Downstream — used in report:** not used in any visual field binding found in the report.")
        p()

    p("---")
    p("_Generated by pbip_document.py_")

    return "\n".join(lines)


# ---------------------------------------------------------------------------
# DAX dependency analyser
# ---------------------------------------------------------------------------

def extract_dax_refs(dax: str, all_tables: list[str], all_measures: list[str]) -> dict:
    refs: dict = {"tables": set(), "columns": set(), "measures": set()}

    for match in re.finditer(r"'?([A-Za-z0-9_ ]+)'?\[([^\]]+)\]", dax):
        t, c = match.group(1).strip(), match.group(2).strip()
        if t:
            refs["tables"].add(t)
            refs["columns"].add(f"{t}[{c}]")

    for match in re.finditer(r"(?<!')\[([^\]]+)\]", dax):
        name = match.group(1).strip()
        if name in all_measures:
            refs["measures"].add(name)

    for t in all_tables:
        if not t:
            continue
        pattern = rf"\b{re.escape(t)}\b"
        if re.search(pattern, dax, re.IGNORECASE):
            refs["tables"].add(t)

    return {k: sorted(v) for k, v in refs.items()}


def describe_dax(dax: str) -> str:
    dax_upper = dax.strip().upper()
    patterns = [
        (r"^CALCULATE\s*\(", "A CALCULATE expression that evaluates a value in a modified filter context."),
        (r"^DIVIDE\s*\(", "A safe division (DIVIDE) that returns a result or an alternate value on divide-by-zero."),
        (r"^SUM\s*\(", "A simple SUM aggregation over a column."),
        (r"^SUMX\s*\(", "A row-by-row SUMX iterator that sums an expression over a table."),
        (r"^COUNT\s*\(", "A COUNT aggregation."),
        (r"^COUNTROWS\s*\(", "Counts the number of rows in a table or filtered table."),
        (r"^COUNTX\s*\(", "A row-by-row COUNTX iterator."),
        (r"^DISTINCTCOUNT\s*\(", "Counts the number of distinct values in a column."),
        (r"^AVERAGE\s*\(", "An AVERAGE aggregation over a column."),
        (r"^AVERAGEX\s*\(", "A row-by-row AVERAGEX iterator."),
        (r"^IF\s*\(", "A conditional expression (IF) that returns one of two values based on a condition."),
        (r"^SWITCH\s*\(", "A SWITCH expression that returns different values based on matching conditions."),
        (r"^VAR\s+", "Uses variables (VAR/RETURN) to break the calculation into named intermediate steps."),
        (r"^FILTER\s*\(", "A FILTER expression that returns a filtered table."),
        (r"^ALL\s*\(", "Removes filters using ALL() to return all rows or all values."),
        (r"^RELATED\s*\(", "Looks up a value from a related table via a relationship."),
        (r"^LOOKUPVALUE\s*\(", "Looks up a value based on matching criteria (LOOKUPVALUE)."),
        (r"^FORMAT\s*\(", "Formats a value as text using a format string."),
        (r"^DATEADD\s*\(", "A time intelligence calculation shifting dates by a specified interval."),
        (r"^SAMEPERIODLASTYEAR\s*\(", "A time intelligence calculation comparing to the same period last year."),
        (r"^TOTALYTD\s*\(", "Calculates a year-to-date total."),
        (r"^TOTALMTD\s*\(", "Calculates a month-to-date total."),
        (r"^TOTALQTD\s*\(", "Calculates a quarter-to-date total."),
        (r"^RANKX\s*\(", "Ranks values within a table using RANKX."),
        (r"^TOPN\s*\(", "Returns the top N rows of a table."),
        (r"^SELECTEDVALUE\s*\(", "Returns the selected value from a slicer or filter context (SELECTEDVALUE)."),
        (r"^HASONEVALUE\s*\(", "Checks whether a column is filtered to a single value."),
        (r"^ISBLANK\s*\(", "Checks whether a value is blank."),
        (r"^BLANK\s*\(", "Returns a BLANK value."),
        (r"^MIN\s*\(", "Returns the minimum value in a column."),
        (r"^MAX\s*\(", "Returns the maximum value in a column."),
        (r"^MAXX\s*\(", "A row-by-row MAXX iterator returning the maximum expression value."),
        (r"^MINX\s*\(", "A row-by-row MINX iterator returning the minimum expression value."),
        (r"^CONCATENATE\s*\(", "Concatenates two text strings."),
        (r"^CONCATENATEX\s*\(", "Concatenates text values across rows of a table."),
    ]
    for pattern, description in patterns:
        if re.match(pattern, dax_upper):
            return description
    return "A DAX expression."


# ---------------------------------------------------------------------------
# Data lineage / impact analysis
#
# Builds, per table, an upstream view (which physical source it's loaded from)
# and a downstream view (which calculated columns, measures, relationships and
# report visuals depend on it) — so a reader can answer "what would break if I
# changed this table?" without cross-referencing every other section by hand.
# ---------------------------------------------------------------------------

# M function names that identify a connector, checked in priority order — most
# specific first, since e.g. a Fabric Warehouse connection still calls Sql.Database().
_CONNECTOR_PATTERNS = [
    (re.compile(r"\.datawarehouse\.fabric\.microsoft\.com", re.IGNORECASE), "Fabric SQL Endpoint (Warehouse / Lakehouse SQL analytics endpoint)"),
    (re.compile(r"\bLakehouse\.Contents\s*\(", re.IGNORECASE), "Fabric Lakehouse (OneLake)"),
    (re.compile(r"\bDataflow\b|\bPowerPlatform\.Dataflows\s*\(", re.IGNORECASE), "Power BI / Fabric Dataflow"),
    (re.compile(r"sharepoint\.com|\bSharePoint\.(Files|Contents)\s*\(", re.IGNORECASE), "SharePoint"),
    (re.compile(r"\bOneDrive\.", re.IGNORECASE), "OneDrive"),
    (re.compile(r"\bSql\.Database\s*\(", re.IGNORECASE), "SQL Server / Azure SQL"),
    (re.compile(r"\bPostgreSQL\.Database\s*\(", re.IGNORECASE), "PostgreSQL"),
    (re.compile(r"\bMySQL\.Database\s*\(", re.IGNORECASE), "MySQL"),
    (re.compile(r"\bOdbc\.(Query|DataSource)\s*\(", re.IGNORECASE), "ODBC source"),
    (re.compile(r"\bOData\.Feed\s*\(", re.IGNORECASE), "OData feed"),
    (re.compile(r"\bWeb\.Contents\s*\(", re.IGNORECASE), "Web API"),
    (re.compile(r"\bExcel\.Workbook\s*\(", re.IGNORECASE), "Excel workbook"),
    (re.compile(r"\bCsv\.Document\s*\(", re.IGNORECASE), "CSV file"),
    (re.compile(r"\bJson\.Document\s*\(", re.IGNORECASE), "JSON source"),
    (re.compile(r"\bFolder\.(Files|Contents)\s*\(", re.IGNORECASE), "Folder"),
]

_SCHEMA_ITEM_RE = re.compile(
    r'\[\s*(?:Schema\s*=\s*(?:"(?P<schema_lit>[^"]*)"|(?P<schema_var>[A-Za-z_][A-Za-z0-9_]*))\s*,\s*)?'
    r'Item\s*=\s*(?:"(?P<item_lit>[^"]*)"|(?P<item_var>[A-Za-z_][A-Za-z0-9_]*))',
    re.IGNORECASE,
)


def parse_source_chain(m_expression: str) -> Optional[dict]:
    """Classifies a table's M-query source without leaking any org-specific literal
    value — only the connector *type* and the non-sensitive schema/object navigation
    (already established as safe to show; see sanitize()) are returned."""
    if not m_expression or not m_expression.strip():
        return None

    connector = next(
        (label for pattern, label in _CONNECTOR_PATTERNS if pattern.search(m_expression)),
        "Other / custom M source",
    )

    schema = table_obj = None
    m = _SCHEMA_ITEM_RE.search(m_expression)
    if m:
        schema = m.group("schema_lit") or (
            f"(from step/parameter: {m.group('schema_var')})" if m.group("schema_var") else None
        )
        table_obj = m.group("item_lit") or (
            f"(from step/parameter: {m.group('item_var')})" if m.group("item_var") else None
        )

    return {"connector": connector, "schema": schema, "object": table_obj}


def build_lineage_index(
    parser,
    tables: list[dict],
    all_table_names: list[str],
    all_measure_names: list[str],
    report_data: Optional[dict],
) -> dict[str, dict]:
    # Every calculated column's and measure's own DAX dependencies, computed once.
    col_deps = []   # (owner_table, column_name, refs)
    measure_deps = []  # (owner_table, measure_name, refs)

    for t in tables:
        owner = t.get("name", "")
        for col in parser.get_columns(t):
            expr = col.get("expression", "") if col.get("isCalculated") else ""
            if expr:
                col_deps.append((owner, col.get("name", ""), extract_dax_refs(expr, all_table_names, all_measure_names)))
        for measure in parser.get_measures(t):
            expr = measure.get("expression", "").strip()
            if expr:
                measure_deps.append((owner, measure.get("name", ""), extract_dax_refs(expr, all_table_names, all_measure_names)))

    # Which report visuals surface a field from each table.
    table_visual_usage: dict[str, list[str]] = {}
    if report_data:
        for page in report_data.get("pages", []):
            page_name = page.get("name", "Unknown")
            for visual in page.get("visuals", []):
                vtype = visual.get("visualType", "unknown")
                title = visual.get("title", "")
                label = f'"{title}" ({vtype}) on page "{page_name}"' if title else f'{vtype} on page "{page_name}"'
                seen = set()
                for f in visual.get("fields", []):
                    fm = re.search(r"([A-Za-z0-9_ ]+)\[", f.get("field", ""))
                    if fm:
                        entity = fm.group(1).strip()
                        if entity and entity not in seen:
                            seen.add(entity)
                            table_visual_usage.setdefault(entity, []).append(label)

    lineage: dict[str, dict] = {}
    for t in tables:
        name = t.get("name", "")

        source = None
        for pt in parser.get_partitions(t):
            src = pt.get("source", {})
            expr = src.get("expression", "") if isinstance(src, dict) else ""
            if expr:
                source = parse_source_chain(expr)
                break

        related = []
        for rel in parser.relationships():
            from_t = rel.get("fromTable", rel.get("fromTableId", ""))
            to_t = rel.get("toTable", rel.get("toTableId", ""))
            from_c = rel.get("fromColumn", rel.get("fromColumnId", ""))
            to_c = rel.get("toColumn", rel.get("toColumnId", ""))
            if from_t == name or to_t == name:
                related.append(f"{from_t}[{from_c}] → {to_t}[{to_c}]")

        # Deliberately includes same-table dependents too (e.g. a measure defined on
        # this table that sums one of its own columns) — those break just as much if
        # this table's columns change, so they belong in the impact analysis.
        lineage[name] = {
            "source": source,
            "related": related,
            "ref_columns": [(owner, cname) for owner, cname, refs in col_deps if name in refs["tables"]],
            "ref_measures": [(owner, mname) for owner, mname, refs in measure_deps if name in refs["tables"]],
            "used_in_visuals": table_visual_usage.get(name, []),
        }

    return lineage


# ---------------------------------------------------------------------------
# Copilot knowledge base renderer
# ---------------------------------------------------------------------------

def render_copilot_kb(
    project_name: str,
    parser,
    report_data: Optional[dict],
    model_format: str,
    row_counts: Optional[dict[str, int]] = None,
) -> str:
    row_counts = row_counts or {}
    tables = parser.tables()
    all_table_names = [t.get("name", "") for t in tables]
    all_measure_names = [
        m.get("name", "")
        for t in tables
        for m in parser.get_measures(t)
    ]

    lines: list[str] = []
    def w(s=""): lines.append(s)

    w("=" * 80)
    w(f"POWER BI KNOWLEDGE BASE: {project_name.upper()}")
    w(f"Semantic model format: {model_format}")
    w("=" * 80)
    w()
    w("INSTRUCTIONS FOR AI ASSISTANT")
    w("-" * 40)
    w("This file is the complete technical knowledge base for the Power BI project")
    w(f'"{project_name}". Use it to answer questions about:')
    w("- How specific measures or fields are calculated (including full DAX logic)")
    w("- What columns exist in which tables and what they represent")
    w("- How tables are related to each other")
    w("- Which measures depend on other measures or columns")
    w("- Which pages and visuals exist in the report, and which fields each visual uses")
    w("- Where a table's data comes from, and what would be impacted if it changed "
      "(see Section 8: Data Lineage and Impact Analysis)")
    w()
    w("When answering questions about a measure, always explain:")
    w("1. What the measure calculates in plain language")
    w("2. The exact DAX formula")
    w("3. Which tables and columns it touches")
    w("4. Which other measures it calls (if any)")
    w("5. Which report visuals use this measure (if applicable)")
    w()

    w("=" * 80)
    w("INDEX")
    w("=" * 80)
    w(f"Tables ({len(tables)}): " + ", ".join(all_table_names))
    w(f"Total columns: {sum(len(parser.get_columns(t)) for t in tables)}")
    w(f"Total measures: {len(all_measure_names)}")
    w(f"Relationships: {len(parser.relationships())}")
    w(f"RLS Roles: {len(parser.roles())}")
    if report_data:
        pages = report_data.get("pages", [])
        total_visuals = sum(len(p.get("visuals", [])) for p in pages)
        w(f"Report pages: {len(pages)} ({total_visuals} visuals total)")
    if row_counts:
        w()
        w("Table row counts (manually pasted from a DAX row-count query, may be stale):")
        for t in tables:
            name = t.get("name", "")
            if name in row_counts:
                w(f"  {name}: {row_counts[name]:,} rows")
    w()

    w("=" * 80)
    w("SECTION 1: TABLES AND COLUMNS")
    w("=" * 80)
    w()

    for table in tables:
        table_name = table.get("name", "Unknown")
        columns = parser.get_columns(table)
        measures = parser.get_measures(table)
        partitions = parser.get_partitions(table)

        w(f"TABLE: {table_name}")
        w("-" * 60)

        if table_name in row_counts:
            w(f"Row count: {row_counts[table_name]:,} (manually pasted from a DAX query, may be stale)")

        for pt in partitions:
            source = pt.get("source", {})
            if isinstance(source, dict):
                query = source.get("query", source.get("expression", ""))
                if query:
                    w(f"This table is loaded from the following query:")
                    # ↓ sanitize vóór output
                    w(sanitize(query.strip()))
                    w()

        if columns:
            w(f"This table has {len(columns)} column(s):")
            w()
            for col in columns:
                col_name = col.get("name", "")
                dtype = col.get("dataType", col.get("type", "unknown"))
                desc = col.get("description", "")
                hidden = col.get("isHidden", False)
                is_calc = col.get("isCalculated", False)
                expr = col.get("expression", "")

                line = f'  Column "{col_name}" (data type: {dtype})'
                if is_calc:
                    line += " [calculated column]"
                if hidden:
                    line += " [hidden from report view]"
                if desc:
                    line += f": {desc}"
                w(line)

                if is_calc and expr:
                    w("    This is a calculated column, defined by the following DAX formula:")
                    for dax_line in sanitize(expr).splitlines():
                        w(f"      {dax_line}")
        else:
            w("This table has no regular columns (may be a calculated or virtual table).")

        if measures:
            w()
            w(f"This table contains {len(measures)} measure(s): "
              + ", ".join(f'"{m.get("name", "")}"' for m in measures))

        w()

    w("=" * 80)
    w("SECTION 2: MEASURES (FULL DAX AND DEPENDENCIES)")
    w("=" * 80)
    w()

    measure_visual_usage: dict[str, list[str]] = {}
    if report_data:
        for page in report_data.get("pages", []):
            page_name = page["name"]
            for visual in page.get("visuals", []):
                vtype = visual["visualType"]
                title = visual.get("title", "")
                label = f'"{title}" ({vtype}) on page "{page_name}"' if title else f'{vtype} on page "{page_name}"'
                for f in visual.get("fields", []):
                    field_str = f.get("field", "")
                    m = re.search(r"\[([^\]]+)\]", field_str)
                    if m:
                        fname = m.group(1)
                        if fname in all_measure_names:
                            measure_visual_usage.setdefault(fname, []).append(label)

    for table in tables:
        measures = parser.get_measures(table)
        if not measures:
            continue
        table_name = table.get("name", "Unknown")

        for measure in measures:
            name = measure.get("name", "")
            dax = measure.get("expression", "").strip()
            desc = measure.get("description", "")
            fmt = measure.get("formatString", "")

            w(f"MEASURE: {name}")
            w(f"Defined in table: {table_name}")
            w("-" * 60)

            if desc:
                w(f"Description: {desc}")
            if fmt:
                w(f"Display format: {fmt}")

            if dax:
                plain = describe_dax(dax)
                w(f"What it does: {plain}")
            w()

            if dax:
                w("Full DAX formula:")
                # ↓ sanitize vóór output
                w(sanitize(dax))
                w()

                deps = extract_dax_refs(dax, all_table_names, all_measure_names)

                if deps["measures"]:
                    w("This measure calls the following other measures:")
                    for m in deps["measures"]:
                        if m != name:
                            w(f'  - "{m}"')
                    w()

                if deps["columns"]:
                    w("This measure references the following columns:")
                    for c in deps["columns"]:
                        w(f"  - {c}")
                    w()

                if deps["tables"]:
                    w("This measure references the following tables:")
                    for t in deps["tables"]:
                        w(f"  - {t}")
                    w()
            else:
                w("No DAX expression was found for this measure.")
                w()

            usages = measure_visual_usage.get(name, [])
            if usages:
                w(f"This measure is used in {len(usages)} visual(s):")
                for usage in usages:
                    w(f"  - {usage}")
                w()

            w()

    w("=" * 80)
    w("SECTION 3: RELATIONSHIPS")
    w("=" * 80)
    w()

    rels = parser.relationships()
    if rels:
        w(f"The data model has {len(rels)} relationship(s):")
        w()
        for i, rel in enumerate(rels, 1):
            from_t = rel.get("fromTable", rel.get("fromTableId", ""))
            from_c = rel.get("fromColumn", rel.get("fromColumnId", ""))
            to_t = rel.get("toTable", rel.get("toTableId", ""))
            to_c = rel.get("toColumn", rel.get("toColumnId", ""))
            card = rel.get("cardinality", rel.get("fromCardinality", "many-to-one"))
            cf = rel.get("crossFilter", rel.get("crossFilteringBehavior", "single"))

            w(f"Relationship {i}:")
            w(f'  The column "{from_c}" in table "{from_t}" is linked to '
              f'the column "{to_c}" in table "{to_t}".')
            w(f"  Cardinality: {card}. Cross-filter direction: {cf}.")
            w()
    else:
        w("No relationships are defined in this model.")
        w()

    roles = parser.roles()
    if roles:
        w("=" * 80)
        w("SECTION 4: ROW-LEVEL SECURITY (RLS)")
        w("=" * 80)
        w()
        for role in roles:
            role_name = role.get("name", "")
            perms = role.get("tablePermissions", [])
            w(f"RLS Role: {role_name}")
            if perms:
                for perm in perms:
                    t = perm.get("table", perm.get("name", ""))
                    expr = perm.get("filterExpression", perm.get("expression", ""))
                    if expr:
                        w(f'  Table "{t}" is filtered by: {expr}')
                    else:
                        w(f'  Table "{t}" is included in this role (no filter expression).')
            w()

    if report_data:
        w("=" * 80)
        w("SECTION 5: REPORT STRUCTURE AND VISUAL FIELD BINDINGS")
        w("=" * 80)
        w()
        lines.extend(_report_structure_copilot_lines(report_data))

    w("=" * 80)
    w("SECTION 6: POWER QUERY (M) SOURCES")
    w("=" * 80)
    w()

    has_pq = False
    for table in tables:
        partitions = parser.get_partitions(table)
        for pt in partitions:
            source = pt.get("source", {})
            expr = source.get("expression", "") if isinstance(source, dict) else source
            if not expr:
                continue
            has_pq = True
            table_name = table.get("name", "Unknown")
            w(f"Table: {table_name} (mode: {pt.get('mode', 'import')})")
            w("-" * 60)
            w("Power Query expression:")
            # ↓ sanitize vóór output
            w(sanitize(expr.strip()))
            w()

    if not has_pq:
        w("No Power Query expressions found (model may use DirectQuery or live connection).")
        w()

    shared = parser.shared_expressions()
    if shared:
        w("=" * 80)
        w("SECTION 7: SHARED POWER QUERY FUNCTIONS")
        w("=" * 80)
        w()
        for expr in shared:
            name = expr.get("name", "")
            kind = expr.get("kind", "")
            body = expr.get("expression", "").strip()
            w(f"Shared expression: {name} (kind: {kind})")
            w("-" * 60)
            if body:
                # ↓ sanitize vóór output
                w(sanitize(body))
            w()

    w("=" * 80)
    w("SECTION 8: DATA LINEAGE AND IMPACT ANALYSIS")
    w("=" * 80)
    w()
    w("For each table: where its data comes from (upstream), and what depends on it")
    w("within the model and report (downstream). Use this to answer questions like")
    w('"what would break if I changed table X?" or "where does table Y\'s data come from?".')
    w()

    lineage = build_lineage_index(parser, tables, all_table_names, all_measure_names, report_data)
    for table in tables:
        name = table.get("name", "Unknown")
        info = lineage.get(name, {})

        w(f"TABLE: {name}")
        w("-" * 60)

        source = info.get("source")
        if source:
            detail = f"Upstream source: {source['connector']}"
            bits = []
            if source.get("schema"):
                bits.append(f"schema: {source['schema']}")
            if source.get("object"):
                bits.append(f"object: {source['object']}")
            if bits:
                detail += " (" + ", ".join(bits) + ")"
            w(detail)
        else:
            w("Upstream source: no Power Query source found (calculated table, or "
              "DirectQuery/live connection).")

        ref_columns = info.get("ref_columns", [])
        ref_measures = info.get("ref_measures", [])
        related = info.get("related", [])
        used_in_visuals = info.get("used_in_visuals", [])

        if ref_columns or ref_measures or related:
            w("Downstream — referenced by:")
            if ref_columns:
                w(f"  Calculated columns ({len(ref_columns)}): "
                  + ", ".join(f"{t}[{c}]" for t, c in ref_columns))
            if ref_measures:
                w(f"  Measures ({len(ref_measures)}): " + ", ".join(m for _, m in ref_measures))
            if related:
                w(f"  Relationships ({len(related)}): " + "; ".join(related))
        else:
            w("Downstream — referenced by: no other calculated column, measure, or "
              "relationship references this table.")

        if used_in_visuals:
            w(f"Downstream — used in report ({len(used_in_visuals)} visual(s)):")
            for usage in used_in_visuals:
                w(f"  - {usage}")
        else:
            w("Downstream — used in report: not used in any visual field binding found in the report.")

        w()

    w("=" * 80)
    w("END OF KNOWLEDGE BASE")
    w("=" * 80)

    return "\n".join(lines)


# ---------------------------------------------------------------------------
# Row-count helper — no data lives in the .pbip files themselves, so getting a
# real row count per table means running a DAX query against the loaded model.
# Rather than requiring a live XMLA/ADOMD.NET connection (extra dependency,
# only works while the model is open), this generates a query you paste into
# Power BI Desktop's DAX query view or DAX Studio by hand; the pasted-back
# result is then merged into the documentation via --rowcounts.
# ---------------------------------------------------------------------------

def generate_rowcount_query(tables: list[dict]) -> str:
    lines = [
        "// Row-count helper query — generated by pbip_extract.py",
        "//",
        "// How to use:",
        "//   1. Open this .pbip project in Power BI Desktop (the model must be loaded),",
        "//      or connect to it with DAX Studio.",
        "//   2. Power BI Desktop: View > DAX query view, paste the query below, run it.",
        "//      DAX Studio: paste the query below, run it (F5).",
        "//   3. Select all results (Ctrl+A), copy (Ctrl+C), and paste them into a new",
        "//      text file (headers included).",
        "//   4. Re-run pbip_extract.py with --rowcounts <that file> to include the",
        "//      counts in the generated documentation.",
        "",
    ]

    names = [t.get("name", "") for t in tables if t.get("name")]
    if not names:
        lines.append("// No tables found — nothing to query.")
        return "\n".join(lines)

    rows = []
    for name in names:
        literal = name.replace('"', '""')
        identifier = name.replace("'", "''")
        rows.append(f'\tROW("Table", "{literal}", "Rows", COUNTROWS(\'{identifier}\'))')

    lines.append("EVALUATE")
    lines.append("UNION(")
    lines.append(",\n".join(rows))
    lines.append(")")
    lines.append("ORDER BY [Table]")
    return "\n".join(lines)


def parse_rowcount_file(path: Path) -> dict[str, int]:
    """Parses a DAX Studio / Power BI Desktop result grid pasted as plain text
    (tab-, comma-, or semicolon-separated, header included) into {table_name: row_count}.
    Semicolon is included because Excel's clipboard/CSV list separator is ";" rather
    than "," on many regional (e.g. Dutch/German) system locales."""
    text = path.read_text(encoding="utf-8-sig")
    lines = [ln for ln in text.splitlines() if ln.strip()]
    counts: dict[str, int] = {}
    if len(lines) < 2:
        return counts

    if "\t" in lines[0]:
        delim = "\t"
    elif ";" in lines[0]:
        delim = ";"
    elif "," in lines[0]:
        delim = ","
    else:
        delim = None

    def split(line: str) -> list[str]:
        cols = line.split(delim) if delim else [line]
        return [c.strip().strip('"').strip("[]").strip() for c in cols]

    header = [h.lower() for h in split(lines[0])]
    table_idx = next((i for i, h in enumerate(header) if "table" in h), 0)
    rows_idx = next((i for i, h in enumerate(header) if "row" in h), len(header) - 1)

    for line in lines[1:]:
        cols = split(line)
        if len(cols) <= max(table_idx, rows_idx):
            continue
        name = cols[table_idx]
        digits = re.sub(r"[^\d]", "", cols[rows_idx])
        if name and digits:
            counts[name] = int(digits)

    return counts


# ---------------------------------------------------------------------------
# Project discovery
#
# A .pbip file and the .pbir file inside its report folder already contain the
# exact, authoritative references to the report and semantic-model folders
# (artifacts[].report.path, and datasetReference.byPath/.byConnection). Reading
# those is faster and more correct than guessing via a directory scan — it can't
# pick up an unrelated report/model that happens to sit elsewhere in the tree,
# and it correctly reports a live/remote connection instead of silently finding
# nothing. The scan-based functions below are kept only as a fallback for
# projects where those references can't be read for some reason.
# ---------------------------------------------------------------------------

def find_report_dir_from_pbip(pbip_file: Path) -> Optional[Path]:
    """Resolves the report folder from the .pbip file's `artifacts` list."""
    try:
        data = read_json(pbip_file)
    except Exception:
        return None
    for artifact in data.get("artifacts", []):
        report = artifact.get("report")
        if report and report.get("path"):
            candidate = (pbip_file.parent / report["path"]).resolve()
            if candidate.is_dir():
                return candidate
    return None


def find_semantic_model_from_report(report_dir: Path) -> tuple[Optional[Path], str]:
    """Resolves the semantic model from the report's definition.pbir file.

    Returns (path, "TMDL"/"TMSL") on success, (None, "connection") if the report
    is bound to a live connection to a remote/published model (nothing local to
    document), or (None, "") if no usable reference was found (caller should fall
    back to find_semantic_model()).
    """
    pbir_path = report_dir / "definition.pbir"
    if not pbir_path.exists():
        return None, ""
    try:
        data = read_json(pbir_path)
    except Exception:
        return None, ""

    ds_ref = data.get("datasetReference", {})
    if "byConnection" in ds_ref:
        return None, "connection"

    rel = ds_ref.get("byPath", {}).get("path", "")
    if not rel:
        return None, ""

    model_dir = (report_dir / rel).resolve()
    if not model_dir.is_dir():
        return None, ""

    db_tmdl = model_dir / "definition" / "database.tmdl"
    if db_tmdl.exists():
        return db_tmdl.parent, "TMDL"

    definition_dir = model_dir / "definition"
    if definition_dir.is_dir() and any(definition_dir.glob("*.tmdl")):
        return definition_dir, "TMDL"

    bim_files = list(model_dir.glob("*.bim"))
    if bim_files:
        return bim_files[0], "TMSL"

    return None, ""


def resolve_report_definition_dir(report_dir: Path) -> Optional[Path]:
    """Modern report definition folder, given the report folder is already known."""
    definition_dir = report_dir / "definition"
    if (definition_dir / "pages" / "pages.json").exists():
        return definition_dir
    return None


def resolve_report_json(report_dir: Path) -> Optional[Path]:
    """Legacy report.json, given the report folder is already known."""
    candidate = report_dir / "report.json"
    if candidate.exists():
        try:
            data = read_json(candidate)
            if data.get("sections"):
                return candidate
        except Exception:
            pass
    return None


def find_semantic_model(pbip_root: Path) -> tuple[Optional[Path], str]:
    """Fallback: scans the project folder for a semantic model when the .pbip/.pbir
    references couldn't be resolved."""
    for candidate in pbip_root.rglob("database.tmdl"):
        return candidate.parent, "TMDL"

    for candidate in pbip_root.rglob("definition"):
        if candidate.is_dir() and any(candidate.glob("*.tmdl")):
            return candidate, "TMDL"

    for candidate in pbip_root.rglob("*.bim"):
        return candidate, "TMSL"

    return None, ""


def find_report_json(pbip_root: Path) -> Optional[Path]:
    """Fallback: scans the project folder for a legacy report.json."""
    for candidate in pbip_root.rglob("report.json"):
        try:
            data = read_json(candidate)
            if data.get("sections"):
                return candidate
        except Exception:
            pass
    return None


def find_report_definition_dir(pbip_root: Path) -> Optional[Path]:
    """Fallback: scans the project folder for a modern report definition folder."""
    for candidate in pbip_root.rglob("pages.json"):
        pages_dir = candidate.parent
        if pages_dir.name == "pages":
            return pages_dir.parent
    return None


# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------

def main():
    ap = argparse.ArgumentParser(description="Document a Power BI project from its .pbip file")
    ap.add_argument("pbip_file", help="Path to the .pbip project file (not the project folder)")
    ap.add_argument("--output", "-o", default="", help="Output file path")
    ap.add_argument(
        "--copilot",
        action="store_true",
        help=(
            "Generate a plain-text knowledge base optimised for Copilot / ChatGPT "
            "context injection instead of Markdown docs."
        ),
    )
    ap.add_argument(
        "--rowcounts",
        default="",
        help=(
            "Path to a text file with a pasted DAX row-count result (table + row count, "
            "tab-, comma-, or semicolon-separated, header included) to include per-table row counts "
            "in the documentation. Get this file by running the query from "
            "--rowcount-query in Power BI Desktop's DAX query view or DAX Studio and "
            "pasting the result grid into a text file."
        ),
    )
    ap.add_argument(
        "--no-rowcount-query",
        action="store_true",
        help="Don't write the companion *_rowcount_query.dax helper file (see --rowcounts).",
    )
    args = ap.parse_args()

    input_path = Path(args.pbip_file)
    if not input_path.is_absolute():
        input_path = Path.cwd() / args.pbip_file.lstrip('/\\')

    if not input_path.exists():
        print(f"ERROR: Path does not exist: {input_path.resolve()}", file=sys.stderr)
        sys.exit(1)
    if input_path.is_dir():
        guesses = sorted(input_path.glob("*.pbip"))
        hint = f" Did you mean: {guesses[0].name}?" if guesses else ""
        print(
            f"ERROR: Expected a .pbip file, got a folder: {input_path.resolve()}\n"
            f"       Point this at the *.pbip file inside that folder instead.{hint}",
            file=sys.stderr,
        )
        sys.exit(1)
    if input_path.suffix.lower() != ".pbip":
        print(f"ERROR: Expected a .pbip file, got: {input_path.resolve()}", file=sys.stderr)
        sys.exit(1)

    pbip_root = input_path.parent
    project_name = input_path.stem.replace("_", " ").title()

    # Prefer the exact references from the .pbip/.pbir files; fall back to a
    # directory scan only if those can't be read (e.g. non-standard project layout).
    report_dir = find_report_dir_from_pbip(input_path)
    if report_dir:
        print(f"Found report artifact (from .pbip): {report_dir}")
    else:
        print(
            "WARNING: Could not read the report artifact reference from the .pbip file; "
            "falling back to scanning the project folder.",
            file=sys.stderr,
        )

    report_data = None

    report_def_dir = resolve_report_definition_dir(report_dir) if report_dir else None
    if not report_def_dir:
        report_def_dir = find_report_definition_dir(pbip_root)
    if report_def_dir:
        print(f"Found modern report definition: {report_def_dir}")
        try:
            report_data = parse_report_definition(report_def_dir)
        except Exception as e:
            print(f"WARNING: Could not parse modern report definition: {e}", file=sys.stderr)

    if not report_data or not report_data.get("pages"):
        report_json_path = resolve_report_json(report_dir) if report_dir else None
        if not report_json_path:
            report_json_path = find_report_json(pbip_root)
        if report_json_path:
            print(f"Found legacy report.json: {report_json_path}")
            try:
                report_data = parse_report(report_json_path)
            except Exception as e:
                print(f"WARNING: Could not parse report.json: {e}", file=sys.stderr)
        else:
            print("No report found — skipping report structure section.")

    if report_data and report_data.get("pages"):
        total_visuals = sum(len(p.get("visuals", [])) for p in report_data["pages"])
        total_fields = sum(
            len(v.get("fields", []))
            for p in report_data["pages"]
            for v in p.get("visuals", [])
        )
        print(f"  Pages: {len(report_data['pages'])}, Visuals: {total_visuals}, Field bindings: {total_fields}")

    model_path, model_format = (None, "")
    if report_dir:
        model_path, model_format = find_semantic_model_from_report(report_dir)
        if model_path:
            print(f"Found semantic model ({model_format}, via .pbip/.pbir reference): {model_path}")

    if not model_path and model_format != "connection":
        model_path, model_format = find_semantic_model(pbip_root)
        if model_path:
            print(f"Found semantic model ({model_format}, by scanning the project folder): {model_path}")

    if model_format == "connection":
        # No local semantic model to document (report is bound to a live connection
        # to a remote/published model) — write just the report structure instead of
        # failing outright.
        print(
            "This report uses a live connection to a remote/published semantic model "
            "— there is no local semantic model to document. Writing the report "
            "structure (pages and visuals) only."
        )

        if args.copilot:
            content = render_report_only_copilot_kb(project_name, report_data)
            default_name = f"{project_name.replace(' ', '_')}_copilot_kb.txt"
            suffix = "Copilot knowledge base"
        else:
            content = render_report_only_markdown(project_name, report_data)
            default_name = f"{project_name.replace(' ', '_')}_docs.md"
            suffix = "Markdown documentation"

        output_path = args.output or default_name
        out = Path(output_path)
        out.write_text(content, encoding="utf-8")

        print(f"\nDone. {suffix} written to: {out.resolve()}")
        if report_data and report_data.get("pages"):
            print(f"  Pages:   {len(report_data['pages'])}")
            print(f"  Visuals: {sum(len(p.get('visuals', [])) for p in report_data['pages'])}")
        return

    if not model_path:
        print("ERROR: No semantic model found (.bim or TMDL definition folder).", file=sys.stderr)
        sys.exit(1)

    if model_format == "TMSL":
        parser = TMSLParser(model_path)
    else:
        parser = TMLDParser(model_path)

    tables = parser.tables()
    n_measures = sum(len(parser.get_measures(t)) for t in tables)

    row_counts: dict[str, int] = {}
    if args.rowcounts:
        rc_path = Path(args.rowcounts)
        if rc_path.exists():
            row_counts = parse_rowcount_file(rc_path)
            print(f"Loaded row counts for {len(row_counts)} table(s) from: {rc_path}")
        else:
            print(f"WARNING: --rowcounts file not found: {rc_path}", file=sys.stderr)
    elif not args.no_rowcount_query:
        dax_query = generate_rowcount_query(tables)
        out_dir = Path(args.output).parent if args.output else Path.cwd()
        dax_path = out_dir / f"{project_name.replace(' ', '_')}_rowcount_query.dax"
        dax_path.write_text(dax_query, encoding="utf-8")
        print(f"Row-count helper query written to: {dax_path.resolve()}")
        print("  Run it in Power BI Desktop's DAX query view or DAX Studio, paste the")
        print("  results into a text file, then re-run with --rowcounts <that file> to")
        print("  include row counts in the documentation.")

    if args.copilot:
        content = render_copilot_kb(project_name, parser, report_data, model_format, row_counts)
        default_name = f"{project_name.replace(' ', '_')}_copilot_kb.txt"
        suffix = "Copilot knowledge base"
    else:
        content = render_markdown(project_name, parser, report_data, model_format, row_counts)
        default_name = f"{project_name.replace(' ', '_')}_docs.md"
        suffix = "Markdown documentation"

    output_path = args.output or default_name
    out = Path(output_path)
    out.write_text(content, encoding="utf-8")

    print(f"\nDone. {suffix} written to: {out.resolve()}")
    print(f"  Tables:        {len(tables)}")
    print(f"  Measures:      {n_measures}")
    print(f"  Relationships: {len(parser.relationships())}")
    if row_counts:
        print(f"  Row counts:    {len(row_counts)}/{len(tables)} table(s)")

    if args.copilot:
        size_kb = out.stat().st_size / 1024
        print(f"  File size:     {size_kb:.1f} KB")
        print()
        print("How to use:")
        print("  Copilot for M365: Upload this .txt file in a Teams/Word chat with Copilot.")
        print("  Claude Project:   Add this file as project knowledge.")
        print("  ChatGPT:          Upload via paperclip, or paste into a custom GPT system prompt.")
        print("  Copilot Studio:   Use as a knowledge source document.")


if __name__ == "__main__":
    main()
