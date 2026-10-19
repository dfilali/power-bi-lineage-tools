#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Power BI PBIT inventory + DAX lineage (standard library only).

Usage:
    python pbit_lineage_inventory.py "rapport.pbit"
    python pbit_lineage_inventory.py "rapport.pbit" -o "inventaire_lineage.csv"

Outputs ONE CSV, one row per discovered table, column, measure, model expression,
and (where present) partition. The script does not modify the source file.

Important:
- PBIT internal schemas vary by Power BI version. The extractor is deliberately
  defensive and records warnings in the CSV rather than silently discarding
  objects.
- DAX dependency parsing is lexical/static analysis, not a full Microsoft DAX
  parser. Dynamic references and some unusual syntax may need manual review.
"""

import argparse
import csv
import json
import os
import re
import sys
import zipfile
from collections import defaultdict, deque
from pathlib import Path


CSV_FIELDS = [
    "object_type",
    "object_name",
    "table_name",
    "object_id",
    "is_hidden",
    "data_type",
    "expression_language",
    "expression",
    "source_kind",
    "source_name",
    "source_expression",
    "direct_dependencies",
    "upstream_dependencies",
    "source_tables",
    "lineage_status",
    "warnings",
]


def s(value):
    """Stable text conversion for values that may be absent or structured."""
    if value is None:
        return ""
    if isinstance(value, str):
        return value
    if isinstance(value, (dict, list)):
        try:
            return json.dumps(value, ensure_ascii=False, sort_keys=True)
        except Exception:
            return str(value)
    return str(value)


def bool_text(value):
    if value is None:
        return ""
    if isinstance(value, bool):
        return "true" if value else "false"
    if isinstance(value, (int, float)):
        return "true" if value != 0 else "false"
    return s(value)


def norm_name(value):
    return re.sub(r"\s+", " ", s(value).strip()).casefold()


def get_any(d, *keys, default=None):
    if not isinstance(d, dict):
        return default
    for key in keys:
        if key in d and d[key] is not None:
            return d[key]
    return default


def list_of_dicts(value):
    if isinstance(value, list):
        return [x for x in value if isinstance(x, dict)]
    return []


def read_json_bytes(raw):
    # JSON in DataModelSchema is normally UTF-8; BOM and UTF-16 are handled too.
    for encoding in ("utf-8-sig", "utf-16", "utf-16-le", "cp1252"):
        try:
            return json.loads(raw.decode(encoding))
        except (UnicodeDecodeError, json.JSONDecodeError):
            continue
    raise ValueError("Impossible de lire DataModelSchema comme JSON UTF-8/UTF-16.")


def find_model_schema(pbit_path):
    """Return (parsed JSON, schema path in archive, warnings)."""
    warnings = []
    try:
        with zipfile.ZipFile(pbit_path, "r") as zf:
            names = zf.namelist()
            candidates = [
                n for n in names
                if Path(n).name.casefold() == "datamodelschema"
                or Path(n).name.casefold() == "datamodelschema.json"
            ]
            if not candidates:
                # Some exports or wrappers may use an extension or different casing.
                candidates = [
                    n for n in names
                    if "datamodelschema" in Path(n).name.casefold()
                    and not Path(n).name.casefold().endswith((".bak", ".txt"))
                ]
            if not candidates:
                raise FileNotFoundError(
                    "Aucun fichier DataModelSchema trouvé dans l'archive PBIT. "
                    "Le fichier peut être chiffré, non standard, ou ne pas être un PBIT valide."
                )
            # Prefer the shortest/root-most candidate.
            candidates.sort(key=lambda n: (n.count("/"), len(n)))
            selected = candidates[0]
            if len(candidates) > 1:
                warnings.append("Plusieurs fichiers DataModelSchema trouvés; utilisé: " + selected)
            return read_json_bytes(zf.read(selected)), selected, warnings
    except zipfile.BadZipFile as exc:
        raise ValueError("Le fichier fourni n'est pas une archive ZIP/PBIT valide.") from exc


def clean_dax_for_scan(expr):
    """
    Mask comments and double-quoted DAX string literals while preserving offsets.
    Single quotes are kept because they quote table names in DAX (e.g. 'Sales Data').
    """
    text = s(expr)
    out = list(text)
    i = 0
    n = len(text)
    while i < n:
        # DAX line comments: -- or //
        if text.startswith("--", i) or text.startswith("//", i):
            j = text.find("\n", i)
            if j < 0:
                j = n
            for k in range(i, j):
                if out[k] != "\n":
                    out[k] = " "
            i = j
            continue
        # Block comment
        if text.startswith("/*", i):
            j = text.find("*/", i + 2)
            j = n if j < 0 else j + 2
            for k in range(i, j):
                if out[k] != "\n":
                    out[k] = " "
            i = j
            continue
        # DAX string literals use double quotes; doubled quotes escape quote.
        if text[i] == '"':
            j = i + 1
            while j < n:
                if text[j] == '"':
                    if j + 1 < n and text[j + 1] == '"':
                        j += 2
                        continue
                    j += 1
                    break
                j += 1
            for k in range(i, j):
                if out[k] != "\n":
                    out[k] = " "
            i = j
            continue
        i += 1
    return "".join(out)


def unquote_table_name(name):
    name = s(name).strip()
    if len(name) >= 2 and name[0] == "'" and name[-1] == "'":
        return name[1:-1].replace("''", "'").strip()
    return name


# Captures 'Table Name'[Column], Table[Column], including spaces in unquoted names
# conservatively. Quoted table names are preferred where names contain spaces.
TABLE_COLUMN_RE = re.compile(
    r"(?P<table>'(?:''|[^'])*'|[A-Za-z_][A-Za-z0-9_ .-]*?)\s*"
    r"\[\s*(?P<column>[^\]\r\n]+?)\s*\]"
)
BRACKET_REF_RE = re.compile(r"\[\s*(?P<name>[^\]\r\n]+?)\s*\]")


def extract_dax_references(expression):
    """Return (table-column references, unqualified bracket references)."""
    cleaned = clean_dax_for_scan(expression)
    table_column_refs = []
    occupied = []
    for match in TABLE_COLUMN_RE.finditer(cleaned):
        table = unquote_table_name(match.group("table").strip())
        column = match.group("column").strip()
        # Avoid capturing a DAX function/identifier as a table name.
        if not table or not column:
            continue
        table_column_refs.append((table, column))
        occupied.append((match.start(), match.end()))

    bracket_refs = []
    for match in BRACKET_REF_RE.finditer(cleaned):
        if any(start <= match.start() < end for start, end in occupied):
            continue
        name = match.group("name").strip()
        if name:
            bracket_refs.append(name)
    return table_column_refs, bracket_refs


def get_expression_text(obj):
    if not isinstance(obj, dict):
        return ""
    # In common schemas DAX is "expression"; M query definitions also use expression.
    return s(get_any(obj, "expression", "formula", "expressionText", "query", default=""))


def hidden_value(obj, parent=None):
    if not isinstance(obj, dict):
        return ""
    value = get_any(obj, "isHidden", "hidden", "IsHidden", default=None)
    if value is None and isinstance(parent, dict):
        value = get_any(parent, "isHidden", "hidden", default=None)
    return bool_text(value)


def stable_id(obj, fallback):
    return s(get_any(obj, "lineageTag", "id", "objectId", "name", default=fallback))


def main():
    parser = argparse.ArgumentParser(
        description="Inventaire Power BI PBIT et lineage DAX dans un seul CSV (stdlib Python)."
    )
    parser.add_argument("pbit", help="Chemin du fichier .pbit")
    parser.add_argument("-o", "--output", help="CSV de sortie (défaut: <nom>_inventaire_lineage.csv)")
    args = parser.parse_args()

    pbit_path = Path(args.pbit).expanduser()
    if not pbit_path.is_file():
        print("ERREUR: fichier introuvable: " + str(pbit_path), file=sys.stderr)
        return 2

    output_path = Path(args.output).expanduser() if args.output else (
        pbit_path.with_name(pbit_path.stem + "_inventaire_lineage.csv")
    )

    try:
        schema, schema_member, global_warnings = find_model_schema(pbit_path)
    except Exception as exc:
        print("ERREUR: " + str(exc), file=sys.stderr)
        return 2

    model = schema.get("model", schema) if isinstance(schema, dict) else {}
    if not isinstance(model, dict):
        print("ERREUR: structure JSON inattendue: objet model introuvable.", file=sys.stderr)
        return 2

    tables = list_of_dicts(get_any(model, "tables", "Tables", default=[]))
    expressions = list_of_dicts(get_any(model, "expressions", "Expressions", default=[]))
    rows = []
    objects = {}       # canonical object key -> row
    name_index = defaultdict(list)
    table_index = {}   # normalized table name -> table object
    column_index = defaultdict(list)
    measure_index = defaultdict(list)
    expression_index = defaultdict(list)
    warnings_by_key = defaultdict(list)

    def add_object(key, row, aliases=()):
        if key in objects:
            # Do not lose duplicate names; make key unique.
            base = key
            n = 2
            while key in objects:
                key = f"{base}#{n}"
                n += 1
            row["warnings"] = (row.get("warnings", "") + "; duplicate canonical key").strip("; ")
        objects[key] = row
        row["_key"] = key
        rows.append(row)
        for alias in set([row.get("object_name", ""), *aliases]):
            if alias:
                name_index[norm_name(alias)].append(key)
        return key

    def new_row(object_type, object_name, table_name="", obj=None, **kwargs):
        obj = obj if isinstance(obj, dict) else {}
        return {
            "object_type": object_type,
            "object_name": s(object_name),
            "table_name": s(table_name),
            "object_id": stable_id(obj, object_name),
            "is_hidden": hidden_value(obj),
            "data_type": s(get_any(obj, "dataType", "type", "data_type", default="")),
            "expression_language": "",
            "expression": "",
            "source_kind": "",
            "source_name": "",
            "source_expression": "",
            "direct_dependencies": "",
            "upstream_dependencies": "",
            "source_tables": "",
            "lineage_status": "not_scanned",
            "warnings": "",
            **kwargs,
        }

    # Named model expressions, often used for shared M expressions or parameters.
    for expr in expressions:
        name = s(get_any(expr, "name", "Name", default="(expression sans nom)"))
        language = s(get_any(expr, "kind", "language", "expressionKind", default=""))
        text_expr = get_expression_text(expr)
        row = new_row("Expression", name, obj=expr)
        row["expression_language"] = language or ("M/Power Query probable" if text_expr else "")
        row["expression"] = text_expr
        row["source_expression"] = text_expr
        row["source_kind"] = "model_expression"
        row["source_name"] = name
        row["lineage_status"] = "expression_exported"
        key = add_object("EXPR::" + norm_name(name), row)
        expression_index[norm_name(name)].append(key)

    # Tables, partitions, columns, and measures.
    for table in tables:
        tname = s(get_any(table, "name", "Name", default="(table sans nom)"))
        tnorm = norm_name(tname)
        table_index[tnorm] = table

        partitions = list_of_dicts(get_any(table, "partitions", "Partitions", default=[]))
        table_expression = ""
        table_source_kind = ""
        table_source_name = ""
        table_source_expression = ""
        partition_descriptions = []

        for partition in partitions:
            pname = s(get_any(partition, "name", "Name", default=""))
            source = get_any(partition, "source", "Source", default={})
            source = source if isinstance(source, dict) else {}
            ptype = s(get_any(source, "type", "Type", default=""))
            pexpr = get_any(source, "expression", "Expression", "query", "Query", default="")
            pexpr_text = s(pexpr)
            source_ref = s(get_any(source, "entityName", "entity", "expressionName", "queryName", default=""))
            if ptype.casefold() == "calculated" or (pexpr_text and "calculated" in ptype.casefold()):
                table_expression = pexpr_text or table_expression
                table_source_kind = "DAX calculated partition"
            elif ptype.casefold() in ("m", "powerquery", "power_query"):
                table_source_kind = "Power Query M partition"
                table_source_expression = pexpr_text or table_source_expression
                table_source_name = source_ref or pname or table_source_name
            elif ptype:
                table_source_kind = table_source_kind or ("partition:" + ptype)
            elif pexpr_text:
                table_source_expression = pexpr_text or table_source_expression
                table_source_kind = table_source_kind or "partition expression (type unspecified)"
            if pname or ptype or pexpr_text:
                partition_descriptions.append(
                    "partition=" + pname + "|type=" + ptype + "|source=" + source_ref
                )

        # Table-level expression appears in some schema variants.
        table_expression = table_expression or s(get_any(table, "expression", "calculatedExpression", default=""))
        if table_expression:
            table_source_kind = table_source_kind or "DAX calculated table"

        row = new_row("CalculatedTable" if table_expression else "Table", tname, obj=table)
        row["expression_language"] = "DAX" if table_expression else ("M" if table_source_expression else "")
        row["expression"] = table_expression
        row["source_kind"] = table_source_kind
        row["source_name"] = table_source_name
        row["source_expression"] = table_source_expression
        row["lineage_status"] = "pending"
        if partition_descriptions:
            row["warnings"] = " | ".join(partition_descriptions)
        if not partitions and not table_expression and not table_source_expression:
            row["warnings"] = (row["warnings"] + "; aucune partition/source explicite trouvée").strip("; ")
        table_key = add_object("TABLE::" + tnorm, row, aliases=(tname,))
        table_index[tnorm] = {"_row_key": table_key, "_raw": table}

        # Columns
        columns = list_of_dicts(get_any(table, "columns", "Columns", default=[]))
        for column in columns:
            cname = s(get_any(column, "name", "Name", default="(column sans nom)"))
            ctype = s(get_any(column, "type", "columnType", "Type", default=""))
            cexpr = s(get_any(column, "expression", "Expression", "formula", default=""))
            # A column can be calculated even when type is omitted.
            is_calc = bool(cexpr) or "calculated" in ctype.casefold()
            row = new_row("CalculatedColumn" if is_calc else "Column", cname, tname, column)
            row["expression_language"] = "DAX" if cexpr else ""
            row["expression"] = cexpr
            row["source_kind"] = "DAX calculated column" if is_calc else "model column"
            row["source_name"] = tname + "[" + cname + "]"
            row["lineage_status"] = "pending"
            ckey = add_object(
                "COLUMN::" + tnorm + "::" + norm_name(cname),
                row,
                aliases=(tname + "[" + cname + "]",)
            )
            column_index[(tnorm, norm_name(cname))].append(ckey)

        # Measures
        measures = list_of_dicts(get_any(table, "measures", "Measures", default=[]))
        for measure in measures:
            mname = s(get_any(measure, "name", "Name", default="(measure sans nom)"))
            mexpr = get_expression_text(measure)
            row = new_row("Measure", mname, tname, measure)
            row["expression_language"] = "DAX"
            row["expression"] = mexpr
            row["source_kind"] = "DAX measure"
            row["source_name"] = tname + "[" + mname + "]"
            row["lineage_status"] = "pending"
            mkey = add_object(
                "MEASURE::" + tnorm + "::" + norm_name(mname),
                row,
                aliases=(tname + "[" + mname + "]",)
            )
            measure_index[norm_name(mname)].append(mkey)

    # Also capture any top-level named expressions found in the model even if they
    # are not referenced by a table partition.
    if not tables:
        global_warnings.append("Aucune table trouvée sous model.tables.")

    # Build table name aliases and object lookup helpers.
    table_key_by_name = {}
    for key, row in objects.items():
        if row["object_type"] in ("Table", "CalculatedTable"):
            table_key_by_name[norm_name(row["object_name"])] = key

    def resolve_table_column(table_name, column_name):
        tnorm = norm_name(table_name)
        candidates = column_index.get((tnorm, norm_name(column_name)), [])
        if len(candidates) == 1:
            return candidates[0]
        # If column wasn't present in JSON, still attach the table dependency.
        tkey = table_key_by_name.get(tnorm)
        if tkey:
            return tkey
        return None

    def resolve_unqualified(name, current_table):
        n = norm_name(name)
        # Measures first: [Measure Name] is the normal syntax for a measure.
        candidates = measure_index.get(n, [])
        if len(candidates) == 1:
            return candidates[0]
        if len(candidates) > 1:
            # Prefer measure in the current table when possible.
            same = [k for k in candidates if norm_name(objects[k]["table_name"]) == norm_name(current_table)]
            if len(same) == 1:
                return same[0]
            return None
        # Column names without table qualification are uncommon and context-dependent.
        cols = []
        for (tn, cn), keys in column_index.items():
            if cn == n:
                cols.extend(keys)
        if len(cols) == 1:
            return cols[0]
        # Model expressions may be referenced by name in M, not usually DAX.
        exprs = expression_index.get(n, [])
        if len(exprs) == 1:
            return exprs[0]
        return None

    # Extract direct dependencies from DAX expressions and M/source references.
    dependency_graph = defaultdict(set)
    unresolved_by_key = defaultdict(list)
    for key, row in list(objects.items()):
        expr = row.get("expression", "")
        if row["object_type"] in ("Measure", "CalculatedColumn", "Table", "CalculatedTable"):
            tc_refs, bracket_refs = extract_dax_references(expr)
            direct = set()
            for ref_table, ref_column in tc_refs:
                resolved = resolve_table_column(ref_table, ref_column)
                if resolved:
                    direct.add(resolved)
                else:
                    unresolved_by_key[key].append(ref_table + "[" + ref_column + "]")
            for ref_name in bracket_refs:
                resolved = resolve_unqualified(ref_name, row.get("table_name", ""))
                if resolved and resolved != key:
                    direct.add(resolved)
                elif not resolved:
                    unresolved_by_key[key].append("[" + ref_name + "]")

            # A calculated table's formula often names tables without [Column] refs.
            # Infer table references from model table identifiers used as DAX arguments.
            if row["object_type"] in ("Table", "CalculatedTable") and expr:
                cleaned = clean_dax_for_scan(expr)
                for tname_norm, tkey in table_key_by_name.items():
                    if tkey == key:
                        continue
                    target_name = objects[tkey]["object_name"]
                    # Conservative word/name-boundary check; quoted table names supported.
                    escaped = re.escape(target_name)
                    patterns = [
                        r"(?<![A-Za-z0-9_])" + escaped + r"(?![A-Za-z0-9_])",
                        r"'" + re.escape(target_name.replace("'", "''")) + r"'",
                    ]
                    if any(re.search(pat, cleaned, flags=re.IGNORECASE) for pat in patterns):
                        direct.add(tkey)

            dependency_graph[key].update(direct)
            row["direct_dependencies"] = " | ".join(
                sorted((objects[d]["object_type"] + ":" +
                        (objects[d]["table_name"] + "[" if objects[d]["table_name"] else "") +
                        objects[d]["object_name"] +
                        ("]" if objects[d]["table_name"] else "")) for d in direct)
            )
            row["lineage_status"] = "parsed"
        elif row["object_type"] == "Expression":
            # M expressions are exported as sources, but M syntax needs a separate parser.
            row["lineage_status"] = "exported_M_or_model_expression"

    # Associate a table row with its columns and measures for context, but do not treat
    # ownership as a DAX dependency. This allows table lineage to list its contents.
    children_by_table = defaultdict(list)
    for key, row in objects.items():
        if row["object_type"] in ("Column", "CalculatedColumn", "Measure"):
            children_by_table[norm_name(row["table_name"])].append(key)

    def pretty_object(key):
        row = objects[key]
        if row["object_type"] in ("Column", "CalculatedColumn", "Measure"):
            return row["table_name"] + "[" + row["object_name"] + "]"
        return row["object_name"] + "{" + row["object_type"] + "}"

    def get_upstream(start_key):
        seen = set()
        stack = list(dependency_graph.get(start_key, set()))
        while stack:
            item = stack.pop()
            if item == start_key or item in seen:
                continue
            seen.add(item)
            stack.extend(dependency_graph.get(item, set()) - seen)
        return seen

    def source_tables_for(start_key):
        """Find upstream physical/model tables and M-source tables through dependencies."""
        result = set()
        visited = set()
        stack = [start_key]
        while stack:
            key = stack.pop()
            if key in visited:
                continue
            visited.add(key)
            row = objects.get(key)
            if not row:
                continue
            if row["object_type"] in ("Table", "CalculatedTable"):
                # Treat non-DAX tables as terminal source tables. Calculated tables are
                # retained too, but traversal continues through their DAX dependencies.
                if row.get("source_kind") != "DAX calculated partition" and not row.get("expression"):
                    result.add(row["object_name"])
                elif row.get("source_kind") == "Power Query M partition" or row.get("source_expression"):
                    result.add(row["object_name"] + " (M)")
                elif not dependency_graph.get(key):
                    result.add(row["object_name"])
            stack.extend(dependency_graph.get(key, set()))
        return result

    # Finalize recursive lineage and warnings.
    for key, row in objects.items():
        if row["object_type"] in ("Measure", "CalculatedColumn", "Table", "CalculatedTable"):
            upstream = get_upstream(key)
            row["upstream_dependencies"] = " | ".join(sorted(pretty_object(k) for k in upstream))
            row["source_tables"] = " | ".join(sorted(source_tables_for(key)))
            unresolved = unresolved_by_key.get(key, [])
            if unresolved:
                row["lineage_status"] = "partial_unresolved_references"
                row["warnings"] = (row.get("warnings", "") + "; unresolved: " +
                                   ", ".join(sorted(set(unresolved)))).strip("; ")
            elif row.get("expression"):
                row["lineage_status"] = "static_lineage_extracted"
            else:
                row["lineage_status"] = "no_expression_or_source"
        elif row["object_type"] in ("Column", "CalculatedColumn"):
            if row.get("expression"):
                row["source_tables"] = " | ".join(sorted(source_tables_for(key)))
            else:
                # Physical columns originate from their containing table.
                row["source_tables"] = row.get("table_name", "")
                row["lineage_status"] = "physical_column_no_DAX"
        elif row["object_type"] in ("Table", "CalculatedTable"):
            pass

    # Add extraction-wide warnings to every row so the one-CSV deliverable preserves them.
    warning_prefix = " | ".join(global_warnings)
    for row in rows:
        if warning_prefix:
            row["warnings"] = (warning_prefix + "; " + row.get("warnings", "")).strip("; ")
        row.pop("_key", None)

    # Stable output order: tables, expressions, columns, calculated columns, measures.
    type_order = {
        "Table": 0, "CalculatedTable": 1, "Expression": 2,
        "Column": 3, "CalculatedColumn": 4, "Measure": 5,
    }
    rows.sort(key=lambda r: (
        type_order.get(r["object_type"], 99),
        norm_name(r["table_name"]),
        norm_name(r["object_name"]),
    ))

    try:
        output_path.parent.mkdir(parents=True, exist_ok=True)
        with output_path.open("w", newline="", encoding="utf-8-sig") as f:
            writer = csv.DictWriter(f, fieldnames=CSV_FIELDS, extrasaction="ignore", delimiter=";")
            writer.writeheader()
            for row in rows:
                writer.writerow({field: row.get(field, "") for field in CSV_FIELDS})
    except OSError as exc:
        print("ERREUR lors de l'écriture du CSV: " + str(exc), file=sys.stderr)
        return 2

    counts = defaultdict(int)
    for row in rows:
        counts[row["object_type"]] += 1
    print("Extraction terminée.")
    print("Fichier source :", pbit_path)
    print("Schéma utilisé :", schema_member)
    print("CSV généré     :", output_path)
    print("Objets extraits:")
    for typ in ("Table", "CalculatedTable", "Expression", "Column", "CalculatedColumn", "Measure"):
        if counts[typ]:
            print(f"  - {typ}: {counts[typ]}")
    print("Séparateur CSV : ;")
    print("Encodage       : UTF-8 avec BOM (compatible Excel)")
    print("Note           : vérifie les lignes avec lineage_status=partial_unresolved_references.")
    return 0


if __name__ == "__main__":
    sys.exit(main())
