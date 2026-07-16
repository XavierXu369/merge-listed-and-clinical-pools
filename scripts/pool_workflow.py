#!/usr/bin/env python3
"""Deterministic workbook stages for merging listed and clinical molecule pools."""

from __future__ import annotations

import argparse
import hashlib
import json
import re
import shutil
import sys
from collections import Counter, defaultdict
from copy import copy
from difflib import SequenceMatcher
from pathlib import Path
from typing import Any, Iterable

try:
    from openpyxl import Workbook, load_workbook
    from openpyxl.styles import Alignment, Font, PatternFill
    from openpyxl.utils import get_column_letter
except ImportError as exc:  # pragma: no cover - environment failure
    raise SystemExit("This script requires openpyxl: python -m pip install openpyxl") from exc


META_SHEET = "__workflow_meta"
RELATION_HEADERS = [
    "判断编号",
    "临床来源行ID",
    "临床分子序号",
    "候选已上市来源行ID",
    "候选已上市分子序号",
    "临床登记号",
    "项目中国阶段",
    "品种中国阶段",
    "试验分期",
    "试验状态",
    "临床药品资产名称",
    "候选已上市药品资产名称",
    "候选已上市通用名",
    "临床剂型",
    "已上市剂型",
    "临床对应疾病",
    "已上市对应疾病",
    "临床相关适应症",
    "已上市相关适应症",
    "临床集团",
    "已上市集团",
    "名称匹配层级",
    "初始关系类型",
    "初始判断依据",
    "适应症核对结论",
    "Disease修正建议",
    "Disease修正值",
    "资产名称修正建议",
    "资产名称修正值",
    "修正后关系类型",
    "最终动作代码",
    "最终承接上市分子序号",
    "判断建议依据",
    "证据来源",
    "集中确认项",
    "判断状态",
    "主表同步状态",
    "阶段执行结果",
    "最终分子序号",
]

ALLOWED_ACTIONS = {
    "KEEP",
    "DELETE_LISTED_COVERED",
    "EXCLUDE_NON_THERAPEUTIC",
    "EXCLUDE_NON_FIXED_REGIMEN",
    "EXCLUDE_PRE_SCOPE_MARKETED",
    "HOLD",
}

DOSAGE_TERMS = sorted(
    {
        "缓释胶囊",
        "控释胶囊",
        "分散片",
        "缓释片",
        "控释片",
        "咀嚼片",
        "肠溶片",
        "口服溶液剂",
        "冻干粉针剂",
        "注射液",
        "注射剂",
        "胶囊剂",
        "颗粒剂",
        "混悬剂",
        "吸入剂",
        "喷雾剂",
        "贴剂",
        "乳膏",
        "软膏",
        "片剂",
        "胶囊",
        "颗粒",
        "口服液",
        "注射用",
        "片",
    },
    key=len,
    reverse=True,
)


def emit(payload: Any) -> None:
    print(json.dumps(payload, ensure_ascii=False, indent=2, default=str))


def load_json(path: str | Path) -> dict[str, Any]:
    p = Path(path)
    if not p.is_file():
        raise ValueError(f"JSON file was not found: {p}")
    with p.open("r", encoding="utf-8-sig") as handle:
        value = json.load(handle)
    if not isinstance(value, dict):
        raise ValueError(f"JSON root must be an object: {p}")
    return value


def require_new_output(output: str | Path, inputs: Iterable[str | Path] = ()) -> Path:
    out = Path(output).expanduser().resolve()
    if out.exists():
        raise ValueError(f"Output already exists and will not be overwritten: {out}")
    for item in inputs:
        if out == Path(item).expanduser().resolve():
            raise ValueError("Output path must differ from every input path.")
    if not out.parent.is_dir():
        raise ValueError(f"Output directory does not exist: {out.parent}")
    if out.suffix.lower() != ".xlsx":
        raise ValueError("Output workbook must use .xlsx.")
    return out


def is_blank(value: Any) -> bool:
    return value is None or (isinstance(value, str) and not value.strip())


def display_value(value: Any, missing: str) -> Any:
    return missing if is_blank(value) else value


def clean_text(value: Any, missing: str = "—") -> str:
    if value is None:
        return ""
    text = str(value).strip()
    if text in {"", missing, "-", "--"}:
        return ""
    return text


def stable_key(value: Any) -> str:
    if value is None:
        return ""
    if isinstance(value, float) and value.is_integer():
        return str(int(value))
    return str(value).strip()


def parse_carrier_sequences(value: Any) -> list[str]:
    text = stable_key(value)
    if not text:
        return []
    return [item for item in (part.strip() for part in re.split(r"[;；,/，\s]+", text)) if item]


def used_bounds(ws) -> tuple[int, int]:
    last_row = 1
    last_col = 1
    for row in ws.iter_rows():
        row_has_value = False
        for cell in row:
            if cell.value is not None:
                row_has_value = True
                last_col = max(last_col, cell.column)
        if row_has_value:
            last_row = max(last_row, row[0].row)
    return last_row, last_col


def header_map(ws) -> tuple[list[str], dict[str, int]]:
    _, last_col = used_bounds(ws)
    headers: list[str] = []
    mapping: dict[str, int] = {}
    for col in range(1, last_col + 1):
        name = clean_text(ws.cell(1, col).value, missing="")
        if not name:
            raise ValueError(f"Blank header at {ws.title}!{get_column_letter(col)}1")
        if name in mapping:
            raise ValueError(f"Duplicate header in {ws.title}: {name}")
        headers.append(name)
        mapping[name] = col
    return headers, mapping


def records_from_sheet(ws) -> tuple[list[str], dict[str, int], list[dict[str, Any]]]:
    headers, mapping = header_map(ws)
    last_row, _ = used_bounds(ws)
    records: list[dict[str, Any]] = []
    for row_number in range(2, last_row + 1):
        values = [ws.cell(row_number, col).value for col in range(1, len(headers) + 1)]
        if all(is_blank(value) for value in values):
            continue
        record = {header: values[index] for index, header in enumerate(headers)}
        record["__row_number__"] = row_number
        records.append(record)
    return headers, mapping, records


def sheet_fingerprint(ws) -> str:
    last_row, last_col = used_bounds(ws)
    digest = hashlib.sha256()
    digest.update(ws.title.encode("utf-8"))
    for row in range(1, last_row + 1):
        for col in range(1, last_col + 1):
            cell = ws.cell(row, col)
            if is_blank(cell.value):
                payload = f"{row}|{col}|BLANK\n"
            else:
                payload = f"{row}|{col}|{cell.data_type}|{type(cell.value).__name__}|{repr(cell.value)}\n"
            digest.update(payload.encode("utf-8", errors="replace"))
    for merged in sorted(str(item) for item in ws.merged_cells.ranges):
        digest.update(f"MERGED|{merged}\n".encode("utf-8"))
    return digest.hexdigest()


def formula_stats(ws) -> tuple[int, int]:
    last_row, last_col = used_bounds(ws)
    formulas = 0
    external = 0
    for row in range(1, last_row + 1):
        for col in range(1, last_col + 1):
            value = ws.cell(row, col).value
            if isinstance(value, str) and value.startswith("="):
                formulas += 1
                if "[" in value:
                    external += 1
    return formulas, external


def validate_config_shape(config: dict[str, Any]) -> None:
    for section in ("listed", "clinical", "output", "final_fields", "run_schema"):
        if section not in config:
            raise ValueError(f"Config is missing section: {section}")
    if not isinstance(config["run_schema"], list) or not config["run_schema"]:
        raise ValueError("run_schema must be a nonempty list.")
    names = [item.get("name") for item in config["run_schema"]]
    if any(not name for name in names) or len(names) != len(set(names)):
        raise ValueError("Run Schema field names must be nonblank and unique.")
    valid_modes = {"sequence", "constant", "column", "column_or_missing", "coalesce_columns", "missing"}
    for item in config["run_schema"]:
        for source in ("listed", "clinical"):
            rule = item.get(source)
            if not isinstance(rule, dict) or rule.get("mode") not in valid_modes:
                raise ValueError(f"Invalid {source} mapping rule for field: {item.get('name')}")
            if rule["mode"] in {"column", "column_or_missing"} and not rule.get("column"):
                raise ValueError(f"Column mapping lacks a column name: {item.get('name')} / {source}")
            if rule["mode"] == "coalesce_columns":
                columns = rule.get("columns")
                if not isinstance(columns, list) or not columns or any(not clean_text(column, missing="") for column in columns):
                    raise ValueError(f"Coalesced mapping needs a nonempty columns list: {item.get('name')} / {source}")
    final_names = set(names)
    required_logical_fields = {
        "sequence",
        "asset",
        "indication",
        "disease",
        "registration_no",
        "project_stage",
        "product_stage",
        "trial_phase",
        "trial_status",
        "common_name",
        "dosage_form",
        "group",
    }
    missing_logical_fields = sorted(required_logical_fields - set(config["final_fields"]))
    if missing_logical_fields:
        raise ValueError("final_fields is missing logical fields: " + ", ".join(missing_logical_fields))
    for logical, field in config["final_fields"].items():
        if field not in final_names:
            raise ValueError(f"final_fields.{logical} points outside Run Schema: {field}")


def source_required_fields(config: dict[str, Any], source: str) -> list[str]:
    section = config[source]
    logical = ["asset", "indication", "disease", "dosage_form"]
    if source == "clinical":
        logical.extend(["project_stage", "product_stage", "trial_phase", "trial_status"])
    required = [section.get(item, "") for item in logical]
    source_id = section.get("source_id", "")
    if source_id:
        required.append(source_id)
    for item in config["run_schema"]:
        rule = item[source]
        if rule["mode"] == "column":
            required.append(rule["column"])
    return sorted({item for item in required if item})


def inspect_one_source(config: dict[str, Any], source: str) -> tuple[dict[str, Any], list[str]]:
    section = config[source]
    path = Path(section["file"])
    blockers: list[str] = []
    if not path.is_file():
        return {"file": str(path), "sheet": section.get("sheet"), "readable": False}, [f"Missing {source} file: {path}"]
    if path.suffix.lower() not in {".xlsx", ".xlsm"}:
        blockers.append(f"Unsupported {source} workbook extension: {path.suffix}")
    wb = load_workbook(path, data_only=False, read_only=False)
    sheet_name = section["sheet"]
    if sheet_name not in wb.sheetnames:
        return {"file": str(path), "sheet": sheet_name, "readable": False, "sheets": wb.sheetnames}, blockers + [f"Missing {source} sheet: {sheet_name}"]
    ws = wb[sheet_name]
    try:
        headers, _, records = records_from_sheet(ws)
    except ValueError as exc:
        return {"file": str(path), "sheet": sheet_name, "readable": False}, blockers + [str(exc)]
    missing_required = [field for field in source_required_fields(config, source) if field not in headers]
    blockers.extend(f"{source} missing required field: {field}" for field in missing_required)
    for item in config["run_schema"]:
        rule = item[source]
        if rule["mode"] == "coalesce_columns" and not any(column in headers for column in rule["columns"]):
            blockers.append(
                f"{source} lacks every fallback column for {item['name']}: {', '.join(rule['columns'])}"
            )

    source_id = section.get("source_id", "")
    blank_ids = 0
    duplicate_ids: list[str] = []
    generated_ids = not source_id or source_id not in headers
    if not generated_ids:
        values = [stable_key(record.get(source_id)) for record in records]
        blank_ids = sum(not value for value in values)
        counts = Counter(value for value in values if value)
        duplicate_ids = [value for value, count in counts.items() if count > 1]
        if blank_ids:
            blockers.append(f"{source} has {blank_ids} blank source IDs")
        if duplicate_ids:
            blockers.append(f"{source} has {len(duplicate_ids)} duplicate source IDs")

    blank_by_field: dict[str, int] = {}
    for logical in ("asset", "indication", "disease", "dosage_form"):
        field = section.get(logical, "")
        if field in headers:
            blank_by_field[field] = sum(is_blank(record.get(field)) for record in records)
    disease_field = section.get("disease", "")
    multi_disease_rows: list[int] = []
    if disease_field in headers:
        for record in records:
            value = clean_text(record.get(disease_field))
            if re.search(r"[;；\r\n]", value):
                multi_disease_rows.append(record["__row_number__"])
        if blank_by_field.get(disease_field, 0):
            blockers.append(f"{source} has blank Disease values: {blank_by_field[disease_field]}")
        if multi_disease_rows:
            blockers.append(f"{source} has unresolved multivalue Disease rows: {len(multi_disease_rows)}")

    duplicate_signatures: dict[str, list[int]] = defaultdict(list)
    asset = section.get("asset", "")
    form = section.get("dosage_form", "")
    if all(field in headers for field in (asset, form, disease_field)):
        for record in records:
            signature = " | ".join(clean_text(record.get(field)).casefold() for field in (asset, form, disease_field))
            duplicate_signatures[signature].append(record["__row_number__"])
    duplicate_signatures = {key: rows for key, rows in duplicate_signatures.items() if key.strip(" |") and len(rows) > 1}
    if duplicate_signatures and not section.get("within_source_duplicates_approved", False):
        blockers.append(
            f"{source} has {len(duplicate_signatures)} within-source asset/form/Disease duplicate signatures that require upstream resolution or explicit approval"
        )

    status_counts: dict[str, dict[str, int]] = {}
    if source == "clinical":
        for logical in ("project_stage", "product_stage", "trial_phase", "trial_status"):
            field = section.get(logical, "")
            if field in headers:
                status_counts[field] = dict(Counter(clean_text(record.get(field)) for record in records))
    formulas, external = formula_stats(ws)
    last_row, last_col = used_bounds(ws)
    result = {
        "file": str(path.resolve()),
        "sheet": sheet_name,
        "readable": True,
        "data_rows": len(records),
        "used_rows_including_header": last_row,
        "columns": last_col,
        "headers": headers,
        "missing_required_fields": missing_required,
        "source_id_field": source_id or None,
        "generated_source_ids": generated_ids,
        "blank_source_ids": blank_ids,
        "duplicate_source_ids": duplicate_ids,
        "blank_key_fields": blank_by_field,
        "multivalue_disease_rows": multi_disease_rows,
        "within_source_duplicate_signature_count": len(duplicate_signatures),
        "within_source_duplicate_signatures": duplicate_signatures,
        "status_counts": status_counts,
        "formula_cells": formulas,
        "potential_external_formula_cells": external,
        "fingerprint": sheet_fingerprint(ws),
    }
    return result, blockers


def inspect_config(config: dict[str, Any]) -> dict[str, Any]:
    validate_config_shape(config)
    listed, listed_blockers = inspect_one_source(config, "listed")
    clinical, clinical_blockers = inspect_one_source(config, "clinical")
    run_schema = [item["name"] for item in config["run_schema"]]
    return {
        "state": "PREFLIGHT_PASSED" if not (listed_blockers + clinical_blockers) else "PREFLIGHT_FAILED",
        "ta_name": config.get("ta_name"),
        "schema_approved": bool(config.get("schema_approved")),
        "run_schema_field_count": len(run_schema),
        "run_schema": run_schema,
        "listed": listed,
        "clinical": clinical,
        "blockers": listed_blockers + clinical_blockers,
    }


def copy_sheet(source_ws, target_wb, target_name: str):
    target = target_wb.create_sheet(target_name)
    last_row, last_col = used_bounds(source_ws)
    for row in range(1, last_row + 1):
        for col in range(1, last_col + 1):
            source_cell = source_ws.cell(row, col)
            target_cell = target.cell(row, col, source_cell.value)
            if source_cell.has_style:
                target_cell.font = copy(source_cell.font)
                target_cell.fill = copy(source_cell.fill)
                target_cell.border = copy(source_cell.border)
                target_cell.alignment = copy(source_cell.alignment)
                target_cell.protection = copy(source_cell.protection)
                target_cell.number_format = source_cell.number_format
            if source_cell.hyperlink:
                target_cell._hyperlink = copy(source_cell.hyperlink)
            if source_cell.comment:
                target_cell.comment = copy(source_cell.comment)
    for key, dimension in source_ws.column_dimensions.items():
        target.column_dimensions[key].width = dimension.width
        target.column_dimensions[key].hidden = dimension.hidden
        target.column_dimensions[key].outlineLevel = dimension.outlineLevel
    for key, dimension in source_ws.row_dimensions.items():
        target.row_dimensions[key].height = dimension.height
        target.row_dimensions[key].hidden = dimension.hidden
        target.row_dimensions[key].outlineLevel = dimension.outlineLevel
    for merged in source_ws.merged_cells.ranges:
        target.merge_cells(str(merged))
    target.freeze_panes = source_ws.freeze_panes
    target.auto_filter.ref = source_ws.auto_filter.ref
    target.sheet_view.showGridLines = source_ws.sheet_view.showGridLines
    return target


def style_table(ws, header_fill: str = "1F4E78") -> None:
    if ws.max_column < 1:
        return
    for cell in ws[1]:
        cell.font = Font(bold=True, color="FFFFFF")
        cell.fill = PatternFill("solid", fgColor=header_fill)
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    for row in ws.iter_rows(min_row=2):
        for cell in row:
            cell.alignment = Alignment(vertical="top", wrap_text=True)
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions
    for col in range(1, ws.max_column + 1):
        width = max(10, min(32, max(len(str(ws.cell(row, col).value or "")) for row in range(1, min(ws.max_row, 50) + 1)) + 2))
        ws.column_dimensions[get_column_letter(col)].width = width


def get_source_id(record: dict[str, Any], section: dict[str, Any], raw_sheet: str) -> str:
    source_id = section.get("source_id", "")
    if source_id and not is_blank(record.get(source_id)):
        return stable_key(record[source_id])
    return f"{raw_sheet}!R{record['__row_number__']}"


def mapped_value(rule: dict[str, Any], record: dict[str, Any], sequence: int, missing: str) -> Any:
    mode = rule["mode"]
    if mode == "sequence":
        return sequence
    if mode == "constant":
        return display_value(rule.get("value"), missing)
    if mode == "missing":
        return missing
    if mode == "coalesce_columns":
        for column in rule["columns"]:
            if column in record and not is_blank(record.get(column)):
                return display_value(record.get(column), missing)
        return missing
    column = rule["column"]
    if column not in record:
        if mode == "column_or_missing":
            return missing
        raise ValueError(f"Required mapped column is missing: {column}")
    return display_value(record.get(column), missing)


def command_inspect(args) -> None:
    config = load_json(args.config)
    emit(inspect_config(config))


def command_build_draft(args) -> None:
    config = load_json(args.config)
    preflight = inspect_config(config)
    if preflight["blockers"]:
        raise ValueError("Preflight failed: " + " | ".join(preflight["blockers"]))
    if not args.confirmed_schema or not config.get("schema_approved"):
        raise ValueError("Draft generation requires explicit Schema approval and schema_approved=true.")
    listed_path = Path(config["listed"]["file"])
    clinical_path = Path(config["clinical"]["file"])
    output = require_new_output(args.output, [listed_path, clinical_path])
    listed_wb = load_workbook(listed_path, data_only=False)
    clinical_wb = load_workbook(clinical_path, data_only=False)
    listed_ws = listed_wb[config["listed"]["sheet"]]
    clinical_ws = clinical_wb[config["clinical"]["sheet"]]
    listed_headers, _, listed_records = records_from_sheet(listed_ws)
    clinical_headers, _, clinical_records = records_from_sheet(clinical_ws)

    wb = Workbook()
    wb.remove(wb.active)
    draft_name = config["output"]["draft_sheet"]
    draft = wb.create_sheet(draft_name)
    run_schema = config["run_schema"]
    headers = [item["name"] for item in run_schema]
    draft.append(headers)
    meta = wb.create_sheet(META_SHEET)
    meta.append(["数据来源", "待判断分子序号", "来源行ID", "原始Sheet行号"])
    missing = config.get("missing_display", "—")
    listed_source = config["output"]["listed_source_value"]
    clinical_source = config["output"]["clinical_source_value"]
    sequence = 1

    for source, records, section, source_value in (
        ("listed", listed_records, config["listed"], listed_source),
        ("clinical", clinical_records, config["clinical"], clinical_source),
    ):
        raw_name = section["raw_output_sheet"]
        for record in records:
            row = [mapped_value(item[source], record, sequence, missing) for item in run_schema]
            draft.append(row)
            meta.append([source_value, sequence, get_source_id(record, section, raw_name), record["__row_number__"]])
            sequence += 1

    listed_raw = copy_sheet(listed_ws, wb, config["listed"]["raw_output_sheet"])
    clinical_raw = copy_sheet(clinical_ws, wb, config["clinical"]["raw_output_sheet"])
    meta["F1"] = "state"
    meta["G1"] = "DRAFT_BUILT"
    meta["F2"] = "listed_raw_fingerprint"
    meta["G2"] = sheet_fingerprint(listed_raw)
    meta["F3"] = "clinical_raw_fingerprint"
    meta["G3"] = sheet_fingerprint(clinical_raw)
    meta["F4"] = "run_schema"
    meta["G4"] = json.dumps(headers, ensure_ascii=False)
    meta.sheet_state = "hidden"
    style_table(draft)
    wb.save(output)

    verify = load_workbook(output, data_only=False)
    verify_draft = verify[draft_name]
    _, _, verify_rows = records_from_sheet(verify_draft)
    listed_fp = sheet_fingerprint(verify[config["listed"]["raw_output_sheet"]])
    clinical_fp = sheet_fingerprint(verify[config["clinical"]["raw_output_sheet"]])
    expected = len(listed_records) + len(clinical_records)
    if len(verify_rows) != expected:
        output.unlink(missing_ok=True)
        raise ValueError("Draft row-count verification failed.")
    if listed_fp != meta["G2"].value or clinical_fp != meta["G3"].value:
        output.unlink(missing_ok=True)
        raise ValueError("Raw-sheet fingerprint verification failed after draft generation.")
    emit(
        {
            "state": "DRAFT_BUILT",
            "output": str(output),
            "listed_rows": len(listed_records),
            "clinical_rows": len(clinical_records),
            "draft_rows": len(verify_rows),
            "run_schema_fields": len(headers),
            "raw_sheets_preserved": True,
            "hidden_meta_sheet": META_SHEET,
        }
    )


def normalize_name(value: Any, missing: str) -> str:
    text = clean_text(value, missing).casefold().replace("＋", "+")
    return re.sub(r"[\s·•,，;；:：()（）\[\]【】{}\-_/\\+]", "", text)


def strip_dosage_form(value: Any, dosage_form: Any, missing: str) -> str:
    base = normalize_name(value, missing)
    form = normalize_name(dosage_form, missing)
    if form and base.endswith(form) and len(base) > len(form):
        base = base[: -len(form)]
    for term in DOSAGE_TERMS:
        normalized_term = normalize_name(term, missing)
        if normalized_term and base.endswith(normalized_term) and len(base) > len(normalized_term):
            base = base[: -len(normalized_term)]
            break
    return base


def aliases(asset: Any, common: Any, dosage_form: Any, missing: str) -> set[str]:
    result = {
        normalize_name(asset, missing),
        strip_dosage_form(asset, dosage_form, missing),
        normalize_name(common, missing),
        strip_dosage_form(common, dosage_form, missing),
    }
    return {item for item in result if item}


def components(asset: Any, dosage_form: Any, missing: str) -> set[str]:
    raw = clean_text(asset, missing)
    parts = re.split(r"[+＋;；、,，\r\n/&]", raw)
    result = {strip_dosage_form(part, dosage_form, missing) for part in parts}
    return {item for item in result if item}


def normalized_equal(left: Any, right: Any, missing: str) -> bool:
    lvalue = normalize_name(left, missing)
    rvalue = normalize_name(right, missing)
    return bool(lvalue and rvalue and lvalue == rvalue)


def read_meta(ws) -> tuple[dict[str, dict[str, str]], dict[str, str]]:
    mapping: dict[str, dict[str, str]] = {}
    last_row, _ = used_bounds(ws)
    for row in range(2, last_row + 1):
        source = clean_text(ws.cell(row, 1).value)
        sequence = stable_key(ws.cell(row, 2).value)
        if source and sequence:
            mapping[sequence] = {
                "source": source,
                "source_id": stable_key(ws.cell(row, 3).value),
                "raw_row": stable_key(ws.cell(row, 4).value),
            }
    state: dict[str, str] = {}
    for row in range(1, max(last_row, 10) + 1):
        key = clean_text(ws.cell(row, 6).value)
        if key:
            state[key] = clean_text(ws.cell(row, 7).value, missing="")
    return mapping, state


def record_source(
    record: dict[str, Any],
    fields: dict[str, str],
    meta_map: dict[str, dict[str, str]],
) -> str:
    """Return source provenance without requiring a visible output column."""
    source_field = fields.get("source", "")
    if source_field and source_field in record:
        source = clean_text(record.get(source_field), missing="")
        if source:
            return source
    sequence = stable_key(record.get(fields["sequence"]))
    return clean_text(meta_map.get(sequence, {}).get("source"), missing="")


def get_record(record: dict[str, Any], field_name: str) -> Any:
    return record.get(field_name, "")


def direct_candidates(clinical: dict[str, Any], listed: list[dict[str, Any]], fields: dict[str, str], missing: str):
    c_alias = aliases(get_record(clinical, fields["asset"]), "", get_record(clinical, fields["dosage_form"]), missing)
    result = []
    for listed_record in listed:
        l_alias = aliases(
            get_record(listed_record, fields["asset"]),
            get_record(listed_record, fields["common_name"]),
            get_record(listed_record, fields["dosage_form"]),
            missing,
        )
        if c_alias & l_alias:
            result.append((listed_record, "direct", 1.0))
    return result


def partial_candidates(clinical: dict[str, Any], listed: list[dict[str, Any]], fields: dict[str, str], missing: str):
    c_parts = components(get_record(clinical, fields["asset"]), get_record(clinical, fields["dosage_form"]), missing)
    result = []
    for listed_record in listed:
        l_alias = aliases(
            get_record(listed_record, fields["asset"]),
            get_record(listed_record, fields["common_name"]),
            get_record(listed_record, fields["dosage_form"]),
            missing,
        )
        if c_parts & l_alias:
            result.append((listed_record, "partial", 0.8))
    return result


def approximate_candidates(
    clinical: dict[str, Any],
    listed: list[dict[str, Any]],
    fields: dict[str, str],
    missing: str,
    threshold: float,
    limit: int,
):
    clinical_name = strip_dosage_form(get_record(clinical, fields["asset"]), get_record(clinical, fields["dosage_form"]), missing)
    if len(clinical_name) < 4:
        return []
    scored = []
    for listed_record in listed:
        l_alias = aliases(
            get_record(listed_record, fields["asset"]),
            get_record(listed_record, fields["common_name"]),
            get_record(listed_record, fields["dosage_form"]),
            missing,
        )
        score = max((SequenceMatcher(None, clinical_name, alias).ratio() for alias in l_alias if len(alias) >= 4), default=0.0)
        if score >= threshold:
            scored.append((listed_record, "approximate", score))
    return sorted(scored, key=lambda item: item[2], reverse=True)[:limit]


def initial_relation(clinical: dict[str, Any], listed: dict[str, Any] | None, match_type: str, fields: dict[str, str], missing: str) -> tuple[str, str, str]:
    if listed is None:
        stages = " ".join(
            clean_text(get_record(clinical, fields[item]), missing)
            for item in ("project_stage", "product_stage")
            if item in fields
        )
        relation = "已上市但当前上市池无直接候选" if "已上市" in stages else "临床独有（无直接候选）"
        return "无直接名称候选", relation, "No listed product/generic candidate was generated; review scope and marketed-gap evidence if relevant."
    same_form = normalized_equal(get_record(clinical, fields["dosage_form"]), get_record(listed, fields["dosage_form"]), missing)
    same_disease = clean_text(get_record(clinical, fields["disease"]), missing) == clean_text(get_record(listed, fields["disease"]), missing)
    if match_type == "partial":
        return "部分成分与上市资产一致", "部分成分重合，非同一资产", "At least one normalized clinical component matches a listed product or generic name; inspect the complete component set."
    if match_type == "approximate":
        return "名称近似但关键表达可能不同", "近似名称/盐型或化学形式差异", "Name similarity generated this review candidate; do not treat it as identity evidence."
    if same_form and same_disease:
        relation = "高度疑似重复"
    elif same_form:
        relation = "同分子同剂型，不同疾病"
    elif same_disease:
        relation = "同分子不同剂型，同疾病"
    else:
        relation = "同分子，剂型与疾病均不同"
    return "名称与上市药品/通用名直接一致", relation, "Normalized clinical asset matched a listed product or generic name; dosage form and Disease were compared separately."


def command_build_candidates(args) -> None:
    config = load_json(args.config)
    validate_config_shape(config)
    if not config.get("schema_approved"):
        raise ValueError("Candidate generation requires an approved Run Schema.")
    source = Path(args.workbook)
    if not source.is_file():
        raise ValueError(f"Draft workbook was not found: {source}")
    output = require_new_output(args.output, [source])
    shutil.copy2(source, output)
    wb = load_workbook(output, data_only=False)
    draft_name = config["output"]["draft_sheet"]
    relation_name = config["output"]["relation_sheet"]
    if draft_name not in wb.sheetnames or META_SHEET not in wb.sheetnames:
        output.unlink(missing_ok=True)
        raise ValueError("Workbook lacks the draft or workflow metadata sheet.")
    if relation_name in wb.sheetnames:
        output.unlink(missing_ok=True)
        raise ValueError(f"Relation sheet already exists: {relation_name}")
    draft = wb[draft_name]
    headers, _, records = records_from_sheet(draft)
    expected_headers = [item["name"] for item in config["run_schema"]]
    if headers != expected_headers:
        output.unlink(missing_ok=True)
        raise ValueError("Draft headers do not equal the approved Run Schema.")
    meta_map, _ = read_meta(wb[META_SHEET])
    fields = config["final_fields"]
    sequence_field = fields["sequence"]
    listed_source = config["output"]["listed_source_value"]
    clinical_source = config["output"]["clinical_source_value"]
    listed = [record for record in records if record_source(record, fields, meta_map) == listed_source]
    clinical = [record for record in records if record_source(record, fields, meta_map) == clinical_source]
    missing = config.get("missing_display", "—")
    threshold = float(config.get("candidate_similarity_threshold", 0.78))
    approx_limit = int(config.get("approximate_candidate_limit", 5))
    relation = wb.create_sheet(relation_name, 1)
    relation.append(RELATION_HEADERS)
    relation_index = {name: index for index, name in enumerate(RELATION_HEADERS)}
    relation_rows = 0
    one_to_many = 0
    relation_counts: Counter[str] = Counter()

    for clinical_record in clinical:
        candidates = direct_candidates(clinical_record, listed, fields, missing)
        if not candidates:
            candidates = partial_candidates(clinical_record, listed, fields, missing)
        if not candidates:
            candidates = approximate_candidates(clinical_record, listed, fields, missing, threshold, approx_limit)
        if not candidates:
            candidates = [(None, "none", 0.0)]
        if len(candidates) > 1:
            one_to_many += 1
        clinical_sequence = stable_key(clinical_record[sequence_field])
        clinical_meta = meta_map.get(clinical_sequence, {})
        for listed_record, match_type, score in candidates:
            relation_rows += 1
            match_level, relation_type, basis = initial_relation(clinical_record, listed_record, match_type, fields, missing)
            if match_type == "approximate":
                basis += f" Similarity score={score:.3f}."
            relation_counts[relation_type] += 1
            row = [""] * len(RELATION_HEADERS)

            def put(name: str, value: Any) -> None:
                row[relation_index[name]] = "" if value is None else value

            put("判断编号", relation_rows)
            put("临床来源行ID", clinical_meta.get("source_id", f"临床序号:{clinical_sequence}"))
            put("临床分子序号", clinical_record[sequence_field])
            put("临床登记号", get_record(clinical_record, fields["registration_no"]))
            put("项目中国阶段", get_record(clinical_record, fields["project_stage"]))
            put("品种中国阶段", get_record(clinical_record, fields["product_stage"]))
            put("试验分期", get_record(clinical_record, fields["trial_phase"]))
            put("试验状态", get_record(clinical_record, fields["trial_status"]))
            put("临床药品资产名称", get_record(clinical_record, fields["asset"]))
            put("临床剂型", get_record(clinical_record, fields["dosage_form"]))
            put("临床对应疾病", get_record(clinical_record, fields["disease"]))
            put("临床相关适应症", get_record(clinical_record, fields["indication"]))
            put("临床集团", get_record(clinical_record, fields["group"]))
            if listed_record is not None:
                listed_sequence = stable_key(listed_record[sequence_field])
                listed_meta = meta_map.get(listed_sequence, {})
                put("候选已上市来源行ID", listed_meta.get("source_id", f"上市序号:{listed_sequence}"))
                put("候选已上市分子序号", listed_record[sequence_field])
                put("候选已上市药品资产名称", get_record(listed_record, fields["asset"]))
                put("候选已上市通用名", get_record(listed_record, fields["common_name"]))
                put("已上市剂型", get_record(listed_record, fields["dosage_form"]))
                put("已上市对应疾病", get_record(listed_record, fields["disease"]))
                put("已上市相关适应症", get_record(listed_record, fields["indication"]))
                put("已上市集团", get_record(listed_record, fields["group"]))
            put("名称匹配层级", match_level)
            put("初始关系类型", relation_type)
            put("初始判断依据", basis)
            put("判断状态", "待判断")
            relation.append(row)

    style_table(relation, header_fill="7030A0")
    wb[META_SHEET]["G1"] = "RELATIONS_BUILT"
    wb.save(output)
    emit(
        {
            "state": "RELATIONS_BUILT",
            "output": str(output),
            "listed_rows": len(listed),
            "unique_clinical_rows": len(clinical),
            "candidate_rows": relation_rows,
            "one_to_many_clinical_rows": one_to_many,
            "initial_relation_types": dict(relation_counts),
            "automatic_deletions": 0,
        }
    )


def relation_header_map(ws) -> dict[str, int]:
    headers, mapping = header_map(ws)
    missing = [header for header in RELATION_HEADERS if header not in mapping]
    if missing:
        raise ValueError("Relation sheet lacks required columns: " + ", ".join(missing))
    return mapping


def command_import_decisions(args) -> None:
    source = Path(args.workbook)
    if not source.is_file():
        raise ValueError(f"Candidate workbook was not found: {source}")
    output = require_new_output(args.output, [source])
    payload = load_json(args.decisions)
    decisions = payload.get("decisions")
    if not isinstance(decisions, list):
        raise ValueError("Decision JSON must contain a decisions list.")
    shutil.copy2(source, output)
    wb = load_workbook(output, data_only=False)
    relation_name = next((name for name in wb.sheetnames if name == "跨库关系判断"), None)
    if relation_name is None:
        output.unlink(missing_ok=True)
        raise ValueError("Relation sheet was not found.")
    ws = wb[relation_name]
    mapping = relation_header_map(ws)
    last_row, _ = used_bounds(ws)
    by_source: dict[str, list[int]] = defaultdict(list)
    by_sequence: dict[str, list[int]] = defaultdict(list)
    for row in range(2, last_row + 1):
        by_source[stable_key(ws.cell(row, mapping["临床来源行ID"]).value)].append(row)
        by_sequence[stable_key(ws.cell(row, mapping["临床分子序号"]).value)].append(row)
    assigned: set[str] = set()
    field_map = {
        "indication_check": "适应症核对结论",
        "disease_advice": "Disease修正建议",
        "disease_correction": "Disease修正值",
        "asset_name_advice": "资产名称修正建议",
        "asset_name_correction": "资产名称修正值",
        "revised_relation_type": "修正后关系类型",
        "final_action": "最终动作代码",
        "carrier_listed_sequence": "最终承接上市分子序号",
        "rationale": "判断建议依据",
        "evidence_source": "证据来源",
        "confirmation_group": "集中确认项",
        "decision_status": "判断状态",
    }
    for decision in decisions:
        if not isinstance(decision, dict):
            raise ValueError("Each decision must be an object.")
        source_id = stable_key(decision.get("clinical_source_id"))
        sequence = stable_key(decision.get("clinical_sequence"))
        source_rows = by_source.get(source_id, []) if source_id else []
        sequence_rows = by_sequence.get(sequence, []) if sequence else []
        if source_id and sequence and set(source_rows) != set(sequence_rows):
            raise ValueError(f"Decision source ID and sequence identify different groups: {source_id} / {sequence}")
        rows = source_rows or sequence_rows
        if not rows:
            raise ValueError(f"Decision did not match a clinical group: {source_id or sequence}")
        group_key = stable_key(ws.cell(rows[0], mapping["临床分子序号"]).value)
        if group_key in assigned:
            raise ValueError(f"Duplicate decision for clinical sequence: {group_key}")
        assigned.add(group_key)
        action = clean_text(decision.get("final_action"), missing="")
        if action and action not in ALLOWED_ACTIONS:
            raise ValueError(f"Unsupported action: {action}")
        for row in rows:
            for source_field, relation_field in field_map.items():
                if source_field in decision:
                    ws.cell(row, mapping[relation_field]).value = decision[source_field]
            has_correction = bool(clean_text(decision.get("disease_correction"), "") or clean_text(decision.get("asset_name_correction"), ""))
            ws.cell(row, mapping["主表同步状态"]).value = "待同步" if has_correction else "无需修正"
    wb[META_SHEET]["G1"] = "DECISIONS_PENDING"
    wb.save(output)
    unique_sequences = set(by_sequence)
    pending = unique_sequences - assigned
    emit(
        {
            "state": "DECISIONS_CLOSED" if not pending else "DECISIONS_PENDING",
            "output": str(output),
            "unique_clinical_rows": len(unique_sequences),
            "decisions_imported": len(assigned),
            "pending_unique_clinical_rows": len(pending),
            "pending_sequences": sorted(pending),
        }
    )


def unique_field_values(ws, rows: list[int], col: int) -> list[str]:
    values = []
    seen = set()
    for row in rows:
        value = clean_text(ws.cell(row, col).value, missing="")
        if value not in seen:
            seen.add(value)
            values.append(value)
    return values


def finalization_preview(config: dict[str, Any], workbook: str | Path) -> tuple[dict[str, Any], dict[str, Any]]:
    wb = load_workbook(workbook, data_only=False)
    draft_name = config["output"]["draft_sheet"]
    relation_name = config["output"]["relation_sheet"]
    if draft_name not in wb.sheetnames or relation_name not in wb.sheetnames or META_SHEET not in wb.sheetnames:
        raise ValueError("Workbook lacks the draft, relation, or metadata sheet.")
    draft = wb[draft_name]
    relation = wb[relation_name]
    draft_headers, _, draft_records = records_from_sheet(draft)
    expected_headers = [item["name"] for item in config["run_schema"]]
    if draft_headers != expected_headers:
        raise ValueError("Draft headers do not equal the approved Run Schema.")
    relation_map = relation_header_map(relation)
    relation_last_row, _ = used_bounds(relation)
    groups: dict[str, list[int]] = defaultdict(list)
    for row in range(2, relation_last_row + 1):
        groups[stable_key(relation.cell(row, relation_map["临床分子序号"]).value)].append(row)
    meta_map, _ = read_meta(wb[META_SHEET])
    fields = config["final_fields"]
    sequence_field = fields["sequence"]
    listed_source = config["output"]["listed_source_value"]
    clinical_source = config["output"]["clinical_source_value"]
    listed_records = [record for record in draft_records if record_source(record, fields, meta_map) == listed_source]
    clinical_records = [record for record in draft_records if record_source(record, fields, meta_map) == clinical_source]
    listed_sequences = {stable_key(record[sequence_field]) for record in listed_records}
    clinical_sequences = {stable_key(record[sequence_field]) for record in clinical_records}
    errors: list[str] = []
    missing_groups = sorted(clinical_sequences - set(groups))
    extra_groups = sorted(set(groups) - clinical_sequences)
    if missing_groups:
        errors.append(f"Clinical rows missing relation groups: {missing_groups}")
    if extra_groups:
        errors.append(f"Relation groups lack clinical draft rows: {extra_groups}")
    decisions: dict[str, dict[str, str]] = {}
    action_counts: Counter[str] = Counter()
    open_sequences: list[str] = []
    for sequence, rows in groups.items():
        values = {}
        for field in ("最终动作代码", "最终承接上市分子序号", "判断状态", "Disease修正值", "资产名称修正值"):
            unique = unique_field_values(relation, rows, relation_map[field])
            if len(unique) > 1:
                errors.append(f"Conflicting {field} values for clinical sequence {sequence}: {unique}")
            values[field] = unique[0] if unique else ""
        action = values["最终动作代码"]
        status = values["判断状态"]
        carrier = values["最终承接上市分子序号"]
        if action not in ALLOWED_ACTIONS:
            errors.append(f"Missing or unsupported action for clinical sequence {sequence}: {action}")
        if action == "HOLD" or not status or status.startswith("待"):
            open_sequences.append(sequence)
        if action == "DELETE_LISTED_COVERED":
            carriers = parse_carrier_sequences(carrier)
            invalid_carriers = [item for item in carriers if item not in listed_sequences]
            if not carriers or invalid_carriers:
                errors.append(f"Invalid listed carrier(s) for clinical sequence {sequence}: {carrier}")
        if action and action != "HOLD":
            action_counts[action] += 1
        decisions[sequence] = {
            "action": action,
            "carrier": carrier,
            "status": status,
            "disease_correction": values["Disease修正值"],
            "asset_correction": values["资产名称修正值"],
        }
    if open_sequences:
        errors.append(f"Open/HOLD clinical sequences remain: {sorted(open_sequences)}")
    keep = action_counts["KEEP"]
    expected_final = len(listed_records) + keep
    preview = {
        "state": "DECISIONS_CLOSED" if not errors else "DECISIONS_PENDING",
        "listed_rows": len(listed_records),
        "clinical_rows": len(clinical_records),
        "unique_relation_groups": len(groups),
        "candidate_rows": relation_last_row - 1,
        "action_counts": dict(action_counts),
        "expected_final_rows": expected_final,
        "expected_final_source_counts": {listed_source: len(listed_records), clinical_source: keep},
        "pending_or_errors": errors,
    }
    context = {
        "wb": wb,
        "draft": draft,
        "relation": relation,
        "draft_headers": draft_headers,
        "draft_records": draft_records,
        "relation_map": relation_map,
        "groups": groups,
        "decisions": decisions,
        "listed_sequences": listed_sequences,
        "listed_records": listed_records,
        "clinical_records": clinical_records,
        "meta_map": meta_map,
    }
    return preview, context


def copy_row_style(source_ws, source_row: int, target_ws, target_row: int, columns: int) -> None:
    for col in range(1, columns + 1):
        source = source_ws.cell(source_row, col)
        target = target_ws.cell(target_row, col)
        if source.has_style:
            target.font = copy(source.font)
            target.fill = copy(source.fill)
            target.border = copy(source.border)
            target.alignment = copy(source.alignment)
            target.protection = copy(source.protection)
            target.number_format = source.number_format


def count_formulas(ws) -> int:
    last_row, last_col = used_bounds(ws)
    return sum(
        1
        for row in range(1, last_row + 1)
        for col in range(1, last_col + 1)
        if isinstance(ws.cell(row, col).value, str) and ws.cell(row, col).value.startswith("=")
    )


def command_finalize(args) -> None:
    config = load_json(args.config)
    validate_config_shape(config)
    if not config.get("schema_approved"):
        raise ValueError("Finalization requires an approved Run Schema.")
    workbook = Path(args.workbook)
    if not workbook.is_file():
        raise ValueError(f"Decided workbook was not found: {workbook}")
    preview, context = finalization_preview(config, workbook)
    if args.mode == "preview":
        emit(preview)
        return
    if not args.confirmed:
        raise ValueError("Generate mode requires --confirmed after preview and human approval.")
    if preview["pending_or_errors"]:
        raise ValueError("Finalization preview failed: " + " | ".join(preview["pending_or_errors"]))
    if not args.output:
        raise ValueError("Generate mode requires --output.")
    output = require_new_output(args.output, [workbook])
    context["wb"].close()
    shutil.copy2(workbook, output)
    wb = load_workbook(output, data_only=False)
    draft_name = config["output"]["draft_sheet"]
    relation_name = config["output"]["relation_sheet"]
    final_name = config["output"]["final_sheet"]
    if final_name in wb.sheetnames:
        output.unlink(missing_ok=True)
        raise ValueError(f"Final sheet already exists: {final_name}")
    draft = wb[draft_name]
    relation = wb[relation_name]
    _, _, draft_records = records_from_sheet(draft)
    relation_map = relation_header_map(relation)
    relation_last_row, _ = used_bounds(relation)
    groups: dict[str, list[int]] = defaultdict(list)
    for row in range(2, relation_last_row + 1):
        groups[stable_key(relation.cell(row, relation_map["临床分子序号"]).value)].append(row)
    decisions: dict[str, dict[str, str]] = {}
    for sequence, rows in groups.items():
        decisions[sequence] = {
            "action": unique_field_values(relation, rows, relation_map["最终动作代码"])[0],
            "carrier": (unique_field_values(relation, rows, relation_map["最终承接上市分子序号"]) or [""])[0],
            "disease_correction": (unique_field_values(relation, rows, relation_map["Disease修正值"]) or [""])[0],
            "asset_correction": (unique_field_values(relation, rows, relation_map["资产名称修正值"]) or [""])[0],
        }
    fields = config["final_fields"]
    sequence_field = fields["sequence"]
    listed_source = config["output"]["listed_source_value"]
    clinical_source = config["output"]["clinical_source_value"]
    meta_map, _ = read_meta(wb[META_SHEET])
    headers = [item["name"] for item in config["run_schema"]]
    final_rows: list[tuple[dict[str, Any], int, str, str]] = []
    clinical_to_final: dict[str, int] = {}
    for record in draft_records:
        source_value = record_source(record, fields, meta_map)
        old_sequence = stable_key(record[sequence_field])
        if source_value == listed_source:
            final_rows.append((record, record["__row_number__"], old_sequence, source_value))
        elif source_value == clinical_source and decisions[old_sequence]["action"] == "KEEP":
            corrected = dict(record)
            if decisions[old_sequence]["disease_correction"]:
                corrected[fields["disease"]] = decisions[old_sequence]["disease_correction"]
            if decisions[old_sequence]["asset_correction"]:
                corrected[fields["asset"]] = decisions[old_sequence]["asset_correction"]
            final_rows.append((corrected, record["__row_number__"], old_sequence, source_value))
    final = wb.create_sheet(final_name, 0)
    for col, header in enumerate(headers, start=1):
        final.cell(1, col).value = header
        source_cell = draft.cell(1, col)
        if source_cell.has_style:
            final.cell(1, col)._style = copy(source_cell._style)
        final.column_dimensions[get_column_letter(col)].width = draft.column_dimensions[get_column_letter(col)].width
    for new_sequence, (record, source_row, old_sequence, source_value) in enumerate(final_rows, start=1):
        record[sequence_field] = new_sequence
        output_row = new_sequence + 1
        for col, header in enumerate(headers, start=1):
            final.cell(output_row, col).value = record.get(header)
        copy_row_style(draft, source_row, final, output_row, len(headers))
        if source_value == clinical_source:
            clinical_to_final[old_sequence] = new_sequence
    final.freeze_panes = "A2"
    final.auto_filter.ref = final.dimensions
    for row in range(2, relation_last_row + 1):
        sequence = stable_key(relation.cell(row, relation_map["临床分子序号"]).value)
        action = decisions[sequence]["action"]
        if action == "KEEP":
            result = "已保留"
            final_sequence = clinical_to_final[sequence]
            sync = "已同步并保留" if decisions[sequence]["disease_correction"] or decisions[sequence]["asset_correction"] else "阶段三已执行：保留"
        elif action == "DELETE_LISTED_COVERED":
            result = "已删除：上市候选承接"
            final_sequence = ""
            sync = "阶段三已执行：删除临床行"
        else:
            result = "已排除"
            final_sequence = ""
            sync = "阶段三已执行：排除"
        relation.cell(row, relation_map["阶段执行结果"]).value = result
        relation.cell(row, relation_map["最终分子序号"]).value = final_sequence
        relation.cell(row, relation_map["主表同步状态"]).value = sync

    _, meta_state = read_meta(wb[META_SHEET])
    expected_listed_fp = meta_state.get("listed_raw_fingerprint", "")
    expected_clinical_fp = meta_state.get("clinical_raw_fingerprint", "")
    wb.remove(wb[draft_name])
    wb.remove(wb[META_SHEET])
    desired = [
        final_name,
        relation_name,
        config["listed"]["raw_output_sheet"],
        config["clinical"]["raw_output_sheet"],
    ]
    if any(name not in wb.sheetnames for name in desired):
        output.unlink(missing_ok=True)
        raise ValueError("Final workbook lacks a required output sheet.")
    extras = [name for name in wb.sheetnames if name not in desired]
    if extras:
        output.unlink(missing_ok=True)
        raise ValueError(f"Unexpected sheets would violate the four-sheet output contract: {extras}")
    wb._sheets = [wb[name] for name in desired]
    wb.save(output)

    verify = load_workbook(output, data_only=False)
    final_ws = verify[final_name]
    final_headers, _, final_records = records_from_sheet(final_ws)
    final_sequences = [record[sequence_field] for record in final_records]
    source_counts = Counter(source_value for _, _, _, source_value in final_rows)
    listed_fp = sheet_fingerprint(verify[config["listed"]["raw_output_sheet"]])
    clinical_fp = sheet_fingerprint(verify[config["clinical"]["raw_output_sheet"]])
    checks = {
        "sheet_order": verify.sheetnames == desired,
        "run_schema": final_headers == headers,
        "final_rows": len(final_records) == preview["expected_final_rows"],
        "sequence": final_sequences == list(range(1, len(final_records) + 1)),
        "source_counts": dict(source_counts) == preview["expected_final_source_counts"],
        "listed_raw_fingerprint": listed_fp == expected_listed_fp,
        "clinical_raw_fingerprint": clinical_fp == expected_clinical_fp,
        "final_formula_cells": count_formulas(final_ws) == 0,
        "no_trailing_blank_rows": final_ws.max_row == len(final_records) + 1,
    }
    if not all(checks.values()):
        output.unlink(missing_ok=True)
        raise ValueError("Final workbook QC failed: " + json.dumps(checks, ensure_ascii=False))
    emit(
        {
            "state": "QC_PASSED",
            "output": str(output),
            "listed_rows": preview["listed_rows"],
            "clinical_rows": preview["clinical_rows"],
            "action_counts": preview["action_counts"],
            "final_rows": len(final_records),
            "final_source_counts": dict(source_counts),
            "final_sequence": f"1-{len(final_records)}",
            "pending": 0,
            "checks": checks,
        }
    )


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(description="Merge listed and clinical molecule pools through gated workbook stages.")
    subparsers = parser.add_subparsers(dest="command", required=True)

    inspect_parser = subparsers.add_parser("inspect", help="Read-only workbook and schema preflight.")
    inspect_parser.add_argument("--config", required=True)
    inspect_parser.set_defaults(func=command_inspect)

    draft_parser = subparsers.add_parser("build-draft", help="Build an approved-schema combined draft.")
    draft_parser.add_argument("--config", required=True)
    draft_parser.add_argument("--output", required=True)
    draft_parser.add_argument("--confirmed-schema", action="store_true")
    draft_parser.set_defaults(func=command_build_draft)

    candidate_parser = subparsers.add_parser("build-candidates", help="Generate high-recall cross-library candidates.")
    candidate_parser.add_argument("--config", required=True)
    candidate_parser.add_argument("--workbook", required=True)
    candidate_parser.add_argument("--output", required=True)
    candidate_parser.set_defaults(func=command_build_candidates)

    decision_parser = subparsers.add_parser("import-decisions", help="Import structured decisions into every candidate row in a clinical group.")
    decision_parser.add_argument("--workbook", required=True)
    decision_parser.add_argument("--decisions", required=True)
    decision_parser.add_argument("--output", required=True)
    decision_parser.set_defaults(func=command_import_decisions)

    final_parser = subparsers.add_parser("finalize", help="Preview or generate the final pool from closed decisions.")
    final_parser.add_argument("--mode", choices=("preview", "generate"), required=True)
    final_parser.add_argument("--config", required=True)
    final_parser.add_argument("--workbook", required=True)
    final_parser.add_argument("--output")
    final_parser.add_argument("--confirmed", action="store_true")
    final_parser.set_defaults(func=command_finalize)
    return parser


def main() -> int:
    parser = build_parser()
    args = parser.parse_args()
    try:
        args.func(args)
        return 0
    except Exception as exc:  # noqa: BLE001 - CLI boundary
        print(json.dumps({"state": "ERROR", "error": str(exc)}, ensure_ascii=False, indent=2), file=sys.stderr)
        return 1


if __name__ == "__main__":
    raise SystemExit(main())
