from __future__ import annotations

import hashlib
import html
import json
import re
import zipfile
from collections import defaultdict
from dataclasses import dataclass
from datetime import UTC, date, datetime, timedelta
from pathlib import Path
from typing import Any
from xml.etree import ElementTree as ET


MAIN_NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
REL_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
PKG_REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships"
NS = {"main": MAIN_NS, "rel": REL_NS, "pkg": PKG_REL_NS}


@dataclass(frozen=True)
class WarnRow:
    county: str
    county_normalized: str
    notice_date: date | None
    processed_date: date | None
    effective_date: date | None
    company: str
    layoff_closure: str
    employees: int
    address: str
    industry: str


def normalize_whitespace(value: str) -> str:
    return re.sub(r"\s+", " ", value or "").strip()


def normalize_text(value: str) -> str:
    text = html.unescape(value or "")
    text = text.replace("\\", " ")
    return normalize_whitespace(text)


def normalize_county(value: str) -> str:
    county = normalize_text(value)
    county = re.sub(r"\s+County$", "", county, flags=re.IGNORECASE)
    county = re.sub(r"\s+Parish$", "", county, flags=re.IGNORECASE)
    return county


def format_date(value: date | None) -> str:
    if value is None:
        return "—"
    return value.strftime("%b %-d, %Y")


def format_number(value: int) -> str:
    return f"{int(value):,}"


def excel_serial_to_date(serial: float) -> date:
    return (datetime(1899, 12, 30) + timedelta(days=float(serial))).date()


def parse_excel_date(value: Any) -> date | None:
    if value is None:
        return None

    if isinstance(value, date) and not isinstance(value, datetime):
        return value

    text = normalize_text(str(value))
    if not text:
        return None

    if re.fullmatch(r"-?\d+(?:\.\d+)?", text):
        try:
            return excel_serial_to_date(float(text))
        except ValueError:
            return None

    for fmt in ("%m/%d/%Y", "%m/%d/%y", "%Y-%m-%d", "%b %d, %Y", "%B %d, %Y"):
        try:
            return datetime.strptime(text, fmt).date()
        except ValueError:
            continue

    return None


def parse_shared_strings(zf: zipfile.ZipFile) -> list[str]:
    try:
        root = ET.fromstring(zf.read("xl/sharedStrings.xml"))
    except KeyError:
        return []

    strings: list[str] = []
    for item in root.findall("main:si", NS):
        text_parts = [node.text or "" for node in item.findall(".//main:t", NS)]
        strings.append(normalize_text("".join(text_parts)))
    return strings


def parse_workbook_sheet_target(zf: zipfile.ZipFile) -> str:
    workbook_root = ET.fromstring(zf.read("xl/workbook.xml"))
    rels_root = ET.fromstring(zf.read("xl/_rels/workbook.xml.rels"))
    rel_map = {
        rel.attrib["Id"]: rel.attrib["Target"]
        for rel in rels_root.findall("pkg:Relationship", NS)
    }

    sheets = workbook_root.findall("main:sheets/main:sheet", NS)
    normalized = []
    for sheet in sheets:
      name = normalize_text(sheet.attrib.get("name", ""))
      normalized.append((name, sheet.attrib.get(f"{{{REL_NS}}}id", "")))

    target_rel = ""
    for name, rel_id in normalized:
        lowered = name.lower()
        if "detailed" in lowered and "warn" in lowered:
            target_rel = rel_id
            break

    if not target_rel:
        for name, rel_id in normalized:
            if "warn" in name.lower():
                target_rel = rel_id
                break

    if not target_rel and normalized:
        target_rel = normalized[0][1]

    if not target_rel or target_rel not in rel_map:
        raise ValueError("Unable to locate a detailed WARN worksheet in the workbook.")

    target = rel_map[target_rel]
    target = target.lstrip("/")
    if not target.startswith("xl/"):
        target = f"xl/{target}"
    return target


def column_index_from_ref(cell_ref: str) -> int:
    letters = re.sub(r"\d+", "", cell_ref).upper()
    index = 0
    for char in letters:
        index = (index * 26) + (ord(char) - 64)
    return max(index - 1, 0)


def cell_value(cell: ET.Element, shared_strings: list[str]) -> str:
    cell_type = cell.attrib.get("t")
    value_node = cell.find("main:v", NS)
    inline_node = cell.find("main:is", NS)

    if cell_type == "s" and value_node is not None and value_node.text is not None:
        try:
            return shared_strings[int(value_node.text)]
        except (IndexError, ValueError):
            return ""

    if cell_type == "inlineStr" and inline_node is not None:
        return normalize_text("".join(node.text or "" for node in inline_node.findall(".//main:t", NS)))

    if value_node is None or value_node.text is None:
        return ""

    return normalize_text(value_node.text)


def parse_sheet_rows(zf: zipfile.ZipFile, sheet_target: str, shared_strings: list[str]) -> list[list[str]]:
    sheet_root = ET.fromstring(zf.read(sheet_target))
    rows: list[list[str]] = []

    for row in sheet_root.findall("main:sheetData/main:row", NS):
        values: dict[int, str] = {}
        max_index = -1
        for cell in row.findall("main:c", NS):
            ref = cell.attrib.get("r", "")
            col_index = column_index_from_ref(ref)
            values[col_index] = cell_value(cell, shared_strings)
            max_index = max(max_index, col_index)
        if max_index < 0:
            rows.append([])
            continue
        rows.append([values.get(index, "") for index in range(max_index + 1)])

    return rows


def load_warn_rows(workbook_path: Path) -> list[WarnRow]:
    with zipfile.ZipFile(workbook_path) as zf:
        shared_strings = parse_shared_strings(zf)
        sheet_target = parse_workbook_sheet_target(zf)
        matrix = parse_sheet_rows(zf, sheet_target, shared_strings)

    rows: list[WarnRow] = []
    for raw_row in matrix[2:]:
        county = normalize_text(raw_row[0] if len(raw_row) > 0 else "")
        company = normalize_text(raw_row[4] if len(raw_row) > 4 else "")
        employees_text = normalize_text(raw_row[6] if len(raw_row) > 6 else "")
        employees = int(float(re.sub(r"[^\d.-]", "", employees_text) or 0))

        if not county or not company or employees <= 0:
            continue

        rows.append(
            WarnRow(
                county=county,
                county_normalized=normalize_county(county),
                notice_date=parse_excel_date(raw_row[1] if len(raw_row) > 1 else ""),
                processed_date=parse_excel_date(raw_row[2] if len(raw_row) > 2 else ""),
                effective_date=parse_excel_date(raw_row[3] if len(raw_row) > 3 else ""),
                company=company,
                layoff_closure=normalize_text(raw_row[5] if len(raw_row) > 5 else ""),
                employees=employees,
                address=normalize_text(raw_row[7] if len(raw_row) > 7 else ""),
                industry=normalize_text(raw_row[8] if len(raw_row) > 8 else ""),
            )
        )

    return rows


def compute_date_range(rows: list[WarnRow]) -> tuple[date | None, date | None]:
    dates = sorted(row.effective_date for row in rows if row.effective_date is not None)
    if not dates:
        return None, None
    return dates[0], dates[-1]


def existing_archive_index(archive_root: Path) -> dict[str, Any]:
    archive_root.mkdir(parents=True, exist_ok=True)
    fingerprints: set[str] = set()
    archive_keys: list[str] = []

    for post_path in sorted(archive_root.glob("*/post.json")):
        archive_keys.append(post_path.parent.name)
        try:
            payload = json.loads(post_path.read_text(encoding="utf-8"))
        except json.JSONDecodeError:
            continue
        fingerprint = str(payload.get("fingerprint", "")).strip()
        if fingerprint:
            fingerprints.add(fingerprint)

    return {
        "archive_keys": archive_keys,
        "fingerprints": fingerprints,
    }


def fingerprint_for(parts: list[str]) -> str:
    digest = hashlib.sha1("||".join(parts).encode("utf-8")).hexdigest()
    return digest


def event_sort_key(row: WarnRow) -> tuple[Any, ...]:
    return (
        row.effective_date or date.max,
        -row.employees,
        row.company.lower(),
        row.county_normalized.lower(),
        row.address.lower(),
    )


def build_next_event_candidates(rows: list[WarnRow], publish_date: date) -> list[dict[str, Any]]:
    upcoming_rows = [
        row for row in rows if row.effective_date is not None and row.effective_date >= publish_date
    ]
    candidates: list[dict[str, Any]] = []

    for row in sorted(upcoming_rows, key=event_sort_key):
        effective_date = row.effective_date
        assert effective_date is not None

        linkedin = (
            f"The next scheduled layoff in California's current WARN file is at {row.company} "
            f"in {row.county}, effective {format_date(effective_date)}."
            f"\n\nThe notice lists {format_number(row.employees)} affected workers."
        )

        if row.address:
            linkedin += f"\n\nLocation: {row.address}."

        if row.industry:
            linkedin += f"\n\nIndustry: {row.industry}."

        fingerprint = fingerprint_for(
            [
                "next-event",
                row.company,
                row.county_normalized,
                row.address,
                effective_date.isoformat(),
                str(row.employees),
            ]
        )

        candidates.append(
            {
                "fingerprint": fingerprint,
                "linkedin": linkedin,
                "post_type": "next_event",
                "event": {
                    "company": row.company,
                    "county": row.county,
                    "county_normalized": row.county_normalized,
                    "address": row.address,
                    "effective_date": effective_date.isoformat(),
                    "employees": row.employees,
                    "industry": row.industry,
                    "layoff_closure": row.layoff_closure,
                },
                "priority": "upcoming",
            }
        )

    return candidates


def build_summary_candidates(rows: list[WarnRow], publish_date: date) -> list[dict[str, Any]]:
    range_start, range_end = compute_date_range(rows)
    total_layoffs = sum(row.employees for row in rows)
    county_count = len({row.county_normalized for row in rows})

    upcoming_rows = [
        row for row in rows if row.effective_date is not None and row.effective_date >= publish_date
    ]
    next_candidates: list[dict[str, Any]] = []

    if upcoming_rows:
        first_date = min(row.effective_date for row in upcoming_rows if row.effective_date is not None)
        same_day_rows = [row for row in upcoming_rows if row.effective_date == first_date]
        impacted_workers = sum(row.employees for row in same_day_rows)
        impacted_counties = len({row.county_normalized for row in same_day_rows})
        impacted_locations = len({(row.company, row.address, row.county_normalized) for row in same_day_rows})

        next_day_copy = (
            f"The next wave of California WARN notices takes effect on {format_date(first_date)}."
            f"\n\nThe current file shows {format_number(impacted_workers)} affected workers across "
            f"{format_number(impacted_locations)} location{'' if impacted_locations == 1 else 's'} in "
            f"{format_number(impacted_counties)} count{'y' if impacted_counties == 1 else 'ies'} on that date."
        )

        next_candidates.append(
            {
                "fingerprint": fingerprint_for(
                    [
                        "next-day-summary",
                        first_date.isoformat(),
                        str(impacted_workers),
                        str(impacted_locations),
                        str(impacted_counties),
                    ]
                ),
                "linkedin": next_day_copy,
                "post_type": "next_day_summary",
                "summary": {
                    "effective_date": first_date.isoformat(),
                    "workers": impacted_workers,
                    "locations": impacted_locations,
                    "counties": impacted_counties,
                },
                "priority": "upcoming_summary",
            }
        )

    range_label = f"{format_date(range_start)} to {format_date(range_end)}"
    total_copy = (
        f"The current California WARN file spans {range_label}."
        f"\n\nAcross that period, it lists {format_number(total_layoffs)} workers affected by layoffs statewide."
    )
    counties_copy = (
        f"The current California WARN file spans {range_label}."
        f"\n\nIt includes layoff notices across {format_number(county_count)} California "
        f"count{'y' if county_count == 1 else 'ies'}."
    )

    return next_candidates + [
        {
            "fingerprint": fingerprint_for(
                [
                    "range-total",
                    range_start.isoformat() if range_start else "",
                    range_end.isoformat() if range_end else "",
                    str(total_layoffs),
                ]
            ),
            "linkedin": total_copy,
            "post_type": "total_layoffs",
            "summary": {
                "range_start": range_start.isoformat() if range_start else None,
                "range_end": range_end.isoformat() if range_end else None,
                "total_layoffs": total_layoffs,
            },
            "priority": "range_total",
        },
        {
            "fingerprint": fingerprint_for(
                [
                    "county-impact",
                    range_start.isoformat() if range_start else "",
                    range_end.isoformat() if range_end else "",
                    str(county_count),
                ]
            ),
            "linkedin": counties_copy,
            "post_type": "county_impact",
            "summary": {
                "range_start": range_start.isoformat() if range_start else None,
                "range_end": range_end.isoformat() if range_end else None,
                "county_count": county_count,
            },
            "priority": "county_impact",
        },
    ]


def select_post(rows: list[WarnRow], publish_date: date, existing_fingerprints: set[str]) -> dict[str, Any] | None:
    candidates = build_next_event_candidates(rows, publish_date) + build_summary_candidates(rows, publish_date)
    for candidate in candidates:
        if candidate["fingerprint"] not in existing_fingerprints:
            return candidate
    return None


def archive_key_for_date(publish_date: date) -> str:
    return publish_date.isoformat()


def archive_dir_for_date(archive_root: Path, publish_date: date) -> Path:
    return archive_root / archive_key_for_date(publish_date)


def read_post_archive(archive_root: Path, publish_date: date) -> tuple[Path, dict[str, Any]] | None:
    archive_dir = archive_dir_for_date(archive_root, publish_date)
    post_path = archive_dir / "post.json"
    if not post_path.exists():
        return None
    return archive_dir, json.loads(post_path.read_text(encoding="utf-8"))


def write_post_archive(
    archive_root: Path,
    publish_date: date,
    workbook_path: Path,
    rows: list[WarnRow],
    candidate: dict[str, Any],
) -> tuple[str, Path]:
    archive_key = archive_key_for_date(publish_date)
    output_dir = archive_root / archive_key
    output_dir.mkdir(parents=True, exist_ok=True)

    range_start, range_end = compute_date_range(rows)
    payload = {
        "archive_key": archive_key,
        "planned_date": publish_date.isoformat(),
        "generated_at": datetime.now(UTC).replace(microsecond=0).isoformat().replace("+00:00", "Z"),
        "source_workbook": str(workbook_path),
        "post_type": candidate["post_type"],
        "fingerprint": candidate["fingerprint"],
        "stats": {
            "total_layoffs": sum(row.employees for row in rows),
            "county_count": len({row.county_normalized for row in rows}),
            "row_count": len(rows),
            "range_start": range_start.isoformat() if range_start else None,
            "range_end": range_end.isoformat() if range_end else None,
        },
        "social": {
            "linkedin": candidate["linkedin"],
        },
    }

    if "event" in candidate:
        payload["event"] = candidate["event"]
    if "summary" in candidate:
        payload["summary"] = candidate["summary"]

    (output_dir / "post.json").write_text(json.dumps(payload, indent=2) + "\n", encoding="utf-8")

    post_copy = [
        f"Archive key: {archive_key}",
        f"Planned date: {publish_date.isoformat()}",
        f"Post type: {candidate['post_type']}",
        f"Fingerprint: {candidate['fingerprint']}",
        "",
        "LinkedIn",
        candidate["linkedin"],
        "",
    ]
    (output_dir / "post-copy.txt").write_text("\n".join(post_copy), encoding="utf-8")

    return archive_key, output_dir


def summarize_archives_by_date(archive_root: Path) -> dict[str, list[str]]:
    grouped: dict[str, list[str]] = defaultdict(list)
    for post_path in sorted(archive_root.glob("*/post.json")):
        archive_key = post_path.parent.name
        grouped[archive_key[:10]].append(archive_key)
    return dict(grouped)
