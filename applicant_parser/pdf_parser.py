from __future__ import annotations

import re
from contextlib import redirect_stderr, redirect_stdout
from io import StringIO
import warnings
from pathlib import Path

try:
    import fitz
except ImportError as exc:  # pragma: no cover - user-facing guard
    raise SystemExit(
        "Не установлен PyMuPDF. Выполните: python -m pip install -r requirements.txt"
    ) from exc

from .config import PRIORITY_SUFFIX, QUOTAS, GroupInfo
from .models import ApplicantRecord, ApplicantSummary, ParseResult

GROUP_RE = re.compile(r"Конкурсная группа\s*-\s*(.+)")
FIELD_PATTERNS = {
    "study_form": re.compile(r"Форма обучения\s*-\s*(.+)"),
    "level": re.compile(r"Уровень подготовки\s*-\s*(.+)"),
    "direction": re.compile(r"УГС/Направление подготовки/специальность\s*-\s*(.+)"),
}


def parse_deans_list_pdf(pdf_path: str | Path) -> ParseResult:
    path = Path(pdf_path).expanduser().resolve()
    if not path.exists():
        raise FileNotFoundError(f"Файл не найден: {path}")
    if path.suffix.lower() != ".pdf":
        raise ValueError("Ожидается PDF-файл.")

    warnings.filterwarnings("ignore", message="Consider using the pymupdf_layout")

    records: list[ApplicantRecord] = []
    groups: set[str] = set()
    current_group: GroupInfo | None = None

    with fitz.open(path) as doc:
        for page_index, page in enumerate(doc, start=1):
            text = page.get_text()
            if "Конкурсная группа" in text:
                current_group = _extract_group_info(text)

            if current_group is None:
                continue

            groups.add(current_group.raw_name)
            for row in _extract_table_rows(page):
                record = _row_to_record(row, current_group, page_index)
                if record is not None:
                    records.append(record)

        applicants = _aggregate_applicants(records)
        return ParseResult(
            source_path=str(path),
            faculty="Все",
            records=records,
            applicants=applicants,
            groups=groups,
            pages_count=doc.page_count,
        )


def split_result_by_faculty(parse_result: ParseResult) -> list[ParseResult]:
    faculty_order: list[str] = []
    records_by_faculty: dict[str, list[ApplicantRecord]] = {}

    for record in parse_result.records:
        if record.faculty not in records_by_faculty:
            records_by_faculty[record.faculty] = []
            faculty_order.append(record.faculty)
        records_by_faculty[record.faculty].append(record)

    results: list[ParseResult] = []
    for faculty in faculty_order:
        records = records_by_faculty[faculty]
        results.append(
            ParseResult(
                source_path=parse_result.source_path,
                faculty=faculty,
                records=records,
                applicants=_aggregate_applicants(records),
                groups={record.group_name for record in records},
                pages_count=parse_result.pages_count,
            )
        )
    return results


def _extract_group_info(page_text: str) -> GroupInfo | None:
    group_match = GROUP_RE.search(page_text)
    if not group_match:
        return None

    raw_name = group_match.group(1).strip()
    faculty = _extract_faculty(raw_name)
    if not faculty:
        return None

    fields = {
        name: (pattern.search(page_text).group(1).strip() if pattern.search(page_text) else "")
        for name, pattern in FIELD_PATTERNS.items()
    }
    direction = fields["direction"]
    if not direction:
        return None

    quota_raw = _extract_quota_raw(raw_name, faculty, direction)
    quota = _normalize_quota(quota_raw)

    return GroupInfo(
        faculty=faculty,
        raw_name=raw_name,
        direction=direction,
        quota=quota,
        quota_raw=quota_raw,
        study_form=fields["study_form"],
        level=fields["level"],
    )


def _extract_faculty(group_name: str) -> str:
    return group_name.split("_", 1)[0].strip()


def _extract_quota_raw(group_name: str, faculty: str, direction: str) -> str:
    prefix = f"{faculty}_{direction}_"
    if group_name.startswith(prefix):
        return group_name[len(prefix) :].strip()

    # Fallback for rare reports where separators are normalized by the PDF layer.
    without_faculty = group_name.removeprefix(f"{faculty}_")
    return without_faculty.replace(direction, "", 1).strip("_ ").strip()


def _normalize_quota(quota_raw: str) -> str:
    for quota in QUOTAS:
        if quota != "Другое" and quota in quota_raw:
            return quota
    return "Другое"


def _extract_table_rows(page) -> list[list[str]]:
    table_rows: list[list[str]] = []
    try:
        with redirect_stdout(StringIO()), redirect_stderr(StringIO()):
            tables = page.find_tables().tables
    except Exception:
        return table_rows

    for table in tables:
        for row in table.extract():
            normalized = [(cell or "").strip() for cell in row]
            if len(normalized) >= 13 and normalized[0].isdigit():
                table_rows.append(normalized)
    return table_rows


def _row_to_record(row: list[str], group: GroupInfo, page_number: int) -> ApplicantRecord | None:
    fio = _clean_human_text(row[1])
    if not fio:
        return None

    return ApplicantRecord(
        faculty=group.faculty,
        fio=fio,
        email=_clean_compact(row[6]),
        phone=_clean_phone(row[7]),
        total_score=_parse_int(row[2]),
        subject_score=_parse_int(row[3]),
        achievement_score=_parse_int(row[4]),
        consent=bool(_clean_human_text(row[5])),
        direction=group.direction,
        quota=group.quota,
        quota_raw=group.quota_raw,
        priority=_clean_human_text(row[11]),
        group_name=group.raw_name,
        study_form=group.study_form,
        level=group.level,
        page_number=page_number,
    )


def _aggregate_applicants(records: list[ApplicantRecord]) -> list[ApplicantSummary]:
    by_email: dict[str, ApplicantSummary] = {}

    for record in records:
        key = _applicant_key(record)
        applicant = by_email.setdefault(key, ApplicantSummary(fio=record.fio))

        if record.total_score > applicant.total_score:
            applicant.fio = record.fio
            applicant.total_score = record.total_score
            applicant.subject_score = record.subject_score
            applicant.achievement_score = record.achievement_score

        applicant.email = applicant.email or record.email
        applicant.phone = _merge_multiline_values(applicant.phone, record.phone)
        applicant.consent = applicant.consent or record.consent
        applicant.source_groups.add(record.group_name)

        priority_value = _format_priority(record)
        cell_key = (record.direction, record.quota)
        if priority_value:
            applicant.priorities[cell_key] = _merge_cell_values(
                applicant.priorities.get(cell_key, ""),
                priority_value,
            )

    return sorted(by_email.values(), key=lambda item: (-item.total_score, item.fio.lower()))


def _format_priority(record: ApplicantRecord) -> str:
    if not record.priority:
        return ""
    suffix = PRIORITY_SUFFIX.get(record.quota, "?")
    value = record.priority if not suffix else f"{record.priority}{suffix}"

    if record.quota == "Целевая":
        target_name = record.quota_raw.removeprefix("Целевая").strip("_ ")
        if target_name:
            value = f"{value} ({target_name})"
    return value


def _merge_cell_values(existing: str, new_value: str) -> str:
    if not existing:
        return new_value
    values = [part.strip() for part in existing.split(";")]
    if new_value not in values:
        return f"{existing}; {new_value}"
    return existing


def _merge_multiline_values(existing: str, new_value: str) -> str:
    values: list[str] = []
    for source in (existing, new_value):
        for value in source.splitlines():
            value = value.strip()
            if value and value not in values:
                values.append(value)
    return "\n".join(values)


def _clean_human_text(value: str) -> str:
    text = value.replace("\u00a0", " ")
    lines = [line.strip() for line in text.splitlines() if line.strip()]
    return re.sub(r"\s+", " ", " ".join(lines)).strip()


def _clean_compact(value: str) -> str:
    return re.sub(r"\s+", "", value.replace("\u00a0", ""))


def _clean_phone(value: str) -> str:
    text = _clean_human_text(value)
    if not text:
        return ""

    phones: list[str] = []
    for part in re.split(r"[;,\n]+", text):
        formatted_phone = _format_phone(part)
        if formatted_phone and formatted_phone not in phones:
            phones.append(formatted_phone)
    return "\n".join(phones)


def _format_phone(value: str) -> str:
    digits = re.sub(r"\D+", "", value)
    if not digits:
        return ""

    if len(digits) == 10:
        digits = f"7{digits}"
    elif len(digits) == 11 and digits.startswith("8"):
        digits = f"7{digits[1:]}"

    if len(digits) == 11 and digits.startswith("7"):
        return f"+7 ({digits[1:4]}) {digits[4:7]}-{digits[7:9]}-{digits[9:11]}"

    return f"+{digits}"


def _parse_int(value: str) -> int:
    match = re.search(r"\d+", value or "")
    return int(match.group(0)) if match else 0


def _applicant_key(record: ApplicantRecord) -> str:
    if record.email:
        return f"email:{record.email.casefold()}"
    return f"fio:{re.sub(r'\s+', ' ', record.fio).strip().casefold()}"
