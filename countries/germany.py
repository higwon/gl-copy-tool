import re
from datetime import datetime

from .common import is_empty_value, normalize_amount, record, report_progress


ACCOUNT_TITLE_RE = re.compile(r"Jahreskonto\s+(\S+)(?:\s+(.*))?$")
DATE_FORMAT = "%d.%m.%Y"


def build_records(sheet, progress_callback=None):
    records = []
    current_account = None
    total_rows = max(sheet.max_row, 1)

    for row_number, row in enumerate(sheet.iter_rows(values_only=True), start=1):
        first_value = row[0] if row else None
        if isinstance(first_value, str) and "Jahreskonto" in first_value:
            match = ACCOUNT_TITLE_RE.search(first_value)
            current_account = match.group(1) if match else None
            continue

        if first_value == "Datum" or not current_account:
            continue

        date = parse_date(first_value)
        if date is None:
            continue

        debit = row[7] if len(row) > 7 else None
        credit = row[8] if len(row) > 8 else None
        if is_empty_value(debit) and is_empty_value(credit):
            continue

        records.append(
            record(
                date,
                current_account,
                normalize_amount(debit, True),
                normalize_amount(credit, True),
                description=row[4] if len(row) > 4 else None,
            )
        )
        if len(records) % 1000 == 0:
            percent = 35 + int((row_number / total_rows) * 7)
            report_progress(progress_callback, min(percent, 42), f"Germany Kontoblatt 거래행을 읽는 중... ({len(records)}행)")

    if not records:
        raise ValueError("Germany Kontoblatt에서 날짜가 있는 거래 데이터를 찾지 못했습니다.")
    return records


def parse_date(value):
    if hasattr(value, "year"):
        return value
    if not isinstance(value, str):
        return None
    try:
        return datetime.strptime(value, DATE_FORMAT)
    except ValueError:
        return None
