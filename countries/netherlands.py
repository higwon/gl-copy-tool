import re

from .common import find_header_row_and_columns, is_empty_value, normalize_amount, record, report_progress

ACCOUNT_BLOCK_RE = re.compile(r"^\s*([^\s]+)\s+-\s+.+$")


def build_records(sheet, progress_callback=None):
    required = ["Date", "Description", "Debit", "Credit", "Account name"]
    header_row, columns = find_header_row_and_columns(sheet, required)
    current_account = None
    records = []
    total_rows = max(sheet.max_row - header_row, 1)

    for row_number in range(header_row + 1, sheet.max_row + 1):
        first_value = sheet.cell(row_number, 1).value
        second_value = sheet.cell(row_number, 2).value
        match = ACCOUNT_BLOCK_RE.match(first_value) if isinstance(first_value, str) else None
        if match and is_empty_value(second_value):
            current_account = match.group(1).strip()
            continue
        if not hasattr(first_value, "year") or not current_account:
            continue

        debit = sheet.cell(row_number, columns["Debit"]).value
        credit = sheet.cell(row_number, columns["Credit"]).value
        description = sheet.cell(row_number, columns["Description"]).value
        if is_empty_value(debit) and is_empty_value(credit) and is_empty_value(description):
            continue
        records.append(
            record(
                first_value,
                current_account,
                normalize_amount(debit, True),
                normalize_amount(credit, True),
                sheet.cell(row_number, columns["Account name"]).value,
                description,
            )
        )
        if len(records) % 1000 == 0:
            percent = 35 + int(((row_number - header_row) / total_rows) * 7)
            report_progress(progress_callback, min(percent, 42), f"Netherlands GL 거래행을 읽는 중... ({len(records)}행)")

    if not records:
        raise ValueError("Netherlands GL에서 날짜가 있는 거래 데이터를 찾지 못했습니다.")
    return records
