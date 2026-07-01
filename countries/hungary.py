from .common import find_header_row_and_columns, is_empty_value, normalize_amount, record, report_progress


def build_records(sheet, progress_callback=None):
    required = ["Budgetary account", "Számviteli teljesítés", "Tartozik", "Követel", "Megjegyzés", "Partner"]
    header_row, columns = find_header_row_and_columns(sheet, required)
    records = []
    total_rows = max(sheet.max_row - header_row, 1)

    for row_number in range(header_row + 1, sheet.max_row + 1):
        date = sheet.cell(row_number, columns["Számviteli teljesítés"]).value
        if not hasattr(date, "year"):
            continue
        account = sheet.cell(row_number, columns["Budgetary account"]).value
        debit = sheet.cell(row_number, columns["Tartozik"]).value
        credit = sheet.cell(row_number, columns["Követel"]).value
        if is_empty_value(account) and is_empty_value(debit) and is_empty_value(credit):
            continue
        records.append(record(date, account, normalize_amount(debit, True), normalize_amount(credit, True), sheet.cell(row_number, columns["Partner"]).value, sheet.cell(row_number, columns["Megjegyzés"]).value))
        if len(records) % 1000 == 0:
            percent = 35 + int(((row_number - header_row) / total_rows) * 7)
            report_progress(progress_callback, min(percent, 42), f"Hungary Novitax 거래행을 읽는 중... ({len(records)}행)")
    if not records:
        raise ValueError("Hungary Novitax Export에서 날짜가 있는 거래 데이터를 찾지 못했습니다.")
    return records
