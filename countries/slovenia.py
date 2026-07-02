from .common import find_header_row_and_columns, is_empty_value, normalize_amount, record, report_progress


def build_records(sheet, progress_callback=None):
    required = [
        "KONTO",
        "DATUM_KNJIZENJA",
        "OPIS_DOKUMENTA",
        "DEBET",
        "KREDIT",
        "PARTNER_NAZIV",
    ]
    header_row, columns = find_header_row_and_columns(sheet, required)
    indexes = {header: column - 1 for header, column in columns.items()}
    records = []
    total_rows = max(sheet.max_row - header_row, 1)

    for row_number, row in enumerate(
        sheet.iter_rows(min_row=header_row + 1, values_only=True),
        start=header_row + 1,
    ):
        date = row[indexes["DATUM_KNJIZENJA"]]
        if not hasattr(date, "year"):
            continue

        account = row[indexes["KONTO"]]
        debit = row[indexes["DEBET"]]
        credit = row[indexes["KREDIT"]]
        if is_empty_value(account) and is_empty_value(debit) and is_empty_value(credit):
            continue

        records.append(
            record(
                date,
                account,
                normalize_amount(debit, True),
                normalize_amount(credit, True),
                row[indexes["PARTNER_NAZIV"]],
                row[indexes["OPIS_DOKUMENTA"]],
            )
        )
        if len(records) % 1000 == 0:
            percent = 35 + int(((row_number - header_row) / total_rows) * 7)
            report_progress(progress_callback, min(percent, 42), f"Slovenia GL 거래행을 읽는 중... ({len(records)}행)")

    if not records:
        raise ValueError("Slovenia GL에서 날짜가 있는 거래 데이터를 찾지 못했습니다.")
    return records
