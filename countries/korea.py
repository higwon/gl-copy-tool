from .common import MATCH_HEADERS, find_header_row_and_columns, normalize_amount, report_progress


def build_records(sheet, progress_callback=None):
    header_row, columns = find_header_row_and_columns(sheet, MATCH_HEADERS)
    records = []
    total_rows = max(sheet.max_row - header_row, 1)

    for row_number in range(header_row + 1, sheet.max_row + 1):
        values = {
            header: sheet.cell(row_number, columns[header]).value
            for header in MATCH_HEADERS
        }
        if not any(value not in (None, "") for value in values.values()):
            continue
        values[MATCH_HEADERS[2]] = normalize_amount(values[MATCH_HEADERS[2]], absolute=True)
        values[MATCH_HEADERS[3]] = normalize_amount(values[MATCH_HEADERS[3]], absolute=True)
        records.append(values)

        if len(records) % 1000 == 0:
            percent = 35 + int(((row_number - header_row) / total_rows) * 7)
            report_progress(progress_callback, min(percent, 42), f"Source GL을 읽는 중... ({len(records)}행)")
    return records
