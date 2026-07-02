import re


HEADER_SCAN_ROWS = 30
MATCH_HEADERS = ["날짜", "계정코드", "차변(EUR)", "대변(EUR)", "거래처명", "적요"]


def report_progress(progress_callback, percent, message):
    if progress_callback:
        progress_callback(percent, message)


def normalize_header(value):
    if value is None:
        return ""
    normalized = "".join(str(value).split()).lower()
    if re.fullmatch(r"차변\(.*\)", normalized):
        return "차변()"
    if re.fullmatch(r"대변\(.*\)", normalized):
        return "대변()"
    return normalized


def find_header_row_and_columns(sheet, required_headers):
    required = {normalize_header(header): header for header in required_headers}
    scan_limit = min(sheet.max_row, HEADER_SCAN_ROWS)

    for row in sheet.iter_rows(max_row=scan_limit):
        found = {}
        for cell in row:
            normalized = normalize_header(cell.value)
            if normalized in required and normalized not in found:
                found[normalized] = cell.column

        if all(normalize_header(header) in found for header in required_headers):
            return row[0].row, {
                required[normalized]: column for normalized, column in found.items()
            }

    missing_text = ", ".join(required_headers)
    raise ValueError(
        f'"{sheet.title}" 시트의 상위 {HEADER_SCAN_ROWS}행에서 '
        f"필수 헤더를 찾지 못했습니다: {missing_text}"
    )


def is_empty_value(value):
    return value is None or value == "" or (isinstance(value, str) and not value.strip())


def normalize_account_code(value):
    if value is None:
        return None
    if isinstance(value, float) and value.is_integer():
        return str(int(value))
    return str(value).strip()


def normalize_amount(value, absolute=False):
    if is_empty_value(value):
        return None
    if isinstance(value, (int, float)):
        return abs(value) if absolute else value
    return value


def record(date, account, debit, credit, partner=None, description=None):
    return {
        MATCH_HEADERS[0]: date,
        MATCH_HEADERS[1]: normalize_account_code(account),
        MATCH_HEADERS[2]: debit,
        MATCH_HEADERS[3]: credit,
        MATCH_HEADERS[4]: partner,
        MATCH_HEADERS[5]: description,
    }


def build_double_entry_records(
    sheet,
    headers,
    progress_label,
    progress_callback=None,
    partner_fallback_header=None,
    description_fallback_header=None,
):
    required = [
        headers["date"],
        headers["debit_account"],
        headers["credit_account"],
        headers["amount"],
        headers["partner"],
        headers["description"],
    ]
    if partner_fallback_header:
        required.append(partner_fallback_header)
    if description_fallback_header:
        required.append(description_fallback_header)

    header_row, columns = find_header_row_and_columns(sheet, required)
    records = []
    total_rows = max(sheet.max_row - header_row, 1)

    for row_number in range(header_row + 1, sheet.max_row + 1):
        date_value = sheet.cell(row_number, columns[headers["date"]]).value
        if not hasattr(date_value, "year"):
            continue

        debit_account = sheet.cell(row_number, columns[headers["debit_account"]]).value
        credit_account = sheet.cell(row_number, columns[headers["credit_account"]]).value
        signed_amount = normalize_amount(
            sheet.cell(row_number, columns[headers["amount"]]).value
        )
        if (
            is_empty_value(debit_account)
            or is_empty_value(credit_account)
            or not isinstance(signed_amount, (int, float))
        ):
            continue

        partner = sheet.cell(row_number, columns[headers["partner"]]).value
        if partner_fallback_header and is_empty_value(partner):
            partner = sheet.cell(row_number, columns[partner_fallback_header]).value

        description = sheet.cell(row_number, columns[headers["description"]]).value
        if description_fallback_header and is_empty_value(description):
            description = sheet.cell(row_number, columns[description_fallback_header]).value

        amount = abs(signed_amount)
        if signed_amount < 0:
            debit_account, credit_account = credit_account, debit_account

        records.append(
            record(date_value, debit_account, amount, None, partner, description)
        )
        records.append(
            record(date_value, credit_account, None, amount, partner, description)
        )

        processed = len(records) // 2
        if processed % 1000 == 0:
            percent = 35 + int(((row_number - header_row) / total_rows) * 7)
            report_progress(
                progress_callback,
                min(percent, 42),
                f"{progress_label} 거래행을 읽는 중... "
                f"({processed}행, 출력 {len(records)}행)",
            )

    if not records:
        raise ValueError(f"{progress_label}에서 날짜가 있는 거래 데이터를 찾지 못했습니다.")
    return records
