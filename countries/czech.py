from .common import build_double_entry_records


def build_records(sheet, progress_callback=None):
    return build_double_entry_records(
        sheet,
        {
            "date": "Datum",
            "debit_account": "MD",
            "credit_account": "DAL",
            "amount": "Částka",
            "partner": "Firma",
            "description": "Text",
        },
        "Czech GL",
        progress_callback,
        partner_fallback_header="Jméno",
        description_fallback_header="English",
    )
