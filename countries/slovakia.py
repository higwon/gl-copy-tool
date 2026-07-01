from .common import build_double_entry_records


def build_records(sheet, progress_callback=None):
    return build_double_entry_records(
        sheet,
        {
            "date": "Dátum",
            "debit_account": "MD",
            "credit_account": "DAL",
            "amount": "Čiastka",
            "partner": "Firma",
            "description": "Text",
        },
        "Slovakia GL",
        progress_callback,
    )
