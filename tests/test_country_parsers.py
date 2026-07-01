import unittest
from datetime import datetime

from openpyxl import Workbook

from countries import IMPORT_FORMATS, build_source_records
from countries.common import MATCH_HEADERS


def worksheet(headers, rows):
    workbook = Workbook()
    sheet = workbook.active
    sheet.append(headers)
    for row in rows:
        sheet.append(row)
    return workbook, sheet


class CountryParserTests(unittest.TestCase):
    def label_for(self, module_name):
        return next(
            label
            for label, parser in IMPORT_FORMATS.items()
            if parser.__module__.endswith(module_name)
        )

    def test_korea(self):
        workbook, sheet = worksheet(
            MATCH_HEADERS,
            [[datetime(2026, 1, 1), "1000", -10, 0, "Partner", "Memo"]],
        )
        records = build_source_records(sheet, self.label_for("korea"))
        workbook.close()
        self.assertEqual(10, records[0][MATCH_HEADERS[2]])

    def test_netherlands(self):
        workbook, sheet = worksheet(
            ["Date", "Description", "Debit", "Credit", "Account name"],
            [
                ["1000 - Cash", None, None, None, None],
                [datetime(2026, 1, 1), "Memo", 25, None, "Cash"],
            ],
        )
        records = build_source_records(sheet, self.label_for("netherlands"))
        workbook.close()
        self.assertEqual("1000", records[0][MATCH_HEADERS[1]])

    def test_austria(self):
        workbook, sheet = worksheet(
            ["Kto-Nr", "Beleg-Dat", "Text", "GW-Soll", "GW-Haben"],
            [["1000", datetime(2026, 1, 1), "Memo", -30, None]],
        )
        records = build_source_records(sheet, self.label_for("austria"))
        workbook.close()
        self.assertEqual(30, records[0][MATCH_HEADERS[2]])

    def test_hungary_without_hint_row(self):
        workbook, sheet = worksheet(
            [
                "Budgetary account",
                "Számviteli teljesítés",
                "Tartozik",
                "Követel",
                "Megjegyzés",
                "Partner",
            ],
            [["1000", datetime(2026, 1, 1), 40, 0, "Memo", "Partner"]],
        )
        records = build_source_records(sheet, self.label_for("hungary"))
        workbook.close()
        self.assertEqual(1, len(records))

    def test_poland_ignores_total_row(self):
        workbook, sheet = worksheet(
            ["Konto", "Data księgowania", "Kwota Wn", "Kwota Ma", "Nazwa podmiotu", "Opis"],
            [
                ["1000", datetime(2026, 1, 1), 50, None, "Partner", "Memo"],
                [None, None, 50, 50, None, None],
            ],
        )
        records = build_source_records(sheet, self.label_for("poland"))
        workbook.close()
        self.assertEqual(1, len(records))

    def test_czech_reverses_negative_double_entry(self):
        workbook, sheet = worksheet(
            ["Datum", "MD", "DAL", "Částka", "Firma", "Jméno", "Text", "English"],
            [[datetime(2026, 1, 1), "1000", "2000", -60, None, "Person", None, "Memo"]],
        )
        records = build_source_records(sheet, self.label_for("czech"))
        workbook.close()
        self.assertEqual("2000", records[0][MATCH_HEADERS[1]])
        self.assertEqual("Person", records[0][MATCH_HEADERS[4]])
        self.assertEqual("Memo", records[0][MATCH_HEADERS[5]])

    def test_slovakia_splits_double_entry(self):
        workbook, sheet = worksheet(
            ["Dátum", "MD", "DAL", "Čiastka", "Firma", "Text"],
            [[datetime(2026, 1, 1), "1000", "2000", 70, "Partner", "Memo"]],
        )
        records = build_source_records(sheet, self.label_for("slovakia"))
        workbook.close()
        self.assertEqual(2, len(records))
        self.assertEqual(70, records[0][MATCH_HEADERS[2]])
        self.assertEqual(70, records[1][MATCH_HEADERS[3]])


if __name__ == "__main__":
    unittest.main()
