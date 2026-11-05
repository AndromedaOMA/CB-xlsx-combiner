"""Regression checks for the Operations review workflow and both merge engines."""

import json
import subprocess
import tempfile
import unittest
from pathlib import Path

from openpyxl import Workbook, load_workbook

import main


ROOT = Path(__file__).resolve().parents[1]


def write_book(path, rows, title="Hotels"):
    path.parent.mkdir(parents=True, exist_ok=True)
    book = Workbook()
    book.active.title = title
    for row in rows:
        book.active.append(row)
    book.save(path)
    book.close()


def browser_merge(paths, directory):
    sources = []
    for path in paths:
        book = load_workbook(path)
        sources.append({"name": path.name, "sheets": {
            sheet.title: [list(row) for row in sheet.values] for sheet in book.worksheets
        }})
        book.close()
    fixture = directory / "browser-input.json"
    fixture.write_text(json.dumps(sources, default=str), encoding="utf-8")
    result = subprocess.run(
        ["node", str(ROOT / "tests/test_browser_merge.cjs"), str(fixture)],
        capture_output=True, text=True, encoding="utf-8", check=True,
    )
    return json.loads(result.stdout)


class MergeTests(unittest.TestCase):
    def setUp(self):
        self.temporary = tempfile.TemporaryDirectory()
        self.addCleanup(self.temporary.cleanup)
        self.directory = Path(self.temporary.name)

    def merge(self, left, right):
        # Same basenames must still have independent source references and highlights.
        paths = [self.directory / "a/Hotels.xlsx", self.directory / "b/Hotels.xlsx"]
        for path, rows in zip(paths, [left, right]):
            write_book(path, rows)
        output = self.directory / "review.xlsx"
        summary = main.merge_workbooks(paths, output)
        book = load_workbook(output)
        self.addCleanup(book.close)
        return paths, book, summary

    def test_hotel_blocks_fields_emails_and_browser_parity(self):
        left = [
            ["Hotel name", "Stars", "Email", "email", "Phone", "Custom note"],
            ["Zulu Hotel", "4*", "info@zulu.test; SALES@zulu.test", "person@zulu.test", "+31 123", "Keep me"],
            ["Beta Hotel", "3*", "beta@example.test"],
            ["Alpha Hotel", 3, "alpha@example.test"],
            ["Luxury Hotel", "5*", None],
            ["Unknown Hotel", "NS", None],
        ]
        right = [
            ["name", "google_stars", "google_rating", "Email adress 2", "email", "address", "website"],
            ["ZULU HOTEL", "4", "4.7", "sales@zulu.test, groups@zulu.test; reservations@zulu.test", "events@zulu.test; other@zulu.test", "1 Main Street", "https://zulu.test"],
        ]
        paths, book, summary = self.merge(left, right)
        sheet = book["Merged Hotels"]
        rows = list(sheet.values)
        headers = list(rows[0])
        status = headers.index("Row type")
        primary = [row for row in rows[1:] if row[status] != "SOURCE DUPLICATE"]
        self.assertEqual([row[0] for row in primary], ["Alpha Hotel", "Beta Hotel", "Zulu Hotel", "Luxury Hotel", "Unknown Hotel"])
        self.assertEqual(summary["records"], 6)
        self.assertEqual(summary["merged_groups"], 1)
        zulu = next(i for i, row in enumerate(rows, 1) if row[status] == "MERGED - REVIEW")
        row = rows[zulu - 1]
        self.assertEqual(row[:3], ("Zulu Hotel", "4*", 4.7))
        self.assertEqual(row[3:8], ("SALES@zulu.test", "groups@zulu.test", "reservations@zulu.test", "events@zulu.test", "info@zulu.test"))
        self.assertEqual(row[8:11], ("1 Main Street", "+31 123", "https://zulu.test"))
        self.assertEqual(row[headers.index("Custom note")], "Keep me")
        self.assertIn("2 additional email(s)", row[-1])
        self.assertEqual(sheet.cell(zulu, 1).fill.fgColor.rgb[-6:], "FCE4D6")
        for number in [zulu + 1, zulu + 2]:
            self.assertEqual(sheet.cell(number, status + 1).value, "SOURCE DUPLICATE")
            self.assertEqual(sheet.cell(number, 8).fill.fgColor.rgb[-6:], "F4CCCC")
        sources = [book[name] for name in book.sheetnames if name.startswith("Source")]
        self.assertEqual(sources[0]["C2"].value, left[1][2])
        self.assertEqual(sources[1]["D2"].value, right[1][3])
        self.assertEqual(sources[0]["A2"].fill.fgColor.rgb[-6:], "F4CCCC")
        self.assertEqual(sources[1]["A2"].fill.fgColor.rgb[-6:], "F4CCCC")
        browser = browser_merge(paths, self.directory)
        self.assertEqual(browser["rows"], [list(row) for row in rows])
        self.assertEqual(browser["fills"][str(zulu - 1)], "FCE4D6")
        self.assertEqual(browser["fills"][str(zulu)], "F4CCCC")
        self.assertEqual(browser["firstSheet"], "Merged Hotels")

    def test_conflicting_stars_are_not_selected_arbitrarily(self):
        paths, book, _ = self.merge(
            [["name", "stars", "Google rating"], ["One Hotel", 3, "4,8"]],
            [["hotel", "google_stars", "googlerating"], ["One Hotel", 4, 4.8]],
        )
        row = list(book["Merged Hotels"].values)[1]
        self.assertEqual(row[1:3], ("Conflicting / review", 4.8))
        self.assertIn("Conflicting stars: 3*, 4*", row[-1])
        self.assertEqual(browser_merge(paths, self.directory)["rows"][1], list(row))

    def test_location_conflict_cannot_be_bridged_by_an_empty_record(self):
        paths, book, summary = self.merge(
            [["name", "email", "address"], ["Bridge Hotel", "a@bridge.test", "1 First Road"], ["Bridge Hotel", "b@bridge.test", None]],
            [["name", "email", "address"], ["Bridge Hotel", "c@bridge.test", "2 Second Road"]],
        )
        self.assertEqual(summary["merged_groups"], 1)
        self.assertEqual(summary["single_records"], 1)
        self.assertEqual(len(list(book["Merged Hotels"].values)), 5)
        self.assertEqual(browser_merge(paths, self.directory)["rows"], [list(row) for row in book["Merged Hotels"].values])

    def test_headerless_contacts_and_category_rows(self):
        paths, book, _ = self.merge(
            [[None], ["Headerless Hotel", "4*", "one@test.example; two@test.example", "groups@test.example"]],
            [["name", "google_stars", "email", "google_rating"], ["3 Stele Athens"], ["Section Hotel", None, None, 4.9], [None, None, "unassigned@test.example"]],
        )
        rows = list(book["Merged Hotels"].values)
        self.assertEqual(rows[1][:3], ("Section Hotel", "3*", 4.9))
        self.assertNotIn("field_4", rows[0])
        self.assertEqual(rows[2][3:6], ("groups@test.example", "one@test.example", "two@test.example"))
        self.assertTrue(any("unassigned@test.example" in str(row) for row in book["Unassigned Rows"].values))
        self.assertEqual(browser_merge(paths, self.directory)["rows"], [list(row) for row in rows])

    def test_distinct_emails_and_a_shared_chain_contact_do_not_merge_properties(self):
        paths, book, summary = self.merge(
            [["name", "email", "address"], ["The Meridian Hotel", "a.b@meridian.test", None], ["Brand East Hotel", "groups@brand.test", "1 East Road"]],
            [["name", "email", "address"], ["Meridian Hotel", "ab@meridian.test", None], ["Brand West Hotel", "groups@brand.test", "2 West Road"]],
        )
        self.assertEqual(summary["merged_groups"], 0)
        self.assertEqual(summary["single_records"], 4)
        self.assertEqual(browser_merge(paths, self.directory)["rows"], [list(row) for row in book["Merged Hotels"].values])

    def test_real_workbooks_preserve_every_record_and_keep_blocks_sorted(self):
        paths = [ROOT / "Athens.xlsx", ROOT / "Athens (1).xlsx"]
        output = self.directory / "athens-review.xlsx"
        summary = main.merge_workbooks(paths, output)
        self.assertEqual(summary["records"], 338)
        book = load_workbook(output)
        self.addCleanup(book.close)
        rows = list(book["Merged Hotels"].values)
        headers = list(rows[0])
        status, count = headers.index("Row type"), headers.index("Source record count")
        primary = [row for row in rows[1:] if row[status] != "SOURCE DUPLICATE"]
        self.assertEqual(sum(row[count] for row in primary), summary["records"])
        order = lambda row: ({"3*": 0, "4*": 1, "5*": 2}.get(row[1], 3), main.normalized(row[0]))
        self.assertEqual(primary, sorted(primary, key=order))
        for i, row in enumerate(rows):
            if row[status] == "MERGED - REVIEW":
                self.assertTrue(all(original[status] == "SOURCE DUPLICATE" for original in rows[i + 1:i + 1 + row[count]]))
        browser = browser_merge(paths, self.directory)
        self.assertEqual(len(browser["rows"]), len(rows))
        self.assertEqual(browser["rows"][0], list(rows[0]))
        self.assertEqual(browser["rows"], [list(row) for row in rows])

    def test_amsterdam_reference_retains_manually_entered_contacts(self):
        records = main.workbook_rows(ROOT / "Amsterdam.xlsx")
        self.assertEqual(len(records), 114)
        hotel = next(record for record in records if record["name"].startswith("Anantara"))
        values, _ = main.hotel_values([hotel])
        self.assertEqual(values[1:5], ["5*", 4.6, "rsv.krasnapolsky@anantara-hotels.com", "groups.krasnapolsky@anantara-hotels.com"])
        self.assertFalse(any(record["sheet"] == "CATALONIA" for record in records))


if __name__ == "__main__":
    unittest.main()
