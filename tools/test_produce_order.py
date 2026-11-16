import unittest
import io

from openpyxl import load_workbook

from build_till_binders import normalize_printed_produce_text, quick_sheet
from produce_order import display_unique_rows, order_produce, product_key


def row(source_id, plu, name, counterparts=""):
    return dict(source_id=source_id, source_ids=["original-" + source_id], plu=plu,
                name=name, organic="organic" in name.lower(),
                counterpart_source_ids=counterparts)


class ProduceOrderTests(unittest.TestCase):
    def test_orders_by_main_item_and_variety_across_categories(self):
        rows = [
            row("1", "3", "Organic Royal Gala Apples"),
            row("2", "2", "Apples Gala Small"),
            row("3", "1", "Apples Ambrosia"),
            row("4", "4", "Apricots"),
            row("5", "5", "Avocado Large Hass"),
            row("6", "6", "Banana Baby"),
            row("7", "7", "Bananas"),
            row("8", "8", "Organic Bananas"),
        ]
        ordered = order_produce(rows)
        self.assertEqual([r["source_id"] for r in ordered], ["3", "2", "1", "4", "5", "7", "6", "8"])
        self.assertEqual([r["display_section"] for r in ordered], ["A"] * 5 + ["B"] * 3)
        self.assertEqual(ordered[0]["display_name"], "APPLES AMBROSIA")
        self.assertEqual(rows[0]["name"], "Organic Royal Gala Apples")

    def test_subgroups_by_type_with_asian_produce_last(self):
        rows = [row("1", "1", "Bok Choy"), row("2", "2", "Herbs Basil"),
                row("3", "3", "Eggplant Italian"), row("4", "4", "Apples Fuji"),
                row("5", "5", "Mint"), row("6", "6", "Apples Gala"),
                row("7", "7", "Zucchini"), row("8", "8", "Napa Cabbage")]
        ordered = order_produce(rows)
        self.assertEqual([(r["display_section"], r["display_group"], r["name"]) for r in ordered], [
            ("A", "Apples", "Apples Fuji"), ("A", "Apples", "Apples Gala"),
            ("E", "Eggplant", "Eggplant Italian"), ("H", "Herbs", "Herbs Basil"),
            ("H", "Herbs", "Mint"), ("Z", "Zucchini", "Zucchini"),
            ("ASIAN PRODUCE", "", "Bok Choy"), ("ASIAN PRODUCE", "", "Napa Cabbage")])

    def test_organic_follows_its_conventional_variety_within_subgroup(self):
        rows = [row("1", "94131", "Organic Apples Fuji", "3"), row("2", "4133", "Apples Gala"),
                row("3", "4131", "Apples Fuji"), row("4", "94133", "Organic Apples Gala", "2")]
        self.assertEqual([r["source_id"] for r in order_produce(rows)], ["3", "1", "2", "4"])

    def test_specific_names_and_reviewed_pairing(self):
        self.assertEqual(product_key("Melon Cantaloupe Large")[0], "CANTALOUPE")
        self.assertEqual(product_key("Organic Oranges Mango")[0], "ORANGE")
        self.assertEqual(product_key("Pepper Sweet Banana")[0], "PEPPER")
        self.assertEqual(product_key("Squash Acorn/Pepper")[0], "SQUASH")
        rows = [row("1", "01", "Lettuce Red Leaf"),
                row("2", "02", "Organic Leaf Lettuce Red", "1"),
                row("3", "03", "Lettuce Romaine")]
        self.assertEqual([r["source_id"] for r in order_produce(rows)], ["1", "2", "3"])
        self.assertFalse(product_key("Unrecognized Produce")[2])

    def test_exact_duplicates_are_suppressed_but_identity_and_conflicts_remain(self):
        rows = [row("1", "00042", "Apple Gala"), row("2", "00042", "Apple Gala"),
                row("3", "00042", "Apple Fuji"), row("4", "42", "Apple Gala")]
        displayed, duplicates = display_unique_rows(rows)
        self.assertEqual([r["source_id"] for r in displayed], ["1", "3", "4"])
        self.assertEqual(duplicates, [dict(
            source_id="2", source_ids="original-2", plu="00042", name="Apple Gala",
            displayed_source_id="1", displayed_source_ids="original-1")])

    def test_sheet_uses_letter_heading_code_first_and_uppercase(self):
        rows = [dict(r, duplicate_count=1) for r in order_produce([
            row("1", "00042", "Organic Apple Gala"),
            row("2", "42", "Apple Gala"),
            row("3", "50", "Bananas"),
        ])]
        content, coordinates, _ = quick_sheet(rows, {}, False)
        workbook = load_workbook(io.BytesIO(content))
        self.addCleanup(workbook.close)
        sheet = workbook["Produce"]
        self.assertEqual(sheet["A4"].value, "A")
        self.assertEqual(sheet["A5"].value, "Apples")
        self.assertEqual([(sheet[e["code_cell"]].value, sheet[e["description_cell"]].value)
                          for e in coordinates],
                         [("42", "APPLE GALA"), ("00042", "ORGANIC APPLE GALA"),
                          ("50", "BANANAS")])

    def test_wrapped_hyphenated_names_remain_verifiable(self):
        self.assertIn("MEYER-BUCKBRAND",
                      normalize_printed_produce_text("ORGANIC LEMONS MEYER-\nBUCKBRAND"))


if __name__ == "__main__":
    unittest.main()
