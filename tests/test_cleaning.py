"""
tests.test_cleaning
~~~~~~~~~~~~~~~~~~~

Tests for whole book cleaning and validation.

The scenarios here focus on problems that only appear after cell values
are normalised: duplicate keys that emerge from stripping or type
conversion, and references that break or resolve on the cleaned values.
"""

import copy
from unittest import TestCase

import pyexcel as p
from pyexcel import (
    CleaningError,
    CleaningPreview,
    ColumnCleaningRule,
    ReferenceRule,
    UniqueRule,
)
from pyexcel.cleaning import (
    BLOCKING,
    ERROR,
    REPORT_HEADERS,
    clean_float,
    clean_int,
    clean_strip,
    clean_str,
    run_cleaning,
)


def make_book(sheet_data, named=True):
    """Build a book, optionally naming each sheet's columns by row 0."""
    book = p.get_book(bookdict=copy.deepcopy(sheet_data))
    if named:
        for sheet in book:
            sheet.name_columns_by_row(0)
    return book


CUSTOMERS = {
    "customers": [
        ["customer_id", "name", "note"],
        [" 1001", "Alice", "x"],
        [1002, "Carol", "y"],
        [1003, "Dan", "z"],
    ],
}

ORDERS = {
    "orders": [
        ["order_id", "customer_id", "amount"],
        [1, " 1001 ", 10],
        [2, 1002, 20],
        [3, "", 40],
    ],
}


def passing_rules():
    return [
        ColumnCleaningRule("customers", "customer_id", "strip"),
        UniqueRule("customers", "customer_id", "int", name="customer pk"),
        ColumnCleaningRule("orders", "customer_id", "strip"),
        ReferenceRule("orders", "customer_id", "customers", cleaner="int"),
    ]


class TestBuiltInCleaners(TestCase):
    def test_strip_only_touches_strings(self):
        assert clean_strip("  a ") == "a"
        assert clean_strip(1) == 1
        assert clean_strip(1.5) == 1.5
        assert clean_strip(False) is False
        assert clean_strip(None) is None
        assert clean_strip("") == ""

    def test_none_empty_false_and_numbers_are_distinct(self):
        assert clean_strip(None) is None
        assert clean_strip("") == ""
        assert clean_strip(False) is False
        assert clean_strip(0) == 0

    def test_clean_str_keeps_none_and_renders_bools(self):
        assert clean_str(None) is None
        assert clean_str("") == ""
        assert clean_str(False) == "false"
        assert clean_str(True) == "true"
        assert clean_str(1) == "1"

    def test_clean_int_strict(self):
        assert clean_int("12") == 12
        assert clean_int(" 12 ") == 12
        assert clean_int("12.0") == 12
        assert clean_int(12.0) == 12
        assert clean_int(12) == 12
        assert clean_int(None) is None
        for bad in ["abc", "12.5", "", "1,000"]:
            with self.assertRaises(Exception):
                clean_int(bad)
        with self.assertRaises(Exception):
            clean_int(12.5)
        with self.assertRaises(Exception):
            clean_int(False)

    def test_clean_float_strict(self):
        assert clean_float("1.5") == 1.5
        assert clean_float(2) == 2.0
        assert clean_float(None) is None
        for bad in ["abc", ""]:
            with self.assertRaises(Exception):
                clean_float(bad)
        with self.assertRaises(Exception):
            clean_float(True)

    def test_legacy_format_interface_unchanged(self):
        # the legacy to_format silently keeps unconvertible values;
        # the strict cleaners must not change that behaviour
        from pyexcel.internal.sheets.formatters import to_format

        assert to_format(int, "abc") == "abc"
        assert to_format(float, "abc") == "abc"
        # strict cleaners raise where the legacy interface stays lenient
        with self.assertRaises(Exception):
            clean_int("abc")


class TestPreviewIsolation(TestCase):
    def setUp(self):
        self.book = make_book(
            {
                "customers": CUSTOMERS["customers"],
                "orders": ORDERS["orders"],
            },
        )
        self.original_orders = [
            list(row) for row in self.book["orders"].to_array()
        ]
        self.original_customers = [
            list(row) for row in self.book["customers"].to_array()
        ]

    def test_preview_does_not_mutate_book(self):
        preview = self.book.clean(passing_rules())
        assert preview.passed is True
        assert self.book["orders"].to_array() == self.original_orders
        assert (
            self.book["customers"].to_array() == self.original_customers
        )
        # still raw string with spaces
        assert self.book["orders"].cell_value(0, 1) == " 1001 "

    def test_apply_writes_cleaned_values_once(self):
        preview = self.book.clean(passing_rules())
        returned_book = preview.apply()
        assert returned_book is self.book
        orders = self.book["orders"]
        # whitespace removed and value converted to int
        assert orders.cell_value(0, 1) == 1001
        assert isinstance(orders.cell_value(0, 1), int)
        assert orders.cell_value(1, 1) == 1002
        # absent relation left as empty string
        assert orders.cell_value(2, 1) == ""
        customers = self.book["customers"]
        # the unique rule's int cleaner is applied to its key column too
        assert customers.cell_value(0, 0) == 1001
        assert isinstance(customers.cell_value(0, 0), int)
        # strip alone keeps text
        assert customers.cell_value(0, 1) == "Alice"

    def test_apply_refuses_when_errors_exist(self):
        data = {
            "customers": [
                ["customer_id"], ["1001"], ["1001"],
            ],
            "orders": [["customer_id"], ["9999"]],
        }
        book = make_book(data)
        preview = book.clean(
            [
                UniqueRule("customers", "customer_id", "int"),
                ReferenceRule(
                    "orders", "customer_id", "customers", cleaner="int",
                ),
            ],
        )
        assert preview.passed is False
        with self.assertRaises(CleaningError):
            preview.apply()
        # nothing was written
        assert book["customers"].cell_value(0, 0) == "1001"

    def test_half_cleaned_book_is_impossible(self):
        # customers fails (duplicate) so orders must stay untouched too
        data = {
            "customers": [["customer_id"], [" 1"], ["1 "]],
            "orders": [["customer_id"], ["1"]],
        }
        book = make_book(data)
        rules = [
            ColumnCleaningRule("customers", "customer_id", "strip"),
            UniqueRule("customers", "customer_id"),
            ColumnCleaningRule("orders", "customer_id", "strip"),
        ]
        preview = book.clean(rules)
        assert preview.passed is False
        with self.assertRaises(CleaningError):
            preview.apply()
        assert book["orders"].cell_value(0, 0) == "1"
        assert book["customers"].cell_value(0, 0) == " 1"

    def test_existing_sheet_and_accessors_see_final_values(self):
        held_sheet = self.book["orders"]
        held_column = held_sheet.column
        held_row = held_sheet.row
        self.book.clean(passing_rules()).apply()
        assert held_sheet.cell_value(0, 1) == 1001
        assert held_column["customer_id"] == [1001, 1002, ""]
        assert held_row[0] == [1, 1001, 10]

    def test_sheet_order_names_and_unselected_fields_preserved(self):
        self.book.clean(passing_rules()).apply()
        assert self.book.sheet_names() == ["customers", "orders"]
        assert self.book["customers"].colnames == [
            "customer_id", "name", "note",
        ]
        # unselected column never touched
        assert self.book["customers"].column["note"] == ["x", "y", "z"]
        assert self.book["orders"].colnames == [
            "order_id", "customer_id", "amount",
        ]


class TestUniquenessChecks(TestCase):
    def test_duplicate_emerges_after_strip(self):
        book = make_book(
            {"items": [["code"], [" A"], ["A"], ["B"]]},
        )
        preview = book.clean(
            [UniqueRule("items", "code", "strip", name="code uk")],
        )
        assert preview.passed is False
        rows = preview.error_sheet().to_array()
        severities = [row[0] for row in rows[1:]]
        assert severities == [ERROR, ERROR]
        messages = [row[-1] for row in rows[1:]]
        assert all("duplicate key" in message for message in messages)
        locations = sorted(row[4] for row in rows[1:])
        assert locations == [0, 1]

    def test_duplicate_emerges_after_type_conversion(self):
        book = make_book(
            {"items": [["code"], [1], ["1"], [2.0]]},
        )
        preview = book.clean(
            [UniqueRule("items", "code", "int")],
        )
        # 1, "1" and 2.0 -> 1, 1, 2: the two 1s collide
        assert preview.passed is False
        rows = preview.error_sheet().to_array()[1:]
        assert len(rows) == 2
        assert all(row[0] == ERROR for row in rows)
        assert sorted(row[4] for row in rows) == [0, 1]

    def test_composite_unique_key(self):
        book = make_book(
            {
                "lines": [
                    ["head", "tail", "v"],
                    ["a", "b", 1],
                    ["a ", " b", 2],   # same composite after strip
                    ["a", "c", 3],
                ],
            },
        )
        preview = book.clean(
            [UniqueRule("lines", ["head", "tail"], "strip")],
        )
        assert preview.passed is False
        rows = preview.error_sheet().to_array()[1:]
        assert len(rows) == 2
        assert rows[0][6] == ("a", "b")

    def test_unique_passes_when_distinct(self):
        book = make_book({"items": [["code"], ["A"], ["B"]]})
        preview = book.clean([UniqueRule("items", "code", "strip")])
        assert preview.passed is True
        assert preview.error_sheet().to_array() == [REPORT_HEADERS]

    def test_absent_key_rows_are_not_duplicates(self):
        book = make_book(
            {"items": [["code"], [""], [""], ["A"]]},
        )
        preview = book.clean([UniqueRule("items", "code", "strip")])
        assert preview.passed is True


class TestReferenceChecks(TestCase):
    def test_reference_resolves_after_cleaning(self):
        book = make_book(
            {
                "customers": [["customer_id"], [" 1001"], [1002]],
                "orders": [["customer_id"], ["1001 "], [1002], [""]],
            },
        )
        preview = book.clean(
            [
                ReferenceRule(
                    "orders", "customer_id", "customers", cleaner="strip",
                ),
            ],
        )
        assert preview.passed is True

    def test_broken_reference_is_reported(self):
        book = make_book(
            {
                "customers": [["customer_id"], [1001]],
                "orders": [["customer_id"], [1001], [9999]],
            },
        )
        preview = book.clean(
            [
                ReferenceRule(
                    "orders", "customer_id", "customers", cleaner="int",
                ),
            ],
        )
        assert preview.passed is False
        rows = preview.error_sheet().to_array()[1:]
        assert len(rows) == 1
        assert rows[0][0] == ERROR
        assert rows[0][3] == "orders"
        assert rows[0][4] == 1
        assert "broken reference" in rows[0][7]

    def test_absent_foreign_key_is_not_missing_association(self):
        book = make_book(
            {
                "customers": [["customer_id"], [1001]],
                "orders": [["customer_id"], [1001], [""], [None]],
            },
        )
        preview = book.clean(
            [
                ReferenceRule(
                    "orders", "customer_id", "customers", cleaner="int",
                ),
            ],
        )
        # the empty rows carry no relation: no false missing association
        assert preview.passed is True

    def test_composite_reference(self):
        book = make_book(
            {
                "parent": [
                    ["a", "b"], [1, 2], [3, 4],
                ],
                "child": [
                    ["x", "y"], [" 1", "2"], ["3", "9"], ["", ""],
                ],
            },
        )
        preview = book.clean(
            [
                ReferenceRule(
                    "child",
                    ["x", "y"],
                    "parent",
                    referenced_columns=["a", "b"],
                    cleaner="int",
                ),
            ],
        )
        rows = preview.error_sheet().to_array()[1:]
        assert len(rows) == 1
        assert rows[0][4] == 1
        assert rows[0][6] == (3, 9)

    def test_mismatched_composite_width_rejected(self):
        make_book(
            {
                "parent": [["a", "b"], [1, 2]],
                "child": [["x"], [1]],
            },
        )
        with self.assertRaises(ValueError):
            ReferenceRule(
                "child", ["x"], "parent",
                referenced_columns=["a", "b"],
            )


class TestBlockingSemantics(TestCase):
    def test_conversion_failure_is_blocking_not_missing(self):
        book = make_book(
            {
                "customers": [["customer_id"], [1001]],
                "orders": [["customer_id"], ["nope"]],
            },
        )
        preview = book.clean(
            [
                ReferenceRule(
                    "orders", "customer_id", "customers", cleaner="int",
                ),
            ],
        )
        rows = preview.error_sheet().to_array()[1:]
        assert any(row[0] == BLOCKING for row in rows)
        # must not be mislabelled as a broken reference
        assert not any(
            "broken reference" in row[7] for row in rows
        )

    def test_unbuildable_target_keyset_blocks_whole_rule(self):
        book = make_book(
            {
                "customers": [["customer_id"], ["bad"]],
                "orders": [["customer_id"], [9999]],
            },
        )
        preview = book.clean(
            [
                ReferenceRule(
                    "orders", "customer_id", "customers", cleaner="int",
                ),
            ],
        )
        rows = preview.error_sheet().to_array()[1:]
        # rule blocked on the referenced side; no missing association error
        assert all(row[0] == BLOCKING for row in rows)
        assert any("not judged" in row[7] for row in rows)

    def test_unique_rule_marks_unjudged_rows_as_blocking(self):
        book = make_book(
            {"items": [["code"], ["1"], ["bad"], ["1"]]},
        )
        preview = book.clean([UniqueRule("items", "code", "int")])
        rows = preview.error_sheet().to_array()[1:]
        # duplicate 1s are still detected among the convertible rows
        assert any(
            row[0] == ERROR and "duplicate" in row[7] for row in rows
        )
        # and the bad cell plus a rule-level blocking note are reported
        assert any(
            row[0] == BLOCKING and row[4] == 1 for row in rows
        )
        assert any(
            row[0] == BLOCKING and "could not be fully judged" in row[7]
            for row in rows
        )
        assert preview.blocking_errors

    def test_custom_cleaner_exception_blocks(self):
        def boom(value):
            if value == "x":
                raise RuntimeError("kaboom")
            return value

        book = make_book({"items": [["code"], ["x"], ["y"]]})
        preview = book.clean(
            [ColumnCleaningRule("items", "code", boom)],
        )
        rows = preview.error_sheet().to_array()[1:]
        assert len(rows) == 1
        assert rows[0][0] == BLOCKING
        assert "kaboom" in rows[0][7]


class TestErrorSheet(TestCase):
    def test_report_is_native_sheet_with_headers(self):
        book = make_book(
            {
                "customers": [["customer_id"], [1], [1]],
                "orders": [["customer_id"], [9]],
            },
        )
        preview = book.clean(
            [
                UniqueRule("customers", "customer_id", "int"),
                ReferenceRule(
                    "orders", "customer_id", "customers", cleaner="int",
                ),
            ],
        )
        report = preview.error_sheet()
        assert isinstance(report, p.Sheet)
        assert report.name == "cleaning report"
        assert report.colnames == REPORT_HEADERS
        array = report.to_array()
        assert array[0] == REPORT_HEADERS
        # every data row has the full column count
        for row in array[1:]:
            assert len(row) == len(REPORT_HEADERS)

    def test_all_errors_are_collected(self):
        book = make_book(
            {
                "customers": [["customer_id"], [1], [1]],
                "orders": [["customer_id"], [8], [9]],
            },
        )
        preview = book.clean(
            [
                UniqueRule("customers", "customer_id", "int"),
                ReferenceRule(
                    "orders", "customer_id", "customers", cleaner="int",
                ),
            ],
        )
        # 2 duplicate rows + 2 broken references = 4 errors
        assert len(preview.errors) == 4
        assert len(preview.error_sheet().to_array()) == 5

    def test_rule_ids_are_generated_or_custom(self):
        book = make_book({"items": [["code"], ["1"], ["1"]]})
        preview = book.clean(
            [
                UniqueRule("items", "code", "int", name="my-id"),
                ColumnCleaningRule("items", "code", "strip"),
            ],
        )
        ids = {row[1] for row in preview.error_sheet().to_array()[1:]}
        assert "my-id" in ids


class TestValueSemantics(TestCase):
    def test_none_empty_false_zero_not_conflated_in_cells(self):
        sheet = p.Sheet([[None, "", False, 0, 0.0]], name="t")
        book = p.Book({"t": sheet})
        preview = book.clean(
            [ColumnCleaningRule("t", [0, 1, 2, 3, 4], "strip")],
        )
        assert preview.passed
        preview.apply()
        row = book["t"].row[0]
        # None and "" both normalise to "" on load, but False and 0 survive
        # as their genuine types; strip never coerces them
        assert row[2] is False
        assert row[3] == 0 and isinstance(row[3], int)
        assert row[4] == 0.0 and isinstance(row[4], float)

    def test_false_is_not_treated_as_zero_by_strict_int(self):
        sheet = p.Sheet([[False, 0, "0"]], name="t")
        book = p.Book({"t": sheet})
        preview = book.clean([UniqueRule("t", [0], "int")])
        rows = preview.error_sheet().to_array()[1:]
        assert any(row[0] == BLOCKING for row in rows)

    def test_empty_string_is_a_real_value_outside_key_columns(self):
        sheet = p.Sheet([["", "  "]], name="t")
        book = p.Book({"t": sheet})
        preview = book.clean(
            [ColumnCleaningRule("t", [0, 1], "strip")],
        )
        preview.apply()
        row = book["t"].row[0]
        assert row[0] == ""
        assert row[1] == ""


class TestRuleDeclaration(TestCase):
    def setUp(self):
        self.book = make_book(
            {
                "customers": [["customer_id"], [1]],
                "orders": [["customer_id"], [1]],
            },
        )

    def test_unknown_sheet_rejected(self):
        with self.assertRaises(KeyError):
            self.book.clean(
                [ColumnCleaningRule("ghost", "customer_id", "strip")],
            )

    def test_unknown_reference_sheet_rejected(self):
        with self.assertRaises(KeyError):
            self.book.clean(
                [ReferenceRule("orders", "customer_id", "ghost")],
            )

    def test_unknown_column_rejected(self):
        with self.assertRaises(KeyError):
            self.book.clean(
                [UniqueRule("customers", "missing_col", "int")],
            )

    def test_unknown_cleaner_rejected(self):
        with self.assertRaises(ValueError):
            ColumnCleaningRule("customers", "customer_id", "bogus")

    def test_empty_rules_rejected(self):
        with self.assertRaises(ValueError):
            run_cleaning(self.book, [])

    def test_integer_column_indices_supported(self):
        # unnamed sheet: use raw column indices
        book = p.get_book(
            bookdict={"items": [[" x"], ["x"], ["y"]]},
        )
        preview = book.clean([UniqueRule("items", 0, "strip")])
        assert preview.passed is False
        rows = preview.error_sheet().to_array()[1:]
        assert len(rows) == 2

    def test_default_cleaner_is_strip(self):
        book = make_book({"items": [["code"], [" a"], ["a "], ["b"]]})
        preview = book.clean([UniqueRule("items", "code")])
        assert preview.passed is False


class TestEndToEndExport(TestCase):
    def test_apply_then_export_round_trip(self):
        book = make_book(
            {
                "customers": [["customer_id"], [" 1001"], [1002]],
                "orders": [
                    ["order_id", "customer_id"],
                    [1, "1001 "],
                    [2, 1002],
                ],
            },
        )
        book.clean(
            [
                ColumnCleaningRule("customers", "customer_id", "strip"),
                ColumnCleaningRule("orders", "customer_id", "strip"),
                UniqueRule("customers", "customer_id", "int"),
                ReferenceRule(
                    "orders", "customer_id", "customers", cleaner="int",
                ),
            ],
        ).apply()

        stream = book.save_to_memory("xls")
        reloaded = p.get_book(
            file_type="xls",
            file_content=stream.getvalue(),
        )
        # every sheet re-exported, order intact, headers preserved
        assert reloaded.sheet_names() == ["customers", "orders"]
        assert reloaded["customers"].cell_value(0, 0) == "customer_id"
        assert reloaded["customers"].cell_value(1, 0) == 1001
        assert isinstance(reloaded["customers"].cell_value(1, 0), int)
        assert reloaded["orders"].cell_value(1, 1) == 1001

    def test_legacy_book_format_still_works(self):
        book = make_book(
            {"items": [["code", "v"], ["1", 1.5], ["2", 2.5]]},
        )
        # legacy whole-sheet format keeps its own (lenient) semantics
        sheet = book["items"]
        sheet.format(str)
        assert sheet.cell_value(0, 1) == "1.5"

    def test_preview_object_api(self):
        book = make_book({"items": [["code"], ["1"], ["1"]]})
        preview = book.clean([UniqueRule("items", "code", "int")])
        assert isinstance(preview, CleaningPreview)
        assert preview.passed is False
        assert preview.book is book
        assert len(preview.blocking_errors) == 0
        assert preview.errors

    def test_overlapping_rules_do_not_duplicate_writes(self):
        book = make_book(
            {
                "customers": [["customer_id"], [" 1001"], [1002]],
                "orders": [["customer_id"], [1001], [1002]],
            },
        )
        rules = [
            ColumnCleaningRule("customers", "customer_id", "strip"),
            UniqueRule("customers", "customer_id", "int"),
            ColumnCleaningRule("orders", "customer_id", "strip"),
            ReferenceRule(
                "orders", "customer_id", "customers", cleaner="int",
            ),
        ]
        preview = book.clean(rules)
        assert preview.passed
        # apply() collapses repeated schedules of the same cell to one
        # final write, in rule order
        written = []
        original_writer = p.Sheet.cell_value

        def spy(self, row, column, new_value=None):
            if new_value is not None:
                written.append((self.name, row, column))
            return original_writer(self, row, column, new_value)

        p.Sheet.cell_value = spy
        try:
            preview.apply()
        finally:
            p.Sheet.cell_value = original_writer
        assert len(written) == len(set(written))
        assert book["customers"].cell_value(0, 0) == 1001
        assert isinstance(book["customers"].cell_value(0, 0), int)
        assert book["orders"].cell_value(0, 0) == 1001
