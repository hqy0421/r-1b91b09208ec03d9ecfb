"""
pyexcel.cleaning
~~~~~~~~~~~~~~~~

Whole book cleaning and validation.

Data quality problems often only surface after cell values are normalised:
two identifiers such as ``" A01"`` and ``"A01"`` look different until the
surrounding spaces are removed, and ``1`` (a number) and ``"1"`` (text)
only collide once the text is converted.  This module cleans the declared
key columns of every sheet first, and then runs uniqueness and cross
sheet reference checks against the *cleaned* composite keys.

The workflow is strictly two phased:

1. :meth:`~pyexcel.Book.clean` previews every rule on private working
   copies.  The original book is never touched, and every error found is
   collected into a native :class:`~pyexcel.Sheet`.
2. :meth:`CleaningPreview.apply` writes the cleaned values back to the
   original book in one go, and only when the whole book passed.  A book
   can never end up half cleaned.

:copyright: (c) 2014-2026 by C Wang
:license: New BSD License, see LICENSE for more details
"""

import copy
import datetime
from decimal import Decimal

from pyexcel import constants
from pyexcel._compact import OrderedDict
from pyexcel.sheet import Sheet

#: rule executed successfully, no problem found
PASSED = "passed"
#: an executable check found a data violation
ERROR = "error"
#: the rule could not be judged because a value could not be cleaned
BLOCKING = "blocking"

#: report column name for the rule severity
SEVERITY = "severity"
RULE_ID = "rule id"
RULE_TYPE = "rule type"
SHEET_NAME = "sheet"
ROW_INDEX = "row"
COLUMN = "column"
ORIGINAL_VALUE = "original value"
MESSAGE = "message"

REPORT_HEADERS = [
    SEVERITY,
    RULE_ID,
    RULE_TYPE,
    SHEET_NAME,
    ROW_INDEX,
    COLUMN,
    ORIGINAL_VALUE,
    MESSAGE,
]

REPORT_SHEET_NAME = "cleaning report"

# built in cleaner names, accepted alongside the callables themselves
STRIP = "strip"
TO_STR = "str"
TO_INT = "int"
TO_FLOAT = "float"


class ConversionError(ValueError):
    """Raised when a cell value cannot be converted as declared.

    A conversion failure is different from a data violation: the rule
    depending on the value cannot be judged at all, so it is reported as
    blocking instead of producing a potentially wrong error.
    """


def clean_strip(value):
    """Trim surrounding whitespace from strings.

    ``None``, booleans, numbers and every other non string value are
    returned untouched so that ``None``, ``""``, ``False`` and numbers are
    never conflated.
    """
    if isinstance(value, str):
        return value.strip()
    return value


def clean_str(value):
    """Strict conversion to :class:`str`.

    ``None`` stays ``None`` (it is not an empty string) while every other
    value is rendered with :func:`str`.
    """
    if value is None:
        return None
    if isinstance(value, str):
        return value
    if isinstance(value, bool):
        return "true" if value else "false"
    if isinstance(value, datetime.datetime):
        return value.strftime("%d/%m/%y")
    if isinstance(value, datetime.date):
        return value.strftime("%d/%m/%y")
    if isinstance(value, datetime.time):
        return value.strftime("%H:%M:%S")
    if isinstance(value, datetime.timedelta):
        return str(value)
    return str(value)


def clean_int(value):
    """Strict conversion to :class:`int`.

    Unlike the legacy :func:`~pyexcel.internal.sheets.formatters.to_format`
    helper, a value that cannot be converted raises :class:`ConversionError`
    instead of being silently returned, and booleans are not accepted as
    integers so that ``False`` and ``0`` remain distinguishable.
    """
    if value is None:
        return None
    if isinstance(value, bool):
        raise ConversionError(f"boolean {value!r} cannot be converted to int")
    if isinstance(value, int):
        return value
    if isinstance(value, float):
        if value.is_integer():
            return int(value)
        raise ConversionError(f"float {value!r} is not an integral number")
    if isinstance(value, Decimal):
        if value == value.to_integral_value():
            return int(value)
        raise ConversionError(f"decimal {value!r} is not an integral number")
    if isinstance(value, str):
        text = value.strip()
        if text == "":
            raise ConversionError("empty string cannot be converted to int")
        try:
            return int(text)
        except ValueError:
            try:
                floating = float(text)
            except ValueError:
                raise ConversionError(
                    f"{value!r} cannot be converted to int",
                )
            if floating.is_integer():
                return int(floating)
            raise ConversionError(f"{value!r} is not an integral number")
    raise ConversionError(f"{value!r} cannot be converted to int")


def clean_float(value):
    """Strict conversion to :class:`float`.

    Booleans are rejected so that ``False`` and numeric values stay
    distinguishable, and unparsable values raise :class:`ConversionError`.
    """
    if value is None:
        return None
    if isinstance(value, bool):
        raise ConversionError(
            f"boolean {value!r} cannot be converted to float",
        )
    if isinstance(value, (int, float, Decimal)):
        return float(value)
    if isinstance(value, str):
        text = value.strip()
        if text == "":
            raise ConversionError(
                "empty string cannot be converted to float",
            )
        try:
            return float(text)
        except ValueError:
            raise ConversionError(f"{value!r} cannot be converted to float")
    raise ConversionError(f"{value!r} cannot be converted to float")


BUILT_IN_CLEANERS = OrderedDict(
    [
        (STRIP, clean_strip),
        (TO_STR, clean_str),
        (TO_INT, clean_int),
        (TO_FLOAT, clean_float),
    ],
)


def resolve_cleaner(cleaner):
    """Turn a cleaner declaration into a callable."""
    if cleaner is None:
        return clean_strip
    if callable(cleaner):
        return cleaner
    if isinstance(cleaner, str):
        if cleaner in BUILT_IN_CLEANERS:
            return BUILT_IN_CLEANERS[cleaner]
        raise ValueError(
            f"unknown cleaner {cleaner!r}; "
            f"available cleaners: {', '.join(BUILT_IN_CLEANERS)}",
        )
    raise TypeError(
        "cleaner must be a callable, one of "
        f"{list(BUILT_IN_CLEANERS)} or None, got {type(cleaner)}",
    )


def _normalise_columns(columns):
    """Accept a single column declaration or a sequence of columns."""
    if columns is None:
        raise ValueError("at least one column must be declared")
    if isinstance(columns, (list, tuple)):
        if len(columns) == 0:
            raise ValueError("at least one column must be declared")
        return list(columns)
    return [columns]


class CleaningRule:
    """Base class of all whole book cleaning rules."""

    rule_type = "rule"

    def __init__(self, sheet, columns=None, name=None):
        self.sheet = sheet
        self.columns = _normalise_columns(columns)
        self.name = name

    def display_name(self, index):
        """Return the user supplied name or a generated one."""
        if self.name:
            return self.name
        return f"{self.rule_type}-{index + 1}"


class ColumnCleaningRule(CleaningRule):
    """Clean declared columns before any key based check runs.

    :param str sheet: name of the sheet the rule applies to
    :param columns: one column or a list of columns.  A column is either a
                    column name (when the sheet names its columns with a
                    header row) or a zero based column index
    :param cleaner: ``"strip"`` (default), ``"str"``, ``"int"``,
                    ``"float"`` or any callable taking one value and
                    returning its cleaned value
    :param str name: optional rule id shown in the report
    """

    rule_type = "cleaning"

    def __init__(self, sheet, columns, cleaner=None, name=None):
        CleaningRule.__init__(self, sheet, columns, name)
        self.cleaner = resolve_cleaner(cleaner)


class UniqueRule(CleaningRule):
    """Assert that the cleaned composite key of a sheet stays unique.

    :param str sheet: name of the sheet the rule applies to
    :param columns: the column or list of columns forming the composite key
    :param cleaner: optional cleaner applied to the key columns before the
                    uniqueness check
    :param str name: optional rule id shown in the report
    """

    rule_type = "unique"

    def __init__(self, sheet, columns, cleaner=None, name=None):
        CleaningRule.__init__(self, sheet, columns, name)
        self.cleaner = resolve_cleaner(cleaner)


class ReferenceRule(CleaningRule):
    """Assert that cleaned foreign keys exist in the referenced key set.

    :param str sheet: the sheet holding the foreign key
    :param columns: foreign key column or composite foreign key columns
    :param str references: the sheet providing the referenced key
    :param referenced_columns: referenced key column or columns.  Defaults
                               to ``columns`` so same named columns line up
    :param cleaner: optional cleaner applied to both sides of the relation
    :param str name: optional rule id shown in the report

    Rows whose foreign key components are all absent (``None`` or empty
    strings, the latter being how pyexcel represents a missing cell) are
    considered to carry no relation and are skipped, so missing
    associations are never reported for a row that declares no relation
    at all.  A conversion failure, by contrast, is reported as blocking.
    """

    rule_type = "reference"

    def __init__(
        self,
        sheet,
        columns,
        references,
        referenced_columns=None,
        cleaner=None,
        name=None,
    ):
        CleaningRule.__init__(self, sheet, columns, name)
        self.references = references
        if referenced_columns is None:
            self.referenced_columns = list(self.columns)
        else:
            self.referenced_columns = _normalise_columns(referenced_columns)
        if len(self.referenced_columns) != len(self.columns):
            raise ValueError(
                "reference rule %s: %d foreign key column(s) reference %d "
                "column(s); the composite keys must have the same width"
                % (
                    self.name or self.rule_type,
                    len(self.columns),
                    len(self.referenced_columns),
                ),
            )
        self.cleaner = resolve_cleaner(cleaner)


class CleaningPreview:
    """Result of a whole book cleaning dry run.

    :ivar book: the book the preview was produced from (left untouched)
    :ivar errors: list of error dictionaries, one per report row
    :ivar changes: list of ``(sheet_name, row, column, value)`` tuples
                   ready to be replayed when :meth:`apply` is called
    """

    def __init__(self, book, errors, changes):
        self.book = book
        self.errors = errors
        self.changes = changes

    @property
    def passed(self):
        """``True`` when no error (blocking or violation) was found."""
        return len(self.errors) == 0

    @property
    def blocking_errors(self):
        """All errors whose severity is :data:`BLOCKING`."""
        return [error for error in self.errors if error[SEVERITY] == BLOCKING]

    def error_sheet(self):
        """Return every error as a native :class:`~pyexcel.Sheet`.

        The sheet carries one header row followed by one row per error.
        An empty (header only) sheet is returned when the book passed so
        callers always receive the same structure.
        """
        array = [list(REPORT_HEADERS)]
        for error in self.errors:
            array.append([error[header] for header in REPORT_HEADERS])
        return Sheet(array, REPORT_SHEET_NAME, name_columns_by_row=0)

    def apply(self):
        """Write all cleaned values back to the original book.

        :raises CleaningError: when the preview did not pass.  Refusing to
                               apply guarantees that no sheet is ever
                               cleaned half way through.
        """
        if not self.passed:
            raise CleaningError(
                "cleaning cannot be applied: %d error(s) reported; "
                "inspect error_sheet() first" % len(self.errors),
            )
        # several rules may clean the same cell in sequence; keep only the
        # last scheduled value, which is the final value seen by every rule
        final_changes = OrderedDict()
        for sheet_name, row, column, value in self.changes:
            final_changes[(sheet_name, row, column)] = value
        for (sheet_name, row, column), value in final_changes.items():
            sheet = self.book.sheet_by_name(sheet_name)
            sheet.cell_value(row, column, value)
        return self.book


class CleaningError(Exception):
    """Raised when a failing preview is asked to apply its changes."""


class _SheetWork:
    """Private working state kept for one sheet during a dry run."""

    def __init__(self, sheet):
        self.sheet = sheet
        # deep copy so the real book is completely isolated
        self.array = copy.deepcopy(sheet.get_internal_array())
        self.colnames = list(sheet.colnames)
        # (rule id, column) -> set of rows already reported, so a single
        # broken cell is reported once even if several rules use it
        self._reported = set()

    def column_index(self, column):
        if isinstance(column, str):
            if column not in self.colnames:
                raise KeyError(
                    f"sheet {self.sheet.name!r} has no column named "
                    f"{column!r}; available columns: {self.colnames}",
                )
            return self.colnames.index(column)
        index = int(column)
        if index < 0 or index >= self.sheet.number_of_columns():
            raise KeyError(
                f"sheet {self.sheet.name!r} has no column index {column}",
            )
        return index

    def cell(self, row, column_index):
        return self.array[row][column_index]


def _make_error(severity, rule, work, row, column, value, message):
    return OrderedDict(
        [
            (SEVERITY, severity),
            (RULE_ID, rule.display_id),
            (RULE_TYPE, rule.rule_type),
            (SHEET_NAME, work.sheet.name),
            (ROW_INDEX, row),
            (COLUMN, column),
            (ORIGINAL_VALUE, value),
            (MESSAGE, message),
        ],
    )


def _clean_column(rule, work, column, errors, changes):
    """Run a cleaner over one column on the working copy."""
    column_index = work.column_index(column)
    cleaner = rule.cleaner
    nrows = work.sheet.number_of_rows()
    for row in range(nrows):
        original = work.cell(row, column_index)
        try:
            cleaned = cleaner(original)
        except Exception as exc:  # any failure blocks the dependent rules
            key = (rule.display_id, column, row)
            if key not in work._reported:
                work._reported.add(key)
                errors.append(
                    _make_error(
                        BLOCKING,
                        rule,
                        work,
                        row,
                        column,
                        original,
                        f"conversion failed: {exc}",
                    ),
                )
            continue
        if cleaned != original or type(cleaned) is not type(original):
            work.array[row][column_index] = cleaned
            changes.append(
                (work.sheet.name, row, column_index, cleaned),
            )


def _attach_ids(rules):
    for index, rule in enumerate(rules):
        rule.display_id = rule.display_name(index)


def _group_by_sheet(rules):
    grouped = OrderedDict()
    for rule in rules:
        grouped.setdefault(rule.sheet, []).append(rule)
    return grouped


def run_cleaning(book, rules):
    """Preview all cleaning and validation rules on a book.

    :param Book book: the book to validate.  It is not modified
    :param list rules: :class:`ColumnCleaningRule`, :class:`UniqueRule`
                       and :class:`ReferenceRule` instances
    :returns: a :class:`CleaningPreview`
    """
    if rules is None or len(rules) == 0:
        raise ValueError("at least one cleaning rule is required")
    _attach_ids(rules)

    sheet_names = book.sheet_names()
    for rule in rules:
        if rule.sheet not in sheet_names:
            raise KeyError(
                f"rule {rule.display_id!r} references unknown sheet "
                f"{rule.sheet!r}; available sheets: {sheet_names}",
            )
        if (
            isinstance(rule, ReferenceRule)
            and rule.references not in sheet_names
        ):
            raise KeyError(
                f"rule {rule.display_id!r} references unknown sheet "
                f"{rule.references!r}; available sheets: {sheet_names}",
            )

    works = OrderedDict(
        (name, _SheetWork(book.sheet_by_name(name))) for name in sheet_names
    )
    errors = []
    changes = []

    grouped = _group_by_sheet(rules)

    # phase 1: every cleaning rule, sheet by sheet in declaration order
    for sheet_name in sheet_names:
        sheet_rules = grouped.get(sheet_name, [])
        work = works[sheet_name]
        for rule in sheet_rules:
            if not isinstance(rule, ColumnCleaningRule):
                continue
            for column in rule.columns:
                _clean_column(rule, work, column, errors, changes)

    # phase 2: validation rules read the cleaned working copies
    for rule in rules:
        if isinstance(rule, UniqueRule):
            _run_unique(rule, works[rule.sheet], errors, changes)
        elif isinstance(rule, ReferenceRule):
            _run_reference(rule, works, errors, changes)

    return CleaningPreview(book, errors, changes)


def _clean_for_validation(
    rule, work, columns, errors, changes, skip_absent=False,
):
    """Clean key columns in place for a validation rule.

    Returns ``(column_indices, blocked_rows)``.  Conversion failures are
    recorded as blocking errors and the affected rows are excluded from
    the actual check instead of producing false data violations.

    When ``skip_absent`` is true, rows whose key components are all
    absent (``None`` or empty string) bypass the cleaner: they carry no
    relation and must not trigger conversion failures.
    """
    column_indices = [work.column_index(column) for column in columns]
    blocked_rows = set()
    nrows = work.sheet.number_of_rows()
    for column, column_index in zip(columns, column_indices):
        for row in range(nrows):
            original = work.cell(row, column_index)
            if skip_absent and _is_absent(original):
                continue
            try:
                cleaned = rule.cleaner(original)
            except Exception as exc:
                blocked_rows.add(row)
                key = (rule.display_id, column, row)
                if key not in work._reported:
                    work._reported.add(key)
                    errors.append(
                        _make_error(
                            BLOCKING,
                            rule,
                            work,
                            row,
                            column,
                            original,
                            f"conversion failed: {exc}",
                        ),
                    )
                continue
            if cleaned != original or type(cleaned) is not type(original):
                work.array[row][column_index] = cleaned
                changes.append(
                    (work.sheet.name, row, column_index, cleaned),
                )
    return column_indices, blocked_rows


def _composite_key(work, row, column_indices):
    return tuple(work.array[row][index] for index in column_indices)


def _is_absent(value):
    """A cell carries no key value.

    ``None`` means absent while an empty string is a real cell value in
    every other context, but pyexcel normalises missing cells to empty
    strings while loading, so both forms mean "no relation declared" when
    they occupy a key column on their own.  ``False`` and numbers are
    never considered absent.
    """
    return value is None or value == ""


def _all_none(key):
    return len(key) > 0 and all(_is_absent(component) for component in key)


def _run_unique(rule, work, errors, changes):
    column_indices, blocked_rows = _clean_for_validation(
        rule, work, rule.columns, errors, changes,
    )
    seen = {}
    nrows = work.sheet.number_of_rows()
    duplicates = set()
    for row in range(nrows):
        if row in blocked_rows:
            continue
        key = _composite_key(work, row, column_indices)
        if _all_none(key):
            # rows without a key do not participate in uniqueness
            continue
        if key in seen:
            duplicates.add(seen[key])
            duplicates.add(row)
        else:
            seen[key] = row
    for row in sorted(duplicates):
        key = _composite_key(work, row, column_indices)
        errors.append(
            _make_error(
                ERROR,
                rule,
                work,
                row,
                ",".join(str(column) for column in rule.columns),
                key if len(key) > 1 else key[0],
                "duplicate key %r" % (key,),
            ),
        )
    if blocked_rows:
        errors.append(
            _make_error(
                BLOCKING,
                rule,
                work,
                constants.DEFAULT_NA,
                ",".join(str(column) for column in rule.columns),
                constants.DEFAULT_NA,
                "uniqueness could not be fully judged because %d row(s) "
                "failed to convert" % len(blocked_rows),
            ),
        )


def _run_reference(rule, works, errors, changes):
    source_work = works[rule.sheet]
    target_work = works[rule.references]

    target_indices, target_blocked = _clean_for_validation(
        rule, target_work, rule.referenced_columns, errors, changes,
    )
    source_indices, source_blocked = _clean_for_validation(
        rule, source_work, rule.columns, errors, changes, skip_absent=True,
    )

    if target_blocked:
        errors.append(
            _make_error(
                BLOCKING,
                rule,
                target_work,
                constants.DEFAULT_NA,
                ",".join(str(column) for column in rule.referenced_columns),
                constants.DEFAULT_NA,
                "referenced key set could not be built because %d row(s) "
                "failed to convert; reference check not judged"
                % len(target_blocked),
            ),
        )
        # no trustworthy key set exists: do not report missing associations
        return

    referenced = set()
    for row in range(target_work.sheet.number_of_rows()):
        key = _composite_key(target_work, row, target_indices)
        if not _all_none(key):
            referenced.add(key)

    nrows = source_work.sheet.number_of_rows()
    for row in range(nrows):
        if row in source_blocked:
            # already reported as a blocking conversion failure above
            continue
        key = _composite_key(source_work, row, source_indices)
        if _all_none(key):
            # no relation declared on this row
            continue
        if key not in referenced:
            errors.append(
                _make_error(
                    ERROR,
                    rule,
                    source_work,
                    row,
                    ",".join(str(column) for column in rule.columns),
                    key if len(key) > 1 else key[0],
                    "broken reference to %r: key %r not found in sheet %r"
                    % (rule.referenced_columns, key, rule.references),
                ),
            )
    if source_blocked:
        errors.append(
            _make_error(
                BLOCKING,
                rule,
                source_work,
                constants.DEFAULT_NA,
                ",".join(str(column) for column in rule.columns),
                constants.DEFAULT_NA,
                "reference check could not be fully judged because %d "
                "row(s) failed to convert" % len(source_blocked),
            ),
        )
