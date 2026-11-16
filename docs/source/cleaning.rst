Whole book cleaning and validation
==================================

.. testcode::
   :hide:

   >>> import pyexcel as pe
   >>> from pyexcel import (
   ...     ColumnCleaningRule,
   ...     UniqueRule,
   ...     ReferenceRule,
   ... )

Data quality problems often only appear after cell values are normalised.
Two identifiers such as ``" A01"`` and ``"A01"`` look different until the
surrounding spaces are removed, and ``1`` (a number) and ``"1"`` (text) only
collide once the text is converted.  pyexcel therefore cleans the declared
key columns first, and checks uniqueness and cross sheet references against
the *cleaned* composite keys.

The workflow is strictly two phased:

#. :meth:`pyexcel.Book.clean` runs every rule on private working copies and
   returns a :class:`pyexcel.CleaningPreview`.  The original book is not
   touched.
#. :meth:`pyexcel.CleaningPreview.apply` writes the cleaned values back in
   one transaction, and only when the whole book passed.  A book can never
   end up half cleaned.

Declaring the rules
-------------------

Suppose a customers sheet holds an ``id`` column with stray spaces and mixed
text/number values, while an orders sheet references those ids::

   >>> customers = [
   ...     ["id", "name"],
   ...     [" 1001", "Alice"],
   ...     [1002, "Bob"],
   ...     [1003, "Carol"],
   ... ]
   >>> orders = [
   ...     ["order_id", "customer_id", "amount"],
   ...     [1, " 1001 ", 10],
   ...     [2, 1002, 20],
   ...     [3, "", 40],
   ... ]
   >>> book = pe.get_book(
   ...     bookdict={"customers": customers, "orders": orders},
   ... )
   >>> book["customers"].name_columns_by_row(0)
   >>> book["orders"].name_columns_by_row(0)

Three rule types are available:

* :class:`pyexcel.ColumnCleaningRule` -- clean declared columns before any
  key based check, using ``"strip"`` (the default), ``"str"``, ``"int"``,
  ``"float"`` or any callable.
* :class:`pyexcel.UniqueRule` -- assert the cleaned composite key stays
  unique within a sheet.
* :class:`pyexcel.ReferenceRule` -- assert every cleaned foreign key exists
  in the referenced key set.

::

   >>> rules = [
   ...     ColumnCleaningRule("customers", "id", "strip"),
   ...     UniqueRule("customers", "id", "int", name="customer pk"),
   ...     ColumnCleaningRule("orders", "customer_id", "strip"),
   ...     ReferenceRule(
   ...         "orders",
   ...         "customer_id",
   ...         "customers",
   ...         referenced_columns="id",
   ...         cleaner="int",
   ...     ),
   ... ]
   >>> preview = book.clean(rules)
   >>> preview.passed
   True

Preview first, apply once
-------------------------

While the preview is open the original cells keep their raw values::

   >>> book["orders"].cell_value(0, 1)
   ' 1001 '

After :meth:`~pyexcel.CleaningPreview.apply`, the sheet and row/column
accessors you already hold all see the final values, while sheet order,
column names and unselected fields are preserved::

   >>> held_sheet = book["orders"]
   >>> preview.apply() is book
   True
   >>> held_sheet.cell_value(0, 1)
   1001
   >>> held_sheet.column["customer_id"]
   [1001, 1002, '']
   >>> book.sheet_names()
   ['customers', 'orders']
   >>> book["customers"].column["name"]
   ['Alice', 'Bob', 'Carol']

Rows whose foreign key components are all absent (``None`` or empty
strings) carry no relation and are skipped, so they are never reported as
missing associations.

Reading the error report
------------------------

Every executable error is collected into a native
:class:`pyexcel.Sheet` by :meth:`pyexcel.CleaningPreview.error_sheet`.  A
duplicate that only emerges after cleaning is a data violation with the
``error`` severity::

   >>> dirty = pe.get_book(
   ...     bookdict={"items": [["code"], [" A"], ["A"], ["B"]]},
   ... )
   >>> dirty["items"].name_columns_by_row(0)
   >>> preview = dirty.clean([UniqueRule("items", "code", "strip")])
   >>> preview.passed
   False
   >>> for row in preview.error_sheet().to_array():
   ...     print(row)
   ['severity', 'rule id', 'rule type', 'sheet', 'row', 'column', 'original value', 'message']
   ['error', 'unique-1', 'unique', 'items', 0, 'code', 'A', "duplicate key ('A',)"]
   ['error', 'unique-1', 'unique', 'items', 1, 'code', 'A', "duplicate key ('A',)"]

A value that cannot be converted is different: the rule depending on it
cannot be judged, so it is reported as ``blocking`` instead of producing a
potentially wrong error.  When the referenced key set itself cannot be
built, no broken-reference errors are emitted at all::

   >>> broken = pe.get_book(
   ...     bookdict={
   ...         "customers": [["id"], ["bad"]],
   ...         "orders": [["customer_id"], [9999]],
   ...     },
   ... )
   >>> broken["customers"].name_columns_by_row(0)
   >>> broken["orders"].name_columns_by_row(0)
   >>> preview = broken.clean(
   ...     [ReferenceRule(
   ...         "orders",
   ...         "customer_id",
   ...         "customers",
   ...         referenced_columns="id",
   ...         cleaner="int",
   ...     )],
   ... )
   >>> for row in preview.error_sheet().to_array()[1:]:
   ...     print(row[0], "-", row[-1])
   blocking - conversion failed: 'bad' cannot be converted to int
   blocking - referenced key set could not be built because 1 row(s) failed to convert; reference check not judged

Applying a failing preview is refused, guaranteeing that no sheet is ever
cleaned half way through::

   >>> preview.apply()
   Traceback (most recent call last):
   ...
   pyexcel.cleaning.CleaningError: cleaning cannot be applied: 2 error(s) reported; inspect error_sheet() first

Value semantics
---------------

The strict cleaners deliberately keep values distinguishable:

* :func:`pyexcel.cleaning.clean_strip` only trims strings; ``None``,
  booleans and numbers pass through untouched.
* ``False`` is not accepted as ``0`` and booleans never silently become
  numbers.
* ``None`` and empty strings are not turned into zeros or placeholder
  text.

The legacy :meth:`pyexcel.Sheet.format` interface and all existing export
functions keep their original lenient semantics and are unaffected by
these rules.
