# xltablediff

Show value differences between two Excel (.xlsx) spreadsheet tables (old and new).
Each table may also have lines before
and after the table; the leading and trailing lines are also compared.

The old and new tables must both have a key column of
the same name, which is specified by the --key option.  The keys
must uniquely identify the rows in the table.  Added or deleted rows
are detected by comparing the keys in the rows-- not by row order.
Similarly, added or deleted columns
are detected by comparing the old and new column names, i.e., the headers.

Only cell values are compared -- not cell formatting or formulas --
and trailing empty rows or cells are ignored.  If a cell somehow contains
any tabs they will be silently converted to spaces prior to comparison.

Deleted columns, rows or cells are highlighted with light red; added
columns, rows or cells are highlighted with light green.
If a row in the table contains values that changed (from old to new),
that row will be highlighted  in light yellow and repeated:
the first of the resulting two rows will show the old values;
the second row will show the new values.
The table's header row and key column are otherwise highlighted in gray-blue.

Limitations:
1. Only one table in one sheet is compared with one table in one other sheet.

Test:
```
./xltablediff.py --newSheet=Sheet2 --key=ID test1in.xlsx test1in.xlsx --out=test1out.xlsx
```
