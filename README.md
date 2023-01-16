# xltablediff: Compare two Excel spreadsheet tables

Yet another Excel (.xlsx) diff program, but it differs from other such programs in several important ways:
 - It compares two spreadsheet tables as though they are old and new versions of a **relational table**, with a shared key column.  The shared key column is used to determine what rows of the table were added, deleted or changed, regardless of their order in the table.
 - It also shows additions or deletions of entire columns, which helps avoid reporting spurious changes to individual rows.
 - Each table may be preceded or followed by rows that are not a part of the table.  
 - Cell formats and formulas are ignored in comparing cell content: only the values matter.
 - It can merge columns of one table into another table, based on the row keys -- see the `--merge` option.  Or it can add columns of one table to another table, based on the row keys -- see the `--appendOld` and `--appendNew` options.  These functions can be helpful when different people have been separately updating different copies of a spreadsheet table, and you want to merge their changes back into the master table.

Technically, a table is treated as a rectangular region of cells in a spreadsheet, beginning with a header row, which defines the names of the columns.  A current limitation is that the spreadsheet cannot have leading columns or non-empty trailing columns that are not a part of the table.  The location of the table -- i.e., the region of the sheet to be treated as a table -- is guessed by xltablediff, by looking for a row that looks like a header row.  (It would be nice to add an option to allow the specific region to be specified in an option, like `B12:f99`, but so far the author has not needed it.  Please submit a pull request if you are willing to contribute such a feature.)

The old and new tables must have a key column in common, which is specified by the `--key K` option, where K is the name of the key column.  By default it must be the same in the old and new tables, but if it differs you can specify the old and new keys like this: `--key OLDKEY=NEWKEY`.  The key
must uniquely identify the rows in the table, as in a relational table.  
Added or deleted columns
are detected by comparing the old and new column names, i.e., the column headers.

Rows before and after the table are compared differently than rows of the table: they are compared as lines, more like with a conventional diff.

## Output

The `--out OUTFILE.xlsx` "option" is required, and tells xltablediff to write the resulting comparison to OUTFILE.xslx .  In the output file, deleted columns, rows or cells are highlighted with light red; and added
columns, rows or cells are highlighted with light green.
If a row in the table contains values that changed (from old to new),
that row will be highlighted in light yellow and repeated:
the first of the resulting two rows will show the old values, and the second row will show the new values.
The table's header row and key column are otherwise highlighted in gray-blue.

## Limitations

 - xltablediff compares only one old versus one new table, even if the given Excel files contain multiple sheets and/or tables.  If you want to compare multiple tables in different sheets, you will have to run xltablediff multiple times to compare them all, using the `--sheet`, `--oldSheet` and/or `--newSheet` options to specify which tables to compare.
 - If a cell somehow contains any tabs they will be silently converted to spaces prior to comparison.

## Testing
Some rudimenary test files are included.  You can use them to try out the software, such as: 
```
./xltablediff.py --newSheet=Sheet2 --key=ID test1in.xlsx test1in.xlsx --out=test1out.xlsx
```

## Options
TODO
