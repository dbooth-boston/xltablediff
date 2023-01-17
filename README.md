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

The `--out OUTFILE.xlsx` "option" is required, and tells xltablediff to write the resulting comparison to OUTFILE.xslx .  

The resulting output file highlights differences found between the
old and new tables.  The first column in the output file
contains a marker indicating whether the row changed:
```
    -   Row deleted
    +   Row added
    =   Row unchanged (excluding columns added or deleted)
    c-  Row changed: this row shows the old content
    c+  Row changed: this row shows the new content
```

Deleted columns, rows or cells are highlighted with light red; and added
columns, rows or cells are highlighted with light green.
If a row in the table contains values that changed (from old to new),
that row will be highlighted in light yellow and repeated:
the first of the resulting two rows will show the old values, and the second row will show the new values.
The table's header row and key column are otherwise highlighted in gray-blue.

## Limitations

 - xltablediff compares only one old versus one new table, even if the given Excel files contain multiple sheets and/or tables.  If you want to compare multiple tables in different sheets, you will have to run xltablediff multiple times to compare them all, using the `--sheet`, `--oldSheet` and/or `--newSheet` options to specify which tables to compare.
 - If a cell somehow contains any tabs they will be silently converted to spaces prior to comparison.
 - Cell formatting might not be retained.

## Testing
Some rudimenary test files are included.  You can use them to try out the software, such as: 
```
./xltablediff.py --newSheet=Sheet2 --key=ID test1in.xlsx test1in.xlsx --out=test1out.xlsx
```

## Usage and options
```

usage: xltablediff.py [-h] [--key KEY] [--oldSheet OLDSHEET]
                      [--newSheet NEWSHEET] [--sheet SHEET] [--ignore IGNORE]
                      [--oldAppend] [--newAppend] [--merge MERGE] [--mergeAll]
                      [--maxColumns MAXCOLUMNS] --out OUT
                      oldFile.xlsx newFile.xlsx

Compare tables in two .xlsx spreadsheets

positional arguments:
  oldFile.xlsx          Old spreadsheet (*.xlsx)
  newFile.xlsx          New spreadsheet (*.xlsx)

optional arguments:
  -h, --help            show this help message and exit
  --key KEY             Specifies KEY as the name of the key column, i.e., its
                        header. If KEY is of the form "OLDKEY=NEWKEY" then
                        OLDKEY and NEWKEY are the corresponding key columns of
                        the old and new tables, respectively.
  --oldSheet OLDSHEET   Specifies the sheet in oldFile to be compared.
                        Default: the first sheet with a table with a KEY
                        column.
  --newSheet NEWSHEET   Specifies the sheet in newFile to be compared.
                        Default: the first sheet with a table with a KEY
                        column.
  --sheet SHEET         Specifies the sheet to be compared, in both oldFile
                        and newFile. Default: the first sheet with a table
                        with a KEY column.
  --ignore IGNORE       Ignore the specified column when comparing old and new
                        table rows. This option may be repeated to ignore
                        multiple columns. The specified column must exist in
                        both old and new tables.
  --oldAppend           Copy the values of oldFile sheet, appending columns of
                        newFile, keeping only the rows of oldFile. Rows of
                        newFile that do not exist in oldFile are discarded.
                        Leading and trailing rows of newFile (before and after
                        the table) are also discarded. The number of rows in
                        the output file will be the same is in oldFile.
  --newAppend           Copy the values of oldFile sheet, appending columns of
                        newFile, but forcing the resulting table body to have
                        the same rows as in newFile, based on the keys in
                        newFile. Rows of newFile that do not exist in oldFile
                        are inserted (with the key value from newFile), and
                        rows of oldFile that do not exist in newFile are
                        discarded. Leading and trailing rows of newFile
                        (before and after the table) are also discarded. The
                        number of rows in the resulting table body will be the
                        same is in newFile.
  --merge MERGE         Output a copy of the old sheet, with values from the
                        specifed MERGE column from the new table merged in.
                        This option may be repeated to merge more than one
                        column. The number of rows in the output file will be
                        the same is in oldFile.
  --mergeAll            Same as '--merge C' for all non-key columns C that are
                        in both the old and new tables.
  --maxColumns MAXCOLUMNS
                        Delete all columns after column N, where N is an
                        integer (origin 1)
  --out OUT             Output file of differences. This "option" is actually
                        REQUIRED.
```

## Examples

```
# Diff test:
  xltablediff.py  --key ID test1old.xlsx test1new.xlsx --out test1diff.xlsx

# Ignore test:
  xltablediff.py  --key ID --ignore Color test1old.xlsx test1new.xlsx --out test1ignore.xlsx

# Merge test:
  xltablediff.py  --key ID --merge Color test1old.xlsx test1new.xlsx --out test1merge.xlsx

# oldAppend test:
  xltablediff.py  --key ID --oldAppend test1old.xlsx test1new.xlsx --out test1oldAppend.xlsx

# newAppend test:
  xltablediff.py  --key ID --newAppend test1old.xlsx test1new.xlsx --out test1newAppend.xlsx
```
  
