#!/usr/bin/env python3.8

# Copyright 2022 by David Booth
# License: Apache 2.0
# Note that this code uses simplediff, which is used 
# under the zlib/libpng license
# <http://www.opensource.org/licenses/zlib-license.php>


# Compare two spreadsheet tables.
#
# Usage:
#   xldiff --sheet=MySheet1 --key=ID oldFile.xlsx newFile.xlsx
#
# Where:
#   --key=K     
#       Specifies K as the name of the key column that uniquely
#       identifies each row in the table, where K must be one
#       of the fields in the header row.  
#
#   --sheet=S
#       Specifies S as the name of the sheet containing the table
#       to be compared.  Leading and trailing whitespace are
#       stripped from sheet names prior to comparing with S.

# Strategy:
# 1. Ignore empty trailing rows and columns.
# 2. Diff the rows before the table only as rows.
# 3. Each row is uniquely identified by the keys in the key column;
# each column is uniquely identified by the headers in the header row.
# 4. Detect added or deleted rows or columns by comparing the old
# and new keys and headers.
# 5. The order of the keys and headers does not matter for diffing.
# It is only used for display purposes.
# 6. Cells are uniquely identify by header and key.
# 7. Detect changed cells by comparing them.
# 8. Use the newFile ordering of headers and keys for
# displaying the results.  If a row or column was deleted, show it
# after the row/column that preceded it in oldFile.  If no
# row/column preceded it in oldFile, then show it before the row/column
# that followed it in oldFile.
# 9. Optionally show unchanged rows and/or columns.
# 10. Highlight changes in yellow, additions in green, deletions in red.



import sys
import os
import os.path
import openpyxl
import re
import json
import keyword
import argparse
import simplediff

###################### TODO #######################
todoWarned = {}   
def TODO(msg):
    """Warn (only once) that a feature is not yet implemented.
    """
    if not msg in todoWarned:
        todoWarned[msg] = True
        if msg[-1] != "\n":
            msg += "\n"
        sys.stderr.write(f"TODO: {msg}")

################################ Globals ###################################
# Command line options:
optionHeader = ""   # Specifies the first column header indicating
                    # the table start.
optionKey = "id"    # Header of key column for the table.

keepLeadingSpaces = None

##################### Dumps #####################
def Dumps(obj): 
    """Format the given Object as a string.
    """
    if isinstance(obj, Object) : 
        return repr(obj)
    if isinstance(obj, list) : 
        return f'[{", ".join([repr(ob) for ob in obj])}]'
    if isinstance(obj, dict) : 
        return f'{{ {", ".join([repr(k) + ": " + repr(obj[k]) for k in obj.keys()])} }}'
    result = ""
    try:
        result = json.dumps(obj, indent=2)
    except TypeError as e:
        result = f'"({obj.__class__.__name__} OBJECT IS NOT JSON SERIALIZABLE)"'
    return result

##################### Object #####################
class Object: 
    """Generic object for attaching attributes.  IDK if this is the
    preferred way to do this in python.  Maybe I should be using
    a named tuple?   But a named tuple is not mutable.
    https://stackoverflow.com/questions/2970608/what-are-named-tuples-in-python
    """
    def __repr__(self):
        attrs = ",\n  ".join([ f'"{k}": {Dumps(getattr(self, k, "None"))}' 
            for k in dir(self) if not re.match("__", k) ])
        return f"{{ {attrs} }}\n"

##################### Unique #####################
def Unique(theList) :
    """Return a list of the unique items in theList, retaining order.
    """
    seen = set()
    u = []
    for item in theList :
        if item in seen :
            continue
        u.append(item)
        seen.add(item)
    return u

##################### Trim #####################
def Trim(s) :
    """Trim leading and trailing whitespace from the given string.
    """
    ss = TrimLeading(TrimTrailing(s))
    # print(f"Trim: [{s}] -> [{ss}]")
    return ss

##################### TrimLeading #####################
def TrimLeading(s) :
    """Trim leading whitespace from the given string.
    """
    ss = re.sub(r'\A[\s\n\r]+', '', s)
    return ss

##################### TrimTrailing #####################
def TrimTrailing(s) :
    """Trim trialing whitespace from the given string.
    """
    ss = re.sub(r'[\s\n\r]+\Z', '', s)
    return ss

##################### CleanSheet #####################
def CleanSheet(sheet, keepLeadingSpaces=None) :
    """Trim empty trailing rows and columns, and replace any
    empty cells with the empty string, for easier processing.
    Trailing whitespace is always trimmed; leading whitespace
    is trimmed unless keepLeadingSpaces is true.
    The given sheet is MODIFIED IN PLACE.
    Returns the number of rows found.
    """

    sys.stderr.write(f"CleanSheet '{sheet.title}' trimming whitespace...\n")
    # Change None to empty string, and trim whitespace.
    for row in sheet.rows :
        for cell in row:
            try:
                if cell.value is None :
                    cell.value = ""
                else :
                    if keepLeadingSpaces :
                        cell.value = TrimTrailing(str(cell.value))
                    else :
                        cell.value = Trim(str(cell.value))
            except AttributeError as e:  
               raise AttributeError(str(e) + f"\n at row {cell.row} column {cell.column}")

    rows = list(sheet.rows)
    columns = list(sheet.columns)
    nRows = len(rows)
    nColumns = len(columns)
    sys.stderr.write(f"Trimming empty rows and columns. nRows: {nRows} nColumns: {nColumns} ...\n")
    # Delete empty trailing rows:
    deleteFrom = nRows+1
    while 1 :
        if nRows <= 0 :
            break
        dataFound = 0
        row = rows[nRows-1]
        for cell in row :
            if cell.value != "" :
                dataFound = 1
                break
        if dataFound :
            break
        deleteFrom = nRows
        # print(f"  Deleted empty row {nRows}")
        nRows -= 1
    if deleteFrom < len(rows)+1 :
        sheet.delete_rows(deleteFrom)    # Numbered from 1
    # Delete empty trailing columns:
    deleteFrom = nColumns+1
    while 1 :
        if nColumns <= 0 :
            break
        dataFound = 0
        column = columns[nColumns-1]
        for cell in column :
            if cell.value != "" :
                # print(f"  column {nColumns} data found: {cell.value}")
                dataFound = 1
                break
        if dataFound :
            break
        deleteFrom = nColumns
        # print(f"  Deleted empty column {nColumns}")
        nColumns -= 1
    if deleteFrom < len(columns) :
        sheet.delete_cols(deleteFrom)    # Numbered from 1
    sys.stderr.write(f"Done trimming.  nRows: {nRows} nColumns: {nColumns}\n")
    return nRows    

##################### WriteTable #####################
def WriteTable(table, filename) :
    """Write an xslx table produced by TableParser or similar.
    The resulting xslx file will include the global properties
    from table, and the headers and content will begin after a 
    BEGINTABLE row.
    """
    # Make a new workbook and copy the table into it.
    outWb = openpyxl.Workbook()
    outSheet = outWb.active
    firstRow = table.rows[0].rowNum-1     # -1 because of 1-based index
    # Subtract 1 because the header row is before the first row:
    outHeaders = [ c.value for c in list(table.sheet.rows)[firstRow-1] ]
    emptyRow = ['' for _ in table.sheet.columns]
    # Global properties first:
    for p in table.gPropsList :
        k, v = p
        outRow = [ f"{k}: {v}" ]
        outRow.extend( emptyRow[1:] )
        outSheet.append(outRow)
    # Add empty rows to make the table begin on the same row as it did
    # originally:
    while outSheet.max_row + 2 < firstRow:
        outSheet.append(emptyRow)
    # Start the table
    outRow = [ f"BEGINTABLE" ]
    outRow.extend( emptyRow[1:] )
    outSheet.append(outRow)
    outSheet.append(outHeaders)
    outSheet.title = table.sheet.title
    # Dump all the rows
    for row in table.rows :
        outRow = []
        for h in outHeaders :
            outRow.append(getattr(row, h, ''))
        outSheet.append(outRow)
    outWb.save(filename)

################# findall_sub #####################
# Python3 function to perform string substitution while
# also returning group matches.  Returns a pair:
#     (newString, matches) 
# where newString is the new string, and matches is a list
# of all matches (potentially an empty list, if no matches).
def findall_sub(pattern, repl, string, count=0, flags=0):
        """ Call findall and sub, and return a pair containing
        both the new string (resulting from the substitution)
        and a list of group matches.
        """
        matches = re.findall(pattern, string, flags)
        newString = re.sub(pattern, repl, string, count, flags)
        return (newString, matches)

##################### CellValue #####################
def CellValue(cell):
    '''Return the text value of the given cell
    '''
    try:
        if cell.value is None :
            return ""
        else :
            return str(cell.value).strip()
    except AttributeError as e:
        sys.stderr.write(f"[WARNING] Ignoring cell at row {cell.row} column {cell.column}: {str(e)}\n")
    return ""

##################### TrimRowsAndColumns #####################
def TrimRowsAndColumns(rows):
    ''' Trim trailing empty rows and columns, and pad
    every row to have the same number of values.
    Modifies the give rows of values in place.
    '''
    iLastRow = -1
    jLastColumn = -1
    for i in range(len(rows)):
        row = rows[i]
        # Look for the last non-empty cell in the row:
        jLastThisRow = -1
        for j in range(len(row)-1, -1, -1):
            if row[j]:
                jLastThisRow = j
                if j > jLastColumn:
                    jLastColumn = j
                break
        if jLastThisRow >= 0:
            iLastRow = i
            if jLastThisRow > jLastColumn:
                jLastColumn = jLastThisRow
    nRows = iLastRow + 1
    nColumns = jLastColumn + 1
    del rows[nRows : ]
    # Now pad (or trim) each row to the max number of non-empty columns.
    for i in range(nRows):
        row = rows[i]
        if len(row) > nColumns:
            # Trim a row with extra cells:
            del row[ nColumns : ]
        # Pad a row with too few cells:
        padding = [ "" for j in range(len(row), nColumns) ]
        row.extend(padding)
    return rows

##################### GuessHeaderRow #####################
def GuessHeaderRow(rows, key, title, ignoreLastRow=True):
    ''' Guess the header row, as the first row with the most non-empty
    cells.  If ignoreLastRow (the default) then the last row is ignored
    because it should not normally be the header row.  If key 
    was specified, the header row also must contain the key.
    Returns:
        iHeader = the 0-based index of the header row, or None if not found.
    '''
    iHeader = None
    maxValues = 0
    sys.stderr.write(f"[INFO] Sheet '{title}' n rows: {len(rows)}\n")
    # r is 0-based index.
    # Ignore the last row, because it should not be the header row.
    nRows = len(rows)
    if ignoreLastRow:
        nRows = nRows - 1
    for r in range(nRows):
        # sys.stderr.write(f"[INFO] Processing sheet '{title}' row: {r+1}\n")
        row = rows[r]
        if key and key not in row:
            continue
        nValues = len( [ v for v in row if v ] )
        if nValues > maxValues:
            maxValues = nValues
            iHeader = r
    return iHeader

##################### FindTable #####################
def FindTable(file, sheetTitle, key):
    ''' Read the given xlsx file and find the desired table.
    Raises an exception if the table is not found.
    Returns a tuple:
        title = title of the sheet in which the table was found.
        rows = rows (values only) of the sheet in which the table was found.
        iHeader = 0-based index of the header row in rows.
    '''
    wb = openpyxl.load_workbook(file, data_only=True) 
    # Potentially look through all sheets for the one to compare.
    # If the --sheet option was specified, then only that one will be 
    # checked for the desired header.
    sheet = None
    iHeader = None
    rows = None
    title = ""
    for s in wb:
        if sheet:
            # Already found the right sheet.  No need to look at others.
            break
        if sheetTitle:
            if sheetTitle == s.title.strip(): 
                sheet = s
                sys.stderr.write(f"[INFO] Found sheet: '{s.title}'\n")
            else:
                sys.stderr.write(f"[INFO] Skipping unwanted sheet: '{s.title}'\n")
                continue
        # Get the rows of cells:
        rows = [ [str(v or "").strip() for v in col] for col in s.values ]
        TrimRowsAndColumns(rows)
        title = s.title
        # Look for the header row.
        iHeader = GuessHeaderRow(rows, key, s.title)
        if iHeader is not None:
            sheet = s
            sys.stderr.write(f"[INFO] In sheet '{s.title}' found header at row {iHeader+1}\n")
            break
    if not sheet:
        if sheetTitle:
            sys.stderr.write(f"[ERROR] Sheet not found: '{sheetTitle}'\n")
            sys.stderr.write(f" in file: '{file}'\n")
            sys.exit(1)
        if key:
            sys.stderr.write(f"[ERROR] Key not found: {key}\n")
            sys.stderr.write(f" in file: '{file}'\n")
            sys.exit(1)
        sys.stderr.write(f"[ERROR] Unable to find header row\n")
        sys.stderr.write(f" in file: '{file}'\n")
        sys.exit(1)
    return (title, rows, iHeader)

##################### CompareTables #####################
def CompareTables(oldRows, iOldHeader, newRows, iNewHeader):
    ''' Compare the old and new tables.  Return a diffs workbook.
    '''
    ###### First compare the lines before the start of the table.
    sys.stderr.write(f"CompareTable\n")
    oldLeadingRows = oldRows[0 : iOldHeader]
    oldLeadingLines = [ "\t".join( [c.replace("\t", " ") for c in r] ) for r in oldLeadingRows ]
    newLeadingRows = newRows[0 : iNewHeader]
    newLeadingLines = [ "\t".join( [c.replace("\t", " ") for c in r] ) for r in newLeadingRows ]
    oldAll = "\n".join(oldLeadingLines)
    sys.stderr.write(f"old all:\n{oldAll}\n")
    newAll = "\n".join(newLeadingLines)
    sys.stderr.write(f"new all:\n{newAll}\n")
    rawDiffs = simplediff.diff(oldLeadingLines, newLeadingLines)
    # flattened = [val for sublist in list_of_lists for val in sublist]
    diffs = [ (d[0], v) for d in rawDiffs for v in d[1] ]
    sys.stderr.write(f"diffs:\n{repr(diffs)}\n")
    diffLeadingLines = [ d[0] + "\t" + d[1] for d in diffs ]
    diffAll = "\n".join(diffLeadingLines)
    sys.stderr.write(f"diff all:\n{diffAll}\n")
    ###### Compare the headers, i.e., the column names.
    oldHeaders = oldRows[iOldHeader]
    newHeaders = newRows[iNewHeader]
    rawDiffHeaders = simplediff.diff(oldHeaders, newHeaders)
    # flattened = [val for sublist in list_of_lists for val in sublist]
    diffHeaders = [ (d[0], v) for d in rawDiffHeaders for v in d[1] ]
    sys.stderr.write(f"diff headers:\n{diffHeaders}\n")

    # *** STOPPED HERE ***
    sys.exit(0)

##################### main #####################
def main():
    # Parse command line options:
    argParser = argparse.ArgumentParser(description='Parse AATS data dictionary artifacts')
    argParser.add_argument('--oldSheet',
                    help='Specifies the sheet in oldFile to be compared.\n If no sheet is specified, the first sheet containing the\n specified header will be used.')
    argParser.add_argument('--newSheet',
                    help='Specifies the sheet in newFile to be compared.\n If no sheet is specified, the first sheet containing the\n specified header will be used.')
    argParser.add_argument('--sheet',
                    help='Specifies the sheet to be compared, in both oldFile and newFile.\n If no sheet is specified, the first sheet containing the\n specified header will be used.')
    argParser.add_argument('--header',
                    help='Specifies the header H of the first column.\nA row offset (from H) may optionally be included as: --header=H+N .\n If no header is specified, the first row will be used as the header row,\n or the 0-based Nth row, if N is specified as: --Header=+N .')
    argParser.add_argument('--key',
                    help='Specifies the name of the key column, i.e., its header')
    argParser.add_argument('--out',
                    help='Output file of differences')
    argParser.add_argument('oldFile', metavar='oldFile.xlsx', type=str,
                    help='Old spreadsheet (*.xlsx)')
    argParser.add_argument('newFile', metavar='newFile.xlsx', type=str,
                    help='New spreadsheet (*.xlsx)')
    # sys.stderr.write(f"[INFO] calling print_using....\n")
    global args
    (args, otherArgs) = argParser.parse_known_args()
    if args.sheet and args.oldSheet:
        raise ValueError(f"[ERROR] Illegal combination of options: --sheet with --oldSheet")
    if args.sheet and args.newSheet:
        raise ValueError(f"[ERROR] Illegal combination of options: --sheet with --newSheet")
    oldSheetTitle = args.sheet
    newSheetTitle = args.sheet
    if args.oldSheet:
        oldSheetTitle = args.oldSheet
    if args.newSheet:
        newSheetTitle = args.newSheet
    # sys.stderr.write("args: \n" + json.dumps(args, indent=2))
    sys.stderr.write("args: \n" + repr(args) + "\n\n")
    sys.stderr.write("otherArgs: \n" + repr(otherArgs) + "\n\n")
    # These will be rows of values-only:
    (oldTitle, oldRows, oldIHeader) = FindTable(args.oldFile, oldSheetTitle, args.key)
    (newTitle, newRows, newIHeader) = FindTable(args.newFile, newSheetTitle, args.key)
    sys.stderr.write(f"[INFO] Old file: '{args.oldFile}' sheet: '{oldTitle}' row: {oldIHeader+1}\n")
    sys.stderr.write(f"[INFO] New file: '{args.newFile}' sheet: '{newTitle}' row: {newIHeader+1}\n")

    diffsRows = CompareTables(oldRows, oldIHeader, newRows, newIHeader)
    sys.exit(0)
        # *** STOPPED HERE ***

    sys.stderr.write(("=" * 72) + "\n")
    sys.stderr.write(f"File: {args.oldFile}\n")
    sys.stderr.write(f"Sheet:{s.title}\n")

    try:
        if not CleanSheet(sheet):
            sys.stderr.write(f"[INFO] Skipping empty sheet '{sheet.title}'\n")
    except AttributeError:
        raise
    pass

######################################################
if __name__ == '__main__':
    main()
    exit(0)

