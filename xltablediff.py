#!/usr/bin/env python3.8

# Copyright 2022 by David Booth
# License: Apache 2.0
# Repo: https://github.com/dbooth-boston/xltablediff
# This code also uses simplediff, which is licensed 
# under the zlib/libpng license:
# <http://www.opensource.org/licenses/zlib-license.php>

# Show value differences between two spreadsheet tables (old and new).
# The tables may also have lines before
# and after the tables; the leading and trailing lines are also compared.
#
# The old and new tables must both have a key column of
# the same name, which is specified by the --key option.  The keys
# must uniquely identify the rows in the table.  Added or deleted rows
# are detected by comparing the keys in the rows-- not by row order.
# Similarly, added or deleted columns
# are detected by comparing the old and new column names, i.e., the headers.  
#
# Only cell values are compared -- not cell formatting or formulas --
# and trailing empty rows or cells are ignored.  If a cell somehow contains
# any tabs they will be silently converted to spaces prior to comparison.
# 
# Deleted columns, rows or cells are highlighted with light red; added
# columns, rows or cells are highlighted with light green.
# If a row in the table contains values that changed (from old to new),
# that row will be highlighted  in light yellow and repeated: 
# the first of the resulting two rows will show the old values;
# the second row will show the new values.
# The table's header row and key column are otherwise highlighted in gray-blue.
#
# Limitations:
#  1. Only one table in one sheet is compared with one table in one other 
#     sheet.  
#
# Test:
#   ./xltablediff.py --newSheet=Sheet2 --key=ID test1in.xlsx test1in.xlsx --out=test1out.xlsx

def Usage():
    return f'''Usage:
   xltablediff [ --sheet=MySheet1 ] [ --key=K ] oldFile.xlsx newFile.xlsx --out=outFile.xlsx

Where:
    --key=K     
        Specifies K as the name of the key column that uniquely
        identifies each row in the old and new tables.
        Key defaults to 'id'.

    --sheet=S
    --oldSheet=S
    --newSheet=S
        Specifies S as the name of the sheet containing the table
        to be compared.  The sheet will be guessed if it is not specified.
        "--sheet=S" is shorthand for "--oldSheet=S --newSheet=S".

    --out=outFile.xslx
        Specifies outFile.xlsx as the differences output file to write.  
        This "option" is actually REQUIRED.

    --help
        This help.

The oldFile and newFile tables to
be compared must begin with a header row containing the names of the
columns.  Headers -- i.e. column names -- must be unique.  They are
used to determine whether a column was deleted or added.  The order of
the columns does not affect the comparison of whether a column was
deleted or added, because each column is uniquely identified by its
header.

The key column must exist in both the oldFile and newFile headers.
It is used to uniquely identify each row,
to determine whether that row was deleted, added or changed.  The order of
rows in the file does not affect this comparison, because the row is
identified by its key, not by its position in the file.  

The resulting outFile highlights differences found between the
oldFile and newFile tables.  The first column in the outFile
contains a marker indicating whether the row changed:
    -   Row was deleted
    +   Row was added
    =   Row was not changed (excluding columns added or deleted)
    c-  Row was changed; this row shows the old content
    c+  Row was changed; this row shows the new content
'''

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
# 9. Highlight changes using cell fill colors.

# Ideas for future enhancements: 
# 1. Allow the table to specified (within
# the spreadsheet) by specifying a range, such as: --table=B10:G16
# 2. Optionally suppress unchanged rows and/or columns.


import sys
import os
import os.path
import openpyxl
from openpyxl.styles import PatternFill, Fill, Font
import re
import json
import keyword
import argparse
import simplediff

################################ Globals ###################################
DEFAULT_KEY = "id"    # Header of key column for the table.

##################### NoTabs #####################
def NoTabs(rows):
    ''' Change tabs to spaces in all cells.
    Modifies the given rows of values in place.
    '''
    for i in range(len(rows)):
        row = rows[i]
        for j in range(len(row)):
            row[j] = row[j].replace("\t", " ")
    return rows

##################### TrimAndPad #####################
def TrimAndPad(rows):
    ''' Trim trailing empty rows and columns, and pad
    every row to have the same number of values.
    Modifies the given rows of values in place.
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
def GuessHeaderRow(rows, key, title):
    ''' Guess the header row, as the first row with the most non-empty
    cells, which much be unique within that row.  
    If key was specified, the header row also must contain the key.
    Returns:
        iHeaders = the 0-based index of the header row, or None if not found.
    '''
    iHeaders = None
    maxValues = 0
    # sys.stderr.write(f"[INFO] Sheet '{title}' n rows: {len(rows)}\n")
    # r is 0-based index.
    for r in range(len(rows)):
        # sys.stderr.write(f"[INFO] Processing sheet '{title}' row: {r+1}\n")
        row = rows[r]
        if key and key not in row:
            continue
        values = [ v for v in row if v ]
        nValues = len(values)
        nSetItems = len(set(values))
        if nSetItems != nValues:
            # Cannot be a header row, because the values are not unique.
            continue
        if nValues > maxValues:
            maxValues = nValues
            iHeaders = r
    return iHeaders

##################### FindTable #####################
def FindTable(file, sheetTitle, key):
    ''' Read the given xlsx file and find the desired table.
    Raises an exception if the table is not found.
    Returns a tuple:
        title = title of the sheet in which the table was found.
        rows = rows (values only) of the sheet in which the table was found.
        iHeaders = 0-based index of the header row in rows.
        iTrailing = 0-based index of rows after the table, which begins
            with the first row that lacks a key.
    '''
    wb = openpyxl.load_workbook(file, data_only=True) 
    # Potentially look through all sheets for the one to compare.
    # If the --sheet option was specified, then only that one will be 
    # checked for the desired header.
    sheet = None
    iHeaders = None
    rows = None
    title = ""
    for s in wb:
        if sheet:
            # Already found the right sheet.  No need to look at others.
            break
        if sheetTitle:
            if sheetTitle == s.title.strip(): 
                sheet = s
                # sys.stderr.write(f"[INFO] Found sheet: '{s.title}'\n")
            else:
                # sys.stderr.write(f"[INFO] Skipping unwanted sheet: '{s.title}'\n")
                continue
        # Get the rows of cells:
        rows = [ [str(v if v is not None else "").strip() for v in col] for col in s.values ]
        # sys.stderr.write(f"[INFO] file: '{file}' c.value rows: \n{repr(rows)} \n")
        TrimAndPad(rows)
        NoTabs(rows)
        title = s.title
        # Look for the header row.
        iHeaders = GuessHeaderRow(rows, key, s.title)
        if iHeaders is not None:
            sheet = s
            # sys.stderr.write(f"[INFO] In sheet '{s.title}' found table headers at row {iHeaders+1}\n")
            break
    if not sheet:
        if sheetTitle:
            sys.stderr.write(f"[ERROR] Sheet not found: '{sheetTitle}'\n")
            sys.stderr.write(f" in file: '{file}'\n")
            sys.exit(1)
        if key:
            sys.stderr.write(f"[ERROR] Key not found: '{key}'\n")
            sys.stderr.write(f" in file: '{file}'\n")
            sys.exit(1)
        sys.stderr.write(f"[ERROR] Unable to find header row\n")
        sys.stderr.write(f" in file: '{file}'\n")
        sys.exit(1)
    # Find the key
    jKey = next( (j for j,v in enumerate(rows[iHeaders]) if v == key), -1 )
    if jKey < 0:
        sys.stderr.write(f"[ERROR] Key not found in header row {iHeaders+1}: '{key}'\n")
        sys.stderr.write(f" in file: '{file}'  sheet: '{sheet.title}\n")
        sys.exit(1)
    # Find the end of the table: the first row with an empty key (if any).
    iTrailing = next( (i for i in range(iHeaders, len(rows)) if not rows[i][jKey]), len(rows) )
    # sys.stderr.write(f"[INFO] iTrailing: {iTrailing} file: '{file}'\n")
    # sys.stderr.write(f"[INFO] file: '{file}' rows: \n{repr(rows)} \n")
    return (title, rows, iHeaders, iTrailing)

##################### RemoveTrailingEmpties #####################
def RemoveTrailingEmpties(items):
    ''' Remove trailing empty items from the given list of items,
    returning a new resulting list.
    '''
    nItems = next( (i+1 for i in range(len(items)-1, -1, -1) if items[i]), 0 )
    return items[0 : nItems]

##################### CompareRows #####################
def CompareRows(diffRows, oldRows, iOldStart, nOldRows, newRows, iNewStart, nNewRows, nDiffHeaders):
    ''' Compare the old and new leading or trailing rows.   Trailing empty cells
    are ignored.  The first item in each returned diffRow is one of {=, -, +}.
    Modifies diffRows in place.
    '''
    # sys.stderr.write(f"CompareRows\n")
    # Concatenate the cells in each row, separated by tabs.
    oldLeadingLines = [ "\t".join( RemoveTrailingEmpties(oldRows[i]) ) for i in range(iOldStart, nOldRows) ]
    newLeadingLines = [ "\t".join( RemoveTrailingEmpties(newRows[i]) ) for i in range(iNewStart, nNewRows) ]
    rawDiffs = simplediff.diff(oldLeadingLines, newLeadingLines)
    # simplediff.diff returns pairs: (d, dList)
    # where:    
    #   d = diff mark: one of {=, -, +} 
    #   dList = a list of items that were the same, deleted or added.
    # So we need to flatten the dLists, so that we just have pairs: (d, item).
    # flattened = [val for sublist in list_of_lists for val in sublist]
    flatDiffs = [ (d[0], v) for d in rawDiffs for v in d[1] ]
    # Prepend the diff mark:
    diffLeadingLines = [ d[0] + "\t" + d[1] for d in flatDiffs ]
    for line in diffLeadingLines:
        partialRow = line.split("\t")
        diffRow = [ (partialRow[i] if i < len(partialRow) else '') for i in range(nDiffHeaders+1) ]
        diffRows.append(diffRow)
    # sys.stderr.write(f"diffRows:\n{repr(diffRows)}\n")

##################### CompareHeaders #####################
def CompareHeaders(oldHeaders, oldHeaderIndex, newHeaders, newHeaderIndex):
    ''' Compare the oldHeaders with newHeaders, returning a combined
    list of headers.
    Return:
        diffHeaderMarks = {=, -, +}, one for each diffHeader
        diffHeaders = Combined old and new headers
    '''
    # Headers are treated as column keys: they must be unique.
    if '' in oldHeaders:
        raise ValueError(f"[ERROR] Old header row {iOldHeaders+1} contains an empty header\n")
    if '' in newHeaders:
        raise ValueError(f"[ERROR] New header row {iNewHeaders+1} contains an empty header\n")
    # Build up the diffHeaders from the newHeaders plus any deleted oldHeaders
    diffHeaders = []
    diffHeaderMarks = []    # {=, -, +}
    # First copy any initial deleted headers
    for i, h in enumerate(oldHeaders):
        if h not in newHeaderIndex:
            diffHeaders.append(h)
            diffHeaderMarks.append('-')
        else:
            break
    # Now copy the newHeaders into diffHeader, but each time one has
    # a corresponding oldHeader that has any deleted headers after it,
    # also copy them in.
    for i, h in enumerate(newHeaders):
        diffHeaders.append(h)
        if h in oldHeaderIndex:
            diffHeaderMarks.append('=')
            j = oldHeaderIndex[h] + 1
            while j < len(oldHeaders) and oldHeaders[j] not in newHeaderIndex:
                diffHeaders.append(oldHeaders[j])
                diffHeaderMarks.append('-')
                j += 1
        else:
            diffHeaderMarks.append('+')
    return diffHeaderMarks, diffHeaders

##################### CompareBody #####################
def CompareBody(diffRows, diffHeaders, key, 
        oldRows, oldHeaders, iOldHeaders, iOldTrailing, oldHeaderIndex, 
        newRows, newHeaders, iNewHeaders, iNewTrailing, newHeaderIndex):
    ''' Compare rows in the body of the table.  Modifies diffRows
    by appending diffRows for the table body.  The first cell of each diffRow 
    will be one of {=, -, +, c-, c+}.
    '''
    # Make lists of oldKeys and newKeys.
    iOldKey = oldHeaderIndex[key]
    oldKeys = [ oldRows[i][iOldKey] for i in range(iOldHeaders+1, iOldTrailing) ]
    # oldKeyIndex will index directly into oldRows, which means that
    # the index is offset by iOldHeaders+1, to get past the leading lines
    # and the header row.
    oldKeyIndex = {}
    for i, v in enumerate(oldKeys):
        r = i+iOldHeaders+1 
        if v == '':
            raise  ValueError(f"[ERROR] Table in oldFile contains an empty key at row {r+1}\n")
        if v in oldKeyIndex:
            raise  ValueError(f"[ERROR] Table in oldFile contains a duplicate key on row {r+1}: '{v}'\n")
        oldKeyIndex[v] = r
    jNewKey = newHeaderIndex[key]
    # sys.stderr.write(f"jNewKey: {jNewKey} iNewHeaders: {iNewHeaders} iNewTrailing: {iNewTrailing}\n")
    newKeys = [ newRows[i][jNewKey] for i in range(iNewHeaders+1, iNewTrailing) ]
    # newKeyIndex will index directly into newRows, which means that
    # the index is offset by iNewHeaders+1, to get past the leading lines
    # and the header row.
    newKeyIndex = {}
    for i, v in enumerate(newKeys):
        r = i+iNewHeaders+1 
        if v == '':
            raise  ValueError(f"[INTERNAL ERROR] Table in newFile contains an empty key at row {r+1}\n")
        if v in newKeyIndex:
            raise  ValueError(f"[ERROR] Table in newFile contains a duplicate key on row {r+1}: '{v}'\n")
        newKeyIndex[v] = r
    # Make the diff list of rows.
    # diffRows will not include the header row, but each row in it
    # will have a diff mark {-, +, c-, c+} as its first item.
    # First copy any initial deleted rows
    # sys.stderr.write(f"diffHeaders: {repr(diffHeaders)}\n")
    for k in oldKeys:
        if k not in newKeyIndex:
            oldRow = oldRows[oldKeyIndex[k]]
            oldDiffRow = [ '-' ]
            oldDiffRow.extend( [ (oldRow[oldHeaderIndex[h]] if h in oldHeaderIndex else '') for h in diffHeaders ] )
            diffRows.append(oldDiffRow)
        else:
            break
    # Now copy the newRows into diffRows, marking each diff row 
    # as one of {=, -, +, c-, c+}.  Each time a new row has
    # a corresponding old row that has any deleted rows after it,
    # also copy them in.  
    commonHeaders = set(oldHeaderIndex.keys()).intersection(set(newHeaderIndex.keys()))
    # sys.stderr.write(f"commonHeaders: {repr(commonHeaders)}\n")
    for ii, k in enumerate(newKeys):
        i = ii+iNewHeaders+1
        newRow = newRows[i]
        newDiffRow = [ '+' ]    # This might later become = or c+
        newDiffRowValues = [ (newRow[newHeaderIndex[h]] if h in newHeaderIndex else '') for h in diffHeaders ]
        newDiffRow.extend(newDiffRowValues)
        if k in oldKeyIndex:
            # Key k is in both oldRows and newRows.  
            oldRow = oldRows[oldKeyIndex[k]]
            # Include old values:
            for j, h in enumerate(diffHeaders):
                if h in oldHeaderIndex and h not in newHeaderIndex:
                    newDiffRow[j+1] = oldRow[oldHeaderIndex[h]]
            # Did the rows change (excluding added/deleted columns)?
            isEqual = next( (False for h in commonHeaders if oldRow[oldHeaderIndex[h]] != newRow[newHeaderIndex[h]]), True )
            if isEqual:
                # No values changed in columns that are in common in this row.
                # Only add one row to diffRows.
                newDiffRow[0] = '='     # Change the marker
                diffRows.append(newDiffRow)
            else:
                # Values changed from oldRow to newRow.  
                oldDiffRow = [ 'c-' ]
                for h in diffHeaders:
                    v = ''
                    if h in newHeaderIndex: v = newRow[newHeaderIndex[h]]
                    if h in oldHeaderIndex: v = oldRow[oldHeaderIndex[h]]
                    oldDiffRow.append(v)
                diffRows.append(oldDiffRow)
                newDiffRow[0] = 'c+'     # Change the marker
                diffRows.append(newDiffRow)
            # Also add any following deleted old rows.
            for j in range(oldKeyIndex[k]+1, iOldTrailing):
                if oldKeys[j - (iOldHeaders+1)] in newKeyIndex:
                    break
                oldRow = oldRows[j]
                oldDiffRow = [ '-' ]
                oldDiffRowValues = [ (oldRow[oldHeaderIndex[h]] if h in oldHeaderIndex else '') for h in diffHeaders ]
                oldDiffRow.extend(oldDiffRowValues)
                diffRows.append(oldDiffRow)
            # Finished k in oldKeyIndex
        else:
            # k is not in oldKeyIndex.  k is a new key.
            newDiffRow = [ '+' ]    # This might later become = or c+
            newDiffRowValues = [ (newRow[newHeaderIndex[h]] if h in newHeaderIndex else '') for h in diffHeaders ]
            newDiffRow.extend( newDiffRowValues )
            diffRows.append(newDiffRow)
        # end of loop: for ii, k in enumerate(newKeys):
    # sys.stderr.write(f"diffRows: \n{repr(diffRows)}\n")
    return diffRows

##################### CompareTables #####################
def CompareTables(oldRows, iOldHeaders, iOldTrailing, newRows, iNewHeaders, iNewTrailing, key):
    ''' Compare the old and new tables, and any leading or trailing rows.  
    Returns:
        diffRows = Rows of diff cells.  The first cell of each row is
            a marker with one of: {=, +, -, c-, c+}.
        iDiffHeaders = Index of the first header row.  There will be two
            header rows (old and new) if any headers changed.
        iDiffBody = Index of the table body (either iDiffHeaders+1
            or iDiffHeaders+2).
        iDiffTrailing = Index of any trailing rows after the table, if any.
    '''
    ###### Compare the headers, i.e., the column names.
    # This is done first to figure out how many resulting columns we'll need.
    # The resulting list of headers will be the union of old and new headers,
    # some headers being deleted, some added and some unchanged.
    # The main difficulty is in deciding the best ordering for them,
    # which will basically be the same order as the new headers,
    # but with deleted headers interspersed.
    # Old and new headers are treated as column keys: they must be unique.
    # And they must not contain any empty header.
    oldHeaders = oldRows[iOldHeaders]
    oldHeaderIndex = { v: i for i, v in enumerate(oldHeaders) }
    newHeaders = newRows[iNewHeaders]
    newHeaderIndex = { v: i for i, v in enumerate(newHeaders) }
    diffHeaderMarks, diffHeaders = CompareHeaders(oldHeaders, oldHeaderIndex, newHeaders, newHeaderIndex)
    ###### Compare the lines before the start of the table.
    nDiffHeaders = len(diffHeaders)     # N columns in diffRows table
    diffRows = []
    CompareRows(diffRows, oldRows, 0, iOldHeaders, newRows, 0, iNewHeaders, nDiffHeaders)
    iDiffHeaders = len(diffRows)
    iDiffBody = iDiffHeaders + 2    # 2 for old and new header rows
    # sys.stderr.write(f"iOldTrailing: {iOldTrailing} iNewTrailing: {iNewTrailing}\n")
    ###### Add old and new headers to diffRows.
    if len(diffHeaders) == len(oldHeaders):
        # No columns were added or deleted.
        iDiffBody = iDiffHeaders + 1    # Only one header row after all
        diffRow = [ '=' ]
        diffRow.extend(oldHeaders)
        diffRows.append(diffRow)
    else:
        # At least one column was added or deleted.
        diffRow = [ 'c-' ]
        diffRow.extend( [ (oldHeaders[oldHeaderIndex[h]] if h in oldHeaderIndex else '') for h in diffHeaders ] )
        diffRows.append(diffRow)
        diffRow = [ 'c+' ]
        diffRow.extend( [ (newHeaders[newHeaderIndex[h]] if h in newHeaderIndex else '') for h in diffHeaders ] )
        diffRows.append(diffRow)
    ###### Compare the table body rows.
    # Compare rows, excluding columns that were added or deleted.
    CompareBody(diffRows, diffHeaders, key, 
        oldRows, oldHeaders, iOldHeaders, iOldTrailing, oldHeaderIndex, 
        newRows, newHeaders, iNewHeaders, iNewTrailing, newHeaderIndex)
    iDiffTrailing = len(diffRows)
    ###### Compare trailing rows (after the table).
    CompareRows(diffRows, oldRows, iOldTrailing, len(oldRows), newRows, iNewTrailing, len(newRows), nDiffHeaders)
    return diffRows, iDiffHeaders, iDiffBody, iDiffTrailing
    sys.exit(0)

##################### WriteDiffFile #####################
def WriteDiffFile(diffRows, iDiffHeaders, iDiffBody, iDiffTrailing, key, outFile):
    ''' Write the diffs to the outFile as XLSX, highlighting
    changed rows/columns/cells.
    '''
    nColumns = len(diffRows[0])     # Includes marker column
    oldHeaders = diffRows[iDiffHeaders]
    newHeaders = diffRows[iDiffHeaders]
    if iDiffBody > iDiffHeaders + 1:
        # At least one header changed from old to new, so we'll have
        # two header rows instead of one.
        newHeaders = diffRows[iDiffHeaders+1]
    jKey = next( j for j in range(nColumns) if oldHeaders[j] == key )
    # Create the Excel spreadsheet and fill it with data.
    outWb = openpyxl.Workbook()
    outSheet = outWb.active
    # Fill the sheet with data
    for diffRow in diffRows:
        outSheet.append(diffRow)
    # Make a new workbook and copy the table into it.
    fillChangeRow = PatternFill("solid", fgColor="FFFFDD")
    fillDelRow =    PatternFill("solid", fgColor="FFB6C1")
    fillAddRow =    PatternFill("solid", fgColor="B6FFC1")
    fillDelCol =    PatternFill("solid", fgColor="FFDDE2")
    fillAddCol =    PatternFill("solid", fgColor="DDFFE2")
    fillKeyCol =    PatternFill("solid", fgColor="E8E8FF")
    fillIgnore =    PatternFill("solid", fgColor="E0E0E0")
    # Determine column highlights for added/deleted columns. 
    # colFills will highlight added or deleted columns.
    colFills = [ None for j in range(nColumns) ]
    for j in range(nColumns):
        if oldHeaders[j] == '':  colFills[j] = fillAddCol
        if newHeaders[j] == '':  colFills[j] = fillDelCol

    # Highlight the spreadsheet.
    wsRows = tuple(outSheet.rows)
    for i, diffRow in enumerate(diffRows):
        rowMark = diffRow[0]
        rowFill = None
        if rowMark == '-': rowFill = fillDelRow 
        if rowMark == '+': rowFill = fillAddRow 
        if rowMark == 'c-' or rowMark == 'c+': rowFill = fillChangeRow 
        if i < iDiffHeaders or i >= iDiffTrailing:
            # This is a leading or trailing row.
            # Only use row fills.
            if rowFill:
                for j in range(nColumns):
                    wsRows[i][j].fill = rowFill
        else:
            # This is either a header row or a body row.
            # Apply row fills first, so they'll be overridden
            # by column fills.
            for j in range(nColumns):
                if oldHeaders[j] == key: 
                    wsRows[i][j].fill = fillKeyCol
                if rowMark and rowFill:
                    wsRows[i][j].fill = rowFill
                if i < iDiffBody:
                    # A header row
                    wsRows[i][j].fill = fillKeyCol
                if colFills[j]:
                    wsRows[i][j].fill = colFills[j]
                if j>0 and rowMark == 'c+' and (not colFills[j]) and diffRow[j] != diffRows[i-1][j]:
                    # Highlight this cell and the one above it, if non-empty
                    if diffRows[i-1][j]: wsRows[i-1][j].fill = fillDelRow
                    if diffRows[i][j]:   wsRows[i][j].fill =   fillAddRow

    outSheet.title = 'Differences'
    outWb.save(outFile)
    sys.stderr.write(f"[INFO] Wrote: '{outFile}'\n")

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
                    help='Output file of differences', required=True)
    argParser.add_argument('oldFile', metavar='oldFile.xlsx', type=str,
                    help='Old spreadsheet (*.xlsx)')
    argParser.add_argument('newFile', metavar='newFile.xlsx', type=str,
                    help='New spreadsheet (*.xlsx)')
    # sys.stderr.write(f"[INFO] calling print_using....\n")
    args = argParser.parse_args()
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
    global DEFAULT_KEY
    key = DEFAULT_KEY
    if args.key:
        key = args.key
    outFile = args.out
    if not outFile:
        sys.stderr.write("[ERROR] Output filename must be specified: --out=outFile.xlsx\n")
        sys.stderr.write(Usage())
        sys.exit(1)
    # sys.stderr.write("args: \n" + repr(args) + "\n\n")
    # These will be rows of values-only:
    (oldTitle, oldRows, iOldHeaders, iOldTrailing) = FindTable(args.oldFile, oldSheetTitle, key)
    if iOldHeaders is None:
        raise ValueError(f"[ERROR] Could not find header row in newFile: '{args.oldFile}'")
    (newTitle, newRows, iNewHeaders, iNewTrailing) = FindTable(args.newFile, newSheetTitle, key)
    if iNewHeaders is None:
        raise ValueError(f"[ERROR] Could not find header row in newFile: '{args.newFile}'")
    sys.stderr.write(f"[INFO] Old table rows: {iOldHeaders+1}-{iOldTrailing} file: '{args.oldFile}' sheet: '{oldTitle}'\n")
    sys.stderr.write(f"[INFO] New table rows: {iNewHeaders+1}-{iNewTrailing} file: '{args.newFile}' sheet: '{newTitle}'\n")

    diffRows, iDiffHeaders, iDiffBody, iDiffTrailing = CompareTables(oldRows, iOldHeaders, iOldTrailing, newRows, iNewHeaders, iNewTrailing, key)
    WriteDiffFile(diffRows, iDiffHeaders, iDiffBody, iDiffTrailing, key, outFile)
    sys.exit(0)


######################################################
if __name__ == '__main__':
    main()
    exit(0)

