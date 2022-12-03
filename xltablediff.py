#!/usr/bin/env python3.8

# Copyright 2022 by David Booth
# License: Apache 2.0
# Repo: https://github.com/dbooth-boston/xltablediff
# This code also uses simplediff, https://github.com/paulgb/simplediff/
# which is licensed under the zlib/libpng license:
# <http://www.opensource.org/licenses/zlib-license.php>

# Show value differences between two spreadsheet tables (old and new).
# The tables may also have unrelated leading and/or trailing rows before
# and after the tables.   The leading and trailing rows are also compared,
# though they are compared as lines.
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
# Optionally columns of the new table may be appended to old table,
# or merged into the old table.  Use './xltablediff.py --help' for options.
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
        possibleKeys = list of headers with unique non-empty column values
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
    possibleKeys = []
    if iHeaders is not None:
        headers = rows[iHeaders]
        # sys.stderr.write(f"headers: {repr(headers)}\n")
        for j in range(len(headers)):
            colValues = set()
            dupeFound = None
            for i in range(iHeaders+1, len(rows)):
                v = rows[i][j]
                if v in colValues:
                    dupeFound = True
                    break
                colValues.add(v)
            if not dupeFound: possibleKeys.append(headers[j])
    # sys.stderr.write(f"possibleKeys: {repr(possibleKeys)}\n")
    return iHeaders, possibleKeys

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
        key = Key used: either the one that was passed in or the first
            possibleKey if no key was specified.
    '''
    sys.stderr.write(f"[INFO] Reading file: '{file}'\n")
    wb = None
    try:
        wb = openpyxl.load_workbook(file, data_only=True) 
    except ValueError as e:
        s = str(e)
        # sys.stderr.write(f"[INFO] Caught exception: '{s}'\n")
        if s.startswith('Value does not match pattern'):
            sys.stderr.write(f"[ERROR] Unable to load file: '{file}'\n If a sheet uses a filter, try eliminating the filter.\n")
            sys.exit(1)
        raise e
    # Potentially look through all sheets for the one to compare.
    # If the --sheet option was specified, then only that one will be 
    # checked for the desired header.
    sheet = None
    iHeaders = None
    rows = None
    title = ""
    allPossibleKeys = set()
    for s in wb:
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
        title = s.title.strip()
        # Look for the header row.
        iHeaders, possibleKeys = GuessHeaderRow(rows, key, title)
        allPossibleKeys.update(possibleKeys)
        if iHeaders is None:
            # Not found.  Maybe the key is wrong.
            if key:
                # Try again without specifying the key, in order to
                # suggest possible keys.
                iOther, otherPossibleKeys = GuessHeaderRow(rows, None, title)
                allPossibleKeys.update(otherPossibleKeys)
        else:
            # Found a header row.
            sheet = s
            # sys.stderr.write(f"[INFO] In sheet '{s.title}' found table headers at row {iHeaders+1}\n")
        if sheet:
            break
    if not sheet:
        if sheetTitle:
            sys.stderr.write(f"[ERROR] Sheet not found: '{sheetTitle}'\n")
            sys.stderr.write(f" in file: '{file}'\n")
            sys.exit(1)
        if key:
            sys.stderr.write(f"[ERROR] Key not found: '{key}'\n")
            sys.stderr.write(f" in file: '{file}'\n")
            pKeys = " ".join(sorted(map((lambda v: f"'{v}'"), allPossibleKeys)))
            if allPossibleKeys:
                sys.stderr.write(f" Potential keys: {pKeys}\n")
            sys.exit(1)
        sys.stderr.write(f"[ERROR] Unable to find header row\n")
        sys.stderr.write(f" in file: '{file}'\n")
        sys.exit(1)
    # Find the key
    if key: allPossibleKeys = set([key])
    jKey = next( (j for j,v in enumerate(rows[iHeaders]) if v in allPossibleKeys), -1 )
    if jKey < 0:
        sys.stderr.write(f"[ERROR] Key not found in header row {iHeaders+1}: '{key}'\n")
        sys.stderr.write(f" in file: '{file}'  sheet: '{sheet.title}\n")
        sys.exit(1)
    if not key:
        key = rows[iHeaders][jKey]
        sys.stderr.write(f"[INFO] Assuming key: '{key}'\n")
    # Find the end of the table: the first row with an empty key (if any).
    iTrailing = next( (i for i in range(iHeaders, len(rows)) if not rows[i][jKey]), len(rows) )
    # sys.stderr.write(f"[INFO] iTrailing: {iTrailing} file: '{file}'\n")
    # sys.stderr.write(f"[INFO] file: '{file}' rows: \n{repr(rows)} \n")
    return (title, rows, iHeaders, iTrailing, key)

##################### RemoveTrailingEmpties #####################
def RemoveTrailingEmpties(items):
    ''' Remove trailing empty items from the given list of items,
    returning a new resulting list.
    '''
    nItems = next( (i+1 for i in range(len(items)-1, -1, -1) if items[i]), 0 )
    return items[0 : nItems]

##################### CompareLeadingTrailingRows #####################
def CompareLeadingTrailingRows(diffRows, oldRows, iOldStart, nOldRows, newRows, iNewStart, nNewRows, nDiffHeaders):
    ''' Compare the old and new leading or trailing rows.   Trailing empty cells
    are ignored.  The first item in each returned diffRow is one of {=, -, +}.
    Modifies diffRows in place.
    '''
    # sys.stderr.write(f"CompareLeadingTrailingRows\n")
    # Concatenate the cells in each row, separated by tabs.
    oldLeadingLines = [ "\t".join( RemoveTrailingEmpties(oldRows[i]) ) for i in range(iOldStart, nOldRows) ]
    newLeadingLines = [ "\t".join( RemoveTrailingEmpties(newRows[i]) ) for i in range(iNewStart, nNewRows) ]
    rawDiffs = diff(oldLeadingLines, newLeadingLines)
    # diff returns pairs: (d, dList)
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
        iEmpty = next( (i for i in range(len(oldHeaders)) if oldHeaders[i] == ''), -1 )
        letter = openpyxl.utils.cell.get_column_letter(iEmpty+1)
        raise ValueError(f"[ERROR] Empty header in column {letter} of old table\n")
    if '' in newHeaders:
        iEmpty = next( (i for i in range(len(newHeaders)) if newHeaders[i] == ''), -1 )
        letter = openpyxl.utils.cell.get_column_letter(iEmpty+1)
        raise ValueError(f"[ERROR] Empty header in column {letter} of new table\n")
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
def CompareBody(diffRows, diffHeaders, key, ignoreHeaders,
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
    # Remove from commonHeaders columns that should be ignored:
    commonHeaders = set(oldHeaderIndex.keys()).intersection(set(newHeaderIndex.keys()))
    # sys.stderr.write(f"commonHeaders: {repr(commonHeaders)}\n")
    ignoreSet = set(ignoreHeaders)
    remainingHeaders = ignoreSet.difference(commonHeaders)
    if remainingHeaders:
        h = sorted(remainingHeaders).join(" ")
        sys.stderr.write(f"[ERROR] Bad --ignore column name(s): h\n Column headers specified with --ignore must exist in both old and new tables\n")
        sys.exit(1)
    compareHeaders = commonHeaders.difference(ignoreSet)
    # Now copy the newRows into diffRows, marking each diff row 
    # as one of {=, -, +, c-, c+}.  Each time a new row has
    # a corresponding old row that has any deleted rows after it,
    # also copy them in.  
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
            isEqual = next( (False for h in compareHeaders if oldRow[oldHeaderIndex[h]] != newRow[newHeaderIndex[h]]), True )
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
def CompareTables(oldRows, iOldHeaders, iOldTrailing, newRows, iNewHeaders, iNewTrailing, key, ignoreHeaders, command):
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
    # nDiffHeader does not include the marker column.
    # The total number of columns in diffRows will be nDiffHeaders+1.
    nDiffHeaders = len(diffHeaders)     
    diffRows = []
    ###### Echo the command on the first row?
    if command:
        # Add one to the length for the marker column:
        cmdRow = [ '' for j in range(len(diffHeaders)+1) ]
        cmdRow[0] = '#'
        cmdRow[1] = command
        diffRows.append(cmdRow)
    ###### Compare the lines before the start of the table.
    CompareLeadingTrailingRows(diffRows, oldRows, 0, iOldHeaders, newRows, 0, iNewHeaders, nDiffHeaders)
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
    CompareBody(diffRows, diffHeaders, key, ignoreHeaders,
        oldRows, oldHeaders, iOldHeaders, iOldTrailing, oldHeaderIndex, 
        newRows, newHeaders, iNewHeaders, iNewTrailing, newHeaderIndex)
    iDiffTrailing = len(diffRows)
    ###### Compare trailing rows (after the table).
    CompareLeadingTrailingRows(diffRows, oldRows, iOldTrailing, len(oldRows), newRows, iNewTrailing, len(newRows), nDiffHeaders)
    return diffRows, iDiffHeaders, iDiffBody, iDiffTrailing

##################### AppendTable #####################
def AppendTable(oldRows, iOldHeaders, iOldTrailing, newRows, iNewHeaders, iNewTrailing, key, outFile):
    ''' Append columns of new table to old table.
    Write the resulting spreadsheet to outFile.
    '''
    # Copy oldRows from the beginning through the headers. 
    oldHeaders = oldRows[iOldHeaders]
    newHeaders = newRows[iNewHeaders]
    outRows = []
    emptyNewRow = [ '' for j in range(len(newHeaders)) ]
    for i in range(iOldHeaders):
        outRow = oldRows[i].copy()
        outRow.extend(emptyNewRow)
        outRows.append(outRow)
    # Find the key column in oldHeaders
    jOldKey = next( j for j in range(len(oldHeaders)) if oldHeaders[j] == key )
    # Find the key column in newHeaders
    jNewKey = next( j for j in range(len(newHeaders)) if newHeaders[j] == key )
    # Make a newKeyIndex
    newKeyIndex = { newRows[i][jNewKey]: i for i in range(iNewHeaders+1, len(newRows)) }
    # Copy oldRows, appending newRows that exist
    for i in range(iOldHeaders, iOldTrailing):
        outRow = oldRows[i].copy()
        newRow = emptyNewRow
        k = outRow[jOldKey]
        if i == iOldHeaders:
            newRow = newHeaders
        elif k in newKeyIndex:
            newRow = newRows[newKeyIndex[k]]
        outRow.extend(newRow)
        outRows.append(outRow)
    # Copy trailing oldRows
    for i in range(iOldTrailing, len(oldRows)):
        outRow = oldRows[i].copy()
        outRow.extend(emptyNewRow)
        outRows.append(outRow)
    # Create the output spreadsheet and fill it with data.
    outWb = openpyxl.Workbook()
    outSheet = outWb.active
    for outRow in outRows:
        outSheet.append(outRow)
    # Highlight the new columns.
    fillAddCol =    PatternFill("solid", fgColor="DDFFE2")
    wsRows = tuple(outSheet.rows)
    nOldColumns = len(oldHeaders)
    # Determine how many columns are needed.
    nTotalColumns = len(oldHeaders) + len(newHeaders)
    for i in range(iOldHeaders, iOldTrailing):
        outRow = outRows[i]
        for j in range(nOldColumns, nTotalColumns):
            wsRows[i][j].fill = fillAddCol
    # Write the output file
    outSheet.title = 'Appended'
    outWb.save(outFile)
    sys.stderr.write(f"[INFO] Wrote: '{outFile}'\n")

##################### MergeTable #####################
def MergeTable(oldRows, iOldHeaders, iOldTrailing, newRows, iNewHeaders, iNewTrailing, key, outFile, mergeHeaders):
    ''' Merge specified columns of old table to new table.
    Write the resulting spreadsheet to outFile.
    '''
    oldHeaders = oldRows[iOldHeaders]
    newHeaders = newRows[iNewHeaders]
    # Create the output spreadsheet and fill it with old data.
    outWb = openpyxl.Workbook()
    outSheet = outWb.active
    for oldRow in oldRows:
        outSheet.append(oldRow)
    # Prepare to highlight the new columns.
    fillAddCol = PatternFill("solid", fgColor="CCFFC2")
    fillChange = PatternFill("solid", fgColor="FFFFAA")
    wsRows = tuple(outSheet.rows)
    # newKeyIndex will be used to look up rows in newRows.
    newKeyIndex = {}
    jNewKey = next( j for j in range(len(newHeaders)) if newHeaders[j] == key )
    jOldKey = next( j for j in range(len(oldHeaders)) if oldHeaders[j] == key )
    for i in range(iNewHeaders+1, iNewTrailing):
        v = newRows[i][jNewKey]
        if v == '':
            raise  ValueError(f"[ERROR] Table in newFile contains an empty key at row {i+1}\n")
        if v in newKeyIndex:
            raise  ValueError(f"[ERROR] Table in newFile contains a duplicate key on row {i+1}: '{v}'\n")
        newKeyIndex[v] = i
    newHeaderIndex = { h: j for j, h in enumerate(newHeaders) }
    for jOld in range(len(oldHeaders)):
        h = oldHeaders[jOld]
        if h not in mergeHeaders: 
            continue
        jNew = newHeaderIndex[h]
        for i in range(iOldHeaders+1, iOldTrailing):
            oldRow = oldRows[i]
            oldValue = oldRow[jOld]
            oldKey = oldRow[jOldKey]
            newValue = ''
            if oldKey in newKeyIndex:
                newValue = newRows[newKeyIndex[oldKey]][jNew]
                if newValue != oldValue:
                    wsRows[i][jOld].value = newValue
                    if oldValue == '':
                        # New value
                        wsRows[i][jOld].fill = fillAddCol
                    else:
                        # Value changed
                        wsRows[i][jOld].fill = fillChange

    # Write the output file
    outSheet.title = 'Merged'
    outWb.save(outFile)
    sys.stderr.write(f"[INFO] Wrote: '{outFile}'\n")

##################### WriteDiffFile #####################
def WriteDiffFile(diffRows, iDiffHeaders, iDiffBody, iDiffTrailing, key, ignoreHeaders, outFile):
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
    ignoreSet = set(ignoreHeaders)
    for j in range(nColumns):
        if oldHeaders[j] == '':  colFills[j] = fillAddCol
        if newHeaders[j] == '':  colFills[j] = fillDelCol
        # It's enough to check oldHeaders hear, because ignoreSet
        # only includes headers that are in both old and new:
        if oldHeaders[j] in ignoreSet: colFills[j] = fillIgnore
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
    argParser.add_argument('--append', action='store_true',
                    help='Copy the values of oldFile sheet, appending columns of newFile.\n Rows of newFile that do not exist in oldFile are discarded, and leading and trailing rows\n of newFile (before and after the table) are also discarded.  The number of rows in the output file will be the same is in oldFile.')
    argParser.add_argument('--merge', nargs=1, action='extend',
                    help='Output a copy of the old sheet, with values from the specifed MERGE column from the new table merged in.  This option may be repeated to merge more than one column.  The number of rows in the output file will be the same is in oldFile.')
    argParser.add_argument('--ignore', nargs=1, action='extend',
                    help='Ignore the specified column when comparing old and new table rows.  This option may be repeated to ignore multiple columns.  The specified column must exist in both old and new tables.')
    argParser.add_argument('--key',
                    help='Specifies the name of the key column, i.e., its header')
    argParser.add_argument('--out',
                    help='Output file of differences.  This "option" is actually REQUIRED.', required=True)
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
    key = None
    if args.key:
        key = args.key
    outFile = args.out
    if not outFile:
        sys.stderr.write("[ERROR] Output filename must be specified: --out=outFile.xlsx\n")
        sys.stderr.write(Usage())
        sys.exit(1)
    # sys.stderr.write("args: \n" + repr(args) + "\n\n")
    # These will be rows of values-only:
    (oldTitle, oldRows, iOldHeaders, iOldTrailing, key) = FindTable(args.oldFile, oldSheetTitle, key)
    if iOldHeaders is None:
        raise ValueError(f"[ERROR] Could not find header row in newFile: '{args.oldFile}'")
    (newTitle, newRows, iNewHeaders, iNewTrailing, key) = FindTable(args.newFile, newSheetTitle, key)
    if iNewHeaders is None:
        raise ValueError(f"[ERROR] Could not find header row in newFile: '{args.newFile}'")
    sys.stderr.write(f"[INFO] Old table rows: {iOldHeaders+1}-{iOldTrailing} file: '{args.oldFile}' sheet: '{oldTitle}'\n")
    sys.stderr.write(f"[INFO] New table rows: {iNewHeaders+1}-{iNewTrailing} file: '{args.newFile}' sheet: '{newTitle}'\n")
    if args.append and args.merge:
        sys.stderr.write(f"[ERROR] Options --append and --merge cannot be used together.\n")
        sys.exit(1)
    if args.merge:
        # sys.stderr.write(f"args.merge: {repr(args.merge)}\n")
        oldHeadersSet = set(oldRows[iOldHeaders])
        newHeadersSet = set(newRows[iNewHeaders])
        mergeHeaders = set()
        for h in args.merge:
            # sys.stderr.write(f"h: {repr(h)}\n")
            if (h not in oldHeadersSet):
                sys.stderr.write(f"[ERROR] Column specified in --merge='{h}' does not exist in old table.\n")
                sys.exit(1)
            if (h not in newHeadersSet):
                sys.stderr.write(f"[ERROR] Column specified in --merge='{h}' does not exist in new table.\n")
                sys.exit(1)
            mergeHeaders.add(h)
        MergeTable(oldRows, iOldHeaders, iOldTrailing, newRows, iNewHeaders, iNewTrailing, key, outFile, mergeHeaders)
        sys.exit(0)
    if args.append:
        AppendTable(oldRows, iOldHeaders, iOldTrailing, newRows, iNewHeaders, iNewTrailing, key, outFile)
        sys.exit(0)
    ignoreHeaders = args.ignore if args.ignore else []
    # command will be the command string to echo in the first row output.
    command = ''
    if 1:
        argv = sys.argv.copy()
        argv[0] = os.path.basename(argv[0])
        command = " ".join(argv)
    diffRows, iDiffHeaders, iDiffBody, iDiffTrailing = CompareTables(oldRows, iOldHeaders, iOldTrailing, newRows, iNewHeaders, iNewTrailing, key, ignoreHeaders, command)
    WriteDiffFile(diffRows, iDiffHeaders, iDiffBody, iDiffTrailing, key, ignoreHeaders, outFile)
    sys.exit(0)


############################################################################
############################## simplediff ##################################
############################################################################
# Lines below here are from the python version of simplediff
# https://github.com/paulgb/simplediff/
# and used under its specified license:
'''
Copyright (c) 2008 - 2013 Paul Butler and contributors

This sofware may be used under a zlib/libpng-style license:

This software is provided 'as-is', without any express or implied warranty. In
no event will the authors be held liable for any damages arising from the use
of this software.

Permission is granted to anyone to use this software for any purpose, including
commercial applications, and to alter it and redistribute it freely, subject to
the following restrictions:

1. The origin of this software must not be misrepresented; you must not claim
that you wrote the original software. If you use this software in a product, an
acknowledgment in the product documentation would be appreciated but is not
required.

2. Altered source versions must be plainly marked as such, and must not be
misrepresented as being the original software.

3. This notice may not be removed or altered from any source distribution.
'''

'''
Simple Diff for Python version 1.0

Annotate two versions of a list with the values that have been
changed between the versions, similar to unix's `diff` but with
a dead-simple Python interface.

(C) Paul Butler 2008-2012 <http://www.paulbutler.org/>
May be used and distributed under the zlib/libpng license
<http://www.opensource.org/licenses/zlib-license.php>
'''

__all__ = ['diff', 'string_diff', 'html_diff']
__version__ = '1.0'


def diff(old, new):
    '''
    Find the differences between two lists. Returns a list of pairs, where the
    first value is in ['+','-','='] and represents an insertion, deletion, or
    no change for that list. The second value of the pair is the list
    of elements.

    Params:
        old     the old list of immutable, comparable values (ie. a list
                of strings)
        new     the new list of immutable, comparable values
   
    Returns:
        A list of pairs, with the first part of the pair being one of three
        strings ('-', '+', '=') and the second part being a list of values from
        the original old and/or new lists. The first part of the pair
        corresponds to whether the list of values is a deletion, insertion, or
        unchanged, respectively.

    Examples:
        >>> diff([1,2,3,4],[1,3,4])
        [('=', [1]), ('-', [2]), ('=', [3, 4])]

        >>> diff([1,2,3,4],[2,3,4,1])
        [('-', [1]), ('=', [2, 3, 4]), ('+', [1])]

        >>> diff('The quick brown fox jumps over the lazy dog'.split(),
        ...      'The slow blue cheese drips over the lazy carrot'.split())
        ... # doctest: +NORMALIZE_WHITESPACE
        [('=', ['The']),
         ('-', ['quick', 'brown', 'fox', 'jumps']),
         ('+', ['slow', 'blue', 'cheese', 'drips']),
         ('=', ['over', 'the', 'lazy']),
         ('-', ['dog']),
         ('+', ['carrot'])]

    '''

    # Create a map from old values to their indices
    old_index_map = dict()
    for i, val in enumerate(old):
        old_index_map.setdefault(val,list()).append(i)

    # Find the largest substring common to old and new.
    # We use a dynamic programming approach here.
    # 
    # We iterate over each value in the `new` list, calling the
    # index `inew`. At each iteration, `overlap[i]` is the
    # length of the largest suffix of `old[:i]` equal to a suffix
    # of `new[:inew]` (or unset when `old[i]` != `new[inew]`).
    #
    # At each stage of iteration, the new `overlap` (called
    # `_overlap` until the original `overlap` is no longer needed)
    # is built from the old one.
    #
    # If the length of overlap exceeds the largest substring
    # seen so far (`sub_length`), we update the largest substring
    # to the overlapping strings.

    overlap = dict()
    # `sub_start_old` is the index of the beginning of the largest overlapping
    # substring in the old list. `sub_start_new` is the index of the beginning
    # of the same substring in the new list. `sub_length` is the length that
    # overlaps in both.
    # These track the largest overlapping substring seen so far, so naturally
    # we start with a 0-length substring.
    sub_start_old = 0
    sub_start_new = 0
    sub_length = 0

    for inew, val in enumerate(new):
        _overlap = dict()
        for iold in old_index_map.get(val,list()):
            # now we are considering all values of iold such that
            # `old[iold] == new[inew]`.
            _overlap[iold] = (iold and overlap.get(iold - 1, 0)) + 1
            if(_overlap[iold] > sub_length):
                # this is the largest substring seen so far, so store its
                # indices
                sub_length = _overlap[iold]
                sub_start_old = iold - sub_length + 1
                sub_start_new = inew - sub_length + 1
        overlap = _overlap

    if sub_length == 0:
        # If no common substring is found, we return an insert and delete...
        return (old and [('-', old)] or []) + (new and [('+', new)] or [])
    else:
        # ...otherwise, the common substring is unchanged and we recursively
        # diff the text before and after that substring
        return diff(old[ : sub_start_old], new[ : sub_start_new]) + \
               [('=', new[sub_start_new : sub_start_new + sub_length])] + \
               diff(old[sub_start_old + sub_length : ],
                       new[sub_start_new + sub_length : ])


def string_diff(old, new):
    '''
    Returns the difference between the old and new strings when split on
    whitespace. Considers punctuation a part of the word

    This function is intended as an example; you'll probably want
    a more sophisticated wrapper in practice.

    Params:
        old     the old string
        new     the new string

    Returns:
        the output of `diff` on the two strings after splitting them
        on whitespace (a list of change instructions; see the docstring
        of `diff`)

    Examples:
        >>> string_diff('The quick brown fox', 'The fast blue fox')
        ... # doctest: +NORMALIZE_WHITESPACE
        [('=', ['The']),
         ('-', ['quick', 'brown']),
         ('+', ['fast', 'blue']),
         ('=', ['fox'])]

    '''
    return diff(old.split(), new.split())


def html_diff(old, new):
    '''
    Returns the difference between two strings (as in stringDiff) in
    HTML format. HTML code in the strings is NOT escaped, so you
    will get weird results if the strings contain HTML.

    This function is intended as an example; you'll probably want
    a more sophisticated wrapper in practice.

    Params:
        old     the old string
        new     the new string

    Returns:
        the output of the diff expressed with HTML <ins> and <del>
        tags.

    Examples:
        >>> html_diff('The quick brown fox', 'The fast blue fox')
        'The <del>quick brown</del> <ins>fast blue</ins> fox'
    '''
    con = {'=': (lambda x: x),
           '+': (lambda x: "<ins>" + x + "</ins>"),
           '-': (lambda x: "<del>" + x + "</del>")}
    return " ".join([(con[a])(" ".join(b)) for a, b in string_diff(old, new)])


def check_diff(old, new):
    '''
    This tests that diffs returned by `diff` are valid. You probably won't
    want to use this function, but it's provided for documentation and
    testing.

    A diff should satisfy the property that the old input is equal to the
    elements of the result annotated with '-' or '=' concatenated together.
    Likewise, the new input is equal to the elements of the result annotated
    with '+' or '=' concatenated together. This function compares `old`,
    `new`, and the results of `diff(old, new)` to ensure this is true.

    Tests:
        >>> check_diff('ABCBA', 'CBABA')
        >>> check_diff('Foobarbaz', 'Foobarbaz')
        >>> check_diff('Foobarbaz', 'Boobazbam')
        >>> check_diff('The quick brown fox', 'Some quick brown car')
        >>> check_diff('A thick red book', 'A quick blue book')
        >>> check_diff('dafhjkdashfkhasfjsdafdasfsda', 'asdfaskjfhksahkfjsdha')
        >>> check_diff('88288822828828288282828', '88288882882828282882828')
        >>> check_diff('1234567890', '24689')
    '''
    old = list(old)
    new = list(new)
    result = diff(old, new)
    _old = [val for (a, vals) in result if (a in '=-') for val in vals]
    assert old == _old, 'Expected %s, got %s' % (old, _old)
    _new = [val for (a, vals) in result if (a in '=+') for val in vals]
    assert new == _new, 'Expected %s, got %s' % (new, _new)

############################################################################
############################## main ##################################
############################################################################

if __name__ == '__main__':
    main()
    exit(0)

