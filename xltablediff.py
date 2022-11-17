#!/usr/bin/env python3.8

# Copyright 2022 by David Booth
# License: Apache 2.0
# Note that this code uses simplediff, which is used 
# under the zlib/libpng license
# <http://www.opensource.org/licenses/zlib-license.php>


# Compare two spreadsheet tables, which may have lines before
# the tables.  We try to be a little intelligent about detecting
# added or removed columns and rows.  Added or removed columns
# are detected by differences in the table headers.  Added or
# removed rows are detected by added or deleted keys.
#

def Usage():
    sys.stderr.write(f'''Usage:
   xltablediff [ --sheet=MySheet1 ] [ --key=K ] oldFile.xlsx newFile.xlsx --out=outFile.xlsx

Where:
    --key=K     
        Specifies K as the name of the key column that uniquely
        identifies each row in the table, where K must be one
        of the fields in the header row.  Key defaults to 'id'.

    --sheet=S
        Specifies S as the name of the sheet containing the table
        to be compared.  The sheet will be guessed if it is not specified.

    --out=outFile.xslx
        Specifies outFile.xlsx as the output file to write.  This "option"
        is mandatory.
''')

# Idea for future enhancement: Allow the table to specified (within
# the spreadsheet) by specifying a range, such as: --table=B10:G16

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
from openpyxl.styles import PatternFill, Fill, Font
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
DEFAULT_KEY = "id"    # Header of key column for the table.

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
        rows = [ [str(v or "").strip() for v in col] for col in s.values ]
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
            sys.stderr.write(f"[ERROR] Key not found: {key}\n")
            sys.stderr.write(f" in file: '{file}'\n")
            sys.exit(1)
        sys.stderr.write(f"[ERROR] Unable to find header row\n")
        sys.stderr.write(f" in file: '{file}'\n")
        sys.exit(1)
    return (title, rows, iHeaders)

##################### RemoveTrailingEmpties #####################
def RemoveTrailingEmpties(items):
    ''' Remove trailing empty items from the given list of items,
    returning a new resulting list.
    '''
    nItems = next( (i+1 for i in range(len(items)-1, -1, -1) if items[i]), 0 )
    return items[0 : nItems]

##################### MakeDiffRow #####################
# isEqual, oldDiffRow, newDiffRow = 
def MakeDiffRow(diffHeaders, commonHeaders, oldRow, oldHeaderIndex, newRow, newHeaderIndex):
    ''' Create an old and new diffRow having the columns of diffHeaders.
    Parameters:
        diffHeaders = List of both old and new headers
        commonHeaders = set of headers common to both old and new
        oldRow = Old row of values to be compared
        oldHeaderIndex = Dict from old header to its index in oldRow
        newRow = New row of values to be compared
        newHeaderIndex = Dict from new header to its index in newRow
    Returns:
        isEqual = True iff all old and new values in commonHeaders columns 
                are equal.  Other columns are ignored.
        oldDiffRow = the old diffRow that was created
        newDiffRow = the new diffRow that was created
    '''
    oldDiffRow = [ (oldRow[oldHeaderIndex[h]] if h in oldHeaderIndex else '') for h in diffHeaders ]
    newDiffRow = [ (newRow[newHeaderIndex[h]] if h in newHeaderIndex else '') for h in diffHeaders ]
    isEqual = next( (False for h in commonHeaders if oldRow[oldHeaderIndex[h]] == newRow[newHeaderIndex[h]]), True )
    return isEqual, oldDiffRow, newDiffRow

##################### CompareLeadingRows #####################
def CompareLeadingRows(oldRows, iOldHeaders, newRows, iNewHeaders, nDiffHeaders):
    ''' Compare the old and new leading rows.   Trailing empty cells
    are ignored.  The first item in each returned diffRow is one of {=, -, +}.
    Return:
        diffRows = diffed rows, padded to nDiffHeaders+1
    '''
    # sys.stderr.write(f"CompareLeadingRows\n")
    # Concatenate the cells in each row, separated by tabs.
    oldLeadingLines = [ "\t".join( RemoveTrailingEmpties(oldRows[i]) ) for i in range(0, iOldHeaders) ]
    newLeadingLines = [ "\t".join( RemoveTrailingEmpties(newRows[i]) ) for i in range(0, iNewHeaders) ]
    rawDiffs = simplediff.diff(oldLeadingLines, newLeadingLines)
    # simplediff.diff returns pairs: (d, dList)
    # where:    
    #   d = diff mark: one of {=, -, +} 
    #   dList = a list of items that were the same, deleted or added.
    # So we need to flatten the dLists, so that we just have pairs: (d, item).
    # flattened = [val for sublist in list_of_lists for val in sublist]
    flatDiffs = [ (d[0], v) for d in rawDiffs for v in d[1] ]
    diffLeadingLines = [ d[0] + "\t" + d[1] for d in flatDiffs ]
    diffRows = []
    for line in diffLeadingLines:
        partialRow = line.split("\t")
        diffRow = [ (partialRow[i] if i < len(partialRow) else '') for i in range(nDiffHeaders+1) ]
        diffRows.append(diffRow)
    # sys.stderr.write(f"diffRows:\n{repr(diffRows)}\n")
    return diffRows

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

##################### CompareTableRows #####################
def CompareTableRows(diffRows, diffHeaders, key, 
        oldRows, oldHeaders, iOldHeaders, oldHeaderIndex, 
        newRows, newHeaders, iNewHeaders, newHeaderIndex):
    ''' Compare rows in the body of the table.  Modifies diffRows
    by appending diffRows for the table body.  The first cell of each diffRow 
    will be one of {=, -, +, c-, c+}.
    '''
    # Make lists of oldKeys and newKeys.
    iOldKey = oldHeaderIndex[key]
    oldKeys = [ oldRows[i][iOldKey] for i in range(iOldHeaders+1, len(oldRows)) ]
    # oldKeyIndex will index directly into oldRows, which means that
    # the index is offset by iOldHeaders+1, to get past the leading lines
    # and the header row.
    oldKeyIndex = { v: i+iOldHeaders+1 for i, v in enumerate(oldKeys) }
    if '' in oldKeyIndex:
        raise  ValueError("[ERROR] Table in oldFile contains an empty key\n")
    if len(oldKeys) != len(oldKeyIndex):
        raise  ValueError("[ERROR] Table in oldFile contains a duplicate key\n")
    iNewKey = newHeaderIndex[key]
    newKeys = [ newRows[i][iNewKey] for i in range(iNewHeaders+1, len(newRows)) ]
    # newKeyIndex will index directly into newRows, which means that
    # the index is offset by iNewHeaders+1, to get past the leading lines
    # and the header row.
    newKeyIndex = { v: i+iNewHeaders+1 for i, v in enumerate(newKeys) }
    if '' in newKeyIndex:
        raise  ValueError("[ERROR] Table in newFile contains an empty key\n")
    if len(newKeys) != len(newKeyIndex):
        raise  ValueError("[ERROR] Table in newFile contains a duplicate key\n")
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
            # Key k is in both oldRows and newRows.  Did the rows change?
            oldRow = oldRows[oldKeyIndex[k]]
            isEqual = next( (False for h in commonHeaders if oldRow[oldHeaderIndex[h]] != newRow[newHeaderIndex[h]]), True )
            if isEqual:
                # No values changed in columns that are in common in this row.
                # Only add one row to diffRows.
                newDiffRow[0] = '='     # Change the marker
                diffRows.append(newDiffRow)
            else:
                # Values changed from oldRow to newRow.  
                oldDiffRow = [ 'c-' ]
                oldDiffRowValues = [ (oldRow[oldHeaderIndex[h]] if h in oldHeaderIndex else '') for h in diffHeaders ]
                oldDiffRow.extend(oldDiffRowValues)
                diffRows.append(oldDiffRow)
                newDiffRow[0] = 'c+'     # Change the marker
                diffRows.append(newDiffRow)
            # Also add any following deleted old rows.
            for j in range(oldKeyIndex[k]+1, len(oldRows)):
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
def CompareTables(oldRows, iOldHeaders, newRows, iNewHeaders, key):
    ''' Compare the old and new tables.  Return a diffs workbook.
    '''
    ###### Compare the headers, i.e., the column names.
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
    diffRows = CompareLeadingRows(oldRows, iOldHeaders, newRows, iNewHeaders, nDiffHeaders)
    iDiffHeaders = len(diffRows)
    iDiffBody = iDiffHeaders + 2    # 2 for old and new header rows
    ###### Add old and new headers to diffRows.
    if len(diffHeaders) == len(oldHeaders):
        # No columns were added or deleted.
        iDiffBody = iDiffHeaders + 1    # Only one header row after all
        diffRow = [ '=' ]
        diffRow.extend(oldHeaders)
        diffRows.append(diffRow)
    else:
        # At least one column was added or deleted.
        diffRow = [ '-' ]
        diffRow.extend( [ (oldHeaders[oldHeaderIndex[h]] if h in oldHeaderIndex else '') for h in diffHeaders ] )
        diffRows.append(diffRow)
        diffRow = [ '+' ]
        diffRow.extend( [ (newHeaders[newHeaderIndex[h]] if h in newHeaderIndex else '') for h in diffHeaders ] )
        diffRows.append(diffRow)
    ##### Compare the table body rows.
    # Compare rows, excluding columns that were added or deleted.
    CompareTableRows(diffRows, diffHeaders, key, 
        oldRows, oldHeaders, iOldHeaders, oldHeaderIndex, 
        newRows, newHeaders, iNewHeaders, newHeaderIndex)
    return diffRows, iDiffHeaders, iDiffBody
    sys.exit(0)
    
##################### WriteDiffFile #####################
def WriteDiffFile(diffRows, iDiffHeaders, iDiffBody, key, outFile):
    ''' Write the diffs to the outFile as XLSX, highlighting
    changed rows/columns/cells.
    '''
    nColumns = len(diffRows[0])     # Includes marker column
    # Make a new workbook and copy the table into it.
    fillChangeRow = PatternFill("solid", fgColor="FFFFDD")
    fillDelRow =    PatternFill("solid", fgColor="FFBBBB")
    fillAddRow =    PatternFill("solid", fgColor="BBFFBB")
    fillDelCol =    PatternFill("solid", fgColor="FFDDDD")
    fillAddCol =    PatternFill("solid", fgColor="DDFFDD")
    fillKeyCol =    PatternFill("solid", fgColor="DDDDFF")
    fillIgnore =    PatternFill("solid", fgColor="DDDDDD")
    if iDiffBody > iDiffHeaders + 1:
        oldHeaders = diffRows[iDiffHeaders]
        newHeaders = diffRows[iDiffHeaders+1]
    for j in range(nColumns):
        if oldHeaders[j] == '':  colFills[j] = fillAddCol
        if newHeaders[j] == '':  colFills[j] = fillDelCol
        if oldHeaders[j] == key: colFills[j] = fillKeyCol
    # Determine column highlights for added/deleted columns. 
    # colFills will highlight added or deleted columns.
    colFills = [ None for j in range(nColumns) ]
    outWb = openpyxl.Workbook()
    outSheet = outWb.active
    # Fill the sheet with data
    for diffRow in diffRows:
        outSheet.append(diffRow)
    if 1:
        # Highlight the spreadsheet.
        wsRows = tuple(outSheet.rows)
        for i, diffRow in enumerate(diffRows):
            rowMark = diffRow[0]
            rowFill = None
            if rowMark == '-': rowFill = fillDelRow 
            if rowMark == '+': rowFill = fillAddRow 
            if rowMark == 'c-' or rowMark == 'c+': rowFill = fillChangeRow 
            if i < iDiffHeaders:
                # This is a leading row (before the headers).
                # Only use row fills.
                if rowFill:
                    for j in range(nColumns):
                        wsRows[i][j].fill = rowFill
            else:
                        # **** STOPPED HERE ***
                # This is either a header row or a body row.
                # Apply row fills first, so they'll be overridden
                # by column fills.
                for j in range(nColumns):
                    if rowMark and rowFill:
                        wsRows[i][j].fill = rowFill
                    if j>0 and colFills[j]:
                        wsRows[i][j].fill = colFills[j]
                    if j>0 and rowMark == 'c+' and (not colFills[j]) and diffRow[j] != diffRows[i-1][j]:
                        # Highlight this cell and the one above it
                        wsRows[i-1][j].fill = fillDelRow
                        wsRows[i][j].fill =   fillAddRow
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
    key = DEFAULT_KEY
    if args.key:
        key = args.key
    outFile = args.out
    if not outFile:
        sys.stderr.write("[ERROR] Output filename must be specified: --out=outFile.xlsx\n")
        sys.exit(1)
    sys.stderr.write("args: \n" + repr(args) + "\n\n")
    sys.stderr.write("otherArgs: \n" + repr(otherArgs) + "\n\n")
    # These will be rows of values-only:
    (oldTitle, oldRows, oldIHeader) = FindTable(args.oldFile, oldSheetTitle, key)
    if oldIHeader is None:
        raise ValueError(f"[ERROR] Could not find header row in newFile: '{args.oldFile}'")
    (newTitle, newRows, newIHeader) = FindTable(args.newFile, newSheetTitle, key)
    if newIHeader is None:
        raise ValueError(f"[ERROR] Could not find header row in newFile: '{args.newFile}'")
    sys.stderr.write(f"[INFO] Old file: '{args.oldFile}' sheet: '{oldTitle}' row: {oldIHeader+1}\n")
    sys.stderr.write(f"[INFO] New file: '{args.newFile}' sheet: '{newTitle}' row: {newIHeader+1}\n")

    diffRows, iDiffHeaders, iDiffBody = CompareTables(oldRows, oldIHeader, newRows, newIHeader, key)
    WriteDiffFile(diffRows, iDiffHeaders, iDiffBody, key, outFile)
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

