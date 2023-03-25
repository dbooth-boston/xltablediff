#! /bin/bash -x

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


echo '========================================================='
echo 'Comparing expected vs new results...'

xltablediff.py  test1diff-expected.xlsx test1diff.xlsx --out /dev/null
xltablediff.py  test1ignore-expected.xlsx test1ignore.xlsx --out /dev/null
xltablediff.py  test1merge-expected.xlsx test1merge.xlsx --out /dev/null
xltablediff.py  test1oldAppend-expected.xlsx test1oldAppend.xlsx --out /dev/null
xltablediff.py  test1newAppend-expected.xlsx test1newAppend.xlsx --out /dev/null

