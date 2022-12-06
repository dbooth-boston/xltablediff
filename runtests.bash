#! /bin/bash -x

# Diff test:
  xltablediff.py  --key ID test1old.xlsx test1new.xlsx --out test1diff.xlsx

# Ignore test:
  xltablediff.py  --key ID --ignore Color test1old.xlsx test1new.xlsx --out test1ignore.xlsx

# Merge test:
  xltablediff.py  --key ID --merge Color test1old.xlsx test1new.xlsx --out test1merge.xlsx

# Append test:
  xltablediff.py  --key ID --append test1old.xlsx test1new.xlsx --out test1append.xlsx



