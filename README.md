# CocoFile
Compare Coco mock data

# CocoComparisonProject

## Overview
This repository contains a Python script, **`Coco10.py`**, that compares data between two Excel sheets:
- **Coco Coco** (Table1)
- **Coco Coco Land** (Table2)

The script handles:
- Parsing and comparing Noel columns (with or without `_0001`, `_0002`, etc.).
- Managing differences in row counts (i.e., when one sheet has more rows than the other).
- Generating a final Excel output with color-coded styling:
  - **Green** cells for Table1.
  - **Blue** cells for Table2.
  - **Red** fill for missing data.
  - Etc. (adjust the details here to match your script's logic).
