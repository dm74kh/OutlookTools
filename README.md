[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](LICENSE)
[![Python 3.10+](https://img.shields.io/badge/Python-3.10%2B-blue)]()
[![Platform](https://img.shields.io/badge/Platform-Windows-blue)]()

# Outlook Duplicates Cleaning

OutlookTools provides a set of reproducible scripts and Jupyter notebooks for managing and sanitizing Microsoft Outlook data files (.PST).  
The project focuses on identifying and safely removing duplicate emails across folders, with detailed logging, configurable matching criteria,  and read-only “dry-run” testing modes for non-destructive verification.

## Features
- Recursive duplicate detection across the entire PST hierarchy.  
- Configurable parameters for matching subjects, recipients, and timestamps.  
- Safe “Dry Run” mode for verification before performing moves.  
- Full export of logs and duplicate reports to `.csv` and `.txt`.  
- Compatible with Windows + Outlook Desktop (MAPI/COM) interface. 

## Requirements
Install the dependencies before running:
- Windows 10/11 with installed Microsoft Outlook Desktop
- Python ≥ 3.10
- Packages: pandas, numpy, openpyxl, win32com.client (via pywin32)

To install the required dependencies, run:  
`pip install -r requirements.txt`

## Installation

```bash
git clone https://github.com/dm74kh/OutlookTools.git
cd OutlookTools
pip install -r requirements.txt
```

## Usage
Launch the main notebook in Jupyter:
```bash
jupyter notebook Outlook_Duplicates_Cleaning.ipynb
```
Then follow the step-by-step structure:
1. Initialization and Imports — check Outlook connection and loaded stores.
2. Duplicate Scan (Test Mode) — safe identification of potential duplicates.
3. Duplicate Cleanup (Recursive Mode) — automated removal or relocation to Duplicates_YYYYMMDD folder.

## License

Distributed under the MIT License  
Copyright © 2025 Dmytro Mykhailychenko.