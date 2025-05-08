# Attendance List Combiner

This script processes CSV attendance lists and extracts the "Participants" section from each file. It combines all the participant data into a single Excel file for easier reporting and analysis.

## Features

- Automatically scans a folder of `.csv` files.
- Extracts section **2. Participants** from each file (assumes UTF-16 encoding and tab-separated values).
- Combines all participant data into one Excel file.
- Adds a column indicating the source file for traceability.

## Requirements

Install dependencies using:

```bash
pip install -r requirements.txt
