# Ausbildungsnachweis Converter

Python tool for converting German vocational training reports (Ausbildungsnachweise) from Excel to Markdown format.

Built this during my Fachinformatiker training to keep Markdown copies of my weekly reports in Obsidian while maintaining Excel originals for printing and signatures.

## Requirements

- Python 3.8+
- pandas
- openpyxl

## Installation

pip install pandas openpyxl

## Usage

1. Create `input` and `output` directories in your project folder
2. Place your Excel files in the `input` directory
3. Run the converter:

python ausbildungsnachweis_converter.py

The script will:
- Convert all Excel files to Markdown
- Name files with calendar week prefix (e.g., `2025-KW02-...`)
- Generate a ZIP archive with all converted files
- Save everything to the `output` directory

## Expected File Format

Excel files should follow this naming pattern:


AusbildungsnachweisU27_DD.MM-DD.MM.xlsx
AusbildungsnachweisU27_DD.MM.YY-DD.MM.YY.xlsx

Examples:
- `AusbildungsnachweisU27_06.01.25-10.01.25.xlsx`
- `AusbildungsnachweisU27_10.03-14.03.xlsx`

## Output Format

Converts to clean Markdown with:
- KW (calendar week) prefix for chronological sorting
- Daily activity lists
- Hour tracking per day
- Metadata (name, course, year)

Example output:

KW02 - Ausbildungsnachweis (06.01.2025 - 10.01.2025)

Name: Nicolay, Dimitrios
Ausbildung: Fachinformatiker SI - U27B (IHK)
Jahr: 2025
06.01.2025

    Introduction to network protocols

    Setting up development environment

    Team meeting and sprint planning

Stunden: 9


## Notes

Year detection uses Excel metadata when available, otherwise falls back to filename or default year. This handles the 2024-2025 transition correctly.