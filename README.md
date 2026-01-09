# SANS Index Creator

A Python utility to generate a professionally formatted alphabetical index document from Excel data.

## Description

This script reads index entries from an Excel file and creates a beautifully formatted Word document (`SANS_Index.docx`) with:

- **Alphabetical Organization**: Entries grouped by their first letter with letter headings
- **Professional Formatting**:
  - Arial font with optimized sizing
  - Page numbers in the footer
  - Two-column tables for organized layout
  - Minimal margins for better readability
  - Fixed row heights for consistent appearance
- **Sortable Entries**: Entries are automatically sorted alphabetically (case-insensitive)

## Requirements

- Python 3.x
- `openpyxl` - for reading Excel files
- `python-docx` - for creating Word documents

## Installation

Install the required dependencies:

```bash
pip install openpyxl python-docx
```

## Usage

1. **Prepare your data** in an Excel file named `index.xlsx`:
   - Column A: Entry labels (the text to be indexed)
   - Column B: Page references (e.g., "1.234" or page numbers)

2. **Run the script**:
   ```bash
   python sans-index-creator.py
   ```

3. **Output**: The script generates `SANS_Index.docx` with the formatted index.

## Example Input (index.xlsx)

| Column A (Label) | Column B (Page Reference) |
|---|---|
| Apple | 1.001 |
| Banana | 2.045 |
| Cherry | 3.100 |
| Apricot | 1.200 |

## Example Output Structure

The generated document will:
- Group entries by first letter (A, B, C, etc.)
- Create a heading for each letter section
- Display entries in a table with label and page reference columns
- Include automatic page numbering in the footer

## Error Handling

If no data is found in `index.xlsx`, the script will:
- Generate a document with a message indicating no entries were found
- Exit gracefully with a notification

## Notes

- Entries starting with non-alphabetic characters are ignored
- The script processes the first two columns of the spreadsheet
- All text is converted to strings for consistency
- Document margins are set to 0.5 inches on all sides
