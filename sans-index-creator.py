from openpyxl import load_workbook
from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_ROW_HEIGHT_RULE
from docx.shared import Pt, Inches
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
from collections import defaultdict
import string
import os
import sys

def add_page_number(run):
    fldChar1 = OxmlElement('w:fldChar')
    fldChar1.set(qn('w:fldCharType'), 'begin')
    run._r.append(fldChar1)
    instrText = OxmlElement('w:instrText')
    instrText.text = 'PAGE'
    run._r.append(instrText)
    fldChar2 = OxmlElement('w:fldChar')
    fldChar2.set(qn('w:fldCharType'), 'end')
    run._r.append(fldChar2)

# Ensure Excel file exists
if not os.path.exists('index.xlsx'):
    print("Error: 'index.xlsx' not found. Place the file 'index.xlsx' in the current directory.")
    sys.exit(1)

# Load Excel file
wb = load_workbook('index.xlsx')
ws = wb.active


# Parse data: (label, page_ref)
index_entries = []

for row in ws.iter_rows(min_row=1, values_only=True):
    if row[0]:  # Assuming column A = label, column B = x.yyy format
        label = str(row[0]).strip()
        page_ref = str(row[1]).strip() if row[1] else ""
        index_entries.append((label, page_ref))

# Sort alphabetically
index_entries.sort(key=lambda x: x[0].lower())

# Group by first letter
grouped = defaultdict(list)
for label, page_ref in index_entries:
    first_letter = label[0].upper()
    if first_letter.isalpha():
        grouped[first_letter].append((label, page_ref))

# Create Word document
doc = Document()

# Set default font size to smaller
doc.styles['Normal'].font.size = Pt(11)
doc.styles['Normal'].font.name = 'Arial'

# Set heading font to Arial
doc.styles['Heading 1'].font.name = 'Arial'
doc.styles['Heading 1'].font.size = Pt(22)

# Set document to one column
section = doc.sections[0]
sectPr = section._sectPr
cols = sectPr.xpath('./w:cols')[0]
cols.set(qn('w:num'), '1')

# Set smaller margins
section.left_margin = Inches(0.5)
section.right_margin = Inches(0.5)
section.top_margin = Inches(0.5)
section.bottom_margin = Inches(0.5)

# Add page numbers to footer
footer = section.footer
footer_paragraph = footer.paragraphs[0]
footer_paragraph.text = "Page "
run = footer_paragraph.add_run()
add_page_number(run)

if not index_entries:
    doc.add_paragraph("No index entries found in index.xlsx. Please add data to the Excel file.")
    doc.save('SANS_Index.docx')
    print("No data found. Generated document with message.")
    exit()


# Add main index with letter sections
for letter in sorted(grouped.keys()):
    heading = doc.add_paragraph(letter)
    heading.style = 'Heading 1'
    
    table = doc.add_table(rows=1, cols=2)
    table.width = Inches(2.5)
    table.columns[0].width = Inches(1.5)
    table.columns[1].width = Inches(1)
        
    # Add entries for this letter
    for label, page_ref in grouped[letter]:
        row_cells = table.add_row().cells
        row_cells[0].text = label
        row_cells[1].text = page_ref
    
    # Set smaller row height
    for row in table.rows:
        row.height = Pt(15)
        row.height_rule = WD_ROW_HEIGHT_RULE.EXACTLY

# Save document
doc.save('SANS_Index.docx')
print("✅ Index generated: SANS_Index.docx")
