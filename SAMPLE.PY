import os
from win32com import client

def _base_dir():
    try:
        return os.path.dirname(os.path.abspath(__file__))
    except NameError:
        return os.getcwd()

# Absolute paths to your files (adjust filenames as needed)
excel_path = os.path.abspath(os.path.join(_base_dir(), '..', 'Efficiency Test 15V.xlsx'))
word_path = os.path.abspath(os.path.join(_base_dir(), '..', 'output.docx'))

if not os.path.exists(excel_path):
    raise FileNotFoundError(f"Excel file not found: {excel_path}")
if not os.path.exists(word_path):
    raise FileNotFoundError(f"Word file not found: {word_path}")

excel = client.Dispatch("Excel.Application")
word = client.Dispatch("Word.Application")

doc = word.Documents.Open(word_path)
book = excel.Workbooks.Open(excel_path)

# Adjust the sheet and range as needed
sheet = book.Worksheets(1)  # or use sheet name
named_range = "table"  # Change if your range is named differently

# Validate named range exists
try:
    rng = sheet.Range(named_range)
except Exception:
    raise ValueError(f"Named range '{named_range}' not found in Excel sheet.")
rng.Copy()

# Find placeholder in Word and paste table
findtext = "table here"  # Set this to the placeholder text in your Word doc (required)
if not findtext:
    raise ValueError("Please set 'findtext' to the placeholder text in the Word document where the table should be pasted.")

target = doc.Content
if target.Find.Execute(FindText=findtext):
    table_range = doc.Range(Start=target.Start, End=target.End)
    table_range.PasteExcelTable(False, False, False)
else:
    print(f"Placeholder '{findtext}' not found in Word document.")