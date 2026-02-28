import pandas as pd
import docx

try:
    doc = docx.Document(r'd:\Data MODEL\Co attainment Examples.docx')
    doc_text = '\n'.join([p.text for p in doc.paragraphs])
    print("DOCX Content:")
    print(doc_text)
except Exception as e:
    print("Error reading DOCX:", e)

print("\nExcel Content:")
try:
    xls = pd.ExcelFile(r'd:\Data MODEL\Input 1.xlsx')
    for sheet in xls.sheet_names:
        print(f"\nSheet {sheet}:")
        df = pd.read_excel(r'd:\Data MODEL\Input 1.xlsx', sheet_name=sheet)
        print(df.head(10))
        print("Columns:", df.columns.tolist())
except Exception as e:
    print("Error reading Excel:", e)
