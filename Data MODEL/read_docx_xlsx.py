import zipfile
import xml.etree.ElementTree as ET
import pandas as pd

def extract_text_from_docx(docx_path):
    try:
        with zipfile.ZipFile(docx_path) as docx:
            xml_content = docx.read('word/document.xml')
        tree = ET.fromstring(xml_content)
        ns = {'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'}
        paragraphs = []
        for p in tree.findall('.//w:p', ns):
            texts = [node.text for node in p.findall('.//w:t', ns) if node.text]
            if texts:
                paragraphs.append(''.join(texts))
        return '\n'.join(paragraphs)
    except Exception as e:
        return str(e)

with open(r'd:\Data MODEL\output.txt', 'w', encoding='utf-8') as f:
    f.write("--- DOCX Content ---\n")
    f.write(extract_text_from_docx(r'd:\Data MODEL\Co attainment Examples.docx'))
    
    f.write("\n\n--- Excel Content ---\n")
    try:
        xls = pd.ExcelFile(r'd:\Data MODEL\Input 1.xlsx')
        for sheet in xls.sheet_names:
            f.write(f"\nSheet: {sheet}\n")
            df = pd.read_excel(xls, sheet_name=sheet)
            f.write(df.head(10).to_string())
            f.write("\nColumns: " + str(df.columns.tolist()) + "\n")
    except Exception as e:
        f.write(f"Error reading Excel: {e}\n")
