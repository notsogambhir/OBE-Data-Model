import zipfile
import xml.etree.ElementTree as ET

def read_xlsx_all_sheets(path):
    out = []
    try:
        with zipfile.ZipFile(path) as z:
            # Get shared strings
            shared_strings = []
            if 'xl/sharedStrings.xml' in z.namelist():
                tree = ET.fromstring(z.read('xl/sharedStrings.xml'))
                ns = {'main': 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'}
                for si in tree.findall('main:si', ns):
                    texts = [t.text for t in si.findall('.//main:t', ns) if t.text is not None]
                    shared_strings.append("".join(texts))
            
            # Read all sheets
            workbook_tree = ET.fromstring(z.read('xl/workbook.xml'))
            ns = {'main': 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'}
            sheets = workbook_tree.find('main:sheets', ns)
            
            sheet_names = {}
            for i, sheet in enumerate(sheets.findall('main:sheet', ns)):
                # sheet.get('name') is the actual name
                # find the corresponding file, usually 'xl/worksheets/sheet{i+1}.xml'
                # but let's just guess by order
                sheet_names[f'xl/worksheets/sheet{i+1}.xml'] = sheet.get('name')
                
            for sheet_path, sheet_name in sheet_names.items():
                if sheet_path in z.namelist():
                    out.append(f"\n--- Sheet: {sheet_name} ---")
                    sheet_tree = ET.fromstring(z.read(sheet_path))
                    rows = []
                    sheetData = sheet_tree.find('main:sheetData', ns)
                    for row in sheetData.findall('main:row', ns):
                        row_data = []
                        for c in row.findall('main:c', ns):
                            val = c.find('main:v', ns)
                            if val is not None:
                                if c.get('t') == 's':
                                    row_data.append(shared_strings[int(val.text)])
                                else:
                                    row_data.append(val.text)
                            else:
                                row_data.append("")
                        rows.append(row_data)
                        if len(rows) > 15:
                            break
                    for r in rows[:10]:
                        out.append(str(r))
    except Exception as e:
        out.append(str(e))
    return '\n'.join(out)

print(read_xlsx_all_sheets(r'd:\Data MODEL\Input 1.xlsx'))
