import zipfile
import xml.etree.ElementTree as ET
import json
import re

file_path = r'c:\Users\abidf\OneDrive\Documents\dev\experiment\padangmanjadi2\MIB for Exhibition.xlsx'

with zipfile.ZipFile(file_path, 'r') as z:
    # 1. Workbook info (sheet names)
    wb_tree = ET.fromstring(z.read('xl/workbook.xml'))
    ns = {'ns': 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'}
    sheets = [s.get('name') for s in wb_tree.findall('.//ns:sheet', ns)]
    print('Sheets:', sheets)

    # 2. Shared Strings
    shared_strings = []
    if 'xl/sharedStrings.xml' in z.namelist():
        ss_tree = ET.fromstring(z.read('xl/sharedStrings.xml'))
        for si in ss_tree.findall('ns:si', ns):
            t_elems = si.findall('.//ns:t', ns)
            text = ''.join([t.text for t in t_elems if t.text is not None])
            shared_strings.append(text)

    print(f'Total shared strings: {len(shared_strings)}')

    # 3. Sheet data
    sheet_tree = ET.fromstring(z.read('xl/worksheets/sheet1.xml'))
    all_rows = []
    for row in sheet_tree.findall('.//ns:row', ns):
        r_num = int(row.get('r'))
        row_dict = {}
        for c in row.findall('ns:c', ns):
            cell_ref = c.get('r')
            col_letter = re.match(r'([A-Z]+)', cell_ref).group(1)
            t = c.get('t')
            v = c.find('ns:v', ns)
            val = v.text if v is not None else None
            if t == 's' and val is not None:
                val = shared_strings[int(val)]
            row_dict[col_letter] = val
        all_rows.append((r_num, row_dict))

print(f'Total rows extracted: {len(all_rows)}')
for r_num, row_dict in all_rows[:50]:
    print(f'Row {r_num}: {row_dict}')
