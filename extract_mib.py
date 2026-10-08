import zipfile
import xml.etree.ElementTree as ET
import json
import re

file_path = r'c:\Users\abidf\OneDrive\Documents\dev\experiment\padangmanjadi2\MIB for Exhibition.xlsx'

with zipfile.ZipFile(file_path, 'r') as z:
    shared_strings = []
    if 'xl/sharedStrings.xml' in z.namelist():
        ss_tree = ET.fromstring(z.read('xl/sharedStrings.xml'))
        ns = {'ns': 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'}
        for si in ss_tree.findall('ns:si', ns):
            t_elems = si.findall('.//ns:t', ns)
            text = ''.join([t.text for t in t_elems if t.text is not None])
            shared_strings.append(text)

    sheet_tree = ET.fromstring(z.read('xl/worksheets/sheet1.xml'))
    ns = {'ns': 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'}
    
    rows = []
    headers = []
    
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
            
        if r_num == 1:
            headers = [row_dict.get(chr(65+i), f'Col{i+1}') for i in range(7)]
            print('Headers:', headers)
        else:
            if any(row_dict.values()):
                rows.append({
                    'row_id': r_num,
                    'lokasi_raw': (row_dict.get('A') or '').strip(),
                    'lokasi_area': (row_dict.get('B') or '').strip(),
                    'sector': (row_dict.get('C') or '').strip(),
                    'brand': (row_dict.get('D') or '').strip(),
                    'aplikasi_masuk': (row_dict.get('E') or '').strip(),
                    'foto_links': [link.strip() for link in (row_dict.get('F') or '').split(',') if link.strip()],
                    'info_tambahan': (row_dict.get('G') or '').strip()
                })

print(f'Total records parsed: {len(rows)}')

# Unique values check
sectors = set(r['sector'] for r in rows if r['sector'])
brands = set(r['brand'] for r in rows if r['brand'])
areas = set(r['lokasi_area'] for r in rows if r['lokasi_area'])

print('Sectors:', sorted(list(sectors)))
print('Brands:', sorted(list(brands)))
print('Areas:', sorted(list(areas)))

with open(r'c:\Users\abidf\OneDrive\Documents\dev\experiment\padangmanjadi2\data.json', 'w', encoding='utf-8') as f:
    json.dump(rows, f, indent=2, ensure_ascii=False)
