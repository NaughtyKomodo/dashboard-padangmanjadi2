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
            
        if r_num > 1 and any(row_dict.values()):
            raw_brand = (row_dict.get('D') or '').strip()
            brand_norm = raw_brand
            if raw_brand.lower() == 'sany':
                brand_norm = 'SANY'
            elif raw_brand.lower() in ['mercedes bens', 'mercedes benz']:
                brand_norm = 'Mercedes-Benz'
            elif raw_brand.lower() == 'caterpillar':
                brand_norm = 'Caterpillar (CAT)'
            elif raw_brand.lower() == 'hytachi':
                brand_norm = 'Hitachi'
                
            raw_links = row_dict.get('F') or ''
            # parse links
            links = []
            for lk in raw_links.replace('\n', ',').split(','):
                lk = lk.strip()
                if lk:
                    # extract google drive id
                    match = re.search(r'id=([a-zA-Z0-9_-]+)', lk)
                    if not match:
                        match = re.search(r'/d/([a-zA-Z0-9_-]+)', lk)
                    drive_id = match.group(1) if match else None
                    
                    # Construct embed and preview URLs
                    preview_iframe_url = f'https://drive.google.com/file/d/{drive_id}/preview' if drive_id else lk
                    thumbnail_url = f'https://drive.google.com/thumbnail?id={drive_id}&sz=w800' if drive_id else lk
                    
                    links.append({
                        'url': lk,
                        'drive_id': drive_id,
                        'iframe_url': preview_iframe_url,
                        'thumbnail_url': thumbnail_url
                    })

            rows.append({
                'id': r_num - 1,
                'lokasi_raw': (row_dict.get('A') or '').strip(),
                'lokasi_area': (row_dict.get('B') or '').strip(),
                'sector': (row_dict.get('C') or '').strip(),
                'brand_raw': raw_brand,
                'brand': brand_norm,
                'aplikasi_masuk': (row_dict.get('E') or '').strip(),
                'photos': links,
                'info_tambahan': (row_dict.get('G') or '').strip()
            })

# Save as data.js
js_content = f"window.MIB_DATA = {json.dumps(rows, indent=2, ensure_ascii=False)};\n"
with open(r'c:\Users\abidf\OneDrive\Documents\dev\experiment\padangmanjadi2\data.js', 'w', encoding='utf-8') as f:
    f.write(js_content)

print(f"Generated data.js successfully with {len(rows)} items with iframe_url support.")
