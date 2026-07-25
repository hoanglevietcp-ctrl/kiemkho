#!/usr/bin/env node
import fs from 'node:fs/promises';
import path from 'node:path';
import { execFileSync } from 'node:child_process';

const source = process.argv[2];
const output = process.argv[3] || 'test-uiux/products.json';

if (!source) {
  throw new Error('Usage: node scripts/import-crm-stock.mjs <source.xlsx> [output.json]');
}

// The CRM export is an XLSX archive. Its first sheet is a simple shared-string table,
// so this small extractor avoids introducing a runtime dependency into the static site.
const python = String.raw`
import json, sys, zipfile, xml.etree.ElementTree as ET
source=sys.argv[1]
ns={'m':'http://schemas.openxmlformats.org/spreadsheetml/2006/main'}
with zipfile.ZipFile(source) as z:
  strings=[]
  if 'xl/sharedStrings.xml' in z.namelist():
    root=ET.fromstring(z.read('xl/sharedStrings.xml'))
    strings=[''.join(t.text or '' for t in node.iter('{http://schemas.openxmlformats.org/spreadsheetml/2006/main}t')) for node in root.findall('m:si', ns)]
  root=ET.fromstring(z.read('xl/worksheets/sheet1.xml'))
  rows=[]
  for row in root.findall('.//m:sheetData/m:row', ns):
    record={}
    for cell in row.findall('m:c', ns):
      value=cell.find('m:v', ns)
      text='' if value is None else value.text
      if cell.attrib.get('t')=='s' and text: text=strings[int(text)]
      col=''.join(char for char in cell.attrib['r'] if char.isalpha())
      record[col]=text
    rows.append(record)
items=[]
for row in rows[1:]:
  full=(row.get('B') or '').strip()
  first, separator, rest=full.partition(',')
  items.append({
    'sku': (row.get('A') or '').strip(),
    'name': first.strip() if separator else full,
    'specs': rest.strip() if separator else '',
    'fullName': full,
    'origin': (row.get('C') or '').strip(),
    'stock': float(row.get('D') or 0),
    'price': float(row.get('E') or 0),
    'note': (row.get('F') or '').strip(),
    'unit': (row.get('G') or '').strip(),
    'stockValue': float(row.get('H') or 0),
  })
print(json.dumps({'source':'CRM_Stock_25.07.2026_11.31.11_793.xlsx','exportedAt':'2026-07-25T11:31:11+07:00','items':items}, ensure_ascii=False, separators=(',',':')))
`;

const json = execFileSync('python3', ['-c', python, source], { encoding: 'utf8' });
await fs.mkdir(path.dirname(output), { recursive: true });
await fs.writeFile(output, `${json}\n`);
console.log(`Wrote ${output}`);
