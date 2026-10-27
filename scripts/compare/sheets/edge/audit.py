"""Conservative whole-workbook snapshot audit. Reject ambiguous formula representations."""
import collections
import json
import math
import re
import sys
import xml.etree.ElementTree as ET
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[3] / 'office-quality'))
from xlsx_calc import NS, address, package, sheets, text_value
sys.path.insert(0, str(Path(__file__).resolve().parents[1] / 'acc'))
from score import same

def normalized(f):
    return re.sub(r'_xlfn\._xlws\.|_xlfn\.|_xlws\.|_xlpm\.', '', f)

def check(root):
    p = package((root / 'input.xlsx').read_bytes())
    data = json.loads((root / 'univer.json').read_text())
    reasons = collections.Counter()
    if any(n.startswith(('xl/tables/', 'xl/pivotTables/', 'xl/externalLinks/')) for n in p):
        reasons['tables-pivots-or-external-links'] += 1
    wb = ET.fromstring(p['xl/workbook.xml'])
    props=wb.find('s:workbookPr',NS)
    calc=wb.find('s:calcPr',NS)
    if props is not None and props.get('date1904') in ('1','true'):
        reasons['date1904-not-audited'] += 1
    if calc is not None and (calc.get('iterate') in ('1','true') or calc.get('fullPrecision') in ('0','false')):
        reasons['iteration-or-display-precision-not-audited'] += 1
    if any(not n.get('name','').startswith('_xlnm.') for n in wb.findall('s:definedNames/s:definedName', NS)):
        reasons['defined-names'] += 1
    ss = [text_value(n) for n in ET.fromstring(p.get('xl/sharedStrings.xml', f'<sst xmlns="{NS["s"]}"/>')).findall('s:si', NS)]
    source_sheets = sheets(p)
    visibility_sensitive=any(sh.find('s:autoFilter',NS) is not None or any(row.get('hidden') in ('1','true') for row in sh.findall('s:sheetData/s:row',NS)) for _,_,sh in source_sheets)
    snapshot_sheets = [data['sheets'][s] for s in data['sheetOrder']]
    if [n for n,_,_ in source_sheets] != [s['name'] for s in snapshot_sheets]:
        reasons['sheet-identity'] += 1
    snapshot_by_name = {s['name']:s for s in snapshot_sheets}
    formulas = inputs = normalized_count = 0
    for name, _, sheet in source_sheets:
        mapped = snapshot_by_name.get(name,{}).get('cellData',{})
        seen = set()
        for cell in sheet.findall('s:sheetData/s:row/s:c',NS):
            pos = address(cell.get('r'))
            seen.add(pos)
            m = mapped.get(str(pos[0]),{}).get(str(pos[1]),{})
            f = cell.find('s:f',NS)
            if f is not None:
                formulas += 1
                if f.get('t') not in (None,'normal') or list(f.attrib):
                    reasons['shared-array-or-special-formula'] += 1
                text = f.text or ''
                body=re.sub(r'"(?:[^"]|"")*"','""',text)
                if re.search(r'(?<![\w.])\$?[A-Z]{1,3}:\$?[A-Z]{1,3}(?!\w)|(?<![\w.])\$?\d+:\$?\d+(?!\w)',body):
                    reasons['whole-row-or-column-reference-not-audited'] += 1
                if '_xlpm.' in text:
                    reasons['lambda-parameter-prefix-rewrite-not-audited'] += 1
                if re.search(r'\bLAMBDA\(\s*[rRcC]\s*,',body):
                    reasons['reserved-lambda-parameter-not-used-as-excel-oracle'] += 1
                if re.search(r'\b(?:INDIRECT|OFFSET)\s*\(',body,re.I):
                    reasons['dynamic-reference-extents-not-audited'] += 1
                if visibility_sensitive and re.search(r'\b(?:SUBTOTAL|AGGREGATE)\s*\(',body,re.I):
                    reasons['visibility-sensitive-formula-not-audited'] += 1
                normalized_count += normalized(text) != text
                if m.get('f') != '=' + normalized(text):
                    reasons['formula-text-mismatch'] += 1
                if m.get('v') is not None:
                    reasons['formula-cache-retained'] += 1
                continue
            v = cell.find('s:v',NS)
            typ = cell.get('t')
            if typ == 'inlineStr':
                expected = text_value(cell)
            elif v is None or v.text is None:
                continue
            elif typ == 's': expected = ss[int(v.text)]
            elif typ == 'b': expected = v.text == '1'
            elif typ in ('str','e'): expected = v.text
            elif typ in (None,'n'): expected = float(v.text)
            else:
                reasons['unsupported-input-type'] += 1
                continue
            inputs += 1
            actual = m.get('v')
            if isinstance(expected,bool):
                ok = m.get('t') == 3 and actual in (expected,int(expected))
            elif isinstance(expected,float):
                ok = isinstance(actual,(int,float)) and not isinstance(actual,bool) and math.isclose(expected,actual,rel_tol=1e-12,abs_tol=1e-9)
            else: ok = actual == expected
            if not ok: reasons['input-value-mismatch'] += 1
        for r,row in mapped.items():
            for c,value in row.items():
                if (int(r),int(c)) not in seen and (value.get('v') is not None or value.get('f')):
                    reasons['extra-cell-value-or-formula'] += 1
    return dict(accepted=not reasons,reasons=dict(reasons),formulas=formulas,inputs=inputs,normalizedPrefixFormulas=normalized_count)

def main(work,out):
    records=[]
    total=collections.Counter()
    samples=[]
    for expected_path in sorted(work.glob('*/*/expected.json')):
        root=expected_path.parent
        e=json.loads(expected_path.read_text())
        try: audit=check(root)
        except Exception as ex: audit=dict(accepted=False,reasons={'audit-error':str(ex)[:160]})
        b=json.loads((root/'bo.values.json').read_text())
        u=json.loads((root/'univer.values.json').read_text())
        base=[(i,c) for i,c in enumerate(e['cells']) if not e.get('timeRandom') and 'volatile' not in c['families'] and not c.get('exclude')]
        scored=[(i,c) for i,c in base if c['value']['kind']!='error' and not (c['value']['kind']=='text' and re.match(r'^#(?:NULL!|DIV/0!|VALUE!|REF!|NAME\?|NUM!|N/A|SPILL!|CALC!)$',c['value']['value']))]
        record=dict(id=e['id'],set=e['set'],audit=audit,scored=len(scored),eligibleBeforeOracleFilter=len(base),excludedOracleErrorOrErrorText=len(base)-len(scored),bo=0,univer=0,status={'bo':b['status'],'univer':u['status']})
        for i,c in scored:
            bv=(b.get('values') or [None]*len(e['cells']))[i]
            uv=(u.get('values') or [None]*len(e['cells']))[i]
            record['bo']+=same(c['value'],bv)
            record['univer']+=same(c['value'],uv)
            if audit['accepted'] and same(c['value'],bv) and not same(c['value'],uv):
                samples.append(dict(id=e['id'],cell=c['cell'],sheet=c['name'],formula=c['formula'],expected=c['value'],bo=bv,univer=uv))
        records.append(record)
        if audit['accepted']:
            total['books']+=1
            total['cells']+=len(scored)
            total['bo']+=record['bo']
            total['univer']+=record['univer']
        else:
            total['excludedBooks']+=1
            total['excludedScoredCells']+=len(scored)
    out.mkdir(parents=True,exist_ok=True)
    (out/'audit.json').write_text(json.dumps(dict(total=dict(total),records=records,boOnlyCorrect=samples),indent=1))
    print(json.dumps(dict(total)))

if __name__=='__main__': main(Path(sys.argv[1]),Path(sys.argv[2]))
