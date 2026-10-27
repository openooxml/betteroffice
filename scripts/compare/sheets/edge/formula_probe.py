"""Deterministic, public-domain one-formula XLSX probes. No cached result in input."""
import hashlib
import json
import sys
from pathlib import Path
from xml.sax.saxutils import escape
from zipfile import ZipFile, ZipInfo, ZIP_DEFLATED

PROBES = [
    ('convert', 'CONVERT(12,"in","m")', 0.3048),
    ('mround-time', 'MROUND(0.378,"0:15")', 0.375),
    ('month-text', 'MONTH("November"&1)', 11),
    ('isoweeknum', 'ISOWEEKNUM(DATE(2021,1,4))', 1),
    ('seriessum', 'SERIESSUM(2,3,-1,{1,2,3})', 22),
    ('hyperlink', 'HYPERLINK("#A2","Go to target")', 'Go to target'),
    ('let', 'LET(x,5,x*2)', 10),
    ('lambda', 'LAMBDA(x,x*2)(5)', 10),
    ('map', 'SUM(MAP({1,2,3},LAMBDA(x,x*2)))', 12),
    ('reduce', 'REDUCE(0,{1,2,3},LAMBDA(a,b,a+b))', 6),
    ('scan', 'SUM(SCAN(0,{1,2,3},LAMBDA(a,b,a+b)))', 10),
    ('byrow', 'SUM(BYROW({1,2;3,4},LAMBDA(row_value,SUM(row_value))))', 10),
    ('bycol', 'SUM(BYCOL({1,2;3,4},LAMBDA(col_value,SUM(col_value))))', 10),
    ('makearray', 'SUM(MAKEARRAY(2,3,LAMBDA(row_index,col_index,row_index*col_index)))', 18),
    ('textsplit', 'COUNTA(TEXTSPLIT("red,green,blue",","))', 3),
    ('textbefore', 'TEXTBEFORE("red-green","-")', 'red'),
    ('textafter', 'TEXTAFTER("red-green","-")', 'green'),
    ('regex-test', 'REGEXTEST("abc123","[0-9]+")', True),
    ('regex-replace', 'REGEXREPLACE("abc123","[0-9]+","X")', 'abcX'),
    ('regex-extract', 'REGEXEXTRACT("abc123","[0-9]+")', '123'),
    ('tocol', 'SUM(TOCOL({1,2;3,4}))', 10),
    ('torow', 'SUM(TOROW({1,2;3,4}))', 10),
    ('wraprows', 'SUM(WRAPROWS({1,2,3,4},2))', 10),
    ('wrapcols', 'SUM(WRAPCOLS({1,2,3,4},2))', 10),
    ('take', 'SUM(TAKE({1,2;3,4},1))', 3),
    ('drop', 'SUM(DROP({1,2;3,4},1))', 7),
    ('choosecols', 'SUM(CHOOSECOLS({1,2;3,4},2))', 6),
    ('chooserows', 'SUM(CHOOSEROWS({1,2;3,4},2))', 7),
    ('expand', 'SUM(EXPAND({1,2},2,2,0))', 3),
    ('vstack', 'SUM(VSTACK({1,2},{3,4}))', 10),
    ('hstack', 'SUM(HSTACK({1;2},{3;4}))', 10),
    ('filter', 'SUM(FILTER({1;2;3},{1;0;1}))', 4),
    ('sortby', 'INDEX(SORTBY({10;20;30},{3;1;2}),1)', 20),
    ('unique', 'COUNTA(UNIQUE({1;1;2;3}))', 3),
    ('sequence', 'SUM(SEQUENCE(2,3))', 21),
    ('xlookup', 'XLOOKUP(1,{0;1;0},{10;20;30})', 20),
    ('xmatch', 'XMATCH(20,{10;20;30})', 2),
    ('aggregate', 'AGGREGATE(9,6,{1,2,3})', 6),
    ('arabic', 'ARABIC("MMXXIV")', 2024),
    ('roman', 'ROMAN(2024)', 'MMXXIV'),
    ('encodeurl', 'ENCODEURL("hello world")', 'hello%20world'),
    ('numbervalue', 'NUMBERVALUE("1.234,5",",",".")', 1234.5),
    ('rsq', 'RSQ({2,4,6},{1,2,3})', 1),
    ('forecast', 'FORECAST.LINEAR(4,{2,4,6},{1,2,3})', 8),
    ('linest', 'INDEX(LINEST({2,4,6},{1,2,3}),1,1)', 2),
    ('intercept', 'INTERCEPT({3,5,7},{1,2,3})', 1),
    ('trend', 'INDEX(TREND({2,4,6},{1,2,3},{4,5}),1,1)', 8),
    ('growth', 'INDEX(GROWTH({2,4,8},{1,2,3},{4,5}),1,1)', 16),
    ('rankeq-emptystrings', 'RANK.EQ(1,{1,2,3,"","",""},1)', 1),
    ('rankavg-emptystrings', 'RANK.AVG(1,{1,2,3,"","",""},1)', 1),
    ('rank-legacy-emptystrings', 'RANK(1,{1,2,3,"","",""},1)', 1),
    ('rankeq-control', 'RANK.EQ(1,{1,2,3},1)', 1),
    ('convert-fraction', 'CONVERT("1 1/4","in","m")', 0.03175),
    ('convert-fraction-half', 'CONVERT("1 1/2","in","m")', 0.0381),
    ('rankref-eq', 'RANK.EQ(1,E1:E6,1)', 1),
    ('rankref-legacy', 'RANK(1,E1:E6,1)', 1),
    ('rankref-control', 'RANK.EQ(1,E1:E3,1)', 1),
    ('isoweeknum-20181231', 'ISOWEEKNUM(DATE(2018,12,31))', 1),
    ('isoweeknum-20191230', 'ISOWEEKNUM(DATE(2019,12,30))', 1),
    ('isoweeknum-20201231', 'ISOWEEKNUM(DATE(2020,12,31))', 53),
    ('isoweeknum-20210111', 'ISOWEEKNUM(DATE(2021,1,11))', 2),
    ('isoweeknum-20210118', 'ISOWEEKNUM(DATE(2021,1,18))', 3),
    ('isoweeknum-20210201', 'ISOWEEKNUM(DATE(2021,2,1))', 5),
    ('isoweeknum-20210607', 'ISOWEEKNUM(DATE(2021,6,7))', 23),
    ('isoweeknum-20211227', 'ISOWEEKNUM(DATE(2021,12,27))', 52),
    ('isoweeknum-20230102', 'ISOWEEKNUM(DATE(2023,1,2))', 1),
    ('mround-time-0126', 'MROUND(0.126,"0:15")', 0.125),
    ('mround-time-0510', 'MROUND(0.51,"0:15")', 0.5104166666666666),
    ('mround-time-0751', 'MROUND(0.751,"0:15")', 0.75),
    ('mround-time-0900', 'MROUND(0.9,"0:15")', 0.8958333333333334),
]

NS = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'
REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'

def workbook(formula, references=False):
    first=f'<row r="1"><c r="A1"><f>{escape(formula)}</f></c>'
    if references:
        first+='<c r="E1"><v>1</v></c>'
    body=first+'</row>'
    if references:
        body+='<row r="2"><c r="E2"><v>2</v></c></row><row r="3"><c r="E3"><v>3</v></c></row>'
        body+=''.join(f'<row r="{r}"><c r="E{r}"><f>""</f></c></row>' for r in range(4,7))
    dimension='A1:E6' if references else 'A1'
    return {
        '[Content_Types].xml': '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/></Types>',
        '_rels/.rels': f'<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="{REL}/officeDocument" Target="xl/workbook.xml"/></Relationships>',
        'xl/workbook.xml': f'<workbook xmlns="{NS}" xmlns:r="{REL}"><sheets><sheet name="Sheet1" sheetId="1" r:id="rId1"/></sheets><calcPr fullCalcOnLoad="1"/></workbook>',
        'xl/_rels/workbook.xml.rels': f'<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="{REL}/worksheet" Target="worksheets/sheet1.xml"/></Relationships>',
        'xl/worksheets/sheet1.xml': f'<worksheet xmlns="{NS}"><dimension ref="{dimension}"/><sheetData>{body}</sheetData></worksheet>',
    }

def main(out, only=None):
    for name, formula, value in PROBES:
        if only and name not in only:continue
        root = out / 'sp' / name
        root.mkdir(parents=True, exist_ok=True)
        path = root / 'input.xlsx'
        with ZipFile(path, 'w', compression=ZIP_DEFLATED) as z:
            for part, xml in workbook(formula,name.startswith('rankref-')).items():
                info = ZipInfo(part, (2000, 1, 1, 0, 0, 0))
                info.compress_type = ZIP_DEFLATED
                z.writestr(info, xml)
        kind = 'bool' if isinstance(value, bool) else 'number' if isinstance(value, (int,float)) else 'text'
        expected = dict(id=name, set='sp', sheets=['Sheet1'], timeRandom=False, oracle='analytical-synthetic', cells=[dict(sheet=0,name='Sheet1',cell='A1',r=0,c=0,formula=formula,value=dict(kind=kind,value=value),families=['other'],unknown=[],tags=[],exclude=None)])
        (root / 'expected.json').write_text(json.dumps(expected))
        print(name, hashlib.sha256(path.read_bytes()).hexdigest())

if __name__ == '__main__':
    main(Path(sys.argv[1]),sys.argv[2].split(',') if len(sys.argv)>2 else None)
