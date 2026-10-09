import argparse
from datetime import datetime, timezone
import hashlib
import json
from pathlib import Path
import platform
import re
import xml.etree.ElementTree as ET
import shutil
import subprocess
import tempfile

from xlsx_calc import NS, formula_targets, package, results, sheets, uncached_workbook

METHOD = 'excel-full-rebuild-v1'


def validate_source(parts):
    if any('vbaproject' in name.lower() or '/macrosheets/' in name.lower() or '/externallinks/' in name.lower() for name in parts):
        raise ValueError('Calculation references require self-contained workbooks without macros')
    formulas = [node.text or '' for _,_,sheet in sheets(parts) for node in sheet.findall('.//s:f',NS)]
    formulas += [node.text or '' for node in ET.fromstring(parts['xl/workbook.xml']).findall('s:definedNames/s:definedName',NS)]
    if any(re.search(r'(?i)\b(?:NOW|TODAY|RAND|RANDBETWEEN|RANDARRAY|CELL|INFO|STOCKHISTORY|WEBSERVICE|RTD)\s*\(',formula) for formula in formulas):
        raise ValueError('Volatile or external calculation inputs require a controlled oracle')


def capture(source, output, timeout=60):
    data = source.read_bytes()
    parts = package(data)
    validate_source(parts)
    targets, count = formula_targets(parts)
    if not targets:
        raise ValueError('Workbook has no formulas')
    prepared = uncached_workbook(data, targets)
    container = Path.home() / 'Library/Containers/com.microsoft.Excel/Data/Documents/BetterOfficeBenchmark'
    container.mkdir(parents=True, exist_ok=True)
    stage = Path(tempfile.mkdtemp(prefix='calculation-',dir=container))
    (stage/'input.xlsx').write_bytes(prepared)
    script = f'''on run argv
    set inputPath to item 1 of argv
    set firstPath to item 2 of argv
    set secondPath to item 3 of argv
    with timeout of {timeout} seconds
        tell application "Microsoft Excel"
            if (count of workbooks) is not 0 then error "Close existing workbooks before calculation capture"
            set previousAlerts to display alerts
            set display alerts to false
            set wb to missing value
            try
                set wb to open workbook workbook file name inputPath update links do not update links read only true password "" write reserved password "" ignore read only recommended true add to mru false
                calculate full rebuild
                save workbook as wb filename firstPath file format Excel XML file format
                set wb to workbook "first.xlsx"
                if (full name of wb as text) is not firstPath then error "Unexpected reference workbook path"
                if (count of workbooks) is not 1 then error "Close existing workbooks before calculation capture"
                calculate full rebuild
                save workbook as wb filename secondPath file format Excel XML file format
                set wb to workbook "second.xlsx"
                if (full name of wb as text) is not secondPath then error "Unexpected reference workbook path"
                set excelVersion to version
                close wb saving no
                set display alerts to previousAlerts
                return excelVersion
            on error messageText number errorNumber
                try
                    set ownedPath to full name of wb as text
                    if ownedPath is inputPath or ownedPath is firstPath or ownedPath is secondPath then close wb saving no
                end try
                set display alerts to previousAlerts
                error messageText number errorNumber
            end try
        end tell
    end timeout
end run
'''
    success = False
    try:
        completed = subprocess.run(['osascript','-',str(stage/'input.xlsx'),str(stage/'first.xlsx'),str(stage/'second.xlsx')],input=script,text=True,capture_output=True,check=True,timeout=timeout+5)
        first = package((stage/'first.xlsx').read_bytes())
        second = package((stage/'second.xlsx').read_bytes())
        if [name for name,_,_ in sheets(first)] != [name for name,_,_ in sheets(parts)]:
            raise ValueError('Excel changed worksheet identity')
        values = results(first,targets)
        if values != results(second,targets):
            raise ValueError('Formula results changed between full recalculations')
        scored = [cell for cell in values if cell['value']['kind'] != 'empty']
        if not scored:
            raise ValueError('Excel produced no nonempty formula results')
        record = dict(schema_version=1,method=METHOD,engine='Microsoft Excel',version=completed.stdout.strip(),os=platform.mac_ver()[0],
            captured_at_utc=datetime.now(timezone.utc).isoformat(),source_sha256=hashlib.sha256(data).hexdigest(),
            input_sha256=hashlib.sha256(prepared).hexdigest(),formula_cells=count,empty_results=len(values)-len(scored),
            clear_cells=targets,cells=scored,sheets=[name for name,_,_ in sheets(parts)],
            verification='Caches cleared before opening; identical results after two full dependency rebuilds')
        output.parent.mkdir(parents=True,exist_ok=True)
        output.with_suffix('.xlsx').write_bytes(prepared)
        output.write_text(json.dumps(record,indent=2)+'\n')
        success = True
        return record
    except subprocess.CalledProcessError as error:
        raise ValueError(error.stderr or str(error)) from error
    finally:
        if success:
            shutil.rmtree(stage)


if __name__ == '__main__':
    parser = argparse.ArgumentParser(description='Record fresh Excel formula results from an idle local Excel instance.')
    parser.add_argument('source',type=Path)
    parser.add_argument('--out',type=Path,required=True)
    parser.add_argument('--timeout',type=int,default=60)
    args = parser.parse_args()
    if platform.system() != 'Darwin' or args.timeout <= 0:
        parser.error('macOS with Microsoft Excel and a positive timeout are required')
    record = capture(args.source,args.out,args.timeout)
    print(json.dumps({key:record[key] for key in ['source_sha256','formula_cells','empty_results','version']} | {'results':len(record['cells'])}))
