"""BO arm: the office-quality native recalc CLI (Workbook::open_for_read + recalculate_all + save) on each cache-cleared input.
Writes bo.values.json next to expected.json. Usage: bo_calc.py <xlsx-native-bench> <work dir> [workers]"""
import concurrent.futures as cf
import json
import subprocess
import sys
import time
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[3] / 'office-quality'))
from xlsx_calc import package, results, sheets  # noqa: E402


def run(binary, root):
    expected = json.loads((root / 'expected.json').read_text())
    out = root / 'bo.xlsx'
    record = dict(id=root.name, status='ok')
    start = time.perf_counter()
    try:
        proc = subprocess.run([binary, str(root / 'input.xlsx'), str(out)], capture_output=True, timeout=180, text=True)
        record['ms'] = round((time.perf_counter() - start) * 1000)
        if proc.returncode:
            raise RuntimeError(f'exit {proc.returncode}: {proc.stderr.strip()[:300]}')
        parts = package(out.read_bytes())
        if [name for name, _, _ in sheets(parts)] != expected['sheets']:
            raise RuntimeError('worksheet identity changed')
        values = [cell['value'] for cell in results(parts, [dict(sheet=c['sheet'], cell=c['cell']) for c in expected['cells']])]
    except subprocess.TimeoutExpired:
        record.update(status='timeout', ms=180_000)
        values = None
    except Exception as error:
        record.update(status='failed', error=str(error)[:300])
        values = None
    out.unlink(missing_ok=True)
    (root / 'bo.values.json').write_text(json.dumps(dict(record, values=values)))
    return record


if __name__ == '__main__':
    binary, work = sys.argv[1], Path(sys.argv[2])
    roots = sorted(p.parent for p in work.glob('*/*/expected.json'))
    with cf.ThreadPoolExecutor(int(sys.argv[3]) if len(sys.argv) > 3 else 3) as pool:
        records = list(pool.map(lambda r: run(binary, r), roots))
    bad = [r for r in records if r['status'] != 'ok']
    print(f'bo calc: {len(records) - len(bad)}/{len(records)} ok; ' + '; '.join(f"{r['id']}={r['status']}" for r in bad[:10]))
