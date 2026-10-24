"""Scores BO and Univer recalc against Excel per function family. Usage: score.py <work dir> <out dir>

Numbers match within max(1e-9, 1e-12 x magnitude) (office-quality xlsx_benchmark tolerances); text, booleans and error codes
must match exactly. A failed or timed-out workbook counts every one of its cells as wrong for that arm.
Volatile cells (NOW/TODAY/RAND/OFFSET/INDIRECT/CELL/INFO) and every cell of a time- or random-dependent workbook are
excluded from value scoring and only checked for "computed without an error the oracle does not have".
The exclusion scope additionally drops, from both arms, the cells prep marked with an exclusion reason (exclude.py):
GETPIVOTDATA, external-workbook references, and cells depending on those. Full-scope numbers are always printed too.
"""
import collections
import json
import math
import sys
from pathlib import Path

ARMS = ('bo', 'univer')
ABS, REL = 1e-9, 1e-12


def same(expected, actual):
    if not actual:
        return False
    if expected['kind'] == actual.get('kind') == 'number':
        return math.isclose(expected['value'], actual['value'], abs_tol=ABS, rel_tol=REL)
    if expected['kind'] == 'text' and actual.get('kind') == 'empty' and expected['value'] == '':
        return True
    return expected == {k: actual[k] for k in ('kind', 'value') if k in actual}


def main(work, out):
    fam = collections.defaultdict(lambda: collections.Counter())
    loaded = collections.defaultdict(lambda: collections.Counter())
    fam_x = collections.defaultdict(lambda: collections.Counter())
    loaded_x = collections.defaultdict(lambda: collections.Counter())
    reasons = collections.Counter()
    tagged = collections.defaultdict(lambda: collections.Counter())
    analysis_errors = []
    nocheck = collections.defaultdict(lambda: collections.Counter())
    books = []
    mismatches = {arm: collections.Counter() for arm in ARMS}
    names = {arm: collections.Counter() for arm in ARMS}
    examples = {arm: [] for arm in ARMS}
    for path in sorted(work.glob('*/*/expected.json')):
        root = path.parent
        expected = json.loads(path.read_text())
        runs = {}
        per_book = collections.Counter()
        for arm in ARMS:
            try:
                runs[arm] = json.loads((root / f'{arm}.values.json').read_text())
            except FileNotFoundError:
                runs[arm] = dict(status='missing', values=None)
        report = {}
        if (root / 'univer.report.json').exists():
            report = json.loads((root / 'univer.report.json').read_text())
        if expected.get('exclusionError'):
            analysis_errors.append((expected['id'], expected['exclusionError']))
        book = dict(id=expected['id'], set=expected['set'], cells=len(expected['cells']), timeRandom=expected['timeRandom'],
                    status={arm: runs[arm]['status'] for arm in ARMS}, correct=collections.Counter(), scored=0,
                    notLoadable=report.get('notLoadable'), errors={arm: runs[arm].get('error') for arm in ARMS if runs[arm].get('error')})
        for index, cell in enumerate(expected['cells']):
            volatile = expected['timeRandom'] or 'volatile' in cell['families']
            families = cell['families'] + ['all'] + [f"set:{expected['set']}"]
            excluded = cell.get('exclude')
            if excluded and not volatile:
                reasons[(excluded, expected['set'])] += 1
            for arm in ARMS:
                values = runs[arm].get('values')
                actual = values[index] if values else None
                kind = (actual or {}).get('kind')
                if volatile:
                    ok = actual is not None and (kind != 'error' or cell['value']['kind'] == 'error')
                    for f in families:
                        nocheck[f][f'{arm}:n'] += 1
                        nocheck[f][f'{arm}:ok'] += ok
                    continue
                ok = same(cell['value'], actual)
                computed = runs['univer']['status'] in ('ok', 'calc-timeout')
                if computed:
                    for f in families:
                        loaded[f][f'{arm}:ok'] += ok
                        loaded[f]['n'] += arm == 'bo'
                if not excluded:
                    for f in families:
                        fam_x[f][f'{arm}:ok'] += ok
                        fam_x[f]['n'] += arm == 'bo'
                        if computed:
                            loaded_x[f][f'{arm}:ok'] += ok
                            loaded_x[f]['n'] += arm == 'bo'
                    for t in (cell.get('tags') or []) + ([f"univer:{runs['univer']['status']}"] if not computed else []):
                        tagged[t][f'{arm}:ok'] += ok
                        tagged[t]['n'] += arm == 'bo'
                for f in families:
                    fam[f][f'{arm}:ok'] += ok
                    if kind == 'error' and cell['value']['kind'] != 'error':
                        fam[f][f'{arm}:err'] += 1
                        if actual.get('value') == '#NAME?':
                            fam[f][f'{arm}:name'] += 1
                    if actual is None:
                        fam[f][f'{arm}:nores'] += 1
                if ok:
                    book['correct'][arm] += 1
                else:
                    if kind == 'error' and actual.get('value') == '#NAME?':
                        for fn in cell['unknown'] or [x for x in cell['families']]:
                            names[arm][fn] += 1
                    mismatches[arm][','.join(cell['families'])] += 1
                    per_book[arm] += 1
                    if per_book[arm] <= 5 and len(examples[arm]) < 3000 and actual is not None:
                        examples[arm].append(dict(id=expected['id'], sheet=cell['name'], cell=cell['cell'], formula=cell['formula'][:160],
                                                  expected=cell['value'], actual=actual, tags=cell.get('tags'), exclude=excluded))
            if not volatile:
                book['scored'] += 1
                for f in families:
                    fam[f]['n'] += 1
        books.append(book)
    out.mkdir(parents=True, exist_ok=True)
    order = ['all', 'set:sp', 'set:mf', 'plain', 'math', 'statistical', 'logical', 'text', 'lookup', 'date', 'financial', 'cross-sheet',
             'dynamic-array', 'array-output', 'lambda-let', 'structured-ref', 'reference', 'other']
    lines = ['| family | cells | BO match | Univer match | BO errors (#NAME?) | Univer errors (#NAME?) | no result BO/Univer |', '|---|---|---|---|---|---|---|']
    pct = lambda a, b: f'{100 * a / b:.1f}%' if b else '-'
    for f in order + sorted(set(fam) - set(order)):
        c = fam.get(f)
        if not c or not c['n']:
            continue
        lines.append(f"| {f} | {c['n']} | {pct(c['bo:ok'], c['n'])} | {pct(c['univer:ok'], c['n'])} | {c['bo:err']} ({c['bo:name']}) | "
                     f"{c['univer:err']} ({c['univer:name']}) | {c['bo:nores']}/{c['univer:nores']} |")
    lines += ['', 'Only workbooks Univer converted and computed (same cells both arms):', '',
              '| family | cells | BO match | Univer match |', '|---|---|---|---|']
    for f in order + sorted(set(loaded) - set(order)):
        c = loaded.get(f)
        if c and c['n']:
            lines.append(f"| {f} | {c['n']} | {pct(c['bo:ok'], c['n'])} | {pct(c['univer:ok'], c['n'])} |")
    lines += ['', 'Exclusion scope: cells the converter cannot represent are dropped from both arms (GETPIVOTDATA, '
              'external-workbook references, and formula cells that reference them directly or through further '
              'references, defined names or table references).', '',
              'Excluded cells: ' + (', '.join(f'{r} ({s}) {n}' for (r, s), n in sorted(reasons.items())) or 'none')
              + (f'; exclusion analysis failed for {len(analysis_errors)} workbooks: ' + '; '.join(f'{i}: {e}' for i, e in analysis_errors[:5])
                 if analysis_errors else ''), '',
              '| family | cells | BO match | Univer match |', '|---|---|---|---|']
    for f in order + sorted(set(fam_x) - set(order)):
        c = fam_x.get(f)
        if c and c['n']:
            lines.append(f"| {f} | {c['n']} | {pct(c['bo:ok'], c['n'])} | {pct(c['univer:ok'], c['n'])} |")
    lines += ['', 'Exclusion scope, only workbooks Univer converted and computed (same cells both arms):', '',
              '| family | cells | BO match | Univer match |', '|---|---|---|---|']
    for f in order + sorted(set(loaded_x) - set(order)):
        c = loaded_x.get(f)
        if c and c['n']:
            lines.append(f"| {f} | {c['n']} | {pct(c['bo:ok'], c['n'])} | {pct(c['univer:ok'], c['n'])} |")
    lines += ['', 'Exclusion scope by cell tag (diagnostic, overlapping; univer:<status> = workbook Univer did not compute):', '',
              '| tag | cells | BO match | Univer match | Univer wrong |', '|---|---|---|---|---|']
    for t, c in sorted(tagged.items(), key=lambda kv: kv[1]['univer:ok'] - kv[1]['n']):
        lines.append(f"| {t} | {c['n']} | {pct(c['bo:ok'], c['n'])} | {pct(c['univer:ok'], c['n'])} | {c['n'] - c['univer:ok']} |")
    lines += ['', 'Volatile and time/random-dependent cells (no-error check only):', '',
              '| family | cells | BO no-error | Univer no-error |', '|---|---|---|---|']
    for f in ['all', 'volatile', 'set:sp', 'set:mf']:
        c = nocheck.get(f)
        if c and c['bo:n']:
            lines.append(f"| {f} | {c['bo:n']} | {pct(c['bo:ok'], c['bo:n'])} | {pct(c['univer:ok'], c['univer:n'])} |")
    status = {arm: collections.Counter(b['status'][arm] for b in books) for arm in ARMS}
    lines += ['', f'Workbooks: {len(books)} (sp {sum(b["set"] == "sp" for b in books)}, mf {sum(b["set"] == "mf" for b in books)}); '
              f'status BO {dict(status["bo"])}; Univer {dict(status["univer"])}',
              f'Workbooks with content not loadable in Univer OSS: {sum(1 for b in books if b["notLoadable"])}',
              '', 'Top #NAME? functions BO: ' + ', '.join(f'{k} {v}' for k, v in names['bo'].most_common(15)),
              'Top #NAME? functions Univer: ' + ', '.join(f'{k} {v}' for k, v in names['univer'].most_common(15)),
              'Top mismatch family sets BO: ' + '; '.join(f'{k} {v}' for k, v in mismatches['bo'].most_common(10)),
              'Top mismatch family sets Univer: ' + '; '.join(f'{k} {v}' for k, v in mismatches['univer'].most_common(10))]
    worst = sorted(books, key=lambda b: b['correct']['bo'] - b['correct']['univer'])
    lines += ['', 'Workbooks where BO trails Univer most (scored, BO ok, Univer ok):'] + [
        f"- {b['id']}: {b['scored']}, {b['correct']['bo']}, {b['correct']['univer']} ({b['status']['bo']}/{b['status']['univer']})" for b in worst[:12]
        if b['correct']['bo'] < b['correct']['univer']]
    (out / 'summary.md').write_text('# xlsx recalc accuracy: BO vs Univer vs Excel\n\n' + '\n'.join(lines) + '\n')
    (out / 'families.json').write_text(json.dumps(dict(families=fam, loaded=loaded, nocheck=nocheck, families_excl=fam_x, loaded_excl=loaded_x,
                                                       tags=tagged, excluded={f'{r}:{s}': n for (r, s), n in reasons.items()}), indent=1))
    (out / 'books.json').write_text(json.dumps(books, indent=0))
    (out / 'mismatches.json').write_text(json.dumps(examples, indent=0))
    print('\n'.join(lines))


if __name__ == '__main__':
    main(Path(sys.argv[1]), Path(sys.argv[2]))
