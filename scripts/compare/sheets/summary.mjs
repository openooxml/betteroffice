// results.json -> summary.md. Timeouts rank as +inf and print as ">Ns".
import { readFileSync, writeFileSync } from 'node:fs';

const rank = (sample) => (sample && !sample.timedOut && sample.ms != null ? sample.ms : Infinity);
function pct(values, p) {
  if (!values.length) return null;
  const sorted = [...values].sort((a, b) => a - b);
  return sorted[Math.min(sorted.length - 1, Math.max(0, Math.ceil((p / 100) * sorted.length) - 1))];
}
const show = (v, ceil) => (v == null ? '-' : v === Infinity ? `>${ceil / 1000}s` : String(Math.round(v)));
const median = (values) => { const v = values.filter((x) => x != null && Number.isFinite(x)); return v.length ? pct(v, 50) : null; };
const fixed = (v, d = 1) => (v == null ? '-' : Number(v).toFixed(d));

export function summarize({ meta, sessions }) {
  const c = meta.ceilings;
  const lines = [`# sheets at scale: bo-xlsx vs univer-sheets`, '',
    `BO ${meta.boSha ?? '?'} | Univer ${meta.univer ?? '?'} | Chromium ${sessions.find((s) => s.chromium)?.chromium ?? '?'} | region ${meta.region ?? '?'} | nproc ${meta.nproc} | load ${meta.loadStart} -> ${meta.loadEnd ?? '?'} | ${meta.smoke ? 'SMOKE' : `n=${meta.reps} measured + ${meta.burners} burner per arm x workbook`}`,
    '', 'Univer opens an IWorkbookData snapshot of the same cells (the OSS build imports no xlsx), so its open times exclude file parsing. Recalc is triggered by each arm\'s API write to A1; paint = first frame matching the settled grid, value = the arm\'s own read of the last J (polled every 50 ms).', ''];
  const groups = new Map();
  for (const s of sessions.filter((x) => x.measured)) {
    const key = `${s.wb}|${s.arm}`;
    if (!groups.has(key)) groups.set(key, []);
    groups.get(key).push(s);
  }
  const header = (cols) => [`| wb | arm | ${cols.join(' | ')} |`, `|${'---|'.repeat(cols.length + 2)}`];
  const row = (wb, arm, cells) => `| ${wb} | ${arm} | ${cells.join(' | ')} |`;
  const dist = (samples, ceil) => {
    const v = samples.map(rank);
    const to = v.filter((x) => x === Infinity).length;
    return [show(pct(v, 50), ceil), show(pct(v, 95), ceil), show(pct(v, 99), ceil), show(pct(v, 100), ceil), `${to}/${v.length}`];
  };

  lines.push('## Open (ms from open start)', '', ...header(['first paint p50', 'p95', 'p99', 'max', 'TO', 'accepts input p50', 'p95', 'p99', 'max', 'TO', 'skipped/crashed']));
  for (const [key, list] of groups) {
    const [wb, arm] = key.split('|');
    lines.push(row(wb, arm, [...dist(list.map((s) => s.firstPaint), c.openMs), ...dist(list.map((s) => s.acceptsInput), c.openMs),
      `${list.filter((s) => s.skipped || s.crashed).length}`]));
  }
  lines.push('', '## Scroll (5 s wheel mid-sheet) and jump (Ctrl+Down to the last row)', '',
    ...header(['fps med', 'fps min', 'frame gap p95', 'rAF p50', 'rAF p95', 'rAF >50ms (sum)', 'longtasks n/sum/max', 'blank frames (sum)', 'jump p50', 'jump max', 'jump TO']));
  for (const [key, list] of groups) {
    const [wb, arm] = key.split('|');
    const sc = list.map((s) => s.scroll).filter(Boolean);
    const fps = sc.map((x) => x.fps);
    const lt = sc.map((x) => x.longtasks);
    const jumps = list.map((s) => s.jump ?? { timedOut: true });
    const jv = jumps.map(rank);
    lines.push(row(wb, arm, [fixed(median(fps)), fixed(fps.length ? Math.min(...fps) : null), fixed(median(sc.map((x) => x.frameGapP95)), 0),
      fixed(median(sc.map((x) => x.rafP50))), fixed(median(sc.map((x) => x.rafP95))), String(sc.reduce((a, x) => a + x.rafOver50, 0)),
      `${lt.reduce((a, x) => a + x.count, 0)}/${lt.reduce((a, x) => a + x.sumMs, 0)}/${lt.length ? Math.max(...lt.map((x) => x.maxMs)) : 0}`,
      `${sc.reduce((a, x) => a + x.blankFrames, 0)}/${sc.reduce((a, x) => a + x.frames, 0)}`,
      show(pct(jv, 50), c.jumpMs), show(pct(jv, 100), c.jumpMs), `${jv.filter((x) => x === Infinity).length}/${jv.length}`]));
  }
  lines.push('', `## Edit (5 digits + Enter, ${meta.edits} edits per rep) and recalc (A1 -> last J of a running-total chain)`, '',
    ...header(['key p50', 'key p95', 'key p99', 'key max', 'key TO', 'commit p50', 'commit p95', 'commit max', 'commit TO', 'recalc paint p50', 'paint max', 'paint TO', 'recalc value p50', 'value max', 'value TO', 'recalc correct']));
  for (const [key, list] of groups) {
    const [wb, arm] = key.split('|');
    const keys = list.flatMap((s) => (s.edits ?? []).flatMap((e) => e.keys));
    const commits = list.flatMap((s) => (s.edits ?? []).map((e) => e.commit));
    const k = dist(keys, c.keyMs), m = dist(commits, c.keyMs);
    const rc = list.map((s) => s.recalc ?? { timedOut: true });
    const rv = rc.map(rank);
    const vv = rc.map((x) => rank(x.value ?? { timedOut: true }));
    lines.push(row(wb, arm, [...k, m[0], m[1], m[3], m[4], show(pct(rv, 50), c.recalcMs), show(pct(rv, 100), c.recalcMs),
      `${rv.filter((x) => x === Infinity).length}/${rv.length}`, ...[50, 100].map((p) => show(pct(vv, p), c.recalcMs)),
      `${vv.filter((x) => x === Infinity).length}/${vv.length}`, `${rc.filter((x) => x.correct).length}/${rc.length}`]));
  }
  lines.push('', '## Memory (medians; PSS = all renderer processes of the measured browser; heap after forced GC)', '',
    ...header(['PSS settled MB', 'PSS after MB', 'tree PSS after MB', 'page heap settled', 'page heap after', 'worker heap settled', 'worker heap after']));
  for (const [key, list] of groups) {
    const [wb, arm] = key.split('|');
    const m = (f) => fixed(median(list.map(f)));
    lines.push(row(wb, arm, [m((s) => s.memSettled?.rendererPssMb), m((s) => s.memAfter?.rendererPssMb), m((s) => s.memAfter?.treePssMb),
      m((s) => s.memSettled?.heap?.pageMb), m((s) => s.memAfter?.heap?.pageMb), m((s) => s.memSettled?.heap?.workersMb), m((s) => s.memAfter?.heap?.workersMb)]));
  }
  const notes = [];
  for (const s of sessions) {
    const bad = (s.edits ?? []).filter((e) => !e.colOk || !e.committed).length;
    if (s.crashed || s.skipped || bad || (s.errors ?? []).length || (s.pageErrors ?? []).length || (s.recalc && !s.recalc.correct))
      notes.push(`- ${s.wb} rep${s.rep} ${s.arm}: ${[s.crashed && `crashed ${s.crashed}`, s.skipped && `skipped ${s.skipped}`, bad && `${bad} edits not committed or off A-H`,
        s.recalc && !s.recalc.correct && `recalc read ${JSON.stringify(s.recalc.read)} expected ${s.recalc.expected} (${s.recalc.setResult})`,
        (s.errors ?? []).length && `errors ${JSON.stringify(s.errors.slice(0, 2))}`, (s.pageErrors ?? []).length && `page ${JSON.stringify(s.pageErrors.slice(0, 2))}`].filter(Boolean).join('; ')}`);
  }
  if (notes.length) lines.push('', '## Anomalies', '', ...notes.slice(0, 60));
  return `${lines.join('\n')}\n`;
}

if (process.argv[1]?.endsWith('summary.mjs') && process.argv[2]) {
  const data = JSON.parse(readFileSync(process.argv[2], 'utf8'));
  writeFileSync(process.argv[3] ?? process.argv[2].replace(/results\.json$/, 'summary.md'), summarize(data));
}
