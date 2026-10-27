import { type AgentEvals, parseAgentEvals } from "./agent-evals";
import { CHART_FONT_B64 } from "./chart-font";
import { type FidelityFormat, fidelityCharts, quantity } from "./fidelity";

export type Theme = "light" | "dark";

// the same values as the benchmarks site (apps/fidelity, shared/site-tokens.css): ours is --bo-acc, the other
// engine --lo = --bo-dim, and each bar sits on a --bo-line-soft track; text and rules use the ink tokens
const PALETTE = {
  light: { fg: "#0b0b0b", dim: "#6b6b6b", line: "#e4e4e4", ours: "#047857", other: "#6b6b6b", track: "#000000", trackOpacity: 0.065 },
  dark: { fg: "#f2f2f0", dim: "#8c8c8c", line: "#262626", ours: "#34d399", other: "#8c8c8c", track: "#ffffff", trackOpacity: 0.07 },
} as const;

// the landing page's Geist, embedded because an SVG loaded through <img> cannot fetch fonts
const FONT = "Geist, -apple-system, BlinkMacSystemFont, 'Segoe UI', 'Noto Sans', Helvetica, Arial, sans-serif";
const FONT_FACE = CHART_FONT_B64
  ? `<defs><style>@font-face{font-family:Geist;font-weight:100 900;src:url(data:font/woff2;base64,${CHART_FONT_B64}) format("woff2")}</style></defs>`
  : "";

// no side margin: content starts at x = 0 and ends at the canvas edge, so the image lines up with the README text
const WIDTH = 560;

export const themeOf = (value: string | null): Theme => (value === "dark" ? "dark" : "light");

const esc = (value: string) => value.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/"/g, "&quot;");
const percent = (share: number) => `${Math.round(share * 100)}%`;

function frame(width: number, height: number, title: string, body: string) {
  return `<svg xmlns="http://www.w3.org/2000/svg" width="${width}" height="${height}" viewBox="0 0 ${width} ${height}" role="img" aria-label="${esc(title)}" font-family="${FONT}">${FONT_FACE}<title>${esc(title)}</title>${body}</svg>\n`;
}

type Bar = { share: number; text: string; ours: boolean; basic?: boolean };

const rate = (row: string | undefined) => (row === undefined ? null : quantity(row));
const counts = (value: string) => value.split("/").map(Number) as [number, number];

export function evalsModel(section: string): string | null {
  return section.match(/`([^`]+)` via `codex exec`/)?.[1] ?? null;
}

function evalBars(basic: AgentEvals | null, complex: AgentEvals | null, format: string): Bar[] {
  const bars: Bar[] = [];
  for (const [evals, tier] of [[complex, "complex"], [basic, "basic"]] as const) {
    if (!evals) continue;
    let ours: number | null;
    let theirs: number | null;
    let oursText: string;
    let theirsText: string;
    if (format === "total") {
      const sum = (key: "ours" | "theirs") => evals.rows.reduce((t, r) => [t[0] + counts(r[key])[0], t[1] + counts(r[key])[1]], [0, 0]);
      const [op, on] = sum("ours");
      const [tp, tn] = sum("theirs");
      ours = on > 0 ? op / on : null;
      theirs = tn > 0 ? tp / tn : null;
      oursText = `${op}/${on}`;
      theirsText = `${tp}/${tn}`;
    } else {
      const row = evals.rows.find((r) => r.format === format);
      if (!row) continue;
      ours = rate(row.ours);
      theirs = rate(row.theirs);
      oursText = row.ours;
      theirsText = row.theirs;
    }
    if (theirs !== null) bars.push({ share: theirs, text: theirsText, ours: false, basic: tier === "basic" });
    if (ours !== null) bars.push({ share: ours, text: oursText, ours: true, basic: tier === "basic" });
  }
  return bars;
}

export function evalsHeadline(basic: AgentEvals | null, complex: AgentEvals | null, model: string | null = null): string {
  const parts: string[] = [];
  for (const [evals, tier] of [[basic, "basic"], [complex, "complex"]] as const) {
    if (!evals) continue;
    const sum = (key: "ours" | "theirs") => evals.rows.reduce((t, r) => [t[0] + counts(r[key])[0], t[1] + counts(r[key])[1]], [0, 0]);
    const [oursPass, oursRuns] = sum("ours");
    const [theirsPass, theirsRuns] = sum("theirs");
    if (oursRuns > 0 && theirsRuns > 0) parts.push(`${tier} ${percent(oursPass / oursRuns)} with BetterOffice MCP vs ${percent(theirsPass / theirsRuns)} with generic file tools`);
  }
  const subject = model ? `Agent edit pass rates with ${model}` : "Agent edit pass rates";
  return parts.length ? `${subject}: ${parts.join("; ")}.` : `${subject}.`;
}

export function evalsChartSvg(basic: AgentEvals | null, complex: AgentEvals | null, theme: Theme, model: string | null = null): string | null {
  const formats = [...new Set([...(basic?.rows ?? []), ...(complex?.rows ?? [])].map((row) => row.format)), "total"];
  const groups = formats.map((format) => ({ format, bars: evalBars(basic, complex, format) })).filter((group) => group.bars.length > 0);
  if (groups.length === 0) return null;
  const c = PALETTE[theme];
  const width = WIDTH;
  const left = 38;
  const top = 74;
  const plotHeight = 170;
  const height = top + plotHeight + 62;
  const groupWidth = (width - left) / groups.length;
  const perGroup = Math.max(...groups.map((group) => group.bars.length));
  const barWidth = Math.min(34, (groupWidth - 30) / perGroup - 6);
  const title = model ? `Agent edits with ${model}: pass rate by format` : "Agent edits: pass rate by format";
  let body = `<text x="0" y="14" font-size="15" font-weight="600" fill="${c.fg}">${esc(title)}</text>`;
  const legend = [
    { label: "Generic complex", fill: c.other, basic: false },
    { label: "BetterOffice complex", fill: c.ours, basic: false },
    { label: "Generic basic", fill: c.other, basic: true },
    { label: "BetterOffice basic", fill: c.ours, basic: true },
  ];
  legend.forEach((item, index) => {
    const x = (index % 2) * 190;
    const y = 26 + Math.floor(index / 2) * 18;
    body += `<rect x="${x + 0.75}" y="${y}" width="11" height="11" fill="${item.fill}" fill-opacity="${item.basic ? 0.2 : 1}" stroke="${item.fill}" stroke-width="1.5"/><text x="${x + 17}" y="${y + 10}" font-size="12" fill="${c.dim}">${item.label}</text>`;
  });
  for (const tick of [0, 0.5, 1]) {
    const y = top + plotHeight * (1 - tick);
    body += `<line x1="${left}" y1="${y}" x2="${width}" y2="${y}" stroke="${c.line}"/><text x="0" y="${y + 4}" font-size="11" fill="${c.dim}">${percent(tick)}</text>`;
  }
  groups.forEach((group, index) => {
    const start = left + groupWidth * index + (groupWidth - group.bars.length * (barWidth + 6) + 6) / 2;
    group.bars.forEach((bar, position) => {
      const h = Math.max(1, plotHeight * Math.min(1, bar.share));
      const bx = start + position * (barWidth + 6);
      const fill = bar.ours ? c.ours : c.other;
      body += `<rect x="${bx.toFixed(1)}" y="${(top + plotHeight - h).toFixed(1)}" width="${barWidth.toFixed(1)}" height="${h.toFixed(1)}" fill="${fill}" fill-opacity="${bar.basic ? 0.2 : 1}"${bar.basic ? ` stroke="${fill}" stroke-width="1.5"` : ""}/>`;
      body += `<text x="${(bx + barWidth / 2).toFixed(1)}" y="${(top + plotHeight - h - 5).toFixed(1)}" font-size="10" text-anchor="middle" fill="${c.fg}">${percent(bar.share)}</text>`;
    });
    body += `<text x="${(left + groupWidth * (index + 0.5)).toFixed(1)}" y="${top + plotHeight + 20}" font-size="13" font-weight="600" text-anchor="middle" fill="${c.fg}">${esc(group.format === "total" ? "Total" : group.format.toUpperCase())}</text>`;
  });
  body += `<text x="0" y="${height - 16}" font-size="11" fill="${c.dim}">Basic: single edits. Complex: multi-step tasks.</text>`;
  body += `<text x="0" y="${height - 4}" font-size="11" fill="${c.dim}">A run passes only if the edit is right and nothing else changed.</text>`;
  return frame(width, height, evalsHeadline(basic, complex, model), body);
}

// the headline rows the alt text quotes, in the order the tables carry them
function headlineRows(formats: FidelityFormat[], group: "quality" | "time") {
  return formats.flatMap((format) =>
    fidelityCharts(format)
      .filter((chart) => chart.group === group)
      .map((chart) => ({
        named: `${format.format.toUpperCase()} ${chart.row}`,
        ours: chart.bars.find((bar) => bar.ours)!,
        theirs: chart.bars.find((bar) => !bar.ours)!,
      })),
  );
}

const quote = (rows: ReturnType<typeof headlineRows>) => rows.map((row) => `${row.named} ${row.ours.text} vs ${row.theirs.text}`).join(", ");

export function benchmarksHeadline(formats: FidelityFormat[]): string {
  const quality = headlineRows(formats, "quality");
  const time = headlineRows(formats, "time");
  if (quality.length === 0 && time.length === 0) return "BetterOffice and LibreOffice scored against Microsoft Office.";
  const parts = [] as string[];
  if (quality.length) parts.push(`Scored against Microsoft Office, BetterOffice vs LibreOffice: ${quote(quality)}`);
  if (time.length) parts.push(`${quality.length ? "average" : "Average"} time per file, lower is better${quality.length ? "" : ", BetterOffice vs LibreOffice"}: ${quote(time)}`);
  return `${parts.join("; ")}.`;
}

export type TableRow = {
  format: string;
  label: string;
  better: "higher" | "lower";
  ours: { text: string; share: number | null };
  theirs: { text: string; share: number | null };
};

// rows that stay in the collapsed data tables only
const TABLE_ONLY = /^(scored\/total|parse success)$/i;
const lowerIsBetter = (label: string) => /error|time/i.test(label);
const isTime = (label: string) => /time/i.test(label);

// every other row the README publishes for each format, quality rows before timing rows, sized for a bar:
// ratios, percentages and SSIM sit on a fixed 0 to 1 scale, errors and times on the larger of the two
export function tableRows(formats: FidelityFormat[]): TableRow[] {
  return formats.flatMap((f) =>
    [...f.rows.filter((row) => !isTime(row.label)), ...f.rows.filter((row) => isTime(row.label))]
      .filter((row) => !TABLE_ONLY.test(row.label))
      .map((row) => {
        const ours = quantity(row.ours);
        const theirs = quantity(row.theirs);
        const better = lowerIsBetter(row.label) ? "lower" : "higher";
        const fixed = better === "higher" && ours !== null && theirs !== null && ours <= 1 && theirs <= 1;
        const full = fixed ? 1 : Math.max(ours ?? 0, theirs ?? 0);
        const share = (value: number | null) => (value === null ? null : full > 0 ? value / full : 0);
        return {
          format: f.format,
          label: row.label,
          better,
          ours: { text: row.ours, share: share(ours) },
          theirs: { text: row.theirs, share: share(theirs) },
        } satisfies TableRow;
      }),
  );
}

const shown = (row: TableRow) => row.label.replace(/ \(avg\)$/, "");

export function benchmarksChartSvg(formats: FidelityFormat[], theme: Theme): string | null {
  const rows = tableRows(formats);
  if (rows.length === 0) return null;
  const c = PALETTE[theme];
  const width = WIDTH;
  const barX = 204;
  const plot = width - barX - 62;
  const barHeight = 15;
  const pair = 2 * barHeight + 3;
  const pitch = pair + 11;
  let y = 52;
  let body = `<text x="0" y="14" font-size="16" font-weight="600" fill="${c.fg}">Scored against Microsoft Office</text>`;
  body += `<rect x="0.75" y="26" width="11" height="11" fill="${c.ours}"/><text x="17" y="36" font-size="12" fill="${c.dim}">BetterOffice</text><rect x="110.75" y="26" width="11" height="11" fill="${c.other}"/><text x="127" y="36" font-size="12" fill="${c.dim}">LibreOffice</text>`;
  let section = "";
  for (const row of rows) {
    if (row.format !== section) {
      section = row.format;
      body += `<text x="0" y="${y + 12}" font-size="13" font-weight="600" fill="${c.fg}">${esc(section.toUpperCase())}</text><line x1="${barX}" y1="${y + 8}" x2="${width}" y2="${y + 8}" stroke="${c.line}"/>`;
      y += 24;
    }
    const lower = row.better === "lower";
    body += `<g data-row="${esc(row.format)}:${esc(row.label)}"><text x="0" y="${y + (lower ? 14 : pair / 2 + 4)}" font-size="12" fill="${c.fg}">${esc(shown(row))}</text>`;
    if (lower) body += `<text x="0" y="${y + 27}" font-size="10" font-style="italic" fill="${c.dim}">lower is better</text>`;
    for (const [offset, bar, ours] of [[0, row.ours, true], [barHeight + 3, row.theirs, false]] as const) {
      body += `<rect x="${barX}" y="${y + offset}" width="${plot}" height="${barHeight}" fill="${c.track}" fill-opacity="${c.trackOpacity}"/>`;
      if (bar.share === null) {
        body += `<text x="${barX}" y="${y + offset + 12}" font-size="12" fill="${c.dim}">${esc(bar.text)}</text>`;
        continue;
      }
      const w = Math.max(1, plot * Math.min(1, bar.share));
      body += `<rect x="${barX}" y="${y + offset}" width="${w.toFixed(1)}" height="${barHeight}" fill="${ours ? c.ours : c.other}"/><text x="${(barX + w + 6).toFixed(1)}" y="${y + offset + 12}" font-size="11" fill="${c.fg}">${esc(bar.text)}</text>`;
    }
    body += "</g>";
    y += pitch;
  }
  body += `<text x="0" y="${y + 8}" font-size="11" fill="${c.dim}">Same scorer and Office references for both engines.</text>`;
  return frame(width, y + 16, benchmarksHeadline(formats), body);
}

export const publishedEvalTiers = (section: string) => ({
  basic: parseAgentEvals(section, ["basic"]),
  complex: parseAgentEvals(section, ["complex"]),
});
