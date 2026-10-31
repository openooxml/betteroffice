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
  return section.match(/Model: `([^`]+)`/)?.[1] ?? null;
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

// one legend row, one footnote line; basic tiers are translucent with a full-colour outline
export function evalsChartSvg(basic: AgentEvals | null, complex: AgentEvals | null, theme: Theme, model: string | null = null): string | null {
  const formats = [...new Set([...(basic?.rows ?? []), ...(complex?.rows ?? [])].map((row) => row.format)), "total"];
  const groups = formats.map((format) => ({ format, bars: evalBars(basic, complex, format) })).filter((group) => group.bars.length > 0);
  if (groups.length === 0) return null;
  const c = PALETTE[theme];
  const width = WIDTH;
  const left = 38;
  const top = 58;
  const plotHeight = 160;
  const height = top + plotHeight + 50;
  const groupWidth = (width - left) / groups.length;
  const perGroup = Math.max(...groups.map((group) => group.bars.length));
  const barWidth = Math.min(34, (groupWidth - 26) / perGroup - 5);
  const title = model ? `Agent edits with ${model}: pass rate by format` : "Agent edits: pass rate by format";
  let body = `<text x="0" y="13" font-size="13" font-weight="600" fill="${c.fg}">${esc(title)}</text>`;
  const legend = [
    { label: "Generic complex", fill: c.other, basic: false },
    { label: "BetterOffice complex", fill: c.ours, basic: false },
    { label: "Generic basic", fill: c.other, basic: true },
    { label: "BetterOffice basic", fill: c.ours, basic: true },
  ];
  let x = 0;
  for (const item of legend) {
    body += `<rect x="${x + 0.75}" y="24" width="11" height="11" fill="${item.fill}" fill-opacity="${item.basic ? 0.2 : 1}" stroke="${item.fill}" stroke-width="1.5"/><text x="${x + 16}" y="34" font-size="11.5" fill="${c.dim}">${item.label}</text>`;
    x += 16 + item.label.length * 6.1 + 12;
  }
  for (const tick of [0, 0.5, 1]) {
    const y = top + plotHeight * (1 - tick);
    body += `<line x1="${left}" y1="${y}" x2="${width}" y2="${y}" stroke="${c.line}"/><text x="0" y="${y + 4}" font-size="11" fill="${c.dim}">${percent(tick)}</text>`;
  }
  groups.forEach((group, index) => {
    const start = left + groupWidth * index + (groupWidth - group.bars.length * (barWidth + 5) + 5) / 2;
    group.bars.forEach((bar, position) => {
      const h = Math.max(1, plotHeight * Math.min(1, bar.share));
      const bx = start + position * (barWidth + 5);
      const fill = bar.ours ? c.ours : c.other;
      body += `<rect x="${bx.toFixed(1)}" y="${(top + plotHeight - h).toFixed(1)}" width="${barWidth.toFixed(1)}" height="${h.toFixed(1)}" fill="${fill}" fill-opacity="${bar.basic ? 0.2 : 1}"${bar.basic ? ` stroke="${fill}" stroke-width="1.5"` : ""}/>`;
      body += `<text x="${(bx + barWidth / 2).toFixed(1)}" y="${(top + plotHeight - h - 4).toFixed(1)}" font-size="11" text-anchor="middle" fill="${c.fg}">${percent(bar.share)}</text>`;
    });
    body += `<text x="${(left + groupWidth * (index + 0.5)).toFixed(1)}" y="${top + plotHeight + 20}" font-size="14" font-weight="600" text-anchor="middle" fill="${c.fg}">${esc(group.format === "total" ? "Total" : group.format.toUpperCase())}</text>`;
  });
  body += `<text x="0" y="${height - 6}" font-size="11" fill="${c.dim}">Basic: single edits. Complex: multi-step tasks. Pass: right edit, no other change.</text>`;
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
const isSsim = (label: string) => /^ssim$/i.test(label);

// every other row the README publishes for each format, SSIM first, then the other quality rows, then timing rows, sized for a bar:
// ratios, percentages and SSIM sit on a fixed 0 to 1 scale, errors and times on the larger of the two
export function tableRows(formats: FidelityFormat[]): TableRow[] {
  return formats.flatMap((f) =>
    [
      ...f.rows.filter((row) => isSsim(row.label)),
      ...f.rows.filter((row) => !isSsim(row.label) && !isTime(row.label)),
      ...f.rows.filter((row) => isTime(row.label)),
    ]
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

// slim bar pairs; a down arrow after the label marks lower-is-better rows, explained once in the legend
export function benchmarksChartSvg(formats: FidelityFormat[], theme: Theme): string | null {
  const rows = tableRows(formats);
  if (rows.length === 0) return null;
  const c = PALETTE[theme];
  const width = WIDTH;
  const barX = 190;
  const valueColumn = 60;
  const plot = width - barX - valueColumn;
  const barHeight = 8;
  const pair = 2 * barHeight + 2;
  const pitch = pair + 5;
  let y = 40;
  let body = `<text x="0" y="12" font-size="12.8" font-weight="600" fill="${c.fg}">Scored against Microsoft Office</text>`;
  body += `<rect x="0.75" y="22" width="11" height="11" fill="${c.ours}"/><text x="16.4" y="32" font-size="12" fill="${c.dim}">BetterOffice</text>`;
  body += `<rect x="108.75" y="22" width="11" height="11" fill="${c.other}"/><text x="124.4" y="32" font-size="12" fill="${c.dim}">LibreOffice</text>`;
  if (rows.some((row) => row.better === "lower")) body += `<text x="208.8" y="32" font-size="12" fill="${c.dim}">↓ lower is better</text>`;
  let section = "";
  for (const row of rows) {
    if (row.format !== section) {
      section = row.format;
      body += `<text x="0" y="${y + 12}" font-size="13" font-weight="600" fill="${c.fg}">${esc(section.toUpperCase())}</text><line x1="${barX}" y1="${y + 8}" x2="${width}" y2="${y + 8}" stroke="${c.line}"/>`;
      y += 18;
    }
    const arrow = row.better === "lower" ? ` <tspan fill="${c.dim}">↓</tspan>` : "";
    body += `<g data-row="${esc(row.format)}:${esc(row.label)}"><text x="0" y="${y + pair / 2 + 4}" font-size="12" fill="${c.fg}">${esc(shown(row))}${arrow}</text>`;
    for (const [offset, bar, ours] of [[0, row.ours, true], [barHeight + 2, row.theirs, false]] as const) {
      body += `<rect x="${barX}" y="${y + offset}" width="${plot}" height="${barHeight}" fill="${c.track}" fill-opacity="${c.trackOpacity}"/>`;
      if (bar.share !== null) {
        const w = Math.max(1, plot * Math.min(1, bar.share));
        body += `<rect x="${barX}" y="${y + offset}" width="${w.toFixed(1)}" height="${barHeight}" fill="${ours ? c.ours : c.other}"/>`;
      }
      body += `<text x="${width}" y="${y + offset + 8}" font-size="10.5" text-anchor="end" fill="${ours && bar.share !== null ? c.fg : c.dim}">${esc(bar.text)}</text>`;
    }
    body += "</g>";
    y += pitch;
  }
  body += `<text x="0" y="${y + 6}" font-size="11" fill="${c.dim}">Same scorer and Office references for both engines.</text>`;
  return frame(width, y + 12, benchmarksHeadline(formats), body);
}

export const publishedEvalTiers = (section: string) => ({
  basic: parseAgentEvals(section, ["basic"]),
  complex: parseAgentEvals(section, ["complex"]),
});
