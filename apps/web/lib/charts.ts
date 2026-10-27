import { type AgentEvals, parseAgentEvals } from "./agent-evals";
import { type FidelityFormat, fidelityCharts, quantity } from "./fidelity";

export type Theme = "light" | "dark";

const PALETTE = {
  light: { bg: "#ffffff", fg: "#0b0b0b", dim: "#6b6b6b", line: "#e4e4e4", ours: "#047857", other: "#6b6b6b" },
  dark: { bg: "#090909", fg: "#f2f2f0", dim: "#8c8c8c", line: "#262626", ours: "#34d399", other: "#8c8c8c" },
} as const;

const FONT = "ui-sans-serif, system-ui, -apple-system, 'Segoe UI', Helvetica, Arial, sans-serif";

export const themeOf = (value: string | null): Theme => (value === "dark" ? "dark" : "light");

const esc = (value: string) => value.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/"/g, "&quot;");
const percent = (share: number) => `${Math.round(share * 100)}%`;

function frame(theme: Theme, width: number, height: number, title: string, body: string) {
  const c = PALETTE[theme];
  return `<svg xmlns="http://www.w3.org/2000/svg" width="${width}" height="${height}" viewBox="0 0 ${width} ${height}" role="img" aria-label="${esc(title)}" font-family="${FONT}"><title>${esc(title)}</title><rect width="${width}" height="${height}" fill="${c.bg}"/>${body}</svg>\n`;
}

type Bar = { share: number; text: string; ours: boolean; basic?: boolean };

const rate = (row: string | undefined) => (row === undefined ? null : quantity(row));

function evalBars(basic: AgentEvals | null, complex: AgentEvals | null, format: string): Bar[] {
  const bars: Bar[] = [];
  for (const [evals, tier] of [[complex, "complex"], [basic, "basic"]] as const) {
    const row = evals?.rows.find((r) => r.format === format);
    if (!evals || !row) continue;
    const theirs = rate(row.theirs);
    const ours = rate(row.ours);
    if (theirs !== null) bars.push({ share: theirs, text: row.theirs, ours: false, basic: tier === "basic" });
    if (ours !== null) bars.push({ share: ours, text: row.ours, ours: true, basic: tier === "basic" });
  }
  return bars;
}

export function evalsHeadline(basic: AgentEvals | null, complex: AgentEvals | null): string {
  const parts: string[] = [];
  for (const [evals, tier] of [[basic, "basic"], [complex, "complex"]] as const) {
    if (!evals) continue;
    const sum = (key: "ours" | "theirs") => evals.rows.reduce((t, r) => [t[0] + Number(r[key].split("/")[0]), t[1] + Number(r[key].split("/")[1])], [0, 0]);
    const [oursPass, oursRuns] = sum("ours");
    const [theirsPass, theirsRuns] = sum("theirs");
    if (oursRuns > 0 && theirsRuns > 0) parts.push(`${tier} ${percent(oursPass / oursRuns)} with BetterOffice MCP vs ${percent(theirsPass / theirsRuns)} with generic file tools`);
  }
  return parts.length ? `Agent edit pass rates: ${parts.join("; ")}.` : "Agent edit pass rates.";
}

export function evalsChartSvg(basic: AgentEvals | null, complex: AgentEvals | null, theme: Theme): string | null {
  const formats = [...new Set([...(basic?.rows ?? []), ...(complex?.rows ?? [])].map((row) => row.format))];
  const groups = formats.map((format) => ({ format, bars: evalBars(basic, complex, format) })).filter((group) => group.bars.length > 0);
  if (groups.length === 0) return null;
  const c = PALETTE[theme];
  const width = 720;
  const height = 360;
  const left = 56;
  const top = 64;
  const plotHeight = 220;
  const groupWidth = (width - left - 24) / groups.length;
  const perGroup = Math.max(...groups.map((group) => group.bars.length));
  const barWidth = Math.min(34, (groupWidth - 36) / perGroup - 6);
  let body = `<text x="${left}" y="28" font-size="15" font-weight="600" fill="${c.fg}">Agent edits: pass rate by format</text>`;
  const legend = [
    { label: "Generic complex", fill: c.other, basic: false },
    { label: "BetterOffice complex", fill: c.ours, basic: false },
    { label: "Generic basic", fill: c.other, basic: true },
    { label: "BetterOffice basic", fill: c.ours, basic: true },
  ];
  let x = left;
  for (const item of legend) {
    body += `<rect x="${x}" y="40" width="10" height="10" fill="${item.fill}" fill-opacity="${item.basic ? 0.2 : 1}" stroke="${item.fill}" stroke-width="1.5"/><text x="${x + 15}" y="49" font-size="11" fill="${c.dim}">${item.label}</text>`;
    x += 22 + item.label.length * 6.2 + 12;
  }
  for (const tick of [0, 0.5, 1]) {
    const y = top + plotHeight * (1 - tick);
    body += `<line x1="${left}" y1="${y}" x2="${width - 12}" y2="${y}" stroke="${c.line}"/><text x="${left - 8}" y="${y + 4}" font-size="10" text-anchor="end" fill="${c.dim}">${percent(tick)}</text>`;
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
    body += `<text x="${(left + groupWidth * (index + 0.5)).toFixed(1)}" y="${top + plotHeight + 22}" font-size="12" font-weight="600" text-anchor="middle" fill="${c.fg}">${esc(group.format.toUpperCase())}</text>`;
  });
  body += `<text x="${left}" y="${height - 14}" font-size="10" fill="${c.dim}">Basic: single edits. Complex: multi-step tasks. A run passes only if the edit is right and the rest of the file is untouched.</text>`;
  return frame(theme, width, height, evalsHeadline(basic, complex), body);
}

export function benchmarkRows(formats: FidelityFormat[]) {
  return formats.flatMap((format) =>
    fidelityCharts(format)
      .filter((chart) => chart.group === "quality")
      .map((chart) => ({
        title: `${format.format.toUpperCase()} ${chart.label.toLowerCase()}`,
        named: `${format.format.toUpperCase()} ${chart.row}`,
        ours: chart.bars.find((bar) => bar.ours)!,
        theirs: chart.bars.find((bar) => !bar.ours)!,
      })),
  );
}

export function benchmarksHeadline(formats: FidelityFormat[]): string {
  const rows = benchmarkRows(formats);
  if (rows.length === 0) return "BetterOffice and LibreOffice scored against Microsoft Office.";
  return `Scored against Microsoft Office, BetterOffice vs LibreOffice: ${rows.map((row) => `${row.named} ${row.ours.text} vs ${row.theirs.text}`).join(", ")}.`;
}

export function benchmarksChartSvg(formats: FidelityFormat[], theme: Theme): string | null {
  const rows = benchmarkRows(formats);
  if (rows.length === 0) return null;
  const c = PALETTE[theme];
  const width = 720;
  const left = 190;
  const plot = width - left - 70;
  const rowHeight = 62;
  const top = 64;
  const height = top + rows.length * rowHeight + 36;
  let body = `<text x="24" y="28" font-size="15" font-weight="600" fill="${c.fg}">BetterOffice vs LibreOffice, scored against Microsoft Office</text>`;
  body += `<rect x="24" y="40" width="10" height="10" fill="${c.ours}"/><text x="39" y="49" font-size="11" fill="${c.dim}">BetterOffice</text><rect x="120" y="40" width="10" height="10" fill="${c.other}"/><text x="135" y="49" font-size="11" fill="${c.dim}">LibreOffice</text>`;
  rows.forEach((row, index) => {
    const y = top + index * rowHeight;
    body += `<text x="24" y="${y + 24}" font-size="12" fill="${c.fg}">${esc(row.title)}</text>`;
    for (const [offset, bar] of [[0, row.ours], [22, row.theirs]] as const) {
      const w = Math.max(1, plot * Math.min(1, bar.share));
      body += `<rect x="${left}" y="${y + 6 + offset}" width="${w.toFixed(1)}" height="16" fill="${bar.ours ? c.ours : c.other}"/><text x="${(left + w + 8).toFixed(1)}" y="${y + 19 + offset}" font-size="11" fill="${c.fg}">${esc(bar.text)}</text>`;
    }
  });
  body += `<text x="24" y="${height - 14}" font-size="10" fill="${c.dim}">Higher is better. Both engines scored with the same scorer against the same Office references.</text>`;
  return frame(theme, width, height, benchmarksHeadline(formats), body);
}

export const publishedEvalTiers = (section: string) => ({
  basic: parseAgentEvals(section, ["basic"]),
  complex: parseAgentEvals(section, ["complex"]),
});
