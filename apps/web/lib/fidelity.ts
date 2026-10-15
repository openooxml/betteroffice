// the benchmark summary the Benchmarks action writes into the repository README;
// values stay the exact strings it published and are only parsed to size chart bars
export type FidelityRow = { label: string; ours: string; theirs: string };
export type FidelityFormat = {
  format: string;
  commit: string;
  commitUrl: string;
  libreoffice: string;
  rows: FidelityRow[];
};

const BEGIN = "<!-- BEGIN GENERATED VISUAL FIDELITY -->";
const END = "<!-- END GENERATED VISUAL FIDELITY -->";

const text = (html: string) => html.replace(/<[^>]+>/g, "").trim();

export function parseFidelity(readme: string): FidelityFormat[] {
  const start = readme.indexOf(BEGIN);
  const end = readme.indexOf(END);
  if (start < 0 || end < start) return [];
  const formats: FidelityFormat[] = [];
  for (const section of readme.slice(start, end).split(/^### /m).slice(1)) {
    const format = section.slice(0, section.indexOf("\n")).trim().toLowerCase();
    const table = section.match(/<table>([\s\S]*?)<\/table>/)?.[1];
    if (!table) continue;
    const head = table.match(/BetterOffice \(<a href="([^"]+)">([0-9a-f]+)<\/a>\)[\s\S]*?LibreOffice \(([^)]+)\)/);
    if (!head) continue;
    const rows: FidelityRow[] = [];
    for (const m of table.matchAll(/<tr><td>([\s\S]*?)<\/td><td[^>]*>([\s\S]*?)<\/td><td[^>]*>([\s\S]*?)<\/td><\/tr>/g)) {
      rows.push({ label: text(m[1]), ours: text(m[2]), theirs: text(m[3]) });
    }
    formats.push({ format, commit: head[2], commitUrl: head[1], libreoffice: head[3], rows });
  }
  return formats;
}

// next.config.ts inlines the README block at build time, so the page, /index.md and the
// markdown the middleware serves on "/" (no file system there) all read the same results
export function publishedFidelity(): FidelityFormat[] {
  return parseFidelity(process.env.BETTEROFFICE_BENCHMARKS ?? "");
}

export type FidelityBar = { who: string; ours: boolean; text: string; share: number };
export type FidelityChart = {
  row: string;
  label: string;
  hint: string;
  better: "higher" | "lower";
  group: "quality" | "time";
  bars: FidelityBar[];
};

const OFFICE: Record<string, string> = { docx: "Word", pptx: "PowerPoint", xlsx: "Excel" };
const NOUN: Record<string, string> = { docx: "document", pptx: "presentation", xlsx: "workbook" };

// the rows drawn as bars, in this order, quality before time; every row stays in the full
// table. Quality sits on a fixed 0 to 1 scale, times on the slower of the two
const CHARTS: {
  row: string;
  label: string;
  hint: (office: string, noun: string) => string;
  better: "higher" | "lower";
  group: "quality" | "time";
}[] = [
  { row: "SSIM", label: "Visual similarity", hint: (office) => `SSIM against ${office}, 0 to 1`, better: "higher", group: "quality" },
  { row: "Exact page counts", label: "Exact page counts", hint: (office, noun) => `${noun}s with ${office}'s page count`, better: "higher", group: "quality" },
  { row: "Recalc accuracy", label: "Formula accuracy", hint: (office) => `results matching ${office} after recalculation`, better: "higher", group: "quality" },
  { row: "Render time (avg)", label: "Render time", hint: (_, noun) => `average per ${noun}`, better: "lower", group: "time" },
  { row: "Recalc time (avg)", label: "Recalculation time", hint: (_, noun) => `average per ${noun}`, better: "lower", group: "time" },
];

// "63/63", "99.60%", "430 ms" and "0.8519" as numbers; anything else (a dash, a new unit) is not drawn
export function quantity(value: string): number | null {
  const plain = value.replace(/,/g, "");
  const ratio = plain.match(/^(\d+)\/(\d+)$/);
  if (ratio) return Number(ratio[2]) > 0 ? Number(ratio[1]) / Number(ratio[2]) : null;
  const number = plain.match(/^(\d+(?:\.\d+)?)\s*(%|ms)?$/);
  if (!number) return null;
  return number[2] === "%" ? Number(number[1]) / 100 : Number(number[1]);
}

export function fidelityCharts(f: FidelityFormat): FidelityChart[] {
  const charts: FidelityChart[] = [];
  for (const spec of CHARTS) {
    const row = f.rows.find((r) => r.label === spec.row);
    if (!row) continue;
    const ours = quantity(row.ours);
    const theirs = quantity(row.theirs);
    if (ours === null || theirs === null) continue;
    const full = spec.group === "quality" ? 1 : Math.max(ours, theirs);
    if (!(full > 0) || ours > full || theirs > full) continue;
    charts.push({
      row: spec.row,
      label: spec.label,
      hint: spec.hint(OFFICE[f.format] ?? "Office", NOUN[f.format] ?? "file"),
      better: spec.better,
      group: spec.group,
      bars: [
        { who: "BetterOffice", ours: true, text: row.ours, share: ours / full },
        { who: "LibreOffice", ours: false, text: row.theirs, share: theirs / full },
      ],
    });
  }
  return charts;
}

// one version line for the section when every format was measured on the same builds
export function sharedVersion(formats: FidelityFormat[]): Pick<FidelityFormat, "commit" | "commitUrl" | "libreoffice"> | null {
  const [first] = formats;
  if (!first || formats.some((f) => f.commit !== first.commit || f.libreoffice !== first.libreoffice)) return null;
  return { commit: first.commit, commitUrl: first.commitUrl, libreoffice: first.libreoffice };
}
