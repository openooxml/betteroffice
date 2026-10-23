import { publishedTiers, type PublishedTier } from "../../../shared/eval-publication";
import { quantity, type FidelityBar } from "./fidelity";

export type AgentEvalRow = { format: string; ours: string; theirs: string };
export type AgentEvals = { ours: string; theirs: string; rows: AgentEvalRow[] };
export type AgentEvalCard = { format: string; bars: FidelityBar[] };

const SECTION = /^## Evals[ \t]*\r?$[\s\S]*?(?=^## |<!--|(?![\s\S]))/m;

// the README's evals section, up to the next section or marker; next.config.ts inlines it at build
export function agentEvalsSection(readme: string): string {
  return readme.match(SECTION)?.[0] ?? "";
}

export function parseAgentEvals(readme: string, allowlist: readonly PublishedTier[] = publishedTiers): AgentEvals | null {
  const section = agentEvalsSection(readme);
  const tables = [...section.matchAll(/^\|.*\|[ \t]*\r?$(?:\r?\n^\|.*\|[ \t]*\r?$)+/gm)];
  const table = tables.find(match => {
    const labels = [...section.slice(0, match.index).matchAll(/^(Basic|Complex) tier:/gim)];
    const tier = (labels.at(-1)?.[1]?.toLowerCase() ?? "basic") as PublishedTier;
    return allowlist.includes(tier);
  })?.[0];
  if (!table) return null;
  const cells = table.split(/\r?\n/).map((line) => line.trim().slice(1, -1).split("|").map((cell) => cell.replace(/\*\*|`/g, "").trim()));
  if (cells.some((row) => row.length !== 3) || cells[1].some((cell) => !/^:?-+:?$/.test(cell))) return null;
  const [header, , ...rows] = cells;
  const ours = header.findIndex((cell, index) => index > 0 && cell.startsWith("BetterOffice"));
  if (ours < 0) return null;
  const theirs = ours === 1 ? 2 : 1;
  return {
    ours: header[ours],
    theirs: header[theirs],
    rows: rows.map((row) => ({ format: row[0].toLowerCase(), ours: row[ours], theirs: row[theirs] })),
  };
}

export function publishedAgentEvals(): AgentEvals | null {
  return parseAgentEvals(process.env.BETTEROFFICE_AGENT_EVALS ?? "");
}

export function agentEvalCards(evals: AgentEvals | null): AgentEvalCard[] {
  const cards: AgentEvalCard[] = [];
  if (!evals) return cards;
  for (const row of evals.rows) {
    if (![row.ours, row.theirs].every((value) => /^\d+\/\d+$/.test(value))) continue;
    const ours = quantity(row.ours);
    const theirs = quantity(row.theirs);
    if (ours === null || theirs === null || ours > 1 || theirs > 1) continue;
    cards.push({
      format: row.format,
      bars: [
        { who: evals.ours, ours: true, text: row.ours, share: ours },
        { who: evals.theirs, ours: false, text: row.theirs, share: theirs },
      ],
    });
  }
  return cards;
}

export function agentEvalRuns(evals: AgentEvals | null): number {
  return agentEvalCards(evals).reduce((total, card) => total + Number(card.bars[0].text.split("/")[1]), 0);
}
