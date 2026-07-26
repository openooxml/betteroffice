/**
 * True when the client ranks markdown above HTML. A browser's
 * `text/html,application/xhtml+xml,...` therefore keeps getting HTML.
 */
export function prefersMarkdown(accept: string | null): boolean {
  if (!accept) return false;

  let markdown = -1;
  let html = -1;

  for (const part of accept.split(",")) {
    const [rawType, ...params] = part.split(";");
    const type = rawType.trim().toLowerCase();
    const q = params
      .map((param) => param.trim().toLowerCase())
      .find((param) => param.startsWith("q="));
    const weight = q ? Number.parseFloat(q.slice(2)) : 1;
    if (!Number.isFinite(weight) || weight <= 0) continue;

    if (type === "text/markdown") markdown = Math.max(markdown, weight);
    if (type === "text/html") html = Math.max(html, weight);
  }

  return markdown > 0 && markdown > html;
}
