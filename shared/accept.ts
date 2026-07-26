// Effective quality for one media type. The most specific matching range wins,
// so a wildcard range still ranks a type it covers. RFC 9110 12.5.1.
function quality(accept: string, type: string): number {
  const topLevel = `${type.split("/")[0]}/*`;
  let bestSpecificity = 0;
  let best = 0;

  for (const part of accept.split(",")) {
    const [rawRange, ...params] = part.split(";");
    const range = rawRange.trim().toLowerCase();

    const specificity =
      range === type ? 3 : range === topLevel ? 2 : range === "*/*" ? 1 : 0;
    if (specificity === 0) continue;

    const q = params
      .map((param) => param.trim().toLowerCase())
      .find((param) => param.startsWith("q="));
    const weight = q ? Number.parseFloat(q.slice(2)) : 1;
    if (!Number.isFinite(weight)) continue;

    if (
      specificity > bestSpecificity ||
      (specificity === bestSpecificity && weight > best)
    ) {
      bestSpecificity = specificity;
      best = weight;
    }
  }

  return best;
}

// Markdown only when it ranks strictly above HTML, so a browser's catch-all
// Accept keeps getting HTML.
export function prefersMarkdown(accept: string | null): boolean {
  if (!accept) return false;
  const markdown = quality(accept, "text/markdown");
  return markdown > 0 && markdown > quality(accept, "text/html");
}
