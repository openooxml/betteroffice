export type UsedByEntry = {
  name: string;
  href: string;
  // a file under public/logos with its viewBox trimmed to the ink, drawn monochrome in the text colour.
  // box is that viewBox's width and height; baseline (from the top) and xHeight are the wordmark's, in the
  // same units, so every wordmark gets one x-height and one baseline. Without a logo the name renders as text.
  logo?: { src: string; box: [number, number]; baseline: number; xHeight: number };
  enabled: boolean;
};

export const USED_BY_LABEL = "Used by";
export const USED_BY_MORE = "and others.";
// true turns the static row into the drifting ticker
export const USED_BY_ROTATE = false;

export const USED_BY: UsedByEntry[] = [
  {
    name: "OpenCode",
    href: "https://opencode.ai",
    logo: { src: "/logos/opencode.svg", box: [234, 42], baseline: 36, xHeight: 30 },
    enabled: true,
  },
  {
    name: "atira",
    href: "https://atira.ai",
    logo: { src: "/logos/atira.svg", box: [958, 253], baseline: 236.43, xHeight: 152.94 },
    enabled: true,
  },
  {
    name: "Ketryx",
    href: "https://www.ketryx.com",
    logo: { src: "/logos/ketryx.svg", box: [145.1, 41.25], baseline: 34.07, xHeight: 18.16 },
    enabled: true,
  },
];

// the band as one markdown line, shared by the markdown twin and checked against llms.txt
export function usedByLine(): string {
  const names = USED_BY.filter((entry) => entry.enabled).map((entry) => `[${entry.name}](${entry.href})`);
  return `${USED_BY_LABEL}: ${names.join(", ")} ${USED_BY_MORE}`;
}
