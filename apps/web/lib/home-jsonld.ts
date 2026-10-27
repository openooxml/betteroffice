import { CRATES, DOCS, NPM, OPENOOXML, PYPI, REPO, SITE, SITE_DESCRIPTION } from "../app/content";
import { type CachedCount, type Registry } from "./download-stats";
import { readDownloads } from "./stats-cache";

export type DownloadCounts = Partial<Record<Registry, CachedCount | null>>;

const REGISTRIES: { key: Registry; name: string; url: string }[] = [
  { key: "npm", name: "npm", url: NPM },
  { key: "crates", name: "crates.io", url: CRATES },
  { key: "pypi", name: "PyPI", url: PYPI },
];

export async function liveCounts(read: (registry: Registry) => Promise<CachedCount | null> = readDownloads): Promise<DownloadCounts> {
  const [npm, crates, pypi] = await Promise.all(REGISTRIES.map(({ key }) => read(key)));
  return { npm, crates, pypi };
}

const counter = (name: string, description: string, count: number, service?: { name: string; url: string }) => ({
  "@type": "InteractionCounter",
  name,
  description,
  interactionType: { "@type": "DownloadAction" },
  userInteractionCount: count,
  ...(service ? { interactionService: { "@type": "WebSite", name: service.name, url: service.url } } : {}),
});

export function downloadStatistics(counts: DownloadCounts) {
  const present = REGISTRIES.filter(({ key }) => counts[key]);
  if (present.length === 0) return null;
  const total = present.reduce((sum, { key }) => sum + counts[key]!.downloads, 0);
  const newest = Math.max(...present.map(({ key }) => counts[key]!.at));
  return {
    dateModified: new Date(newest).toISOString(),
    interactionStatistic: [
      counter("Downloads in the last 30 days", `Combined ${present.map(({ name }) => name).join(", ")} downloads of BetterOffice packages over the last 30 days.`, total),
      ...present.map(({ key, name, url }) =>
        counter(`${name} downloads in the last 30 days`, `Downloads of BetterOffice packages from ${name} over the last 30 days.`, counts[key]!.downloads, { name, url }),
      ),
    ],
  };
}

export function softwareJsonLd(counts: DownloadCounts = {}) {
  return {
    "@context": "https://schema.org",
    "@type": "SoftwareApplication",
    name: "BetterOffice",
    url: SITE,
    image: `${SITE}/logo.svg`,
    applicationCategory: "BusinessApplication",
    operatingSystem: "Web",
    description: SITE_DESCRIPTION,
    offers: {
      "@type": "Offer",
      price: "0",
      priceCurrency: "USD",
    },
    license: "https://www.apache.org/licenses/LICENSE-2.0",
    softwareHelp: DOCS,
    creator: {
      "@type": "Organization",
      name: "OpenOOXML",
      url: OPENOOXML,
    },
    sameAs: [REPO, NPM, CRATES, PYPI, OPENOOXML],
    ...downloadStatistics(counts),
  };
}
