import {
  CAPABILITIES,
  COLLABORATION,
  DEMO,
  DOCS,
  ECOSYSTEMS,
  EDITORS,
  FOUNDATION,
  HERO,
  OPENOOXML,
  PACKAGES,
  PACKAGES_SECTION,
  PEERS,
  REPO,
  SITE,
  SUITE,
} from "./content";

export const MARKDOWN_MEDIA_TYPE = "text/markdown; charset=utf-8";

function named(items: { name: string; desc: string }[]): string {
  return items.map((item) => `- **${item.name}** — ${item.desc}`).join("\n");
}

export function homepageMarkdown(): string {
  const editors = EDITORS.map(
    (editor) =>
      `- **${editor.name}** (\`.${editor.format}\`) — ${editor.desc} [Demo](${DEMO}/${editor.format})`,
  ).join("\n");

  const packages = PACKAGES.map(
    (pkg) =>
      `- [\`${pkg.name}\`](https://www.npmjs.com/package/${pkg.name}) — ${pkg.desc}`,
  ).join("\n");

  const ecosystems = ECOSYSTEMS.map(
    (eco) =>
      `- **${eco.name}** (\`${eco.install}\`) — ${eco.desc} [${eco.registry}](${eco.url}), [guide](${eco.docs})`,
  ).join("\n");

  return `# ${HERO.title}

${HERO.tagline}

- [Demos](${DEMO})
- [Documentation](${DOCS})
- [GitHub](${REPO})
- [OpenOOXML](${OPENOOXML})
- [Full index for agents](${SITE}/llms.txt)

## ${SUITE.heading}

${SUITE.prose}

${editors}

## ${PACKAGES_SECTION.heading}

${PACKAGES_SECTION.prose}

${packages}

One install line per ecosystem:

${ecosystems}

## ${FOUNDATION.heading}

${FOUNDATION.prose}

${named(CAPABILITIES)}

## ${COLLABORATION.heading}

${COLLABORATION.prose}

${named(PEERS)}
`;
}
