import {
  CAPABILITIES,
  BENCHMARKS,
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
  RELEASES,
  SITE,
  SUITE,
} from "./content";

export const MARKDOWN_MEDIA_TYPE = "text/markdown; charset=utf-8";

function named(items: { name: string; desc: string }[]): string {
  return items.map((item) => `- **${item.name}.** ${item.desc}`).join("\n");
}

export function homepageMarkdown(): string {
  const editors = EDITORS.map(
    (editor) =>
      `- **${editor.name}** (\`.${editor.format}\`, ${editor.status}) — ${editor.desc} [Demo](${DEMO}/${editor.format})`,
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

[Package guide](${DOCS}/docs/packages) · [Release notes](${RELEASES})

## ${FOUNDATION.heading}

${FOUNDATION.prose}

${named(CAPABILITIES)}

[Try it out](${DEMO}), or [explore the docs](${DOCS}) for setup, APIs, and format support.

See how published releases and source builds compare with Microsoft Office in our [visual fidelity results](${BENCHMARKS}).

## ${COLLABORATION.heading}

${COLLABORATION.prose}

${named(PEERS)}

[Collaboration guide](${DOCS}/docs/collaboration)
`;
}
