import { formats, getFormat } from "./formats";

export const MARKDOWN_MEDIA_TYPE = "text/markdown; charset=utf-8";

export const SITE = "https://demo.betteroffice.dev";
export const WEBSITE = "https://betteroffice.dev";
export const DOCS = "https://docs.betteroffice.dev";
export const REPO = "https://github.com/openooxml/betteroffice";

const INTRO =
  "Live demos of the BetterOffice engines. Each editor opens a real file in the browser and runs on native OOXML engines written in Rust and compiled to WebAssembly. Nothing is uploaded; parsing, editing and rendering happen on the page.";

export function indexMarkdown(): string {
  const list = formats
    .map(
      (format) =>
        `- [${format.id.toUpperCase()}](${SITE}/${format.id}) — ${format.kind}. ${format.tagline}`,
    )
    .join("\n");

  return `# BetterOffice demos

${INTRO}

${list}

- [Website](${WEBSITE})
- [Documentation](${DOCS})
- [GitHub](${REPO})
`;
}

export function formatMarkdown(id: string): string | null {
  const format = getFormat(id);
  if (!format) return null;

  return `# BetterOffice ${format.id.toUpperCase()} demo

${format.kind}. ${format.tagline}

This page is an interactive editor, so the demo itself needs a browser. The
engine behind it also runs headless — see the documentation for the packages.

- [Open the demo](${SITE}/${format.id})
- [All demos](${SITE})
- [Documentation](${DOCS})
- [\`@betteroffice/${format.id}\`](https://www.npmjs.com/package/@betteroffice/${format.id}) — framework-free core
- [\`@betteroffice/${format.id}-react\`](https://www.npmjs.com/package/@betteroffice/${format.id}-react) — React editor
`;
}
