import { CRATES, DOCS, NPM, PYPI, REPO } from "../../../shared/sites";
export { BENCHMARKS, BENCHMARKS_SITE, CRATES, DEMO, DOCS, EVALS_SITE, NPM, OPENOOXML, PYPI, RELEASES, REPO, SITE } from "../../../shared/sites";

export const SITE_TITLE = "BetterOffice — DOCX, XLSX and PPTX editors";
export const SITE_DESCRIPTION =
  "Open-source DOCX, XLSX and PPTX editors for React, with Rust and WebAssembly cores, collaboration, and headless APIs. Apache-2.0.";

export const BLOG = {
  heading: "Blog",
  description: "Posts from the BetterOffice project.",
};

export const HERO = {
  title: "BetterOffice",
  tagline:
    "The open-source office suite. High-fidelity editing and real-time collaboration for people and agents, on Rust engines that run anywhere.",
};

export const ECOSYSTEMS = [
  {
    name: "JavaScript",
    registry: "npm",
    install: "npm install @betteroffice/docx-react",
    url: NPM,
    docs: `${DOCS}/docs/javascript`,
    desc: "Published React editors and framework-free cores for DOCX, XLSX and PPTX.",
  },
  {
    name: "Rust",
    registry: "crates.io",
    install: "cargo add betteroffice-docx",
    url: CRATES,
    docs: `${DOCS}/docs/rust`,
    desc: "The same engines natively, for servers, CLIs and agent pipelines.",
  },
  {
    name: "Python",
    registry: "PyPI",
    install: "pip install betteroffice-xlsx",
    url: PYPI,
    docs: `${DOCS}/docs/python`,
    desc: "Published DOCX, XLSX and PPTX packages, with installation and examples in the Python guide.",
  },
];

export const SUITE = {
  label: "Suite",
  heading: "One suite, four editors",
  prose:
    "Every editor edits the file itself, with no conversion on open or save.",
};

export const EDITORS = [
  {
    name: "Documents",
    format: "docx",
    demo: true,
    desc: "Redline a contract with your team while an agent proposes the next clause.",
    status: "available",
  },
  {
    name: "Spreadsheets",
    format: "xlsx",
    demo: true,
    desc: "Build a quarterly budget and watch every dependent formula update as you edit.",
    status: "available",
  },
  {
    name: "Slides",
    format: "pptx",
    demo: true,
    desc: "Rework a pitch deck in the browser and hand it back as a real .pptx.",
    status: "available",
  },
  {
    name: "Diagrams",
    format: "vsdx",
    demo: false,
    desc: "Update a process map's shape data and drop it into a Word or PowerPoint report.",
    status: "source preview",
  },
];

// a format's place in the suite, so other per-format lists follow the same order
export const editorOrder = (format: string) => EDITORS.findIndex((editor) => editor.format === format);

export const PACKAGES_SECTION = {
  label: "Packages",
  heading: "Ships as components, not iframes",
  prose:
    "The DOCX, XLSX and PPTX editors install from npm and render inside your app — no embeds, no external services, documents never leave the page.",
};

export const PACKAGES = [
  {
    name: "@betteroffice/agents",
    desc: "Tools for agents to search, read, edit and preview DOCX, XLSX and PPTX files, through a TypeScript SDK or a local MCP server.",
  },
  {
    name: "@betteroffice/docx",
    desc: "Framework-free DOCX core: parsing, CRDT editing, page layout and rendering on the Rust engine.",
  },
  {
    name: "@betteroffice/docx-react",
    desc: "The DOCX editor and viewer as a drop-in React component.",
  },
  {
    name: "@betteroffice/xlsx",
    desc: "Framework-free spreadsheet core: parsing, formula calculation and rendering on the Rust engine.",
  },
  {
    name: "@betteroffice/xlsx-react",
    desc: "The spreadsheet editor and viewer as a drop-in React component.",
  },
  {
    name: "@betteroffice/pptx",
    desc: "Framework-free slides core: parsing, editing and rendering on the Rust engine.",
  },
  {
    name: "@betteroffice/pptx-react",
    desc: "The slides editor and viewer as a drop-in React component.",
  },
];

export const FOUNDATION = {
  label: "Features",
  heading: "Features",
  prose:
    "Every BetterOffice editor runs on Rust engines, from opening and editing files to layout and rendering. The same engines run in your browser, on your servers and inside agent workflows.",
};

export const FEATURES_LINK = { label: "All features", href: `${DOCS}/docs` };

export const BENCHMARKS_SECTION = {
  label: "Benchmarks",
  heading: "Benchmarks",
  workflow: `${REPO}/actions/workflows/visual-fidelity.yml`,
  methodology: `${REPO}/blob/main/scripts/office-quality/README.md`,
  more: "Per-document results",
};

export const AGENT_EVALS = {
  heading: "Agent edits",
  lead: (tasks: number) => `Same model, same ${tasks} editing tasks, two toolsets. An edit counts only if the requested change is there and nothing else in the file moved.`,
  method: `${REPO}#evals`,
  more: "Method, tasks and fixtures",
};

export const CAPABILITIES = [
  {
    name: "Documents, spreadsheets, and slides",
    desc: "Open, edit, render, and save DOCX, XLSX and PPTX files with high fidelity. Native OOXML editing preserves untouched file parts losslessly when round-tripping.",
  },
  {
    name: "Agent editing with human review",
    desc: "Review attributed agent edits through tracked changes, inline diffs, and before-and-after previews. Accept or reject changes directly in the editor.",
  },
  {
    name: "Real-time collaboration",
    desc: "People and agents edit the same file together, with live cursors and selections. Concurrent changes merge automatically, and offline edits sync when peers reconnect.",
  },
  {
    name: "Undo and redo",
    desc: "Navigate editing history and undo accepted agent proposals as a single step.",
  },
  {
    name: "Embed or automate",
    desc: "Drop React editors into your app, build on the framework-free JavaScript cores, or use Rust and Python APIs for headless processing and agent workflows.",
  },
  {
    name: "Open source and self-hostable",
    desc: "Apache-2.0 licensed, with control over your document storage, deployment, and collaboration infrastructure.",
  },
];

export const COLLABORATION = {
  label: "Collaboration",
  heading: "People and agents, one document",
  prose:
    "The document itself is a CRDT: every editor — every person, every AI agent — is a peer on the same data structure, and concurrent edits merge in the engine. Agents don't get a sidebar; they get a cursor, with the same undo and the same tracked-changes attribution as any co-author.",
};

export const PEERS = [
  {
    name: "People",
    desc: "Live co-editing over any WebSocket relay. Offline edits converge on reconnect.",
  },
  {
    name: "Agents",
    desc: "An agent edits through the same operations as a person, and human review is suggesting mode — accept or reject tracked changes, not a diff dialog.",
  },
];

export const CODA = { heading: "The future is yours to edit." };
