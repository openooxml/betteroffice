import { CRATES, DOCS, NPM, PYPI } from "../../../shared/sites";
export { BENCHMARKS, CRATES, DEMO, DOCS, NPM, OPENOOXML, PYPI, RELEASES, REPO, SITE } from "../../../shared/sites";

export const SITE_TITLE = "BetterOffice — DOCX, XLSX and PPTX editors";
export const SITE_DESCRIPTION =
  "Open-source DOCX, XLSX and PPTX editors for React, with Rust and WebAssembly cores, collaboration, and headless APIs. Apache-2.0.";

export const HERO = {
  title: "BetterOffice",
  tagline:
    "The open-source office suite. Word-faithful editing and real-time collaboration on engines we build ourselves — running entirely in your browser, by the OpenOOXML project.",
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
    "DOCX, XLSX and PPTX editors are published on npm and render inside your app. The VSDX diagram editor is available from source.",
};

export const EDITORS = [
  {
    name: "Documents",
    format: "docx",
    desc: "Word-faithful editing: fonts, theme colors, styles, tables, headers & footers, tracked changes.",
    status: "available",
  },
  {
    name: "Spreadsheets",
    format: "xlsx",
    desc: "Calculation graph, grid rendering and number formats on the same shared core.",
    status: "available",
  },
  {
    name: "Slides",
    format: "pptx",
    desc: "Slide model, masters and shape editing on the same shared core.",
    status: "available",
  },
  {
    name: "Diagrams",
    format: "vsdx",
    desc: "Source preview: edit diagrams with pan and zoom, update shape data in batches, and export to Word or PowerPoint through Rust. Controls the engine would refuse are disabled rather than offered.",
    status: "source preview",
  },
];

export const PACKAGES_SECTION = {
  label: "Packages",
  heading: "Ships as components, not iframes",
  prose:
    "The DOCX, XLSX and PPTX editors install from npm and render inside your app — no embeds, no external services, documents never leave the page. The same engines publish to crates.io for native Rust and to PyPI for Python.",
};

export const PACKAGES = [
  {
    name: "@betteroffice/docx",
    desc: "Framework-free .docx core — parsing, CRDT editing and page layout on the Rust engine, with version-checked atomic text and paragraph batches.",
  },
  {
    name: "@betteroffice/docx-react",
    desc: "The DOCX editor as a drop-in React component, with host-controlled saving, awaited input flushing, and version-checked edit batches.",
  },
  {
    name: "@betteroffice/xlsx",
    desc: "Framework-free spreadsheet core — parsing, calculation and rendering on the Rust engine, with opt-in operation timings.",
  },
  {
    name: "@betteroffice/xlsx-react",
    desc: "The spreadsheet editor as a drop-in React component.",
  },
  {
    name: "@betteroffice/pptx",
    desc: "Framework-free slides core — parsing, editing and rendering on the Rust engine, with opt-in operation timings, manual undo boundaries, and comment repositioning.",
  },
  {
    name: "@betteroffice/pptx-react",
    desc: "The slides editor as a drop-in React component, with host-controlled saving, awaited input flushing, and pointer position queries.",
  },
];

export const FOUNDATION = {
  label: "Features",
  heading: "Office editing for people and agents",
  prose:
    "OpenOOXML's Rust engines power every BetterOffice editor, from opening and editing files to layout and rendering. The same engines run in your browser and in headless workflows.",
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
    desc: "Navigate editing history and undo accepted agent proposals as a single step. DOCX hosts can set explicit boundaries and choose automatic or manual grouping.",
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
    desc: "Live co-editing over any WebSocket relay. Offline edits converge on reconnect — merging is the data structure, not a server feature.",
  },
  {
    name: "Agents",
    desc: "An agent edits through the same operations as a person, and human review is suggesting mode — accept or reject tracked changes, not a diff dialog.",
  },
];
