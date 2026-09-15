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
  heading: "One suite, three editors",
  prose:
    "DOCX, XLSX and PPTX editors are published on npm and render inside your app.",
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
    desc: "Slide model, masters and shape editing, with browser rendering of supported bitmap-only WMF images.",
    status: "available",
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
    desc: "Framework-free .docx core — parsing, CRDT editing and page layout in Rust, compiled to WebAssembly. Standard legacy horizontal rules render and survive editing and saving. Elliptical pictures render with crops and borders. Other picture presets retain rectangular rendering; soft-edge effects are unsupported. Hidden content is omitted by default and can be revealed through the render options.",
  },
  {
    name: "@betteroffice/docx-react",
    desc: "The full DOCX editor as a React component — toolbar, pages, comments, tracked changes.",
  },
  {
    name: "@betteroffice/xlsx",
    desc: "Framework-free spreadsheet core — parsing, calculation and rendering on the Rust engine.",
  },
  {
    name: "@betteroffice/xlsx-react",
    desc: "The spreadsheet editor as a drop-in React component.",
  },
  {
    name: "@betteroffice/pptx",
    desc: "Framework-free slides core — parsing, editing and rendering on the Rust engine. Custom browser image loaders can use presentationImageBlob for supported bitmap-only WMF wrappers.",
  },
  {
    name: "@betteroffice/pptx-react",
    desc: "The slides editor as a drop-in React component.",
  },
];

export const FOUNDATION = {
  label: "Foundation",
  heading: "Built on our own engines",
  prose:
    "BetterOffice is built by OpenOOXML, the open-source project writing native OOXML engines in Rust — parsing, layout, editing and rendering, from the file format up. Owning the whole stack is what makes the output Word-faithful.",
};

export const CAPABILITIES = [
  {
    name: "Own engines",
    desc: "We build the OOXML engines ourselves, in Rust — from the file format up. No wrapper around someone else's suite.",
  },
  {
    name: "Native OOXML editing",
    desc: "Documents are edited in their own format. No lossy conversion on open, none on save.",
  },
  {
    name: "Word-faithful output",
    desc: "What you see is what Word shows — layout, pagination and styling match the original.",
  },
  {
    name: "Real-time collaboration",
    desc: "The document is a CRDT — concurrent edits merge in the engine, not on a server.",
  },
  {
    name: "Agent-ready",
    desc: "Runs headless too — parse, edit and render documents server-side or inside agent pipelines.",
  },
  {
    name: "Apache 2.0",
    desc: "Permissive license, developed in the open, self-hostable without exceptions.",
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
