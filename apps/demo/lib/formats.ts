export type FormatStatus = "live" | "soon";

export interface Format {
  id: "docx" | "xlsx" | "pptx" | "vsdx";
  name: string;
  kind: string;
  tagline: string;
  status: FormatStatus;
}

export const formats: Format[] = [
  {
    id: "docx",
    name: "Docx",
    kind: ".docx documents",
    tagline: "Faithful .docx editing, rendered on our Rust engine to canvas.",
    status: "live",
  },
  {
    id: "xlsx",
    name: "Xlsx",
    kind: "Spreadsheets",
    tagline: "Cells, formulas, and rendering on a native Rust engine.",
    status: "live",
  },
  {
    id: "pptx",
    name: "Pptx",
    kind: "Slides",
    tagline: "Collaborative slides, shaped and rendered by a native Rust engine.",
    status: "live",
  },
  {
    id: "vsdx",
    name: "Vsdx",
    kind: ".vsdx diagrams",
    tagline: "VSDX support is in development.",
    status: "soon",
  },
];

export const liveFormats = formats.filter((format) => format.status === "live");

export function getFormat(id: string): Format | undefined {
  return liveFormats.find((format) => format.id === id);
}
