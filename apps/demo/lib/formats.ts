export type FormatStatus = "live" | "soon";

export interface Format {
  id: "docx" | "xlsx" | "pptx";
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
    tagline: "Presentations — on the same engine stack.",
    status: "soon",
  },
];

export function getFormat(id: string): Format | undefined {
  return formats.find((f) => f.id === id);
}
