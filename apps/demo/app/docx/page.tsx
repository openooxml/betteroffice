import type { Metadata } from "next";
import { DemoStage } from "../components/DemoStage";
import { getFormat } from "../../lib/formats";

export const metadata: Metadata = { title: "Docx" };

export default function DocxDemo() {
  return <DemoStage format={getFormat("docx")!} />;
}
