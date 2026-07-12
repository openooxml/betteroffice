import type { Metadata } from "next";
import { DemoStage } from "../components/DemoStage";
import { getFormat } from "../../lib/formats";

export const metadata: Metadata = { title: "Pptx" };

export default function PptxDemo() {
  return <DemoStage format={getFormat("pptx")!} />;
}
