import type { Metadata } from "next";
import { DemoStage } from "../components/DemoStage";
import { getFormat } from "../../lib/formats";

export const metadata: Metadata = { title: "Xlsx" };

export default function XlsxDemo() {
  return <DemoStage format={getFormat("xlsx")!} />;
}
