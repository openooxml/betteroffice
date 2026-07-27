import type { Metadata } from "next";
import { Suspense } from "react";
import { PptxDemoClient } from "./PptxDemoClient";

export const metadata: Metadata = {
  title: "PPTX",
  description:
    "Open and edit a PowerPoint deck in the browser — masters, layouts, shapes and collaborative editing on the BetterOffice Rust engine.",
  alternates: { canonical: "/pptx" },
};

export default function PptxDemo() {
  return (
    <Suspense fallback={null}>
      <PptxDemoClient />
    </Suspense>
  );
}
