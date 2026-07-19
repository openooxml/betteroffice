import type { Metadata } from "next";
import { Suspense } from "react";
import { PptxDemoClient } from "./PptxDemoClient";
import "./pptx.css";

export const metadata: Metadata = { title: "PPTX" };

export default function PptxDemo() {
  return (
    <Suspense fallback={null}>
      <PptxDemoClient />
    </Suspense>
  );
}
