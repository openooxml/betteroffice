import type { Metadata } from "next";
import { Suspense } from "react";
import { DocxDemoClient } from "./DocxDemoClient";
import "@betteroffice/docx-react/styles.css";
import "../collab/collab.css";
import "./docx.css";

export const metadata: Metadata = { title: "DOCX" };

export default function DocxDemo() {
  return (
    <Suspense fallback={null}>
      <DocxDemoClient />
    </Suspense>
  );
}
