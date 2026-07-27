import type { Metadata } from "next";
import { Suspense } from "react";
import { DocxDemoClient } from "./DocxDemoClient";
import "@betteroffice/docx-react/styles.css";

export const metadata: Metadata = {
  title: "DOCX",
  description:
    "Open and edit a Word document in the browser — paginated layout, styles, tables and tracked changes on the BetterOffice Rust engine.",
  alternates: { canonical: "/docx" },
};

export default function DocxDemo() {
  return (
    <Suspense fallback={null}>
      <DocxDemoClient />
    </Suspense>
  );
}
