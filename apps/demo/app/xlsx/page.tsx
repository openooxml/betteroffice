import type { Metadata } from "next";
import { Suspense } from "react";
import { XlsxDemoClient } from "./XlsxDemoClient";

export const metadata: Metadata = {
  title: "XLSX",
  description:
    "Open and edit an Excel workbook in the browser — formulas, recalculation, number formats and grid rendering on the BetterOffice Rust engine.",
  alternates: { canonical: "/xlsx" },
};

export default function XlsxDemo() {
  return (
    <Suspense fallback={null}>
      <XlsxDemoClient />
    </Suspense>
  );
}
