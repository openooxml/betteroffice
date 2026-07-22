import type { Metadata } from "next";
import { Suspense } from "react";
import { XlsxDemoClient } from "./XlsxDemoClient";

export const metadata: Metadata = { title: "XLSX" };

export default function XlsxDemo() {
  return (
    <Suspense fallback={null}>
      <XlsxDemoClient />
    </Suspense>
  );
}
