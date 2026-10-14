import type { Metadata } from "next";
import { Suspense } from "react";
import { VsdxDemoClient } from "./VsdxDemoClient";

export const metadata: Metadata = {
  title: "VSDX",
  description:
    "Open Visio diagrams in the browser — pages, shapes, connectors and themes on the BetterOffice Rust engine.",
  alternates: { canonical: "/vsdx" },
};

export default function VsdxDemo() {
  return (
    <Suspense fallback={null}>
      <VsdxDemoClient />
    </Suspense>
  );
}
