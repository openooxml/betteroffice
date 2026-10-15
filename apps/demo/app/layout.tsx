import type { Metadata, Viewport } from "next";
import "./globals.css";

const SITE = "https://demo.betteroffice.dev";

export const metadata: Metadata = {
  metadataBase: new URL(SITE),
  title: {
    default: "BetterOffice — Demos",
    template: "%s — BetterOffice Demos",
  },
  description:
    "Live demos of the BetterOffice engines — Word documents, spreadsheets, and slides, rendered by native OOXML engines in Rust.",
};

export const viewport: Viewport = {
  themeColor: "#ffffff",
};

export default function RootLayout({
  children,
}: {
  children: React.ReactNode;
}) {
  return (
    <html lang="en">
      <body>{children}</body>
    </html>
  );
}
