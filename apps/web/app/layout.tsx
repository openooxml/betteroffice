import type { Metadata } from "next";
import "./globals.css";

const SITE = "https://betteroffice.dev";

export const metadata: Metadata = {
  metadataBase: new URL(SITE),
  title: {
    default: "BetterOffice — The open-source office suite",
    template: "%s — BetterOffice",
  },
  description:
    "BetterOffice is an open-source office suite by the OpenOOXML project — we build our own native OOXML engines in Rust, from the file format up.",
  openGraph: {
    type: "website",
    siteName: "BetterOffice",
    url: SITE,
  },
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
