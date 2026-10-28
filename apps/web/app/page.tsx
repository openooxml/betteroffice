import type { Metadata } from "next";
import { FEED_URL, publishedPosts } from "../lib/blog";
import { liveCounts } from "../lib/home-jsonld";
import { Home } from "./home";

export const dynamic = "force-dynamic";

export const metadata: Metadata = {
  alternates: {
    canonical: "/",
    types: {
      "text/markdown": "/index.md",
      ...(publishedPosts().length > 0 ? { "application/atom+xml": FEED_URL } : {}),
    },
  },
};

export default async function Page() {
  return <Home downloads={await liveCounts()} />;
}
