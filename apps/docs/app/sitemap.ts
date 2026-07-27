import type { MetadataRoute } from "next";
import { source } from "@/lib/source";
import { docsRoute, siteUrl } from "@/lib/shared";

export const revalidate = false;

export default function sitemap(): MetadataRoute.Sitemap {
  return source.getPages().map((page) => ({
    url: `${siteUrl}${page.url}`,
    changeFrequency: "weekly",
    priority: page.url === docsRoute ? 1 : 0.8,
  }));
}
