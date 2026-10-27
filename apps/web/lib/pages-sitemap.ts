import type { MetadataRoute } from "next";
import { SITE } from "../../../shared/sites";
import { blogSitemap, publishedPosts } from "./blog";

export function pagesSitemap(): MetadataRoute.Sitemap {
  return [
    {
      url: SITE,
      changeFrequency: "weekly",
      priority: 1,
    },
    ...blogSitemap(publishedPosts()),
  ];
}
