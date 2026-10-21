import type { MetadataRoute } from "next";
import { SITE } from "../../../shared/sites";
import { blogSitemap, publishedPosts } from "../lib/blog";

export default function sitemap(): MetadataRoute.Sitemap {
  return [
    {
      url: SITE,
      changeFrequency: "weekly",
      priority: 1,
    },
    ...blogSitemap(publishedPosts()),
  ];
}
