import type { MetadataRoute } from "next";
import { SITE } from "../../../shared/sites";

export default function sitemap(): MetadataRoute.Sitemap {
  return [
    {
      url: SITE,
      changeFrequency: "weekly",
      priority: 1,
    },
  ];
}
