import type { MetadataRoute } from "next";
import { liveFormats } from "../lib/formats";

const SITE = "https://demo.betteroffice.dev";

export default function sitemap(): MetadataRoute.Sitemap {
  return [
    { url: SITE, changeFrequency: "weekly", priority: 1 },
    ...liveFormats.map((format) => ({
      url: `${SITE}/${format.id}`,
      changeFrequency: "weekly" as const,
      priority: 0.8,
    })),
  ];
}
