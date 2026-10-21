import { defineCloudflareConfig } from "@opennextjs/cloudflare";
import staticAssetsIncrementalCache from "@opennextjs/cloudflare/overrides/incremental-cache/static-assets-incremental-cache";

// every page is prerendered at build and never revalidated, so the prerendered
// HTML (including the blog's generateStaticParams pages) is served from static assets
export default defineCloudflareConfig({ incrementalCache: staticAssetsIncrementalCache });
