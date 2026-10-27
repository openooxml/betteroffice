import { SITEMAP_HEADERS, sitemapIndexXml } from "../../lib/sitemap-xml";

export const dynamic = "force-static";

export function GET() {
  return new Response(sitemapIndexXml(), { headers: SITEMAP_HEADERS });
}
