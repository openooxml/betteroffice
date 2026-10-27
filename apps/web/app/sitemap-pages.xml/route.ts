import { pagesSitemap } from "../../lib/pages-sitemap";
import { SITEMAP_HEADERS, urlsetXml } from "../../lib/sitemap-xml";

export const dynamic = "force-static";

export function GET() {
  return new Response(urlsetXml(pagesSitemap()), { headers: SITEMAP_HEADERS });
}
