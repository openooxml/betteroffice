import { MARKDOWN_MEDIA_TYPE, homepageMarkdown } from "../markdown";

export const dynamic = "force-static";

// Negotiating on `/` depends on the CDN keying by Accept, which Cloudflare's
// cache does not do. This URL is the cache-safe way to fetch the same markdown.
export function GET() {
  const body = homepageMarkdown();
  return new Response(body, {
    headers: {
      "Content-Type": MARKDOWN_MEDIA_TYPE,
      "Cache-Control": "public, max-age=0, s-maxage=3600",
      "x-markdown-tokens": String(Math.ceil(body.length / 4)),
    },
  });
}
