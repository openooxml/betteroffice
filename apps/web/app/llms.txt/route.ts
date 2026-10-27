import { llmsTxt, publishedPosts } from "../../lib/blog";
import { liveDownloadsLine, withDownloads } from "../../lib/llms-downloads";

// The download line reads STATS_KV on each request; the CDN caches the response.
export const dynamic = "force-dynamic";

export async function GET() {
  const body = withDownloads(llmsTxt(process.env.BETTEROFFICE_LLMS ?? "", publishedPosts()), await liveDownloadsLine());
  return new Response(body, {
    headers: {
      "Content-Type": "text/plain; charset=utf-8",
      "Cache-Control": "public, max-age=0, s-maxage=3600",
    },
  });
}
