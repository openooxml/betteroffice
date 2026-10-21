import { llmsTxt, publishedPosts } from "../../lib/blog";

export const dynamic = "force-static";

export function GET() {
  const body = llmsTxt(process.env.BETTEROFFICE_LLMS ?? "", publishedPosts());
  return new Response(body, {
    headers: {
      "Content-Type": "text/plain; charset=utf-8",
      "Cache-Control": "public, max-age=0, s-maxage=3600",
    },
  });
}
