import { FEED_TYPE, atomFeed, publishedPosts } from "../../../lib/blog";
import { postHtml } from "../../../lib/blog-html";

export const dynamic = "force-static";

export function GET() {
  const posts = publishedPosts();
  if (posts.length === 0) return new Response("Not found\n", { status: 404 });
  return new Response(atomFeed(posts, postHtml), {
    headers: {
      "Content-Type": FEED_TYPE,
      "Cache-Control": "public, max-age=0, s-maxage=3600",
    },
  });
}
