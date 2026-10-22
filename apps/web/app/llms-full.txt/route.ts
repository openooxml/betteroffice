import { publishedFidelity } from "../../lib/fidelity";
import { publishedAgentEvals } from "../../lib/agent-evals";
import { llmsFull, publishedPosts } from "../../lib/blog";
import { MARKDOWN_MEDIA_TYPE, homepageMarkdown } from "../markdown";

export const dynamic = "force-static";

export function GET() {
  const posts = publishedPosts();
  const body = llmsFull(homepageMarkdown(publishedFidelity(), publishedAgentEvals(), posts), posts);
  return new Response(body, {
    headers: {
      "Content-Type": MARKDOWN_MEDIA_TYPE,
      "Cache-Control": "public, max-age=0, s-maxage=3600",
      "x-markdown-tokens": String(Math.ceil(body.length / 4)),
    },
  });
}
