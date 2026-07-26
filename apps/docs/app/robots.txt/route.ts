import { siteUrl } from "@/lib/shared";

export const dynamic = "force-static";

const BODY = `# Content preferences: https://contentsignals.org
# ai-train: training or fine-tuning a generative model
# search:   building a search index and linking back
# ai-input: retrieving this page to ground a generated answer

User-agent: *
Content-Signal: ai-train=yes, search=yes, ai-input=yes
Allow: /

Sitemap: ${siteUrl}/sitemap.xml
`;

export function GET() {
  return new Response(BODY, {
    headers: {
      "Content-Type": "text/plain; charset=utf-8",
      "Cache-Control": "public, max-age=3600",
    },
  });
}
