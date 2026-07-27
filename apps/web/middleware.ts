import { NextResponse, type NextRequest } from "next/server";
import { DEMO, DOCS, REPO, SITE } from "./app/content";
import { MARKDOWN_MEDIA_TYPE, homepageMarkdown } from "./app/markdown";
import { prefersMarkdown } from "../../shared/accept";

export const config = { matcher: "/" };

const LINKS = [
  `<${DOCS}>; rel="service-doc"`,
  `<${SITE}/llms.txt>; rel="describedby"; type="text/plain"`,
  `<${SITE}/index.md>; rel="alternate"; type="text/markdown"`,
  `<${DEMO}>; rel="related"`,
  `<${REPO}>; rel="related"`,
].join(", ");

export function middleware(request: NextRequest) {
  if (!prefersMarkdown(request.headers.get("accept"))) {
    const response = NextResponse.next();
    response.headers.set("Link", LINKS);
    response.headers.append("Vary", "Accept");
    return response;
  }

  const body = homepageMarkdown();
  return new NextResponse(body, {
    headers: {
      "Content-Type": MARKDOWN_MEDIA_TYPE,
      // Cloudflare's CDN keys on Accept-Encoding only, so a cacheable markdown
      // body at the same URL could be served to a browser. Rebuilding the
      // string per request is cheaper than that failure.
      "Cache-Control": "no-store",
      Link: LINKS,
      Vary: "Accept",
      "x-markdown-tokens": String(Math.ceil(body.length / 4)),
    },
  });
}
