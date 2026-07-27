import { NextResponse, type NextRequest } from "next/server";
import { prefersMarkdown } from "../../shared/accept";

export const config = { matcher: ["/", "/docs/:path*"] };

const WEBSITE = "https://betteroffice.dev";
const REPO = "https://github.com/openooxml/betteroffice";
const SITE = "https://docs.betteroffice.dev";

const LINKS = [
  `<${SITE}/docs>; rel="service-doc"`,
  `<${SITE}/llms.txt>; rel="describedby"; type="text/markdown"`,
  `<${SITE}/llms-full.txt>; rel="alternate"; type="text/markdown"`,
  `<${WEBSITE}>; rel="related"`,
  `<${REPO}>; rel="related"`,
].join(", ");

// fumadocs already renders every page as markdown, so negotiation rewrites to
// those routes instead of converting anything here.
function markdownPath(pathname: string): string {
  if (pathname === "/" || pathname === "/docs") return "/llms.mdx/docs/content.md";
  const slug = pathname.replace(/^\/docs\//, "").replace(/\/$/, "");
  return `/llms.mdx/docs/${slug}/content.md`;
}

export function middleware(request: NextRequest) {
  if (!prefersMarkdown(request.headers.get("accept"))) {
    const response = NextResponse.next();
    response.headers.set("Link", LINKS);
    return response;
  }

  const target = new URL(markdownPath(request.nextUrl.pathname), request.url);
  const response = NextResponse.rewrite(target);
  response.headers.set("Link", LINKS);
  response.headers.set("Vary", "Accept");
  return response;
}
