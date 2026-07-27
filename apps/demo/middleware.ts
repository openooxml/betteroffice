import { NextResponse, type NextRequest } from "next/server";
import { prefersMarkdown } from "../../shared/accept";
import {
  DOCS,
  MARKDOWN_MEDIA_TYPE,
  REPO,
  SITE,
  WEBSITE,
  formatMarkdown,
  indexMarkdown,
} from "./lib/markdown";

export const config = { matcher: ["/", "/docx", "/xlsx", "/pptx"] };

const LINKS = [
  `<${DOCS}>; rel="service-doc"`,
  `<${WEBSITE}/llms.txt>; rel="describedby"; type="text/plain"`,
  `<${WEBSITE}>; rel="related"`,
  `<${REPO}>; rel="related"`,
].join(", ");

export function middleware(request: NextRequest) {
  if (!prefersMarkdown(request.headers.get("accept"))) {
    const response = NextResponse.next();
    response.headers.set("Link", LINKS);
    return response;
  }

  const path = request.nextUrl.pathname.replace(/^\/|\/$/g, "");
  const body = path ? formatMarkdown(path) : indexMarkdown();
  if (body === null) return NextResponse.next();

  return new NextResponse(body, {
    headers: {
      "Content-Type": MARKDOWN_MEDIA_TYPE,
      "Cache-Control": "no-store",
      Link: `${LINKS}, <${SITE}${request.nextUrl.pathname}>; rel="canonical"`,
      Vary: "Accept",
      "x-markdown-tokens": String(Math.ceil(body.length / 4)),
    },
  });
}
