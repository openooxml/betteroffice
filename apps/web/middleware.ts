import { NextResponse, type NextRequest } from "next/server";
import { DEMO, DOCS, REPO, SITE } from "./app/content";
import { MARKDOWN_MEDIA_TYPE, homepageMarkdown } from "./app/markdown";
import { publishedFidelity } from "./lib/fidelity";
import { publishedAgentEvals } from "./lib/agent-evals";
import { blogMarkdownFor, blogPosts, blogTwin, publishedPosts } from "./lib/blog";
import { prefersMarkdown } from "../../shared/accept";

export const config = { matcher: ["/", "/blog", "/blog/:path*", "/blog.md", "/blog.txt"] };

const LINKS = [
  `<${DOCS}>; rel="service-doc"`,
  `<${SITE}/llms.txt>; rel="describedby"; type="text/plain"`,
  `<${SITE}/index.md>; rel="alternate"; type="text/markdown"`,
  `<${DEMO}>; rel="related"`,
  `<${REPO}>; rel="related"`,
].join(", ");

const NOINDEX = { "X-Robots-Tag": "noindex, nofollow" };

function negotiate(accept: string | null, links: string, markdown: () => string, draft = false): NextResponse {
  if (!prefersMarkdown(accept)) {
    const response = NextResponse.next();
    response.headers.set("Link", links);
    response.headers.append("Vary", "Accept");
    if (draft) response.headers.set("X-Robots-Tag", NOINDEX["X-Robots-Tag"]);
    return response;
  }

  const body = markdown();
  return new NextResponse(body, {
    headers: {
      ...(draft ? NOINDEX : {}),
      "Content-Type": MARKDOWN_MEDIA_TYPE,
      // Cloudflare's CDN keys on Accept-Encoding only, so a cacheable markdown
      // body at the same URL could be served to a browser. Rebuilding the
      // string per request is cheaper than that failure.
      "Cache-Control": "no-store",
      Link: links,
      Vary: "Accept",
      "x-markdown-tokens": String(Math.ceil(body.length / 4)),
    },
  });
}

export function middleware(request: NextRequest) {
  const { pathname } = request.nextUrl;
  const accept = request.headers.get("accept");
  if (pathname === "/") {
    return negotiate(accept, LINKS, () =>
      homepageMarkdown(publishedFidelity(), publishedAgentEvals(), publishedPosts()),
    );
  }

  const posts = blogPosts();
  // a twin has its own URL, so it caches like any static file
  const twin = blogTwin(pathname, posts);
  if (twin) {
    return new NextResponse(twin.body, {
      headers: {
        "Content-Type": twin.type,
        "Cache-Control": "public, max-age=0, s-maxage=3600",
        Link: `<${twin.canonical}>; rel="canonical"`,
        "x-markdown-tokens": String(Math.ceil(twin.body.length / 4)),
        ...(twin.draft ? NOINDEX : {}),
      },
    });
  }

  const page = blogMarkdownFor(pathname, posts);
  if (!page) return NextResponse.next();
  return negotiate(accept, `<${page.url}>; rel="alternate"; type="text/markdown"`, () => page.body, page.draft);
}
