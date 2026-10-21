import { describe, expect, test } from "bun:test";
import { readFileSync } from "node:fs";
import { join } from "node:path";
import { createElement } from "react";
import { renderToStaticMarkup } from "react-dom/server";
import { NextRequest } from "next/server";
import { loadPosts, localImages, parsePost } from "./blog-source.ts";
import {
  BLOG_MARKDOWN_URL,
  BLOG_URL,
  FEED_URL,
  blogTwin,
  canonicalUrl,
  jsonLdScript,
  listed,
  llmsFull,
  llmsTxt,
  markdownUrl,
  parseBlog,
  postUrl,
  textUrl,
} from "./blog.ts";
import { postHtml } from "./blog-html.ts";
import { homepageMarkdown } from "../app/markdown.ts";

const FIXTURES = join(import.meta.dir, "__fixtures__", "blog");
const PUBLIC = join(import.meta.dir, "..", "public");
const LLMS_BASE = readFileSync(join(import.meta.dir, "..", "content", "llms.txt"), "utf8");
const ALL = parseBlog(JSON.stringify(loadPosts(FIXTURES, { drafts: true, publicDir: PUBLIC })));
const PUBLISHED = listed(ALL);
const DRAFT = ALL.find((post) => post.draft);

// what next.config.ts injects for a preview build of the fixtures
process.env.BETTEROFFICE_BLOG = JSON.stringify(loadPosts(FIXTURES, { drafts: true, publicDir: PUBLIC }));
process.env.BETTEROFFICE_LLMS = LLMS_BASE;

const withPosts = async (posts, run) => {
  const saved = process.env.BETTEROFFICE_BLOG;
  process.env.BETTEROFFICE_BLOG = JSON.stringify(posts);
  try {
    return await run();
  } finally {
    process.env.BETTEROFFICE_BLOG = saved;
  }
};

const ldBlocks = (html) =>
  [...html.matchAll(/<script type="application\/ld\+json">([\s\S]*?)<\/script>/g)].map((m) => JSON.parse(m[1]));

const minimal = (frontmatter, body = "Lorem ipsum.") => `---\n${frontmatter}\n---\n\n${body}\n`;
const BASE_FM = "title: Lorem\ndescription: Ipsum dolor.\ndate: 2026-01-02\nauthors: [Ada Lorem]";

describe("blog sources", () => {
  test("reads every fixture, drafts only when asked", () => {
    expect(ALL.map((post) => post.slug)).toEqual(["consectetur-draft", "lorem-ipsum", "dolor-sit-amet"]);
    expect(loadPosts(FIXTURES, { drafts: false, publicDir: PUBLIC }).map((post) => post.slug).sort()).toEqual(["dolor-sit-amet", "lorem-ipsum"]);
    expect(PUBLISHED.map((post) => post.slug)).toEqual(["lorem-ipsum", "dolor-sit-amet"]);
  });

  test("parses frontmatter into typed fields and keeps the file as written", () => {
    const post = ALL.find((p) => p.slug === "lorem-ipsum");
    expect(post).toMatchObject({
      title: "Lorem ipsum dolor sit amet & consectetur",
      date: "2026-10-01",
      updated: "2026-10-10",
      authors: [{ name: "Ada Lorem", url: "https://example.com/ada" }, { name: "Brutus Ipsum" }],
      tags: ["lorem", "ipsum"],
      draft: false,
    });
    expect(post.source).toBe(readFileSync(join(FIXTURES, "lorem-ipsum.md"), "utf8"));
    expect(post.body.startsWith("Lorem ipsum dolor sit amet")).toBe(true);
  });

  test("rejects posts the pages could not render correctly", () => {
    const bad = {
      "Bad_Slug.md": minimal(BASE_FM),
      "no-frontmatter.md": "Lorem ipsum.\n",
      "no-title.md": minimal("description: Ipsum.\ndate: 2026-01-02\nauthors: [Ada]"),
      "no-authors.md": minimal("title: Lorem\ndescription: Ipsum.\ndate: 2026-01-02"),
      "bad-date.md": minimal("title: Lorem\ndescription: Ipsum.\ndate: 02.01.2026\nauthors: [Ada]"),
      "impossible-date.md": minimal("title: Lorem\ndescription: Ipsum.\ndate: 2026-02-30\nauthors: [Ada]"),
      "early-update.md": minimal(`${BASE_FM}\nupdated: 2026-01-01`),
      "typo.md": minimal(`${BASE_FM}\ntitel: Lorem`),
      "relative-canonical.md": minimal(`${BASE_FM}\ncanonical: /elsewhere`),
      "empty.md": minimal(BASE_FM, ""),
    };
    for (const [file, source] of Object.entries(bad)) expect(() => parsePost(file, source)).toThrow(`blog post ${file}`);
    expect(parsePost("fine.md", minimal(BASE_FM)).slug).toBe("fine");
    expect(() => localImages(parsePost("pic.md", minimal(BASE_FM, "![x](/missing.png)")), PUBLIC)).toThrow("not in public/");
    const examples = parsePost("code.md", minimal(BASE_FM, "```md\n![x](/missing.png)\n```\n\nUse `![x](/missing.png)` for images."));
    expect(localImages(examples, PUBLIC)).toEqual({});
    const hidden = [
      "    ![x](/missing.png)",
      "   ```\n![x](/missing.png)\n   ```",
      "- item\n\n      ![x](/missing.png)",
      "<!-- ![x](/missing.png) -->",
      "Lorem <!-- <img src=\"/missing.png\"> --> ipsum.",
      "<!--\n![x](/missing.png)\n-->",
      "[unused]: /missing.png",
      "![x](//cdn.example.com/x.png) ![y](https://example.com/y.png)",
    ];
    for (const body of hidden) expect(localImages(parsePost("hidden.md", minimal(BASE_FM, body)), PUBLIC)).toEqual({});
    const reference = parsePost("ref.md", minimal(BASE_FM, "![Logo][logo] and ![Logo]\n\n[logo]: /logo.svg"));
    expect(localImages(reference, PUBLIC)).toEqual({ "/logo.svg": [276, 206] });
    expect(postHtml({ ...reference, images: localImages(reference, PUBLIC) })).toContain('<img src="/logo.svg" alt="Logo" width="276" height="206"');
    const missingReference = parsePost("gone.md", minimal(BASE_FM, "![x][gone]\n\n[gone]: /missing.png"));
    expect(() => localImages(missingReference, PUBLIC)).toThrow("image /missing.png is not in public/");
    const rawImage = parsePost("raw-img.md", minimal(BASE_FM, "<p><IMG SRC='/missing.png'></p>"));
    expect(() => localImages(rawImage, PUBLIC)).toThrow("not in public/");
  });

  test("keeps the indentation of the first body line", () => {
    expect(parsePost("code.md", minimal(BASE_FM, "    indented code")).body).toBe("    indented code\n");
  });

  test("renders the body with one heading level below the page title", () => {
    const html = postHtml(ALL.find((p) => p.slug === "lorem-ipsum"));
    expect(html).toContain('<h2 id="duis-aute-irure">Duis aute irure</h2>');
    expect(html).toContain('<h3 id="nemo-enim-ipsam">Nemo enim ipsam</h3>');
    expect(html).toContain('<div class="post-table"><table>');
    expect(html).toContain('<img src="/logo.svg" alt="Lorem ipsum" width="276" height="206" loading="lazy" decoding="async" />');
    const raw = parsePost("raw.md", minimal(BASE_FM, '<img src="/logo.svg" alt="A > B">'));
    expect(postHtml({ ...raw, images: { "/logo.svg": [276, 206] } })).toContain(
      '<img src="/logo.svg" alt="A > B" width="276" height="206" loading="lazy" decoding="async" />',
    );
    expect(html).not.toContain("<h1");
    const titled = parsePost("titled.md", minimal(BASE_FM, "# Lorem\n\nIpsum."));
    expect(() => postHtml(titled)).toThrow("only h1");
    for (const body of ['<H1 class="lead">Lorem</H1>', "<h1\nclass=x>Lorem</h1>", "Lorem <h1>ipsum</h1>"]) {
      expect(() => postHtml(parsePost("raw-h1.md", minimal(BASE_FM, body)))).toThrow("only h1");
    }
    const note = postHtml(parsePost("note.md", minimal(BASE_FM, "<!-- <h1>draft title</h1> -->\n\nIpsum.")));
    expect(note).toContain("<p>Ipsum.</p>");
    expect(note).not.toContain("draft title");
    const quoted = postHtml(parsePost("quoted.md", minimal(BASE_FM, '<div title="<!--">Visible</div>\n\n## Remaining')));
    expect(quoted).toContain('<div title="<!--">Visible</div>');
    expect(quoted).toContain('<h2 id="remaining">Remaining</h2>');
    const idsOf = (body) => [...postHtml(parsePost("ids.md", minimal(BASE_FM, body))).matchAll(/\sid=["']([^"']+)/gi)].map((m) => m[1]);
    expect(idsOf("## Title\n\n## Setup\n\n## Setup\n\n## Setup 2")).toEqual(["title-2", "setup", "setup-2", "setup-2-2"]);
    expect(idsOf('<h2 ID="setup">Setup</h2>\n\n## Setup\n\n<div id=\'usage\'></div>\n\n## Usage')).toEqual(["setup", "setup-2", "usage", "usage-2"]);
    expect(postHtml(parsePost("attrs.md", minimal(BASE_FM, '<H2 class="lead">Intro</H2>')))).toContain('<H2 class="lead" id="intro">Intro</H2>');
  });
});

describe("production content", () => {
  test("the lorem fixtures are test-only: the site's content dir carries none of them", () => {
    const production = loadPosts(join(import.meta.dir, "..", "content", "blog"), { drafts: true, publicDir: PUBLIC });
    const fixtures = new Set(ALL.flatMap((post) => [post.slug, post.title]));
    expect(production.filter((post) => fixtures.has(post.slug) || fixtures.has(post.title))).toEqual([]);
    expect(readFileSync(join(import.meta.dir, "..", "next.config.ts"), "utf8")).toContain('?? "content/blog"');
  });

  test("with no published post the blog stays dark", async () => {
    const { default: BlogIndex } = await import("../app/blog/page.tsx");
    const { GET } = await import("../app/llms.txt/route.ts");
    await withPosts([DRAFT], async () => {
      expect(() => BlogIndex()).toThrow();
      expect(await GET().text()).not.toContain("## Blog");
      expect(blogTwin("/blog.md", [DRAFT])).toBeNull();
    });
    expect(homepageMarkdown([], null, [])).not.toContain("/blog");
  });
});

describe("text and markdown twins", () => {
  test("every post, drafts included in preview, has a .md and a .txt twin", () => {
    for (const post of ALL) {
      const md = blogTwin(`/blog/${post.slug}.md`, ALL);
      expect(md).toEqual({ body: post.source, type: "text/markdown; charset=utf-8", canonical: canonicalUrl(post), draft: post.draft });
      const txt = blogTwin(`/blog/${post.slug}.txt`, ALL);
      expect(txt.type).toBe("text/plain; charset=utf-8");
      expect(txt.body.startsWith(`# ${post.title}\n\n> ${post.description}\n`)).toBe(true);
      expect(txt.body).toContain(post.body.trim());
    }
    expect(blogTwin("/blog/missing.md", ALL)).toBeNull();
    expect(blogTwin("/blog/lorem-ipsum.html", ALL)).toBeNull();
  });

  test("the index twins list published posts newest first, with no drafts", () => {
    const md = blogTwin("/blog.md", ALL).body;
    const at = PUBLISHED.map((post) => md.indexOf(`(${markdownUrl(post)})`));
    expect(at.every((index) => index > 0)).toBe(true);
    expect(at).toEqual([...at].sort((a, b) => a - b));
    expect(md).not.toContain(DRAFT.slug);
    const txt = blogTwin("/blog.txt", ALL).body;
    for (const post of PUBLISHED) expect(txt).toContain(textUrl(post));
    expect(txt).not.toContain(DRAFT.slug);
    expect(blogTwin("/blog.md", [DRAFT])).toBeNull();
  });

  test("the middleware serves twins and negotiates markdown on blog pages", async () => {
    const { middleware } = await import("../middleware.ts");
    const post = PUBLISHED[0];
    const twin = middleware(new NextRequest(`${postUrl(post)}.md`));
    expect(twin.status).toBe(200);
    expect(twin.headers.get("content-type")).toBe("text/markdown; charset=utf-8");
    expect(twin.headers.get("link")).toBe(`<${postUrl(post)}>; rel="canonical"`);
    expect(await twin.text()).toBe(post.source);

    const text = middleware(new NextRequest(`${postUrl(post)}.txt`));
    expect(text.headers.get("content-type")).toBe("text/plain; charset=utf-8");

    const negotiated = middleware(new NextRequest(postUrl(post), { headers: { accept: "text/markdown" } }));
    expect(negotiated.headers.get("vary")).toBe("Accept");
    expect(negotiated.headers.get("cache-control")).toBe("no-store");
    expect(await negotiated.text()).toBe(post.source);

    const html = middleware(new NextRequest(BLOG_URL, { headers: { accept: "text/html,*/*;q=0.8" } }));
    expect(html.headers.get("link")).toBe(`<${BLOG_MARKDOWN_URL}>; rel="alternate"; type="text/markdown"`);
    expect(html.headers.get("x-middleware-next")).toBe("1");
    for (const response of [twin, text, negotiated, html]) expect(response.headers.get("x-robots-tag")).toBeNull();

    const drafts = [
      middleware(new NextRequest(`${postUrl(DRAFT)}.md`)),
      middleware(new NextRequest(`${postUrl(DRAFT)}.txt`)),
      middleware(new NextRequest(postUrl(DRAFT), { headers: { accept: "text/markdown" } })),
      middleware(new NextRequest(postUrl(DRAFT), { headers: { accept: "text/html" } })),
    ];
    expect(await drafts[2].text()).toBe(DRAFT.source);
    for (const response of drafts) expect(response.headers.get("x-robots-tag")).toBe("noindex, nofollow");
  });
});

describe("indexes for agents and crawlers", () => {
  test("llms.txt lists each published post's markdown, and nothing without posts", async () => {
    const llms = llmsTxt(LLMS_BASE, PUBLISHED);
    for (const post of PUBLISHED) expect(llms).toContain(`- [${post.title}](${markdownUrl(post)}): ${post.description}`);
    expect(llms).toContain(`(${BLOG_MARKDOWN_URL})`);
    expect(llms).not.toContain(DRAFT.slug);
    expect(llms.indexOf("## Blog")).toBeLessThan(llms.indexOf("## Links"));
    expect(llmsTxt(LLMS_BASE, [])).toBe(LLMS_BASE);

    const { GET } = await import("../app/llms.txt/route.ts");
    expect(await GET().text()).toBe(llms);
    expect(await withPosts([DRAFT], async () => (await GET()).text())).toBe(LLMS_BASE);
  });

  test("llms-full.txt carries the landing page and every published post in full", async () => {
    const { GET } = await import("../app/llms-full.txt/route.ts");
    const full = await GET().text();
    expect(full.startsWith("# BetterOffice")).toBe(true);
    for (const post of PUBLISHED) expect(full).toContain(post.body.trim());
    expect(full).not.toContain(DRAFT.body.trim());
    expect(llmsFull("# Home\n", [])).toBe("# Home");
  });

  test("the landing page markdown links the blog only when a post is published", () => {
    expect(homepageMarkdown([], null, PUBLISHED)).toContain(`- [Blog](${BLOG_MARKDOWN_URL})`);
    expect(homepageMarkdown([], null, [])).not.toContain("/blog");
  });

  test("the sitemap lists published posts with their last change, and no drafts", async () => {
    const { default: sitemap } = await import("../app/sitemap.ts");
    const entries = sitemap();
    const lorem = entries.find((entry) => entry.url === `${BLOG_URL}/lorem-ipsum`);
    expect(lorem.lastModified).toBe("2026-10-10");
    expect(entries.find((entry) => entry.url === `${BLOG_URL}/dolor-sit-amet`).lastModified).toBe("2026-09-15");
    expect(entries.find((entry) => entry.url === BLOG_URL).lastModified).toBe("2026-10-10");
    expect(entries.some((entry) => entry.url.includes(DRAFT.slug))).toBe(false);
    expect(await withPosts([DRAFT], () => sitemap().map((entry) => entry.url))).toEqual(["https://betteroffice.dev"]);
  });

  test("the Atom feed has one escaped entry per published post", async () => {
    const { GET } = await import("../app/blog/feed.xml/route.ts");
    const response = GET();
    expect(response.headers.get("content-type")).toBe("application/atom+xml; charset=utf-8");
    const feed = await response.text();
    expect(feed).toContain(`<link rel="self" type="application/atom+xml" href="${FEED_URL}"/>`);
    expect(feed.match(/<entry>/g)).toHaveLength(PUBLISHED.length);
    expect(feed).toContain("<title>Lorem ipsum dolor sit amet &amp; consectetur</title>");
    expect(feed).toContain("<updated>2026-10-10T00:00:00Z</updated>");
    expect(feed).toContain("&lt;h2 id=&quot;duis-aute-irure&quot;&gt;");
    expect(feed).toContain(`<content type="html" xml:base="${postUrl(PUBLISHED[0])}">`);
    expect(feed).not.toContain(DRAFT.slug);
    expect((await withPosts([DRAFT], () => GET())).status).toBe(404);
  });
});

describe("blog pages", () => {
  const params = (slug) => ({ params: Promise.resolve({ slug }) });

  test("each post renders one h1 inside an article, with machine-readable dates", async () => {
    const { default: Post, generateStaticParams } = await import("../app/blog/[slug]/page.tsx");
    expect(generateStaticParams().map((p) => p.slug).sort()).toEqual(ALL.map((post) => post.slug).sort());
    for (const post of ALL) {
      const html = renderToStaticMarkup(await Post(params(post.slug)));
      expect(html.match(/<h1[\s>]/g)).toHaveLength(1);
      expect(html).toMatch(/<article[^>]*>[\s\S]*<h1[\s\S]*<\/article>/);
      expect(html).toContain(`<time dateTime="${post.date}">`);
      expect(html).toContain(`href="/blog/${post.slug}.md"`);
    }
  });

  test("each post carries valid BlogPosting and BreadcrumbList JSON-LD", async () => {
    const { default: Post } = await import("../app/blog/[slug]/page.tsx");
    const post = ALL.find((p) => p.slug === "lorem-ipsum");
    const [posting, crumbs] = ldBlocks(renderToStaticMarkup(await Post(params(post.slug))));
    expect(posting).toEqual({
      "@context": "https://schema.org",
      "@type": "BlogPosting",
      headline: post.title,
      description: post.description,
      datePublished: "2026-10-01",
      dateModified: "2026-10-10",
      author: [
        { "@type": "Person", name: "Ada Lorem", url: "https://example.com/ada" },
        { "@type": "Person", name: "Brutus Ipsum" },
      ],
      publisher: { "@type": "Organization", name: "OpenOOXML", url: "https://openooxml.org" },
      image: [`${postUrl(post)}/og.png`],
      url: postUrl(post),
      mainEntityOfPage: { "@type": "WebPage", "@id": postUrl(post) },
      inLanguage: "en",
      keywords: "lorem, ipsum",
      isPartOf: { "@type": "Blog", "@id": BLOG_URL, name: "BetterOffice blog" },
    });
    expect(crumbs["@type"]).toBe("BreadcrumbList");
    expect(crumbs.itemListElement.map((item) => [item.position, item.item])).toEqual([
      [1, "https://betteroffice.dev"],
      [2, BLOG_URL],
      [3, postUrl(post)],
    ]);
    expect(jsonLdScript({ headline: "</script><script>" })).not.toContain("</script>");
  });

  test("post metadata sets canonical, alternates, cards, and noindex on drafts", async () => {
    const { generateMetadata } = await import("../app/blog/[slug]/page.tsx");
    const post = PUBLISHED[0];
    const meta = await generateMetadata(params(post.slug));
    expect(meta.alternates).toEqual({
      canonical: postUrl(post),
      types: { "text/markdown": markdownUrl(post), "application/atom+xml": FEED_URL },
    });
    expect(meta.openGraph).toMatchObject({ type: "article", publishedTime: "2026-10-01", modifiedTime: "2026-10-10" });
    expect(meta.openGraph.images[0]).toEqual({ url: `${postUrl(post)}/og.png`, alt: post.title, width: 1200, height: 630 });
    expect(meta.twitter.card).toBe("summary_large_image");
    expect(meta.robots).toBeUndefined();
    expect((await generateMetadata(params(DRAFT.slug))).robots).toEqual({ index: false, follow: false });
  });

  test("the index lists published posts newest first and links each one", async () => {
    const { default: BlogIndex } = await import("../app/blog/page.tsx");
    const html = renderToStaticMarkup(createElement(BlogIndex));
    expect(html.match(/<h1[\s>]/g)).toHaveLength(1);
    const links = [...html.matchAll(/<h2[^>]*><a [^>]*?href="([^"]+)"/g)].map((m) => m[1]);
    expect(links).toEqual(PUBLISHED.map((post) => `/blog/${post.slug}`));
    expect(html).not.toContain(DRAFT.title);
    const [blog, crumbs] = ldBlocks(html);
    expect(blog["@type"]).toBe("Blog");
    expect(blog.blogPost.map((entry) => entry.url)).toEqual(PUBLISHED.map(postUrl));
    expect(crumbs.itemListElement).toHaveLength(2);
  });

  test("the index keeps the site's social image", async () => {
    const { metadata } = await import("../app/blog/page.tsx");
    expect(metadata.openGraph.images[0].url).toBe("/opengraph-image");
    expect(metadata.twitter.images[0].url).toBe("/opengraph-image");
  });
});
