import { BLOG, SITE } from "../app/content";

export type BlogAuthor = { name: string; url?: string };

export type BlogPost = {
  slug: string;
  title: string;
  description: string;
  date: string;
  updated?: string;
  authors: BlogAuthor[];
  tags: string[];
  canonical?: string;
  image?: string;
  imageAlt?: string;
  draft: boolean;
  source: string;
  body: string;
  images?: Record<string, [number, number]>;
};

export const MARKDOWN_TYPE = "text/markdown; charset=utf-8";
export const TEXT_TYPE = "text/plain; charset=utf-8";
export const FEED_TYPE = "application/atom+xml; charset=utf-8";
export const OG_SIZE = { width: 1200, height: 630 };

export const BLOG_PATH = "/blog";
export const BLOG_URL = `${SITE}${BLOG_PATH}`;
export const BLOG_MARKDOWN_URL = `${SITE}/blog.md`;
export const BLOG_TEXT_URL = `${SITE}/blog.txt`;
export const FEED_URL = `${SITE}/blog/feed.xml`;

export function parseBlog(json: string): BlogPost[] {
  const posts = JSON.parse(json || "[]") as BlogPost[];
  return posts.sort(
    (a, b) =>
      b.date.localeCompare(a.date) ||
      (b.updated ?? b.date).localeCompare(a.updated ?? a.date) ||
      a.slug.localeCompare(b.slug),
  );
}

// every post this build renders: drafts too in preview builds (see next.config.ts)
export function blogPosts(): BlogPost[] {
  return parseBlog(process.env.BETTEROFFICE_BLOG ?? "");
}

export function listed(posts: BlogPost[]): BlogPost[] {
  return posts.filter((post) => !post.draft);
}

// what the index, nav, feed, sitemap and llms.txt show
export function publishedPosts(): BlogPost[] {
  return listed(blogPosts());
}

export const postPath = (post: Pick<BlogPost, "slug">) => `${BLOG_PATH}/${post.slug}`;
export const postUrl = (post: Pick<BlogPost, "slug">) => `${SITE}${postPath(post)}`;
export const markdownUrl = (post: Pick<BlogPost, "slug">) => `${postUrl(post)}.md`;
export const textUrl = (post: Pick<BlogPost, "slug">) => `${postUrl(post)}.txt`;
export const canonicalUrl = (post: BlogPost) => post.canonical ?? postUrl(post);
export const modified = (post: BlogPost) => post.updated ?? post.date;

export function imageUrl(post: BlogPost): string {
  if (!post.image) return `${postUrl(post)}/og.png`;
  return post.image.startsWith("/") ? `${SITE}${post.image}` : post.image;
}

export function formatDate(day: string): string {
  return new Date(`${day}T00:00:00Z`).toLocaleDateString("en-US", {
    timeZone: "UTC",
    year: "numeric",
    month: "long",
    day: "numeric",
  });
}

export function readingMinutes(post: BlogPost): number {
  const words = post.body.replace(/```[\s\S]*?```/g, " ").split(/\s+/).filter(Boolean).length;
  return Math.max(1, Math.round(words / 230));
}

export function authorNames(post: BlogPost): string {
  const names = post.authors.map((author) => author.name);
  return names.length <= 2 ? names.join(" and ") : `${names.slice(0, -1).join(", ")} and ${names.at(-1)}`;
}

// the post as one plain document: a readable header instead of YAML, then the body as written
export function postDocument(post: BlogPost): string {
  const lines = [
    `# ${post.title}`,
    "",
    `> ${post.description}`,
    "",
    `- Published: ${post.date}`,
    ...(post.updated ? [`- Updated: ${post.updated}`] : []),
    `- Authors: ${post.authors.map((author) => author.name).join(", ")}`,
    ...(post.tags.length ? [`- Tags: ${post.tags.join(", ")}`] : []),
    `- URL: ${canonicalUrl(post)}`,
    `- Markdown: ${markdownUrl(post)}`,
    "",
    post.body.trimEnd(),
    "",
  ];
  return lines.join("\n");
}

export function blogIndexMarkdown(posts: BlogPost[]): string {
  const entries = posts.map(
    (post) => `- [${post.title}](${markdownUrl(post)}) (${post.date}): ${post.description}`,
  );
  return [
    `# ${BLOG.heading}`,
    "",
    BLOG.description,
    "",
    `Every post is also served as markdown at its URL plus \`.md\` and as plain text plus \`.txt\`. Feed: ${FEED_URL}`,
    "",
    ...entries,
    "",
  ].join("\n");
}

export function blogIndexText(posts: BlogPost[]): string {
  const entries = posts.flatMap((post) => [
    post.title,
    `${post.date}, ${authorNames(post)}`,
    post.description,
    `${postUrl(post)} (text: ${textUrl(post)}, markdown: ${markdownUrl(post)})`,
    "",
  ]);
  return [BLOG.heading, "", BLOG.description, "", ...entries].join("\n");
}

const TWIN = /^\/blog\/([a-z0-9]+(?:-[a-z0-9]+)*)\.(md|txt)$/;

export type BlogTwin = { body: string; type: string; canonical: string; draft: boolean };

// /blog.md, /blog.txt, /blog/<slug>.md and /blog/<slug>.txt; null for any other path
export function blogTwin(pathname: string, posts: BlogPost[]): BlogTwin | null {
  const index = listed(posts);
  if (pathname === "/blog.md" || pathname === "/blog.txt") {
    if (index.length === 0) return null;
    return pathname === "/blog.md"
      ? { body: blogIndexMarkdown(index), type: MARKDOWN_TYPE, canonical: BLOG_URL, draft: false }
      : { body: blogIndexText(index), type: TEXT_TYPE, canonical: BLOG_URL, draft: false };
  }
  const match = pathname.match(TWIN);
  const post = match && posts.find((candidate) => candidate.slug === match[1]);
  if (!post) return null;
  return match[2] === "md"
    ? { body: post.source, type: MARKDOWN_TYPE, canonical: canonicalUrl(post), draft: post.draft }
    : { body: postDocument(post), type: TEXT_TYPE, canonical: canonicalUrl(post), draft: post.draft };
}

// the markdown twin of an HTML blog URL, for Accept negotiation
export function blogMarkdownFor(pathname: string, posts: BlogPost[]): { body: string; url: string; draft: boolean } | null {
  const path = pathname.replace(/\/$/, "");
  if (path === BLOG_PATH) {
    return listed(posts).length ? { body: blogIndexMarkdown(listed(posts)), url: BLOG_MARKDOWN_URL, draft: false } : null;
  }
  const post = posts.find((candidate) => postPath(candidate) === path);
  return post ? { body: post.source, url: markdownUrl(post), draft: post.draft } : null;
}

// the index and every post whose canonical URL is on this site
export function blogSitemap(posts: BlogPost[]) {
  if (posts.length === 0) return [];
  return [
    { url: BLOG_URL, lastModified: posts.map(modified).sort().at(-1), changeFrequency: "weekly" as const, priority: 0.7 },
    ...posts
      .filter((post) => canonicalUrl(post) === postUrl(post))
      .map((post) => ({ url: postUrl(post), lastModified: modified(post), changeFrequency: "monthly" as const, priority: 0.6 })),
  ];
}

const PUBLISHER = { "@type": "Organization", name: "OpenOOXML", url: "https://openooxml.org" };

function person(author: BlogAuthor) {
  return { "@type": "Person", name: author.name, ...(author.url ? { url: author.url } : {}) };
}

export function postJsonLd(post: BlogPost) {
  return {
    "@context": "https://schema.org",
    "@type": "BlogPosting",
    headline: post.title,
    description: post.description,
    datePublished: post.date,
    dateModified: modified(post),
    author: post.authors.map(person),
    publisher: PUBLISHER,
    image: [imageUrl(post)],
    url: postUrl(post),
    mainEntityOfPage: { "@type": "WebPage", "@id": canonicalUrl(post) },
    inLanguage: "en",
    ...(post.tags.length ? { keywords: post.tags.join(", ") } : {}),
    isPartOf: { "@type": "Blog", "@id": BLOG_URL, name: `BetterOffice ${BLOG.heading.toLowerCase()}` },
  };
}

export function blogJsonLd(posts: BlogPost[]) {
  return {
    "@context": "https://schema.org",
    "@type": "Blog",
    "@id": BLOG_URL,
    name: `BetterOffice ${BLOG.heading.toLowerCase()}`,
    description: BLOG.description,
    url: BLOG_URL,
    publisher: PUBLISHER,
    blogPost: posts.map((post) => ({
      "@type": "BlogPosting",
      headline: post.title,
      url: postUrl(post),
      datePublished: post.date,
      dateModified: modified(post),
    })),
  };
}

export function breadcrumbJsonLd(trail: { name: string; url: string }[]) {
  return {
    "@context": "https://schema.org",
    "@type": "BreadcrumbList",
    itemListElement: trail.map((crumb, index) => ({
      "@type": "ListItem",
      position: index + 1,
      name: crumb.name,
      item: crumb.url,
    })),
  };
}

// safe inside <script>: no "</script>", no HTML comment openers
export function jsonLdScript(data: unknown): string {
  return JSON.stringify(data).replace(/</g, "\\u003c").replace(/>/g, "\\u003e").replace(/&/g, "\\u0026");
}

function xml(value: string): string {
  return value
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&apos;");
}

const stamp = (day: string) => `${day}T00:00:00Z`;

export function atomFeed(posts: BlogPost[], html: (post: BlogPost) => string): string {
  const updated = posts.map(modified).sort().at(-1) ?? "1970-01-01";
  const entries = posts.map((post) =>
    [
      "  <entry>",
      `    <id>${xml(postUrl(post))}</id>`,
      `    <title>${xml(post.title)}</title>`,
      `    <link rel="alternate" type="text/html" href="${xml(canonicalUrl(post))}"/>`,
      `    <link rel="alternate" type="text/markdown" href="${xml(markdownUrl(post))}"/>`,
      `    <published>${stamp(post.date)}</published>`,
      `    <updated>${stamp(modified(post))}</updated>`,
      ...post.authors.map(
        (author) =>
          `    <author><name>${xml(author.name)}</name>${author.url ? `<uri>${xml(author.url)}</uri>` : ""}</author>`,
      ),
      ...post.tags.map((tag) => `    <category term="${xml(tag)}"/>`),
      `    <summary>${xml(post.description)}</summary>`,
      `    <content type="html" xml:base="${xml(postUrl(post))}">${xml(html(post))}</content>`,
      "  </entry>",
    ].join("\n"),
  );
  return [
    '<?xml version="1.0" encoding="utf-8"?>',
    '<feed xmlns="http://www.w3.org/2005/Atom">',
    `  <id>${BLOG_URL}</id>`,
    `  <title>BetterOffice ${xml(BLOG.heading.toLowerCase())}</title>`,
    `  <subtitle>${xml(BLOG.description)}</subtitle>`,
    `  <link rel="self" type="application/atom+xml" href="${FEED_URL}"/>`,
    `  <link rel="alternate" type="text/html" href="${BLOG_URL}"/>`,
    `  <updated>${stamp(updated)}</updated>`,
    ...entries,
    "</feed>",
    "",
  ].join("\n");
}

// llms.txt with a Blog section listing each post's markdown, placed before the closing Links section
export function llmsTxt(base: string, posts: BlogPost[]): string {
  if (posts.length === 0) return base;
  const section = [
    "## Blog",
    "",
    `- [Blog index](${BLOG_MARKDOWN_URL}): every post, newest first. Atom feed: ${FEED_URL}`,
    ...posts.map((post) => `- [${post.title}](${markdownUrl(post)}): ${post.description}`),
    "",
  ].join("\n");
  const links = base.indexOf("\n## Links");
  if (links < 0) return `${base.trimEnd()}\n\n${section}`;
  return `${base.slice(0, links + 1)}${section}\n${base.slice(links + 1)}`;
}

// the landing page and every post in one markdown file
export function llmsFull(home: string, posts: BlogPost[]): string {
  return [home.trimEnd(), ...posts.map(postDocument)].join("\n\n---\n\n");
}
