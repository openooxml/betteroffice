import { micromark } from "micromark";
import { gfm, gfmHtml } from "micromark-extension-gfm";
import type { BlogPost } from "./blog";

// ids the post page already uses outside the body
const RESERVED_IDS = ["title"];

const decodeHtml = (value: string) =>
  value.replace(/&quot;/g, '"').replace(/&#39;/g, "'").replace(/&lt;/g, "<").replace(/&gt;/g, ">").replace(/&amp;/g, "&");

// comments, scanned past whole tags so a "<!--" inside a quoted attribute stays put
const MARKUP = /<[a-z](?:[^>"']|"[^"]*"|'[^']*')*>|<!--[\s\S]*?(?:-->|$)/gi;
const ATTRS = String.raw`((?:[^>"']|"[^"]*"|'[^']*')*?)`;
const IMG = new RegExp(String.raw`<img\s${ATTRS}\/?>`, "gi");
const TAG = new RegExp(String.raw`<[a-z][a-z0-9-]*\s${ATTRS}\/?>`, "gi");
const HEADING = new RegExp(String.raw`<(h[2-6])((?:\s(?:[^>"']|"[^"]*"|'[^']*')*?)?)\/?>([\s\S]*?)<\/\1\s*>`, "gi");

function attributes(attrs: string): Map<string, string> {
  const found = new Map<string, string>();
  for (const [, name, double, single, bare] of attrs.matchAll(/([^\s"'=<>\/]+)(?:\s*=\s*(?:"([^"]*)"|'([^']*)'|([^\s"'=<>`]+)))?/g)) {
    const key = name.toLowerCase();
    if (!found.has(key)) found.set(key, decodeHtml(double ?? single ?? bare ?? ""));
  }
  return found;
}

// posts are maintainer-written markdown, so raw HTML (figures with explicit sizes, video) is allowed;
// HTML comments are author notes and never reach the page
export function renderMarkdown(markdown: string): string {
  return micromark(markdown, {
    allowDangerousHtml: true,
    allowDangerousProtocol: false,
    extensions: [gfm()],
    htmlExtensions: [gfmHtml()],
  }).replace(MARKUP, (markup) => (markup.startsWith("<!--") ? "" : markup));
}

// every src the rendered post shows: markdown images, reference-style ones resolved, and raw <img> tags
export function imageSources(html: string): string[] {
  return [...html.matchAll(IMG)].flatMap((match) => attributes(match[1]).get("src") ?? []);
}

function anchor(text: string): string {
  return (
    text
      .replace(/<[^>]+>/g, "")
      .replace(/&[a-z#0-9]+;/gi, "")
      .toLowerCase()
      .replace(/[^a-z0-9]+/g, "-")
      .replace(/^-|-$/g, "") || "section"
  );
}

export function postHtml(post: BlogPost): string {
  const html = renderMarkdown(post.body);
  if (/<h1[\s>/]/i.test(html)) {
    throw new Error(`blog post ${post.slug}: use ## and below in the body; the title is the page's only h1`);
  }
  const used = new Set([...RESERVED_IDS, ...[...html.matchAll(TAG)].flatMap((match) => attributes(match[1]).get("id") ?? [])]);
  const unique = (base: string) => {
    let id = base;
    for (let n = 2; used.has(id); n++) id = `${base}-${n}`;
    used.add(id);
    return id;
  };
  const sizes = post.images ?? {};
  return html
    .replace(HEADING, (heading: string, tag: string, attrs: string, inner: string) =>
      attributes(attrs).has("id") ? heading : `<${tag}${attrs.trimEnd()} id="${unique(anchor(inner))}">${inner}</${tag}>`,
    )
    .replace(IMG, (tag: string, attrs: string) => {
      const has = attributes(attrs);
      const src = has.get("src");
      const size = src !== undefined && !has.has("width") ? sizes[src] : undefined;
      const extra = [
        ...(size ? [`width="${size[0]}" height="${size[1]}"`] : []),
        ...(has.has("loading") ? [] : ['loading="lazy"']),
        ...(has.has("decoding") ? [] : ['decoding="async"']),
      ];
      return extra.length ? `<img ${attrs.trimEnd()} ${extra.join(" ")} />` : tag;
    })
    .replace(/<table>/g, '<div class="post-table"><table>')
    .replace(/<\/table>/g, "</table></div>");
}
