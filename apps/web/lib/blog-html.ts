import { micromark } from "micromark";
import { gfm, gfmHtml } from "micromark-extension-gfm";
import type { BlogPost } from "./blog";

// ids the post page already uses outside the body
const RESERVED_IDS = ["title"];

const decodeHtml = (value: string) => value.replace(/&amp;/g, "&").replace(/&quot;/g, '"');

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

// posts are maintainer-written markdown, so raw HTML (figures with explicit sizes, video) is allowed
export function postHtml(post: BlogPost): string {
  const html = micromark(post.body, {
    allowDangerousHtml: true,
    allowDangerousProtocol: false,
    extensions: [gfm()],
    htmlExtensions: [gfmHtml()],
  });
  if (/<h1[\s>]/.test(html)) {
    throw new Error(`blog post ${post.slug}: use ## and below in the body; the title is the page's only h1`);
  }
  const used = new Set(RESERVED_IDS);
  const unique = (base: string) => {
    let id = base;
    for (let n = 2; used.has(id); n++) id = `${base}-${n}`;
    used.add(id);
    return id;
  };
  const sizes = post.images ?? {};
  return html
    .replace(/<(h[2-6])>([\s\S]*?)<\/\1>/g, (_, tag: string, inner: string) => `<${tag} id="${unique(anchor(inner))}">${inner}</${tag}>`)
    .replace(/<img ([^>]*?)\/?>/g, (tag: string, attrs: string) => {
      const src = attrs.match(/\bsrc="([^"]*)"/)?.[1];
      const size = src && !/\bwidth=/.test(attrs) ? sizes[decodeHtml(src)] : undefined;
      const extra = [
        ...(size ? [`width="${size[0]}" height="${size[1]}"`] : []),
        ...(/\bloading=/.test(attrs) ? [] : ['loading="lazy"']),
        ...(/\bdecoding=/.test(attrs) ? [] : ['decoding="async"']),
      ];
      return extra.length ? `<img ${attrs.trimEnd()} ${extra.join(" ")} />` : tag;
    })
    .replace(/<table>/g, '<div class="post-table"><table>')
    .replace(/<\/table>/g, "</table></div>");
}
