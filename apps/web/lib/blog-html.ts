import { micromark } from "micromark";
import { gfm, gfmHtml } from "micromark-extension-gfm";
import type { BlogPost } from "./blog";

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
  const seen = new Map<string, number>();
  return html
    .replace(/<(h[2-6])>([\s\S]*?)<\/\1>/g, (_, tag: string, inner: string) => {
      const base = anchor(inner);
      const count = seen.get(base) ?? 0;
      seen.set(base, count + 1);
      return `<${tag} id="${count ? `${base}-${count + 1}` : base}">${inner}</${tag}>`;
    })
    .replace(/<img (?![^>]*\bloading=)/g, '<img loading="lazy" decoding="async" ')
    .replace(/<table>/g, '<div class="post-table"><table>')
    .replace(/<\/table>/g, "</table></div>");
}
