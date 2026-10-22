import { existsSync, readdirSync, readFileSync } from "node:fs";
import { join } from "node:path";
import { parse } from "yaml";
import type { BlogAuthor, BlogPost } from "./blog";
import { imageSources, renderMarkdown } from "./blog-html";
import { imageSize } from "./image-size";

const SLUG = /^[a-z0-9]+(?:-[a-z0-9]+)*$/;
const DAY = /^\d{4}-\d{2}-\d{2}$/;
const FRONTMATTER = /^---\r?\n([\s\S]*?)\r?\n---[ \t]*(?:\r?\n|$)/;

function fail(file: string, message: string): never {
  throw new Error(`blog post ${file}: ${message}`);
}

function text(file: string, data: Record<string, unknown>, key: string): string {
  const value = data[key];
  if (typeof value !== "string" || value.trim() === "") fail(file, `\`${key}\` must be a non-empty string`);
  return value.trim();
}

function optionalText(file: string, data: Record<string, unknown>, key: string): string | undefined {
  return data[key] === undefined ? undefined : text(file, data, key);
}

function day(file: string, data: Record<string, unknown>, key: string): string | undefined {
  const value = data[key] instanceof Date ? (data[key] as Date).toISOString().slice(0, 10) : data[key];
  if (value === undefined) return undefined;
  if (typeof value !== "string" || !DAY.test(value) || !sameDay(value)) {
    fail(file, `\`${key}\` must be a YYYY-MM-DD date`);
  }
  return value;
}

function sameDay(value: string): boolean {
  const parsed = new Date(`${value}T00:00:00Z`);
  return !Number.isNaN(parsed.getTime()) && parsed.toISOString().slice(0, 10) === value;
}

function authors(file: string, value: unknown): BlogAuthor[] {
  const list = typeof value === "string" ? [value] : value;
  if (!Array.isArray(list) || list.length === 0) fail(file, "`authors` must name at least one author");
  return list.map((entry) => {
    if (typeof entry === "string" && entry.trim()) return { name: entry.trim() };
    if (entry && typeof entry === "object" && typeof entry.name === "string" && entry.name.trim()) {
      if (entry.url !== undefined && !isAbsoluteUrl(entry.url)) fail(file, "an author `url` must be an absolute URL");
      return entry.url ? { name: entry.name.trim(), url: entry.url } : { name: entry.name.trim() };
    }
    fail(file, "each author is a name or { name, url }");
  });
}

function tags(file: string, value: unknown): string[] {
  if (value === undefined) return [];
  const list = typeof value === "string" ? [value] : value;
  if (!Array.isArray(list) || !list.every((tag) => typeof tag === "string" && tag.trim())) {
    fail(file, "`tags` must be a list of strings");
  }
  return list.map((tag) => tag.trim());
}

function isAbsoluteUrl(value: unknown): value is string {
  return typeof value === "string" && /^https?:\/\/\S+$/.test(value);
}

export function parsePost(file: string, source: string): BlogPost {
  const slug = file.replace(/\.md$/, "");
  if (!SLUG.test(slug)) fail(file, "the file name must be a lowercase slug such as `my-post.md`");
  const match = source.match(FRONTMATTER);
  if (!match) fail(file, "missing `---` frontmatter");
  const data = parse(match[1]) as unknown;
  if (!data || typeof data !== "object" || Array.isArray(data)) fail(file, "frontmatter must be a mapping");
  const fields = data as Record<string, unknown>;

  const known = new Set(["title", "description", "date", "updated", "authors", "tags", "canonical", "image", "imageAlt", "draft"]);
  const unknown = Object.keys(fields).filter((key) => !known.has(key));
  if (unknown.length > 0) fail(file, `unknown frontmatter ${unknown.map((key) => `\`${key}\``).join(", ")}`);

  const date = day(file, fields, "date");
  if (!date) fail(file, "`date` is required");
  const updated = day(file, fields, "updated");
  if (updated && updated < date) fail(file, "`updated` is earlier than `date`");
  const canonical = fields.canonical;
  if (canonical !== undefined && !isAbsoluteUrl(canonical)) fail(file, "`canonical` must be an absolute URL");
  const image = optionalText(file, fields, "image");
  if (image && !image.startsWith("/") && !isAbsoluteUrl(image)) fail(file, "`image` must be a site path or an absolute URL");
  if (fields.draft !== undefined && typeof fields.draft !== "boolean") fail(file, "`draft` must be true or false");

  const body = source.slice(match[0].length).replace(/^(?:[ \t]*\r?\n)+/, "");
  if (!body.trim()) fail(file, "the post has no body");

  return {
    slug,
    title: text(file, fields, "title"),
    description: text(file, fields, "description"),
    date,
    ...(updated ? { updated } : {}),
    authors: authors(file, fields.authors),
    tags: tags(file, fields.tags),
    ...(canonical ? { canonical: canonical as string } : {}),
    ...(image ? { image } : {}),
    ...(fields.imageAlt !== undefined ? { imageAlt: text(file, fields, "imageAlt") } : {}),
    draft: fields.draft === true,
    source,
    body,
  };
}

// width and height of every site-local image the rendered post shows, read from public/
export function localImages(post: BlogPost, publicDir: string): Record<string, [number, number]> {
  const sizes: Record<string, [number, number]> = {};
  for (const src of imageSources(renderMarkdown(post.body))) {
    if (!src.startsWith("/") || src.startsWith("//") || src in sizes) continue;
    let path: string;
    try {
      path = decodeURI(src.replace(/[?#].*$/, ""));
    } catch {
      fail(`${post.slug}.md`, `image ${src} is not a valid path`);
    }
    const file = join(publicDir, path);
    if (!existsSync(file)) fail(`${post.slug}.md`, `image ${src} is not in public/`);
    const size = imageSize(readFileSync(file));
    if (!size) fail(`${post.slug}.md`, `cannot read the size of ${src}`);
    sizes[src] = size;
  }
  return sizes;
}

export function loadPosts(dir: string, { drafts, publicDir }: { drafts: boolean; publicDir: string }): BlogPost[] {
  if (!existsSync(dir)) return [];
  return readdirSync(dir)
    .filter((file) => file.endsWith(".md"))
    .map((file) => parsePost(file, readFileSync(join(dir, file), "utf8")))
    .filter((post) => drafts || !post.draft)
    .map((post) => ({ ...post, images: localImages(post, publicDir) }));
}
