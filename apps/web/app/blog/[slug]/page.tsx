import type { Metadata } from "next";
import Link from "next/link";
import { notFound } from "next/navigation";
import {
  BLOG_PATH,
  BLOG_URL,
  FEED_URL,
  OG_SIZE,
  blogPosts,
  breadcrumbJsonLd,
  canonicalUrl,
  formatDate,
  imageUrl,
  jsonLdScript,
  markdownUrl,
  modified,
  postJsonLd,
  postUrl,
  postPath,
  readingMinutes,
  type BlogPost,
} from "../../../lib/blog";
import { postHtml } from "../../../lib/blog-html";
import { BLOG, SITE } from "../../content";
import { label, lead, metaLink } from "../styles";

type Params = { params: Promise<{ slug: string }> };

export function generateStaticParams() {
  return blogPosts().map((post) => ({ slug: post.slug }));
}

async function find({ params }: Params): Promise<BlogPost> {
  const { slug } = await params;
  const post = blogPosts().find((candidate) => candidate.slug === slug);
  if (!post) notFound();
  return post;
}

export async function generateMetadata(props: Params): Promise<Metadata> {
  const post = await find(props);
  const image = {
    url: imageUrl(post),
    alt: post.imageAlt ?? post.title,
    ...(post.image ? {} : OG_SIZE),
  };
  return {
    title: post.title,
    description: post.description,
    authors: post.authors.map((author) => ({ name: author.name, url: author.url })),
    ...(post.tags.length ? { keywords: post.tags } : {}),
    alternates: {
      canonical: canonicalUrl(post),
      types: { "text/markdown": markdownUrl(post), "application/atom+xml": FEED_URL },
    },
    openGraph: {
      type: "article",
      siteName: "BetterOffice",
      url: postUrl(post),
      title: post.title,
      description: post.description,
      publishedTime: post.date,
      modifiedTime: modified(post),
      authors: post.authors.map((author) => author.url ?? author.name),
      ...(post.tags.length ? { tags: post.tags } : {}),
      images: [image],
    },
    twitter: { card: "summary_large_image", title: post.title, description: post.description, images: [image] },
    ...(post.draft ? { robots: { index: false, follow: false } } : {}),
  };
}

export default async function Post(props: Params) {
  const post = await find(props);
  const crumbs = breadcrumbJsonLd([
    { name: "BetterOffice", url: SITE },
    { name: BLOG.heading, url: BLOG_URL },
    { name: post.title, url: postUrl(post) },
  ]);
  return (
    <>
      <script type="application/ld+json" dangerouslySetInnerHTML={{ __html: jsonLdScript(postJsonLd(post)) }} />
      <script type="application/ld+json" dangerouslySetInnerHTML={{ __html: jsonLdScript(crumbs) }} />

      <article aria-labelledby="title">
        <header className="wrap pt-[clamp(3.5rem,8vw,6.5rem)] pb-[clamp(2.5rem,5vw,3.5rem)]">
          <nav aria-label="Breadcrumb" className={`${label} mb-7`}>
            <ol className="flex list-none flex-wrap items-center gap-2">
              <li>
                <Link href={BLOG_PATH} className="text-dim no-underline transition-colors hover:text-fg">
                  {BLOG.heading}
                </Link>
              </li>
              <li aria-hidden="true">/</li>
              <li>
                <time dateTime={post.date}>{post.date}</time>
              </li>
              {post.draft && (
                <li className="border border-line px-1.5 py-px text-[0.625rem] text-fg">draft</li>
              )}
            </ol>
          </nav>
          <h1
            id="title"
            className="max-w-[22ch] text-[clamp(2.25rem,5.2vw,4.25rem)] leading-[0.98] font-[560] tracking-[-0.04em]"
          >
            {post.title}
          </h1>
          <p className={`${lead} mt-6 max-w-[40rem]`}>{post.description}</p>
          <p className="mt-8 flex flex-wrap gap-x-6 gap-y-2 font-mono text-xs text-dim">
            <span>
              {post.authors.map((author, index) => (
                <span key={author.name}>
                  {index > 0 && (index === post.authors.length - 1 ? " and " : ", ")}
                  {author.url ? (
                    <a href={author.url} rel="author" className="text-ink no-underline hover:text-fg">
                      {author.name}
                    </a>
                  ) : (
                    <span className="text-ink">{author.name}</span>
                  )}
                </span>
              ))}
            </span>
            <time dateTime={post.date}>{formatDate(post.date)}</time>
            {post.updated && (
              <span>
                updated <time dateTime={post.updated}>{formatDate(post.updated)}</time>
              </span>
            )}
            <span>{readingMinutes(post)} min read</span>
          </p>
        </header>

        <div className="wrap rule py-[clamp(3rem,6vw,4.5rem)]">
          <div className="post" dangerouslySetInnerHTML={{ __html: postHtml(post) }} />
        </div>

        <footer className="wrap rule flex flex-wrap items-center justify-between gap-x-10 gap-y-4 py-10">
          <Link href={BLOG_PATH} className={metaLink}>
            all posts
          </Link>
          {post.tags.length > 0 && (
            <ul aria-label="Tags" className="flex list-none flex-wrap gap-2">
              {post.tags.map((tag) => (
                <li key={tag} className="border border-line px-2 py-0.5 font-mono text-[0.6875rem] text-dim">
                  {tag}
                </li>
              ))}
            </ul>
          )}
          <p className="flex gap-x-5">
            <a href={`${postPath(post)}.md`} className={metaLink} type="text/markdown">
              markdown
            </a>
            <a href={`${postPath(post)}.txt`} className={metaLink} type="text/plain">
              text
            </a>
            <a href="/blog/feed.xml" className={metaLink} type="application/atom+xml">
              feed
            </a>
          </p>
        </footer>
      </article>
    </>
  );
}
