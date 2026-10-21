import type { Metadata } from "next";
import Link from "next/link";
import { notFound } from "next/navigation";
import {
  BLOG_MARKDOWN_URL,
  BLOG_PATH,
  BLOG_URL,
  FEED_URL,
  OG_SIZE,
  authorNames,
  blogJsonLd,
  breadcrumbJsonLd,
  formatDate,
  jsonLdScript,
  postPath,
  publishedPosts,
  readingMinutes,
} from "../../lib/blog";
import { BLOG, SITE } from "../content";
import { label, lead, metaLink } from "./styles";

const TITLE = `BetterOffice ${BLOG.heading.toLowerCase()}`;
// a nested openGraph replaces the root one, so the site card is named again here
const SITE_IMAGE = { url: "/opengraph-image", alt: "BetterOffice", ...OG_SIZE };

export const metadata: Metadata = {
  title: BLOG.heading,
  description: BLOG.description,
  alternates: {
    canonical: BLOG_PATH,
    types: { "text/markdown": BLOG_MARKDOWN_URL, "application/atom+xml": FEED_URL },
  },
  openGraph: {
    type: "website",
    siteName: "BetterOffice",
    url: BLOG_PATH,
    title: TITLE,
    description: BLOG.description,
    images: [SITE_IMAGE],
  },
  twitter: { card: "summary_large_image", title: TITLE, description: BLOG.description, images: [SITE_IMAGE] },
};

export default function BlogIndex() {
  const posts = publishedPosts();
  if (posts.length === 0) notFound();
  const crumbs = breadcrumbJsonLd([
    { name: "BetterOffice", url: SITE },
    { name: BLOG.heading, url: BLOG_URL },
  ]);
  return (
    <>
      <script type="application/ld+json" dangerouslySetInnerHTML={{ __html: jsonLdScript(blogJsonLd(posts)) }} />
      <script type="application/ld+json" dangerouslySetInnerHTML={{ __html: jsonLdScript(crumbs) }} />

      <section aria-labelledby="blog" className="wrap pt-[clamp(4rem,9vw,7rem)] pb-[clamp(2.5rem,5vw,4rem)]">
        <p className={`${label} mb-7 tabular-nums`}>
          {posts.length} {posts.length === 1 ? "post" : "posts"}
        </p>
        <h1 id="blog" className="text-[clamp(2.75rem,7vw,5.5rem)] leading-[0.9] font-[560] tracking-[-0.05em]">
          {BLOG.heading}
        </h1>
        <p className={`${lead} mt-6 max-w-[34rem]`}>{BLOG.description}</p>
        <p className="mt-6 flex flex-wrap gap-x-5 gap-y-2">
          <a href="/blog/feed.xml" className={metaLink} type="application/atom+xml">feed</a>
          <a href="/blog.md" className={metaLink} type="text/markdown">markdown</a>
          <a href="/blog.txt" className={metaLink} type="text/plain">text</a>
        </p>
      </section>

      <div className="wrap rule pb-[clamp(4rem,8vw,7rem)]">
        <ol className="list-none">
          {posts.map((post) => (
            <li key={post.slug} className="border-b border-line-soft">
              <article className="grid gap-x-12 gap-y-3 py-[clamp(1.75rem,3.5vw,2.5rem)] lg:grid-cols-12">
                <time dateTime={post.date} className="font-mono text-xs text-dim tabular-nums lg:col-span-3 lg:pt-2">
                  {formatDate(post.date)}
                </time>
                <div className="lg:col-span-8">
                  <h2 className="text-[clamp(1.375rem,2.4vw,1.75rem)] leading-[1.2] font-[560] tracking-[-0.025em]">
                    <Link href={postPath(post)} className="text-fg no-underline hover:underline hover:decoration-faint hover:underline-offset-4">
                      {post.title}
                    </Link>
                  </h2>
                  <p className="mt-3 max-w-[40rem] text-[0.9375rem] leading-[1.65] text-ink">{post.description}</p>
                  <p className="mt-3 font-mono text-[0.6875rem] text-dim">
                    {authorNames(post)} · {readingMinutes(post)} min read
                  </p>
                </div>
              </article>
            </li>
          ))}
        </ol>
      </div>
    </>
  );
}
