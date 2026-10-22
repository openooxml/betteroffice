import { ImageResponse } from "next/og";
import { OG_SIZE, authorNames, blogPosts, formatDate } from "../../../../lib/blog";
import { logoUri } from "../../../og-logo";

export const dynamic = "force-static";

export function generateStaticParams() {
  return blogPosts().map((post) => ({ slug: post.slug }));
}

export async function GET(_request: Request, { params }: { params: Promise<{ slug: string }> }) {
  const { slug } = await params;
  const post = blogPosts().find((candidate) => candidate.slug === slug);
  if (!post) return new Response("Not found\n", { status: 404 });
  const size = post.title.length > 70 ? 56 : post.title.length > 40 ? 68 : 80;
  return new ImageResponse(
    (
      <div
        style={{
          width: "100%",
          height: "100%",
          display: "flex",
          flexDirection: "column",
          justifyContent: "space-between",
          padding: "72px 96px",
          backgroundColor: "#ffffff",
          color: "#141414",
          fontFamily: "sans-serif",
          position: "relative",
          overflow: "hidden",
        }}
      >
        <img
          src={logoUri}
          width={760}
          height={567}
          style={{ position: "absolute", right: -170, top: -80, opacity: 0.045, transform: "rotate(-6deg)" }}
        />
        <div style={{ display: "flex", alignItems: "center", gap: 20, fontSize: 28, color: "#5c5c5c" }}>
          <img src={logoUri} width={64} height={48} />
          BetterOffice blog
        </div>
        <div style={{ display: "flex", fontSize: size, fontWeight: 700, letterSpacing: -2, lineHeight: 1.08, maxWidth: 980 }}>
          {post.title}
        </div>
        <div style={{ display: "flex", fontSize: 28, color: "#5c5c5c" }}>
          {`${formatDate(post.date)} · ${authorNames(post)}`}
        </div>
      </div>
    ),
    { ...OG_SIZE, headers: { "Cache-Control": "public, max-age=0, s-maxage=86400" } },
  );
}
