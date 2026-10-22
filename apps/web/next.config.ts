import { existsSync, readFileSync } from "node:fs";
import { join } from "node:path";
import type { NextConfig } from "next";
import { PHASE_DEVELOPMENT_SERVER } from "next/constants";
import { initOpenNextCloudflareForDev } from "@opennextjs/cloudflare";
import { agentEvalsSection } from "./lib/agent-evals";
import { loadPosts } from "./lib/blog-source";

function readme(): string {
  for (const path of [join(process.cwd(), "..", "..", "README.md"), join(process.cwd(), "README.md")]) {
    if (existsSync(path)) return readFileSync(path, "utf8");
  }
  return "";
}

// the README's generated benchmark block for lib/fidelity.ts, and its evals section for lib/agent-evals.ts
const README = readme();

export default function nextConfig(phase: string): NextConfig {
  // drafts render in `next dev` and in builds with BETTEROFFICE_BLOG_DRAFTS=1, and are never listed
  const drafts = phase === PHASE_DEVELOPMENT_SERVER || process.env.BETTEROFFICE_BLOG_DRAFTS === "1";
  const blogDir = join(process.cwd(), process.env.BETTEROFFICE_BLOG_DIR ?? "content/blog");
  return {
    reactStrictMode: true,
    env: {
      BETTEROFFICE_BENCHMARKS:
        README.match(/<!-- BEGIN GENERATED VISUAL FIDELITY -->[\s\S]*?<!-- END GENERATED VISUAL FIDELITY -->/)?.[0] ?? "",
      BETTEROFFICE_AGENT_EVALS: agentEvalsSection(README),
      BETTEROFFICE_BLOG: JSON.stringify(loadPosts(blogDir, { drafts, publicDir: join(process.cwd(), "public") })),
      BETTEROFFICE_LLMS: readFileSync(join(process.cwd(), "content/llms.txt"), "utf8"),
    },
  };
}

initOpenNextCloudflareForDev();
