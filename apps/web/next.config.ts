import { existsSync, readFileSync } from "node:fs";
import { join } from "node:path";
import type { NextConfig } from "next";
import { initOpenNextCloudflareForDev } from "@opennextjs/cloudflare";
import { agentEvalsSection } from "./lib/agent-evals";

function readme(): string {
  for (const path of [join(process.cwd(), "..", "..", "README.md"), join(process.cwd(), "README.md")]) {
    if (existsSync(path)) return readFileSync(path, "utf8");
  }
  return "";
}

// the README's generated benchmark block for lib/fidelity.ts, and its evals section for lib/agent-evals.ts
const README = readme();

const nextConfig: NextConfig = {
  reactStrictMode: true,
  env: {
    BETTEROFFICE_BENCHMARKS:
      README.match(/<!-- BEGIN GENERATED VISUAL FIDELITY -->[\s\S]*?<!-- END GENERATED VISUAL FIDELITY -->/)?.[0] ?? "",
    BETTEROFFICE_AGENT_EVALS: agentEvalsSection(README),
  },
};

export default nextConfig;

initOpenNextCloudflareForDev();
