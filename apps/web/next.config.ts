import { existsSync, readFileSync } from "node:fs";
import { join } from "node:path";
import type { NextConfig } from "next";
import { initOpenNextCloudflareForDev } from "@opennextjs/cloudflare";

// the README's generated benchmark block, for lib/fidelity.ts
function benchmarks(): string {
  for (const path of [join(process.cwd(), "..", "..", "README.md"), join(process.cwd(), "README.md")]) {
    if (!existsSync(path)) continue;
    const block = readFileSync(path, "utf8").match(
      /<!-- BEGIN GENERATED VISUAL FIDELITY -->[\s\S]*?<!-- END GENERATED VISUAL FIDELITY -->/,
    );
    if (block) return block[0];
  }
  return "";
}

const nextConfig: NextConfig = {
  reactStrictMode: true,
  env: { BETTEROFFICE_BENCHMARKS: benchmarks() },
};

export default nextConfig;

initOpenNextCloudflareForDev();
