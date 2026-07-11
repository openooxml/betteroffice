import type { NextConfig } from "next";
import { createMDX } from "fumadocs-mdx/next";
import { initOpenNextCloudflareForDev } from "@opennextjs/cloudflare";

const withMDX = createMDX();

const config: NextConfig = {
  reactStrictMode: true,
};

export default withMDX(config);

if (process.env.NODE_ENV === "development") {
  initOpenNextCloudflareForDev();
}
