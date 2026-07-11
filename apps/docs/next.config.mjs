import { createMDX } from "fumadocs-mdx/next";
import { initOpenNextCloudflareForDev } from "@opennextjs/cloudflare";

const withMDX = createMDX();

/** @type {import('next').NextConfig} */
const config = {
  reactStrictMode: true,
};

export default withMDX(config);

// unguarded init crashes `next build` with EPIPE when combined with fumadocs-mdx
if (process.env.NODE_ENV === "development") {
  initOpenNextCloudflareForDev();
}
