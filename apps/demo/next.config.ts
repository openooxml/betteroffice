import type { NextConfig } from "next";
import { initOpenNextCloudflareForDev } from "@opennextjs/cloudflare";

const nextConfig: NextConfig = {
  reactStrictMode: true,
  transpilePackages: [
    "@betteroffice/xlsx",
    "@betteroffice/xlsx-react",
    "@betteroffice/fonts",
    "@betteroffice/fonts-cjk",
    "@betteroffice/vsdx",
    "@betteroffice/vsdx-react",
  ],
};

export default nextConfig;

initOpenNextCloudflareForDev();
