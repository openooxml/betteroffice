import type { NextConfig } from "next";
import { initOpenNextCloudflareForDev } from "@opennextjs/cloudflare";

const nextConfig: NextConfig = {
  reactStrictMode: true,
  transpilePackages: ["@betteroffice/xlsx", "@betteroffice/xlsx-react"],
};

export default nextConfig;

initOpenNextCloudflareForDev();
