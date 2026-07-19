import { flatSquareBadge } from "../../../lib/badge";

const ALLOWED_LOGOS = new Set(["npm", "rust"]);
const HEX = /^[0-9a-fA-F]{3,8}$/;

function color(value: string | null, fallback: string): string {
  if (!value) return fallback;
  return HEX.test(value) ? `#${value}` : fallback;
}

export function GET(request: Request) {
  const params = new URL(request.url).searchParams;
  const label = (params.get("label") ?? "").slice(0, 64);
  const message = (params.get("message") ?? "").slice(0, 64);
  const logoParam = params.get("logo");
  const logo = logoParam && ALLOWED_LOGOS.has(logoParam) ? logoParam : undefined;

  const svg = flatSquareBadge({
    label,
    message,
    color: color(params.get("color"), "#4ade80"),
    labelColor: color(params.get("labelColor"), "#555"),
    logo,
  });

  return new Response(svg, {
    headers: {
      "Content-Type": "image/svg+xml; charset=utf-8",
      "Cache-Control": "public, max-age=86400, s-maxage=86400",
    },
  });
}
