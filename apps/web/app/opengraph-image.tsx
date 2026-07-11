import { ImageResponse } from "next/og";

export const size = { width: 1200, height: 630 };
export const contentType = "image/png";
export const alt = "BetterOffice — The open-source office suite";

// og card: logo mark, wordmark, one-liner on plain white
export default function OpengraphImage() {
  return new ImageResponse(
    (
      <div
        style={{
          width: "100%",
          height: "100%",
          display: "flex",
          flexDirection: "column",
          justifyContent: "center",
          padding: "0 96px",
          backgroundColor: "#ffffff",
          color: "#141414",
          fontFamily: "sans-serif",
        }}
      >
        <div style={{ fontSize: 84, fontWeight: 700, letterSpacing: -3 }}>
          BetterOffice
        </div>
        <div style={{ fontSize: 34, color: "#5c5c5c", marginTop: 18 }}>
          The open-source office suite — on engines built in Rust
        </div>
      </div>
    ),
    size
  );
}
