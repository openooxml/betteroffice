import { ImageResponse } from "next/og";
import { logoUri } from "./og-logo";

export const size = { width: 1200, height: 630 };
export const contentType = "image/png";
export const alt = "BetterOffice — The open-source office suite";

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
          position: "relative",
          overflow: "hidden",
        }}
      >
        <img
          src={logoUri}
          width={760}
          height={567}
          style={{
            position: "absolute",
            right: -170,
            top: -80,
            opacity: 0.045,
            transform: "rotate(-6deg)",
          }}
        />

        <img src={logoUri} width={150} height={112} style={{ marginBottom: 44 }} />
        <div style={{ fontSize: 84, fontWeight: 700, letterSpacing: -3 }}>
          BetterOffice
        </div>
        <div style={{ fontSize: 33, color: "#5c5c5c", marginTop: 16 }}>
          The open-source office suite — on engines built in Rust
        </div>
      </div>
    ),
    size
  );
}
