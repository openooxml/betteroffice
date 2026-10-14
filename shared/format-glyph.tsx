// line drawings of each editor's working surface, drawn in the text color
import type { ReactNode } from "react";

const line = {
  stroke: "currentColor",
  strokeWidth: 1,
  vectorEffect: "non-scaling-stroke",
  fill: "none",
} as const;

function Bars({
  x,
  y,
  widths,
  gap = 9,
  h = 3,
  opacity = 0.32,
}: {
  x: number;
  y: number;
  widths: number[];
  gap?: number;
  h?: number;
  opacity?: number;
}) {
  return (
    <g fill="currentColor" fillOpacity={opacity}>
      {widths.map((w, i) => (
        <rect key={i} x={x} y={y + i * gap} width={w} height={h} />
      ))}
    </g>
  );
}

function Docx() {
  return (
    <>
      <rect x="62" y="16" width="116" height="150" {...line} fill="var(--color-bg)" strokeOpacity={0.45} />
      <rect x="76" y="32" width="54" height="6" fill="currentColor" fillOpacity={0.75} />
      <Bars x={76} y={48} widths={[88, 84, 88, 52]} />
      <Bars x={76} y={92} widths={[88, 80, 34]} />
      <rect x="76" y="117" width="30" height="3" fill="currentColor" fillOpacity={0.32} />
      <line x1="76" x2="106" y1="118.5" y2="118.5" stroke="var(--color-acc)" />
      <rect x="110" y="117" width="30" height="3" fill="var(--color-acc)" fillOpacity={0.9} />
      <rect x="143" y="113" width="1.5" height="11" fill="currentColor" />
      <Bars x={76} y={130} widths={[88, 70]} />
    </>
  );
}

function Xlsx() {
  const cols = [44, 70, 100, 130, 160, 190];
  const rows = [36, 52, 68, 84, 100, 116, 132, 148];
  const values: [number, number, number][] = [
    [1, 1, 14], [2, 1, 20], [3, 1, 12], [4, 1, 18],
    [1, 2, 16], [2, 2, 10], [3, 2, 22], [4, 2, 14],
    [1, 3, 20], [2, 3, 16], [3, 3, 12], [4, 3, 18],
    [1, 4, 12], [2, 4, 18], [4, 4, 16],
    [1, 5, 18], [2, 5, 14], [3, 5, 20], [4, 5, 12],
  ];
  return (
    <>
      <rect x="20" y="14" width="200" height="12" {...line} strokeOpacity={0.35} />
      <text x="25" y="23" fontSize="7" fontFamily="var(--font-mono)" fill="currentColor" fillOpacity={0.55}>
        fx
      </text>
      <rect x="36" y="18" width="70" height="3" fill="currentColor" fillOpacity={0.3} />
      <rect x="20" y="36" width="200" height="16" fill="currentColor" fillOpacity={0.05} />
      <rect x="20" y="36" width="24" height="128" fill="currentColor" fillOpacity={0.05} />
      {cols.map((x) => (
        <line key={x} x1={x} x2={x} y1={36} y2={164} {...line} strokeOpacity={0.18} />
      ))}
      {rows.map((y) => (
        <line key={y} x1={20} x2={220} y1={y} y2={y} {...line} strokeOpacity={0.18} />
      ))}
      <rect x="20" y="36" width="200" height="128" {...line} strokeOpacity={0.4} />
      <g fill="currentColor" fillOpacity={0.34}>
        {values.map(([c, r, w]) => (
          <rect key={`${c}-${r}`} x={cols[c] - w - 5} y={rows[r] + 6.5} width={w} height={3} />
        ))}
      </g>
      <rect x="130" y="100" width="30" height="16" fill="var(--color-acc-soft)" />
      <rect x="130" y="100" width="30" height="16" fill="none" stroke="var(--color-acc)" strokeWidth={1.5} vectorEffect="non-scaling-stroke" />
      <rect x="157.5" y="113.5" width="5" height="5" fill="var(--color-acc)" />
    </>
  );
}

function Pptx() {
  return (
    <>
      {[22, 58, 94, 130].map((y, i) => (
        <rect
          key={y}
          x="20"
          y={y}
          width="44"
          height="25"
          {...line}
          strokeOpacity={i === 1 ? 0.9 : 0.3}
        />
      ))}
      <rect x="25" y="28" width="22" height="3" fill="currentColor" fillOpacity={0.35} />
      <rect x="25" y="64" width="22" height="3" fill="currentColor" fillOpacity={0.6} />
      <rect x="25" y="100" width="22" height="3" fill="currentColor" fillOpacity={0.35} />
      <rect x="25" y="136" width="22" height="3" fill="currentColor" fillOpacity={0.35} />
      <rect x="76" y="40" width="144" height="81" {...line} strokeOpacity={0.5} />
      <rect x="88" y="54" width="66" height="7" fill="currentColor" fillOpacity={0.75} />
      <Bars x={88} y={70} widths={[54, 46, 50]} gap={8} />
      <circle cx="186" cy="88" r="16" fill="var(--color-acc-soft)" stroke="var(--color-acc)" strokeWidth={1.25} vectorEffect="non-scaling-stroke" />
      <g fill="var(--color-bg)" stroke="var(--color-acc)" strokeWidth={1} vectorEffect="non-scaling-stroke">
        <rect x="167.5" y="69.5" width="4" height="4" />
        <rect x="200.5" y="69.5" width="4" height="4" />
        <rect x="167.5" y="102.5" width="4" height="4" />
        <rect x="200.5" y="102.5" width="4" height="4" />
      </g>
      <rect x="76" y="134" width="144" height="24" {...line} strokeOpacity={0.2} />
      <Bars x={84} y={141} widths={[80, 52]} gap={7} h={2.5} opacity={0.22} />
    </>
  );
}

function Vsdx() {
  return (
    <>
      <g {...line} strokeOpacity={0.45}>
        <path d="M70 96 H90" />
        <path d="M150 96 H160 V58 H176" />
        <path d="M150 96 H160 V134 H170" />
      </g>
      <g fill="currentColor" fillOpacity={0.45}>
        <path d="M86 93 L92 96 L86 99 Z" />
        <path d="M172 55 L178 58 L172 61 Z" />
        <path d="M164 131 L170 134 L164 137 Z" />
      </g>
      <rect x="26" y="84" width="44" height="24" {...line} fill="var(--color-bg)" strokeOpacity={0.6} />
      <rect x="34" y="94.5" width="28" height="3" fill="currentColor" fillOpacity={0.4} />
      <path d="M120 72 L150 96 L120 120 L90 96 Z" {...line} fill="var(--color-bg)" strokeOpacity={0.6} />
      <rect x="178" y="46" width="40" height="24" rx="12" {...line} fill="var(--color-bg)" strokeOpacity={0.6} />
      <rect x="170" y="122" width="46" height="24" fill="var(--color-acc-soft)" stroke="var(--color-acc)" strokeWidth={1.25} vectorEffect="non-scaling-stroke" />
      <rect x="178" y="132.5" width="30" height="3" fill="var(--color-acc)" fillOpacity={0.8} />
      <g fill="var(--color-bg)" stroke="var(--color-acc)" strokeWidth={1} vectorEffect="non-scaling-stroke">
        <rect x="167.5" y="119.5" width="4" height="4" />
        <rect x="214.5" y="119.5" width="4" height="4" />
        <rect x="167.5" y="144.5" width="4" height="4" />
        <rect x="214.5" y="144.5" width="4" height="4" />
      </g>
    </>
  );
}

const GLYPHS: Record<string, () => ReactNode> = {
  docx: Docx,
  xlsx: Xlsx,
  pptx: Pptx,
  vsdx: Vsdx,
};

export function FormatGlyph({ format }: { format: string }) {
  const Glyph = GLYPHS[format];
  if (!Glyph) return null;
  return (
    <svg
      viewBox="0 0 240 180"
      className="h-full w-full"
      preserveAspectRatio="xMidYMid meet"
      aria-hidden="true"
    >
      <Glyph />
    </svg>
  );
}
