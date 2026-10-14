"use client";

import { useEffect, useRef, type CSSProperties, type RefObject } from "react";
import { Bot, User } from "lucide-react";

// one page, two peers, one review: a person selects and types, an agent proposes a tracked change,
// the person accepts it, the agent makes a second edit that stays pending, then the page resets
const LOOP = 16;
// the frame shown before the loop starts and under reduced motion: a tracked change awaiting review
const STILL = 9.6;

type Ease = "io" | "lin" | number;
type Key = [t: number, v: number, ease?: Ease];
type Track = { x?: Key[]; y?: Key[]; sx?: Key[]; s?: Key[]; o?: Key[] };

const TYPE_A = 52;
const TYPE_B = 58;
const DEL = 48;
const INS = 40;
const SHIFT = INS + 6;
const CLOSE = DEL + 6;
const CELL = 34;

const TRACKS: Record<string, Track> = {
  // person: point, click, drag a selection, type over it
  ptr: {
    x: [[0.6, 330], [1.3, 200.5], [1.5, 200.5], [2.3, 316, "lin"], [2.5, 316], [2.9, 350], [8.4, 350], [9.6, 433]],
    y: [[0.6, 64], [1.3, 103], [2.5, 103], [2.9, 132], [8.4, 132], [9.6, 175]],
    o: [[0.3, 0], [0.6, 1], [2.7, 1], [3, 0], [8.3, 0], [8.6, 1], [10.3, 1], [10.6, 0]],
    s: [[1.38, 1], [1.42, 0.85], [1.52, 1], [9.78, 1], [9.82, 0.85], [9.92, 1]],
  },
  clk1: { s: [[1.4, 0.4], [1.9, 1.6, "lin"]], o: [[1.39, 0], [1.4, 0.6], [1.9, 0]] },
  sel: { sx: [[1.5, 0], [2.3, 1, "lin"]], o: [[1.45, 0], [1.5, 1], [2.45, 1], [2.5, 0]] },
  w1: { o: [[2.45, 1], [2.5, 0], [15, 0], [16, 1]] },
  w2: { o: [[2.45, 1], [2.5, 0], [15, 0], [16, 1]] },
  t1: { sx: [[2.6, 0], [3.5, 1, 8]], o: [[15, 1], [16, 0]] },
  t2: { sx: [[3.5, 0], [4.4, 1, 8]], o: [[15, 1], [16, 0]] },
  pc: {
    x: [[1.5, 0], [2.3, 116, "lin"], [2.45, 116], [2.5, 0], [2.6, 0], [3.5, TYPE_A, 8], [3.5, TYPE_A + 6], [4.4, 116, 8]],
    o: [[1.35, 0], [1.4, 1], [15, 1], [15.6, 0]],
  },
  // agent: strike a word, insert its replacement, with a change bar and a proposal card
  ac: {
    x: [[5, 250], [5.8, 0], [6.5, CLOSE, "lin"], [7.6, CLOSE + INS, 6], [10.2, CLOSE + INS], [10.8, INS], [11, INS], [11.8, 186]],
    y: [[5, -40], [5.8, 0], [11, 0], [11.8, 104]],
    o: [[5, 0], [5.3, 1], [15, 1], [15.6, 0]],
  },
  strike: { sx: [[5.8, 0], [6.5, 1, "lin"]], o: [[5.75, 0], [5.8, 1], [10.2, 1], [10.6, 0]] },
  del: {
    sx: [[10.2, 1], [10.8, 0], [14.99, 0], [15, 1]],
    o: [[5.8, 1], [6.3, 0.5], [10.2, 0.5], [10.8, 0], [15, 0], [16, 1]],
  },
  ins: { sx: [[6.5, 0], [7.6, 1, 6]], x: [[10.2, 0], [10.8, -CLOSE]], o: [[10.4, 1], [10.9, 0]] },
  insu: { sx: [[6.5, 0], [7.6, 1, 6]], x: [[10.2, 0], [10.8, -CLOSE]], o: [[10.2, 1], [10.6, 0]] },
  insk: { x: [[10.2, 0], [10.8, -CLOSE]], o: [[10.4, 0], [10.9, 1], [15, 1], [16, 0]] },
  tr: { x: [[6.5, 0], [7.6, SHIFT, 6], [10.2, SHIFT], [10.8, SHIFT - CLOSE], [15, SHIFT - CLOSE], [16, 0]] },
  bar1: { o: [[6.6, 0], [7, 1], [10.3, 1], [10.8, 0]] },
  card: { x: [[7.8, 18], [8.4, 0], [10.2, 0], [10.7, 18]], o: [[7.8, 0], [8.4, 1], [10.2, 1], [10.7, 0]] },
  chk: { o: [[9.8, 0], [9.9, 1]] },
  clk2: { s: [[9.8, 0.4], [10.3, 1.6, "lin"]], o: [[9.79, 0], [9.8, 0.6], [10.3, 0]] },
  // agent: a second edit in a table cell, left pending
  old: { x: [[12.4, 0], [13.2, -CELL, 5], [15, -CELL], [16, 0]], o: [[11.8, 1], [12.3, 0.5], [15, 0.5], [16, 1]] },
  oldx: {
    x: [[12.4, 0], [13.2, -CELL, 5], [15, -CELL], [16, 0]],
    sx: [[11.8, 0], [12.3, 1, "lin"]],
    o: [[11.75, 0], [11.8, 1], [15, 1], [15.8, 0]],
  },
  new: { sx: [[12.4, 0], [13.2, 1, 5]], o: [[15, 1], [15.8, 0]] },
  newu: { sx: [[12.4, 0], [13.2, 1, 5]], o: [[15, 1], [15.8, 0]] },
  bar2: { o: [[12.4, 0], [12.8, 1], [15, 1], [15.8, 0]] },
};

function sample(keys: Key[] | undefined, t: number, rest: number): number {
  if (!keys) return rest;
  if (t <= keys[0][0]) return keys[0][1];
  for (let i = 1; i < keys.length; i++) {
    const [t1, v1, ease = "io"] = keys[i];
    if (t > t1) continue;
    const [t0, v0] = keys[i - 1];
    const p = t1 === t0 ? 1 : (t - t0) / (t1 - t0);
    const q =
      typeof ease === "number"
        ? Math.floor(p * Math.round(ease)) / Math.round(ease)
        : ease === "lin"
          ? p
          : p < 0.5
            ? 4 * p * p * p
            : 1 - (-2 * p + 2) ** 3 / 2;
    return v0 + (v1 - v0) * q;
  }
  return keys[keys.length - 1][1];
}

const r3 = (v: number) => Math.round(v * 1000) / 1000;

function frame(k: string, t: number): { transform: string; opacity: string } {
  const tr = TRACKS[k];
  const s = sample(tr.s, t, 1);
  return {
    transform: `translate(${r3(sample(tr.x, t, 0))}px, ${r3(sample(tr.y, t, 0))}px) scale(${r3(sample(tr.sx, t, 1) * s)}, ${r3(s)})`,
    opacity: String(r3(sample(tr.o, t, 1))),
  };
}

function anim(k: string, origin = "center"): { "data-k": string; style: CSSProperties } {
  return {
    "data-k": k,
    style: { transformBox: "fill-box", transformOrigin: origin, ...frame(k, STILL) },
  };
}

const word = { fill: "currentColor", fillOpacity: 0.26, height: 4 } as const;

function Row({ y, words }: { y: number; words: [number, number][] }) {
  return (
    <>
      {words.map(([x, w]) => (
        <rect key={x} x={x} y={y} width={w} {...word} />
      ))}
    </>
  );
}

function useLoop(ref: RefObject<SVGSVGElement | null>) {
  useEffect(() => {
    const svg = ref.current;
    if (!svg) return;
    const els = [...svg.querySelectorAll<SVGElement>("[data-k]")].map((el) => ({
      el,
      k: el.dataset.k as string,
      last: "",
    }));
    const reduce = window.matchMedia("(prefers-reduced-motion: reduce)");
    let t = STILL;
    let raf = 0;
    let prev = 0;
    let visible = false;

    const paint = (time: number) => {
      for (const item of els) {
        const f = frame(item.k, time);
        const key = f.transform + f.opacity;
        if (key === item.last) continue;
        item.last = key;
        item.el.style.transform = f.transform;
        item.el.style.opacity = f.opacity;
      }
    };
    const tick = (now: number) => {
      t = (t + Math.min(0.1, (now - prev) / 1000)) % LOOP;
      prev = now;
      paint(t);
      raf = requestAnimationFrame(tick);
    };
    const sync = () => {
      cancelAnimationFrame(raf);
      raf = 0;
      if (reduce.matches) {
        t = STILL;
        paint(t);
        return;
      }
      if (!visible || document.hidden) return;
      prev = performance.now();
      raf = requestAnimationFrame(tick);
    };

    const io = new IntersectionObserver(([entry]) => {
      visible = entry.isIntersecting;
      sync();
    });
    io.observe(svg);
    document.addEventListener("visibilitychange", sync);
    reduce.addEventListener("change", sync);
    return () => {
      cancelAnimationFrame(raf);
      io.disconnect();
      document.removeEventListener("visibilitychange", sync);
      reduce.removeEventListener("change", sync);
    };
  }, [ref]);
}

export function CoEdit() {
  const ref = useRef<SVGSVGElement>(null);
  useLoop(ref);
  return (
    <svg ref={ref} viewBox="20 8 520 304" className="block h-auto w-full" aria-hidden="true">
      <rect x="40.5" y="24.5" width="360" height="300" fill="var(--color-bg)" stroke="currentColor" strokeOpacity={0.22} />
      <rect x="72" y="52" width="150" height="8" fill="currentColor" fillOpacity={0.8} />

      <Row y={86} words={[[72, 48], [126, 70], [202, 36], [244, 62], [312, 56]]} />
      <Row y={101} words={[[72, 32], [110, 84], [322, 46]]} />
      <rect {...anim("sel", "left center")} x="198" y="97" width="120" height="12" fill="currentColor" fillOpacity={0.1} />
      <rect {...anim("w1")} x="200" y="101" width="44" {...word} />
      <rect {...anim("w2")} x="250" y="101" width="66" {...word} />
      <rect {...anim("t1", "left center")} x="200" y="101" width={TYPE_A} {...word} />
      <rect {...anim("t2", "left center")} x={200 + TYPE_A + 6} y="101" width={TYPE_B} {...word} />
      <Row y={116} words={[[72, 66], [144, 40], [190, 88]]} />

      <Row y={144} words={[[72, 42], [120, 60], [186, 48], [240, 78], [324, 44]]} />
      <Row y={159} words={[[72, 56], [134, 34]]} />
      <rect {...anim("del", "left center")} x="174" y="159" width={DEL} {...word} />
      <line {...anim("strike", "left center")} x1="172" x2={176 + DEL} y1="161" y2="161" stroke="currentColor" strokeOpacity={0.6} strokeWidth={1.25} />
      <g {...anim("tr")}>
        <Row y={159} words={[[228, 38], [272, 44]]} />
      </g>
      <rect {...anim("insk")} x="228" y="159" width={INS} {...word} />
      <rect {...anim("ins", "left center")} x="228" y="159" width={INS} height="4" fill="var(--color-acc)" fillOpacity={0.85} />
      <line {...anim("insu", "left center")} x1="228" x2={228 + INS} y1="166" y2="166" stroke="var(--color-acc)" strokeOpacity={0.7} />
      <rect {...anim("bar1")} x="55" y="153" width="1.75" height="16" fill="var(--color-acc)" />
      <Row y={174} words={[[72, 74], [152, 50], [208, 90], [304, 42]]} />
      <Row y={189} words={[[72, 38], [116, 66]]} />

      <rect x="72.5" y="210.5" width="296" height="66" fill="none" stroke="currentColor" strokeOpacity={0.3} />
      <rect x="73" y="211" width="295" height="21" fill="currentColor" fillOpacity={0.05} />
      <path d="M72.5 232.5H368.5M72.5 254.5H368.5M168.5 210.5V276.5M264.5 210.5V276.5" stroke="currentColor" strokeOpacity={0.18} />
      <Row y={219} words={[[80, 40], [176, 34], [322, 38]]} />
      <Row y={241} words={[[80, 56], [176, 48], [334, 26]]} />
      <Row y={263} words={[[80, 44], [176, 60]]} />
      <rect {...anim("old")} x="330" y="263" width="30" {...word} />
      <line {...anim("oldx", "left center")} x1="328" x2="362" y1="265" y2="265" stroke="currentColor" strokeOpacity={0.6} strokeWidth={1.25} />
      <rect {...anim("new", "right center")} x="332" y="263" width="28" height="4" fill="var(--color-acc)" fillOpacity={0.85} />
      <line {...anim("newu", "right center")} x1="332" x2="360" y1="270" y2="270" stroke="var(--color-acc)" strokeOpacity={0.7} />
      <rect {...anim("bar2")} x="55" y="256" width="1.75" height="18" fill="var(--color-acc)" />
      <Row y={296} words={[[72, 60], [138, 82], [226, 40]]} />

      <g {...anim("card")}>
        <rect x="412.5" y="140.5" width="116" height="46" fill="var(--color-bg)" stroke="currentColor" strokeOpacity={0.22} />
        <rect x="422" y="150" width="14" height="12" fill="var(--color-acc)" />
        <Bot x={424} y={151} width={10} height={10} color="var(--color-bg)" strokeWidth={2.4} />
        <rect x="442" y="151" width="58" height="3" fill="currentColor" fillOpacity={0.3} />
        <rect x="442" y="157" width="40" height="3" fill="currentColor" fillOpacity={0.2} />
        <rect {...anim("chk")} x="422" y="168" width="20" height="12" fill="var(--color-acc)" fillOpacity={0.22} />
        <rect x="422.5" y="168.5" width="20" height="12" fill="none" stroke="var(--color-acc)" strokeOpacity={0.8} />
        <path d="M428 174.5l2.6 2.6 5-5" fill="none" stroke="var(--color-acc)" strokeWidth={1.4} />
        <rect x="446.5" y="168.5" width="20" height="12" fill="none" stroke="currentColor" strokeOpacity={0.3} />
        <path d="M453.5 171.5l6 6M459.5 171.5l-6 6" stroke="currentColor" strokeOpacity={0.55} strokeWidth={1.2} />
      </g>

      <g {...anim("pc")}>
        <rect className="blink" x="199.25" y="96" width="1.5" height="14" fill="currentColor" />
        <rect x="199.25" y="81" width="17" height="15" fill="currentColor" />
        <User x={202} y={83} width={11} height={11} color="var(--color-bg)" strokeWidth={2.2} />
      </g>
      <g {...anim("ac")}>
        <rect className="blink" style={{ animationDelay: "0.55s" }} x="173.25" y="154" width="1.5" height="14" fill="var(--color-acc)" />
        <rect x="173.25" y="139" width="17" height="15" fill="var(--color-acc)" />
        <Bot x={176} y={141} width={11} height={11} color="var(--color-bg)" strokeWidth={2.2} />
      </g>

      <circle {...anim("clk1")} cx="200.5" cy="103" r="7" fill="none" stroke="currentColor" strokeOpacity={0.5} />
      <circle {...anim("clk2")} cx="433" cy="175" r="7" fill="none" stroke="var(--color-acc)" />
      <g {...anim("ptr", "left top")}>
        <path
          d="M0 0V15.5L3.9 11.8L6.6 17.6L9.1 16.5L6.5 10.9H11.6Z"
          fill="currentColor"
          stroke="var(--color-bg)"
          strokeWidth={1}
          strokeLinejoin="round"
        />
      </g>
    </svg>
  );
}
