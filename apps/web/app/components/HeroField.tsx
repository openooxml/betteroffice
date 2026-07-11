"use client";

import { useEffect, useRef } from "react";

const VERT = `
attribute vec2 a_pos;
void main() {
  gl_Position = vec4(a_pos, 0.0, 1.0);
}
`;

// a document sheet lying in front of the viewer, receding to a horizon.
// paragraphs, headings, bullets, ragged right edges. the pointer is a ghost
// editor: word bars near it resolve into letter-like glyphs, a caret blinks
// on its row, a selection sweeps whole words (always ending at a space), and
// clicks send a wave that briefly turns bars into text as it passes.
const FRAG = `
precision highp float;
uniform vec2 u_res;
uniform float u_t;
uniform vec2 u_mouse;
uniform vec3 u_click;
uniform float u_scroll;
uniform sampler2D u_atlas;
uniform sampler2D u_words;
uniform float u_ready;

float hash(vec2 p) {
  return fract(sin(dot(p, vec2(127.1, 311.7))) * 43758.5453123);
}

float noise(vec2 p) {
  vec2 i = floor(p);
  vec2 f = fract(p);
  vec2 u = f * f * (3.0 - 2.0 * f);
  return mix(
    mix(hash(i), hash(i + vec2(1.0, 0.0)), u.x),
    mix(hash(i + vec2(0.0, 1.0)), hash(i + vec2(1.0, 1.0)), u.x),
    u.y
  );
}

const float PW = 16.0;
const float MARGIN = 1.15;
const float UNIT = 0.62;
const float PARA = 8.0;

// ground-plane projection: near and large at the bottom, receding to a
// horizon above the hero; the pointer nudges it for a gentle parallax
vec2 toPage(vec2 uv, float aspect, float scroll) {
  float h = 2.6 + (u_mouse.y - 0.5) * 0.14;
  float z = max(h - uv.y, 0.3);
  float x = (uv.x - 0.5 + (u_mouse.x - 0.5) * -0.05) * aspect;
  return vec2(x * 22.9 / z + PW * 0.5, 80.0 / z + scroll);
}

// per-row layout: x = indent, y = word unit, z = line length,
// w = kind (-1 blank, 0 body, 1 heading, 2 bullet)
vec4 rowMeta(float row) {
  float paraRow = mod(row, PARA);
  float paraIdx = floor(row / PARA);
  if (paraRow > PARA - 1.5) return vec4(MARGIN, UNIT, 0.0, -1.0);

  bool heading = paraRow < 0.5;
  bool bullet = hash(vec2(paraIdx * 7.7, 2.1)) > 0.55
    && paraRow >= 2.0 && paraRow <= 5.0;

  float unit = heading ? UNIT * 1.6 : UNIT;
  float indent = MARGIN + (heading ? 0.0 : (paraRow < 1.5 ? UNIT * 1.3 : 0.0));
  if (bullet) indent += 1.0;

  float maxLen = PW - MARGIN - indent;
  float len = heading
    ? maxLen * mix(0.32, 0.52, hash(vec2(paraIdx, 9.4)))
    : maxLen * mix(0.85, 1.0, hash(vec2(row * 1.31, 3.7)));
  if (bullet) len = maxLen * mix(0.45, 0.72, hash(vec2(row, 5.5)));
  if (!heading && paraRow > PARA - 2.6) len *= mix(0.4, 0.75, hash(vec2(row, 17.0)));

  return vec4(indent, unit, len, heading ? 1.0 : (bullet ? 2.0 : 0.0));
}

// real typography: picks a word from the word-bank texture and renders its
// letters by sampling the baked font atlas — four monospace chars per word
float letterCov(float cx, float fy0, vec2 seed, float unitScale) {
  float cw = 0.13 * unitScale;
  float ci = floor(cx / cw);
  if (ci < 0.0 || ci > 3.0) return 0.0;
  float fx = fract(cx / cw);

  float wordIdx = floor(hash(seed) * 64.0);
  float code = texture2D(u_words, vec2((ci + 0.5) / 4.0, (wordIdx + 0.5) / 64.0)).r * 255.0;
  if (code < 0.5) return 0.0;

  float letter = code - 1.0;
  float v = (0.22 - fy0) / 0.48;
  if (v < 0.0 || v > 1.0) return 0.0;

  float a = texture2D(u_atlas, vec2(
    (mod(letter, 8.0) + fx) / 8.0,
    (floor(letter / 8.0) + v) / 8.0
  )).a;
  return smoothstep(0.3, 0.6, a);
}

// word coverage at a page position: x = coverage, y = weight, z = bullet dot.
// morph blends solid bars into glyphs; soft is the depth-of-field width
vec3 typeset(vec2 pg, float morph, float soft) {
  float row = floor(pg.y);
  float fy0 = fract(pg.y) - 0.47;
  vec4 meta = rowMeta(row);
  if (meta.w < -0.5) return vec3(0.0);

  bool heading = meta.w > 0.5 && meta.w < 1.5;
  float band = heading ? 0.21 : 0.15;

  float dotm = 0.0;
  if (meta.w > 1.5) {
    vec2 dp = vec2(pg.x - (meta.x - 0.48), fy0);
    dotm = 1.0 - smoothstep(0.05, 0.1, length(dp * vec2(1.0, 1.9)));
  }

  float xx = pg.x - meta.x;
  if (xx < 0.0 || xx > meta.z) return vec3(0.0, 0.0, dotm);

  float c = floor(xx / meta.y);
  if (hash(vec2(c * 1.7 + 11.0, row * 2.3)) < 0.15) return vec3(0.0, 0.0, dotm);

  float fx = fract(xx / meta.y);
  float bar = smoothstep(0.12, 0.2, fx)
    * (1.0 - smoothstep(band - soft, band + soft, abs(fy0)));

  float cx = xx - (c + 0.16) * meta.y;
  float glyph = letterCov(cx, fy0, vec2(c * 7.3, row * 3.9), meta.y / UNIT)
    * step(0.0, cx) * 1.35;

  float cov = mix(bar, glyph, morph);
  return vec3(cov, heading ? 1.6 : 1.0, dotm);
}

void main() {
  vec2 uv = gl_FragCoord.xy / u_res;
  float aspect = u_res.x / u_res.y;
  float scroll = u_t * 0.16 + u_scroll * 3.0;

  vec2 pg = toPage(uv, aspect, scroll);
  vec2 m = toPage(u_mouse, aspect, scroll);
  vec2 cl = toPage(u_click.xy, aspect, scroll);

  float row = floor(pg.y);
  float fy0 = fract(pg.y) - 0.47;

  // aerial fog toward the horizon, with depth-of-field softening the bars
  float fog = smoothstep(0.45, 1.05, uv.y);
  float soft = 0.035 + fog * 0.07;

  // morph field: bars resolve into characters near the caret, and inside
  // an expanding wave where the user clicked
  float d = distance(pg * vec2(1.0, 1.5), m * vec2(1.0, 1.5));
  float focus = exp(-d * d * 0.03);
  float age = u_t - u_click.z;
  float dc = distance(pg * vec2(1.0, 1.5), cl * vec2(1.0, 1.5));
  float wave = step(0.0, age)
    * exp(-pow(abs(dc - age * 6.0), 2.0) * 0.7)
    * exp(-age * 1.1);
  float morph = clamp(smoothstep(0.3, 0.7, focus) + wave, 0.0, 1.0) * u_ready;

  vec3 ts = typeset(pg, morph, soft);

  // per-word shimmer, like edits landing somewhere in the document
  float cWord = floor(pg.x / UNIT);
  float breathe = 0.72 + 0.28 * noise(vec2(cWord * 0.9 + u_t * 0.3, row * 1.3));

  // selection snapped to whole words on the caret row, ending at a space;
  // the word count re-rolls every couple of seconds
  float rowM = floor(m.y);
  float sameRow = step(abs(row - rowM), 0.1);
  vec4 mm = rowMeta(rowM);
  float sel = 0.0;
  if (mm.w > -0.5 && mm.z > 0.0) {
    float xxm = m.x - mm.x;
    float nWords = 1.0 + floor(hash(vec2(rowM * 1.9, floor(u_t * 0.45))) * 4.0);
    float maxW = max(floor(mm.z / mm.y) - 1.0, 0.0);
    float endW = clamp(floor(xxm / mm.y), 0.0, maxW);
    float startW = max(endW - nWords + 1.0, 0.0);
    float left = mm.x + (startW + 0.1) * mm.y;
    float right = mm.x + (endW + 1.0) * mm.y + 0.03;
    sel = sameRow * step(0.0, xxm)
      * step(left, pg.x) * step(pg.x, right)
      * step(abs(fy0), 0.4);
  }

  // blinking caret at the pointer
  float caret = step(abs(pg.x - m.x), 0.05) * sameRow
    * step(abs(fy0), 0.44)
    * step(fract(u_t * 1.1), 0.55);

  // the sheet itself: hairline edges, soft cast shadow on the desk
  float inside = smoothstep(-0.04, 0.08, pg.x) * smoothstep(PW + 0.04, PW - 0.08, pg.x);
  float distOut = max(max(-pg.x, pg.x - PW), 0.0);
  float shadow = exp(-distOut * 1.6) * 0.06 * (1.0 - inside);
  float edge = (1.0 - smoothstep(0.0, 0.05, abs(pg.x)))
    + (1.0 - smoothstep(0.0, 0.05, abs(pg.x - PW)));

  // vignette: keep the copy block clean, let the page live around it
  float mask = length((uv - vec2(0.3, 0.55)) * vec2(1.85, 1.3));
  float fade = smoothstep(0.26, 0.8, mask);
  float depth = 1.0 - fog * 0.95;
  float vis = fade * depth;

  float ink = (ts.x * ts.y * breathe * (0.055 + (0.11 * focus + 0.14 * wave) * depth)
    + ts.z * (0.07 + 0.09 * focus)) * vis * inside;
  ink = min(ink, 0.3);

  vec3 col = vec3(1.0);
  col -= shadow * vis;
  col = mix(col, vec3(0.62, 0.9, 0.8), sel * 0.5 * vis * inside);
  col = mix(col, vec3(0.2), ink);
  col = mix(col, vec3(0.45), edge * 0.18 * vis);
  col = mix(col, vec3(0.13), caret * 0.7 * vis * inside);

  col += (hash(gl_FragCoord.xy * 0.7 + fract(u_t) * 17.0) - 0.5) * 0.016;
  gl_FragColor = vec4(col, 1.0);
}
`;

// paints the reactive document field behind the hero; the caret follows the
// pointer with a soft lag and roams the page on its own when idle
export function HeroField() {
  const ref = useRef<HTMLCanvasElement>(null);

  useEffect(() => {
    const canvas = ref.current;
    if (!canvas) return;
    const gl = canvas.getContext("webgl", {
      antialias: false,
      depth: false,
      stencil: false,
      powerPreference: "low-power",
    });
    if (!gl) return;

    function compile(type: number, src: string) {
      const shader = gl!.createShader(type)!;
      gl!.shaderSource(shader, src);
      gl!.compileShader(shader);
      return shader;
    }

    const program = gl.createProgram()!;
    gl.attachShader(program, compile(gl.VERTEX_SHADER, VERT));
    gl.attachShader(program, compile(gl.FRAGMENT_SHADER, FRAG));
    gl.linkProgram(program);
    if (!gl.getProgramParameter(program, gl.LINK_STATUS)) return;
    gl.useProgram(program);

    const buf = gl.createBuffer();
    gl.bindBuffer(gl.ARRAY_BUFFER, buf);
    gl.bufferData(
      gl.ARRAY_BUFFER,
      new Float32Array([-1, -1, 3, -1, -1, 3]),
      gl.STATIC_DRAW
    );
    const loc = gl.getAttribLocation(program, "a_pos");
    gl.enableVertexAttribArray(loc);
    gl.vertexAttribPointer(loc, 2, gl.FLOAT, false, 0, 0);

    const uRes = gl.getUniformLocation(program, "u_res");
    const uT = gl.getUniformLocation(program, "u_t");
    const uMouse = gl.getUniformLocation(program, "u_mouse");
    const uClick = gl.getUniformLocation(program, "u_click");
    const uScroll = gl.getUniformLocation(program, "u_scroll");
    const uAtlas = gl.getUniformLocation(program, "u_atlas");
    const uWords = gl.getUniformLocation(program, "u_words");
    const uReady = gl.getUniformLocation(program, "u_ready");

    // word bank as a 4x64 data texture: 0 = space, 1..26 = a..z
    const WORDS =
      "the and for with text docx page font word edit open save type line list grid cell form note plan tabs news all was are has not but one two who out use own see now how its may also into over such each than them then when what more most some time very well work year from have this that been were said done live free real fast"
        .split(" ");
    const wordData = new Uint8Array(4 * 64);
    for (let w = 0; w < 64; w++) {
      const word = WORDS[w % WORDS.length];
      for (let i = 0; i < 4; i++) {
        wordData[w * 4 + i] = i < word.length ? word.charCodeAt(i) - 96 : 0;
      }
    }
    const wordsTex = gl.createTexture();
    gl.activeTexture(gl.TEXTURE1);
    gl.bindTexture(gl.TEXTURE_2D, wordsTex);
    gl.pixelStorei(gl.UNPACK_ALIGNMENT, 1);
    gl.texImage2D(gl.TEXTURE_2D, 0, gl.LUMINANCE, 4, 64, 0, gl.LUMINANCE, gl.UNSIGNED_BYTE, wordData);
    gl.texParameteri(gl.TEXTURE_2D, gl.TEXTURE_MIN_FILTER, gl.NEAREST);
    gl.texParameteri(gl.TEXTURE_2D, gl.TEXTURE_MAG_FILTER, gl.NEAREST);
    gl.texParameteri(gl.TEXTURE_2D, gl.TEXTURE_WRAP_S, gl.CLAMP_TO_EDGE);
    gl.texParameteri(gl.TEXTURE_2D, gl.TEXTURE_WRAP_T, gl.CLAMP_TO_EDGE);
    gl.uniform1i(uWords, 1);

    // the baked font atlas ships as a static asset; until it arrives the
    // shader keeps rendering plain bars (u_ready gates the morph)
    let atlasReady = 0;
    const atlasTex = gl.createTexture();
    gl.activeTexture(gl.TEXTURE0);
    gl.bindTexture(gl.TEXTURE_2D, atlasTex);
    gl.texImage2D(gl.TEXTURE_2D, 0, gl.RGBA, 1, 1, 0, gl.RGBA, gl.UNSIGNED_BYTE, new Uint8Array(4));
    gl.texParameteri(gl.TEXTURE_2D, gl.TEXTURE_MIN_FILTER, gl.LINEAR);
    gl.texParameteri(gl.TEXTURE_2D, gl.TEXTURE_MAG_FILTER, gl.LINEAR);
    gl.texParameteri(gl.TEXTURE_2D, gl.TEXTURE_WRAP_S, gl.CLAMP_TO_EDGE);
    gl.texParameteri(gl.TEXTURE_2D, gl.TEXTURE_WRAP_T, gl.CLAMP_TO_EDGE);
    gl.uniform1i(uAtlas, 0);
    const atlasImg = new Image();
    atlasImg.onload = () => {
      gl.activeTexture(gl.TEXTURE0);
      gl.bindTexture(gl.TEXTURE_2D, atlasTex);
      gl.texImage2D(gl.TEXTURE_2D, 0, gl.RGBA, gl.RGBA, gl.UNSIGNED_BYTE, atlasImg);
      gl.texParameteri(gl.TEXTURE_2D, gl.TEXTURE_MIN_FILTER, gl.LINEAR_MIPMAP_LINEAR);
      gl.generateMipmap(gl.TEXTURE_2D);
      atlasReady = 1;
    };
    atlasImg.src = "/glyph-atlas.png";

    const mouse = { x: 0.75, y: 0.6 };
    const target = { x: 0.75, y: 0.6 };
    const click = { x: 0.5, y: 0.5, t: -100 };
    let lastMove = 0;
    const start = performance.now();

    function toLocal(e: PointerEvent) {
      const rect = canvas!.getBoundingClientRect();
      return {
        x: (e.clientX - rect.left) / rect.width,
        y: 1 - (e.clientY - rect.top) / rect.height,
      };
    }

    function onMove(e: PointerEvent) {
      const p = toLocal(e);
      if (p.x < -0.2 || p.x > 1.2 || p.y < -0.2 || p.y > 1.2) return;
      target.x = p.x;
      target.y = p.y;
      lastMove = performance.now();
    }

    function onDown(e: PointerEvent) {
      const p = toLocal(e);
      if (p.x < 0 || p.x > 1 || p.y < 0 || p.y > 1) return;
      click.x = p.x;
      click.y = p.y;
      click.t = (performance.now() - start) / 1000;
    }

    window.addEventListener("pointermove", onMove, { passive: true });
    window.addEventListener("pointerdown", onDown, { passive: true });

    function resize() {
      const dpr = Math.min(window.devicePixelRatio || 1, 1.5);
      const w = Math.round(canvas!.clientWidth * dpr);
      const h = Math.round(canvas!.clientHeight * dpr);
      if (canvas!.width !== w || canvas!.height !== h) {
        canvas!.width = w;
        canvas!.height = h;
        gl!.viewport(0, 0, w, h);
      }
    }

    const reduce = window.matchMedia("(prefers-reduced-motion: reduce)").matches;
    let raf = 0;
    let visible = true;

    function frame() {
      resize();
      const t = (performance.now() - start) / 1000;

      // ghost editor roams the page while the pointer is idle
      if (performance.now() - lastMove > 3500) {
        target.x = 0.5 + 0.42 * Math.sin(t * 0.19) * Math.cos(t * 0.08);
        target.y = 0.52 + 0.36 * Math.sin(t * 0.12 + 1.7);
      }
      mouse.x += (target.x - mouse.x) * 0.055;
      mouse.y += (target.y - mouse.y) * 0.055;

      gl!.uniform2f(uRes, canvas!.width, canvas!.height);
      gl!.uniform1f(uT, t);
      gl!.uniform2f(uMouse, mouse.x, mouse.y);
      gl!.uniform3f(uClick, click.x, click.y, click.t);
      gl!.uniform1f(uScroll, window.scrollY / Math.max(canvas!.clientHeight, 1));
      gl!.uniform1f(uReady, atlasReady);
      gl!.drawArrays(gl!.TRIANGLES, 0, 3);
      if (!reduce && visible) raf = requestAnimationFrame(frame);
    }

    const observer = new IntersectionObserver(([entry]) => {
      const was = visible;
      visible = entry.isIntersecting && !document.hidden;
      if (visible && !was && !reduce) raf = requestAnimationFrame(frame);
      if (!visible) cancelAnimationFrame(raf);
    });
    observer.observe(canvas);

    function onVisibility() {
      const was = visible;
      visible = !document.hidden;
      if (visible && !was && !reduce) raf = requestAnimationFrame(frame);
      if (!visible) cancelAnimationFrame(raf);
    }
    document.addEventListener("visibilitychange", onVisibility);

    raf = requestAnimationFrame(frame);

    return () => {
      cancelAnimationFrame(raf);
      observer.disconnect();
      window.removeEventListener("pointermove", onMove);
      window.removeEventListener("pointerdown", onDown);
      document.removeEventListener("visibilitychange", onVisibility);
      gl.getExtension("WEBGL_lose_context")?.loseContext();
    };
  }, []);

  return (
    <canvas
      ref={ref}
      className="pointer-events-none absolute inset-0 h-full w-full"
      aria-hidden="true"
    />
  );
}
