"use client";

import { useEffect, useRef } from "react";

const PAGE_WIDTH = 16;
const FLOW = 0.2;
const SCROLL_ROWS = 3;
const AGENT_EVERY = 6.5;
const STILL = 14;

const VERT = `
attribute vec2 a_pos;
void main() {
  gl_Position = vec4(a_pos, 0.0, 1.0);
}
`;

// a stream of document pages lying on a ground plane, receding to a misted
// horizon above the hero copy and drifting toward the viewer. a layout front
// crosses the sheet: beyond it the pages are still line boxes, nearer rows are
// typeset left to right as they pass it. the pointer is a reading lens with a
// caret and a word selection, an agent caret types tracked insertions, and a
// click sends a wave that turns words into glyphs.
const FRAG = `
precision highp float;

uniform vec2 u_res;
uniform float u_t;
uniform float u_flow;
uniform vec2 u_mouse;
uniform vec2 u_pointer;
uniform vec3 u_click;
uniform vec4 u_view;
uniform float u_frontY;
uniform vec4 u_clear;
uniform vec3 u_agent;
uniform float u_intro;
uniform float u_dark;
uniform float u_px;
uniform float u_ready;
uniform sampler2D u_atlas;
uniform sampler2D u_words;

const float PW = ${PAGE_WIDTH.toFixed(1)};
const float PH = 21.5;
const float PERIOD = 23.0;
const float TOP = 2.75;
const float LINES = 16.0;
const float ML = 1.6;
const float UNIT = 0.62;
const float PARA = 8.0;

float hash(vec2 p) {
  return fract(sin(dot(p, vec2(127.1, 311.7))) * 43758.5453123);
}

// box-filtered coverage of [a, b] over a pixel fp wide, so edges stay clean at any depth
float span(float x, float a, float b, float fp) {
  return clamp((min(x + 0.5 * fp, b) - max(x - 0.5 * fp, a)) / fp, 0.0, 1.0);
}

float rect(vec2 p, vec4 r, vec2 fp) {
  return span(p.x, r.x, r.z, fp.x) * span(p.y, r.y, r.w, fp.y);
}

float frame(vec2 p, vec4 r, vec2 fp) {
  vec2 w = fp * u_px * 0.5;
  return rect(p, r + vec4(-w, w), fp) - rect(p, r + vec4(w, -w), fp);
}

// screen uv to page space: x across the sheet, y down the document
vec2 plane(vec2 uv, float flow, float lean) {
  float k = max(u_view.x - uv.y, 0.004);
  float x = (uv.x - 0.5 + lean) * (u_res.x / u_res.y) * u_view.y / k + PW * 0.5;
  return vec2(x, -u_view.z / k - flow);
}

// x = global text row, y = position inside the line (0 at its top), z = 1 inside a page's text block
vec3 rowAt(float y) {
  float page = floor(y / PERIOD);
  float ry = y - page * PERIOD - TOP;
  float line = floor(ry);
  return vec3(page * LINES + line, ry - line, step(0.0, ry) * step(ry, LINES - 0.001));
}

// x = indent, y = word unit, z = length,
// w = kind (-1 blank, 0 body, 1 heading, 2 bullet, 3 table, 4 figure)
vec4 rowMeta(float row) {
  float pr = mod(row, PARA);
  float block = floor(row / PARA);
  float kind = hash(vec2(block, 41.3));
  vec4 blank = vec4(ML, UNIT, 0.0, -1.0);
  if (kind > 0.86) return pr < 6.5 ? vec4(ML, UNIT, PW - 2.0 * ML, 3.0) : blank;
  if (kind > 0.74) {
    if (pr < 5.5) return vec4(ML, UNIT, PW - 2.0 * ML, 4.0);
    float cap = mix(3.0, 5.0, hash(vec2(block, 3.1)));
    return pr < 6.5 ? vec4((PW - cap) * 0.5, UNIT * 0.8, cap, 0.0) : blank;
  }
  if (pr > PARA - 1.5) return blank;

  bool heading = pr < 0.5;
  bool bullet = hash(vec2(block * 7.7, 2.1)) > 0.55 && pr >= 2.0 && pr <= 5.0;
  float unit = heading ? UNIT * 1.6 : UNIT;
  float indent = ML + (heading ? 0.0 : (pr < 1.5 ? UNIT * 1.3 : 0.0));
  if (bullet) indent += 1.0;
  float maxLen = PW - ML - indent;
  float len = heading
    ? maxLen * mix(0.32, 0.52, hash(vec2(block, 9.4)))
    : maxLen * mix(0.85, 1.0, hash(vec2(row * 1.31, 3.7)));
  if (bullet) len = maxLen * mix(0.45, 0.72, hash(vec2(row, 5.5)));
  if (!heading && pr > PARA - 2.6) len *= mix(0.4, 0.75, hash(vec2(row, 17.0)));
  return vec4(indent, unit, len, heading ? 1.0 : (bullet ? 2.0 : 0.0));
}

// a word from the word bank, four monospace letters from the baked font atlas
float letterCov(float cx, float fy, vec2 seed, float scale) {
  float cw = 0.13 * scale;
  float ci = floor(cx / cw);
  if (cx < 0.0 || ci > 3.0) return 0.0;
  float fx = fract(cx / cw);
  float word = floor(hash(seed) * 64.0);
  float code = texture2D(u_words, vec2((ci + 0.5) / 4.0, (word + 0.5) / 64.0)).r * 255.0;
  if (code < 0.5) return 0.0;
  float letter = code - 1.0;
  float v = (fy + 0.22) / 0.48;
  if (v < 0.0 || v > 1.0) return 0.0;
  float a = texture2D(u_atlas, vec2(
    (mod(letter, 8.0) + fx) / 8.0,
    (floor(letter / 8.0) + v) / 8.0
  )).a;
  return smoothstep(0.3, 0.6, a);
}

void main() {
  vec2 uv = gl_FragCoord.xy / u_res;
  float aspect = u_res.x / u_res.y;
  float k = u_view.x - uv.y;
  vec3 bg = mix(vec3(1.0), vec3(0.0353), u_dark);

  // clean paper around and below the copy block, mist toward the horizon
  vec2 hs = max((u_clear.zw - u_clear.xy) * 0.5, vec2(0.0));
  vec2 q = (abs(uv - (u_clear.xy + u_clear.zw) * 0.5) - hs) * vec2(aspect, 1.0);
  float around = u_clear.z > u_clear.x ? smoothstep(0.02, 0.2, length(max(q, 0.0))) : 1.0;
  float above = smoothstep(u_clear.w - 0.005, u_clear.w + 0.09, uv.y);
  float vis = around * above * smoothstep(0.004, 0.16, k) * u_intro;
  if (vis < 0.003) {
    gl_FragColor = vec4(bg, 1.0);
    return;
  }

  // page units per device pixel, across and down the sheet
  vec2 fp = vec2(u_view.y / k, u_view.z / (k * k)) / u_res.y;
  float hair = fp.y * u_px;
  float lean = (0.5 - u_mouse.x) * 0.03;
  vec2 p = plane(uv, u_flow, lean);
  float page = floor(p.y / PERIOD);
  vec2 pp = vec2(p.x, p.y - page * PERIOD);

  float paper = rect(pp, vec4(0.0, 0.0, PW, PH), fp);
  float edge = frame(pp, vec4(0.0, 0.0, PW, PH), fp);
  vec2 off = max(abs(pp - vec2(PW * 0.5 + 0.06, PH * 0.5 + 0.2)) - vec2(PW, PH) * 0.5, 0.0);
  float shadow = exp(-length(off) * 6.0) * (1.0 - paper);
  vec3 desk = mix(vec3(0.988), vec3(0.0353), u_dark);
  if (paper + edge < 0.002) {
    gl_FragColor = vec4(mix(bg, desk * (1.0 - shadow * 0.05 * (1.0 - u_dark)), vis), 1.0);
    return;
  }
  float chrome = 0.0;
  if (pp.y < 2.0 || pp.y > PH - 1.6) {
    chrome = rect(pp, vec4(ML, 1.1, ML + 2.4, 1.32), fp)
      + span(pp.y, 1.8, 1.8 + hair, fp.y) * span(pp.x, ML, PW - ML, fp.x) * 0.7
      + rect(pp, vec4(PW * 0.5 - 0.2, PH - 1.32, PW * 0.5 + 0.2, PH - 1.1), fp);
  }

  vec3 r = rowAt(p.y);
  float row = r.x;
  float fy = r.y - 0.5;
  vec4 meta = r.z > 0.5 && paper > 0.002 ? rowMeta(row) : vec4(ML, UNIT, 0.0, -1.0);
  float kind = meta.w;

  float laid = clamp((p.y - r.y + 0.5 - u_frontY) * 3.0, 0.0, 1.0);

  // the pointer is a reading lens of constant screen size; a click sends a wave across the sheet
  vec2 ld = (uv - u_mouse) * vec2(aspect, 1.0);
  float lens = smoothstep(0.35, 0.85, exp(-dot(ld, ld) * 70.0));
  float age = u_t - u_click.z;
  float wave = 0.0;
  if (age >= 0.0 && age < 4.0) {
    float dc = distance(p * vec2(1.0, 1.5), u_click.xy * vec2(1.0, 1.5));
    wave = exp(-pow(abs(dc - age * 6.0), 2.0) * 0.7) * exp(-age * 1.1);
  }
  // glyphs only where a line is tall enough to read them; smaller rows keep their bars
  float morph = clamp(lens + wave, 0.0, 1.0) * u_ready * smoothstep(10.0, 18.0, 1.0 / (fp.y * u_px));

  float ink = 0.0;
  float wire = 0.0;
  float grid = 0.0;
  float shade = 0.0;
  float accInk = 0.0;
  float accMark = 0.0;
  float sel = 0.0;
  float caret = 0.0;
  float weight = 1.0;

  if (kind > -0.5 && kind < 2.5) {
    bool heading = kind > 0.5 && kind < 1.5;
    float band = heading ? 0.21 : 0.15;
    weight = heading ? 1.5 : 1.0;
    float xx = p.x - meta.x;
    float c = floor(xx / meta.y);
    float fx = xx - c * meta.y;
    float skip = step(hash(vec2(c * 1.7 + 11.0, row * 2.3)), 0.15);
    float onLine = span(xx, 0.0, meta.z, fp.x);
    float body = span(fy, -band, band, fp.y) * onLine;
    float bar = span(fx, 0.16 * meta.y, meta.y, fp.x) * body * (1.0 - skip);
    bar = mix(bar, 0.7 * body, smoothstep(0.25, 0.7, fp.x / meta.y));
    if (morph > 0.01 && skip < 0.5) {
      float g = letterCov(fx - 0.16 * meta.y, fy, vec2(c * 7.3, row * 3.9), meta.y / UNIT) * onLine * 1.3;
      bar = mix(bar, g, morph);
    }

    // the agent types the rest of its line as a tracked insertion
    vec2 a = u_agent.xy;
    vec3 ar = rowAt(a.y);
    float pending = 0.0;
    float inserted = 0.0;
    if (ar.z > 0.5 && abs(ar.x - row) < 0.5) {
      float aAge = u_t - u_agent.z;
      float words = floor(meta.z / meta.y);
      float c0 = clamp(floor((a.x - meta.x) / meta.y), 0.0, max(words - 2.0, 0.0));
      float typedW = min(floor(max(aAge, 0.0) * 2.4), words - c0);
      float from = meta.x + (c0 + 0.1) * meta.y;
      float to = meta.x + (c0 + typedW) * meta.y + 0.06;
      float done = (words - c0) / 2.4;
      float on = smoothstep(0.0, 0.3, aAge) * (1.0 - smoothstep(done + 1.6, done + 2.4, aAge));
      float cw = fp.x * u_px * 0.8;
      pending = on * step(from, p.x) * (1.0 - span(p.x, from - 1.0, to, fp.x));
      inserted = on * span(p.x, from, to, fp.x);
      accMark = on * (
        span(p.x, from, to, fp.x) * span(fy, band + 0.07, band + 0.07 + hair * 1.3, fp.y) * 0.8
        + span(p.x, to - cw, to + cw, fp.x) * span(fy, -0.34, 0.42, fp.y) * 0.95
        + rect(vec2(p.x, fy), vec4(to - cw, -0.5, to + 0.5, -0.34), fp) * 0.9
      );
    }

    float typed = span(xx, -1.0, laid * (meta.z + 1.0) - 0.5, fp.x);
    float shown = max(typed, lens) * (1.0 - pending);
    float breathe = 0.74 + 0.26 * sin(u_t * 0.5 + hash(vec2(c, row * 1.3)) * 6.283);
    ink = bar * shown * breathe * (1.0 - inserted);
    accInk = bar * inserted * 0.55;
    wire = frame(vec2(xx, fy), vec4(0.0, -band, meta.z, band), fp) * (1.0 - max(typed, lens));

    // the viewer's own caret and a whole-word selection on the pointer's row
    vec2 m = u_pointer;
    vec3 mr = rowAt(m.y);
    if (mr.z > 0.5 && abs(mr.x - row) < 0.5 && meta.z > 0.0) {
      float xxm = m.x - meta.x;
      float nWords = 1.0 + floor(hash(vec2(row * 1.9, floor(u_t * 0.45))) * 4.0);
      float maxW = max(floor(meta.z / meta.y) - 1.0, 0.0);
      float endW = clamp(floor(xxm / meta.y), 0.0, maxW);
      float startW = max(endW - nWords + 1.0, 0.0);
      sel = step(0.0, xxm) * span(fy, -0.4, 0.4, fp.y)
        * span(p.x, meta.x + (startW + 0.1) * meta.y, meta.x + (endW + 1.0) * meta.y + 0.03, fp.x);
      float cw = fp.x * u_px * 0.75;
      caret = span(p.x, m.x - cw, m.x + cw, fp.x) * span(fy, -0.44, 0.44, fp.y)
        * step(fract(u_t * 1.1), 0.55);
    }
  } else if (kind > 2.5 && kind < 3.5) {
    // a table: ruled grid, header row, cells typeset left to right
    float pr = mod(row, PARA);
    float tw = PW - 2.0 * ML;
    float u = (p.x - ML) / tw;
    float cs = u < 0.34 ? 0.0 : (u < 0.56 ? 0.34 : (u < 0.78 ? 0.56 : 0.78));
    float ce = u < 0.34 ? 0.34 : (u < 0.56 ? 0.56 : (u < 0.78 ? 0.78 : 1.0));
    float ci = u < 0.34 ? 0.0 : (u < 0.56 ? 1.0 : (u < 0.78 ? 2.0 : 3.0));
    float inside = span(p.x, ML, ML + tw, fp.x);
    float hx = fp.x * u_px * 0.5;
    grid = (span(fy, -0.5, -0.5 + hair, fp.y) + (pr > 5.5 ? span(fy, 0.5 - hair, 0.5, fp.y) : 0.0)) * inside
      + span(p.x, ML + cs * tw - hx, ML + cs * tw + hx, fp.x)
      + span(p.x, ML + ce * tw - hx, ML + ce * tw + hx, fp.x);
    shade = pr < 0.5 ? inside : 0.0;
    float x0 = ML + cs * tw + 0.35;
    float clen = ((ce - cs) * tw - 0.7) * mix(0.35, 0.85, hash(vec2(row * 3.1, ci)));
    float tband = pr < 0.5 ? 0.17 : 0.13;
    vec4 cell = vec4(x0, -tband, x0 + clen, tband);
    float shown = max(clamp(laid * 4.0 - ci, 0.0, 1.0), lens);
    ink = rect(vec2(p.x, fy), cell, fp) * shown;
    wire = frame(vec2(p.x, fy), cell, fp) * (1.0 - shown);
    weight = pr < 0.5 ? 1.4 : 1.0;
  } else if (kind > 3.5) {
    // a figure: framed bar chart whose bars rise as it crosses the front
    float pr = mod(row, PARA);
    float block = floor(row / PARA);
    vec2 fq = vec2(p.x, pr + r.y);
    vec4 fr = vec4(ML + 0.6, 0.3, PW - ML - 0.6, 5.7);
    float base = 5.2;
    grid = frame(fq, fr, fp)
      + span(fq.y, base, base + hair, fp.y) * span(p.x, fr.x + 0.5, fr.z - 0.5, fp.x);
    float bw = (fr.z - fr.x - 1.0) / 7.0;
    float bi = floor((p.x - fr.x - 0.5) / bw);
    if (bi >= 0.0 && bi < 7.0) {
      float h = mix(1.0, 4.3, hash(vec2(block, bi * 1.7))) + 0.16 * sin(u_t * 0.55 + bi * 1.9);
      float bx = fr.x + 0.5 + bi * bw;
      vec4 b = vec4(bx + bw * 0.2, base - h, bx + bw * 0.8, base);
      float bar = rect(fq, b, fp);
      float shown = max(clamp((p.y - u_frontY) * 4.0, 0.0, 1.0), lens);
      float hot = step(abs(bi - floor(hash(vec2(block, 8.8)) * 7.0)), 0.5);
      ink = bar * shown * (1.0 - hot);
      accInk = bar * shown * hot * 0.4;
      wire = frame(fq, b, fp) * (1.0 - shown);
      weight = 1.3;
    }
  }

  vec3 inkC = mix(vec3(0.1), vec3(0.93), u_dark);
  vec3 accC = mix(vec3(0.016, 0.471, 0.341), vec3(0.204, 0.827, 0.6), u_dark);
  vec3 col = mix(desk, mix(vec3(1.0), vec3(0.058), u_dark), paper);
  col *= 1.0 - shadow * 0.05 * (1.0 - u_dark);
  col = mix(col, inkC, (shade * 0.035 + chrome * 0.1 + grid * 0.13) * paper + edge * 0.16);
  col = mix(col, accC, sel * 0.22);
  col = mix(col, inkC, wire * 0.15);
  col = mix(col, inkC, min(ink * weight * (0.15 + 0.2 * lens + 0.22 * wave), 0.45));
  col = mix(col, accC, min(accInk, 0.6));
  col = mix(col, accC, min(accMark, 1.0));
  col = mix(col, inkC, caret * 0.75);

  // the layout front, a fine accent line across the sheet
  float fd = abs(gl_FragCoord.y - u_view.w * u_res.y) / u_px;
  col = mix(col, accC, paper * (clamp(1.0 - fd, 0.0, 1.0) * 0.32 + exp(-fd * 0.18) * 0.05));

  col = mix(bg, col, vis);
  col += (hash(gl_FragCoord.xy * 0.7 + fract(u_t) * 17.0) - 0.5) * 0.014 * vis;
  gl_FragColor = vec4(col, 1.0);
}
`;

// word bank as a 4x64 data texture: 0 = space, 1..26 = a..z
const WORDS =
  "the and for with text docx page font word edit open save type line list grid cell form note plan tabs news all was are has not but one two who out use own see now how its may also into over such each than them then when what more most some time very well work year from have this that been were said done live free real fast".split(
    " "
  );

const ease = (x: number) => 1 - (1 - x) ** 3;

// paints the document field behind the hero. it compiles off the main thread
// where the browser allows, pauses off-screen and in hidden tabs, renders one
// still frame under reduced motion, and lowers its resolution if frames run slow
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
    // without a field the hero is plain paper, never a dead canvas
    if (!gl) {
      canvas.style.visibility = "hidden";
      return;
    }
    const parallel = gl.getExtension("KHR_parallel_shader_compile");

    const program = gl.createProgram()!;
    for (const [type, src] of [
      [gl.VERTEX_SHADER, VERT],
      [gl.FRAGMENT_SHADER, FRAG],
    ] as const) {
      const shader = gl.createShader(type)!;
      gl.shaderSource(shader, src);
      gl.compileShader(shader);
      gl.attachShader(program, shader);
    }
    gl.linkProgram(program);

    let state: "compiling" | "ready" | "failed" = "compiling";
    let atlasReady = 0;
    const atlasImg = new Image();
    const u: Record<string, WebGLUniformLocation | null> = {};

    function setup(gl: WebGLRenderingContext) {
      if (!gl.getProgramParameter(program, gl.LINK_STATUS)) {
        const logs = gl.getAttachedShaders(program)?.map((s) => gl.getShaderInfoLog(s)).filter(Boolean);
        console.warn("hero field: shader failed to link", gl.getProgramInfoLog(program), ...(logs ?? []));
        return false;
      }
      gl.useProgram(program);
      for (const name of [
        "res", "t", "flow", "mouse", "pointer", "click", "view", "frontY", "clear",
        "agent", "intro", "dark", "px", "ready", "atlas", "words",
      ]) {
        u[name] = gl.getUniformLocation(program, `u_${name}`);
      }

      gl.bindBuffer(gl.ARRAY_BUFFER, gl.createBuffer());
      gl.bufferData(gl.ARRAY_BUFFER, new Float32Array([-1, -1, 3, -1, -1, 3]), gl.STATIC_DRAW);
      const loc = gl.getAttribLocation(program, "a_pos");
      gl.enableVertexAttribArray(loc);
      gl.vertexAttribPointer(loc, 2, gl.FLOAT, false, 0, 0);

      const wordData = new Uint8Array(4 * 64);
      for (let w = 0; w < 64; w++) {
        const word = WORDS[w % WORDS.length];
        for (let i = 0; i < 4; i++) wordData[w * 4 + i] = i < word.length ? word.charCodeAt(i) - 96 : 0;
      }
      gl.activeTexture(gl.TEXTURE1);
      gl.bindTexture(gl.TEXTURE_2D, gl.createTexture());
      gl.pixelStorei(gl.UNPACK_ALIGNMENT, 1);
      gl.texImage2D(gl.TEXTURE_2D, 0, gl.LUMINANCE, 4, 64, 0, gl.LUMINANCE, gl.UNSIGNED_BYTE, wordData);
      gl.texParameteri(gl.TEXTURE_2D, gl.TEXTURE_MIN_FILTER, gl.NEAREST);
      gl.texParameteri(gl.TEXTURE_2D, gl.TEXTURE_MAG_FILTER, gl.NEAREST);
      gl.texParameteri(gl.TEXTURE_2D, gl.TEXTURE_WRAP_S, gl.CLAMP_TO_EDGE);
      gl.texParameteri(gl.TEXTURE_2D, gl.TEXTURE_WRAP_T, gl.CLAMP_TO_EDGE);
      gl.uniform1i(u.words, 1);

      // the baked font atlas ships as a static asset; until it arrives the
      // lens keeps showing plain word bars
      const atlasTex = gl.createTexture();
      gl.activeTexture(gl.TEXTURE0);
      gl.bindTexture(gl.TEXTURE_2D, atlasTex);
      gl.texImage2D(gl.TEXTURE_2D, 0, gl.RGBA, 1, 1, 0, gl.RGBA, gl.UNSIGNED_BYTE, new Uint8Array(4));
      gl.texParameteri(gl.TEXTURE_2D, gl.TEXTURE_MIN_FILTER, gl.LINEAR);
      gl.texParameteri(gl.TEXTURE_2D, gl.TEXTURE_MAG_FILTER, gl.LINEAR);
      gl.texParameteri(gl.TEXTURE_2D, gl.TEXTURE_WRAP_S, gl.CLAMP_TO_EDGE);
      gl.texParameteri(gl.TEXTURE_2D, gl.TEXTURE_WRAP_T, gl.CLAMP_TO_EDGE);
      gl.uniform1i(u.atlas, 0);
      atlasImg.onload = () => {
        if (state !== "ready") return;
        gl.activeTexture(gl.TEXTURE0);
        gl.bindTexture(gl.TEXTURE_2D, atlasTex);
        gl.texImage2D(gl.TEXTURE_2D, 0, gl.RGBA, gl.RGBA, gl.UNSIGNED_BYTE, atlasImg);
        gl.texParameteri(gl.TEXTURE_2D, gl.TEXTURE_MIN_FILTER, gl.LINEAR_MIPMAP_LINEAR);
        gl.generateMipmap(gl.TEXTURE_2D);
        atlasReady = 1;
        if (reduce) redraw();
      };
      atlasImg.src = "/glyph-atlas.png";
      return true;
    }

    const reduce = window.matchMedia("(prefers-reduced-motion: reduce)").matches;
    const scheme = window.matchMedia("(prefers-color-scheme: dark)");
    let dark = scheme.matches ? 1 : 0;

    // camera fitted to the hero: horizon just under the canvas top, the sheet
    // about as wide as the frame just above the copy block, which stays clean
    const clear = [0, 0, 0, 0];
    const view = { horizon: 0.96, lateral: 4.4, depth: 4.1, base: 0.45 };
    const clearEl = canvas.parentElement?.querySelector("[data-hero-clear]");
    function measure() {
      const c = canvas!.getBoundingClientRect();
      if (!c.width || !c.height) return;
      if (clearEl) {
        const r = clearEl.getBoundingClientRect();
        clear[0] = (r.left - c.left) / c.width;
        clear[1] = 1 - (r.bottom - c.top) / c.height;
        clear[2] = (r.right - c.left) / c.width;
        clear[3] = 1 - (r.top - c.top) / c.height;
      }
      const base = clearEl ? clear[3] : 0.45;
      const horizon = 1 - 26 / c.height;
      const k0 = Math.max(horizon - base - 0.08, 0.12);
      const sheet = c.width * (c.width < 700 ? 1.45 : 0.86);
      const line = Math.min(Math.max(c.width / 46, 15), 30);
      view.base = base;
      view.horizon = horizon;
      view.lateral = (PAGE_WIDTH * k0 * c.height) / sheet;
      view.depth = (k0 * k0 * c.height) / line;
    }
    const restFront = () => view.base + (view.horizon - view.base) * 0.6;

    // the shader's ground-plane projection, for points that are the same for every pixel
    function plane(x: number, y: number, flow: number, lean: number) {
      const k = Math.max(view.horizon - y, 0.004);
      const aspect = canvas!.clientWidth / Math.max(canvas!.clientHeight, 1);
      return [(x - 0.5 + lean) * aspect * (view.lateral / k) + PAGE_WIDTH / 2, -view.depth / k - flow];
    }

    let quality = 1;
    let slow = 0;
    let counted = 0;
    function adapt(dt: number) {
      if (dt <= 0 || dt > 0.25 || ++counted < 40) return;
      slow += ((dt > 0.026 ? 1 : 0) - slow) * 0.05;
      if (slow > 0.6 && quality > 0.5) {
        quality *= 0.75;
        slow = 0;
        counted = 0;
      }
    }

    let dpr = 1;
    function resize(gl: WebGLRenderingContext) {
      dpr = Math.min(window.devicePixelRatio || 1, 1.5) * quality;
      const w = Math.round(canvas!.clientWidth * dpr);
      const h = Math.round(canvas!.clientHeight * dpr);
      if (canvas!.width !== w || canvas!.height !== h) {
        canvas!.width = w;
        canvas!.height = h;
        gl.viewport(0, 0, w, h);
      }
    }

    const mouse = { x: 0.68, y: 0.75 };
    const target = { x: 0.68, y: 0.75 };
    const click = { x: 0.5, y: 0.5, t: -100 };
    const agent = [0, 0, -100];
    let nextAgent = 2.4;
    let lastMove = -1e9;
    let clock = 0;

    function toLocal(e: PointerEvent) {
      const rect = canvas!.getBoundingClientRect();
      return { x: (e.clientX - rect.left) / rect.width, y: 1 - (e.clientY - rect.top) / rect.height };
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
      click.t = clock;
    }

    let raf = 0;
    let running = false;
    let inView = true;
    let last = performance.now();

    function frame(now: number) {
      if (state === "compiling") {
        if (parallel && !gl!.getProgramParameter(program, parallel.COMPLETION_STATUS_KHR)) {
          raf = requestAnimationFrame(frame);
          return;
        }
        state = setup(gl!) ? "ready" : "failed";
        if (state === "failed") {
          canvas!.style.visibility = "hidden";
          return;
        }
        last = now;
      }
      const dt = Math.min((now - last) / 1000, 0.1);
      last = now;
      if (!reduce) {
        clock += dt;
        adapt(dt);
      }
      resize(gl!);

      const t = reduce ? STILL : clock;
      const scroll = window.scrollY / Math.max(canvas!.clientHeight, 1);
      const flow = t * FLOW + scroll * SCROLL_ROWS;
      const rest = restFront();
      const front = reduce ? rest : view.base - 0.04 + (rest - view.base + 0.04) * ease(Math.min(t / 2.4, 1));
      const lo = view.base + 0.12;

      if (reduce) {
        [agent[0], agent[1]] = plane(0.36, lo + (rest - lo) * 0.4, (t - 1.3) * FLOW + scroll * SCROLL_ROWS, 0);
        agent[2] = t - 1.3;
        mouse.x = 0.7;
        mouse.y = (rest + view.horizon) * 0.5;
      } else {
        if (t > nextAgent) {
          [agent[0], agent[1]] = plane(
            0.22 + Math.random() * 0.5,
            lo + Math.random() * Math.max(rest - 0.03 - lo, 0.02),
            flow,
            0
          );
          agent[2] = t;
          nextAgent = t + AGENT_EVERY;
        }
        // the ghost reader roams the visible pages while the pointer is idle
        if (performance.now() - lastMove > 3500) {
          const mid = (lo + view.horizon - 0.1) * 0.5;
          const amp = Math.max(view.horizon - 0.1 - lo, 0) * 0.5;
          target.x = 0.5 + 0.36 * Math.sin(t * 0.19) * Math.cos(t * 0.08);
          target.y = mid + amp * Math.sin(t * 0.12 + 1.7);
        }
        mouse.x += (target.x - mouse.x) * 0.055;
        mouse.y += (target.y - mouse.y) * 0.055;
      }

      gl!.uniform2f(u.res, canvas!.width, canvas!.height);
      gl!.uniform1f(u.t, t);
      gl!.uniform1f(u.flow, flow);
      const lean = (0.5 - mouse.x) * 0.03;
      const pointer = plane(mouse.x, mouse.y, flow, lean);
      const hit = plane(click.x, click.y, flow, lean);
      gl!.uniform2f(u.mouse, mouse.x, mouse.y);
      gl!.uniform2f(u.pointer, pointer[0], pointer[1]);
      gl!.uniform3f(u.click, hit[0], hit[1], click.t);
      gl!.uniform4f(u.view, view.horizon, view.lateral, view.depth, front);
      gl!.uniform1f(u.frontY, plane(0.5, front, flow, 0)[1]);
      gl!.uniform4f(u.clear, clear[0], clear[1], clear[2], clear[3]);
      gl!.uniform3f(u.agent, agent[0], agent[1], agent[2]);
      gl!.uniform1f(u.intro, reduce ? 1 : ease(Math.min(t / 1.2, 1)));
      gl!.uniform1f(u.dark, dark);
      gl!.uniform1f(u.px, dpr);
      gl!.uniform1f(u.ready, atlasReady);
      gl!.drawArrays(gl!.TRIANGLES, 0, 3);
      if (!reduce && running) raf = requestAnimationFrame(frame);
    }

    function redraw() {
      cancelAnimationFrame(raf);
      raf = requestAnimationFrame(frame);
    }

    function sync() {
      const go = inView && !document.hidden && state !== "failed" && !reduce;
      if (go && !running) {
        running = true;
        last = performance.now();
        cancelAnimationFrame(raf);
        raf = requestAnimationFrame(frame);
      } else if (!go && running) {
        running = false;
        if (state === "ready") cancelAnimationFrame(raf);
      }
    }

    const sized = new ResizeObserver(() => {
      measure();
      if (reduce) redraw();
    });
    measure();
    sized.observe(canvas);
    if (clearEl) sized.observe(clearEl);

    function onScheme(e: MediaQueryListEvent) {
      dark = e.matches ? 1 : 0;
      if (reduce) redraw();
    }
    scheme.addEventListener("change", onScheme);

    const observer = new IntersectionObserver(([entry]) => {
      inView = entry.isIntersecting;
      sync();
    });
    observer.observe(canvas);
    document.addEventListener("visibilitychange", sync);

    function onLost() {
      state = "failed";
      running = false;
      cancelAnimationFrame(raf);
      canvas!.style.visibility = "hidden";
    }
    canvas.addEventListener("webglcontextlost", onLost);
    window.addEventListener("pointermove", onMove, { passive: true });
    window.addEventListener("pointerdown", onDown, { passive: true });

    if (reduce) raf = requestAnimationFrame(frame);
    else sync();

    return () => {
      cancelAnimationFrame(raf);
      running = false;
      state = "failed";
      atlasImg.onload = null;
      observer.disconnect();
      sized.disconnect();
      scheme.removeEventListener("change", onScheme);
      document.removeEventListener("visibilitychange", sync);
      canvas.removeEventListener("webglcontextlost", onLost);
      window.removeEventListener("pointermove", onMove);
      window.removeEventListener("pointerdown", onDown);
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
