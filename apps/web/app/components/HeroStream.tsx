"use client";

import { useEffect, useRef } from "react";

const FPS = 30;
const DRIFT = 5;
const FADE = 1.2;

const VERT = `
attribute vec2 a_pos;
void main() {
  gl_Position = vec4(a_pos, 0.0, 1.0);
}
`;

// the page stream seen from above, confined to the band between the nav and the hero copy:
// staggered pages in hairline outline, their text still line boxes, drifting toward the reader.
// everything is in CSS pixels so the lines stay one pixel wide.
const FRAG = `
#ifdef GL_FRAGMENT_PRECISION_HIGH
precision highp float;
#else
precision mediump float;
#endif

uniform vec2 u_res;
uniform float u_px;
uniform float u_x0;
uniform float u_off;
uniform float u_dark;
uniform float u_fade;
uniform vec4 u_band;

const vec2 PAGE = vec2(168.0, 238.0);
const vec2 GAP = vec2(40.0, 46.0);
const float PITCH = 9.0;

float hash(vec2 p) {
  return fract(sin(dot(p, vec2(127.1, 311.7))) * 43758.5453123);
}

float span(float x, float a, float b, float fp) {
  return clamp((min(x + 0.5 * fp, b) - max(x - 0.5 * fp, a)) / fp, 0.0, 1.0);
}

float box(vec2 p, vec4 r, float fp) {
  return span(p.x, r.x, r.z, fp) * span(p.y, r.y, r.w, fp);
}

float frame(vec2 p, vec4 r, float fp) {
  return box(p, r + vec4(-0.5, -0.5, 0.5, 0.5), fp) - box(p, r + vec4(0.5, 0.5, -0.5, -0.5), fp);
}

void main() {
  float fp = 1.0 / u_px;
  vec2 s = vec2(gl_FragCoord.x, u_res.y - gl_FragCoord.y) * fp;
  float band = smoothstep(u_band.x, u_band.x + 48.0, s.x) * smoothstep(u_band.z, u_band.z - 48.0, s.x)
    * smoothstep(u_band.y, u_band.y + 28.0, s.y) * smoothstep(u_band.w, u_band.w - 44.0, s.y);
  if (band < 0.002) {
    gl_FragColor = vec4(0.0);
    return;
  }
  vec2 cell = PAGE + GAP;
  float fx = s.x - u_x0;
  float cx = floor(fx / cell.x);
  float fy = s.y - u_off - mod(cx, 2.0) * 0.5 * cell.y;
  float cy = floor(fy / cell.y);
  vec2 l = vec2(fx - cx * cell.x, fy - cy * cell.y);
  vec2 id = vec2(cx, cy);
  float h = hash(id);

  float outline = frame(l, vec4(0.0, 0.0, PAGE), fp);
  float sheet = box(l, vec4(0.0, 0.0, PAGE), fp);
  float ink = box(l, vec4(16.0, 14.0, 52.0 + 30.0 * h, 16.0), fp);
  ink += box(l, vec4(PAGE.x * 0.5 - 5.0, PAGE.y - 15.0, PAGE.x * 0.5 + 5.0, PAGE.y - 13.0), fp);

  // body text: paragraphs of line boxes, a figure on some pages, a ruled table on others
  float ty = l.y - 30.0;
  if (ty > -2.0 && ty < PAGE.y - 58.0 && l.x > 8.0 && l.x < PAGE.x - 8.0) {
    float line = floor(ty / PITCH);
    float start = 0.0;
    for (int j = 0; j < 8; j++) {
      float fj = float(j);
      float len = 3.0 + floor(hash(id + vec2(fj * 7.1, 3.3)) * 4.0);
      if (line < start + len) {
        if (line >= start) {
          float k = line - start;
          float kind = hash(id + vec2(fj, 9.7));
          vec2 q = vec2(l.x, ty - start * PITCH);
          float bottom = len * PITCH - 6.0;
          if (fj == 1.0 && h > 0.55) {
            float fig = frame(q, vec4(16.0, 0.0, PAGE.x - 16.0, bottom), fp);
            float bars = 0.0;
            for (int b = 0; b < 5; b++) {
              float fb = float(b);
              float bh = (0.3 + 0.6 * hash(id + vec2(fb, 1.9))) * (bottom - 8.0);
              bars += box(q, vec4(28.0 + fb * 24.0, bottom - 4.0 - bh, 40.0 + fb * 24.0, bottom - 4.0), fp);
            }
            ink += fig + bars * 0.6;
          } else if (fj == 2.0 && kind < 0.35) {
            float rows = 0.0;
            for (int r = 0; r < 7; r++) {
              float fr = float(r);
              if (fr > len) break;
              rows += box(q, vec4(16.0, fr * PITCH - 0.5, PAGE.x - 16.0, fr * PITCH + 0.5), fp);
            }
            float cols = box(q, vec4(68.5, 0.0, 69.5, len * PITCH), fp) + box(q, vec4(110.5, 0.0, 111.5, len * PITCH), fp);
            ink += rows + cols;
          } else {
            float w = k == len - 1.0 ? 0.3 + 0.5 * hash(id + vec2(fj, k)) : 1.0;
            ink += box(q, vec4(16.0, k * PITCH, 16.0 + (PAGE.x - 32.0) * w, k * PITCH + 3.0), fp) * 0.7;
          }
        }
        break;
      }
      start += len + 1.0;
    }
  }

  float a = (outline * mix(0.045, 0.06, u_dark) + clamp(ink, 0.0, 1.0) * mix(0.03, 0.04, u_dark)) * u_fade * band;
  float lift = sheet * 0.018 * u_dark * u_fade * band;
  vec3 c = mix(vec3(0.043), vec3(0.95), u_dark);
  gl_FragColor = vec4(c * a + vec3(1.0) * lift, a + lift);
}
`;

// paints the stream in the hero's empty band. it waits for an idle moment to compile, draws at
// 30 fps only while the hero is on screen and the tab visible, lowers its resolution if frames
// run slow, renders one still frame under reduced motion, and hides itself without WebGL
export function HeroStream() {
  const ref = useRef<HTMLCanvasElement>(null);

  useEffect(() => {
    const canvas = ref.current;
    const hero = canvas?.parentElement;
    if (!canvas || !hero) return;
    const frameEl = hero.querySelector<HTMLElement>(".wrap");
    const clearEl = hero.querySelector<HTMLElement>("[data-hero-clear]");
    let gl: WebGLRenderingContext | null = null;
    let program: WebGLProgram | null = null;
    const u: Record<string, WebGLUniformLocation | null> = {};
    let state: "idle" | "compiling" | "ready" | "failed" = "idle";
    let scheduled = false;
    let raf = 0;
    let inView = false;
    let clock = 0;
    let last = 0;
    let drawn = 0;
    let born = -1;
    let px = 1;
    let quality = 1;
    let slow = 0;
    let counted = 0;
    const band = [0, 0, 0, 0];

    const reduce = window.matchMedia("(prefers-reduced-motion: reduce)");
    const scheme = window.matchMedia("(prefers-color-scheme: dark)");

    function fail() {
      state = "failed";
      cancelAnimationFrame(raf);
      canvas!.style.visibility = "hidden";
    }

    // rail to rail across, from the canvas top down to a margin above the copy
    function measure() {
      const c = canvas!.getBoundingClientRect();
      const f = (frameEl ?? canvas!).getBoundingClientRect();
      band[0] = f.left - c.left;
      band[2] = f.right - c.left;
      band[1] = 0;
      band[3] = clearEl ? clearEl.getBoundingClientRect().top - c.top - 20 : c.height * 0.4;
    }

    function init() {
      if (state !== "idle") return;
      gl = canvas!.getContext("webgl", {
        alpha: true,
        antialias: false,
        depth: false,
        stencil: false,
        powerPreference: "low-power",
      });
      if (!gl) {
        fail();
        return;
      }
      const parallel = gl.getExtension("KHR_parallel_shader_compile");
      program = gl.createProgram()!;
      for (const [type, text] of [
        [gl.VERTEX_SHADER, VERT],
        [gl.FRAGMENT_SHADER, FRAG],
      ] as const) {
        const shader = gl.createShader(type)!;
        gl.shaderSource(shader, text);
        gl.compileShader(shader);
        gl.attachShader(program, shader);
      }
      gl.linkProgram(program);
      state = "compiling";
      const poll = () => {
        if (state !== "compiling" || !gl || !program) return;
        if (parallel && !gl.getProgramParameter(program, parallel.COMPLETION_STATUS_KHR)) {
          setTimeout(poll, 50);
          return;
        }
        if (!gl.getProgramParameter(program, gl.LINK_STATUS)) {
          const logs = gl.getAttachedShaders(program)?.map((s) => gl!.getShaderInfoLog(s)).filter(Boolean);
          console.warn("hero stream: shader failed to link", gl.getProgramInfoLog(program), ...(logs ?? []));
          fail();
          return;
        }
        gl.useProgram(program);
        gl.bindBuffer(gl.ARRAY_BUFFER, gl.createBuffer());
        gl.bufferData(gl.ARRAY_BUFFER, new Float32Array([-1, -1, 3, -1, -1, 3]), gl.STATIC_DRAW);
        const loc = gl.getAttribLocation(program, "a_pos");
        gl.enableVertexAttribArray(loc);
        gl.vertexAttribPointer(loc, 2, gl.FLOAT, false, 0, 0);
        for (const name of ["res", "px", "x0", "off", "dark", "fade", "band"]) {
          u[name] = gl.getUniformLocation(program, `u_${name}`);
        }
        state = "ready";
        sync();
      };
      poll();
    }

    function resize() {
      if (!gl) return;
      px = Math.min(window.devicePixelRatio || 1, 1.5) * quality;
      const w = Math.round(canvas!.clientWidth * px);
      const h = Math.round(canvas!.clientHeight * px);
      if (canvas!.width !== w || canvas!.height !== h) {
        canvas!.width = w;
        canvas!.height = h;
        gl.viewport(0, 0, w, h);
      }
    }

    function draw(now: number) {
      if (state !== "ready" || !gl) return;
      if (born < 0) born = now;
      resize();
      const still = reduce.matches;
      gl.uniform2f(u.res, canvas!.width, canvas!.height);
      gl.uniform1f(u.px, px);
      gl.uniform1f(u.x0, (band[0] + band[2]) / 2 - 84);
      gl.uniform1f(u.off, still ? 0 : (clock * DRIFT) % 28400);
      gl.uniform1f(u.dark, scheme.matches ? 1 : 0);
      gl.uniform1f(u.fade, still ? 1 : Math.min((now - born) / 1000 / FADE, 1));
      gl.uniform4f(u.band, band[0], band[1], band[2], band[3]);
      gl.drawArrays(gl.TRIANGLES, 0, 3);
    }

    // steps the resolution down while frames run slow
    function adapt(gap: number) {
      if (gap <= 0 || gap > 250 || ++counted < 30) return;
      slow += ((gap > 26 ? 1 : 0) - slow) * 0.05;
      if (slow > 0.6 && quality > 0.5) {
        quality *= 0.75;
        slow = 0;
        counted = 0;
      }
    }

    let prev = 0;
    function tick(now: number) {
      raf = requestAnimationFrame(tick);
      if (prev) adapt(now - prev);
      prev = now;
      if (now - drawn < 1000 / FPS - 2) return;
      clock += Math.min((now - last) / 1000, 0.1);
      last = now;
      drawn = now;
      draw(now);
    }

    function sync() {
      cancelAnimationFrame(raf);
      raf = 0;
      if (!inView || document.hidden || state === "failed") return;
      if (state === "idle") {
        if (!scheduled) {
          scheduled = true;
          if ("requestIdleCallback" in window) requestIdleCallback(init, { timeout: 1500 });
          else setTimeout(init, 200);
        }
        return;
      }
      if (state !== "ready") return;
      if (reduce.matches) {
        raf = requestAnimationFrame(draw);
        return;
      }
      last = performance.now();
      prev = 0;
      raf = requestAnimationFrame(tick);
    }

    const io = new IntersectionObserver(([entry]) => {
      inView = entry.isIntersecting;
      sync();
    });
    io.observe(canvas);
    const sized = new ResizeObserver(() => {
      measure();
      if (reduce.matches) sync();
    });
    measure();
    sized.observe(canvas);
    if (clearEl) sized.observe(clearEl);
    document.addEventListener("visibilitychange", sync);
    reduce.addEventListener("change", sync);
    scheme.addEventListener("change", sync);
    canvas.addEventListener("webglcontextlost", fail);

    return () => {
      cancelAnimationFrame(raf);
      state = "failed";
      io.disconnect();
      sized.disconnect();
      document.removeEventListener("visibilitychange", sync);
      reduce.removeEventListener("change", sync);
      scheme.removeEventListener("change", sync);
      canvas.removeEventListener("webglcontextlost", fail);
      gl?.getExtension("WEBGL_lose_context")?.loseContext();
    };
  }, []);

  return (
    <canvas ref={ref} className="pointer-events-none absolute inset-0 h-full w-full" aria-hidden="true" />
  );
}
