import { chromium } from "playwright";
import { resolve, extname } from "node:path";
import { mkdir } from "node:fs/promises";

const root = resolve(import.meta.dir, "../../..");
const dist = resolve(root, "apps/desktop/dist");
const output =
  process.env.DESKTOP_SCREENSHOTS ?? "/tmp/betteroffice-desktop-screenshots";
await mkdir(output, { recursive: true });
const types: Record<string, string> = {
  ".html": "text/html",
  ".js": "text/javascript",
  ".css": "text/css",
  ".wasm": "application/wasm",
  ".ttf": "font/ttf",
};
const server = Bun.serve({
  hostname: "127.0.0.1",
  port: 0,
  async fetch(request) {
    const url = new URL(request.url);
    const path = resolve(
      dist,
      `.${url.pathname === "/" ? "/index.html" : url.pathname}`
    );
    if (!path.startsWith(`${dist}/`))
      return new Response(null, { status: 404 });
    const file = Bun.file(path);
    if (!(await file.exists())) return new Response(null, { status: 404 });
    return new Response(file, {
      headers: {
        "Content-Type": types[extname(path)] ?? "application/octet-stream",
      },
    });
  },
});
const browser = await chromium.launch({ headless: true });
try {
  for (const format of ["docx", "xlsx", "pptx"] as const) {
    const page = await browser.newPage({
      viewport: { width: 1180, height: 860 },
    });
    const errors: string[] = [];
    page.on("pageerror", (error) => errors.push(error.message));
    page.on("response", (response) => {
      if (response.status() >= 400)
        errors.push(`${response.status()} ${response.url()}`);
    });
    await page.goto(`http://127.0.0.1:${server.port}/?format=${format}`);
    await page
      .getByRole("heading", {
        name: `Open a ${
          { docx: "document", xlsx: "spreadsheet", pptx: "presentation" }[
            format
          ]
        }`,
      })
      .waitFor();
    if (format === "docx")
      await page.screenshot({ path: `${output}/start.png` });
    const input = page.locator("input[type=file]").last();
    const fixture =
      format === "xlsx" ? "showcase.xlsx" : `betteroffice-demo.${format}`;
    await input.setInputFiles(resolve(root, "apps/demo/public", fixture));
    await page
      .locator(".editor-stage canvas")
      .first()
      .waitFor({ timeout: 120_000 });
    await page.waitForFunction(() =>
      [...document.querySelectorAll("canvas")].some(
        (canvas) => canvas.width > 100 && canvas.height > 100
      )
    );
    await page.waitForTimeout(1500);
    const alert = (await page.locator(".notice").count())
      ? await page.locator(".notice").textContent()
      : null;
    if (alert) throw new Error(`${format}: ${alert}`);
    await page.screenshot({ path: `${output}/${format}.png` });
    const download = page.waitForEvent("download");
    await page
      .getByRole("button", { name: "Save", exact: true })
      .first()
      .click();
    const saved = await download;
    await saved.saveAs(`${output}/saved.${format}`);
    await input.setInputFiles({
      name: `broken.${format}`,
      mimeType: "application/octet-stream",
      buffer: Buffer.from("not a zip"),
    });
    await page.locator(".notice").waitFor();
    if ((await page.locator(".editor-stage canvas").count()) === 0)
      throw new Error(`${format}: failed open discarded the current editor`);
    if (errors.length) throw new Error(`${format}: ${errors.join("\n")}`);
    console.log(
      `${format}: start, open, toolbar, save, failed-open retention passed`
    );
    await page.close();
  }
} finally {
  await browser.close();
  server.stop(true);
}
