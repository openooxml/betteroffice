import { test, expect, type FrameLocator, type Page, type TestInfo } from "playwright/test";
import JSZip from "jszip";

type Format = "docx" | "xlsx" | "pptx";

function urlPath(raw: string): string {
  try {
    return new URL(raw).pathname;
  } catch {
    return raw;
  }
}

interface SocketRecord {
  url: string;
  sent: number;
  received: number;
  closed: boolean;
}

interface Diagnostics {
  lines: string[];
  sockets: SocketRecord[];
}

/** Records browser diagnostics for the host page and every iframe. */
async function instrument(page: Page): Promise<Diagnostics> {
  const lines: string[] = [];
  const sockets: SocketRecord[] = [];
  page.on("console", (message) => {
    lines.push(`[console:${message.type()}] ${urlPath(message.location().url)} ${message.text()}`);
  });
  page.on("pageerror", (error) => {
    lines.push(`[pageerror] ${error.message}`);
  });
  page.on("requestfailed", (request) => {
    lines.push(`[requestfailed] ${request.method()} ${urlPath(request.url())} ${request.failure()?.errorText ?? ""}`);
  });
  page.on("response", (response) => {
    const path = urlPath(response.url());
    if (response.status() >= 400 || path.startsWith("/api/") || path.startsWith("/app/")) {
      lines.push(`[http ${response.status()}] ${path}`);
    }
  });
  page.on("websocket", (socket) => {
    const record: SocketRecord = { url: urlPath(socket.url()), sent: 0, received: 0, closed: false };
    sockets.push(record);
    socket.on("framesent", () => {
      record.sent += 1;
    });
    socket.on("framereceived", () => {
      record.received += 1;
    });
    socket.on("close", () => {
      record.closed = true;
    });
  });
  await page.addInitScript(() => {
    const holder = window as unknown as { __wopiCsp: string[] };
    holder.__wopiCsp = [];
    document.addEventListener("securitypolicyviolation", (event) => {
      holder.__wopiCsp.push(
        `${event.violatedDirective} blocked=${event.blockedURI} source=${event.sourceFile ?? ""} sample=${event.sample ?? ""}`,
      );
    });
  });
  return { lines, sockets };
}

async function dumpDiagnostics(page: Page, diagnostics: Diagnostics, info: TestInfo): Promise<void> {
  const all = [...diagnostics.lines];
  for (const socket of diagnostics.sockets) {
    all.push(`[websocket] ${socket.url} sent=${socket.sent} received=${socket.received} closed=${socket.closed}`);
  }
  const hostMessages = await page.evaluate(() => JSON.stringify((window as unknown as { __wopiMessages?: unknown[] }).__wopiMessages ?? [])).catch(() => "(unavailable)");
  all.push(`[host-messages] ${hostMessages}`);
  for (const frame of page.frames()) {
    const violations = await frame.evaluate(() => (window as unknown as { __wopiCsp?: string[] }).__wopiCsp ?? []).catch(() => []);
    for (const violation of violations) all.push(`[csp] ${urlPath(frame.url())} ${violation}`);
    if (frame === page.mainFrame()) continue;
    const state = await frame.evaluate(() => {
      const holder = window as unknown as { __wopiDebug?: Record<string, unknown> };
      return {
        text: (document.body?.innerText ?? "").slice(0, 400),
        rootChildren: document.getElementById("root")?.childElementCount ?? -1,
        debug: holder.__wopiDebug ?? null,
      };
    }).catch((error: Error) => ({ text: `(evaluate failed: ${error.message})`, rootChildren: -1, debug: null }));
    all.push(`[iframe] ${urlPath(frame.url())} rootChildren=${state.rootChildren} debug=${JSON.stringify(state.debug)} text=${JSON.stringify(state.text)}`);
  }
  const iframeUrls = page.frames().map((frame) => frame.url());
  all.push(`[frames] ${iframeUrls.join(" | ")}`);
  const text = all.length > 0 ? all.join("\n") : "(no diagnostics captured)";
  console.log(`\nWOPI e2e diagnostics:\n${text}\n`);
  await info.attach("wopi-diagnostics", { body: text, contentType: "text/plain" });
}

const FORMATS: Format[] = ["docx", "xlsx", "pptx"];

async function loadedCount(page: Page): Promise<number> {
  return page.evaluate(() => {
    const messages = (window as unknown as { __wopiMessages: Array<{ MessageId: string; Values?: { Status?: string } }> }).__wopiMessages;
    return messages.filter((message) => message.MessageId === "App_LoadingStatus" && message.Values?.Status === "Document_Loaded").length;
  });
}

async function edit(frame: FrameLocator, page: Page, format: Format, marker: string, row = 0): Promise<void> {
  if (format === "docx") {
    await frame.locator(".canvas-page canvas").first().click({ position: { x: 150, y: 150 } });
    const input = frame.locator("[data-testid=yrs-input], .paged-editor__input");
    await expect(input).toBeEditable();
    await input.press("ControlOrMeta+Home");
    await page.keyboard.insertText(marker);
    await expect(frame.locator(".canvas-page-mirror").first()).toContainText(marker);
    return;
  }
  if (format === "xlsx") {
    const editor = frame.getByTestId("xlsx-scroll");
    for (let i = 0; i < row; i += 1) await editor.press("ArrowDown");
    await editor.press("F2");
    const input = frame.getByTestId("xlsx-cell-editor");
    await input.press("ControlOrMeta+A");
    await input.pressSequentially(marker);
    await input.press("Enter");
    await expect(frame.getByRole("gridcell").filter({ hasText: marker })).toHaveCount(1);
    return;
  }
  await frame.getByTestId("pptx-tool-text-box").click();
  const bounds = await frame.getByTestId("pptx-slide-canvas").boundingBox();
  expect(bounds).not.toBeNull();
  await page.mouse.move(bounds!.x + bounds!.width * 0.1, bounds!.y + bounds!.height * 0.65);
  await page.mouse.down();
  await page.mouse.move(bounds!.x + bounds!.width * 0.6, bounds!.y + bounds!.height * 0.85, { steps: 5 });
  await page.mouse.up();
  await page.keyboard.type(marker);
}

async function containsText(frame: FrameLocator, format: Format, marker: string): Promise<void> {
  if (format === "docx") {
    await expect(frame.locator(".canvas-page-mirror").first()).toContainText(marker);
  } else if (format === "xlsx") {
    await expect(frame.getByRole("gridcell").filter({ hasText: marker })).toHaveCount(1);
  }
}

async function sendSave(page: Page): Promise<void> {
  await page.evaluate(() => {
    (window as unknown as { __sendToFrames: (id: string) => void }).__sendToFrames("Action_Save");
  });
}

async function savedText(page: Page, format: Format): Promise<string> {
  const base64 = await (await page.request.get(`/debug/file?format=${format}`)).text();
  const zip = await JSZip.loadAsync(Buffer.from(base64, "base64"));
  const parts = Object.values(zip.files).filter((part) => !part.dir && part.name.endsWith(".xml"));
  const xml = await Promise.all(parts.map((part) => part.async("string")));
  return xml.join("\n");
}

for (const format of FORMATS) {
  test(`${format}: two users co-edit in iframes and Action_Save puts both edits`, async ({ page }, info) => {
    const diagnostics = await instrument(page);
    try {
      const markerA = `WopiAlice${format}1`;
      const markerB = `WopiBob${format}2`;
      await page.goto(`/host-page?format=${format}`);

      await expect.poll(() => loadedCount(page), { timeout: 120_000 }).toBe(2);

      const frameA = page.frameLocator("iframe[title='Alice']");
      const frameB = page.frameLocator("iframe[title='Bob']");

      if (format === "pptx") {
        await edit(frameA, page, format, markerA);
        await sendSave(page);
        await expect.poll(async () => savedText(page, format), { timeout: 60_000 }).toContain(markerA);
        await edit(frameB, page, format, markerB, 1);
        await sendSave(page);
        await expect.poll(async () => savedText(page, format), { timeout: 60_000 }).toContain(markerB);
        await expect.poll(async () => savedText(page, format), { timeout: 60_000 }).toContain(markerA);
        return;
      }

      await edit(frameA, page, format, markerA);
      await containsText(frameB, format, markerA);

      await edit(frameB, page, format, markerB, 1);
      await containsText(frameA, format, markerB);

      await sendSave(page);

      await expect.poll(async () => savedText(page, format), { timeout: 60_000 }).toContain(markerA);
      await expect.poll(async () => savedText(page, format), { timeout: 60_000 }).toContain(markerB);
    } catch (error) {
      await dumpDiagnostics(page, diagnostics, info);
      throw error;
    }
  });
}
