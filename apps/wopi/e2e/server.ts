import { readFileSync } from "node:fs";
import { resolve } from "node:path";

import { generateKeyPair } from "../src/proof";
import { createHandler } from "../src/server";
import { createReferenceHost } from "../test/host/reference-host";

const port = Number(process.env.WOPI_E2E_PORT ?? "4190");
const origin = `http://127.0.0.1:${port}`;
const root = resolve(import.meta.dir, "../../..");

const FIXTURES: Record<string, { file: string; name: string }> = {
  docx: { file: resolve(root, "apps/demo/public/betteroffice-demo.docx"), name: "demo.docx" },
  xlsx: { file: resolve(root, "packages/xlsx/test-fixtures/sample.xlsx"), name: "book.xlsx" },
  pptx: { file: resolve(root, "apps/demo/public/betteroffice-demo.pptx"), name: "demo.pptx" },
};

const checkFileInfo = new Map<string, Record<string, unknown>>();
const host = createReferenceHost({ checkFileInfoExtras: (file) => checkFileInfo.get(file.id) ?? {} });
for (const [format, fixture] of Object.entries(FIXTURES)) {
  host.setFile(`file-${format}`, new Uint8Array(readFileSync(fixture.file)), { name: fixture.name });
}
host.tokens.set("token-a", { userId: "user-a", userName: "Alice A", canWrite: true });
host.tokens.set("token-b", { userId: "user-b", userName: "Bob B", canWrite: true });

const config = {
  publicUrl: origin,
  proofKeyFile: "data/proof-key.pem",
  maxFileBytes: 64 * 1024 * 1024,
  hosts: [{ id: "e2e", wopiSrcOrigins: [origin], frameOrigins: [origin] }],
};

const wopi = createHandler({ config, currentKey: generateKeyPair(), autosave: false });

function hostPage(fileId: string): string {
  const src = `${origin}/wopi-host/files/${fileId}`;
  const action = `/wopi/edit?WOPISrc=${encodeURIComponent(src)}&ui=en-US`;
  return `<!doctype html>
<html><head><meta charset="utf-8"><title>Reference host</title></head>
<body>
<script>
window.__wopiMessages = [];
window.addEventListener("message", (event) => {
  try { window.__wopiMessages.push(JSON.parse(event.data)); } catch {}
});
window.__sendToFrames = (messageId) => {
  for (const frame of document.querySelectorAll("iframe")) {
    frame.contentWindow.postMessage(JSON.stringify({ MessageId: messageId, SendTime: Date.now(), Values: {} }), ${JSON.stringify(origin)});
  }
};
</script>
<form method="POST" action="${action}" target="frame-a">
  <input type="hidden" name="access_token" value="token-a">
  <input type="hidden" name="access_token_ttl" value="0">
</form>
<form method="POST" action="${action}" target="frame-b">
  <input type="hidden" name="access_token" value="token-b">
  <input type="hidden" name="access_token_ttl" value="0">
</form>
<iframe name="frame-a" title="Alice" style="width:780px;height:850px"></iframe>
<iframe name="frame-b" title="Bob" style="width:780px;height:850px"></iframe>
<script>
for (const form of document.querySelectorAll("form")) form.submit();
</script>
</body></html>`;
}

const puts: Array<{ format: string; autosave: string | null }> = [];

const server = Bun.serve({
  port,
  idleTimeout: 120,
  fetch: async (request, srv) => {
    const url = new URL(request.url);
    if (url.pathname === "/host-page") {
      const format = url.searchParams.get("format") ?? "docx";
      const fixture = FIXTURES[format];
      if (!fixture) return new Response("unsupported format", { status: 400 });
      const close = url.searchParams.get("close");
      let fileId = `file-${format}`;
      if (close === "message" || close === "url") {
        fileId += `-${close}`;
        host.setFile(fileId, new Uint8Array(readFileSync(fixture.file)), { name: fixture.name });
        checkFileInfo.set(fileId, { ClosePostMessage: close === "message", CloseUrl: close === "url" ? `${origin}/closed` : "" });
      }
      return new Response(hostPage(fileId), { headers: { "Content-Type": "text/html; charset=utf-8" } });
    }
    if (url.pathname === "/closed") {
      return new Response("<!doctype html><title>Closed</title><p>Document closed.</p>", { headers: { "Content-Type": "text/html; charset=utf-8" } });
    }
    if (url.pathname === "/debug/puts") {
      const format = url.searchParams.get("format");
      return Response.json(puts.filter((put) => put.format === format));
    }
    if (url.pathname === "/debug/file") {
      const format = url.searchParams.get("format") ?? "docx";
      const file = host.files.get(url.searchParams.get("file") ?? `file-${format}`);
      if (!file) return new Response("missing", { status: 404 });
      return new Response(Buffer.from(file.bytes).toString("base64"), { headers: { "Content-Type": "text/plain" } });
    }
    if (url.pathname.startsWith("/wopi-host/")) {
      const rewritten = new Request(request.url.replace("/wopi-host/", "/"), request);
      const response = await host.fetch(rewritten);
      if (request.headers.get("X-WOPI-Override") === "PUT" && response.ok) {
        puts.push({ format: url.pathname.split("/").at(-2)!.replace("file-", ""), autosave: request.headers.get("X-COOL-WOPI-IsAutosave") });
      }
      return response;
    }
    return wopi(request, srv as never);
  },
  websocket: wopi.relay.websocket as never,
});

console.log(`WOPI e2e server listening on ${server.url}`);
