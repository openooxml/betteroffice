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

const host = createReferenceHost();
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

const wopi = createHandler({ config, currentKey: generateKeyPair() });

function hostPage(format: string): string {
  const src = `${origin}/wopi-host/files/file-${format}`;
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

const server = Bun.serve({
  port,
  idleTimeout: 120,
  fetch: async (request, srv) => {
    const url = new URL(request.url);
    if (url.pathname === "/host-page") {
      const format = url.searchParams.get("format") ?? "docx";
      return new Response(hostPage(format), { headers: { "Content-Type": "text/html; charset=utf-8" } });
    }
    if (url.pathname === "/debug/file") {
      const format = url.searchParams.get("format") ?? "docx";
      const file = host.files.get(`file-${format}`);
      if (!file) return new Response("missing", { status: 404 });
      return new Response(Buffer.from(file.bytes).toString("base64"), { headers: { "Content-Type": "text/plain" } });
    }
    if (url.pathname.startsWith("/wopi-host/")) {
      const rewritten = new Request(request.url.replace("/wopi-host/", "/"), request);
      return host.fetch(rewritten);
    }
    return wopi(request, srv as never);
  },
  websocket: wopi.relay.websocket as never,
});

console.log(`WOPI e2e server listening on ${server.url}`);
