import { beforeAll, expect, test } from "bun:test";
import { initEngines } from "../replica/engines";
import { openTestClient, sourceBytes } from "../../test/replica/clients";
import { addToken, makeHarness, SRC } from "../../test/session/harness";

beforeAll(async () => {
  await initEngines();
});

test("without X-WOPI-ItemVersion, a same-length file seen right after PutFile keeps both versions", async () => {
  const h = makeHarness({ format: "xlsx" });
  addToken(h, "tok-a", "user-a");
  let external: Uint8Array | null = null;
  const real = h.host.fetch.bind(h.host);
  h.host.fetch = async (request: Request) => {
    const put = (request.headers.get("X-WOPI-Override") ?? "").toUpperCase() === "PUT";
    const sent = put ? (await request.clone().arrayBuffer()).byteLength : 0;
    const response = await real(request);
    if (!put) return response;
    const file = h.host.files.get("file1")!;
    external = new Uint8Array(sent).fill(0x5a);
    file.bytes = external;
    file.version += 1;
    file.itemVersion = String(file.version);
    const headers = new Headers(response.headers);
    headers.delete("X-WOPI-ItemVersion");
    return new Response(await response.arrayBuffer(), { status: response.status, headers });
  };
  const joined = await h.registry.join({ hostId: "testhost", wopiSrc: SRC, accessToken: "tok-a", accessTokenTtl: 0, mode: "edit" });
  const session = [...h.registry.list()][0]!;
  const client = await openTestClient(joined.joinPayload, sourceBytes("xlsx"), 91003);
  client.edit("first-edit ", 0);
  const replica = session.replica!;
  expect(session.applyUpdate(joined.user.id, client.encodeStateAsUpdate(replica.encodeStateVector())).accepted).toBe(true);
  const first = await session.requestSave(joined.user.id);
  client.destroy();
  expect({
    firstSaved: first.saved,
    relativePuts: h.callsFor("PUT_RELATIVE").length,
    externalKept: Buffer.from(h.host.files.get("file1")!.bytes).equals(Buffer.from(external!)),
  }).toEqual({ firstSaved: true, relativePuts: 1, externalKept: true });
});
