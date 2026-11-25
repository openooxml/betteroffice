import { beforeAll, expect, test } from "bun:test";
import { initEngines } from "../replica/engines";
import { openTestClient, sourceBytes } from "../../test/replica/clients";
import { addToken, makeHarness, SRC } from "../../test/session/harness";

beforeAll(async () => {
  await initEngines();
});

async function run(stripInfoVersion: boolean, seed: number) {
  const h = makeHarness({ format: "xlsx" });
  addToken(h, "tok-a", "user-a");
  let external: Uint8Array | null = null;
  const real = h.host.fetch.bind(h.host);
  h.host.fetch = async (request: Request) => {
    const put = (request.headers.get("X-WOPI-Override") ?? "").toUpperCase() === "PUT" && external === null;
    const sent = put ? (await request.clone().arrayBuffer()).byteLength : 0;
    const response = await real(request);
    if (!put) {
      if (!stripInfoVersion || !(response.headers.get("content-type") ?? "").includes("json")) return response;
      const info = await response.json();
      delete info.Version;
      const headers = new Headers(response.headers);
      headers.delete("content-length");
      return new Response(JSON.stringify(info), { status: response.status, headers });
    }
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
  const client = await openTestClient(joined.joinPayload, sourceBytes("xlsx"), seed);
  const replica = session.replica!;
  client.edit("first-edit ", 0);
  expect(session.applyUpdate(joined.user.id, client.encodeStateAsUpdate(replica.encodeStateVector())).accepted).toBe(true);
  const first = await session.requestSave(joined.user.id);
  client.edit("second-edit ", 1);
  const secondAccepted = session.applyUpdate(joined.user.id, client.encodeStateAsUpdate(replica.encodeStateVector())).accepted;
  const second = await session.requestSave(joined.user.id);
  client.destroy();
  const outcome = {
    firstSaved: first.saved,
    secondAccepted,
    secondSaved: second.saved,
    plainPuts: h.putCount(),
    relativePuts: h.callsFor("PUT_RELATIVE").length,
    externalInFile: Buffer.from(h.host.files.get("file1")!.bytes).equals(Buffer.from(external!)),
  };
  console.log("same-length outcome", JSON.stringify({ stripInfoVersion, ...outcome }));
  return outcome;
}

for (const stripInfoVersion of [false, true]) {
  test(`without X-WOPI-ItemVersion${stripInfoVersion ? " or a file info Version" : ""}, a same-length file seen right after PutFile keeps both versions`, async () => {
    const outcome = await run(stripInfoVersion, stripInfoVersion ? 91004 : 91003);
    expect(outcome.externalInFile).toBe(true);
    expect(outcome.firstSaved).toBe(true);
    expect(outcome.relativePuts).toBe(1);
  });
}
