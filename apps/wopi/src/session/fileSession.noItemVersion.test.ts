import { beforeAll, expect, test } from "bun:test";
import { initEngines } from "../replica/engines";
import { openTestClient, sourceBytes } from "../../test/replica/clients";
import { addToken, makeHarness, SRC } from "../../test/session/harness";

beforeAll(async () => {
  await initEngines();
});

async function run(stripItemVersion: boolean) {
  const h = makeHarness({ format: "xlsx" });
  addToken(h, "tok-a", "user-a");
  if (stripItemVersion) {
    const real = h.host.fetch.bind(h.host);
    h.host.fetch = async (request: Request) => {
      const response = await real(request);
      if ((request.headers.get("X-WOPI-Override") ?? "").toUpperCase() !== "PUT") return response;
      const headers = new Headers(response.headers);
      headers.delete("X-WOPI-ItemVersion");
      return new Response(await response.arrayBuffer(), { status: response.status, headers });
    };
  }
  const joined = await h.registry.join({ hostId: "testhost", wopiSrc: SRC, accessToken: "tok-a", accessTokenTtl: 0, mode: "edit" });
  const session = [...h.registry.list()][0]!;
  const client = await openTestClient(joined.joinPayload, sourceBytes("xlsx"), 91001);
  const apply = (marker: string, slot: number) => {
    client.edit(marker, slot);
    const replica = session.replica!;
    expect(session.applyUpdate(joined.user.id, client.encodeStateAsUpdate(replica.encodeStateVector())).accepted).toBe(true);
  };
  apply("first-edit ", 0);
  const first = await session.requestSave(joined.user.id);
  apply("second-edit ", 1);
  const second = await session.requestSave(joined.user.id);
  client.destroy();
  return {
    firstSaved: first.saved,
    secondSaved: second.saved,
    plainPuts: h.putCount(),
    relativePuts: h.callsFor("PUT_RELATIVE").length,
  };
}

test("a host that sends X-WOPI-ItemVersion saves twice with PutFile", async () => {
  expect(await run(false)).toEqual({ firstSaved: true, secondSaved: true, plainPuts: 2, relativePuts: 0 });
});

test("host that omits X-WOPI-ItemVersion still saves twice with PutFile and no conflict copy", async () => {
  expect(await run(true)).toEqual({ firstSaved: true, secondSaved: true, plainPuts: 2, relativePuts: 0 });
});
