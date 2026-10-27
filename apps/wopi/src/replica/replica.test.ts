import { beforeAll, describe, expect, test } from "bun:test";

import {
  byteEqual,
  differingParts,
  openFreshClient,
  openTestClient,
  sourceBytes,
  stateVectorClientIds,
  unzipParts,
  type TestClient,
} from "../../test/replica/clients";
import { initEngines } from "./engines";
import { openServerReplica, type Format, type ServerReplica } from "./replica";

const FORMATS: Format[] = ["docx", "xlsx", "pptx"];

let nextClientId = 10_000;
function clientId(): number {
  nextClientId += 7;
  return nextClientId;
}

beforeAll(async () => {
  await initEngines();
});

function expectDocxSavesEqual(a: Uint8Array, b: Uint8Array): void {
  const differing = differingParts(a, b);
  expect(differing.filter((part) => part !== "docProps/core.xml")).toEqual([]);
  if (differing.includes("docProps/core.xml")) {
    const strip = (bytes: Uint8Array) =>
      new TextDecoder()
        .decode(bytes)
        .replace(
          /<dcterms:modified[^>]*>[^<]*<\/dcterms:modified>/,
          "<dcterms:modified/>",
        );
    const coreA = unzipParts(a)["docProps/core.xml"]!;
    const coreB = unzipParts(b)["docProps/core.xml"]!;
    expect(strip(coreA)).toBe(strip(coreB));
  }
}

function expectSavesEqual(format: Format, a: Uint8Array, b: Uint8Array): void {
  if (format === "docx") expectDocxSavesEqual(a, b);
  else expect(byteEqual(a, b)).toBe(true);
}

async function syncClient(server: ServerReplica, client: TestClient): Promise<void> {
  client.applyUpdate(server.encodeStateAsUpdate(client.encodeStateVector()));
}

for (const format of FORMATS) {
  describe(`${format} server replica`, () => {
    test("open, join, client edit, save round trip and version", async () => {
      const source = sourceBytes(format);
      const server = await openServerReplica(format, source, { clientId: clientId() });
      const client = await openTestClient(server.joinPayload(), source, clientId());
      try {
        expect(server.format).toBe(format);
        expect(Object.keys(server.joinPayload()).sort()).toEqual([
          "format",
          "initialUpdate",
        ]);
        expect([...client.encodeStateVector()]).toEqual([...server.encodeStateVector()]);
        expect(server.version).toBe(0);

        client.edit("replica-marker ");
        const update = client.encodeStateAsUpdate(server.encodeStateVector());
        expect(server.applyUpdate(update)).toEqual({ accepted: true });
        expect(server.version).toBe(1);

        expect(server.applyUpdate(update)).toEqual({ accepted: true });
        const afterDuplicate = format === "xlsx" ? 1 : 2;
        expect(server.version).toBe(afterDuplicate);
        expect(server.applyUpdate(new Uint8Array(0))).toEqual({ accepted: true });
        expect(server.applyUpdate(new Uint8Array([0, 0]))).toEqual({ accepted: true });
        expect(server.version).toBe(afterDuplicate);

        const saved = await server.save();
        expect(server.version).toBe(afterDuplicate);
        const fresh = await openFreshClient(format, saved, clientId());
        try {
          expect(fresh.contains("replica-marker")).toBe(true);
        } finally {
          fresh.destroy();
        }

        const [saveA, saveB] = await Promise.all([server.save(), server.save()]);
        expectSavesEqual(format, saveA, saveB);
        expect([...client.encodeStateVector()]).toEqual([...server.encodeStateVector()]);
      } finally {
        client.destroy();
        server.destroy();
      }
    });

    test("three replicas with concurrent edits through the server converge", async () => {
      const source = sourceBytes(format);
      const server = await openServerReplica(format, source, { clientId: clientId() });
      const a = await openTestClient(server.joinPayload(), source, clientId());
      const b = await openTestClient(server.joinPayload(), source, clientId());
      try {
        a.edit("AAA ", 0);
        b.edit("BBB ", 1);
        expect(server.applyUpdate(a.encodeStateAsUpdate(server.encodeStateVector())))
          .toEqual({ accepted: true });
        expect(server.applyUpdate(b.encodeStateAsUpdate(server.encodeStateVector())))
          .toEqual({ accepted: true });
        expect(server.version).toBe(2);
        await syncClient(server, a);
        await syncClient(server, b);
        expect([...a.encodeStateVector()]).toEqual([...server.encodeStateVector()]);
        expect([...b.encodeStateVector()]).toEqual([...server.encodeStateVector()]);
        expect(a.contains("BBB")).toBe(true);
        expect(b.contains("AAA")).toBe(true);

        const serverSave = await server.save();
        expectSavesEqual(format, serverSave, await a.save());
        expectSavesEqual(format, serverSave, await b.save());
      } finally {
        a.destroy();
        b.destroy();
        server.destroy();
      }
    });

    test("delete-only update bumps version and saves without the text", async () => {
      const source = sourceBytes(format);
      const server = await openServerReplica(format, source, { clientId: clientId() });
      const client = await openTestClient(server.joinPayload(), source, clientId());
      try {
        client.edit("DELETE-ME ");
        expect(
          server.applyUpdate(client.encodeStateAsUpdate(server.encodeStateVector())),
        ).toEqual({ accepted: true });
        expect(server.version).toBe(1);
        const before = await openFreshClient(format, await server.save(), clientId());
        try {
          expect(before.contains("DELETE-ME")).toBe(true);
        } finally {
          before.destroy();
        }

        client.remove("DELETE-ME");
        const deletion = client.encodeStateAsUpdate(server.encodeStateVector());
        expect(server.applyUpdate(deletion)).toEqual({ accepted: true });
        expect(server.version).toBe(2);

        const after = await openFreshClient(format, await server.save(), clientId());
        try {
          expect(after.contains("DELETE-ME")).toBe(false);
        } finally {
          after.destroy();
        }
      } finally {
        client.destroy();
        server.destroy();
      }
    });

    test("rejected updates leave state vector, version and save unchanged", async () => {
      const source = sourceBytes(format);
      const server = await openServerReplica(format, source, { clientId: clientId() });
      const client = await openTestClient(server.joinPayload(), source, clientId());
      try {
        client.edit("kept-marker ");
        const valid = client.encodeStateAsUpdate(server.encodeStateVector());
        const truncated = valid.slice(0, Math.max(1, Math.floor(valid.length / 2)));
        const corrupt: Uint8Array[] = [
          new Uint8Array([255, 254, 253, 252, 251, 250, 249, 248]),
          truncated,
        ];
        for (const bad of corrupt) {
          const vectorBefore = server.encodeStateVector();
          const versionBefore = server.version;
          const saveBefore = await server.save();
          const result = server.applyUpdate(bad);
          expect(result.accepted).toBe(false);
          if (!result.accepted) expect(result.reason.length).toBeGreaterThan(0);
          expect([...server.encodeStateVector()]).toEqual([...vectorBefore]);
          expect(server.version).toBe(versionBefore);
          expectSavesEqual(format, await server.save(), saveBefore);
        }
      } finally {
        client.destroy();
        server.destroy();
      }
    });

    test("late joiner converges without re-importing", async () => {
      const source = sourceBytes(format);
      const server = await openServerReplica(format, source, { clientId: clientId() });
      const early = await openTestClient(server.joinPayload(), source, clientId());
      let late: TestClient | undefined;
      try {
        early.edit("early-one ", 0);
        expect(
          server.applyUpdate(early.encodeStateAsUpdate(server.encodeStateVector())),
        ).toEqual({ accepted: true });
        early.edit("early-two ", 1);
        expect(
          server.applyUpdate(early.encodeStateAsUpdate(server.encodeStateVector())),
        ).toEqual({ accepted: true });
        expect(server.version).toBe(2);

        late = await openTestClient(server.joinPayload(), source, clientId());
        expect([...late.encodeStateVector()]).toEqual([...server.encodeStateVector()]);
        expect(late.contains("early-one")).toBe(true);
        expect(late.contains("early-two")).toBe(true);

        late.edit("late-marker ", 0);
        expect(
          server.applyUpdate(late.encodeStateAsUpdate(server.encodeStateVector())),
        ).toEqual({ accepted: true });
        expect(server.version).toBe(3);
        await syncClient(server, early);
        expect(early.contains("late-marker")).toBe(true);
        expect([...early.encodeStateVector()]).toEqual([...server.encodeStateVector()]);

        const saved = await server.save();
        const fresh = await openFreshClient(format, saved, clientId());
        try {
          expect(fresh.contains("late-marker")).toBe(true);
          expect(fresh.contains("early-two")).toBe(true);
        } finally {
          fresh.destroy();
        }
      } finally {
        early.destroy();
        late?.destroy();
        server.destroy();
      }
    });

    test("two late joiners edit concurrently under their own client ids", async () => {
      const source = sourceBytes(format);
      const server = await openServerReplica(format, source, { clientId: clientId() });
      const early = await openTestClient(server.joinPayload(), source, clientId());
      let lateA: TestClient | undefined;
      let lateB: TestClient | undefined;
      try {
        early.edit("early-one ", 0);
        server.applyUpdate(early.encodeStateAsUpdate(server.encodeStateVector()));
        early.edit("early-two ", 1);
        server.applyUpdate(early.encodeStateAsUpdate(server.encodeStateVector()));

        lateA = await openTestClient(server.joinPayload(), source, clientId());
        lateB = await openTestClient(server.joinPayload(), source, clientId());
        expect(lateA.contains("early-one")).toBe(true);
        expect(lateB.contains("early-two")).toBe(true);
        expect(new Set([server.clientId, lateA.clientId, lateB.clientId]).size).toBe(3);

        lateA.edit("from-late-a ", 2);
        lateB.edit("from-late-b ", 3);
        expect(stateVectorClientIds(lateA.encodeStateVector())).toContain(lateA.clientId);
        expect(stateVectorClientIds(lateB.encodeStateVector())).toContain(lateB.clientId);

        expect(
          server.applyUpdate(lateA.encodeStateAsUpdate(server.encodeStateVector())),
        ).toEqual({ accepted: true });
        expect(
          server.applyUpdate(lateB.encodeStateAsUpdate(server.encodeStateVector())),
        ).toEqual({ accepted: true });
        await syncClient(server, lateA);
        await syncClient(server, lateB);
        await syncClient(server, early);

        for (const replica of [lateA, lateB, early]) {
          expect([...replica.encodeStateVector()]).toEqual([
            ...server.encodeStateVector(),
          ]);
          expect(replica.contains("from-late-a")).toBe(true);
          expect(replica.contains("from-late-b")).toBe(true);
        }
        const serverIds = stateVectorClientIds(server.encodeStateVector());
        expect(serverIds).toContain(lateA.clientId);
        expect(serverIds).toContain(lateB.clientId);

        const fresh = await openFreshClient(format, await server.save(), clientId());
        try {
          expect(fresh.contains("from-late-a")).toBe(true);
          expect(fresh.contains("from-late-b")).toBe(true);
        } finally {
          fresh.destroy();
        }
      } finally {
        early.destroy();
        lateA?.destroy();
        lateB?.destroy();
        server.destroy();
      }
    });

    test("calls after destroy throw", async () => {
      const source = sourceBytes(format);
      const server = await openServerReplica(format, source, { clientId: clientId() });
      server.destroy();
      expect(() => server.encodeStateVector()).toThrow();
      expect(() => server.encodeStateAsUpdate()).toThrow();
      expect(() => server.applyUpdate(new Uint8Array([0, 0]))).toThrow();
      expect(() => server.joinPayload()).toThrow();
      expect(() => server.save()).toThrow();
      expect(() => server.destroy()).toThrow();
    });
  });
}
