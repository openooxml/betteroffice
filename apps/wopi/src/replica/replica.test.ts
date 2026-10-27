import { beforeAll, describe, expect, test } from "bun:test";

import {
  byteEqual,
  differingParts,
  openFreshClient,
  openTestClient,
  sourceBytes,
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
        expect([...client.encodeStateVector()]).toEqual([...server.encodeStateVector()]);
        expect(server.version).toBe(0);

        client.edit("replica-marker ");
        const update = client.encodeStateAsUpdate(server.encodeStateVector());
        expect(server.applyUpdate(update)).toEqual({ accepted: true });
        expect(server.version).toBe(1);

        expect(server.applyUpdate(update)).toEqual({ accepted: true });
        expect(server.version).toBe(2);
        expect(server.applyUpdate(new Uint8Array(0))).toEqual({ accepted: true });
        expect(server.applyUpdate(new Uint8Array([0, 0]))).toEqual({ accepted: true });
        expect(server.version).toBe(2);

        const saved = await server.save();
        expect(server.version).toBe(2);
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
      const late = await (async () => {
        early.edit("early-one ", 0);
        server.applyUpdate(early.encodeStateAsUpdate(server.encodeStateVector()));
        early.edit("early-two ", 0);
        server.applyUpdate(early.encodeStateAsUpdate(server.encodeStateVector()));
        return openTestClient(server.joinPayload(), source, clientId());
      })();
      try {
        expect(server.version).toBe(2);
        expect([...late.encodeStateVector()]).toEqual([...server.encodeStateVector()]);
        expect(late.contains("early-one")).toBe(true);
        expect(late.contains("early-two")).toBe(true);

        late.edit("late-marker ", 0);
        expect(
          server.applyUpdate(late.encodeStateAsUpdate(server.encodeStateVector())),
        ).toEqual({ accepted: true });
        expect(server.version).toBe(3);
        const saved = await server.save();
        const fresh = await openFreshClient(format, saved, clientId());
        try {
          expect(fresh.contains("late-marker")).toBe(true);
        } finally {
          fresh.destroy();
        }
      } finally {
        early.destroy();
        late.destroy();
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
