import { beforeAll, describe, expect, test } from "bun:test";
import JSZip from "jszip";

import { blankDocument } from "./blank";
import { initEngines } from "./replica/engines";
import { openServerReplica, type Format } from "./replica/replica";
import { openFreshClient } from "../test/replica/clients";

beforeAll(async () => {
  await initEngines();
});

const FORMATS: Format[] = ["docx", "xlsx", "pptx"];

describe("blank documents", () => {
  test.each(FORMATS)("%s blank is a zip with a content types part", async (format) => {
    const bytes = await blankDocument(format);
    expect(bytes.byteLength).toBeGreaterThan(100);
    expect(bytes[0]).toBe(0x50);
    expect(bytes[1]).toBe(0x4b);
    const zip = await JSZip.loadAsync(bytes);
    expect(zip.file("[Content_Types].xml")).not.toBeNull();
  });

  test.each(FORMATS)("%s blank opens in the server replica, saves and reopens", async (format) => {
    const bytes = await blankDocument(format);
    const replica = await openServerReplica(format, bytes, { clientId: 770001 });
    try {
      const saved = (await replica.save()).bytes;
      expect(saved.byteLength).toBeGreaterThan(100);
      const fresh = await openFreshClient(format, saved, 770002);
      try {
        const resaved = await fresh.save();
        expect(resaved.byteLength).toBeGreaterThan(100);
      } finally {
        fresh.destroy();
      }
    } finally {
      replica.destroy();
    }
  });
});
