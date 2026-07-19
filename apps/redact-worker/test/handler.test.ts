import { describe, expect, test } from "bun:test";

import { createWorker, type Env, type Sanitizer } from "../src/handler";

const DOCX_TYPE =
  "application/vnd.openxmlformats-officedocument.wordprocessingml.document";
const TYPES = {
  docx: DOCX_TYPE,
  xlsx: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
  pptx: "application/vnd.openxmlformats-officedocument.presentationml.presentation",
} as const;

interface StoredObject {
  bytes: Uint8Array;
  contentType: string;
  format: string;
}

class MemoryBucket {
  readonly objects = new Map<string, StoredObject>();

  async put(
    key: string,
    value: Uint8Array,
    options: R2PutOptions,
  ): Promise<R2Object> {
    const httpMetadata = options.httpMetadata;
    this.objects.set(key, {
      bytes: value.slice(),
      contentType:
        httpMetadata && !(httpMetadata instanceof Headers)
          ? httpMetadata.contentType ?? "application/octet-stream"
          : "application/octet-stream",
      format: options.customMetadata?.format ?? "",
    });
    return { key } as R2Object;
  }

  async get(key: string): Promise<R2ObjectBody | null> {
    const stored = this.objects.get(key);
    if (!stored) return null;
    return {
      key,
      size: stored.bytes.byteLength,
      body: new Blob([stored.bytes.slice().buffer as ArrayBuffer]).stream(),
      httpEtag: `"${key}"`,
      httpMetadata: { contentType: stored.contentType },
      customMetadata: { format: stored.format },
    } as unknown as R2ObjectBody;
  }
}

describe("redaction worker", () => {
  test("stores only sanitizer output under an opaque id", async () => {
    const bucket = new MemoryBucket();
    let received: Uint8Array | undefined;
    const sanitizer: Sanitizer = (bytes, format) => {
      received = bytes.slice();
      expect(format).toBe("docx");
      return new Uint8Array([9, 8, 7]);
    };
    const worker = createWorker(sanitizer);
    const response = await worker.fetch(uploadRequest(new Uint8Array([1, 2, 3])), env(bucket));
    const result = (await response.json()) as { id: string; url: string };

    expect(response.status).toBe(200);
    expect(received).toEqual(new Uint8Array([1, 2, 3]));
    expect(result.id).toMatch(/^[a-f0-9]{32}$/);
    expect(result.url).toBe(`https://redact.test/f/${result.id}`);
    expect(bucket.objects.get(result.id)?.bytes).toEqual(new Uint8Array([9, 8, 7]));
  });

  test("rejects mismatched and oversized uploads", async () => {
    const bucket = new MemoryBucket();
    const worker = createWorker((bytes) => bytes);
    const mismatch = uploadRequest(new Uint8Array([1]));
    mismatch.headers.set("Content-Type", "application/zip");
    expect((await worker.fetch(mismatch, env(bucket))).status).toBe(415);

    const oversized = uploadRequest(new Uint8Array([1]));
    oversized.headers.set("Content-Length", String(64 * 1024 * 1024 + 1));
    expect((await worker.fetch(oversized, env(bucket))).status).toBe(413);
    expect(bucket.objects.size).toBe(0);
  });

  test("accepts the XLSX and PPTX upload labels", async () => {
    const bucket = new MemoryBucket();
    const formats: string[] = [];
    const worker = createWorker((bytes, format) => {
      formats.push(format);
      return bytes;
    });
    expect(
      (await worker.fetch(uploadRequest(new Uint8Array([1]), "xlsx"), env(bucket))).status,
    ).toBe(200);
    expect(
      (await worker.fetch(uploadRequest(new Uint8Array([2]), "pptx"), env(bucket))).status,
    ).toBe(200);
    expect(formats).toEqual(["xlsx", "pptx"]);
  });

  test("rejects sanitizer failures and limits repeated clients", async () => {
    const bucket = new MemoryBucket();
    const rejecting = createWorker(() => {
      throw new Error("bad zip");
    });
    expect(
      (await rejecting.fetch(uploadRequest(new Uint8Array([1])), env(bucket))).status,
    ).toBe(400);

    const limited = createWorker((bytes) => bytes, { rateLimit: 1 });
    expect(
      (await limited.fetch(uploadRequest(new Uint8Array([1])), env(bucket))).status,
    ).toBe(200);
    expect(
      (await limited.fetch(uploadRequest(new Uint8Array([1])), env(bucket))).status,
    ).toBe(429);
  });

  test("retrieves sanitized files without a filename", async () => {
    const bucket = new MemoryBucket();
    const worker = createWorker(() => new Uint8Array([4, 5, 6]));
    const uploaded = await worker.fetch(uploadRequest(new Uint8Array([1])), env(bucket));
    const { id } = (await uploaded.json()) as { id: string };
    const response = await worker.fetch(
      new Request(`https://redact.test/f/${id}`),
      env(bucket),
    );
    expect(response.status).toBe(200);
    expect(new Uint8Array(await response.arrayBuffer())).toEqual(new Uint8Array([4, 5, 6]));
    expect(response.headers.get("Content-Disposition")).toBe("attachment");
    expect(response.headers.get("Content-Disposition")).not.toContain("filename");
  });

  test("builds safely without an R2 binding", async () => {
    const worker = createWorker((bytes) => bytes);
    const upload = await worker.fetch(uploadRequest(new Uint8Array([1])), {});
    const retrieve = await worker.fetch(
      new Request("https://redact.test/f/0123456789abcdef0123456789abcdef"),
      {},
    );
    expect(upload.status).toBe(503);
    expect(retrieve.status).toBe(503);
  });
});

function uploadRequest(
  bytes: Uint8Array,
  format: keyof typeof TYPES = "docx",
): Request {
  return new Request("https://redact.test/upload", {
    method: "POST",
    headers: {
      "CF-Connecting-IP": "192.0.2.10",
      "Content-Type": TYPES[format],
      "X-BetterOffice-Format": format,
    },
    body: bytes.slice().buffer as ArrayBuffer,
  });
}

function env(bucket: MemoryBucket): Env {
  return { REDACTED_BUCKET: bucket as unknown as R2Bucket };
}
