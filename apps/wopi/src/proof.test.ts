import { createPublicKey, verify } from "node:crypto";
import { mkdtempSync, statSync } from "node:fs";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { describe, expect, test } from "bun:test";
import {
  createProofSigner,
  dotnetTicks,
  exportPrivateKeyPem,
  generateKeyPair,
  keyPairFromPrivatePem,
  loadOrCreateKeyPairFile,
  proofKeyAttributes,
} from "./proof";

function expectedProofBytes(accessToken: string, url: string, ticks: bigint): Buffer {
  const tokenBytes = Buffer.from(accessToken, "utf8");
  const urlBytes = Buffer.from(url.toUpperCase(), "utf8");
  const header = Buffer.alloc(4);
  header.writeInt32BE(tokenBytes.length, 0);
  const urlHeader = Buffer.alloc(4);
  urlHeader.writeInt32BE(urlBytes.length, 0);
  const lengthHeader = Buffer.alloc(4);
  lengthHeader.writeInt32BE(8, 0);
  const timestamp = Buffer.alloc(8);
  timestamp.writeBigInt64BE(ticks, 0);
  return Buffer.concat([header, tokenBytes, urlHeader, urlBytes, lengthHeader, timestamp]);
}

function publicKeyFromAttributes(attributes: { modulus: string; exponent: string }) {
  const toBase64Url = (base64: string) =>
    Buffer.from(base64, "base64").toString("base64url");
  return createPublicKey({
    format: "jwk",
    key: {
      kty: "RSA",
      n: toBase64Url(attributes.modulus),
      e: toBase64Url(attributes.exponent),
    },
  });
}

function verifySignature(attributes: { modulus: string; exponent: string }, data: Buffer, proof: string): boolean {
  return verify("RSA-SHA256", data, publicKeyFromAttributes(attributes), Buffer.from(proof, "base64"));
}

describe("dotnetTicks", () => {
  test("matches known dates", () => {
    expect(dotnetTicks(new Date("1970-01-01T00:00:00.000Z"))).toBe(621355968000000000n);
    expect(dotnetTicks(new Date("2020-01-01T00:00:00.000Z"))).toBe(637134336000000000n);
    expect(dotnetTicks(new Date("1970-01-01T00:00:00.001Z"))).toBe(621355968000010000n);
  });
});

describe("proofKeyAttributes", () => {
  test("builds a CSP PUBLICKEYBLOB with little-endian modulus", () => {
    const pair = generateKeyPair();
    const attributes = proofKeyAttributes(pair.publicKey);
    const blob = Buffer.from(attributes.value, "base64");
    const modulus = Buffer.from(attributes.modulus, "base64");
    const exponent = Buffer.from(attributes.exponent, "base64");

    expect(attributes.value.startsWith("BgIAAACkAABSU0Ex")).toBe(true);
    expect(blob[0]).toBe(0x06);
    expect(blob[1]).toBe(0x02);
    expect(blob[2]).toBe(0);
    expect(blob[3]).toBe(0);
    expect(blob.subarray(8, 12).toString("ascii")).toBe("RSA1");
    expect(blob.readUInt32LE(12)).toBe(2048);
    expect(blob.readUInt32LE(16)).toBe(65537);
    expect(blob.length).toBe(20 + 256);
    expect(modulus.length).toBe(256);
    expect([...exponent]).toEqual([1, 0, 1]);
    expect(Buffer.from(blob.subarray(20)).reverse()).toEqual(modulus);
  });

  test("accepts a private key and a PEM string", () => {
    const pair = generateKeyPair();
    const expected = proofKeyAttributes(pair.publicKey);
    expect(proofKeyAttributes(pair.privateKey)).toEqual(expected);
    expect(proofKeyAttributes(exportPrivateKeyPem(pair.privateKey))).toEqual(expected);
  });

  test("roundtrips a key pair through PEM", () => {
    const pair = generateKeyPair();
    const restored = keyPairFromPrivatePem(exportPrivateKeyPem(pair.privateKey));
    expect(proofKeyAttributes(restored.publicKey)).toEqual(proofKeyAttributes(pair.publicKey));
  });

  test("persists a key file once and reloads the same key with mode 0600", () => {
    const path = join(mkdtempSync(join(tmpdir(), "wopi-proof-")), "nested", "proof-key.pem");
    const first = loadOrCreateKeyPairFile(path);
    const second = loadOrCreateKeyPairFile(path);
    expect(proofKeyAttributes(second.publicKey)).toEqual(proofKeyAttributes(first.publicKey));
    expect(statSync(path).mode & 0o777).toBe(0o600);
  });
});

describe("signProof", () => {
  test("signs the documented byte layout and fails on tampering", () => {
    const pair = generateKeyPair();
    const oldPair = generateKeyPair();
    const signer = createProofSigner(pair, oldPair);
    const ticks = dotnetTicks(new Date("2024-05-06T07:08:09.000Z"));
    const input = {
      accessToken: "token-value",
      url: "https://cloud.example.com/remote.php/dav/files/demo.docx",
      ticks,
    };

    const result = signer.signProof(input);
    expect(result.timestamp).toBe(ticks.toString());
    expect(result.proofOld).toBeDefined();

    const bytes = expectedProofBytes(input.accessToken, input.url, ticks);
    expect(verifySignature(proofKeyAttributes(pair.publicKey), bytes, result.proof)).toBe(true);
    expect(verifySignature(proofKeyAttributes(oldPair.publicKey), bytes, result.proofOld)).toBe(true);

    expect(
      verifySignature(
        proofKeyAttributes(pair.publicKey),
        expectedProofBytes(input.accessToken, "https://cloud.example.com/other.docx", ticks),
        result.proof,
      ),
    ).toBe(false);
    expect(
      verifySignature(
        proofKeyAttributes(pair.publicKey),
        expectedProofBytes("other-token", input.url, ticks),
        result.proof,
      ),
    ).toBe(false);
    expect(
      verifySignature(
        proofKeyAttributes(pair.publicKey),
        expectedProofBytes(input.accessToken, input.url, ticks + 1n),
        result.proof,
      ),
    ).toBe(false);
  });

  test("signs proofOld with the current key without an old key", () => {
    const pair = generateKeyPair();
    const signer = createProofSigner(pair);
    const input = { accessToken: "t", url: "https://example.com/f.docx", ticks: 1n };
    const result = signer.signProof(input);
    expect(result.proofOld).toBe(result.proof);
    expect(signer.oldProofKeyAttributes()).toEqual(proofKeyAttributes(pair.publicKey));
    expect(
      verifySignature(
        proofKeyAttributes(pair.publicKey),
        expectedProofBytes(input.accessToken, input.url, input.ticks),
        result.proofOld,
      ),
    ).toBe(true);
  });
});
