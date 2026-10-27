import { createPublicKey, verify, type KeyObject } from "node:crypto";
import { readFileSync } from "node:fs";
import { describe, expect, test } from "bun:test";
import { buildProofBytes, proofKeyAttributes } from "./proof";

interface ProofCase {
  name: string;
  expect: "proof-current" | "proofold-current" | "proof-old" | "invalid";
  valid: boolean;
  accessToken: string;
  timestamp: string;
  url: string;
  proof: string;
  proofOld: string;
}

interface Fixture {
  source: string;
  discovery: {
    value: string;
    modulus: string;
    exponent: string;
    oldvalue: string;
    oldmodulus: string;
    oldexponent: string;
  };
  cases: ProofCase[];
}

const fixture = JSON.parse(
  readFileSync(new URL("./fixtures/microsoft-proof-keys.json", import.meta.url), "utf8"),
) as Fixture;

function publicKey(modulus: string, exponent: string): KeyObject {
  return createPublicKey({
    format: "jwk",
    key: {
      kty: "RSA",
      n: Buffer.from(modulus, "base64").toString("base64url"),
      e: Buffer.from(exponent, "base64").toString("base64url"),
    },
  });
}

function verifyWith(key: KeyObject, bytes: Buffer, proof: string): boolean {
  return verify("RSA-SHA256", bytes, key, Buffer.from(proof, "base64"));
}

function extractAccessToken(url: string): string {
  const marker = "access_token=";
  const start = url.indexOf(marker);
  expect(start).toBeGreaterThanOrEqual(0);
  return url.slice(start + marker.length);
}

const currentKey = publicKey(fixture.discovery.modulus, fixture.discovery.exponent);
const oldKey = publicKey(fixture.discovery.oldmodulus, fixture.discovery.oldexponent);

describe("Microsoft proof key vectors", () => {
  test("names its source", () => {
    expect(fixture.source).toContain("microsoft/Office-Online-Test-Tools-and-Documentation");
    expect(fixture.source).toContain("samples/python/proof_keys/tests.py");
    expect(fixture.source).toContain("MIT");
    expect(fixture.cases).toHaveLength(8);
  });

  test("proofKeyAttributes reproduces the published values", () => {
    expect(proofKeyAttributes(currentKey)).toEqual({
      value: fixture.discovery.value,
      modulus: fixture.discovery.modulus,
      exponent: fixture.discovery.exponent,
    });
    expect(proofKeyAttributes(oldKey)).toEqual({
      value: fixture.discovery.oldvalue,
      modulus: fixture.discovery.oldmodulus,
      exponent: fixture.discovery.oldexponent,
    });
  });

  test("valid cases verify with the key the sample expects", () => {
    for (const entry of fixture.cases.filter((c) => c.valid)) {
      const token = extractAccessToken(entry.url);
      expect(token).toBe(entry.accessToken);
      const bytes = buildProofBytes(token, entry.url, BigInt(entry.timestamp));
      if (entry.expect === "proof-current") {
        expect(verifyWith(currentKey, bytes, entry.proof)).toBe(true);
      } else if (entry.expect === "proofold-current") {
        expect(verifyWith(currentKey, bytes, entry.proofOld)).toBe(true);
      } else if (entry.expect === "proof-old") {
        expect(verifyWith(oldKey, bytes, entry.proof)).toBe(true);
      } else {
        throw new Error(`unexpected valid case ${entry.name}`);
      }
      const anyKey =
        verifyWith(currentKey, bytes, entry.proof) ||
        verifyWith(oldKey, bytes, entry.proof) ||
        verifyWith(currentKey, bytes, entry.proofOld) ||
        verifyWith(oldKey, bytes, entry.proofOld);
      expect(anyKey).toBe(true);
    }
  });

  test("invalid cases fail with both keys and both headers", () => {
    for (const entry of fixture.cases.filter((c) => !c.valid)) {
      const bytes = buildProofBytes(extractAccessToken(entry.url), entry.url, BigInt(entry.timestamp));
      expect(verifyWith(currentKey, bytes, entry.proof)).toBe(false);
      expect(verifyWith(oldKey, bytes, entry.proof)).toBe(false);
      expect(verifyWith(currentKey, bytes, entry.proofOld)).toBe(false);
      expect(verifyWith(oldKey, bytes, entry.proofOld)).toBe(false);
    }
  });
});
