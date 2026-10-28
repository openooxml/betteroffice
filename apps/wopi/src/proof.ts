import {
  constants,
  createPrivateKey,
  createPublicKey,
  createSign,
  generateKeyPairSync,
  randomUUID,
  type KeyObject,
} from "node:crypto";
import { existsSync, linkSync, mkdirSync, readFileSync, unlinkSync, writeFileSync } from "node:fs";
import { dirname, join } from "node:path";

export interface ProofKeyPair {
  privateKey: KeyObject;
  publicKey: KeyObject;
}

export interface ProofKeyAttributes {
  value: string;
  modulus: string;
  exponent: string;
}

export type ProofKeyInput = KeyObject | string | ProofKeyPair;

export interface SignProofInput {
  accessToken: string;
  url: string;
  ticks: bigint;
}

export interface SignedProof {
  proof: string;
  proofOld: string;
  timestamp: string;
}

const DOTNET_EPOCH_TICKS = 621355968000000000n;
const TICKS_PER_MILLISECOND = 10000n;

/** .NET ticks (100 ns units since 0001-01-01 UTC) for a date. */
export function dotnetTicks(date: Date): bigint {
  return BigInt(date.getTime()) * TICKS_PER_MILLISECOND + DOTNET_EPOCH_TICKS;
}

export function generateKeyPair(): ProofKeyPair {
  const { privateKey, publicKey } = generateKeyPairSync("rsa", { modulusLength: 2048 });
  return { privateKey, publicKey };
}

function publicFromPrivate(privateKey: KeyObject): KeyObject {
  return createPublicKey({ key: privateKey.export({ format: "jwk" }), format: "jwk" });
}

function toPublicKey(key: ProofKeyInput): KeyObject {
  if (typeof key === "object" && "publicKey" in key) return key.publicKey;
  if (typeof key === "string") {
    try {
      return createPublicKey(key);
    } catch {
      return publicFromPrivate(createPrivateKey(key));
    }
  }
  return key.type === "private" ? publicFromPrivate(key) : key;
}

function bigEndianBytes(base64Url: string): Buffer {
  return Buffer.from(base64Url, "base64url");
}

function jwkBytes(key: ProofKeyInput): { modulus: Buffer; exponent: Buffer } {
  const jwk = toPublicKey(key).export({ format: "jwk" }) as { n?: string; e?: string };
  if (!jwk.n || !jwk.e) throw new Error("RSA public key is missing modulus or exponent");
  return { modulus: bigEndianBytes(jwk.n), exponent: bigEndianBytes(jwk.e) };
}

function cspPublicKeyBlob(modulus: Buffer, exponent: Buffer): Buffer {
  const bitLength = modulus.length * 8;
  const exponentValue = exponent.reduce((value, byte) => value * 256 + byte, 0);
  const header = Buffer.alloc(20);
  header[0] = 0x06;
  header[1] = 0x02;
  header.writeUInt16LE(0, 2);
  header.writeUInt32LE(0x0000a400, 4);
  header.write("RSA1", 8, "ascii");
  header.writeUInt32LE(bitLength, 12);
  header.writeUInt32LE(exponentValue, 16);
  return Buffer.concat([header, Buffer.from(modulus).reverse()]);
}

export function proofKeyAttributes(key: ProofKeyInput): ProofKeyAttributes {
  const { modulus, exponent } = jwkBytes(key);
  return {
    value: cspPublicKeyBlob(modulus, exponent).toString("base64"),
    modulus: modulus.toString("base64"),
    exponent: exponent.toString("base64"),
  };
}

export function keyPairFromPrivateKey(privateKey: KeyObject): ProofKeyPair {
  if (privateKey.type !== "private") throw new Error("Expected an RSA private key");
  if (privateKey.asymmetricKeyType !== "rsa") throw new Error("Proof key must be an RSA key");
  const modulusLength = (
    privateKey.asymmetricKeyDetails as { modulusLength?: number } | undefined
  )?.modulusLength;
  if (modulusLength !== 2048) throw new Error("Proof key must be a 2048-bit RSA key");
  return { privateKey, publicKey: publicFromPrivate(privateKey) };
}

export function keyPairFromPrivatePem(pem: string): ProofKeyPair {
  return keyPairFromPrivateKey(createPrivateKey(pem));
}

export function exportPrivateKeyPem(key: KeyObject | string): string {
  const privateKey = typeof key === "string" ? createPrivateKey(key) : key;
  if (privateKey.type !== "private") throw new Error("Expected an RSA private key");
  return privateKey.export({ format: "pem", type: "pkcs8" }).toString();
}

export function loadKeyPairFile(path: string): ProofKeyPair {
  return keyPairFromPrivatePem(readFileSync(path, "utf8"));
}

export function publishKeyPairFile(path: string, pair: ProofKeyPair): ProofKeyPair {
  const directory = dirname(path);
  mkdirSync(directory, { recursive: true });
  const tempPath = join(directory, `.proof-key-${process.pid}-${randomUUID()}.tmp`);
  writeFileSync(tempPath, exportPrivateKeyPem(pair.privateKey), { encoding: "utf8", mode: 0o600 });
  try {
    linkSync(tempPath, path);
    return pair;
  } catch (error) {
    if ((error as NodeJS.ErrnoException).code === "EEXIST") return loadKeyPairFile(path);
    throw error;
  } finally {
    try {
      unlinkSync(tempPath);
    } catch {}
  }
}

export function loadOrCreateKeyPairFile(path: string): ProofKeyPair {
  if (existsSync(path)) return loadKeyPairFile(path);
  return publishKeyPairFile(path, generateKeyPair());
}

/** The exact byte layout signed for a WOPI proof header. */
export function buildProofBytes(accessToken: string, url: string, ticks: bigint): Buffer {
  if (typeof ticks !== "bigint") throw new TypeError("ticks must be a bigint");
  const tokenBytes = Buffer.from(accessToken, "utf8");
  const urlBytes = Buffer.from(url.toUpperCase(), "utf8");
  const timestamp = Buffer.alloc(8);
  timestamp.writeBigInt64BE(ticks, 0);
  const parts: Buffer[] = [];
  const tokenLength = Buffer.alloc(4);
  tokenLength.writeInt32BE(tokenBytes.length, 0);
  parts.push(tokenLength, tokenBytes);
  const urlLength = Buffer.alloc(4);
  urlLength.writeInt32BE(urlBytes.length, 0);
  parts.push(urlLength, urlBytes);
  const ticksLength = Buffer.alloc(4);
  ticksLength.writeInt32BE(8, 0);
  parts.push(ticksLength, timestamp);
  return Buffer.concat(parts);
}

function signBytes(privateKey: KeyObject, bytes: Buffer): string {
  const signer = createSign("RSA-SHA256");
  signer.update(bytes);
  signer.end();
  return signer.sign({ key: privateKey, padding: constants.RSA_PKCS1_PADDING }).toString("base64");
}

export interface ProofSignerKeys {
  privateKey: KeyObject;
  oldPrivateKey?: KeyObject;
}

export function signProof(input: SignProofInput, keys: ProofSignerKeys): SignedProof {
  const ticks = input.ticks;
  const bytes = buildProofBytes(input.accessToken, input.url, ticks);
  return {
    proof: signBytes(keys.privateKey, bytes),
    proofOld: signBytes(keys.oldPrivateKey ?? keys.privateKey, bytes),
    timestamp: ticks.toString(),
  };
}

export interface ProofSigner {
  signProof(input: SignProofInput): SignedProof;
  proofKeyAttributes(): ProofKeyAttributes;
  oldProofKeyAttributes(): ProofKeyAttributes;
}

function toPair(input: ProofKeyPair | KeyObject): ProofKeyPair {
  return "privateKey" in input ? input : keyPairFromPrivateKey(input);
}

export function createProofSigner(
  current: ProofKeyPair | KeyObject,
  old?: ProofKeyPair | KeyObject,
): ProofSigner {
  const currentPair = toPair(current);
  const oldPair = old ? toPair(old) : undefined;
  return {
    signProof: (input) =>
      signProof(input, {
        privateKey: currentPair.privateKey,
        oldPrivateKey: oldPair?.privateKey,
      }),
    proofKeyAttributes: () => proofKeyAttributes(currentPair.publicKey),
    oldProofKeyAttributes: () => proofKeyAttributes((oldPair ?? currentPair).publicKey),
  };
}
