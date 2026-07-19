import { copyFile, mkdir, rm } from "node:fs/promises";
import { spawnSync } from "node:child_process";
import { fileURLToPath } from "node:url";
import { dirname, resolve } from "node:path";

const WASM_PACK_VERSION = "0.15.0";
const root = resolve(dirname(fileURLToPath(import.meta.url)), "../../..");
const crate = resolve(root, "crates/ooxml-opc");
const output = resolve(root, "target/wasm-pack/redact-worker");
const generated = resolve(root, "apps/redact-worker/src/wasm/generated");

const version = spawnSync("wasm-pack", ["--version"], { encoding: "utf8" });
const versionErrorCode =
  version.error && "code" in version.error && typeof version.error.code === "string"
    ? version.error.code
    : undefined;
if (versionErrorCode === "ENOENT") {
  throw new Error(
    `wasm-pack ${WASM_PACK_VERSION} is required; install it with cargo install wasm-pack --version ${WASM_PACK_VERSION} --locked`,
  );
}
if (version.status !== 0) process.exit(version.status ?? 1);
if (version.stdout.trim() !== `wasm-pack ${WASM_PACK_VERSION}`) {
  throw new Error(`expected wasm-pack ${WASM_PACK_VERSION}, got ${version.stdout.trim()}`);
}

await rm(output, { recursive: true, force: true });
const build = spawnSync(
  "wasm-pack",
  ["build", crate, "--release", "--target", "web", "--out-dir", output, "--locked"],
  { stdio: "inherit" },
);
if (build.status !== 0) process.exit(build.status ?? 1);

await mkdir(generated, { recursive: true });
for (const file of [
  "ooxml_opc.js",
  "ooxml_opc.d.ts",
  "ooxml_opc_bg.wasm",
  "ooxml_opc_bg.wasm.d.ts",
]) {
  await copyFile(resolve(output, file), resolve(generated, file));
}
