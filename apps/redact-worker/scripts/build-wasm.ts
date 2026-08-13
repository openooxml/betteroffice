import { copyFile, mkdir, rm } from "node:fs/promises";
import { spawnSync } from "node:child_process";
import { fileURLToPath } from "node:url";
import { dirname, resolve } from "node:path";
import { requireWasmOpt, requireWasmPack } from "../../../scripts/wasm.ts";

const root = resolve(dirname(fileURLToPath(import.meta.url)), "../../..");
const crate = resolve(root, "crates/ooxml-opc");
const output = resolve(root, "target/wasm-pack/redact-worker");
const generated = resolve(root, "apps/redact-worker/src/wasm/generated");

requireWasmPack();
requireWasmOpt();

await rm(output, { recursive: true, force: true });
const build = spawnSync(
  "wasm-pack",
  [
    "build",
    crate,
    "--release",
    "--target",
    "web",
    "--out-dir",
    output,
    "--",
    "--locked",
    "--features",
    "wasm",
  ],
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
