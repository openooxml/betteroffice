import wasmModule from "./wasm/generated/ooxml_opc_bg.wasm";
import { initSync, sanitizeOoxml } from "./wasm/generated/ooxml_opc.js";
import { createWorker } from "./handler";

initSync({ module: wasmModule });

export default createWorker((bytes, format) => sanitizeOoxml(bytes, format));
