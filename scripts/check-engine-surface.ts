#!/usr/bin/env bun
import { readFileSync } from "node:fs";
import { join } from "node:path";

export interface ParamInfo {
  type: string;
  optional: boolean;
}

export interface MemberInfo {
  kind: "static" | "getter" | "method" | "constructor";
  name: string;
  params: ParamInfo[];
  returnType: string;
}

export interface Surface {
  classes: Map<string, Map<string, MemberInfo>>;
  functions: Map<string, MemberInfo>;
}

interface FormatConfig {
  wasm: string;
  native: string;
  classes: string[];
}

export const FORMATS: Record<string, FormatConfig> = {
  xlsx: {
    wasm: "packages/xlsx/src/wasm/generated/xlsx_wasm.d.ts",
    native: "bindings/xlsx-native/index.d.ts",
    classes: ["XlsxDocument", "XlsxSnapshotBuilder"],
  },
};

const IGNORED_FUNCTIONS = new Set([
  "free",
  "install_panic_hook",
  "initSync",
  "default",
  "__wbg_init",
]);

function splitTopLevel(input: string, separator: string): string[] {
  const parts: string[] = [];
  let depth = 0;
  let angle = 0;
  let current = "";
  for (const ch of input) {
    if (ch === "(" || ch === "[" || ch === "{") depth += 1;
    if (ch === ")" || ch === "]" || ch === "}") depth -= 1;
    if (ch === "<") angle += 1;
    if (ch === ">" && angle > 0) angle -= 1;
    if (ch === separator && depth === 0 && angle === 0) {
      parts.push(current);
      current = "";
    } else {
      current += ch;
    }
  }
  if (current.trim() !== "" || parts.length > 0) parts.push(current);
  return parts;
}

export function normalizeType(raw: string): string {
  let type = raw.trim().replace(/\s+/g, " ");
  type = type.replace(/\bBuffer\b/g, "Uint8Array");
  if (type.includes("|")) {
    const parts = splitTopLevel(type, "|")
      .map((part) => part.trim())
      .filter((part) => part !== "null" && part !== "");
    const deduped: string[] = [];
    for (const part of parts) {
      if (!deduped.includes(part)) deduped.push(part);
    }
    return deduped.join(" | ");
  }
  return type;
}

function normalizeParamType(raw: string, optional: boolean): string {
  let type = normalizeType(raw);
  if (optional && type.includes("|")) {
    const parts = splitTopLevel(type, "|")
      .map((part) => part.trim())
      .filter((part) => part !== "undefined" && part !== "null");
    type = parts.join(" | ");
  }
  return type;
}

export function parseParams(raw: string): ParamInfo[] {
  const trimmed = raw.trim();
  if (trimmed === "") return [];
  return splitTopLevel(trimmed, ",").map((part) => {
    const item = part.trim();
    const colon = item.indexOf(":");
    if (colon === -1) return { type: normalizeType(item), optional: false };
    const namePart = item.slice(0, colon).trim();
    const optional = namePart.endsWith("?");
    const typeRaw = item.slice(colon + 1).trim();
    return { type: normalizeParamType(typeRaw, optional), optional };
  });
}

function findClassBodies(text: string): Map<string, string> {
  const bodies = new Map<string, string>();
  const classRe = /(?:export\s+)?(?:declare\s+)?class\s+(\w+)\s*\{/g;
  let match: RegExpExecArray | null;
  while ((match = classRe.exec(text)) !== null) {
    const name = match[1];
    let depth = 1;
    let index = classRe.lastIndex;
    const start = index;
    while (index < text.length && depth > 0) {
      if (text[index] === "{") depth += 1;
      if (text[index] === "}") depth -= 1;
      index += 1;
    }
    bodies.set(name, text.slice(start, index - 1));
    classRe.lastIndex = index;
  }
  return bodies;
}

function parseMember(declaration: string): MemberInfo | null {
  let item = declaration.trim().replace(/\s+/g, " ");
  if (item === "") return null;
  if (item.startsWith("[Symbol.dispose]")) return null;
  if (/^(private|protected)\s+constructor\b/.test(item)) return null;
  const constructor = /^(?:public\s+)?constructor\s*\((.*)\)$/.exec(item);
  if (constructor) {
    return {
      kind: "constructor",
      name: "constructor",
      params: parseParams(constructor[1]),
      returnType: "void",
    };
  }
  let kind: MemberInfo["kind"] = "method";
  if (item.startsWith("static ")) {
    kind = "static";
    item = item.slice("static ".length).trim();
  }
  if (item.startsWith("get ")) {
    const getter = /^get\s+(\w+)\s*\(\s*\)\s*:\s*(.+)$/.exec(item);
    if (!getter) return null;
    return { kind: "getter", name: getter[1], params: [], returnType: normalizeType(getter[2]) };
  }
  if (item.startsWith("readonly ")) {
    const readonlyMatch = /^readonly\s+(\w+)\s*:\s*(.+)$/.exec(item);
    if (!readonlyMatch) return null;
    return {
      kind: "getter",
      name: readonlyMatch[1],
      params: [],
      returnType: normalizeType(readonlyMatch[2]),
    };
  }
  const method = /^(\w+)\s*\((.*)\)\s*(?::\s*(.+))?$/.exec(item);
  if (!method) return null;
  return {
    kind,
    name: method[1],
    params: parseParams(method[2]),
    returnType: normalizeType(method[3] ?? "void"),
  };
}

export function parseSurface(text: string): Surface {
  const withoutComments = text.replace(/\/\*[\s\S]*?\*\//g, " ").replace(/\/\/.*$/gm, " ");
  const classes = new Map<string, Map<string, MemberInfo>>();
  for (const [className, body] of findClassBodies(withoutComments)) {
    const members = new Map<string, MemberInfo>();
    for (const rawDecl of splitTopLevel(body, ";").flatMap((part) => splitTopLevel(part, "\n"))) {
      const member = parseMember(rawDecl);
      if (member) members.set(member.name, member);
    }
    classes.set(className, members);
  }
  let topLevel = withoutComments;
  for (const [, body] of findClassBodies(withoutComments)) {
    topLevel = topLevel.replace(body, " ");
  }
  topLevel = topLevel.replace(/(?:export\s+)?(?:declare\s+)?class\s+\w+\s*\{\s*\}/g, " ");
  const functions = new Map<string, MemberInfo>();
  const fnRe = /export\s+(?:declare\s+)?function\s+(\w+)\s*\(([^)]*)\)\s*:\s*([^;\n]+)[;\n]/g;
  let fnMatch: RegExpExecArray | null;
  while ((fnMatch = fnRe.exec(topLevel)) !== null) {
    const name = fnMatch[1];
    if (IGNORED_FUNCTIONS.has(name) || name.startsWith("__wbg")) continue;
    functions.set(name, {
      kind: "method",
      name,
      params: parseParams(fnMatch[2]),
      returnType: normalizeType(fnMatch[3]),
    });
  }
  const defaultRe = /export\s+default\s+function\s+(\w+)/;
  if (defaultRe.test(topLevel)) {
    // default init is ignored by contract
  }
  return { classes, functions };
}

function describe(member: MemberInfo): string {
  const prefix = member.kind === "static" ? "static " : member.kind === "getter" ? "get " : "";
  const params = member.params
    .map((param) => param.optional ? `${param.type} (optional)` : param.type)
    .join(", ");
  return `${prefix}${member.name}(${params}): ${member.returnType}`;
}

export function compareSurfaces(wasm: Surface, native: Surface, classNames: string[]): string[] {
  const problems: string[] = [];
  for (const className of classNames) {
    const wasmClass = wasm.classes.get(className);
    const nativeClass = native.classes.get(className);
    if (!wasmClass) {
      problems.push(`wasm class ${className} not found`);
      continue;
    }
    if (!nativeClass) {
      problems.push(`native class ${className} is missing`);
      continue;
    }
    if (!wasmClass.has("constructor") && nativeClass.has("constructor")) {
      problems.push(`${className}.constructor: public constructor is only available natively`);
    }
    for (const [name, expected] of wasmClass) {
      const actual = nativeClass.get(name);
      if (!actual) {
        problems.push(`${className}.${name}: missing natively (wasm: ${describe(expected)})`);
        continue;
      }
      if (actual.kind !== expected.kind) {
        problems.push(
          `${className}.${name}: kind mismatch wasm ${expected.kind}, native ${actual.kind} (${describe(expected)} vs ${describe(actual)})`,
        );
        continue;
      }
      if (actual.params.length !== expected.params.length) {
        problems.push(
          `${className}.${name}: arity mismatch wasm ${expected.params.length}, native ${actual.params.length} (${describe(expected)} vs ${describe(actual)})`,
        );
        continue;
      }
      if (expected.params.some((param, index) => param.optional !== actual.params[index].optional)) {
        problems.push(
          `${className}.${name}: parameter optionality mismatch (${describe(expected)} vs ${describe(actual)})`,
        );
        continue;
      }
      const paramsMismatch = expected.params.some(
        (param, index) => param.type !== actual.params[index].type,
      );
      if (paramsMismatch) {
        problems.push(
          `${className}.${name}: parameter types mismatch (${describe(expected)} vs ${describe(actual)})`,
        );
        continue;
      }
      if (actual.returnType !== expected.returnType) {
        problems.push(
          `${className}.${name}: return type mismatch (${describe(expected)} vs ${describe(actual)})`,
        );
      }
    }
  }
  for (const [name, expected] of wasm.functions) {
    if (IGNORED_FUNCTIONS.has(name)) continue;
    const actual = native.functions.get(name);
    if (!actual) {
      problems.push(`function ${name}: missing natively (wasm: ${describe(expected)})`);
      continue;
    }
    if (actual.params.length !== expected.params.length) {
      problems.push(
        `function ${name}: arity mismatch wasm ${expected.params.length}, native ${actual.params.length} (${describe(expected)} vs ${describe(actual)})`,
      );
      continue;
    }
    if (expected.params.some((param, index) => param.optional !== actual.params[index].optional)) {
      problems.push(
        `function ${name}: parameter optionality mismatch (${describe(expected)} vs ${describe(actual)})`,
      );
      continue;
    }
    const paramsMismatch = expected.params.some(
      (param, index) => param.type !== actual.params[index].type,
    );
    if (paramsMismatch) {
      problems.push(
        `function ${name}: parameter types mismatch (${describe(expected)} vs ${describe(actual)})`,
      );
      continue;
    }
    if (actual.returnType !== expected.returnType) {
      problems.push(
        `function ${name}: return type mismatch (${describe(expected)} vs ${describe(actual)})`,
      );
    }
  }
  return problems;
}

function main(): void {
  const format = process.argv[2];
  if (!format || !FORMATS[format]) {
    console.error(`usage: bun scripts/check-engine-surface.ts <${Object.keys(FORMATS).join("|")}>`);
    process.exit(2);
  }
  const config = FORMATS[format];
  const root = join(import.meta.dir, "..");
  const wasmText = readFileSync(join(root, config.wasm), "utf8");
  const nativeText = readFileSync(join(root, config.native), "utf8");
  const problems = compareSurfaces(parseSurface(wasmText), parseSurface(nativeText), config.classes);
  if (problems.length > 0) {
    console.error(`engine surface mismatch for ${format}:`);
    for (const problem of problems) console.error(`  - ${problem}`);
    process.exit(1);
  }
  console.log(`engine surface ok for ${format}`);
}

if (import.meta.main) main();
