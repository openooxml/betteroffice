import { describe, expect, test } from "bun:test";
import { compareSurfaces, normalizeType, parseParams, parseSurface } from "./check-engine-surface";

const wasm = `
export class XlsxDocument {
    private constructor();
    free(): void;
    [Symbol.dispose](): void;
    static open(bytes: Uint8Array): XlsxDocument;
    editCellJson(args: string): string;
    saveBytes(): Uint8Array;
    nextPeerSnapshotChunk(): Uint8Array | undefined;
    recoveryCheckpoint(through_seq: number, preserve_history?: boolean | null): Uint8Array;
    readonly clientId: number;
}
export class XlsxSnapshotBuilder {
    free(): void;
    constructor();
    push(chunk: Uint8Array): void;
    finish(): XlsxDocument;
}
export function displayListCellLimit(): number;
export function install_panic_hook(): void;
export type InitInput = string;
export interface InitOutput {
  readonly memory: WebAssembly.Memory;
}
`;

const native = `
export declare class XlsxDocument {
  free(): void
  static open(bytes: Uint8Array): XlsxDocument
  editCellJson(args: string): string
  saveBytes(): Buffer
  nextPeerSnapshotChunk(): Buffer | undefined
  recoveryCheckpoint(throughSeq: number, preserveHistory?: boolean | undefined | null): Buffer
  get clientId(): number
}
export declare class XlsxSnapshotBuilder {
  constructor()
  push(chunk: Uint8Array): void
  finish(): XlsxDocument
  free(): void
}
export declare function displayListCellLimit(): number
export declare function engineVersion(): string
export interface BorderPatch {
  preset?: string
}
`;

describe("check-engine-surface parser", () => {
  test("normalizes Buffer and null unions", () => {
    expect(normalizeType("Buffer")).toBe("Uint8Array");
    expect(normalizeType("string | null")).toBe("string");
    expect(normalizeType("Buffer | undefined | null")).toBe("Uint8Array | undefined");
  });

  test("parses params ignoring names and optional undefined", () => {
    expect(parseParams("through_seq: number, preserve_history?: boolean | null")).toEqual([
      { type: "number", optional: false },
      { type: "boolean", optional: true },
    ]);
    expect(parseParams("bytes: Uint8Array")).toEqual([{ type: "Uint8Array", optional: false }]);
  });

  test("matching wasm and native surfaces pass", () => {
    const problems = compareSurfaces(parseSurface(wasm), parseSurface(native), [
      "XlsxDocument",
      "XlsxSnapshotBuilder",
    ]);
    expect(problems).toEqual([]);
  });

  test("missing and mismatched members are reported", () => {
    const broken = native.replace("saveBytes(): Buffer", "saveBytes(): string");
    const problems = compareSurfaces(parseSurface(wasm), parseSurface(broken), [
      "XlsxDocument",
      "XlsxSnapshotBuilder",
    ]);
    expect(problems.some((problem) => problem.includes("saveBytes"))).toBe(true);
  });

  test("top-level functions after a class are compared", () => {
    expect(parseSurface(wasm).functions.has("displayListCellLimit")).toBe(true);
    const broken = native.replace("displayListCellLimit(): number", "displayListCellLimit(): string");
    const problems = compareSurfaces(parseSurface(wasm), parseSurface(broken), [
      "XlsxDocument",
      "XlsxSnapshotBuilder",
    ]);
    expect(problems.some((problem) => problem.includes("displayListCellLimit"))).toBe(true);
  });

  test("a missing free() is reported", () => {
    const broken = native.replace("  free(): void\n  static open", "  static open");
    const problems = compareSurfaces(parseSurface(wasm), parseSurface(broken), [
      "XlsxDocument",
      "XlsxSnapshotBuilder",
    ]);
    expect(problems.some((problem) => problem.includes("free"))).toBe(true);
  });

  test.each([
    ["method", "export class Example { read(value?: boolean): void; }", "export class Example { read(value: boolean): void; }"],
    ["static method", "export class Example { static open(value?: boolean): Example; }", "export class Example { static open(value: boolean): Example; }"],
    ["function", "export function read(value?: boolean): void;", "export function read(value: boolean): void;"],
  ])("a required parameter replacing an optional one is reported for a %s", (_, expected, actual) => {
    const classes = expected.includes("class Example") ? ["Example"] : [];
    const problems = compareSurfaces(parseSurface(expected), parseSurface(actual), classes);
    expect(problems).toHaveLength(1);
    expect(problems[0]).toContain("parameter optionality mismatch");
  });

  test("parses public constructors and ignores private constructors", () => {
    expect(parseSurface(wasm).classes.get("XlsxDocument")?.has("constructor")).toBe(false);
    expect(parseSurface(wasm).classes.get("XlsxSnapshotBuilder")?.get("constructor")).toEqual({
      kind: "constructor", name: "constructor", params: [], returnType: "void",
    });
    const surface = parseSurface(`export class Example {
      public constructor(
        bytes: Uint8Array,
        enabled?: boolean
      );
    }`);
    expect(surface.classes.get("Example")?.get("constructor")?.params).toEqual([
      { type: "Uint8Array", optional: false },
      { type: "boolean", optional: true },
    ]);
  });

  test.each([
    ["arity", "constructor(value?: boolean, enabled?: boolean)"],
    ["parameter types", "constructor(value?: number)"],
    ["parameter optionality", "constructor(value: boolean)"],
  ])("a changed public constructor reports a %s mismatch", (kind, constructor) => {
    const expected = parseSurface("export class Example { constructor(value?: boolean); }");
    const actual = parseSurface(`export class Example { ${constructor}; }`);
    const problems = compareSurfaces(expected, actual, ["Example"]);
    expect(problems).toHaveLength(1);
    expect(problems[0]).toContain(`Example.constructor: ${kind} mismatch`);
  });

  test("a missing or added public constructor is reported", () => {
    const constructible = parseSurface("export class Example { constructor(); }");
    const privateConstructor = parseSurface("export class Example { private constructor(); }");
    expect(compareSurfaces(constructible, privateConstructor, ["Example"])[0])
      .toContain("Example.constructor: missing natively");
    expect(compareSurfaces(privateConstructor, constructible, ["Example"])[0])
      .toContain("Example.constructor: public constructor is only available natively");
  });

  test("pending members are skipped and unlisted missing members still fail", () => {
    const expected = parseSurface(`export class Example {
      later(): string;
      required(): number;
    }`);
    const actual = parseSurface("export class Example {}");
    const problems = compareSurfaces(expected, actual, ["Example"], ["Example.later"]);
    expect(problems).toHaveLength(1);
    expect(problems[0]).toContain("Example.required: missing natively");
    expect(compareSurfaces(expected, actual, ["Example"], [
      "Example.later", "Example.required",
    ])).toEqual([]);
  });

  test("pending classes and functions are skipped independently", () => {
    const expected = parseSurface(`
      export class Later { read(): string; }
      export function later(): string;
      export function required(): number;
    `);
    const problems = compareSurfaces(expected, parseSurface(""), ["Later"], [
      "Later", "function later",
    ]);
    expect(problems).toHaveLength(1);
    expect(problems[0]).toContain("function required: missing natively");
  });

  test("stale pending entries fail", () => {
    const expected = parseSurface(`export class Example { later(): string; }
      export function done(): number;`);
    const actual = parseSurface(`export class Example { later(): string; }
      export function done(): number;`);
    const problems = compareSurfaces(expected, actual, ["Example"], [
      "Example.later", "function done", "Example.gone",
    ]);
    expect(problems).toHaveLength(3);
    expect(problems.join("\n")).toContain("Example.later: pending but available natively");
    expect(problems.join("\n")).toContain("function done: pending but available natively");
    expect(problems.join("\n")).toContain("Example.gone: pending but not in the wasm surface");
  });

  test("native-only classes and functions are allowed", () => {
    const extra = `${native}
      export declare class Document { read(): string; }
      export declare function openDocument(): Document;
    `;
    expect(compareSurfaces(parseSurface(wasm), parseSurface(extra), [
      "XlsxDocument", "XlsxSnapshotBuilder",
    ])).toEqual([]);
  });
});
