#!/usr/bin/env bun
// usage: bun scripts/api-surface.ts --write <xlsx|pptx> | bun scripts/api-surface.ts --check <xlsx|pptx>
import { existsSync, mkdirSync, mkdtempSync, readFileSync, readdirSync, rmSync, writeFileSync } from "node:fs";
import { createRequire } from "node:module";
import { tmpdir } from "node:os";
import { dirname, join, relative, resolve } from "node:path";
import { nativeBuildSource } from "./native-engine";

interface FormatConfig {
  packageDir: string;
  packageName: string;
  entries: Record<string, string>;
  /** Declarations of the -native package built from these sources: checked against the wasm package's snapshots. */
  native?: "xlsx" | "pptx";
}

interface DeclarationOptions {
  compilerOptions?: Record<string, unknown>;
  filePaths?: Map<string, string>;
  packageDirAbs?: string;
}

export const FORMATS: Record<string, FormatConfig> = {
  xlsx: {
    packageDir: "packages/xlsx",
    packageName: "@betteroffice/xlsx",
    entries: {
      ".": "src/index.ts",
      "./headless": "src/headless.ts",
      "./worker": "src/worker.ts",
      "./collaboration": "src/collaboration/index.ts",
    },
  },
  pptx: {
    packageDir: "packages/pptx",
    packageName: "@betteroffice/pptx",
    entries: {
      ".": "src/index.ts",
      "./worker": "src/worker.ts",
      "./collaboration": "src/collaboration/index.ts",
    },
  },
};

FORMATS["xlsx-native"] = { ...FORMATS.xlsx, native: "xlsx" };

export function entryFileName(entry: string): string {
  if (entry === ".") return "index.api.d.ts";
  return `${entry.replace(/^\.\//, "").replace(/\//g, "-")}.api.d.ts`;
}

export function reachableFiles(
  entryFile: string,
  files: Map<string, string>,
  options: DeclarationOptions = {},
): string[] {
  const packageDirAbs = options.packageDirAbs ?? join(import.meta.dir, "..", FORMATS.xlsx.packageDir);
  const ts = loadTypeScript(packageDirAbs);
  const compilerOptions = options.compilerOptions ?? {
    moduleResolution: ts.ModuleResolutionKind.Bundler,
    module: ts.ModuleKind.ESNext,
  };
  const paths = new Map<string, string>();
  const directories = new Set<string>();
  for (const file of files.keys()) {
    const absolute = normalizePath(resolve(options.filePaths?.get(file) ?? file));
    paths.set(absolute, file);
    let directory = dirname(absolute);
    while (!directories.has(directory)) {
      directories.add(directory);
      const parent = dirname(directory);
      if (parent === directory) break;
      directory = parent;
    }
  }
  const host = {
    fileExists: (fileName: string) => paths.has(normalizePath(resolve(fileName)))
      || (fileName.endsWith("/package.json") && ts.sys.fileExists(fileName)),
    readFile: (fileName: string) => {
      const file = paths.get(normalizePath(resolve(fileName)));
      return file === undefined
        ? fileName.endsWith("/package.json") ? ts.sys.readFile(fileName) : undefined
        : files.get(file);
    },
    directoryExists: (directory: string) => directories.has(normalizePath(resolve(directory)))
      || ts.sys.directoryExists(directory),
    getCurrentDirectory: () => packageDirAbs,
  };
  const cache = ts.createModuleResolutionCache(packageDirAbs, (fileName: string) => fileName, compilerOptions);
  const seen = new Set<string>();
  const queue = [entryFile];
  while (queue.length > 0) {
    const current = queue.shift() as string;
    if (seen.has(current)) continue;
    const content = files.get(current);
    if (content === undefined) throw new Error(`Missing declaration file: ${current}`);
    seen.add(current);
    const preprocessed = ts.preProcessFile(content, true, true);
    const specifiers = [
      ...preprocessed.importedFiles.map((file: { fileName: string }) => file.fileName),
      ...preprocessed.referencedFiles.map((file: { fileName: string }) =>
        file.fileName.startsWith(".") ? file.fileName : `./${file.fileName}`),
    ];
    for (const specifier of specifiers) {
      const containingFile = resolve(options.filePaths?.get(current) ?? current);
      const resolved = ts.resolveModuleName(specifier, containingFile, compilerOptions, host, cache).resolvedModule;
      const target = resolved && paths.get(normalizePath(resolve(resolved.resolvedFileName)));
      if (target === undefined) {
        if (specifier.startsWith(".")) {
          throw new Error(`Unresolved local declaration import "${specifier}" in ${current}`);
        }
        continue;
      }
      if (!seen.has(target)) queue.push(target);
    }
  }
  return [...seen].sort();
}

function normalizePath(path: string): string {
  return path.replace(/\\/g, "/").replace(/^\.\//, "");
}

export function buildSnapshot(
  packageName: string,
  entry: string,
  entryDeclaration: string,
  files: Map<string, string>,
  options: DeclarationOptions = {},
): string {
  const reachable = reachableFiles(entryDeclaration, files, options);
  const lines: string[] = [
    `// ${packageName} "${entry}" public API surface`,
    `// Regenerate: bun scripts/api-surface.ts --write ${packageName.replace(/^@betteroffice\//, "")}`,
  ];
  for (const file of reachable) {
    const content = (files.get(file) as string)
      .replace(/\/\/# sourceMappingURL=.*$/gm, "")
      .trim();
    lines.push(`// ---- ${file} ----`, content, "");
  }
  return `${lines.join("\n").trim()}\n`;
}

function loadTypeScript(packageDirAbs: string) {
  const requireFromPackage = createRequire(join(packageDirAbs, "package.json"));
  try {
    return requireFromPackage("typescript");
  } catch {
    const requireFromRoot = createRequire(join(import.meta.dir, "..", "package.json"));
    try {
      return requireFromRoot("typescript");
    } catch {
      const requireFromXlsx = createRequire(join(import.meta.dir, "..", FORMATS.xlsx.packageDir, "package.json"));
      return requireFromXlsx("typescript");
    }
  }
}

function emitDeclarations(config: FormatConfig, root: string) {
  const packageDirAbs = resolve(root, config.packageDir);
  const ts = loadTypeScript(packageDirAbs);
  const entryFiles = Object.values(config.entries).map((entry) => join(packageDirAbs, entry));
  const configPath = ts.findConfigFile(packageDirAbs, ts.sys.fileExists, "tsconfig.json");
  let packageOptions = {};
  if (configPath) {
    const loaded = ts.readConfigFile(configPath, ts.sys.readFile);
    if (loaded.error) {
      throw new Error(ts.flattenDiagnosticMessageText(loaded.error.messageText, "\n"));
    }
    const parsed = ts.parseJsonConfigFileContent(loaded.config, ts.sys, dirname(configPath));
    if (parsed.errors.length > 0) {
      throw new Error(parsed.errors
        .map((error: { messageText: unknown }) => ts.flattenDiagnosticMessageText(error.messageText, "\n"))
        .join("\n"));
    }
    packageOptions = parsed.options;
  }
  const options = {
    target: ts.ScriptTarget.ES2020,
    module: ts.ModuleKind.ESNext,
    moduleResolution: ts.ModuleResolutionKind.Bundler,
    lib: ["lib.es2020.d.ts", "lib.dom.d.ts", "lib.dom.iterable.d.ts"],
    ...packageOptions,
    declaration: true,
    emitDeclarationOnly: true,
    declarationMap: false,
    sourceMap: false,
    skipLibCheck: true,
    strict: true,
    allowImportingTsExtensions: true,
    resolveJsonModule: true,
    isolatedModules: true,
    noEmitOnError: false,
    noEmit: false,
    outDir: undefined,
    declarationDir: undefined,
  };
  const host = ts.createCompilerHost(options);
  if (config.native) {
    const native = config.native;
    const getSourceFile = host.getSourceFile;
    host.getSourceFile = (fileName: string, languageVersion: unknown, ...rest: unknown[]) => {
      const replaced = nativeBuildSource(native, fileName);
      return replaced === undefined
        ? getSourceFile.call(host, fileName, languageVersion, ...rest)
        : ts.createSourceFile(fileName, replaced, languageVersion, true);
    };
  }
  const program = ts.createProgram(entryFiles, options, host);
  const files = new Map<string, string>();
  const filePaths = new Map<string, string>();
  const addFile = (fileName: string, text: string) => {
    const absolute = resolve(fileName);
    const sourceRelative = normalizePath(relative(join(packageDirAbs, "src"), absolute));
    const key = sourceRelative.startsWith("../")
      ? normalizePath(relative(root, absolute))
      : sourceRelative;
    files.set(key, text);
    filePaths.set(key, absolute);
  };
  program.emit(undefined, addFile);
  for (const source of program.getSourceFiles()) {
    if (source.isDeclarationFile
      && !program.isSourceFileDefaultLibrary(source)
      && !program.isSourceFileFromExternalLibrary(source)
      && !normalizePath(source.fileName).includes("/node_modules/")) {
      addFile(source.fileName, source.text);
    }
  }
  return { files, options: { compilerOptions: options, filePaths, packageDirAbs } };
}

function entryDeclarationFor(config: FormatConfig, entry: string, sourceEntry: string): string {
  void config;
  void entry;
  let name = sourceEntry.replace(/^src\//, "").replace(/\.ts$/, ".d.ts");
  return normalizePath(name);
}

export function generateSnapshots(format: string): Map<string, string> {
  return generateSnapshotsForPackage(FORMATS[format], join(import.meta.dir, ".."));
}

export function generateSnapshotsForPackage(config: FormatConfig, root: string): Map<string, string> {
  const { files, options } = emitDeclarations(config, root);
  const snapshots = new Map<string, string>();
  for (const [entry, sourceEntry] of Object.entries(config.entries)) {
    const declaration = entryDeclarationFor(config, entry, sourceEntry);
    snapshots.set(entryFileName(entry), buildSnapshot(config.packageName, entry, declaration, files, options));
  }
  return snapshots;
}

function diffText(expected: string, actual: string): string {
  const expectedLines = expected.split("\n");
  const actualLines = actual.split("\n");
  const out: string[] = [];
  const max = Math.max(expectedLines.length, actualLines.length);
  for (let index = 0; index < max; index += 1) {
    if (expectedLines[index] !== actualLines[index]) {
      if (expectedLines[index] !== undefined) out.push(`- ${expectedLines[index]}`);
      if (actualLines[index] !== undefined) out.push(`+ ${actualLines[index]}`);
    }
  }
  return out.slice(0, 200).join("\n");
}

function main(): void {
  const args = process.argv.slice(2);
  const mode = args.find((arg) => arg === "--write" || arg === "--check");
  const format = args.find((arg) => !arg.startsWith("--"));
  if (!mode || !format || !FORMATS[format]) {
    console.error("usage: bun scripts/api-surface.ts --write xlsx | bun scripts/api-surface.ts --check xlsx");
    process.exit(2);
  }
  const config = FORMATS[format];
  if (config.native && mode === "--write") {
    console.error(`${format} declarations are checked against the ${config.packageName} snapshots; write those instead`);
    process.exit(2);
  }
  const root = join(import.meta.dir, "..");
  const apiDir = join(root, config.packageDir, "api");
  const snapshots = generateSnapshots(format);
  if (mode === "--write") {
    mkdirSync(apiDir, { recursive: true });
    for (const existing of existsSync(apiDir) ? readdirSync(apiDir) : []) {
      if (existing.endsWith(".api.d.ts") && !snapshots.has(existing)) {
        rmSync(join(apiDir, existing));
      }
    }
    for (const [fileName, content] of snapshots) {
      writeFileSync(join(apiDir, fileName), content);
    }
    console.log(`wrote ${snapshots.size} API surface snapshot(s) for ${format} to ${relative(root, apiDir)}`);
    return;
  }
  const tempDir = mkdtempSync(join(tmpdir(), "api-surface-"));
  try {
    let failed = false;
    for (const [fileName, content] of snapshots) {
      const expectedPath = join(apiDir, fileName);
      const tempPath = join(tempDir, fileName);
      writeFileSync(tempPath, content);
      if (!existsSync(expectedPath)) {
        console.error(`missing API surface snapshot: ${relative(root, expectedPath)}`);
        console.error("regenerate with: bun scripts/api-surface.ts --write xlsx");
        failed = true;
        continue;
      }
      const committed = readFileSync(expectedPath, "utf8");
      if (committed !== content) {
        console.error(`API surface changed: ${relative(root, expectedPath)}`);
        console.error(diffText(committed, content));
        console.error("regenerate with: bun scripts/api-surface.ts --write xlsx");
        failed = true;
      }
    }
    if (existsSync(apiDir)) {
      for (const existing of readdirSync(apiDir)) {
        if (existing.endsWith(".api.d.ts") && !snapshots.has(existing)) {
          console.error(`stale API surface snapshot: ${join(relative(root, apiDir), existing)}`);
          failed = true;
        }
      }
    }
    if (failed) process.exit(1);
    console.log(`API surface ok for ${format}`);
  } finally {
    rmSync(tempDir, { recursive: true, force: true });
  }
}

if (import.meta.main) main();
