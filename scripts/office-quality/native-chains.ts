import { createHash } from 'node:crypto';
import { appendFile, mkdir, readFile, writeFile } from 'node:fs/promises';
import { createRequire } from 'node:module';
import { dirname, join, resolve } from 'node:path';
import { pathToFileURL } from 'node:url';

type Module = Record<string, any>;

export interface Engine {
  registerFont(bytes: Uint8Array, face?: { family: string; bold: boolean; italic: boolean }): number;
  registerSubstituteFont?(id: number, family: string, bold?: boolean, italic?: boolean): number;
  clearFonts(): void;
}

export type Substitute = { substituteOf: number; family: string; bold?: boolean; italic?: boolean };
export type RecordedFace = { bytes: Uint8Array } | Substitute;

export type RejectedEmbedded = { family: string; reason: string };

export interface ChainsFile {
  faces: Array<{ file: string; sha256: string } | Substitute>;
  fontChains: Record<string, number[]>;
  measurement: Record<string, unknown>;
  rejectedEmbedded: RejectedEmbedded[];
  absentAddons: AbsentAddon[];
}

export type Measurement = { fontChains: Record<string, number[]> } & Record<string, unknown>;

export interface Requirement {
  key: string;
  family: string;
  bold: boolean;
  italic: boolean;
  scripts?: string[];
}

export const ADDONS = { 'fonts-latin': 'LATIN_FONT_ASSET_URLS', 'fonts-intl': 'INTL_FONT_ASSET_URLS', 'fonts-cjk': 'CJK_FONT_ASSET_URLS' } as const;
export type AddonName = keyof typeof ADDONS;

export type AbsentAddon = { addon: AddonName; families: string[]; scripts: string[] };

export interface Addon {
  version?: string;
  error?: string;
  absent?: boolean;
  files: Set<string>;
  manifest: string;
}

const ADDON_IN_ERROR = /@betteroffice\/(fonts-latin|fonts-intl|fonts-cjk)\b|\b(intl|CJK) add-on/i;

export function addonOfError(text: string): AddonName | undefined {
  const match = ADDON_IN_ERROR.exec(text);
  return match ? ((match[1] ?? `fonts-${match[2].toLowerCase()}`) as AddonName) : undefined;
}

export interface FontPackages {
  version: string;
  fonts: Module;
  addons: Record<AddonName, Addon>;
}

export interface Product {
  root: string;
  version: string;
  yrs: Module;
  layout: Module;
  utils: Module;
  editor: Module;
  fonts: FontPackages;
}

export class MissingExport extends Error {}

export class MissingAddon extends Error {}

export class FailedRegistration extends Error {}

const repo = resolve(import.meta.dir, '../..');
const MEASUREMENT_KEYS = new Set(['defaults', 'compat', 'authoritativeShaping']);

const REQUIRED = {
  yrs: ['createYrsSession'],
  layout: ['createRustMeasureSource', 'resolveCompleteFontProvider'],
  utils: ['extractEmbeddedFontFaces'],
  editor: ['buildResidentRegionLayoutRequest', 'documentRenderEnv'],
} as const;
const SESSION = ['openDocx', 'materializeDocx', 'layoutFontRequirementsJson', 'registerFont', 'clearFonts', 'destroy'];

export const sha256 = (bytes: Uint8Array) => createHash('sha256').update(bytes).digest('hex');

export async function packageVersion(root: string): Promise<string> {
  const version = JSON.parse(await readFile(join(root, 'package.json'), 'utf8')).version;
  if (typeof version !== 'string' || !version) throw new Error(`${root}/package.json has no version`);
  return version;
}

export async function loadProduct(root: string, planAbsent: ReadonlySet<AddonName> = new Set()): Promise<Product> {
  root = resolve(root);
  const manifest = JSON.parse(await readFile(join(root, 'package.json'), 'utf8'));
  const label = `${manifest.name}@${manifest.version}`;
  const modules: Record<string, Module> = {};
  for (const [subpath, names] of Object.entries(REQUIRED)) {
    const entry = manifest.exports?.[`./${subpath}`];
    const target = typeof entry === 'string' ? entry : entry?.import ?? entry?.default;
    if (typeof target !== 'string') throw new MissingExport(`${label} has no ./${subpath} export`);
    const module = await import(pathToFileURL(join(root, target)).href);
    const missing = names.filter((name) => typeof module[name] !== 'function');
    if (missing.length) throw new MissingExport(`${label} ./${subpath} lacks ${missing.join(', ')}`);
    modules[subpath] = module;
  }
  const probe = await modules.yrs.createYrsSession();
  try {
    const missing = SESSION.filter((name) => typeof probe[name] !== 'function');
    if (missing.length) throw new MissingExport(`${label} sessions lack ${missing.join(', ')}`);
  } finally {
    probe.destroy?.();
  }
  const fonts = await loadFontPackages(root, planAbsent);
  return { root, version: manifest.version, yrs: modules.yrs, layout: modules.layout, utils: modules.utils, editor: modules.editor, fonts };
}

const message = (error: unknown) => (error instanceof Error ? error.message : String(error));

export async function loadFontPackages(root: string, planAbsent: ReadonlySet<AddonName> = new Set()): Promise<FontPackages> {
  const fromProduct = createRequire(join(root, 'package.json'));
  const fontsManifest = fromProduct.resolve('@betteroffice/fonts/package.json');
  const fonts = await import(pathToFileURL(fromProduct.resolve('@betteroffice/fonts')).href);
  const version = await packageVersion(dirname(fontsManifest));
  const missing = ['resolveBundledFamilyFace', 'resolveScriptFallbackFace'].filter((name) => typeof fonts[name] !== 'function');
  if (missing.length) throw new Error(`@betteroffice/fonts@${version} lacks ${missing.join(', ')}`);
  const fromFonts = createRequire(fontsManifest);
  const fontsPackage = JSON.parse(await readFile(fontsManifest, 'utf8'));
  const addons = {} as Record<AddonName, Addon>;
  for (const [name, exported] of Object.entries(ADDONS) as Array<[AddonName, string]>) {
    const id = `@betteroffice/${name}`;
    let addonRoot: string | undefined;
    try {
      addonRoot = dirname(fromFonts.resolve(`${id}/package.json`));
    } catch {}
    const addon: Addon = { files: new Set(), manifest: join(addonRoot ?? join(repo, 'packages', name), 'assets', 'manifest.json') };
    const declared = ['dependencies', 'peerDependencies', 'optionalDependencies'].some((field) => fontsPackage[field]?.[id] !== undefined);
    if (planAbsent.has(name) && declared) throw new Error(`the plan marks ${id} absent, but @betteroffice/fonts@${version} declares it`);
    if (!addonRoot && !declared) addon.absent = true;
    if (addonRoot) {
      addon.version = await packageVersion(addonRoot);
      try {
        const module = await import(pathToFileURL(fromFonts.resolve(id)).href);
        if (module[exported] === null || typeof module[exported] !== 'object') addon.error = `${id}@${addon.version} has no ${exported} export`;
      } catch (error) {
        addon.error = `${id}@${addon.version} failed to load: ${message(error)}`;
      }
    }
    const faces = JSON.parse(await readFile(addon.manifest, 'utf8')).faces;
    if (!Array.isArray(faces)) throw new Error(`${addon.manifest} lists no faces`);
    for (const face of faces) addon.files.add(face.file);
    addons[name] = addon;
  }
  return { version, fonts, addons };
}

export function describeAddons(packages: FontPackages) {
  return Object.entries(packages.addons).map(([name, addon]) =>
    `${name} ${addon.absent ? 'absent (not declared)' : addon.version === undefined ? 'missing' : addon.error ? `${addon.version} (failed: ${addon.error})` : addon.version}`).join(', ');
}

export function absentAddonUse(packages: FontPackages, requirements: Requirement[], loads: Iterable<[AddonName, string]> = []): AbsentAddon[] {
  const absent = (Object.entries(packages.addons) as Array<[AddonName, Addon]>).filter(([, addon]) => addon.absent);
  const owner = (face: { file: string } | undefined) => face && absent.find(([, addon]) => addon.files.has(face.file))?.[0];
  const use = new Map<AddonName, { families: Set<string>; scripts: Set<string> }>();
  const entry = (name: AddonName) => use.get(name) ?? use.set(name, { families: new Set(), scripts: new Set() }).get(name)!;
  for (const requirement of requirements) {
    const family = owner(packages.fonts.resolveBundledFamilyFace(requirement.family, requirement.bold, requirement.italic));
    if (family) entry(family).families.add(requirement.family);
    for (const script of requirement.scripts ?? []) {
      const name = owner(packages.fonts.resolveScriptFallbackFace(script, false, false, requirement.family));
      if (name) entry(name).scripts.add(script);
    }
  }
  for (const [name, label] of loads) entry(name).families.add(label);
  return [...use].map(([addon, { families, scripts }]) => ({ addon, families: [...families].sort(), scripts: [...scripts].sort() }));
}

export function missingAddons(packages: FontPackages, requirements: Requirement[]) {
  const owner = (file: string) => (Object.entries(packages.addons) as Array<[AddonName, Addon]>).find(([, addon]) => addon.files.has(file));
  const missing = new Set<string>();
  for (const requirement of requirements) {
    const faces: Array<[string, { file: string } | undefined]> = [
      ['family face', packages.fonts.resolveBundledFamilyFace(requirement.family, requirement.bold, requirement.italic)],
      ...(requirement.scripts ?? []).map((script): [string, { file: string } | undefined] =>
        [`${script} script`, packages.fonts.resolveScriptFallbackFace(script, false, false, requirement.family)]),
    ];
    for (const [use, face] of faces) {
      const hit = face && owner(face.file);
      if (!hit || hit[1].absent || (hit[1].version !== undefined && !hit[1].error)) continue;
      missing.add(`${requirement.family} (${use}, ${face!.file}) needs @betteroffice/${hit[0]}, which is ${hit[1].error ?? 'missing'}`);
    }
  }
  return [...missing];
}

export function recordingEngine(engine: Engine): { engine: Engine; faces: RecordedFace[] } {
  const faces: RecordedFace[] = [];
  const record = (id: number, face: RecordedFace) => {
    if (id === faces.length) faces.push(face);
    else if (!Number.isInteger(id) || id < 0 || id > faces.length)
      throw new Error(`the engine returned font id ${id} after ${faces.length} registrations`);
    return id;
  };
  return {
    faces,
    engine: {
      registerFont: (bytes, face) =>
        record(face ? engine.registerFont(bytes, face) : engine.registerFont(bytes), { bytes: bytes.slice() }),
      ...(engine.registerSubstituteFont && {
        registerSubstituteFont: (id: number, family: string, ...style: [bold?: boolean, italic?: boolean]) =>
          record(engine.registerSubstituteFont!(id, family, ...style), {
            substituteOf: id,
            family,
            ...(style.length > 0 && { bold: style[0] }),
            ...(style.length > 1 && { italic: style[1] }),
          }),
      }),
      clearFonts: () => {
        engine.clearFonts();
        faces.length = 0;
      },
    },
  };
}

const RESOLVERS: Record<string, (args: any[]) => string> = {
  resolve: ([family]) => family,
  resolveLastResort: ([family]) => `${family} (last resort)`,
  resolveGlyphFallback: ([family]) => `${family} (coverage)`,
  resolveScriptFallback: ([script, , , family]) => `${family ?? 'no family'} (${script} script)`,
  resolveBaseLatin: () => 'base Latin',
};

export function guardRegistrations(bundled: Module | undefined, engine: Engine, embedded: Array<{ family: string; data: ArrayBuffer }> = [], absent: ReadonlySet<AddonName> = new Set()) {
  const failures = new Set<string>();
  const absentLoads: Array<[AddonName, string]> = [];
  const embeddedFamilies = new Map<object, string>(embedded.map((face) => [face.data, face.family]));
  const rejected = new Map<object, RejectedEmbedded>();
  const sources = new WeakMap<object, string>();
  const guarded = new WeakMap<Function, Function>();
  const read = (source: () => Promise<ArrayBuffer>, label: string) => async () => {
    try {
      const bytes = await source();
      sources.set(bytes, label);
      return bytes;
    } catch (error) {
      const addon = addonOfError(message(error));
      if (addon && absent.has(addon)) absentLoads.push([addon, label]);
      else failures.add(`${label}: ${message(error)}`);
      throw error;
    }
  };
  const guard = (loader: any, label: string) => {
    if (typeof loader !== 'function') return loader;
    if (!guarded.has(loader))
      guarded.set(loader, Object.assign(read(loader, label), loader, loader.readOnlyBytes ? { readOnlyBytes: read(loader.readOnlyBytes, label) } : {}));
    return guarded.get(loader);
  };
  const provider = bundled && new Proxy(bundled, {
    get(target, property, receiver) {
      const value = Reflect.get(target, property, receiver);
      const label = typeof property === 'string' && Object.hasOwn(RESOLVERS, property) ? RESOLVERS[property] : undefined;
      return label && typeof value === 'function' ? (...args: any[]) => guard(Reflect.apply(value, target, args), label(args)) : value;
    },
  });
  return {
    failures,
    absentLoads: () => [...absentLoads],
    rejectedEmbedded: () => [...rejected.values()],
    bundled: provider,
    engine: {
      registerFont: (bytes, face) => {
        try {
          return face ? engine.registerFont(bytes, face) : engine.registerFont(bytes);
        } catch (error) {
          const family = embeddedFamilies.get(bytes.buffer);
          const label = sources.get(bytes.buffer);
          if (family !== undefined) {
            if (!rejected.has(bytes.buffer)) rejected.set(bytes.buffer, { family, reason: message(error) });
          } else if (label) failures.add(`${label}: ${message(error)}`);
          throw error;
        }
      },
      ...(engine.registerSubstituteFont && {
        registerSubstituteFont: (id: number, family: string, ...style: [bold?: boolean, italic?: boolean]) => {
          let substitute: number;
          try {
            substitute = engine.registerSubstituteFont!(id, family, ...style);
          } catch (error) {
            failures.add(`substitute for ${family}: ${message(error)}`);
            throw error;
          }
          if (!Number.isInteger(substitute)) {
            failures.add(`substitute for ${family}: the engine returned ${substitute}`);
            throw new Error(`the engine returned ${substitute} for the substitute for ${family}`);
          }
          return substitute;
        },
      }),
      clearFonts: () => engine.clearFonts(),
    } satisfies Engine,
  };
}

export async function layOutFonts(product: Product, bytes: Uint8Array, engineOf: (session: any) => Engine = (session) => session) {
  const session = await product.yrs.createYrsSession();
  try {
    const host = session.openDocx(bytes, true);
    const document = session.materializeDocx() ?? host.document;
    const embedded = await product.utils.extractEmbeddedFontFaces(document);
    const absent = new Set((Object.entries(product.fonts.addons) as Array<[AddonName, Addon]>).filter(([, addon]) => addon.absent).map(([name]) => name));
    const registrations = guardRegistrations(await product.layout.resolveCompleteFontProvider(), engineOf(session), embedded, absent);
    const { bundled } = registrations;
    const measure = product.layout.createRustMeasureSource({ engine: registrations.engine, ...(bundled ? { bundled } : {}) });
    measure.setCompat(document.package?.settings?.compatibilityFlags);
    measure.setEmbeddedFaces(embedded, []);
    const renderEnv = product.editor.documentRenderEnv(document);
    const request = product.editor.buildResidentRegionLayoutRequest(document, 0, renderEnv);
    const requirements: Requirement[] = JSON.parse(session.layoutFontRequirementsJson(JSON.stringify(request)));
    const missing = missingAddons(product.fonts, requirements);
    if (missing.length) throw new MissingAddon(missing.join('; '));
    await measure.prepareFontRequirements(requirements);
    if (registrations.failures.size)
      throw new FailedRegistration(`${product.version}: the product's registry dropped fonts it failed to register: ${[...registrations.failures].join('; ')}`);
    const measurement = measure.measurementConfigForRequirements(requirements);
    if (!measurement) throw new Error(`${product.version}: the product could not load the fonts this document lays out with`);
    const absentAddons = absentAddonUse(product.fonts, requirements, registrations.absentLoads());
    return { session, document, renderEnv, requirements, measurement: measurement as Measurement, rejectedEmbedded: registrations.rejectedEmbedded(), absentAddons };
  } catch (error) {
    session.destroy();
    throw error;
  }
}

export async function recordChains(product: Product, bytes: Uint8Array) {
  let recorder: ReturnType<typeof recordingEngine> | undefined;
  const laid = await layOutFonts(product, bytes, (session) => (recorder = recordingEngine(session)).engine);
  laid.session.destroy();
  return { faces: recorder!.faces, measurement: laid.measurement, rejectedEmbedded: laid.rejectedEmbedded, absentAddons: laid.absentAddons };
}

export function serializeChains(faces: RecordedFace[], { fontChains, ...measurement }: Measurement, rejectedEmbedded: RejectedEmbedded[] = [], absentAddons: AbsentAddon[] = []) {
  const unmapped = Object.keys(measurement).filter((key) => !MEASUREMENT_KEYS.has(key));
  if (unmapped.length) throw new Error(`the product's measurement carries ${unmapped.join(', ')}, which the chains file does not map to recorded fonts`);
  const files = new Map<string, Uint8Array>();
  const chains: ChainsFile = {
    faces: faces.map((face, index) => {
      if (!('bytes' in face)) {
        if (!(face.substituteOf < index)) throw new Error(`substitute font ${index} precedes its base ${face.substituteOf}`);
        return { ...face };
      }
      const hash = sha256(face.bytes);
      const file = files.has(`${hash}.font`) ? `${hash}-${index}.font` : `${hash}.font`;
      files.set(file, face.bytes);
      return { file, sha256: hash };
    }),
    fontChains: {},
    measurement: structuredClone(measurement),
    rejectedEmbedded: rejectedEmbedded.map((face) => ({ ...face })),
    absentAddons: absentAddons.map((use) => ({ addon: use.addon, families: [...use.families], scripts: [...use.scripts] })),
  };
  for (const [key, chain] of Object.entries(fontChains)) {
    if (!chain.length || chain.some((id) => !Number.isInteger(id) || id < 0 || id >= faces.length))
      throw new Error(`font chain ${key} names an unrecorded font: ${chain.join(',')}`);
    chains.fontChains[key] = [...chain];
  }
  return { chains, files };
}

export async function writeDocumentChains(product: Product, bytes: Uint8Array, directory: string, name: string) {
  const { faces, measurement, rejectedEmbedded, absentAddons } = await recordChains(product, bytes);
  const { chains, files } = serializeChains(faces, measurement, rejectedEmbedded, absentAddons);
  await mkdir(directory, { recursive: true });
  for (const [file, data] of files) await writeFile(join(directory, file), data);
  const path = join(directory, `${name}.chains.json`);
  await writeFile(path, JSON.stringify(chains));
  return { path, absentAddons: chains.absentAddons };
}

export function planAbsentAddons(plan: Record<string, any>): Set<AddonName> {
  return new Set((['fonts-cjk', 'fonts-intl'] as const).filter((name) => plan[`published_${name.replace('-', '_')}`]?.docx === 'absent'));
}

export function describeAbsentUse(uses: AbsentAddon[]) {
  return uses.map(({ addon, families, scripts }) =>
    `${addon} absent (${[families.length && `families ${families.join(', ')}`, scripts.length && `scripts ${scripts.join(', ')}`].filter(Boolean).join('; ')})`).join('; ');
}

export const LAST_RELEASE_WITHOUT_REGISTRY_EXPORTS = '0.5.1';

const SEMVER = /^(\d+)\.(\d+)\.(\d+)(?:-([0-9A-Za-z.-]+))?(?:\+[0-9A-Za-z.-]+)?$/;

export function compareSemver(left: string, right: string): number {
  const parse = (version: string) => {
    const match = SEMVER.exec(version);
    if (!match) throw new Error(`${version} is not a semver version`);
    return match;
  };
  const [a, b] = [parse(left), parse(right)];
  for (const index of [1, 2, 3]) {
    const difference = Number(a[index]) - Number(b[index]);
    if (difference) return Math.sign(difference);
  }
  if (!a[4] || !b[4]) return (a[4] ? -1 : 0) + (b[4] ? 1 : 0);
  const [x, y] = [a[4].split('.'), b[4].split('.')];
  for (let index = 0; index < Math.max(x.length, y.length); index++) {
    if (x[index] === undefined) return -1;
    if (y[index] === undefined) return 1;
    const [numericX, numericY] = [/^\d+$/.test(x[index]), /^\d+$/.test(y[index])];
    if (numericX && numericY) {
      const difference = Number(x[index]) - Number(y[index]);
      if (difference) return Math.sign(difference);
    } else if (numericX !== numericY) return numericX ? -1 : 1;
    else if (x[index] !== y[index]) return x[index] < y[index] ? -1 : 1;
  }
  return 0;
}

export function fontPathFor(channel: string, version: string): 'bench-manifest' | 'product-registry' {
  return channel === 'published' && compareSemver(version, LAST_RELEASE_WITHOUT_REGISTRY_EXPORTS) <= 0 ? 'bench-manifest' : 'product-registry';
}

export function describeFontPath(channel: string, version: string) {
  const last = LAST_RELEASE_WITHOUT_REGISTRY_EXPORTS;
  if (channel !== 'published')
    return `${channel}: product-registry, @betteroffice/docx@${version} is the commit under test, never compared with ${last}, the last release without the registry exports`;
  if (fontPathFor(channel, version) === 'bench-manifest')
    return `${channel}: bench-manifest, @betteroffice/docx@${version} <= ${last}, the last release without the registry exports; its native column stays on the old font map until the next release`;
  return `${channel}: product-registry, @betteroffice/docx@${version} > ${last}, the last release without the registry exports`;
}

export function missingExportsFailure(channel: string, version: string, detail: string) {
  const reason = channel === 'published'
    ? `is newer than ${LAST_RELEASE_WITHOUT_REGISTRY_EXPORTS}, the last release without the registry exports`
    : 'is the commit under test';
  return `${channel}: @betteroffice/docx@${version} ${reason}, so it must carry the registry exports the recorder needs: ${detail}`;
}

export function checkPublishedVersion(installed: string, source: unknown, planned: unknown) {
  if (installed !== source || installed !== planned)
    throw new Error(`published @betteroffice/docx ${installed} differs from the published native source's ${String(source)} (planned ${String(planned)})`);
}

async function main(env: NodeJS.ProcessEnv) {
  const required = (name: string) => {
    const value = env[name]?.trim();
    if (!value) throw new Error(`${name} is required`);
    return resolve(value);
  };
  const output = required('QUALITY_OUTPUT');
  const chainsRoot = required('QUALITY_CHAINS');
  const job = JSON.parse(await readFile(join(output, 'job.json'), 'utf8'));
  const planBytes = await readFile(required('QUALITY_PLAN'));
  if (sha256(planBytes) !== job.plan_sha256) throw new Error('QUALITY_PLAN is not the plan this job was staged from');
  const planAbsent = planAbsentAddons(JSON.parse(planBytes.toString('utf8')));
  const build = JSON.parse(await readFile(join(required('QUALITY_NATIVE'), 'published', 'build.json'), 'utf8'));
  const roots = { published: required('QUALITY_PRODUCT_PUBLISHED'), commit: required('QUALITY_PRODUCT_COMMIT') };
  const paths: Record<string, { path: string; version: string; last_release_without_registry_exports?: string }> = {};
  const lines: string[] = [];
  const failures: string[] = [];
  for (const [channel, root] of Object.entries(roots)) {
    const version = await packageVersion(root);
    if (channel === 'published') checkPublishedVersion(version, build.package_version, job.published_version);
    if (fontPathFor(channel, version) === 'bench-manifest') {
      paths[channel] = { path: 'bench-manifest', version, last_release_without_registry_exports: LAST_RELEASE_WITHOUT_REGISTRY_EXPORTS };
      lines.push(describeFontPath(channel, version));
      console.log(lines.at(-1));
      continue;
    }
    let product: Product;
    try {
      product = await loadProduct(root, channel === 'published' ? planAbsent : new Set());
    } catch (error) {
      if (!(error instanceof MissingExport)) throw error;
      throw new Error(missingExportsFailure(channel, version, error.message));
    }
    lines.push(`${channel}: @betteroffice/fonts@${product.fonts.version}, ${describeAddons(product.fonts)}`);
    console.log(lines.at(-1));
    let absentDocuments = 0;
    for (const sample of job.samples) {
      try {
        const { absentAddons } = await writeDocumentChains(product, new Uint8Array(await readFile(join(output, sample.id, 'source.docx'))), join(chainsRoot, channel), sample.id);
        if (absentAddons.length) {
          absentDocuments++;
          console.log(`${channel} ${sample.id}: fell back for ${describeAbsentUse(absentAddons)}`);
        }
      } catch (error) {
        if (!(error instanceof MissingAddon)) throw error;
        failures.push(`${channel} ${sample.id}: ${error.message}`);
        console.log(failures.at(-1));
      }
    }
    lines.push(`${channel}: ${absentDocuments} documents fell back for an absent font add-on`);
    console.log(lines.at(-1));
    paths[channel] = { path: 'product-registry', version };
    lines.push(describeFontPath(channel, version));
    console.log(lines.at(-1));
  }
  const summary = [...lines, ...failures.map((failure) => `FAILED ${failure}`)];
  if (env.GITHUB_STEP_SUMMARY)
    await appendFile(env.GITHUB_STEP_SUMMARY, `### DOCX native font chains\n\n${summary.map((line) => `- ${line}`).join('\n')}\n`);
  if (failures.length) throw new Error(`documents need font add-ons their channel lacks:\n${failures.join('\n')}`);
  await mkdir(chainsRoot, { recursive: true });
  await writeFile(join(chainsRoot, 'paths.json'), JSON.stringify(paths, null, 2) + '\n');
}

if (import.meta.main) await main(process.env);
