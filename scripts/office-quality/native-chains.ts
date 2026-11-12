import { createHash } from 'node:crypto';
import { appendFile, mkdir, readFile, writeFile } from 'node:fs/promises';
import { join, resolve } from 'node:path';
import { pathToFileURL } from 'node:url';

type Module = Record<string, any>;

export interface Engine {
  registerFont(bytes: Uint8Array, face?: { family: string; bold: boolean; italic: boolean }): number;
  registerSubstituteFont?(id: number, family: string): number;
  clearFonts(): void;
}

export type RecordedFace = { bytes: Uint8Array } | { substituteOf: number; family: string };

export interface ChainsFile {
  faces: Array<{ file: string; sha256: string } | { substituteOf: number; family: string; bold: boolean; italic: boolean }>;
  fontChains: Record<string, number[]>;
}

export interface Product {
  root: string;
  version: string;
  yrs: Module;
  layout: Module;
  utils: Module;
  editor: Module;
}

export class MissingExport extends Error {}

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

export async function loadProduct(root: string): Promise<Product> {
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
  return { root, version: manifest.version, yrs: modules.yrs, layout: modules.layout, utils: modules.utils, editor: modules.editor };
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
        registerSubstituteFont: (id: number, family: string) =>
          record(engine.registerSubstituteFont!(id, family), { substituteOf: id, family }),
      }),
      clearFonts: () => {
        engine.clearFonts();
        faces.length = 0;
      },
    },
  };
}

export async function layOutFonts(product: Product, bytes: Uint8Array, engineOf: (session: any) => Engine = (session) => session) {
  const session = await product.yrs.createYrsSession();
  try {
    const host = session.openDocx(bytes, true);
    const document = session.materializeDocx() ?? host.document;
    const bundled = await product.layout.resolveCompleteFontProvider();
    const measure = product.layout.createRustMeasureSource({ engine: engineOf(session), ...(bundled ? { bundled } : {}) });
    measure.setCompat(document.package?.settings?.compatibilityFlags);
    measure.setEmbeddedFaces(await product.utils.extractEmbeddedFontFaces(document), []);
    const renderEnv = product.editor.documentRenderEnv(document);
    const request = product.editor.buildResidentRegionLayoutRequest(document, 0, renderEnv);
    const requirements = JSON.parse(session.layoutFontRequirementsJson(JSON.stringify(request)));
    await measure.prepareFontRequirements(requirements);
    const measurement = measure.measurementConfigForRequirements(requirements);
    if (!measurement) throw new Error(`${product.version}: the product could not load the fonts this document lays out with`);
    return { session, document, renderEnv, requirements, measurement: measurement as { fontChains: Record<string, number[]> } };
  } catch (error) {
    session.destroy();
    throw error;
  }
}

export async function recordChains(product: Product, bytes: Uint8Array) {
  let recorder: ReturnType<typeof recordingEngine> | undefined;
  const laid = await layOutFonts(product, bytes, (session) => (recorder = recordingEngine(session)).engine);
  laid.session.destroy();
  return { faces: recorder!.faces, fontChains: laid.measurement.fontChains };
}

export function serializeChains(faces: RecordedFace[], fontChains: Record<string, number[]>) {
  const files = new Map<string, Uint8Array>();
  const chains: ChainsFile = {
    faces: faces.map((face, index) => {
      if (!('bytes' in face)) {
        if (!(face.substituteOf < index)) throw new Error(`substitute font ${index} precedes its base ${face.substituteOf}`);
        return { substituteOf: face.substituteOf, family: face.family, bold: false, italic: false };
      }
      const hash = sha256(face.bytes);
      const file = files.has(`${hash}.font`) ? `${hash}-${index}.font` : `${hash}.font`;
      files.set(file, face.bytes);
      return { file, sha256: hash };
    }),
    fontChains: {},
  };
  for (const [key, chain] of Object.entries(fontChains)) {
    if (!chain.length || chain.some((id) => !Number.isInteger(id) || id < 0 || id >= faces.length))
      throw new Error(`font chain ${key} names an unrecorded font: ${chain.join(',')}`);
    chains.fontChains[key] = [...chain];
  }
  return { chains, files };
}

export async function writeDocumentChains(product: Product, bytes: Uint8Array, directory: string, name: string) {
  const { faces, fontChains } = await recordChains(product, bytes);
  const { chains, files } = serializeChains(faces, fontChains);
  await mkdir(directory, { recursive: true });
  for (const [file, data] of files) await writeFile(join(directory, file), data);
  const path = join(directory, `${name}.chains.json`);
  await writeFile(path, JSON.stringify(chains));
  return path;
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
  const build = JSON.parse(await readFile(join(required('QUALITY_NATIVE'), 'published', 'build.json'), 'utf8'));
  const roots = { published: required('QUALITY_PRODUCT_PUBLISHED'), commit: required('QUALITY_PRODUCT_COMMIT') };
  const paths: Record<string, { path: string; version: string }> = {};
  const lines: string[] = [];
  for (const [channel, root] of Object.entries(roots)) {
    const version = await packageVersion(root);
    if (channel === 'published') checkPublishedVersion(version, build.package_version, job.published_version);
    let product: Product;
    try {
      product = await loadProduct(root);
    } catch (error) {
      if (channel !== 'published' || !(error instanceof MissingExport)) throw error;
      paths[channel] = { path: 'bench-manifest', version };
      lines.push(`${channel}: bench-manifest (registry export missing in ${version}): ${error.message}`);
      console.log(lines.at(-1));
      continue;
    }
    for (const sample of job.samples)
      await writeDocumentChains(product, new Uint8Array(await readFile(join(output, sample.id, 'source.docx'))), join(chainsRoot, channel), sample.id);
    paths[channel] = { path: 'product-registry', version };
    lines.push(`${channel}: product-registry, @betteroffice/docx@${version}`);
    console.log(lines.at(-1));
  }
  await mkdir(chainsRoot, { recursive: true });
  await writeFile(join(chainsRoot, 'paths.json'), JSON.stringify(paths, null, 2) + '\n');
  if (env.GITHUB_STEP_SUMMARY)
    await appendFile(env.GITHUB_STEP_SUMMARY, `### DOCX native font chains\n\n${lines.map((line) => `- ${line}`).join('\n')}\n`);
}

if (import.meta.main) await main(process.env);
