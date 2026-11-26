import { writeSync } from 'node:fs';
import type { RustParagraphIds } from '../docx/rustSaveFacade';
import { withSessionRevisionIds } from '../docx/sessionRevisionIds';
import { preloadOpcWasm, unzipContainer } from '../docx/wasm';
import type { BlockContent, Document, DocumentBody } from '../types/document';
import type { YrsSession } from './index';
import { withPartMetadata, type ResidentPartMetadata } from './residentParts';
import { captureSessionSave, type DocxSessionSave } from './saveYrsDocx';
import { sectionSavePlanTestSeam } from './sectionSavePlanTestSeam';
import { DocxSectionError } from './sections';
import { sessionMaterializedSource, sessionSourcePackage } from './sessionInternals';
import type { EditSession } from './wasm/index';
import { projectedSessionKey, stableStringify, yrsToDocument } from './yrsToDocument';

type Plan = NonNullable<RustParagraphIds['sectionEdits']>;
type Result = { plan?: Plan; lossy: boolean; strict: boolean } | { error: string } | null;
type NativeSession = EditSession & { section_save_plan_json(input: string): string };

const sessions = new WeakMap<YrsSession, EditSession>();
const stats = { checked: 0, differences: 0 };

export function sectionSavePlanParityStats(): Readonly<typeof stats> {
  return stats;
}

sectionSavePlanTestSeam.hook = {
  attach: (session, raw) => { sessions.set(session, raw); },
  check: async (session, output, capture, sourceBuffer, useLocalProvenance) => {
    if (!sessions.has(session) && !sessionMaterializedSource(session)) {
      return sectionSavePlanTestSeam.reference!.plan(session, output, capture, sourceBuffer);
    }
    const { plan, refusal } = await compareSectionPlans(session, output, capture, sourceBuffer, useLocalProvenance);
    if (refusal) throw refusal;
    return plan;
  },
};

process.once('exit', () => {
  writeSync(1, `section-plan-parity checked=${stats.checked} differences=${stats.differences}\n`);
});

function sessionKeys(blocks: readonly BlockContent[], keys: Array<string | null>): void {
  for (const block of blocks) {
    if (block.type === 'paragraph') keys.push(projectedSessionKey(block) ?? null);
    else if (block.type === 'table') {
      for (const row of block.rows) for (const cell of row.cells) sessionKeys(cell.content, keys);
    } else if (block.type === 'blockSdt') sessionKeys(block.content, keys);
  }
}

function s13Body(document: Document): DocumentBody {
  const request = { document: document.package.document };
  return withSessionRevisionIds(request, () => JSON.parse(JSON.stringify(request.document)) as DocumentBody);
}

async function referenceResult(
  session: YrsSession,
  output: Document,
  capture: DocxSessionSave,
  sourceBuffer: ArrayBuffer
) {
  const reference = sectionSavePlanTestSeam.reference!.plan;
  let plan: Awaited<ReturnType<typeof reference>> = undefined;
  let refusal: DocxSectionError | undefined;
  try {
    plan = await reference(session, output, capture, sourceBuffer);
  } catch (error) {
    if (!(error instanceof DocxSectionError) || error.code !== 'unsupported') throw error;
    refusal = error;
  }
  const expected: Result = refusal ? { error: refusal.message }
    : plan ? { ...plan, strict: plan.strict ?? false } : null;
  return { expected, plan, refusal };
}

async function compareSectionPlans(
  session: YrsSession,
  output: Document,
  capture: DocxSessionSave,
  sourceBuffer: ArrayBuffer,
  useLocalProvenance: boolean
) {
  const source = sessionMaterializedSource(session);
  const outputSessionKeys: Array<string | null> = [];
  sessionKeys(output.package.document.content, outputSessionKeys);
  const part = capture.identities.paragraphs
    .find(identity => identity.session?.story === 'body' && identity.source)
    ?.source?.partUri.replace(/^\//, '') ?? 'word/document.xml';
  const { metadata } = sectionSavePlanTestSeam.reference!;
  if (source) await preloadOpcWasm();
  const reference = referenceResult(session, output, capture, sourceBuffer);
  let native: Result | { exception: string };
  try {
    const raw = sessions.get(session) as NativeSession | undefined;
    if (!raw) throw new Error('Section parity requires an opened WASM session');
    const bytes = source ? unzipContainer(new Uint8Array(sourceBuffer))[part] : undefined;
    native = JSON.parse(raw.section_save_plan_json(JSON.stringify({
      source: source && s13Body(source),
      output: s13Body(output),
      outputSessionKeys,
      metadataUnchanged: source !== null && metadata(source) === metadata(output),
      sourcePart: bytes === undefined ? null : Array.from(bytes),
      useLocalProvenance,
    }))) as Result;
  } catch (error) {
    native = { exception: String(error) };
  }
  const result = await reference;
  stats.checked += 1;
  const expectedJson = stableStringify(result.expected);
  const nativeJson = stableStringify(native);
  if (nativeJson !== expectedJson) {
    stats.differences += 1;
    process.exitCode = 1;
    throw new Error(`Section save plan differs\nTypeScript: ${expectedJson}\nRust: ${nativeJson}`);
  }
  return result;
}

type ProjectedPlanOptions = {
  base?: Document;
  sourceBuffer?: ArrayBuffer;
  useLocalProvenance?: boolean;
  editOutput?: (output: Document) => void;
};

function projectPlan(
  session: YrsSession,
  {
    base = session.materializeDocx()!,
    sourceBuffer = sessionSourcePackage(session)?.buffer ?? base.originalBuffer!,
    useLocalProvenance = false,
    editOutput,
  }: ProjectedPlanOptions = {}
) {
  const capture = captureSessionSave(session);
  if (useLocalProvenance) capture.plan = session.paragraphSavePlanWithLocalProvenance();
  const output = sectionSavePlanTestSeam.reference!.withoutSynthetic(
    yrsToDocument(session, base), capture.identities
  );
  editOutput?.(output);
  return { output, capture, sourceBuffer, useLocalProvenance };
}

export async function checkProjectedSectionPlan(session: YrsSession, options?: ProjectedPlanOptions): Promise<Result> {
  const { output, capture, sourceBuffer, useLocalProvenance } = projectPlan(session, options);
  return (await compareSectionPlans(session, output, capture, sourceBuffer, useLocalProvenance)).expected;
}

export async function referenceProjectedSectionPlan(session: YrsSession, options?: ProjectedPlanOptions): Promise<Result> {
  const { output, capture, sourceBuffer } = projectPlan(session, options);
  return (await referenceResult(session, output, capture, sourceBuffer)).expected;
}

export function setSectionMetadata(session: YrsSession, metadata: ResidentPartMetadata): ProjectedPlanOptions {
  const raw = sessions.get(session)!;
  raw.set_part_metadata(JSON.stringify(metadata), '[]');
  return { base: withPartMetadata(session.materializeDocx()!, metadata) };
}

export async function checkSectionMetadataPlan(
  session: YrsSession, metadata: ResidentPartMetadata
): Promise<Result> {
  return checkProjectedSectionPlan(session, setSectionMetadata(session, metadata));
}
