import { afterAll, expect } from 'bun:test';
import { appendFileSync } from 'node:fs';
import { relative } from 'node:path';
import type { RustParagraphIds } from '../docx/rustSaveFacade';
import { withSessionRevisionIds } from '../docx/sessionRevisionIds';
import type { BlockContent, Document, DocumentBody } from '../types/document';
import type { YrsSession } from './index';
import { withPartMetadata, type ResidentPartMetadata } from './residentParts';
import { captureSessionSave, type DocxSessionSave } from './saveYrsDocx';
import { sectionSavePlanTestSeam, type SectionPlanInputObserver } from './sectionSavePlanTestSeam';
import { DocxSectionError } from './sections';
import { sessionSourcePackage } from './sessionInternals';
import type { EditSession } from './wasm/index';
import { projectedSessionKey, stableStringify, yrsToDocument } from './yrsToDocument';

type Plan = NonNullable<RustParagraphIds['sectionEdits']>;
type Result = { plan?: Plan; lossy: boolean; strict: boolean } | { error: string } | null;
type NativeSession = EditSession & { section_save_plan_json(input: string): string };

const sessions = new WeakMap<YrsSession, EditSession>();
const stats = { checked: 0, differences: 0, undefined: 0, plan: 0, ordinary: 0, refusal: 0 };

afterAll(() => {
  if (stats.differences > 0) {
    throw new Error(`Section save plan parity found ${stats.differences} differences`);
  }
});

export function sectionSavePlanParityStats(): Readonly<typeof stats> {
  return stats;
}

sectionSavePlanTestSeam.hook = {
  attach: (session, raw) => { sessions.set(session, raw); },
  check: async (session, output, capture, sourceBuffer, useLocalProvenance) => {
    const { plan, refusal } = await compareSectionPlans(session, output, capture, sourceBuffer, useLocalProvenance);
    if (refusal) throw refusal;
    return plan;
  },
};

function parityTestFile(): string {
  const getState = Reflect.get(expect, 'getState') as (() => { testPath?: string }) | undefined;
  const testPath = getState?.call(expect)?.testPath;
  if (testPath) return relative(process.cwd(), testPath);
  const stackPath = new Error().stack?.match(/(?:\(|\s)((?:\/|[A-Za-z]:[\\/]).*?[._](?:test|spec)\.[cm]?[jt]sx?):\d+:\d+\)?$/m)?.[1];
  const entryPath = /[._](?:test|spec)\.[cm]?[jt]sx?$/.test(Bun.main) ? Bun.main : undefined;
  const file = stackPath ?? entryPath;
  return file ? relative(process.cwd(), file) : '<unknown>';
}

function sessionKeys(blocks: readonly BlockContent[], keys: Array<string | null>): void {
  for (const block of blocks) {
    if (block.type === 'paragraph') keys.push(projectedSessionKey(block) ?? null);
    else if (block.type === 'table') {
      for (const row of block.rows) for (const cell of row.cells) sessionKeys(cell.content, keys);
    } else if (block.type === 'blockSdt') sessionKeys(block.content, keys);
  }
}

function s13Body(document: Readonly<Document>): DocumentBody {
  const request = { document: document.package.document };
  return withSessionRevisionIds(request, () => JSON.parse(JSON.stringify(request.document)) as DocumentBody);
}

async function referenceResult(
  session: YrsSession,
  output: Document,
  capture: DocxSessionSave,
  sourceBuffer: ArrayBuffer,
  observeInput?: SectionPlanInputObserver
) {
  const reference = sectionSavePlanTestSeam.reference!.plan;
  let plan: Awaited<ReturnType<typeof reference>> = undefined;
  let refusal: DocxSectionError | undefined;
  try {
    plan = await reference(session, output, capture, sourceBuffer, observeInput);
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
  const reportPath = process.env.SECTION_PLAN_PARITY_OUT;
  const testFile = reportPath ? parityTestFile() : undefined;
  const raw = sessions.get(session) as NativeSession | undefined;
  let native: Result | { exception: string } | undefined;
  const result = await referenceResult(session, output, capture, sourceBuffer, (source, sourcePart) => {
    if (!raw && !source) return;
    try {
      if (!raw) throw new Error('Section parity requires an opened WASM session');
      const outputSessionKeys: Array<string | null> = [];
      sessionKeys(output.package.document.content, outputSessionKeys);
      const { metadata } = sectionSavePlanTestSeam.reference!;
      native = JSON.parse(raw.section_save_plan_json(JSON.stringify({
        source: source && s13Body(source),
        output: s13Body(output),
        outputSessionKeys,
        metadataUnchanged: source !== null && metadata(source) === metadata(output),
        sourcePart: sourcePart === undefined ? null : Array.from(sourcePart),
        useLocalProvenance,
      }))) as Result;
    } catch (error) {
      native = { exception: String(error) };
    }
  });
  if (native === undefined) return result;
  const kind = result.refusal ? 'refusal' : !result.plan ? 'undefined' : result.plan.plan ? 'plan' : 'ordinary';
  stats.checked += 1;
  stats[kind] += 1;
  const expectedJson = stableStringify(result.expected);
  const nativeJson = stableStringify(native);
  const differs = nativeJson !== expectedJson;
  if (differs) stats.differences += 1;
  if (reportPath) {
    const checked = `checked ${testFile} ${kind}\n`;
    const diff = differs ? `DIFF ${testFile}\t${expectedJson}\t${nativeJson}\n` : '';
    appendFileSync(reportPath, checked + diff);
  }
  if (differs) {
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
