import { DocxSectionError, sectionPropertiesEqual } from './sections';
import { preloadOpcWasm, unzipContainer } from '../docx/wasm';
/** Session save that reports where each saved paragraph can be found after reopening. */

import { engineProgress } from '../utils/engineProgress';
import type { RepackOptions } from '../docx/rezip';
import {
  headerFooterAliasOf,
  isWrittenByCanonical,
  markHeaderFooterAlias,
} from '../docx/headerFooterAliasProjection';
import { type RustParagraphIds, writeDocumentWithRust } from '../docx/rustSaveFacade';
import type { Comment, Paragraph } from '../types/content';
import type { BlockContent, Document } from '../types/document';
import type { YrsSession } from './index';
import { resolveCommentMedia } from './hostMedia';
import { editorSaveKeys, noteSeparatorKeys } from './editorSaveKeys';
import { savedCommentIds } from './residentCommentProjection';
import { sessionMaterializedSource, sessionSourceNoteSeparators, sessionSourcePackage } from './sessionInternals';
import type {
  DocxParagraphIdentitySnapshot,
  DocxParagraphSavePlan,
  DocxPersistedParagraphAnchor,
  DocxSessionParagraphAnchor,
} from './paragraphIdentity';
import { projectedSessionKey, yrsToDocument } from './yrsToDocument';

/** A saved paragraph's session anchor and the persisted anchor that resolves to it once reopened. */
export interface DocxSavedParagraph {
  session: DocxSessionParagraphAnchor;
  persisted: DocxPersistedParagraphAnchor;
}

export interface DocxSavedDocument {
  bytes: Uint8Array;
  /**
   * Every session paragraph the bytes hold with a Word paragraph ID, in
   * document order. Stories that share a part are views of its paragraphs:
   * each view the part holds is listed, and its persisted anchor resolves to
   * the first story's view once reopened.
   */
  paragraphs: DocxSavedParagraph[];
  /**
   * Saved paragraphs whose Word paragraph ID the live session reassigned
   * while the save was written, such as by a duplicate repair after a remote
   * update: their persisted anchor finds them in these bytes, not in the
   * live session.
   */
  conflicts: DocxSavedParagraph[];
}

function storyBlocks(document: Document): BlockContent[][] {
  const pkg = document.package;
  const headerFooterBlocks = [pkg.headers, pkg.footers].flatMap((entries) =>
    entries
      ? [...entries].flatMap(([, part]) =>
          isWrittenByCanonical(part, entries) ? [] : [part.content]
        )
      : []
  );
  return [
    pkg.document.content,
    ...headerFooterBlocks,
    ...(pkg.footnotes ?? []).map((note) => note.content),
    ...(pkg.endnotes ?? []).map((note) => note.content),
  ];
}

const COMMENTS_PART = 'word/comments.xml';

function commentContent(body: unknown): Paragraph[] {
  if (
    Array.isArray(body) &&
    body.length > 0 &&
    body.every((entry) => (entry as { type?: unknown } | null)?.type === 'paragraph')
  ) {
    return body as Paragraph[];
  }
  const text = typeof body === 'string' ? body : '';
  return [
    {
      type: 'paragraph',
      formatting: {},
      content: text ? [{ type: 'run', formatting: {}, content: [{ type: 'text', text }] }] : [],
    },
  ];
}

function commentBodyKey(body: unknown): string | undefined {
  return JSON.stringify(body, (_key, value: unknown) => {
    if (!value || typeof value !== 'object' || Array.isArray(value)) return value;
    return Object.fromEntries(Object.entries(value).sort(([left], [right]) => left < right ? -1 : left > right ? 1 : 0));
  });
}

function savedComments(
  session: YrsSession,
  base: Document
): { comments: Comment[]; sessionIds: Map<number, string>; changed: boolean } {
  const source = base.package.document.comments ?? [];
  const listed = session.listComments();
  resolveCommentMedia(listed, (token) => session.mediaDataUrl(token));
  const ids = savedCommentIds(source.map((comment) => comment.id), listed);
  const held = new Map(listed.map((comment) => [comment.id, comment]));
  let changed = false;
  const parentOf = (parentId: string | null) => {
    if (parentId === null) return undefined;
    const id = ids.get(parentId);
    if (id !== undefined) return id;
    const numeric = Number(parentId);
    return Number.isFinite(numeric) && String(numeric) === parentId ? numeric : undefined;
  };
  const comments = source.map((comment) => {
    const entry = held.get(String(comment.id));
    held.delete(String(comment.id));
    if (!entry || (!entry.author && !entry.date && entry.body == null &&
      !entry.done && entry.parentId === null)) return comment;
    const bodyChanged = entry.body != null &&
      commentBodyKey(entry.body) !== commentBodyKey(comment.blockContent ?? comment.content);
    const author = entry.author;
    const date = entry.date === (comment.date ?? '') ? comment.date : entry.date || undefined;
    const parentId = parentOf(entry.parentId);
    if (!bodyChanged && author === comment.author && date === comment.date &&
      entry.done === (comment.done ?? false) && parentId === comment.parentId) return comment;
    changed = true;
    return {
      ...comment, author, date, done: entry.done, parentId,
      ...(comment.status === undefined ? {} : { status: entry.done ? 'resolved' as const : 'active' as const }),
      ...(bodyChanged ? { content: commentContent(entry.body), blockContent: undefined } : {}),
    };
  });
  const sessionIds = new Map<number, string>();
  for (const entry of held.values()) {
    changed = true;
    const id = ids.get(entry.id)!;
    if (String(id) !== entry.id) sessionIds.set(id, entry.id);
    comments.push({
      id,
      author: entry.author,
      done: entry.done,
      parentId: parentOf(entry.parentId),
      ...(entry.date ? { date: entry.date } : {}),
      content: commentContent(entry.body),
    });
  }
  return { comments, sessionIds, changed };
}

/** The session paragraphs a projection writes with a Word paragraph ID. */
function savedParagraphs(
  document: Document,
  identities: DocxParagraphIdentitySnapshot
): DocxSavedParagraph[] {
  const byKey = new Map(
    identities.paragraphs.flatMap((identity) =>
      identity.session ? [[identity.session.paraId, identity] as const] : []
    )
  );
  const saved: DocxSavedParagraph[] = [];
  const visit = (blocks: readonly BlockContent[]): void => {
    for (const block of blocks) {
      if (block.type === 'table') {
        for (const row of block.rows) for (const cell of row.cells) visit(cell.content);
      } else if (block.type === 'blockSdt') {
        visit(block.content);
      } else if (block.type === 'paragraph' && block.paraId) {
        const key = projectedSessionKey(block);
        const identity = key === undefined ? undefined : byKey.get(key);
        if (!identity?.session || !identity.persisted) continue;
        saved.push({
          session: identity.session,
          persisted: { ...identity.persisted, paraId: block.paraId },
        });
      }
    }
  };
  for (const blocks of storyBlocks(document)) visit(blocks);
  return saved;
}

/**
 * `blocks` with each paragraph, in table cells and block content controls
 * too, replaced by `map`'s result and dropped where it is `null`; the same
 * array when nothing changes.
 */
function mapParagraphs(
  blocks: BlockContent[],
  map: (paragraph: Paragraph) => Paragraph | null
): BlockContent[] {
  const kept: BlockContent[] = [];
  let changed = false;
  for (const block of blocks) {
    let next: BlockContent | null = block;
    if (block.type === 'paragraph') {
      next = map(block);
    } else if (block.type === 'table') {
      const rows = block.rows.map((row) => {
        const cells = row.cells.map((cell) => {
          const content = mapParagraphs(cell.content, map);
          return content === cell.content ? cell : { ...cell, content };
        });
        return cells.every((cell, index) => cell === row.cells[index]) ? row : { ...row, cells };
      });
      if (rows.some((row, index) => row !== block.rows[index])) next = { ...block, rows };
    } else if (block.type === 'blockSdt') {
      const content = mapParagraphs(block.content, map);
      if (content !== block.content) next = { ...block, content };
    }
    changed ||= next !== block;
    if (next) kept.push(next);
  }
  return changed ? kept : blocks;
}

/** `document` with `map` applied to the paragraphs of every story; see {@link mapParagraphs}. */
function mapStoryParagraphs(
  document: Document,
  map: (paragraph: Paragraph) => Paragraph | null
): Document {
  const withContent = <T extends { content: BlockContent[] }>(owner: T): T => {
    const content = mapParagraphs(owner.content, map);
    return content === owner.content ? owner : { ...owner, content };
  };
  const pkg = document.package;
  const parts = (stories: typeof pkg.headers) =>
    stories && new Map(
      [...stories].map(([id, part]) => {
        const projected = withContent(part);
        const canonicalRId = headerFooterAliasOf(part);
        if (projected !== part && canonicalRId !== undefined) {
          markHeaderFooterAlias(projected, canonicalRId);
        }
        return [id, projected] as const;
      })
    );
  return {
    ...document,
    package: {
      ...pkg,
      document: withContent(pkg.document),
      headers: parts(pkg.headers),
      footers: parts(pkg.footers),
      footnotes: pkg.footnotes?.map(withContent),
      endnotes: pkg.endnotes?.map(withContent),
    },
  };
}

/** `document` without the editor-only paragraphs the session has not authored into. */
function withoutSyntheticParagraphs(
  document: Document,
  identities: DocxParagraphIdentitySnapshot
): Document {
  const synthetic = new Set(
    identities.paragraphs.flatMap(({ origin, session }) =>
      origin === 'synthetic' && session ? [session.paraId] : []
    )
  );
  if (synthetic.size === 0) return document;
  return mapStoryParagraphs(document, (paragraph) => {
    const key = projectedSessionKey(paragraph);
    return paragraph.content.length === 0 && key !== undefined && synthetic.has(key)
      ? null
      : paragraph;
  });
}

/** `document` with each paragraph of a spliced part marked with the source occurrence it was seeded from. */
function withSourceOrdinals(
  document: Document,
  parts: NonNullable<DocxParagraphSavePlan['splicedParts']>
): Document {
  const ordinals = new Map(
    parts.flatMap(({ paragraphs }) => paragraphs.map(([ordinal, key]) => [key, ordinal] as const))
  );
  if (ordinals.size === 0) return document;
  return mapStoryParagraphs(document, (paragraph) => {
    const key = projectedSessionKey(paragraph);
    const ordinal = key === undefined ? undefined : ordinals.get(key);
    return ordinal === undefined || paragraph.sourceOrdinal === ordinal
      ? paragraph
      : { ...paragraph, sourceOrdinal: ordinal };
  });
}

/** The identities and paragraph ID plan a session save applies. @internal */
export interface DocxSessionSave {
  identities: DocxParagraphIdentitySnapshot;
  plan: DocxParagraphSavePlan;
}

/** Captures a save of `session`, in the task that projects the saved document. @internal */
export function captureSessionSave(session: YrsSession): DocxSessionSave {
  const identities = session.paragraphIdentities();
  engineProgress();
  const plan = session.paragraphSavePlan();
  engineProgress();
  return { identities, plan };
}

function sectionSaveMetadata(document: Document): string {
  const pkg = document.package;
  const keys = editorSaveKeys({ ...document, package: { ...pkg, document: {
    ...pkg.document, sections: undefined, finalSectionProperties: undefined,
  } } });
  return JSON.stringify({
    keys, numbering: pkg.numbering, settings: pkg.settings, styles: pkg.styles,
    theme: pkg.theme, fontTable: pkg.fontTable, properties: pkg.properties,
    background: pkg.background, customRootBindings: pkg.document.customRootBindings,
    commentAuthors: pkg.document.commentAuthors,
    footnoteSeparators: pkg.footnoteSeparators, endnoteSeparators: pkg.endnoteSeparators,
  }, (_key, value) => value instanceof Map ? [...value] : value);
}

const MODELED_SECTION_CHILDREN = new Set(['headerReference', 'footerReference', 'footnotePr', 'endnotePr', 'type', 'pgSz', 'pgMar', 'paperSrc', 'pgBorders', 'lnNumType', 'pgNumType', 'cols', 'vAlign', 'titlePg', 'bidi', 'docGrid']);

/** Whether every `sectPr` in `xml` holds only children the document model carries, so an ordinary save rewrites them whole. */
function sectionsModeled(xml: string): boolean {
  let depth = 0;
  let prefix = '';
  for (const [, close, name, empty] of xml.matchAll(/<(\/?)([\w.-]+(?::[\w.-]+)?)(?:[^>"']|"[^"]*"|'[^']*')*?(\/?)>/g)) {
    const [space, local] = name!.includes(':') ? name!.split(':') as [string, string] : ['', name!];
    if (depth === 0) {
      if (local === 'sectPr' && !close && !empty) { depth = 1; prefix = space; }
      continue;
    }
    if (close) { depth -= 1; continue; }
    if (depth === 1 && (space !== prefix || !MODELED_SECTION_CHILDREN.has(local))) return false;
    if (!empty) depth += 1;
  }
  return true;
}

/**
 * Section edits the save composes into the source body part, or no `plan`
 * when only the ordinary save, writing the body from the model, can carry
 * them. `lossy` marks an added boundary that only the composer can write
 * without dropping section XML the document model does not carry, and
 * `strict` boundaries placed after a structural edit in a body holding
 * paragraphs the model does not list, which the ordinary save would rewrite.
 * Undefined when the session changed no section.
 */
async function sectionSavePlan(session: YrsSession, output: Document, capture: DocxSessionSave, originalBuffer: ArrayBuffer): Promise<{ plan?: NonNullable<RustParagraphIds['sectionEdits']>; lossy: boolean; strict?: boolean } | undefined> {
  const source = sessionMaterializedSource(session);
  if (!source) return undefined;
  const plannedSplices = [...(capture.plan.splicedParts ?? []), ...(capture.plan.sectionSplicedParts ?? [])];
  const sourceParagraphs: Paragraph[] = [];
  const outputParagraphs: Paragraph[] = [];
  const visit = (blocks: readonly BlockContent[], paragraphs: Paragraph[]) => {
    for (const block of blocks) {
      if (block.type === 'paragraph') paragraphs.push(block);
      else if (block.type === 'table') for (const row of block.rows) for (const cell of row.cells) visit(cell.content, paragraphs);
      else if (block.type === 'blockSdt') visit(block.content, paragraphs);
    }
  };
  visit(source.package.document.content, sourceParagraphs);
  visit(output.package.document.content, outputParagraphs);
  const body = new Set(source.package.document.content.filter(block => block.type === 'paragraph'));
  const byOrdinal = new Map(sourceParagraphs.map((paragraph, index) => [paragraph.sourceOrdinal ?? index, paragraph]));
  const donors = sourceParagraphs.flatMap((paragraph, index) => body.has(paragraph) && paragraph.sectionProperties ? [{ ordinal: paragraph.sourceOrdinal ?? index, properties: paragraph.sectionProperties }] : []);
  const structural = plannedSplices.some(part => part.inserted?.length || part.removed?.length);
  const contiguous = sourceParagraphs.every((paragraph, index) => (paragraph.sourceOrdinal ?? index) === index);
  const originalByKey = new Map(capture.identities.paragraphs.flatMap(identity => identity.session && identity.source && identity.session.story === 'body' ? [[identity.session.paraId, identity.source.paragraphOrdinal] as const] : []));
  const sectionChanges = session.listSections?.().some(section => {
    const before = section.closingParagraphId === 'final' ? source.package.document.finalSectionProperties ?? {}
      : byOrdinal.get(originalByKey.get(section.closingParagraphId) ?? -1)?.sectionProperties;
    return !sectionPropertiesEqual(before, section.authored);
  });
  if (!sectionChanges) return undefined;
  const edits: NonNullable<RustParagraphIds['sectionEdits']>['edits'] = [];
  const bodyParagraphs = output.package.document.content.filter(block => block.type === 'paragraph');
  const sourcePositions = new Map<object, number>();
  let legacy = false;
  let addedBoundary = false;
  const partUri = capture.identities.paragraphs.find(identity => identity.session?.story === 'body' && identity.source)?.source?.partUri;
  const part = partUri?.replace(/^\//, '') ?? 'word/document.xml';
  await preloadOpcWasm();
  const bytes = unzipContainer(new Uint8Array(originalBuffer))[part];
  const lossy = () => addedBoundary && !(bytes instanceof Uint8Array && sectionsModeled(new TextDecoder().decode(bytes)));
  const ordinary = (message: string) => {
    if (lossy()) throw new DocxSectionError('unsupported', message);
    return { lossy: false };
  };
  const supported = new Set(['orientation', 'pageWidth', 'pageHeight', 'marginTop', 'marginBottom', 'marginLeft', 'marginRight', 'headerDistance', 'footerDistance', 'gutter', 'sectionStart']);
  const changed = (before: object, after: object): Record<string, unknown> | undefined => {
    const patch: Record<string, unknown> = {};
    for (const key of new Set([...Object.keys(before), ...Object.keys(after)])) {
      const a = (before as Record<string, unknown>)[key];
      const b = (after as Record<string, unknown>)[key];
      if (sectionPropertiesEqual(a, b)) continue;
      if (!supported.has(key) || b === undefined) return undefined;
      patch[key] = b;
    }
    return patch;
  };
  for (const block of output.package.document.content) {
    if (block.type !== 'paragraph' || !block.sectionProperties) continue;
    const key = projectedSessionKey(block);
    const ordinal = key === undefined ? undefined : originalByKey.get(key);
    const prior = ordinal === undefined ? undefined : byOrdinal.get(ordinal)?.sectionProperties;
    if (prior && sectionPropertiesEqual(prior, block.sectionProperties)) continue;
    addedBoundary ||= !prior;
    const at = outputParagraphs.indexOf(block);
    const following = outputParagraphs.slice(at).map(p => originalByKey.get(projectedSessionKey(p) ?? '')).find(ordinal => ordinal !== undefined);
    const preceding = outputParagraphs.slice(0, at).reverse().map(p => originalByKey.get(projectedSessionKey(p) ?? '')).find(ordinal => ordinal !== undefined);
    let donor = prior ? { ordinal: ordinal!, properties: prior } : donors.find(donor => following !== undefined && donor.ordinal >= following);
    let properties = changed(donor?.properties ?? source.package.document.finalSectionProperties ?? {}, block.sectionProperties);
    if (!properties && ordinal === undefined && preceding !== undefined) {
      const previous = donors.find(donor => donor.ordinal === preceding);
      const inherited = previous && changed(previous.properties, block.sectionProperties);
      if (inherited) { donor = previous; properties = inherited; }
    }
    if (!properties) { legacy = true; continue; }
    const edit = { paragraph: bodyParagraphs.indexOf(block), bodyParagraphs: bodyParagraphs.length, donor: donor?.ordinal ?? null, properties };
    edits.push(edit);
    if (!structural) sourcePositions.set(edit, ordinal ?? at);
  }
  const final = output.package.document.finalSectionProperties ?? {};
  if (!sectionPropertiesEqual(final, source.package.document.finalSectionProperties ?? {})) {
    const properties = changed(source.package.document.finalSectionProperties ?? {}, final);
    if (!properties) legacy = true;
    else edits.push({ paragraph: null, donor: null, properties });
  }
  if (legacy) return ordinary('Cannot safely compose section patches with legacy section properties');
  if (edits.length === 0) return undefined;
  if (!(bytes instanceof Uint8Array)) return ordinary('Section source part is unavailable');
  const digest = await crypto.subtle.digest('SHA-256', Uint8Array.from(bytes).buffer);
  const knownParts = new Set([part, ...capture.identities.paragraphs.flatMap(identity => identity.source ? [identity.source.partUri.replace(/^\//, '')] : []), ...plannedSplices.map(part => part.part)]);
  const onlySections = capture.identities.paragraphs.every(identity => !identity.session || identity.source !== null || identity.origin === 'synthetic') && sectionSaveMetadata(source) === sectionSaveMetadata(output) && capture.plan.assignments.length === 0 && capture.plan.patchedParts.every(part => part.paraIds.length === 0) && [...knownParts].every(path => {
    if (capture.plan.patchedParts.some(part => part.part === path)) return true;
    const splice = plannedSplices.find(part => part.part === path);
    return splice !== undefined && path === part && !splice.inserted?.length && !splice.removed?.length && !splice.tablePatches?.length && splice.changed.every(ordinal => splice.sectionOnly?.includes(ordinal));
  });
  const placed = onlySections ? edits.map(edit => {
    const paragraph = sourcePositions.get(edit);
    return paragraph === undefined ? edit : { paragraph, donor: edit.donor, properties: edit.properties };
  }) : edits;
  return { plan: { onlySections, sha256: Array.from(new Uint8Array(digest), byte => byte.toString(16).padStart(2, '0')).join(''), edits: placed }, lossy: lossy(), strict: structural && !contiguous && edits.some(edit => edit.paragraph !== null) };
}

/**
 * Writes `document`, projected from `session` when `capture` was taken, and
 * records the Word paragraph IDs the bytes hold as saved. `patches` selects
 * the planned parts written as `originalBuffer`'s bytes with IDs patched in,
 * and `splices` the story parts written as its bytes with only the
 * paragraphs that need it re-serialized; each selects none unless
 * `originalBuffer` is the session's source package.
 * @internal
 */
export async function writeSessionSave(
  session: YrsSession,
  document: Document,
  capture: DocxSessionSave,
  originalBuffer: ArrayBuffer,
  options: RepackOptions = {},
  patches: (part: string) => boolean = () => true,
  skipMutations = false,
  splices: (part: string) => boolean = patches
): Promise<DocxSavedDocument> {
  engineProgress();
  const output = withoutSyntheticParagraphs(document, capture.identities);
  engineProgress();
  const sourceBuffer = sessionSourcePackage(session)?.buffer ?? originalBuffer;
  const bodyPart = capture.identities.paragraphs.find(identity => identity.session?.story === 'body' && identity.source)?.source?.partUri.replace(/^\//, '') ?? 'word/document.xml';
  const separators = noteSeparatorKeys(document);
  let sourceSeparators: ReturnType<typeof noteSeparatorKeys> | null | undefined;
  const rewritten = (part: string): boolean => {
    const kind = (['footnotes', 'endnotes'] as const).find(
      (name) => part.replace(/^\//, '') === `word/${name}.xml`
    );
    if (!kind) return false;
    if (sourceSeparators === undefined) {
      sourceSeparators = sessionSourceNoteSeparators(session);
      engineProgress();
    }
    return sourceSeparators?.[kind] !== separators[kind];
  };
  const unspliced = (capture.plan.splicedParts ?? [])
    .filter((entry) => !(splices(entry.part) && !rewritten(entry.part)) && (entry.paragraphPatches?.length || entry.legacyParagraphs?.length))
    .map((entry) => ({ part: entry.part, ordinals: entry.changed }));
  if (unspliced.length) {
    console.debug('DOCX_MINIMAL_PATCH unspliced paragraphs', unspliced);
  }
  const splicesFor = (edits: RustParagraphIds['sectionEdits']) =>
    [...(capture.plan.splicedParts ?? []), ...(edits ? capture.plan.sectionSplicedParts ?? [] : [])].filter(({ part, sha256 }) =>
      (splices(part) && !rewritten(part)) || (edits && part === bodyPart && sha256 === edits.sha256)
    );
  const sectionPlan = await sectionSavePlan(session, output, capture, sourceBuffer);
  engineProgress();
  const refuse = (message: string) => new DocxSectionError('unsupported', message);
  const write = async (sectionEdits: RustParagraphIds['sectionEdits']) => {
    const modelBody = (part: string) => sectionPlan !== undefined && !sectionEdits && part === bodyPart;
    const spliced = splicesFor(sectionEdits).filter(({ part }) => !modelBody(part));
    const plan: RustParagraphIds = {
      ...(sectionEdits ? { sectionEdits } : {}),
      assignments: capture.plan.assignments,
      patchedParts: capture.plan.patchedParts.filter(({ part }) => patches(part) && !rewritten(part) && (sectionPlan === undefined || part !== bodyPart)),
      ...(spliced.length === 0
        ? {}
        : {
            splicedParts: spliced.map(({ part, sha256, paragraphs, changed, sectionOnly, inserted, removed, paragraphPatches, legacyParagraphs, tablePatches }) => {
              const kept = <T>(entries: T[], ordinal: (entry: T) => number) =>
                sectionEdits && part === bodyPart ? entries.filter(entry => !sectionOnly?.includes(ordinal(entry))) : entries;
              return {
                part,
                sha256,
                paragraphs: paragraphs.map(([ordinal]) => ordinal),
                changed: kept(changed, ordinal => ordinal),
                ...(inserted?.length ? { inserted } : {}),
                ...(removed?.length ? { removed } : {}),
                ...(paragraphPatches === undefined ? {} : { paragraphPatches: kept(paragraphPatches, patch => patch.ordinal) }),
                ...(legacyParagraphs === undefined ? {} : { legacyParagraphs: kept(legacyParagraphs, legacy => legacy.ordinal) }),
                ...(tablePatches === undefined ? {} : { tablePatches }),
              };
            }),
          }),
    };
    const { buffer, warnings } = await writeDocumentWithRust(
      withSourceOrdinals(output, spliced), sectionEdits ? sourceBuffer : originalBuffer, options, undefined, undefined, plan, skipMutations
    );
    for (const warning of warnings) {
      if (warning.startsWith('DOCX_MINIMAL_PATCH replace counts')) {
        console.debug(warning);
      }
    }
    return { buffer, plan };
  };
  let sectionEdits = sectionPlan?.plan;
  if (sectionEdits && !sectionEdits.onlySections && !splicesFor(sectionEdits).some(({ part }) => part === bodyPart)) {
    if (sectionPlan?.lossy) throw refuse('Cannot safely compose sections with this document edit');
    sectionEdits = undefined;
  }
  let result: Awaited<ReturnType<typeof write>>;
  try {
    result = await write(sectionEdits);
  } catch (error) {
    if (!sectionEdits) throw error;
    if (sectionPlan?.lossy) throw refuse(error instanceof Error ? error.message : String(error));
    if (sectionPlan?.strict) throw refuse('Cannot safely locate section boundaries after this structural edit');
    result = await write(undefined);
  }
  const { buffer, plan } = result;
  const paragraphs = savedParagraphs(output, capture.identities);
  const bytes = new Uint8Array(buffer);
  engineProgress();
  const held = session.writtenParagraphIds(bytes);
  engineProgress();
  const heldSets = new Map(Object.entries(held).map(([part, ids]) => [part, new Set(ids)]));
  const isWritten = (partUri: string, paraId: string) =>
    heldSets.get(partUri)?.has(paraId.toUpperCase()) ?? false;
  const written = paragraphs.filter(({ persisted }) =>
    isWritten(persisted.story.partUri, persisted.paraId)
  );
  const stale = new Set(
    session
      .recordSavedParagraphIds([
        ...written.map(
          ({ session: anchor, persisted }) => [anchor.paraId, persisted.paraId] as const
        ),
        ...plan.assignments
          .filter(({ part, paraId }) => isWritten(`/${part}`, paraId))
          .map(({ part, ordinal, paraId }) => [`/${part}#${ordinal}`, paraId] as const),
      ])
      .map(([owner, paraId]) => `${owner}\u0000${paraId}`)
  );
  const isStale = ({ session: anchor, persisted }: DocxSavedParagraph) =>
    stale.has(`${anchor.paraId}\u0000${persisted.paraId}`);
  return {
    bytes,
    paragraphs: written.filter((paragraph) => !isStale(paragraph)),
    conflicts: written.filter(isStale),
  };
}

/**
 * Saves a session opened from DOCX bytes and reports each saved paragraph's
 * persisted anchor, read from the same projection the bytes are written
 * from and kept only where the written part holds its Word paragraph ID.
 * Parts whose stories are unchanged since the session opened keep
 * their source bytes, paragraph IDs aside. The comments the session holds
 * are saved with their anchors, replies and resolution. A source paragraph without a
 * Word paragraph ID saves without one unless the host called
 * `persistParagraphIds()` first; the IDs a save writes are recorded as
 * saved, so they keep them against unsaved claims from other replicas, once
 * reconciled with the live session.
 */
export async function saveYrsDocx(
  session: YrsSession,
  options: RepackOptions = {}
): Promise<DocxSavedDocument> {
  const base = session.materializeDocx();
  if (!base?.originalBuffer) {
    throw new Error('saveYrsDocx requires a session opened from DOCX bytes');
  }
  const capture = captureSessionSave(session);
  const comments = savedComments(session, base);
  const document = yrsToDocument(
    session,
    comments.changed
      ? {
          ...base,
          package: {
            ...base.package,
            document: { ...base.package.document, comments: comments.comments },
          },
        }
      : base,
    { commentIds: comments.sessionIds }
  );
  return writeSessionSave(
    session,
    document,
    capture,
    base.originalBuffer,
    options,
    (part) => !comments.changed || part !== COMMENTS_PART
  );
}
