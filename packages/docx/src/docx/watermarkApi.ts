/**
 * Watermark document API
 *
 * Platform-agnostic helpers for reading and applying a document watermark.
 * Shared by the React and Vue adapters (and the imperative ref API) so the
 * "Design → Watermark" behavior stays identical across frameworks.
 *
 * A watermark lives on `HeaderFooter.watermark`. MS Word repeats the same
 * watermark across the default, first-page, and even-page headers of every
 * section, so `setDocumentWatermark`:
 *
 * 1. Applies the watermark (a per-header copy) to every existing header, and
 * 2. Creates the header parts a section needs but lacks, so the watermark
 *    still shows on title pages (`w:titlePg`) and even pages
 *    (`w:evenAndOddHeaders`) and on documents that had no header at all.
 *
 * To avoid breaking header inheritance (a section that omits a header
 * reference inherits the previous section's header — Word's "link to
 * previous"), a missing `first`/`even` header part is only created when **no**
 * header of that type exists anywhere in the document, i.e. there is nothing to
 * inherit. All updates are immutable — a new `Document` is returned so the
 * change lands in the host's undo/redo history.
 */

import type {
  Document,
  HeaderFooter,
  HeaderFooterType,
  Watermark,
  Relationship,
  SectionProperties,
} from '../types/document';
import { RELATIONSHIP_TYPES } from './relsParser';

/** Read the document's watermark (the first header that carries one). */
export function getDocumentWatermark(doc: Document | null | undefined): Watermark | undefined {
  const headers = doc?.package.headers;
  if (!headers) return undefined;
  for (const hf of headers.values()) {
    if (hf.watermark) return hf.watermark;
  }
  return undefined;
}

/** Strip the watermark from every header. */
function removeFromAllHeaders(doc: Document): Document {
  const headers = doc.package.headers;
  if (!headers || headers.size === 0) return doc;
  let changed = false;
  const next = new Map<string, HeaderFooter>();
  for (const [rId, hf] of headers) {
    if (hf.watermark) {
      const { watermark: _omit, ...rest } = hf;
      next.set(rId, rest);
      changed = true;
    } else {
      next.set(rId, hf);
    }
  }
  if (!changed) return doc;
  return { ...doc, package: { ...doc.package, headers: next } };
}

/**
 * Apply the watermark to every existing header. Each header gets its own copy
 * of the watermark object: a picture watermark's `relId` is a header-part-local
 * relationship id, so the parts must not share one object or the save step
 * would stamp the same (wrong) rId into every header's rels.
 */
function setOnAllHeaders(doc: Document, watermark: Watermark): Document {
  const headers = doc.package.headers;
  if (!headers || headers.size === 0) return doc;
  const next = new Map<string, HeaderFooter>();
  for (const [rId, hf] of headers) {
    next.set(rId, { ...hf, watermark: { ...watermark } });
  }
  return { ...doc, package: { ...doc.package, headers: next } };
}

/** All section-properties objects in the body (earlier sections + final). */
function collectSectionProperties(doc: Document): SectionProperties[] {
  const body = doc.package.document;
  const out: SectionProperties[] = [];
  if (body.sections) {
    for (const s of body.sections) if (s.properties) out.push(s.properties);
  }
  if (body.finalSectionProperties) out.push(body.finalSectionProperties);
  return out;
}

/**
 * Ensure the watermark is carried by a header for every page a section
 * displays. Creates the header parts that are missing AND would otherwise show
 * nothing (so the watermark would be absent):
 *
 * - `default`: created only for a document with no headers at all.
 * - `first`: created when a section sets `titlePg` but no first-page header
 *   exists anywhere (Word blanks page 1's default header, so without this the
 *   watermark is missing on title pages).
 * - `even`: created when a section sets `evenAndOddHeaders` but no even-page
 *   header exists anywhere (otherwise even pages show no watermark).
 *
 * A created part is shared by reference across every section that needs it.
 */
function ensureWatermarkHeaderCoverage(doc: Document, watermark: Watermark): Document {
  const pkg = doc.package;
  const body = pkg.document;
  const existingHeaders = pkg.headers;
  const sectionProps = collectSectionProperties(doc);

  const existingTypes = new Set<HeaderFooterType>();
  if (existingHeaders) for (const hf of existingHeaders.values()) existingTypes.add(hf.hdrFtrType);

  const hasNoHeaders = !existingHeaders || existingHeaders.size === 0;
  const anyTitlePg = sectionProps.some((sp) => sp.titlePg);
  const anyEvenOdd = sectionProps.some((sp) => sp.evenAndOddHeaders);

  // Which header types need a brand-new part (displayed, but nothing to inherit).
  const createDefault = hasNoHeaders;
  const createFirst = anyTitlePg && !existingTypes.has('first');
  const createEven = anyEvenOdd && !existingTypes.has('even');

  if (!createDefault && !createFirst && !createEven) return doc;

  const rels: Map<string, Relationship> = pkg.relationships
    ? new Map(pkg.relationships)
    : new Map();
  const headers = new Map<string, HeaderFooter>(existingHeaders ?? []);

  // Unique header target filename (header1.xml, header2.xml, …).
  const usedTargets = new Set<string>();
  for (const r of rels.values()) {
    if (r.target) usedTargets.add(r.target.replace(/^\/?word\//, '').toLowerCase());
  }
  function nextHeaderTarget(): string {
    let n = 1;
    while (usedTargets.has(`header${n}.xml`)) n++;
    const target = `header${n}.xml`;
    usedTargets.add(target.toLowerCase());
    return target;
  }

  // Create one part per needed type (shared across the sections that need it).
  // A non-numeric rId avoids the rezip pipeline's numeric `rIdN` allocation for
  // images/hyperlinks (which would otherwise collide). See createHeaderWithWatermark history.
  function createHeaderPart(type: HeaderFooterType): string {
    let rId = `rIdWmHdr${type[0].toUpperCase()}${type.slice(1)}`;
    let suffix = 1;
    while (rels.has(rId)) rId = `rIdWmHdr${type[0].toUpperCase()}${type.slice(1)}${suffix++}`;
    rels.set(rId, { id: rId, type: RELATIONSHIP_TYPES.header, target: nextHeaderTarget() });
    headers.set(rId, {
      type: 'header',
      hdrFtrType: type,
      content: [],
      watermark: { ...watermark },
    });
    return rId;
  }
  const created: Partial<Record<HeaderFooterType, string>> = {};
  function refFor(type: HeaderFooterType): string {
    return (created[type] ??= createHeaderPart(type));
  }

  // Add a reference of `type` to a section's properties when it doesn't have one.
  function withRef(sp: SectionProperties | undefined, type: HeaderFooterType): SectionProperties {
    const refs = sp?.headerReferences ? [...sp.headerReferences] : [];
    if (refs.some((r) => r.type === type)) return sp ?? {};
    refs.push({ type, rId: refFor(type) });
    return { ...(sp ?? {}), headerReferences: refs };
  }

  // Apply the needed references to a section's properties.
  function patchSection(sp: SectionProperties | undefined): SectionProperties | undefined {
    let next = sp;
    if (createDefault) next = withRef(next, 'default');
    if (createFirst && next?.titlePg) next = withRef(next, 'first');
    if (createEven && next?.evenAndOddHeaders) next = withRef(next, 'even');
    return next;
  }

  // A document with no sections still has the implicit final section properties
  // (stays undefined when there's nothing to patch — that field is optional).
  const finalSectionProperties = patchSection(body.finalSectionProperties);
  const sections = body.sections?.map((s) => ({
    ...s,
    // `properties` is required, so fall back to the original (patchSection only
    // returns undefined for an undefined input, which never happens here).
    properties: patchSection(s.properties) ?? s.properties,
  }));

  return {
    ...doc,
    package: {
      ...pkg,
      relationships: rels,
      headers,
      document: {
        ...body,
        finalSectionProperties,
        ...(sections ? { sections } : {}),
      },
    },
  };
}

/**
 * Return a new `Document` with the watermark applied to all headers, or removed
 * when `watermark` is null. Creates the header parts a section needs but lacks
 * (default for a headerless doc; first/even for title/even pages) so the
 * watermark shows on every page MS Word would show it.
 */
export function setDocumentWatermark(doc: Document, watermark: Watermark | null): Document {
  if (!watermark) return removeFromAllHeaders(doc);
  const applied = setOnAllHeaders(doc, watermark);
  return ensureWatermarkHeaderCoverage(applied, watermark);
}
