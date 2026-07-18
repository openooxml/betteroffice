/**
 * The pure layout COMPUTE pass shared by the React and Vue adapters — issue
 * #696 Tier 2, the clean half of the engine spine.
 *
 * This is the 6-step pass from React's `useLayoutPipeline` minus the paint
 * + scroll/event side-effects (which stay adapter-side, where the framework
 * timing lives): PM doc → flow blocks → measure → header/footer resolve →
 * margin extension → Rust pagination (+ two-pass footnote stabilization) →
 * footnote render items. It is pure (no DOM, no refs, no rAF) and returns
 * everything the adapter needs to render.
 *
 * Two injected seams: `measureBlocks` (each adapter passes the Rust
 * measurement implementation) and `paginationSource` (the mandatory Rust
 * pagination engine — every pagination pass, including each footnote
 * stabilization iteration, runs through it). `getHfPmDoc` is the
 * HF-unification seam (prefer the persistent PM doc over re-parsing
 * `HeaderFooter.content`).
 */

import type { EditorTreeNode } from '../types/editorTree';

import {
  toMeasuredBlocks,
  pageGeometryFromPage,
  type ColumnLayout,
  type LayoutBlock,
  type LayoutOptions,
  type LayoutPaginationSource,
  type FootnoteContent,
  type FootnoteRenderItem,
  type HeaderFooterContent,
  type Layout,
  type BlockExtent,
  type PageMargins,
  type PageHeaderFooterRefs,
} from '../layout/pagination';
import {
  toLayoutBlocks,
  computePerBlockWidths,
  demoteBlockLikeFloatingTables,
  collectFootnoteRefs,
  mapFootnotesToPages,
  convertHeaderFooterToContent,
  buildFootnoteRenderItems,
  distributeFootnotesIntoColumns,
  stabilizeFootnoteLayout,
  FOOTNOTE_COLUMN_GAP_PX,
  FOOTNOTE_SEPARATOR_HEIGHT,
  extendMarginsForHeaderFooter,
  twipsToPixels,
  type FontStyle,
  type FloatPageGeometry,
} from '../layout';
import {
  buildNoteContentMap,
  noteReferenceMapId,
  type NotePresentation,
} from '../layout/regions/footnoteLayout';
import type {
  DisplayListFieldWidths,
  DisplayListHeadersFooters,
  DisplayListHfVariant,
} from '../layout/render/rustDisplayList';
import {
  beginRustMeasureLayoutPass,
  measureTextWidthWithActiveRustSource,
} from '../layout/measure/rustMeasureSource';
import type {
  Document,
  EndnoteProperties,
  FootnoteProperties,
  HeaderFooter,
  NoteKind,
  NumberFormat,
  SectionProperties,
  StyleDefinitions,
  Theme,
  Watermark,
} from '../types/document';
import { formatNumber } from '../docx/numberingParser';

interface PageSizePx {
  w: number;
  h: number;
}

/** Adapter-supplied block measurer (React's is caching). */
export type MeasureBlocksFn = (
  blocks: LayoutBlock[],
  contentWidth: number | number[],
  pageGeometry?: FloatPageGeometry
) => BlockExtent[];

export interface ComputeLayoutInputs {
  /** Legacy tree state; null when an adapter supplies yrs body blocks. */
  state: { readonly doc: { readonly attrs: Record<string, any> } } | null;
  document: Document | null;
  pageSize: PageSizePx;
  margins: PageMargins;
  columns: ColumnLayout | undefined;
  finalPageSize: PageSizePx;
  finalMargins: PageMargins;
  finalColumns: ColumnLayout | undefined;
  pageGap: number;
  contentWidth: number;
  theme: Theme | null | undefined;
  styles: StyleDefinitions | null | undefined;
  sectionProperties: SectionProperties | null | undefined;
  finalSectionProperties: SectionProperties | null | undefined;
  /** Resolved HF objects for the section (default + first-page). */
  headerContent: HeaderFooter | null | undefined;
  footerContent: HeaderFooter | null | undefined;
  firstPageHeaderContent: HeaderFooter | null | undefined;
  firstPageFooterContent: HeaderFooter | null | undefined;
  /** Optional adapter-owned body source; null falls back to ProseMirror lowering. */
  yrsBodyBlocks?: (env: { pageContentHeight: number }) => LayoutBlock[] | null;
  /** Optional adapter-owned non-body story source; null preserves PM lowering. */
  yrsStoryBlocks?: (storyId: string) => LayoutBlock[] | null;
  measureBlocks: MeasureBlocksFn;
  /**
   * The mandatory Rust pagination engine
   * (`layout/pagination/rustPagination.ts`). Every pagination pass — full,
   * footnote pass 1, and each footnote-stabilization iteration — runs
   * `paginationSource.paginate(measured, options)`. The caller must gate
   * layout passes on the source's readiness (`isReady()` / `whenReady()`);
   * `paginate` throws on an unloaded engine or an engine failure and the
   * pipeline surfaces the error.
   *
   * The Rust engine keeps no resume checkpoints: every pass is a full
   * relayout in native code — the measured wager that replaced the TS
   * incremental/off-thread kernels.
   */
  paginationSource: LayoutPaginationSource;
}

export interface LayoutComputation {
  blocks: LayoutBlock[];
  measures: BlockExtent[];
  layout: Layout;
  headerContentForRender: HeaderFooterContent | undefined;
  footerContentForRender: HeaderFooterContent | undefined;
  firstPageHeaderForRender: HeaderFooterContent | undefined;
  firstPageFooterForRender: HeaderFooterContent | undefined;
  hasTitlePg: boolean;
  watermark: Watermark | undefined;
  headerDistancePx: number | undefined;
  footerDistancePx: number | undefined;
  pageBorders: SectionProperties['pageBorders'] | undefined;
  footnotesByPage: Map<number, FootnoteRenderItem[]> | undefined;
}

/**
 * Resolve one page's section-aware footnote column layout.
 *
 * Footnotes paint N-up when the owning section opts into multiple columns.
 * Returns `{ columns: 1, columnWidth: fallback }` for the ordinary single-
 * column path. Legacy pages without section identity retain the former first-
 * multi-column-section fallback until Batch D supplies `sectionIndex`.
 */
function sectionPropertiesForPage(
  document: Document | null,
  page: Layout['pages'][number]
): SectionProperties | undefined {
  const body = document?.package?.document;
  if (!body) return undefined;
  const sections = [
    ...(body.sections ?? []).map((section) => section.properties),
    body.finalSectionProperties,
  ];
  if (page.sectionIndex !== undefined) {
    return sections[page.sectionIndex] ?? body.finalSectionProperties;
  }
  // Legacy paginator output has no section identity. Preserve the prior
  // single-document heuristic until Batch D supplies page.sectionIndex.
  return sections.find((properties) => (properties?.footnoteColumns ?? 1) > 1) ?? sections[0];
}

function orderedSectionProperties(document: Document | null): Array<SectionProperties | undefined> {
  const body = document?.package?.document;
  if (!body) return [];
  return [
    ...(body.sections ?? []).map((section) => section.properties),
    body.finalSectionProperties,
  ];
}

/** Resolve Link-to-Previous inheritance into the page-level rId contract. */
function effectiveHeaderFooterRefs(
  sections: Array<SectionProperties | undefined>,
  sectionIndex: number
): PageHeaderFooterRefs | undefined {
  const refs: PageHeaderFooterRefs = {};
  const assign = (
    kind: 'header' | 'footer',
    type: 'default' | 'first' | 'even',
    rId: string
  ): void => {
    const suffix = type === 'default' ? 'Default' : type === 'first' ? 'First' : 'Even';
    const key = `${kind}${suffix}` as keyof PageHeaderFooterRefs;
    refs[key] = rId;
  };
  for (let index = 0; index <= sectionIndex; index++) {
    const section = sections[index];
    for (const ref of section?.headerReferences ?? []) assign('header', ref.type, ref.rId);
    for (const ref of section?.footerReferences ?? []) assign('footer', ref.type, ref.rId);
  }
  return Object.keys(refs).length ? refs : undefined;
}

/**
 * Consume Batch A/D's optional page identity and fill all derivable effective
 * furniture before painter/display-list handoff. Absent section identity keeps
 * the legacy document-global values supplied by the adapters.
 */
function orchestratePageMetadata(prepared: PreparedLayoutCompute, layout: Layout): void {
  const { document, sectionProperties: legacySection } = prepared.inputs;
  const sections = orderedSectionProperties(document);
  const observedSectionPages = new Map<string, number>();

  for (const page of layout.pages) {
    // A one-section document is unambiguous even before Batch D stamps page
    // identity, so it can receive logical numbering/furniture immediately.
    const sectionIndex = page.sectionIndex ?? (sections.length === 1 ? 0 : undefined);
    if (page.sectionIndex === undefined && sectionIndex !== undefined)
      page.sectionIndex = sectionIndex;
    const section =
      sectionIndex !== undefined ? (sections[sectionIndex] ?? legacySection) : legacySection;
    if (!section) {
      if (prepared.watermark !== undefined && page.watermark === undefined) {
        page.watermark = prepared.watermark;
      }
      continue;
    }

    if (sectionIndex !== undefined) {
      const sectionKey = page.sectionId ?? String(sectionIndex);
      const observedIndex = observedSectionPages.get(sectionKey) ?? 0;
      const sectionPageIndex = page.sectionPageIndex ?? observedIndex;
      observedSectionPages.set(sectionKey, Math.max(observedIndex, sectionPageIndex) + 1);
      page.sectionPageIndex = sectionPageIndex;
      page.headerFooterRefs ??= effectiveHeaderFooterRefs(sections, sectionIndex);

      const numbering = page.pageNumbering ?? section.pageNumbering;
      if (numbering) {
        page.pageNumbering = numbering;
        const logicalNumber = page.sectionPageNumber ?? (numbering.start ?? 1) + sectionPageIndex;
        page.sectionPageNumber = logicalNumber;
        page.pageLabel ??= formatNumber(
          logicalNumber,
          (numbering.format ?? 'decimal') as NumberFormat
        );
      }
    }

    if (page.headerDistance === undefined && section.headerDistance != null) {
      page.headerDistance = twipsToPixels(section.headerDistance);
    }
    if (page.footerDistance === undefined && section.footerDistance != null) {
      page.footerDistance = twipsToPixels(section.footerDistance);
    }
    page.pageBorders ??= section.pageBorders;
    page.watermark ??= section.watermark ?? prepared.watermark;
    page.verticalAlign ??= section.verticalAlign;
  }
}

function resolveFootnoteColumnLayout(
  document: Document | null,
  page: Layout['pages'][number],
  fallbackColumnWidth: number
): { columns: number; columnWidth: number } {
  const fnSection = sectionPropertiesForPage(document, page);
  if (!fnSection?.footnoteColumns) {
    return { columns: 1, columnWidth: fallbackColumnWidth };
  }

  const columns = fnSection.footnoteColumns;
  // Footnote columns span the section's full content width, independent of the
  // body's w:cols. Mirror the painter's width math so a footnote measured here
  // wraps exactly as it paints.
  const sectionContentWidthPx =
    fnSection.pageWidth != null
      ? twipsToPixels(
          fnSection.pageWidth - (fnSection.marginLeft ?? 1440) - (fnSection.marginRight ?? 1440)
        )
      : fallbackColumnWidth;
  const columnWidth = (sectionContentWidthPx - (columns - 1) * FOOTNOTE_COLUMN_GAP_PX) / columns;
  return { columns, columnWidth: Math.max(1, columnWidth) };
}

/**
 * Everything `prepareLayoutCompute` produces before pagination: the measured
 * blocks + options the paginator consumes, plus the already-resolved values
 * `finishLayoutComputation` threads into the result.
 */
interface PreparedLayoutCompute {
  measured: ReturnType<typeof toMeasuredBlocks>;
  layoutOpts: LayoutOptions;
  hasFootnotes: boolean;
  footnoteRefs: ReturnType<typeof collectFootnoteRefs>;
  blocks: LayoutBlock[];
  measures: BlockExtent[];
  headerContentForRender: HeaderFooterContent | undefined;
  footerContentForRender: HeaderFooterContent | undefined;
  firstPageHeaderForRender: HeaderFooterContent | undefined;
  firstPageFooterForRender: HeaderFooterContent | undefined;
  evenPageHeader: HeaderFooter | undefined;
  evenPageFooter: HeaderFooter | undefined;
  evenPageHeaderForRender: HeaderFooterContent | undefined;
  evenPageFooterForRender: HeaderFooterContent | undefined;
  evenAndOddHeaders: boolean;
  hasTitlePg: boolean;
  watermark: Watermark | undefined;
  inputs: ComputeLayoutInputs;
}

/**
 * The kernel inputs behind each produced `Layout`, keyed by the layout object
 * itself, so the display-list build can recover the exact `{ measured,
 * options }` envelope without widening the public `LayoutComputation`
 * shape. WeakMap: entries die with their layouts.
 */
const kernelInputsByLayout = new WeakMap<
  Layout,
  {
    measured: ReturnType<typeof toMeasuredBlocks>;
    options: LayoutOptions;
    headersFooters?: DisplayListHeadersFooters;
  }
>();

function rememberKernelInputs(prepared: PreparedLayoutCompute, layout: Layout): Layout {
  orchestratePageMetadata(prepared, layout);
  kernelInputsByLayout.set(layout, {
    measured: prepared.measured,
    options: prepared.layoutOpts,
    headersFooters: buildDisplayListHeadersFooters(prepared, layout.pages),
  });
  return layout;
}

/**
 * Per-page resolved widths of the PAGE/NUMPAGES field runs in one HF part (F2).
 * A centered/right footer line is measured ONCE at the field's fallback text
 * ("1"), so the Rust builder would hold the same centered position on every page.
 * Supplying the resolved-text width per layout page lets it re-center — the DOM
 * painter gets this for free because the browser lays out the real digits under
 * `text-align: center`. Only PAGE/NUMPAGES vary per page; DATE/TIME/OTHER resolve
 * to a constant fallback and are left on the char-distributed path. The `style`
 * mirrors the engine's field-run measurement so `fallbackWidth` equals the
 * width already baked into `line.width`.
 */
function computeHfFieldWidths(
  content: HeaderFooterContent,
  pages: Layout['pages']
): DisplayListFieldWidths[] | undefined {
  const totalPages = String(pages.length);
  const out: DisplayListFieldWidths[] = [];
  // Sole path: the Rust source active during this pass. Inactive (no measure
  // pass ran) or unready chains → skip the field's per-page widths; the
  // builder keeps its char-distributed fallback for that field.
  const measureFieldText = (text: string, style: FontStyle): number | undefined => {
    const rust = measureTextWidthWithActiveRustSource(text, style);
    return rust.active ? rust.width : undefined;
  };
  for (const block of content.blocks) {
    if (block.kind !== 'paragraph') continue;
    for (const run of block.runs) {
      if (run.kind !== 'field') continue;
      if (run.fieldType !== 'PAGE' && run.fieldType !== 'NUMPAGES') continue;
      if (run.pmStart === undefined) continue;
      const style: FontStyle = {
        fontFamily: run.fontFamily ?? 'Calibri',
        fontSize: run.fontSize ?? 11,
        bold: run.bold,
        italic: run.italic,
      };
      const fallbackWidth = measureFieldText(run.fallback || '1', style);
      if (fallbackWidth === undefined) continue;
      const perPage: number[] = [];
      let allPagesMeasured = true;
      for (const page of pages) {
        const width = measureFieldText(
          run.fieldType === 'NUMPAGES' ? totalPages : (page.pageLabel ?? String(page.number)),
          style
        );
        if (width === undefined) {
          allPagesMeasured = false;
          break;
        }
        perPage.push(width);
      }
      if (!allPagesMeasured) continue;
      out.push({ pmStart: run.pmStart, fallbackWidth, perPage });
    }
  }
  return out.length > 0 ? out : undefined;
}

/**
 * Assemble the display-list builder's `headersFooters` envelope field from the
 * exact values the DOM painter renders with: the `HeaderFooterContent` objects
 * `prepareLayoutCompute` already converted (via `convertHeaderFooterPmDocToContent`
 * when the persistent HF PM is mounted), the section's titlePg /
 * evenAndOddHeaders flags, and the section header/footer distances. Pure
 * repackaging — no geometry decisions happen here; the Rust builder ports the
 * painter's band math. Returns undefined when the document has no headers or
 * footers so body-only envelopes stay byte-identical.
 */
function buildDisplayListHeadersFooters(
  prepared: PreparedLayoutCompute,
  pages: Layout['pages']
): DisplayListHeadersFooters | undefined {
  const { inputs, hasTitlePg } = prepared;
  const pkg = inputs.document?.package;

  // relationship id of an HF part: identity search in the package maps, the
  // same association getHfPmDoc uses to route to the persistent HF PM
  const findRid = (
    bag: Map<string, HeaderFooter> | undefined,
    hf: HeaderFooter | null | undefined
  ): string | undefined => {
    if (!bag || !hf) return undefined;
    for (const [rId, value] of bag) {
      if (value === hf) return rId;
    }
    return undefined;
  };

  const variants: DisplayListHfVariant[] = [];
  const pushVariant = (
    kind: 'header' | 'footer',
    type: 'default' | 'first' | 'even',
    hf: HeaderFooter | null | undefined,
    content: HeaderFooterContent | undefined
  ): void => {
    if (!content || content.blocks.length === 0) return;
    const fieldWidths = computeHfFieldWidths(content, pages);
    variants.push({
      rId: findRid(kind === 'header' ? pkg?.headers : pkg?.footers, hf) ?? '',
      kind,
      type,
      measured: toMeasuredBlocks(content.blocks, content.measures),
      height: content.height,
      ...(content.flowHeight !== undefined ? { flowHeight: content.flowHeight } : {}),
      ...(content.visualTop !== undefined ? { visualTop: content.visualTop } : {}),
      ...(content.visualBottom !== undefined ? { visualBottom: content.visualBottom } : {}),
      ...(fieldWidths ? { fieldWidths } : {}),
    });
  };

  pushVariant('header', 'default', inputs.headerContent, prepared.headerContentForRender);
  pushVariant('footer', 'default', inputs.footerContent, prepared.footerContentForRender);
  pushVariant('header', 'first', inputs.firstPageHeaderContent, prepared.firstPageHeaderForRender);
  pushVariant('footer', 'first', inputs.firstPageFooterContent, prepared.firstPageFooterForRender);
  pushVariant('header', 'even', prepared.evenPageHeader, prepared.evenPageHeaderForRender);
  pushVariant('footer', 'even', prepared.evenPageFooter, prepared.evenPageFooterForRender);
  const watermark = prepared.watermark;
  if (variants.length === 0 && !watermark) return undefined;

  const { sectionProperties } = inputs;
  return {
    titlePg: hasTitlePg,
    evenAndOddHeaders: prepared.evenAndOddHeaders,
    // nullish, not truthy — an explicit w:header="0" must keep 0 (#740); when
    // absent the Rust side falls back to page.margins.header ?? 48, matching
    // the painter
    ...(sectionProperties?.headerDistance != null
      ? { headerDistance: twipsToPixels(sectionProperties.headerDistance) }
      : {}),
    ...(sectionProperties?.footerDistance != null
      ? { footerDistance: twipsToPixels(sectionProperties.footerDistance) }
      : {}),
    ...(watermark ? { watermark } : {}),
    variants,
  };
}

/**
 * The `{ measured, options }` pair a Layout was paginated from, when that
 * Layout came out of a compute call in this session, plus the `headersFooters`
 * payload assembled from the same HF contents/page-level watermark the pages
 * render with. The canvas renderer joins this with the Layout itself into the
 * `{ measured, options, layout, headersFooters }` envelope the Rust
 * display-list builder consumes. Every compute path (including footnote
 * stabilization) records its final layout here; undefined means the Layout
 * did not come from a compute call in this session.
 */
export function getLayoutKernelInputs(layout: Layout):
  | {
      measured: ReturnType<typeof toMeasuredBlocks>;
      options: unknown;
      headersFooters?: DisplayListHeadersFooters;
    }
  | undefined {
  return kernelInputsByLayout.get(layout);
}

/**
 * Run the pure layout compute pass (the 6 steps in this file's header).
 * Every pagination pass runs through the injected Rust pagination source —
 * the sole layout kernel. The wasm engine is synchronous native code, so
 * edits and opens alike take this one full-pass path (the Rust engine keeps
 * no resume checkpoints; incremental resume and the off-thread worker died
 * with the TS kernel).
 */
export function computeLayout(inputs: ComputeLayoutInputs): LayoutComputation {
  const prepared = prepareLayoutCompute(inputs);
  const { layout, footnotesByPage } = prepared.hasFootnotes
    ? paginateWithFootnotes(prepared)
    : {
        layout: inputs.paginationSource.paginate(prepared.measured, prepared.layoutOpts),
        footnotesByPage: undefined,
      };
  rememberKernelInputs(prepared, layout);
  return finishLayoutComputation(prepared, layout, footnotesByPage);
}

function prepareLayoutCompute(inputs: ComputeLayoutInputs): PreparedLayoutCompute {
  beginRustMeasureLayoutPass();
  const {
    state,
    document,
    pageSize,
    margins,
    columns,
    finalPageSize,
    finalMargins,
    finalColumns,
    pageGap,
    contentWidth,
    theme,
    styles,
    sectionProperties,
    finalSectionProperties,
    headerContent,
    footerContent,
    firstPageHeaderContent,
    firstPageFooterContent,
    measureBlocks,
  } = inputs;

  // Step 1: PM doc → flow blocks.
  const pageContentHeight = pageSize.h - margins.top - margins.bottom;
  const blocks =
    inputs.yrsBodyBlocks?.({ pageContentHeight }) ??
    (state
      ? toLayoutBlocks(state.doc as EditorTreeNode, { theme, pageContentHeight })
      : (() => {
          throw new Error('computeLayout requires either yrsBodyBlocks or a ProseMirror state');
        })());

  // Step 2: BlockExtent all blocks (per-section widths; full measure for float context).
  const blockWidths = computePerBlockWidths(
    blocks,
    { pageSize, margins, columns },
    { pageSize: finalPageSize, margins: finalMargins, columns: finalColumns }
  );

  // Step 1.5: Demote full-width "floating" tables to inline. A positioned table
  // that leaves no room for text to wrap beside it (a common full-width contract
  // form table) is block-like in Word/Google Docs — it paginates across pages.
  // Our floating path instead paints it as one overflowing fragment AND makes
  // the next paragraph skip past the whole table height (a wrap zone), stranding
  // it off-page. Clearing `floating` here — before measure and layout — routes
  // it through `layoutTable` (which breaks rows across pages) and suppresses the
  // wrap zone. Purely a layout transform on the ephemeral FlowBlocks; the PM doc
  // and the saved DOCX keep the original floating table.
  demoteBlockLikeFloatingTables(blocks, blockWidths, contentWidth);

  const measures = measureBlocks(
    blocks,
    blockWidths,
    pageGeometryFromPage({ size: pageSize, margins })
  );

  // Step 2.5: Footnote references.
  const footnoteRefs = collectFootnoteRefs(blocks);
  const hasFootnotes = footnoteRefs.some((ref) =>
    ref.noteKind === 'endnote'
      ? !!document?.package?.endnotes?.length
      : !!document?.package?.footnotes?.length
  );

  // Step 2.75: Header/footer content for rendering (needed before layout to
  // compute effective margins when HF content exceeds available space).
  const hfMetricsHeader = { section: 'header' as const, pageSize, margins };
  const hfMetricsFooter = { section: 'footer' as const, pageSize, margins };
  const defaultTabStopTwips =
    (state?.doc.attrs?.defaultTabStopTwips as number | null | undefined) ??
    document?.package.settings?.defaultTabStop ??
    null;
  const hfOptions = {
    styles,
    theme,
    measureBlocks,
    defaultTabStopTwips,
    yrsStoryBlocks: inputs.yrsStoryBlocks,
  };

  const findHfStoryId = (hf: HeaderFooter): string | undefined => {
    const pkg = document?.package;
    if (!pkg) return undefined;
    for (const bag of [pkg.headers, pkg.footers]) {
      if (!bag) continue;
      for (const [rId, value] of bag) {
        if (value === hf) return `hf:${rId}`;
      }
    }
    return undefined;
  };

  // HF unification phase 1: prefer the persistent PM doc when mounted.
  const convertHf = (
    hf: HeaderFooter | null | undefined,
    metrics: typeof hfMetricsHeader | typeof hfMetricsFooter
  ): HeaderFooterContent | undefined => {
    if (!hf) return undefined;
    const hfStoryId = inputs.yrsStoryBlocks ? findHfStoryId(hf) : undefined;
    return convertHeaderFooterToContent(hf, contentWidth, metrics, hfOptions, hfStoryId);
  };

  const headerContentForRender = convertHf(headerContent, hfMetricsHeader);
  const footerContentForRender = convertHf(footerContent, hfMetricsFooter);
  const hasTitlePg = sectionProperties?.titlePg === true;
  const firstPageHeaderForRender = hasTitlePg
    ? convertHf(firstPageHeaderContent, hfMetricsHeader)
    : undefined;
  const firstPageFooterForRender = hasTitlePg
    ? convertHf(firstPageFooterContent, hfMetricsFooter)
    : undefined;
  const sectionList = orderedSectionProperties(document);
  const inputSectionIndex = Math.max(
    0,
    sectionList.findIndex((candidate) => candidate === sectionProperties)
  );
  const effectiveRefs = effectiveHeaderFooterRefs(sectionList, inputSectionIndex);
  const evenPageHeader = effectiveRefs?.headerEven
    ? document?.package?.headers?.get(effectiveRefs.headerEven)
    : undefined;
  const evenPageFooter = effectiveRefs?.footerEven
    ? document?.package?.footers?.get(effectiveRefs.footerEven)
    : undefined;
  const evenPageHeaderForRender = convertHf(evenPageHeader, hfMetricsHeader);
  const evenPageFooterForRender = convertHf(evenPageFooter, hfMetricsFooter);
  const settingsWithEvenOdd = document?.package?.settings as
    | { evenAndOddHeaders?: boolean }
    | undefined;
  const evenAndOddHeaders =
    settingsWithEvenOdd?.evenAndOddHeaders === true ||
    sectionProperties?.evenAndOddHeaders === true;

  // Watermark rides PM state as a doc attr (so it's undoable).
  const watermark =
    (state?.doc.attrs?.watermark as Watermark | null | undefined) ??
    finalSectionProperties?.watermark ??
    undefined;

  // Margin extension — push body clear of the header/footer bands (Word grows
  // the band when in-flow content exceeds the authored margin). Shared core
  // helper: uses in-flow `flowHeight` so page/margin-anchored floats (e.g. a
  // letterhead) don't push the body (issue #705), with a content-area clamp;
  // mutates each `sectionBreak.margins` in place.
  const { margins: effectiveMargins, finalMargins: effectiveFinalMargins } =
    extendMarginsForHeaderFooter({
      pageSize,
      margins,
      finalMargins,
      bodyBlocks: blocks,
      headers: [headerContentForRender, firstPageHeaderForRender, evenPageHeaderForRender],
      footers: [footerContentForRender, firstPageFooterForRender, evenPageFooterForRender],
      warn: (msg) => console.warn(`[computeLayout] ${msg}`),
    });

  // Step 3 inputs: pagination options (the paginate step itself runs in
  // `computeLayout`).
  const bodyBreakType = finalSectionProperties?.sectionStart as
    | 'continuous'
    | 'nextPage'
    | 'evenPage'
    | 'oddPage'
    | undefined;
  const layoutOpts = {
    pageSize,
    margins: effectiveMargins,
    finalPageSize,
    finalMargins: effectiveFinalMargins,
    columns: finalColumns,
    bodyBreakType,
    pageGap,
  };

  return {
    measured: toMeasuredBlocks(blocks, measures),
    layoutOpts,
    hasFootnotes,
    footnoteRefs,
    blocks,
    measures,
    headerContentForRender,
    footerContentForRender,
    firstPageHeaderForRender,
    firstPageFooterForRender,
    evenPageHeader,
    evenPageFooter,
    evenPageHeaderForRender,
    evenPageFooterForRender,
    evenAndOddHeaders,
    hasTitlePg,
    watermark,
    inputs,
  };
}

/** Two-pass footnote pagination (measure footnotes, stabilize reserved space). */
function paginateWithFootnotes(prepared: PreparedLayoutCompute): {
  layout: Layout;
  footnotesByPage: Map<number, FootnoteRenderItem[]> | undefined;
} {
  const { measured, layoutOpts, footnoteRefs, blocks, measures, inputs } = prepared;
  const { document, contentWidth, styles, theme, measureBlocks, state } = inputs;
  const defaultTabStopTwips =
    (state?.doc.attrs?.defaultTabStopTwips as number | null | undefined) ??
    document?.package.settings?.defaultTabStop ??
    null;

  const paginate = (m: typeof measured, o: LayoutOptions): Layout =>
    inputs.paginationSource.paginate(m, o);
  const pass1Layout = paginate(measured, layoutOpts);
  // Resolve the authored footnote columns against the page/section that owns
  // each reference. Batch D populates sectionIndex; legacy layouts retain the
  // previous document-wide fallback.
  const initialPageFootnotes = mapFootnotesToPages(pass1Layout.pages, footnoteRefs);
  const footnoteColumnsByPage = new Map<number, number>();
  const footnoteWidthById = new Map<number, number>();
  for (const page of pass1Layout.pages) {
    const { columns, columnWidth } = resolveFootnoteColumnLayout(document, page, contentWidth);
    footnoteColumnsByPage.set(page.number, columns);
    for (const id of initialPageFootnotes.get(page.number) ?? []) {
      footnoteWidthById.set(id, columnWidth);
    }
  }
  const notePresentations = buildNotePresentations(
    footnoteRefs,
    pass1Layout.pages,
    document,
    initialPageFootnotes
  );
  const footnoteContentMap: Map<number, FootnoteContent> = buildNoteContentMap(
    document!.package.footnotes ?? [],
    document!.package.endnotes ?? [],
    footnoteRefs,
    (ref) => footnoteWidthById.get(noteReferenceMapId(ref)) ?? contentWidth,
    {
      styles: styles ?? undefined,
      theme: theme ?? null,
      measureBlocks,
      defaultTabStopTwips,
      yrsStoryBlocks: inputs.yrsStoryBlocks,
    },
    notePresentations
  );
  const stabilized = stabilizeFootnoteLayout({
    blocks,
    measures,
    layoutOpts,
    paginate,
    footnoteRefs,
    footnoteContentMap,
    initialLayout: pass1Layout,
    footnoteColumns: footnoteColumnsByPage,
    mapReferencesToPages: (pages, refs) => mapNotesToPages(pages, refs, document),
  });
  attachTypedNoteAreas(stabilized.layout, stabilized.pageFootnoteMap, footnoteContentMap, document);
  return {
    layout: stabilized.layout,
    footnotesByPage: buildFootnoteRenderItems(
      stabilized.pageFootnoteMap,
      footnoteContentMap,
      document
    ),
  };
}

type NoteProperties = FootnoteProperties | EndnoteProperties;

function notePropertiesForPage(
  document: Document | null,
  page: Layout['pages'][number],
  kind: NoteKind
): NoteProperties {
  // Batch B adds document-wide settings parsing; keep the read optional so
  // Batch E remains compatible with older parsed packages.
  const settings = document?.package?.settings as
    | { footnotePr?: FootnoteProperties; endnotePr?: EndnoteProperties }
    | undefined;
  const section = sectionPropertiesForPage(document, page);
  return kind === 'endnote'
    ? { ...(settings?.endnotePr ?? {}), ...(section?.endnotePr ?? {}) }
    : { ...(settings?.footnotePr ?? {}), ...(section?.footnotePr ?? {}) };
}

function anchorPageForMapId(
  pages: Layout['pages'],
  anchorMap: Map<number, number[]>,
  mapId: number
): Layout['pages'][number] | undefined {
  const pageNumber = [...anchorMap].find(([, ids]) => ids.includes(mapId))?.[0];
  return pages.find((page) => page.number === pageNumber);
}

/** Apply note placement after resolving the page that owns each body anchor. */
function mapNotesToPages(
  pages: Layout['pages'],
  refs: ReturnType<typeof collectFootnoteRefs>,
  document: Document | null
): Map<number, number[]> {
  const anchors = mapFootnotesToPages(pages, refs);
  const result = new Map<number, number[]>();
  const append = (pageNumber: number, mapId: number): void => {
    const ids = result.get(pageNumber) ?? [];
    if (!ids.includes(mapId)) ids.push(mapId);
    result.set(pageNumber, ids);
  };

  for (const ref of refs) {
    const mapId = noteReferenceMapId(ref);
    const anchorPage = anchorPageForMapId(pages, anchors, mapId);
    if (!anchorPage) continue;
    const kind = ref.noteKind ?? 'footnote';
    const properties = notePropertiesForPage(document, anchorPage, kind);
    const placement =
      properties.position ?? (kind === 'endnote' ? ('docEnd' as const) : ('pageBottom' as const));
    let targetPage = anchorPage;
    if (placement === 'docEnd') {
      targetPage = pages[pages.length - 1] ?? anchorPage;
    } else if (placement === 'sectEnd') {
      const sameSection = pages.filter((page) =>
        anchorPage.sectionId !== undefined
          ? page.sectionId === anchorPage.sectionId
          : anchorPage.sectionIndex !== undefined
            ? page.sectionIndex === anchorPage.sectionIndex
            : true
      );
      targetPage = sameSection[sameSection.length - 1] ?? anchorPage;
    }
    append(targetPage.number, mapId);
  }
  return result;
}

/** Resolve starts/restarts/formats without evaluating any field instruction. */
function buildNotePresentations(
  refs: ReturnType<typeof collectFootnoteRefs>,
  pages: Layout['pages'],
  document: Document | null,
  anchorMap: Map<number, number[]>
): Map<number, NotePresentation> {
  const result = new Map<number, NotePresentation>();
  const counters = new Map<string, number>();
  for (const ref of refs) {
    const mapId = noteReferenceMapId(ref);
    if (result.has(mapId)) continue;
    const page = anchorPageForMapId(pages, anchorMap, mapId);
    if (!page) continue;
    const kind = ref.noteKind ?? 'footnote';
    const properties = notePropertiesForPage(document, page, kind);
    const restartKey =
      properties.numRestart === 'eachPage'
        ? `${kind}:page:${page.number}`
        : properties.numRestart === 'eachSect'
          ? `${kind}:section:${page.sectionId ?? page.sectionIndex ?? 0}`
          : `${kind}:continuous`;
    const displayNumber = counters.get(restartKey) ?? properties.numStart ?? 1;
    counters.set(restartKey, displayNumber + 1);
    result.set(mapId, {
      displayNumber,
      displayLabel: formatNumber(displayNumber, (properties.numFmt ?? 'decimal') as NumberFormat),
      anchor: { docStart: ref.pmPos, docEnd: ref.pmPos + 1 },
    });
  }
  return result;
}

/**
 * Populate Batch A's renderer-neutral note regions. The DOM painter consumes
 * `y`/`placement`; Batch F can lower the same measured payload into display
 * primitives without reconstructing notes from document state.
 */
function attachTypedNoteAreas(
  layout: Layout,
  pageFootnoteMap: Map<number, number[]>,
  footnoteContentMap: Map<number, FootnoteContent>,
  document: Document | null
): void {
  for (const page of layout.pages) {
    const ids = pageFootnoteMap.get(page.number);
    if (!ids?.length) continue;
    const contents = ids
      .map((id) => footnoteContentMap.get(id))
      .filter((content): content is FootnoteContent => content !== undefined);
    if (!contents.length) continue;

    const contentBottom = page.size.h - page.margins.bottom;
    const lastBodyBottom = page.fragments.reduce(
      (bottom, fragment) => Math.max(bottom, fragment.y + fragment.height),
      page.margins.top
    );
    const byKind = new Map<NoteKind, FootnoteContent[]>();
    for (const content of contents) {
      const kind = content.noteKind ?? 'footnote';
      const group = byKind.get(kind) ?? [];
      group.push(content);
      byKind.set(kind, group);
    }

    let bottomCursor = contentBottom;
    let beneathTextCursor = lastBodyBottom;
    page.noteAreas = [...byKind].map(([kind, group]) => {
      const properties = notePropertiesForPage(document, page, kind);
      const placement =
        properties.position ?? (kind === 'endnote' ? ('docEnd' as const) : ('pageBottom' as const));
      const columns = kind === 'footnote' ? (page.footnoteColumns ?? 1) : 1;
      const partitions = distributeFootnotesIntoColumns(group, columns);
      const height =
        partitions.reduce(
          (max, partition) =>
            Math.max(
              max,
              partition.reduce((sum, content) => sum + content.height, 0)
            ),
          0
        ) + FOOTNOTE_SEPARATOR_HEIGHT;
      let y: number;
      if (placement === 'beneathText') {
        y = Math.min(beneathTextCursor, contentBottom - height);
        beneathTextCursor = y + height;
      } else {
        bottomCursor -= height;
        y = bottomCursor;
      }
      return {
        kind,
        placement,
        y,
        height,
        columns,
        sectionId: page.sectionId,
        notes: group.map((content) => ({
          kind,
          id: content.id,
          displayLabel: content.displayLabel ?? String(content.displayNumber),
          blocks: content.blocks,
          measures: content.measures,
          height: content.height,
          anchorDocStart: content.anchor?.docStart,
          anchorDocEnd: content.anchor?.docEnd,
          customMarkFollows: content.customMarkFollows,
        })),
      };
    });
  }
}

function finishLayoutComputation(
  prepared: PreparedLayoutCompute,
  layout: Layout,
  footnotesByPage: Map<number, FootnoteRenderItem[]> | undefined
): LayoutComputation {
  const { sectionProperties } = prepared.inputs;
  return {
    blocks: prepared.blocks,
    measures: prepared.measures,
    layout,
    headerContentForRender: prepared.headerContentForRender,
    footerContentForRender: prepared.footerContentForRender,
    firstPageHeaderForRender: prepared.firstPageHeaderForRender,
    firstPageFooterForRender: prepared.firstPageFooterForRender,
    hasTitlePg: prepared.hasTitlePg,
    watermark: prepared.watermark,
    // Nullish, not truthy: an explicit `w:header="0"` must paint the header at
    // the page top, not fall back to the painter's 0.5in default (#740).
    headerDistancePx:
      sectionProperties?.headerDistance != null
        ? twipsToPixels(sectionProperties.headerDistance)
        : undefined,
    footerDistancePx:
      sectionProperties?.footerDistance != null
        ? twipsToPixels(sectionProperties.footerDistance)
        : undefined,
    pageBorders: sectionProperties?.pageBorders,
    footnotesByPage: footnotesByPage?.size ? footnotesByPage : undefined,
  };
}
