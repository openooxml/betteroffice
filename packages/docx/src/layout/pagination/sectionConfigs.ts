/**
 * Per-section geometry collection.
 *
 * Walks a block list once and resolves the effective page geometry of every
 * section (size, margins, columns, header/footer refs, numbering, borders,
 * …), inheriting from the previous section where a break leaves a property
 * unset and applying any upstream section contracts. The live consumer is
 * `layout/regions/sectionGeometry.ts` (per-section measurement widths); the
 * Rust engine mirrors the same resolution natively.
 */

import type {
  LayoutBlock,
  PageMargins,
  ColumnLayout,
  SectionBreakBlock,
  PageHeaderFooterRefs,
  PageNumberingContract,
  PageBorderContract,
  NoteSettingsContract,
  LayoutOptions,
} from './types';

/**
 * Page-flow geometry resolved from a single section's properties.
 * Exported so the React paged editor can reuse the same shape when
 * measuring blocks per section width — keeping pagination and
 * measurement consistent.
 */
export type SectionLayoutConfig = {
  pageSize: { w: number; h: number };
  margins: PageMargins;
  /** Optional. Sections without explicit columns inherit `{ count: 1 }`. */
  columns?: ColumnLayout;
  /** Stable section/page state copied onto every page created in this section. */
  sectionIndex?: number;
  sectionId?: string;
  orientation?: 'portrait' | 'landscape';
  headerFooterRefs?: PageHeaderFooterRefs;
  pageNumbering?: PageNumberingContract;
  pageBorders?: PageBorderContract;
  watermark?: SectionBreakBlock['watermark'];
  verticalAlign?: SectionBreakBlock['verticalAlign'];
  noteSettings?: NoteSettingsContract;
  /** Versioned page metadata is emitted only when an upstream contract opts in. */
  emitPageMetadata?: boolean;
};

/**
 * Walk `blocks` once and collect per-section geometry. `configs` has one
 * entry per section break plus a trailing `finalConfig`. `breakIndices` is
 * 1-to-1 with the inner break entries (same length as `configs.length - 1`).
 * Callers that need the break `type` can read it from
 * `(blocks[breakIndices[i]] as SectionBreakBlock).type`.
 *
 * @internal
 */
export function collectSectionConfigs(
  blocks: LayoutBlock[],
  initialConfig: SectionLayoutConfig,
  finalConfig: SectionLayoutConfig,
  sectionContracts?: LayoutOptions['sections']
): {
  configs: SectionLayoutConfig[];
  breakIndices: number[];
} {
  const configs: SectionLayoutConfig[] = [];
  const breakIndices: number[] = [];
  let previousConfig = initialConfig;
  let sectionIndex = initialConfig.sectionIndex ?? 0;
  for (let i = 0; i < blocks.length; i++) {
    if (blocks[i].kind !== 'sectionBreak') continue;
    const sb = blocks[i] as SectionBreakBlock;
    const contract = sectionContracts?.[sectionIndex];
    const config: SectionLayoutConfig = {
      pageSize: sb.pageSize ?? previousConfig.pageSize,
      margins: sb.margins ?? previousConfig.margins,
      columns: sb.columns,
      sectionIndex: sb.sectionIndex ?? sectionIndex,
      sectionId:
        sb.sectionId ??
        previousConfig.sectionId ??
        (previousConfig.emitPageMetadata ? `section-${sectionIndex}` : undefined),
      orientation: sb.orientation ?? previousConfig.orientation,
      headerFooterRefs: sb.headerFooterRefs ?? previousConfig.headerFooterRefs,
      pageNumbering: sb.pageNumbering ?? previousConfig.pageNumbering,
      pageBorders: sb.pageBorders ?? previousConfig.pageBorders,
      watermark: sb.watermark ?? previousConfig.watermark,
      verticalAlign: sb.verticalAlign ?? previousConfig.verticalAlign,
      noteSettings: sb.noteSettings ?? previousConfig.noteSettings,
      emitPageMetadata:
        previousConfig.emitPageMetadata ||
        sb.sectionId != null ||
        sb.sectionIndex != null ||
        sb.pageNumbering != null ||
        sb.headerFooterRefs != null,
      ...(contract
        ? {
            pageSize: {
              w: contract.pageSize?.w ?? sb.pageSize?.w ?? previousConfig.pageSize.w,
              h: contract.pageSize?.h ?? sb.pageSize?.h ?? previousConfig.pageSize.h,
            },
            margins: { ...previousConfig.margins, ...sb.margins, ...contract.margins },
            columns: contract.columns ?? sb.columns,
            sectionId:
              contract.sectionId ??
              sb.sectionId ??
              (previousConfig.emitPageMetadata ? `section-${sectionIndex}` : undefined),
            headerFooterRefs: contract.headerFooterRefs ?? sb.headerFooterRefs,
            pageNumbering: contract.pageNumbering ?? sb.pageNumbering,
            pageBorders: contract.pageBorders ?? sb.pageBorders,
            watermark: contract.watermark ?? sb.watermark,
            noteSettings: contract.noteSettings ?? sb.noteSettings,
          }
        : {}),
    };
    configs.push(config);
    breakIndices.push(i);
    previousConfig = config;
    sectionIndex += 1;
  }
  const finalContract = sectionContracts?.[sectionIndex];
  configs.push({
    ...finalConfig,
    ...(finalContract
      ? {
          pageSize: {
            w: finalContract.pageSize?.w ?? finalConfig.pageSize.w,
            h: finalContract.pageSize?.h ?? finalConfig.pageSize.h,
          },
          margins: { ...finalConfig.margins, ...finalContract.margins },
          columns: finalContract.columns ?? finalConfig.columns,
          headerFooterRefs: finalContract.headerFooterRefs ?? finalConfig.headerFooterRefs,
          pageNumbering: finalContract.pageNumbering ?? finalConfig.pageNumbering,
          pageBorders: finalContract.pageBorders ?? finalConfig.pageBorders,
          watermark: finalContract.watermark ?? finalConfig.watermark,
          noteSettings: finalContract.noteSettings ?? finalConfig.noteSettings,
        }
      : {}),
    sectionIndex: finalConfig.sectionIndex ?? sectionIndex,
    sectionId:
      finalContract?.sectionId ??
      finalConfig.sectionId ??
      (finalConfig.emitPageMetadata ? `section-${sectionIndex}` : undefined),
  });
  return { configs, breakIndices };
}
