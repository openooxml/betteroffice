/**
 * Style Resolver for ProseMirror Editor
 *
 * Resolves OOXML style definitions to final paragraph and run properties.
 * Handles the cascade:
 * 1. Document defaults (docDefaults)
 * 2. Normal style (if no explicit styleId)
 * 3. Style chain (basedOn inheritance - already resolved by styleParser)
 * 4. Inline properties
 *
 * Based on ECMA-376 style cascade rules.
 */

import type {
  StyleDefinitions,
  Style,
  DocDefaults,
  ParagraphFormatting,
  TextFormatting,
} from '../types/document';
import { mergeTextFormatting } from '../utils/textFormattingMerge';

/**
 * Resolved style properties ready for rendering
 */
export interface ResolvedParagraphStyle {
  /** Paragraph formatting (alignment, spacing, indentation, etc.) */
  paragraphFormatting?: ParagraphFormatting;
  /** Default run formatting from the style */
  runFormatting?: TextFormatting;
}

/**
 * Word's default-template Normal style, used as a last-resort fallback for a
 * bare document that ships neither a `Normal` style nor `w:docDefaults`. These
 * are the values Word writes into a freshly created document's Normal: 8pt
 * (160 twips) after spacing, 1.08x line spacing.
 *
 * It is NOT applied when the document provides `w:docDefaults`. Those defaults
 * are the authoritative paragraph defaults per ECMA-376 §17.7.5; an empty
 * `w:pPrDefault` means zero spacing / single line, and synthesizing this style
 * over it wrongly inflates style-less paragraphs (e.g. table cells whose author
 * relied on the empty default).
 */
const BUILTIN_NORMAL_STYLE: Style = {
  styleId: 'Normal',
  type: 'paragraph',
  name: 'Normal',
  default: true,
  pPr: {
    spaceAfter: 160,
    lineSpacing: 259,
    lineSpacingRule: 'auto',
  },
};

/**
 * StyleResolver provides efficient access to resolved style properties
 */
export class StyleResolver {
  private readonly stylesById: Map<string, Style>;
  private readonly docDefaults: DocDefaults | undefined;
  private readonly defaultParagraphStyle: Style | undefined;
  private readonly defaultTableStyle: Style | undefined;
  private readonly defaultCharacterStyle: Style | undefined;

  constructor(styleDefinitions: StyleDefinitions | undefined) {
    this.stylesById = new Map();
    this.docDefaults = styleDefinitions?.docDefaults;

    // Build lookup map
    if (styleDefinitions?.styles) {
      for (const style of styleDefinitions.styles) {
        if (style.styleId) {
          this.stylesById.set(style.styleId, style);
        }
      }
    }

    // Find defaults — one per type per ECMA-376 §17.7.4.18.
    this.defaultParagraphStyle = this.findDefaultStyle('paragraph');
    this.defaultTableStyle = this.findDefaultStyle('table');
    this.defaultCharacterStyle = this.findDefaultStyle('character');
  }

  /**
   * Get a style by ID
   */
  getStyle(styleId: string): Style | undefined {
    return this.stylesById.get(styleId);
  }

  /**
   * Get all available paragraph styles (for toolbar dropdown)
   */
  getParagraphStyles(): Style[] {
    const styles: Style[] = [];
    for (const style of this.stylesById.values()) {
      if (style.type === 'paragraph' && !style.hidden && !style.semiHidden) {
        styles.push(style);
      }
    }
    // Sort by uiPriority, then by name
    return styles.sort((a, b) => {
      const priorityA = a.uiPriority ?? 99;
      const priorityB = b.uiPriority ?? 99;
      if (priorityA !== priorityB) return priorityA - priorityB;
      return (a.name ?? a.styleId).localeCompare(b.name ?? b.styleId);
    });
  }

  /**
   * Whether a paragraph style with the given id is defined in the
   * document's `styles.xml`. Used by the agent toolkit to refuse
   * `set_paragraph_style({ styleId: 'NoSuchStyle' })` instead of
   * silently writing an invalid `<w:pStyle>` reference.
   */
  hasParagraphStyle(styleId: string): boolean {
    const style = this.stylesById.get(styleId);
    return style?.type === 'paragraph';
  }

  /**
   * Resolve paragraph style properties, including docDefaults cascade
   *
   * @param styleId - The style ID to resolve (e.g., 'Heading1', 'Normal')
   * @returns Resolved paragraph and run formatting
   */
  resolveParagraphStyle(styleId: string | undefined | null): ResolvedParagraphStyle {
    const result: ResolvedParagraphStyle = {};

    // Start with document defaults
    if (this.docDefaults?.pPr) {
      result.paragraphFormatting = { ...this.docDefaults.pPr };
    }
    if (this.docDefaults?.rPr) {
      result.runFormatting = { ...this.docDefaults.rPr };
    }

    // If no styleId, apply Normal style (if exists)
    if (!styleId) {
      if (this.defaultParagraphStyle) {
        this.mergeStyleIntoResult(result, this.defaultParagraphStyle);
      }
      return result;
    }

    // Get the requested style (already has basedOn chain resolved by styleParser)
    const style = this.stylesById.get(styleId);
    if (!style) {
      // Style not found, fall back to Normal
      if (this.defaultParagraphStyle) {
        this.mergeStyleIntoResult(result, this.defaultParagraphStyle);
      }
      return result;
    }

    // Merge style properties into result
    this.mergeStyleIntoResult(result, style);

    return result;
  }

  /**
   * Resolve the style applied to the paragraph that follows one styled with
   * `styleId` when the user presses Enter (OOXML `w:next`, §17.7.4.10).
   *
   * Returns null when the style has no `w:next`, when `next` points back at
   * the same style (the common heading-stays-heading case is handled by the
   * caller), or when the style is unknown.
   */
  getNextStyleId(styleId: string | undefined | null): string | null {
    if (!styleId) return null;
    const next = this.stylesById.get(styleId)?.next;
    return next && next !== styleId ? next : null;
  }

  /**
   * Get all available table styles (for style gallery)
   */
  getTableStyles(): Style[] {
    const styles: Style[] = [];
    for (const style of this.stylesById.values()) {
      if (style.type === 'table' && !style.hidden && !style.semiHidden) {
        styles.push(style);
      }
    }
    return styles.sort((a, b) => {
      const priorityA = a.uiPriority ?? 99;
      const priorityB = b.uiPriority ?? 99;
      if (priorityA !== priorityB) return priorityA - priorityB;
      return (a.name ?? a.styleId).localeCompare(b.name ?? b.styleId);
    });
  }

  /**
   * Resolve run (character) style properties
   *
   * @param styleId - The character style ID to resolve
   * @returns Resolved text formatting
   */
  resolveRunStyle(styleId: string | undefined | null): TextFormatting | undefined {
    // OOXML §17.7.4.18 + §17.3.2 cascade for run formatting:
    //   1. docDefaults.rPr            (rPrDefault)
    //   2. default character style    (the style marked w:default="1")
    //   3. explicit character style   (from <w:rStyle> on the run)
    // Pre-PR this method skipped step 2, so any property set on the default
    // character style (typically "Default Paragraph Font" / "FontePadrao")
    // never reached runs without an explicit <w:rStyle>.
    let result: TextFormatting = {};
    if (this.docDefaults?.rPr) {
      result = { ...this.docDefaults.rPr };
    }

    if (this.defaultCharacterStyle?.rPr) {
      result = this.mergeTextFormatting(result, this.defaultCharacterStyle.rPr) ?? result;
    }

    if (!styleId) {
      return Object.keys(result).length > 0 ? result : undefined;
    }

    const style = this.stylesById.get(styleId);
    if (!style?.rPr) {
      return Object.keys(result).length > 0 ? result : undefined;
    }

    const merged = this.mergeTextFormatting(result, style.rPr);
    return merged && Object.keys(merged).length > 0 ? merged : undefined;
  }

  /**
   * Get a character style's own properties WITHOUT docDefaults.
   * Used when the caller already has docDefaults applied (e.g., from paragraph style resolution).
   * This prevents docDefault fonts from incorrectly overriding paragraph style fonts.
   */
  getRunStyleOwnProperties(styleId: string | undefined | null): TextFormatting | undefined {
    if (!styleId) return undefined;

    const style = this.stylesById.get(styleId);
    if (!style?.rPr) return undefined;

    return Object.keys(style.rPr).length > 0 ? { ...style.rPr } : undefined;
  }

  /**
   * Get document defaults
   */
  getDocDefaults(): DocDefaults | undefined {
    return this.docDefaults;
  }

  /**
   * Get default paragraph style (usually "Normal")
   */
  getDefaultParagraphStyle(): Style | undefined {
    return this.defaultParagraphStyle;
  }

  /**
   * Get the default table style (the one marked `w:default="1"`).
   *
   * Per ECMA-376 §17.7.4.18, tables that don't specify a `w:tblStyle`
   * inherit from this style. The styleId varies by document language
   * ("Normal Table", "TableNormal", "Tabelanormal", etc.) — find it by
   * the parsed `default` flag, not by name.
   */
  getDefaultTableStyle(): Style | undefined {
    return this.defaultTableStyle;
  }

  /**
   * Get the default character style (the one marked `w:default="1"`).
   *
   * Per ECMA-376 §17.7.4.18, runs without an explicit `w:rStyle` reference
   * inherit from this style. The styleId varies by document language
   * ("Default Paragraph Font", "FontePadrao", "Fontepargpadro" in
   * Portuguese fixtures, etc.) — find it by the parsed `default` flag.
   */
  getDefaultCharacterStyle(): Style | undefined {
    return this.defaultCharacterStyle;
  }

  /**
   * Check if a style exists
   */
  hasStyle(styleId: string): boolean {
    return this.stylesById.has(styleId);
  }

  // ============================================================================
  // Private helpers
  // ============================================================================

  private findDefaultStyle(type: 'paragraph' | 'character' | 'table'): Style | undefined {
    // First try to find explicitly marked default
    for (const style of this.stylesById.values()) {
      if (style.type === type && style.default) {
        return style;
      }
    }
    // Fall back to "Normal" for paragraph styles. Only synthesize the
    // default-template Normal for a bare document with no docDefaults; when
    // docDefaults exist they are the authoritative paragraph defaults and an
    // absent Normal must not re-introduce the template's 8pt/1.08 spacing.
    if (type === 'paragraph') {
      return this.stylesById.get('Normal') ?? (this.docDefaults ? undefined : BUILTIN_NORMAL_STYLE);
    }
    return undefined;
  }

  private mergeStyleIntoResult(result: ResolvedParagraphStyle, style: Style): void {
    if (style.pPr) {
      result.paragraphFormatting = this.mergeParagraphFormatting(
        result.paragraphFormatting,
        style.pPr
      );
    }
    if (style.rPr) {
      result.runFormatting = this.mergeTextFormatting(result.runFormatting, style.rPr);
    }
  }

  /**
   * Merge paragraph formatting (source overrides target)
   */
  private mergeParagraphFormatting(
    target: ParagraphFormatting | undefined,
    source: ParagraphFormatting | undefined
  ): ParagraphFormatting | undefined {
    if (!source) return target;
    if (!target) return source ? { ...source } : undefined;

    const result = { ...target };

    for (const key of Object.keys(source) as (keyof ParagraphFormatting)[]) {
      const value = source[key];
      if (value !== undefined) {
        if (key === 'runProperties') {
          result.runProperties = this.mergeTextFormatting(
            result.runProperties,
            source.runProperties
          );
        } else if (key === 'borders' || key === 'numPr' || key === 'frame') {
          const baseValue = result[key] as Record<string, unknown> | undefined;
          const sourceValue = value as Record<string, unknown> | undefined;
          (result as Record<string, unknown>)[key] = {
            ...(baseValue || {}),
            ...(sourceValue || {}),
          };
        } else if (key === 'tabs' && Array.isArray(value)) {
          // Tabs from higher priority source replace lower priority
          result.tabs = [...value];
        } else {
          (result as Record<string, unknown>)[key] = value;
        }
      }
    }

    return result;
  }

  private mergeTextFormatting(
    target: TextFormatting | undefined,
    source: TextFormatting | undefined
  ): TextFormatting | undefined {
    return mergeTextFormatting(target, source);
  }
}

/**
 * Create a style resolver from document's style definitions
 */
export function createStyleResolver(styleDefinitions: StyleDefinitions | undefined): StyleResolver {
  return new StyleResolver(styleDefinitions);
}
