/**
 * Typed value setters for block-level content controls — set a dropdown
 * selection, toggle a checkbox, or set a date. These produce both the visible
 * content (the run text Word shows) and the structured state inside the
 * captured raw `w:sdtPr` (dropdown `w:lastValue`, `w14:checked`, `w:date`'s
 * `w:fullDate`), patched in place so the rest of the control round-trips
 * verbatim. Use these instead of {@link setContentControlContent} for typed
 * controls, which that function refuses by design.
 *
 * Raw `w:sdtPr` is patched with targeted string edits (not a full re-serialize)
 * to preserve the `CT_SdtPr` element order and any unmodeled properties — the
 * same capture-and-replay contract used everywhere else for SDTs.
 */

import type { BlockContent, SdtProperties, Run } from '../types/document';

/** A typed value to apply to a content control. */
export type ContentControlValue =
  | { kind: 'dropdown'; value: string }
  | { kind: 'checkbox'; checked: boolean }
  | { kind: 'date'; date: string };

/** The control doesn't support the requested value kind, or the value is invalid. */
export class ContentControlValueError extends Error {
  constructor(message: string) {
    super(message);
    this.name = 'ContentControlValueError';
  }
}

// ── raw w:sdtPr string patching ─────────────────────────────────────────────

/** Read an attribute from the first `<prefix:local ...>` element in `xml`. */
function readAttr(xml: string, element: string, attr: string): string | undefined {
  const el = new RegExp(`<${element}\\b[^>]*>`).exec(xml);
  if (!el) return undefined;
  const m = new RegExp(`\\b${attr}="([^"]*)"`).exec(el[0]);
  return m ? m[1] : undefined;
}

/**
 * Set (or add) an attribute on the first `<prefix:local ...>` element in `xml`.
 * Returns the patched string; if the element isn't present, `xml` is unchanged.
 */
function setAttr(xml: string, element: string, attr: string, value: string): string {
  const tag = new RegExp(`<${element}\\b[^>]*?(/?)>`);
  const m = tag.exec(xml);
  if (!m) return xml;
  const open = m[0];
  const selfClose = m[1] === '/';
  const body = open.slice(1, selfClose ? -2 : -1); // strip "<" and "/>"/">"
  const hasAttr = new RegExp(`\\b${attr}="[^"]*"`).test(body);
  // Use a replacement *function* so `$`-sequences in `value` aren't interpreted.
  const newBody = hasAttr
    ? body.replace(new RegExp(`\\b${attr}="[^"]*"`), () => `${attr}="${value}"`)
    : `${body} ${attr}="${value}"`;
  return xml.replace(open, () => `<${newBody}${selfClose ? '/>' : '>'}`);
}

/** Escape a string for safe interpolation into an XML attribute value. */
function escapeXmlAttr(s: string): string {
  return s
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;');
}

/** Hex code-point string (e.g. "2612") → the character it denotes. */
function codePointChar(hex: string | undefined, fallback: string): string {
  if (!hex) return fallback;
  const n = parseInt(hex, 16);
  return Number.isNaN(n) || n < 0 || n > 0x10ffff || (n >= 0xd800 && n <= 0xdfff)
    ? fallback
    : String.fromCodePoint(n);
}

// ── date formatting (minimal OOXML w:dateFormat support) ────────────────────

/** Format an ISO date (yyyy-mm-dd) with a subset of OOXML date tokens. */
export function formatSdtDate(iso: string, pattern?: string, language?: string): string {
  const [y, m, d] = iso
    .slice(0, 10)
    .split('-')
    .map((n) => parseInt(n, 10));
  if (!isValidIsoDate(y, m, d)) return iso;
  const fmt = pattern && pattern.trim() ? pattern : 'M/d/yyyy';
  const pad = (n: number) => String(n).padStart(2, '0');
  const date = new Date(Date.UTC(y, m - 1, d));
  const month = (width: 'long' | 'short'): string => {
    try {
      return new Intl.DateTimeFormat(language || 'en', {
        month: width,
        timeZone: 'UTC',
      }).format(date);
    } catch {
      return new Intl.DateTimeFormat('en', { month: width, timeZone: 'UTC' }).format(date);
    }
  };
  // Single pass so an emitted month name (e.g. "March") isn't re-scanned by a
  // later, shorter token like `M` — which would corrupt it to "3arch".
  const tokens: Record<string, string> = {
    yyyy: String(y),
    yy: String(y).slice(-2),
    MMMM: month('long'),
    MMM: month('short'),
    MM: pad(m),
    M: String(m),
    dd: pad(d),
    d: String(d),
  };
  return fmt.replace(/yyyy|yy|MMMM|MMM|MM|M|dd|d/g, (t) => tokens[t]);
}

function isValidIsoDate(year: number, month: number, day: number): boolean {
  if (!year || !month || !day) return false;
  const date = new Date(Date.UTC(year, month - 1, day));
  return (
    date.getUTCFullYear() === year && date.getUTCMonth() === month - 1 && date.getUTCDate() === day
  );
}

// ── value application ───────────────────────────────────────────────────────

/** A one-run paragraph; `font` sets the run's font (for symbol glyphs). */
function paragraph(text: string, font?: string): BlockContent {
  if (!text) return { type: 'paragraph', content: [] };
  const run: Run = {
    type: 'run',
    content: [{ type: 'text', text }],
    ...(font
      ? { formatting: { fontFamily: { ascii: font, hAnsi: font, eastAsia: font, cs: font } } }
      : {}),
  };
  return { type: 'paragraph', content: [run] };
}

function clearShowingPlaceholderXml(raw: string | undefined): string | undefined {
  if (!raw) return raw;
  return raw
    .replace(/<w:showingPlcHdr\b[^>]*\/>/g, '')
    .replace(/<w:showingPlcHdr\b[^>]*>[\s\S]*?<\/w:showingPlcHdr>/g, '');
}

/** Clear a control's placeholder state (real content is being written). */
function withoutPlaceholder(props: SdtProperties, nextRaw: string): SdtProperties {
  const cleaned = clearShowingPlaceholderXml(nextRaw);
  return {
    ...props,
    showingPlaceholder: false,
    rawPropertiesXml: (cleaned ?? nextRaw) || undefined,
  };
}

/**
 * Compute the new properties + display blocks for applying a typed value, without
 * touching a document. Shared by the headless setter and the editor (PM) path.
 * Throws {@link ContentControlValueError} on a type/value mismatch.
 */
export function applyContentControlValue(
  props: SdtProperties,
  value: ContentControlValue
): { properties: SdtProperties; content: BlockContent[] } {
  const raw = props.rawPropertiesXml ?? '';
  switch (value.kind) {
    case 'dropdown': {
      if (props.sdtType !== 'dropDownList' && props.sdtType !== 'comboBox') {
        throw new ContentControlValueError(
          `Control is '${props.sdtType}', not a dropdown/combo box.`
        );
      }
      const item = props.listItems?.find(
        (it) => it.value === value.value || it.displayText === value.value
      );
      if (!item && props.sdtType === 'dropDownList') {
        throw new ContentControlValueError(
          `'${value.value}' is not one of the control's list items.`
        );
      }
      const storedValue = item?.value ?? value.value;
      const displayText = item?.displayText ?? value.value;
      const selectedIndex = item ? props.listItems?.indexOf(item) : undefined;
      // Always write `w:lastValue` (the stored selection): a dropdown/combo box
      // that has never been picked has no such attribute, and without it Word
      // loses the structured selection on reload even though the display text is
      // right. `setAttr` adds it when absent and is a no-op if the element is
      // missing entirely (e.g. an empty raw), so this is safe.
      const element = props.sdtType === 'comboBox' ? 'w:comboBox' : 'w:dropDownList';
      const nextRaw = setAttr(raw, element, 'w:lastValue', escapeXmlAttr(storedValue));
      return {
        properties: {
          ...withoutPlaceholder(props, nextRaw),
          listLastValue: storedValue,
          controlState: {
            ...props.controlState,
            value: displayText,
            selectedValue: storedValue,
            selectedIndex,
            placeholder: false,
          },
        },
        content: [paragraph(displayText)],
      };
    }
    case 'checkbox': {
      if (props.sdtType !== 'checkbox') {
        throw new ContentControlValueError(`Control is '${props.sdtType}', not a checkbox.`);
      }
      if (readAttr(raw, 'w14:checked', 'w14:val') == null) {
        throw new ContentControlValueError(
          'Checkbox control has no <w14:checked> state to update (not a Word checkbox).'
        );
      }
      const stateEl = value.checked ? 'w14:checkedState' : 'w14:uncheckedState';
      const char = codePointChar(readAttr(raw, stateEl, 'w14:val'), value.checked ? '☒' : '☐');
      // The glyph renders in the state's symbol font (e.g. MS Gothic).
      const font = readAttr(raw, stateEl, 'w14:font');
      const nextRaw = setAttr(raw, 'w14:checked', 'w14:val', value.checked ? '1' : '0');
      return {
        properties: {
          ...withoutPlaceholder(props, nextRaw),
          checked: value.checked,
          controlState: {
            ...props.controlState,
            checked: value.checked,
            value: char,
            placeholder: false,
          },
        },
        content: [paragraph(char, font)],
      };
    }
    case 'date': {
      if (props.sdtType !== 'date') {
        throw new ContentControlValueError(`Control is '${props.sdtType}', not a date control.`);
      }
      const iso = value.date.slice(0, 10);
      const [year, month, day] = /^\d{4}-\d{2}-\d{2}$/.test(iso) ? iso.split('-').map(Number) : [];
      if (!isValidIsoDate(year, month, day)) {
        throw new ContentControlValueError(`Date must be ISO yyyy-mm-dd, got '${value.date}'.`);
      }
      // Local-floating (no trailing Z): a UTC midnight would render as the
      // previous day in timezones behind UTC after Word's local conversion.
      const fullDate = `${iso}T00:00:00`;
      const nextRaw = setAttr(raw, 'w:date', 'w:fullDate', fullDate);
      const pattern =
        props.dateState?.format ?? props.dateFormat ?? readAttr(raw, 'w:dateFormat', 'w:val');
      const language = props.dateState?.language;
      const display = formatSdtDate(iso, pattern, language);
      return {
        properties: {
          ...withoutPlaceholder(props, nextRaw),
          dateFormat: pattern,
          dateState: { ...props.dateState, fullDate, format: pattern },
          controlState: {
            ...props.controlState,
            value: iso,
            placeholder: false,
          },
        },
        content: [paragraph(display)],
      };
    }
  }
}
