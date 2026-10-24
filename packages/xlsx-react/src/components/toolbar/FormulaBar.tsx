import { parseHyperlinkLocation } from '@betteroffice/xlsx';
import type { CellAddr } from '@betteroffice/xlsx';
import { createContext, useContext, useLayoutEffect, useRef, useState } from 'react';
import type { CSSProperties, KeyboardEvent, Ref } from 'react';
import {
  TOOLBAR_FONT_FAMILY,
  toolbarFieldStyle,
  toolbarTheme,
  useFocusRing,
} from '../../../../../shared/react-toolbar/theme';
import { useTranslation } from '../../i18n';

/** The editor's formula input, shared with whichever chrome renders the formula bar. */
export interface FormulaBarBinding {
  /** Address of the focused cell. */
  a1: string;
  onNavigate?(cell: CellAddr): void;
  value: string;
  disabled: boolean;
  /** No writable cell selection: the text shows but cannot be edited. */
  readOnly: boolean;
  inputRef: Ref<HTMLInputElement>;
  onChange(value: string): void;
  onCommit(move: 'up' | 'down'): void;
  onCancel(): void;
  onBlur(): void;
  onCompositionStart(): void;
  onCompositionEnd(): void;
  native?: boolean;
  isComposing?(): boolean;
  onKeyDown?(event: KeyboardEvent<HTMLInputElement>): boolean;
  onKeyUp?(): void;
}

export const FormulaBarContext = createContext<FormulaBarBinding | null>(null);

/** Replaces the text of an input the user has not typed into, keeping a whole-text selection whole. */
export function presentInputValue(input: HTMLInputElement, value: string): void {
  const whole = input.ownerDocument.activeElement === input && input.selectionStart === 0 &&
    input.selectionEnd === input.value.length;
  input.value = value;
  if (whole) input.setSelectionRange(0, value.length);
}

const styles: Record<string, CSSProperties> = {
  group: {
    display: 'flex',
    alignItems: 'center',
    gap: 4,
    flex: '1 1 320px',
    minWidth: 240,
  },
  nameBox: {
    ...toolbarFieldStyle({}),
    width: 72,
    flex: '0 0 auto',
    font: `500 13px ${TOOLBAR_FONT_FAMILY}`,
    textAlign: 'center',
    outlineColor: toolbarTheme.accent,
  },
  mark: {
    display: 'grid',
    placeItems: 'center',
    width: 20,
    height: 28,
    flex: '0 0 auto',
    color: toolbarTheme.inkMuted,
    font: `italic 700 13px ${TOOLBAR_FONT_FAMILY}`,
    userSelect: 'none',
  },
};

/** @experimental */
export interface FormulaBarProps {
  className?: string;
  style?: CSSProperties;
}

/**
 * The name box and formula input of the surrounding `XlsxEditor`. It edits
 * the focused cell through the editor's input ordering, so commands issued
 * while a formula is typed land after it. Renders nothing outside an editor.
 * @experimental
 */
export function FormulaBar({ className, style }: FormulaBarProps) {
  const binding = useContext(FormulaBarContext);
  const [nameDraft, setNameDraft] = useState<string | null>(null);
  const input = useRef<HTMLInputElement | null>(null);
  const { t } = useTranslation();
  useLayoutEffect(() => {
    if (binding?.native && input.current && !binding.isComposing?.() && input.current.value !== binding.value) {
      presentInputValue(input.current, binding.value);
    }
  }, [binding]);
  const formulaFocus = useFocusRing<HTMLInputElement>({
    onBlur: () => binding?.onBlur(),
  });
  if (!binding) return null;
  return (
    <div
      className={className}
      style={{ ...styles.group, ...style }}
      role="group"
      aria-label={t('toolbar.formulaBarLabel')}
    >
      <input
        data-testid="xlsx-name-box"
        readOnly={!binding.onNavigate}
        value={nameDraft ?? binding.a1}
        onChange={(event) => setNameDraft(event.target.value)}
        onBlur={() => setNameDraft(null)}
        onKeyDown={(event) => {
          if (event.key !== 'Enter' && event.key !== 'Escape') return;
          event.preventDefault(); event.stopPropagation();
          const value = event.currentTarget.value.trim();
          const cell = /^\$?[A-Za-z]{1,3}\$?[1-9]\d*$/.test(value) ? parseHyperlinkLocation(value, 'Sheet') : null;
          setNameDraft(null);
          if (event.key === 'Enter' && cell) binding.onNavigate?.(cell);
        }}
        placeholder={t('toolbar.nameBoxPlaceholder')}
        aria-label={t('toolbar.nameBoxPlaceholder')}
        style={styles.nameBox}
      />
      <span style={styles.mark} aria-hidden="true">
        fx
      </span>
      <input
        ref={binding.native ? (value) => {
          input.current = value;
          if (typeof binding.inputRef === 'function') binding.inputRef(value);
          else if (binding.inputRef) binding.inputRef.current = value;
        } : binding.inputRef}
        data-testid="xlsx-formula-input"
        value={binding.native ? undefined : binding.value}
        defaultValue={binding.native ? '' : undefined}
        placeholder={t('toolbar.formulaPlaceholder')}
        aria-label={t('toolbar.formulaPlaceholder')}
        disabled={binding.disabled}
        readOnly={binding.readOnly}
        onChange={(event) => binding.onChange(event.target.value)}
        onCompositionStart={binding.onCompositionStart}
        onCompositionEnd={binding.onCompositionEnd}
        onKeyDown={(event) => {
          if (binding.onKeyDown?.(event)) return;
          if (binding.readOnly || event.nativeEvent.isComposing || event.keyCode === 229) return;
          if (event.key === 'Enter') {
            binding.onCommit(event.shiftKey ? 'up' : 'down');
            event.preventDefault();
          } else if (event.key === 'Escape') {
            binding.onCancel();
            event.preventDefault();
          }
        }}
        onKeyUp={binding.onKeyUp}
        onFocus={formulaFocus.focusProps.onFocus}
        onBlur={formulaFocus.focusProps.onBlur}
        style={{
          ...toolbarFieldStyle({ focusRing: formulaFocus.focusRing }),
          flex: '1 1 260px',
          minWidth: 140,
          padding: '0 8px',
        }}
      />
    </div>
  );
}
