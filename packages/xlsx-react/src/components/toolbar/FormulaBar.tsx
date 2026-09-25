import { createContext, useContext } from 'react';
import type { CSSProperties, Ref } from 'react';
import { useTranslation } from '../../i18n';

/** The editor's formula input, shared with whichever chrome renders the formula bar. */
export interface FormulaBarBinding {
  /** Address of the focused cell. */
  a1: string;
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
}

export const FormulaBarContext = createContext<FormulaBarBinding | null>(null);

const styles: Record<string, CSSProperties> = {
  group: {
    display: 'flex',
    alignItems: 'center',
    gap: 4,
    flex: '1 1 320px',
    minWidth: 240,
    padding: '0 6px',
    borderRight: '1px solid rgba(226, 232, 240, 0.9)',
  },
  nameBox: {
    appearance: 'none',
    width: 64,
    height: 28,
    flex: '0 0 auto',
    boxSizing: 'border-box',
    border: '1px solid #e2e8f0',
    borderRadius: 6,
    background: '#f8fafc',
    color: '#0f172a',
    font: '600 12px ui-monospace, SFMono-Regular, Menlo, monospace',
    textAlign: 'center',
    outlineColor: '#2563eb',
  },
  mark: {
    display: 'grid',
    placeItems: 'center',
    width: 20,
    height: 28,
    flex: '0 0 auto',
    color: '#64748b',
    font: 'italic 700 12px Georgia, serif',
    userSelect: 'none',
  },
  input: {
    appearance: 'none',
    flex: '1 1 260px',
    minWidth: 140,
    height: 28,
    boxSizing: 'border-box',
    border: '1px solid #e2e8f0',
    borderRadius: 6,
    padding: '0 8px',
    background: '#ffffff',
    color: '#0f172a',
    font: '13px ui-sans-serif, system-ui, sans-serif',
    outlineColor: '#2563eb',
  },
};

export interface FormulaBarProps {
  className?: string;
  style?: CSSProperties;
}

/**
 * The name box and formula input of the surrounding `XlsxEditor`. It edits
 * the focused cell through the editor's input ordering, so commands issued
 * while a formula is typed land after it. Renders nothing outside an editor.
 */
export function FormulaBar({ className, style }: FormulaBarProps) {
  const binding = useContext(FormulaBarContext);
  const { t } = useTranslation();
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
        readOnly
        value={binding.a1}
        placeholder={t('toolbar.nameBoxPlaceholder')}
        aria-label={t('toolbar.nameBoxPlaceholder')}
        style={styles.nameBox}
      />
      <span style={styles.mark} aria-hidden="true">
        fx
      </span>
      <input
        ref={binding.inputRef}
        data-testid="xlsx-formula-input"
        value={binding.value}
        placeholder={t('toolbar.formulaPlaceholder')}
        aria-label={t('toolbar.formulaPlaceholder')}
        disabled={binding.disabled}
        readOnly={binding.readOnly}
        onChange={(event) => binding.onChange(event.target.value)}
        onCompositionStart={binding.onCompositionStart}
        onCompositionEnd={binding.onCompositionEnd}
        onKeyDown={(event) => {
          if (binding.readOnly || event.nativeEvent.isComposing || event.keyCode === 229) return;
          if (event.key === 'Enter') {
            binding.onCommit(event.shiftKey ? 'up' : 'down');
            event.preventDefault();
          } else if (event.key === 'Escape') {
            binding.onCancel();
            event.preventDefault();
          }
        }}
        onBlur={binding.onBlur}
        style={styles.input}
      />
    </div>
  );
}
