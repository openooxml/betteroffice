import { useEffect, useRef, useState } from 'react';
import { useTranslation } from '../../i18n';
import { MaterialSymbol } from '../ui/Icons';
import { useDisabledDescription } from '../ui/disabledDescription';
import { EDITING_MODES, type EditorMode } from './internals/editing-modes';

export function EditingModeDropdown({
  mode,
  onModeChange,
  disabled = false,
  description,
  optionState,
}: {
  mode: EditorMode;
  onModeChange: (mode: EditorMode) => void;
  disabled?: boolean;
  /** Why the dropdown is disabled. */
  description?: string;
  /** Availability of one mode; unavailable modes explain why. */
  optionState?: (mode: EditorMode) => { enabled: boolean; description?: string };
}) {
  const { t } = useTranslation();
  const reason = useDisabledDescription(disabled, description);
  const [isOpen, setIsOpen] = useState(false);
  const [compact, setCompact] = useState(false);
  const triggerRef = useRef<HTMLButtonElement>(null);
  const dropdownRef = useRef<HTMLDivElement>(null);
  const [pos, setPos] = useState({ top: 0, left: 0 });

  const current = EDITING_MODES.find((m) => m.value === mode)!;

  // Responsive: icon-only below 1400px
  useEffect(() => {
    const mql = window.matchMedia('(max-width: 1400px)');
    setCompact(mql.matches);
    const handler = (e: MediaQueryListEvent) => setCompact(e.matches);
    mql.addEventListener('change', handler);
    return () => mql.removeEventListener('change', handler);
  }, []);

  useEffect(() => {
    if (!isOpen || !triggerRef.current) return;
    const rect = triggerRef.current.getBoundingClientRect();
    // Align dropdown to right edge of trigger so it doesn't overflow the screen
    setPos({ top: rect.bottom + 2, left: rect.right - 220 });
  }, [isOpen]);

  useEffect(() => {
    if (!isOpen) return;
    const close = (e: MouseEvent) => {
      if (
        !triggerRef.current?.contains(e.target as Node) &&
        !dropdownRef.current?.contains(e.target as Node)
      ) {
        setIsOpen(false);
      }
    };
    const esc = (e: KeyboardEvent) => {
      if (e.key === 'Escape') setIsOpen(false);
    };
    document.addEventListener('mousedown', close);
    document.addEventListener('keydown', esc);
    return () => {
      document.removeEventListener('mousedown', close);
      document.removeEventListener('keydown', esc);
    };
  }, [isOpen]);

  return (
    <div style={{ position: 'relative' }}>
      <button
        ref={triggerRef}
        type="button"
        onMouseDown={(e) => e.preventDefault()}
        onClick={() => !disabled && setIsOpen(!isOpen)}
        {...reason.triggerProps}
        aria-label={t(current.labelKey)}
        aria-haspopup="menu"
        aria-expanded={isOpen}
        title={reason.title ? `${t(current.labelKey)}: ${reason.title}` : t(current.labelKey)}
        style={{
          display: 'flex',
          alignItems: 'center',
          gap: compact ? 0 : 4,
          padding: compact ? '2px 4px' : '2px 6px 2px 4px',
          border: 'none',
          background: isOpen ? 'var(--doc-bg-hover)' : 'transparent',
          borderRadius: 4,
          cursor: 'pointer',
          fontSize: 13,
          fontWeight: 400,
          color: 'var(--doc-text)',
          whiteSpace: 'nowrap',
          height: 28,
        }}
      >
        <MaterialSymbol name={current.icon} size={18} />
        {!compact && <span>{t(current.labelKey)}</span>}
        <MaterialSymbol name="arrow_drop_down" size={16} />
      </button>
      {reason.node}

      {isOpen && (
        <div
          ref={dropdownRef}
          data-docx-escape-layer="true"
          role="menu"
          aria-label={t('commands.editingMode')}
          onMouseDown={(e) => e.preventDefault()}
          style={{
            position: 'fixed',
            top: pos.top,
            left: pos.left,
            backgroundColor: 'var(--doc-surface)',
            border: '1px solid var(--doc-border)',
            borderRadius: 8,
            boxShadow: '0 4px 12px var(--doc-shadow)',
            padding: '4px 0',
            zIndex: 10000,
            minWidth: 220,
          }}
        >
          {EDITING_MODES.map((m) => {
            const state = optionState?.(m.value) ?? { enabled: true };
            return (
            <button
              key={m.value}
              type="button"
              role="menuitemradio"
              aria-checked={m.value === mode}
              aria-disabled={state.enabled ? undefined : true}
              title={state.enabled ? undefined : state.description}
              onMouseDown={(e) => e.preventDefault()}
              onClick={() => {
                if (!state.enabled) return;
                onModeChange(m.value);
                setIsOpen(false);
              }}
              onMouseOver={(e) => {
                (e.currentTarget as HTMLButtonElement).style.backgroundColor =
                  'var(--doc-bg-hover)';
              }}
              onMouseOut={(e) => {
                (e.currentTarget as HTMLButtonElement).style.backgroundColor = 'transparent';
              }}
              style={{
                display: 'flex',
                alignItems: 'center',
                gap: 10,
                padding: '8px 12px',
                border: 'none',
                background: 'transparent',
                cursor: state.enabled ? 'pointer' : 'default',
                opacity: state.enabled ? 1 : 0.5,
                fontSize: 13,
                color: 'var(--doc-text)',
                width: '100%',
                textAlign: 'left',
              }}
            >
              <MaterialSymbol name={m.icon} size={20} />
              <span style={{ display: 'flex', flexDirection: 'column', alignItems: 'flex-start' }}>
                <span style={{ fontWeight: 500 }}>{t(m.labelKey)}</span>
                <span style={{ fontSize: 11, color: 'var(--doc-text-muted)' }}>
                  {state.enabled ? t(m.descKey) : state.description}
                </span>
              </span>
              {m.value === mode && (
                <MaterialSymbol
                  name="check"
                  size={18}
                  style={{ marginLeft: 'auto', color: 'var(--doc-primary)' }}
                />
              )}
            </button>
            );
          })}
        </div>
      )}
    </div>
  );
}
