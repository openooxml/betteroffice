import { useEffect, useRef, useState } from 'react';
import type { CSSProperties } from 'react';

import { useTranslation } from '../i18n';
import { Z_INDEX } from '../styles/zIndex';

export interface InlineHeaderFooterEditorProps {
  position: 'header' | 'footer';
  targetRect: { top: number; left: number; width: number; height: number };
  onClose: () => void;
  onRemove?: () => void;
  onInsertField?: (fieldType: 'PAGE' | 'NUMPAGES') => void;
}

const separatorBarStyle: CSSProperties = {
  position: 'absolute',
  left: 0,
  right: 0,
  bottom: '100%',
  display: 'flex',
  alignItems: 'center',
  justifyContent: 'space-between',
  padding: '2px 0',
  fontSize: 11,
  color: 'var(--doc-primary)',
  userSelect: 'none',
  pointerEvents: 'auto',
};

const optionsButtonStyle: CSSProperties = {
  background: 'none',
  border: 'none',
  color: 'var(--doc-primary)',
  cursor: 'pointer',
  fontSize: 11,
  padding: '2px 6px',
  borderRadius: 3,
};

const dropdownStyle: CSSProperties = {
  position: 'absolute',
  right: 0,
  top: '100%',
  background: 'var(--doc-surface)',
  border: '1px solid var(--doc-border-light)',
  borderRadius: 4,
  boxShadow: '0 2px 6px var(--doc-shadow)',
  zIndex: Z_INDEX.dropdown,
  minWidth: 160,
  padding: '4px 0',
};

const dropdownItemStyle: CSSProperties = {
  display: 'block',
  width: '100%',
  padding: '6px 12px',
  border: 'none',
  background: 'none',
  textAlign: 'left',
  cursor: 'pointer',
  fontSize: 12,
  color: 'var(--doc-text)',
};

export function InlineHeaderFooterEditor({
  position,
  targetRect,
  onClose,
  onRemove,
  onInsertField,
}: InlineHeaderFooterEditorProps) {
  const { t } = useTranslation();
  const [showOptions, setShowOptions] = useState(false);
  const optionsRef = useRef<HTMLDivElement>(null);
  const label = position === 'header' ? t('headerFooter.header') : t('headerFooter.footer');

  useEffect(() => {
    const handleKeyDown = (event: KeyboardEvent) => {
      if (event.key !== 'Escape') return;
      event.preventDefault();
      event.stopPropagation();
      onClose();
    };
    document.addEventListener('keydown', handleKeyDown, true);
    return () => document.removeEventListener('keydown', handleKeyDown, true);
  }, [onClose]);

  useEffect(() => {
    if (!showOptions) return;
    const handleClick = (event: MouseEvent) => {
      if (optionsRef.current && !optionsRef.current.contains(event.target as Node)) {
        setShowOptions(false);
      }
    };
    document.addEventListener('mousedown', handleClick);
    return () => document.removeEventListener('mousedown', handleClick);
  }, [showOptions]);

  const containerStyle: CSSProperties = {
    position: 'absolute',
    top: targetRect.top,
    left: targetRect.left,
    width: targetRect.width,
    height: targetRect.height,
    zIndex: Z_INDEX.hfInlineEditor,
    pointerEvents: 'none',
  };
  const barStyle =
    position === 'footer'
      ? ({ ...separatorBarStyle, bottom: 'auto', top: '100%' } satisfies CSSProperties)
      : separatorBarStyle;

  return (
    <div className="hf-inline-editor" style={containerStyle}>
      <div className="hf-separator-bar" style={barStyle}>
        <span style={{ fontWeight: 500, letterSpacing: 0.3 }}>{label}</span>
        <div style={{ position: 'relative' }} ref={optionsRef}>
          <button
            type="button"
            style={optionsButtonStyle}
            onClick={(event) => {
              event.stopPropagation();
              setShowOptions((visible) => !visible);
            }}
            onMouseDown={(event) => event.stopPropagation()}
          >
            {t('headerFooter.options')} ▾
          </button>
          {showOptions && (
            <div style={dropdownStyle} onMouseDown={(event) => event.stopPropagation()}>
              {onInsertField && (
                <>
                  <MenuButton
                    onClick={() => onInsertField('PAGE')}
                    label={t('headerFooter.insertPageNumber')}
                  />
                  <MenuButton
                    onClick={() => onInsertField('NUMPAGES')}
                    label={t('headerFooter.insertTotalPages')}
                  />
                  <div
                    style={{ borderTop: '1px solid var(--doc-border-light)', margin: '4px 0' }}
                  />
                </>
              )}
              {onRemove && (
                <MenuButton
                  onClick={onRemove}
                  label={t('headerFooter.remove', { label: label.toLowerCase() })}
                />
              )}
              <MenuButton
                onClick={onClose}
                label={t('headerFooter.closeEditing', { label: label.toLowerCase() })}
              />
            </div>
          )}
        </div>
      </div>
    </div>
  );
}

function MenuButton({ onClick, label }: { onClick: () => void; label: string }) {
  return (
    <button
      type="button"
      style={dropdownItemStyle}
      onClick={() => {
        onClick();
      }}
      onMouseOver={(event) => {
        event.currentTarget.style.backgroundColor = 'var(--doc-bg-hover)';
      }}
      onMouseOut={(event) => {
        event.currentTarget.style.backgroundColor = 'transparent';
      }}
    >
      {label}
    </button>
  );
}
