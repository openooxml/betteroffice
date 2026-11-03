import { useEffect, useLayoutEffect, useRef, useState } from 'react';
import type { CSSProperties, ReactElement } from 'react';
import type { PasteFormattingOption } from '@betteroffice/docx/yrs';
import { useTranslation } from '../../../i18n';
import { MaterialSymbol } from '../../ui/Icons';
import type { PasteOptionsAnchor } from './usePasteOptions';

const OPTIONS: Array<{ kind: PasteFormattingOption; labelKey: 'pasteOptions.keepSource' | 'pasteOptions.merge' | 'pasteOptions.textOnly' }> = [
  { kind: 'keep-source', labelKey: 'pasteOptions.keepSource' },
  { kind: 'merge', labelKey: 'pasteOptions.merge' },
  { kind: 'text-only', labelKey: 'pasteOptions.textOnly' },
];

/** @internal Where the open menu sits relative to the button so it stays inside the viewport. */
export function menuPlacement(
  button: { left: number; top: number; bottom: number },
  menu: { width: number; height: number },
  viewport: { width: number; height: number },
  margin = 8
): { left: number; top: number } {
  const below = button.bottom + 4;
  const top = below + menu.height <= viewport.height - margin || button.top - 4 - menu.height < margin
    ? below - button.top
    : -4 - menu.height;
  const right = button.left + menu.width;
  const left = right > viewport.width - margin ? Math.max(margin, viewport.width - margin - menu.width) - button.left : 0;
  return { left, top };
}

function clamped(anchor: PasteOptionsAnchor): CSSProperties {
  const width = 36;
  const height = 28;
  const left = Math.max(8, Math.min(anchor.left, Math.max(8, window.innerWidth - width - 8)));
  const top = Math.max(8, Math.min(anchor.top, Math.max(8, window.innerHeight - height - 8)));
  return { position: 'fixed', left, top, zIndex: 80 };
}

/** @internal Icon button and menu shown at the end of a formatted paste. */
export function PasteOptionsControl(props: {
  anchor: PasteOptionsAnchor;
  menuOpen: boolean;
  selected: PasteFormattingOption;
  onMenuOpenChange(open: boolean): void;
  onChoose(option: PasteFormattingOption): void;
  onCloseMenu(): void;
}): ReactElement {
  const { t } = useTranslation();
  const { anchor, menuOpen, selected, onMenuOpenChange, onChoose, onCloseMenu } = props;
  const rootRef = useRef<HTMLDivElement | null>(null);
  const itemRefs = useRef<Array<HTMLButtonElement | null>>([]);
  const menuRef = useRef<HTMLDivElement | null>(null);
  const [placement, setPlacement] = useState<{ left: number; top: number }>({ left: 0, top: 28 });
  const [style, setStyle] = useState<CSSProperties>(() => ({ position: 'fixed', left: anchor.left, top: anchor.top, zIndex: 80 }));
  const [focusIndex, setFocusIndex] = useState(0);

  useLayoutEffect(() => {
    setStyle(clamped(anchor));
  }, [anchor]);

  useLayoutEffect(() => {
    const root = rootRef.current;
    const menu = menuRef.current;
    if (!menuOpen || !root || !menu) return;
    const button = root.getBoundingClientRect();
    const size = menu.getBoundingClientRect();
    const next = menuPlacement(
      { left: button.left, top: button.top, bottom: button.top + 24 },
      { width: size.width, height: size.height },
      { width: window.innerWidth, height: window.innerHeight }
    );
    setPlacement((current) => (current.left === next.left && current.top === next.top ? current : next));
  }, [menuOpen, style]);

  useEffect(() => {
    if (!menuOpen) return;
    const index = Math.max(0, OPTIONS.findIndex(option => option.kind === selected));
    setFocusIndex(index);
    itemRefs.current[index]?.focus();
  }, [menuOpen, selected]);

  const move = (next: number) => {
    const index = (next + OPTIONS.length) % OPTIONS.length;
    setFocusIndex(index);
    itemRefs.current[index]?.focus();
  };

  return (
    <div ref={rootRef} data-paste-options style={style} onKeyDown={(event) => event.stopPropagation()}>
      <button
        type="button"
        aria-haspopup="menu"
        aria-expanded={menuOpen}
        aria-label={t('pasteOptions.button')}
        title={t('pasteOptions.button')}
        onClick={() => onMenuOpenChange(!menuOpen)}
        onKeyDown={(event) => {
          if (event.key === 'Escape' && menuOpen) {
            event.preventDefault();
            onCloseMenu();
          }
          if ((event.key === 'ArrowDown' || event.key === 'ArrowUp') && !menuOpen) {
            event.preventDefault();
            onMenuOpenChange(true);
          }
        }}
        style={{
          width: 28,
          height: 24,
          display: 'flex',
          alignItems: 'center',
          justifyContent: 'center',
          borderRadius: 4,
          border: '1px solid var(--doc-focus-ring, #888)',
          backgroundColor: 'var(--doc-surface, #fff)',
          color: 'var(--doc-primary, #333)',
          cursor: 'pointer',
          boxShadow: '0 1px 3px var(--doc-shadow, rgba(0,0,0,0.25))',
          padding: 0,
        }}
      >
        <MaterialSymbol name="format_paint" size={15} />
        <MaterialSymbol name="arrow_drop_down" size={13} />
      </button>
      {menuOpen && (
        <div
          ref={menuRef}
          role="menu"
          aria-label={t('pasteOptions.button')}
          style={{
            position: 'absolute',
            left: placement.left,
            top: placement.top,
            minWidth: 220,
            padding: 4,
            borderRadius: 6,
            border: '1px solid var(--doc-focus-ring, #888)',
            backgroundColor: 'var(--doc-surface, #fff)',
            boxShadow: '0 4px 16px var(--doc-shadow, rgba(0,0,0,0.3))',
          }}
          onKeyDown={(event) => {
            if (event.key === 'Escape') {
              event.preventDefault();
              onCloseMenu();
            } else if (event.key === 'ArrowDown') {
              event.preventDefault();
              move(focusIndex + 1);
            } else if (event.key === 'ArrowUp') {
              event.preventDefault();
              move(focusIndex - 1);
            } else if (event.key === 'Home') {
              event.preventDefault();
              move(0);
            } else if (event.key === 'End') {
              event.preventDefault();
              move(OPTIONS.length - 1);
            }
          }}
        >
          {OPTIONS.map((option, index) => (
            <button
              key={option.kind}
              ref={(element) => { itemRefs.current[index] = element; }}
              type="button"
              role="menuitemradio"
              aria-checked={selected === option.kind}
              tabIndex={focusIndex === index ? 0 : -1}
              onClick={() => onChoose(option.kind)}
              onKeyDown={(event) => {
                if (event.key === 'Enter' || event.key === ' ') {
                  event.preventDefault();
                  onChoose(option.kind);
                }
              }}
              style={{
                display: 'flex',
                width: '100%',
                alignItems: 'center',
                gap: 8,
                padding: '6px 8px',
                border: 0,
                borderRadius: 4,
                backgroundColor: 'transparent',
                color: 'inherit',
                cursor: 'pointer',
                textAlign: 'left',
                font: 'inherit',
              }}
            >
              <span aria-hidden="true" style={{ width: 14, flex: '0 0 14px', display: 'flex', justifyContent: 'center' }}>
                {selected === option.kind ? <MaterialSymbol name="check" size={14} /> : null}
              </span>
              <span>{t(option.labelKey)}</span>
            </button>
          ))}
        </div>
      )}
    </div>
  );
}

export default PasteOptionsControl;
