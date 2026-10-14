import { useEffect, useRef, useState } from 'react';
import type { CSSProperties, DragEvent, KeyboardEvent, SVGProps } from 'react';
import type { TFunction } from '@betteroffice/vsdx-i18n';

export interface StatusBarPage {
  id: string;
  name: string | null;
}

export interface StatusBarProps {
  pages: ReadonlyArray<StatusBarPage>;
  activeIndex: number;
  onSelectPage: (index: number) => void;
  onReorderPage?: (pageId: string, toIndex: number) => void;
  onAddPage?: () => void;
  zoom: number;
  onZoomChange: (zoom: number) => void;
  onFitToWindow: () => void;
  t: TFunction;
  className?: string;
}

export const ZOOM_STOPS = [0.1, 0.25, 0.5, 0.75, 1, 1.5, 2, 4] as const;
export const MIN_ZOOM = ZOOM_STOPS[0];
export const MAX_ZOOM = ZOOM_STOPS[ZOOM_STOPS.length - 1];

export function sliderPositionForZoom(zoom: number): number {
  const value = clampZoom(zoom);
  return value <= 1
    ? Math.log10(value / MIN_ZOOM) / 2
    : 0.5 + Math.log(value) / Math.log(MAX_ZOOM) / 2;
}

export function zoomForSliderPosition(position: number): number {
  const value = Math.max(0, Math.min(1, Number.isFinite(position) ? position : 0.5));
  return value <= 0.5
    ? MIN_ZOOM * 10 ** (value * 2)
    : MAX_ZOOM ** ((value - 0.5) * 2);
}

export function clampZoom(zoom: number): number {
  if (!Number.isFinite(zoom) || zoom <= 0) return 1;
  return Math.max(MIN_ZOOM, Math.min(MAX_ZOOM, zoom));
}

export function StatusBar({
  pages,
  activeIndex,
  onSelectPage,
  onReorderPage,
  onAddPage,
  zoom,
  onZoomChange,
  onFitToWindow,
  t,
  className,
}: StatusBarProps) {
  const [menuOpen, setMenuOpen] = useState(false);
  const menuTriggerRef = useRef<HTMLButtonElement>(null);
  const tabRefs = useRef(new Map<number, HTMLButtonElement>());
  const draggedPageIdRef = useRef<string | null>(null);
  const currentIndex = pages.length ? Math.max(0, Math.min(activeIndex, pages.length - 1)) : 0;
  const currentZoom = clampZoom(zoom);
  const zoomPercent = Math.round(currentZoom * 100);

  useEffect(() => {
    if (!menuOpen) return;
    const onKeyDown = (event: globalThis.KeyboardEvent) => {
      if (event.key !== 'Escape') return;
      event.preventDefault();
      setMenuOpen(false);
      menuTriggerRef.current?.focus();
    };
    document.addEventListener('keydown', onKeyDown);
    return () => document.removeEventListener('keydown', onKeyDown);
  }, [menuOpen]);

  const pageLabel = (page: StatusBarPage, index: number) =>
    page.name ?? t('pages.fallbackTitle', { number: index + 1 });
  const selectPage = (index: number) => {
    onSelectPage(index);
    tabRefs.current.get(index)?.focus();
  };
  const moveTabFocus = (event: KeyboardEvent<HTMLButtonElement>, index: number) => {
    let nextIndex: number | null = null;
    if (event.key === 'ArrowLeft') nextIndex = Math.max(0, index - 1);
    if (event.key === 'ArrowRight') nextIndex = Math.min(pages.length - 1, index + 1);
    if (event.key === 'Home') nextIndex = 0;
    if (event.key === 'End') nextIndex = pages.length - 1;
    if (nextIndex === null) return;
    event.preventDefault();
    selectPage(nextIndex);
  };
  const changeZoomByStop = (direction: -1 | 1) => {
    const stops = direction > 0 ? ZOOM_STOPS : [...ZOOM_STOPS].reverse();
    const next = direction > 0
      ? stops.find((value) => value > currentZoom + Number.EPSILON) ?? MAX_ZOOM
      : stops.find((value) => value < currentZoom - Number.EPSILON) ?? MIN_ZOOM;
    onZoomChange(next);
  };
  const closeMenu = () => {
    setMenuOpen(false);
    menuTriggerRef.current?.focus();
  };
  const onDropPage = (event: DragEvent<HTMLButtonElement>, index: number) => {
    event.preventDefault();
    const sourcePageId = draggedPageIdRef.current;
    draggedPageIdRef.current = null;
    if (sourcePageId === null) return;
    const sourceIndex = pages.findIndex((page) => page.id === sourcePageId);
    if (sourceIndex < 0 || sourceIndex === index || index >= pages.length) return;
    onReorderPage?.(sourcePageId, index);
  };

  return (
    <footer className={className} style={styles.root}>
      <div style={styles.pagesSection}>
        <div style={styles.pageControls}>
          <button
            ref={menuTriggerRef}
            type="button"
            aria-label={t('statusBar.pageList')}
            aria-haspopup="menu"
            aria-expanded={menuOpen}
            title={t('statusBar.pageList')}
            onClick={() => setMenuOpen((open) => !open)}
            style={iconButtonStyle(false, menuOpen)}
          >
            <Icon name="list" />
          </button>
          {menuOpen && (
            <div role="menu" aria-label={t('statusBar.pageList')} style={styles.menu}>
              {pages.map((page, index) => (
                <button
                  key={page.id}
                  type="button"
                  role="menuitemradio"
                  aria-checked={index === currentIndex}
                  onClick={() => {
                    onSelectPage(index);
                    closeMenu();
                  }}
                  style={menuItemStyle(index === currentIndex)}
                >
                  {pageLabel(page, index)}
                </button>
              ))}
            </div>
          )}
          <button
            type="button"
            disabled={currentIndex === 0 || pages.length === 0}
            aria-label={t('statusBar.previousPage')}
            title={t('statusBar.previousPage')}
            onClick={() => selectPage(currentIndex - 1)}
            style={iconButtonStyle(currentIndex === 0 || pages.length === 0)}
          >
            <Icon name="previous" />
          </button>
          <button
            type="button"
            disabled={pages.length === 0 || currentIndex === pages.length - 1}
            aria-label={t('statusBar.nextPage')}
            title={t('statusBar.nextPage')}
            onClick={() => selectPage(currentIndex + 1)}
            style={iconButtonStyle(pages.length === 0 || currentIndex === pages.length - 1)}
          >
            <Icon name="next" />
          </button>
        </div>
        <div role="tablist" aria-label={t('statusBar.pageTabs')} style={styles.tabs}>
          {pages.map((page, index) => (
            <button
              key={page.id}
              ref={(element) => {
                if (element) tabRefs.current.set(index, element);
                else tabRefs.current.delete(index);
              }}
              type="button"
              role="tab"
              aria-selected={index === currentIndex}
              tabIndex={index === currentIndex ? 0 : -1}
              draggable={Boolean(onReorderPage)}
              onClick={() => selectPage(index)}
              onKeyDown={(event) => moveTabFocus(event, index)}
              onDragStart={() => { draggedPageIdRef.current = page.id; }}
              onDragOver={(event) => { if (onReorderPage) event.preventDefault(); }}
              onDrop={(event) => onDropPage(event, index)}
              style={tabStyle(index === currentIndex)}
            >
              {pageLabel(page, index)}
            </button>
          ))}
        </div>
        {onAddPage && (
          <button
            type="button"
            aria-label={t('statusBar.addPage')}
            title={t('statusBar.addPage')}
            onClick={onAddPage}
            style={iconButtonStyle(false)}
          >
            <Icon name="add" />
          </button>
        )}
      </div>
      <div role="group" aria-label={t('statusBar.zoomControls')} style={styles.zoomSection}>
        <button type="button" aria-label={t('statusBar.zoomOut')} title={t('statusBar.zoomOut')} onClick={() => changeZoomByStop(-1)} style={iconButtonStyle(false)}><Icon name="minus" /></button>
        <input
          type="range"
          min="0"
          max="1000"
          step="1"
          value={Math.round(sliderPositionForZoom(currentZoom) * 1000)}
          aria-label={t('statusBar.zoomSlider')}
          aria-valuemin={MIN_ZOOM * 100}
          aria-valuemax={MAX_ZOOM * 100}
          aria-valuenow={zoomPercent}
          aria-valuetext={t('statusBar.zoomPercentage', { zoom: zoomPercent })}
          onChange={(event) => onZoomChange(clampZoom(zoomForSliderPosition(Number(event.currentTarget.value) / 1000)))}
          style={styles.slider}
        />
        <button type="button" aria-label={t('statusBar.zoomIn')} title={t('statusBar.zoomIn')} onClick={() => changeZoomByStop(1)} style={iconButtonStyle(false)}><Icon name="add" /></button>
        <button type="button" aria-label={t('statusBar.resetZoom')} title={t('statusBar.resetZoom')} onClick={() => onZoomChange(1)} style={styles.zoomReadout}>
          {t('statusBar.zoomPercentage', { zoom: zoomPercent })}
        </button>
        <button type="button" aria-label={t('statusBar.fitToWindow')} title={t('statusBar.fitToWindow')} onClick={onFitToWindow} style={iconButtonStyle(false)}><Icon name="fit" /></button>
      </div>
    </footer>
  );
}

function Icon({ name }: { name: 'add' | 'fit' | 'list' | 'minus' | 'next' | 'previous' }) {
  const paths: Record<typeof name, SVGProps<SVGPathElement>['d']> = {
    add: 'M5 12h14M12 5v14',
    fit: 'M5 9V5h4M15 5h4v4M19 15v4h-4M9 19H5v-4',
    list: 'M5 7h14M5 12h14M5 17h14',
    minus: 'M5 12h14',
    next: 'm9 5 7 7-7 7',
    previous: 'm15 5-7 7 7 7',
  };
  return <svg aria-hidden="true" viewBox="0 0 24 24" width="16" height="16" fill="none" stroke="currentColor" strokeWidth="1.8" strokeLinecap="round" strokeLinejoin="round"><path d={paths[name]} /></svg>;
}

function iconButtonStyle(disabled: boolean, active = false): CSSProperties {
  return {
    ...styles.iconButton,
    background: active ? '#d3e3fd' : 'transparent',
    color: disabled ? '#9aa0a6' : '#3c4043',
    cursor: disabled ? 'default' : 'pointer',
    opacity: disabled ? 0.48 : 1,
  };
}

function tabStyle(active: boolean): CSSProperties {
  return {
    ...styles.tab,
    borderBottomColor: active ? '#1a73e8' : 'transparent',
    color: active ? '#174ea6' : '#3c4043',
    fontWeight: active ? 600 : 400,
  };
}

function menuItemStyle(active: boolean): CSSProperties {
  return {
    ...styles.menuItem,
    background: active ? '#e8f0fe' : 'transparent',
    color: active ? '#174ea6' : '#3c4043',
    fontWeight: active ? 600 : 400,
  };
}

const styles: Record<string, CSSProperties> = {
  root: { display: 'flex', minWidth: 0, height: 38, alignItems: 'stretch', justifyContent: 'space-between', gap: 12, padding: '0 8px', color: '#3c4043', background: '#ffffff', borderTop: '1px solid #d9dde3', font: '400 12px ui-sans-serif, system-ui, sans-serif', boxSizing: 'border-box' },
  pagesSection: { position: 'relative', display: 'flex', minWidth: 0, flex: '1 1 auto', alignItems: 'center', gap: 2 },
  pageControls: { display: 'inline-flex', flex: '0 0 auto', alignItems: 'center', gap: 1 },
  iconButton: { appearance: 'none', display: 'inline-grid', placeItems: 'center', flex: '0 0 auto', width: 28, height: 28, padding: 0, border: 0, borderRadius: 4, lineHeight: 1 },
  tabs: { display: 'flex', minWidth: 0, flex: '1 1 auto', alignSelf: 'stretch', overflowX: 'auto', overflowY: 'hidden', scrollbarWidth: 'thin' },
  tab: { appearance: 'none', flex: '0 0 auto', maxWidth: 180, height: '100%', padding: '0 13px', overflow: 'hidden', border: 0, borderBottom: '2px solid', background: 'transparent', cursor: 'pointer', font: 'inherit', lineHeight: 1, overflowWrap: 'normal', textOverflow: 'ellipsis', whiteSpace: 'nowrap' },
  menu: { position: 'absolute', zIndex: 10, bottom: 34, left: 0, width: 220, maxHeight: 300, overflowY: 'auto', padding: 4, border: '1px solid #c7cacf', borderRadius: 6, background: '#ffffff', boxShadow: '0 4px 16px rgba(60, 64, 67, 0.24)', boxSizing: 'border-box' },
  menuItem: { appearance: 'none', display: 'block', width: '100%', minHeight: 30, padding: '5px 8px', overflow: 'hidden', border: 0, borderRadius: 4, cursor: 'pointer', font: 'inherit', textAlign: 'left', textOverflow: 'ellipsis', whiteSpace: 'nowrap' },
  zoomSection: { display: 'inline-flex', flex: '0 0 auto', alignItems: 'center', gap: 2 },
  slider: { width: 116, accentColor: '#1a73e8' },
  zoomReadout: { appearance: 'none', minWidth: 48, height: 28, padding: '0 6px', border: 0, borderRadius: 4, color: '#3c4043', background: 'transparent', cursor: 'pointer', font: '500 12px ui-sans-serif, system-ui, sans-serif' },
};
