/**
 * Page Setup Dialog
 *
 * Modal for editing page layout properties:
 * - Page size (Letter, A4, Legal, etc.)
 * - Orientation (portrait/landscape)
 * - Margins (top, bottom, left, right) in inches or centimeters
 */

import React, { useState, useCallback, useEffect, useId } from 'react';
import type { CSSProperties } from 'react';
import type { SectionProperties } from '@betteroffice/docx/types/document';
import { TWIPS_PER_INCH } from '@betteroffice/docx/utils';
import { useTranslation } from '../../i18n';

/** Common page sizes in twips (width x height in portrait orientation) */
const PAGE_SIZES = [
  { labelKey: 'dialogs.pageSetup.pageSizes.letter' as const, width: 12240, height: 15840 },
  { labelKey: 'dialogs.pageSetup.pageSizes.a4' as const, width: 11906, height: 16838 },
  { labelKey: 'dialogs.pageSetup.pageSizes.legal' as const, width: 12240, height: 20160 },
  { labelKey: 'dialogs.pageSetup.pageSizes.a3' as const, width: 16838, height: 23811 },
  { labelKey: 'dialogs.pageSetup.pageSizes.a5' as const, width: 8391, height: 11906 },
  { labelKey: 'dialogs.pageSetup.pageSizes.b5' as const, width: 9979, height: 14175 },
  { labelKey: 'dialogs.pageSetup.pageSizes.executive' as const, width: 10440, height: 15120 },
] as const;

// ============================================================================
// TYPES
// ============================================================================

export interface PageSetupDialogProps {
  isOpen: boolean;
  onClose: () => void;
  onApply: (props: Partial<SectionProperties>) => void;
  currentProps?: SectionProperties;
  /** Unit margins are shown and typed in (default: 'inch'). */
  unit?: 'inch' | 'cm';
}

// ============================================================================
// HELPERS
// ============================================================================

const CM_PER_INCH = 2.54;

function twipsPerUnit(unit: 'inch' | 'cm'): number {
  return unit === 'cm' ? TWIPS_PER_INCH / CM_PER_INCH : TWIPS_PER_INCH;
}

function twipsToUnit(twips: number, unit: 'inch' | 'cm'): number {
  return Math.round((twips / twipsPerUnit(unit)) * 100) / 100;
}

function unitToTwips(value: number, unit: 'inch' | 'cm'): number {
  return Math.round(value * twipsPerUnit(unit));
}

/** Find matching page size preset, ignoring orientation */
function findPageSizeIndex(w: number, h: number): number {
  // Normalize to portrait (smaller dimension = width)
  const pw = Math.min(w, h);
  const ph = Math.max(w, h);
  return PAGE_SIZES.findIndex((s) => Math.abs(s.width - pw) < 20 && Math.abs(s.height - ph) < 20);
}

// ============================================================================
// STYLES
// ============================================================================

const overlayStyle: CSSProperties = {
  position: 'fixed',
  top: 0,
  left: 0,
  right: 0,
  bottom: 0,
  backgroundColor: 'var(--doc-overlay)',
  display: 'flex',
  alignItems: 'center',
  justifyContent: 'center',
  zIndex: 10000,
};

const dialogStyle: CSSProperties = {
  backgroundColor: 'var(--doc-surface)',
  borderRadius: 8,
  boxShadow: '0 4px 20px var(--doc-shadow)',
  minWidth: 400,
  maxWidth: 480,
  width: '100%',
  margin: 20,
};

const headerStyle: CSSProperties = {
  padding: '16px 20px 12px',
  borderBottom: '1px solid var(--doc-border)',
  fontSize: 16,
  fontWeight: 600,
};

const bodyStyle: CSSProperties = {
  padding: '16px 20px',
  display: 'flex',
  flexDirection: 'column',
  gap: 14,
};

const sectionLabelStyle: CSSProperties = {
  fontSize: 12,
  fontWeight: 600,
  color: 'var(--doc-text-muted)',
  textTransform: 'uppercase',
  letterSpacing: '0.5px',
};

const rowStyle: CSSProperties = {
  display: 'flex',
  alignItems: 'center',
  gap: 12,
};

const labelStyle: CSSProperties = {
  width: 80,
  fontSize: 13,
  color: 'var(--doc-text-muted)',
};

const inputStyle: CSSProperties = {
  flex: 1,
  padding: '6px 8px',
  border: '1px solid var(--doc-border)',
  borderRadius: 4,
  fontSize: 13,
};

const selectStyle: CSSProperties = {
  ...inputStyle,
};

const unitStyle: CSSProperties = {
  fontSize: 11,
  color: 'var(--doc-text-muted)',
  width: 16,
};

const footerStyle: CSSProperties = {
  padding: '12px 20px 16px',
  borderTop: '1px solid var(--doc-border)',
  display: 'flex',
  justifyContent: 'flex-end',
  gap: 8,
};

const btnStyle: CSSProperties = {
  padding: '6px 16px',
  fontSize: 13,
  border: '1px solid var(--doc-border)',
  borderRadius: 4,
  cursor: 'pointer',
};

// ============================================================================
// COMPONENT
// ============================================================================

// Default Word values (Letter, 1" margins)
const DEFAULT_WIDTH = 12240;
const DEFAULT_HEIGHT = 15840;
const DEFAULT_MARGIN = 1440;

export function PageSetupDialog({
  isOpen,
  onClose,
  onApply,
  currentProps,
  unit = 'inch',
}: PageSetupDialogProps): React.ReactElement | null {
  const { t } = useTranslation();
  const id = useId();
  const [pageWidth, setPageWidth] = useState(DEFAULT_WIDTH);
  const [pageHeight, setPageHeight] = useState(DEFAULT_HEIGHT);
  const [orientation, setOrientation] = useState<'portrait' | 'landscape'>('portrait');
  const [marginTop, setMarginTop] = useState(DEFAULT_MARGIN);
  const [marginBottom, setMarginBottom] = useState(DEFAULT_MARGIN);
  const [marginLeft, setMarginLeft] = useState(DEFAULT_MARGIN);
  const [marginRight, setMarginRight] = useState(DEFAULT_MARGIN);

  useEffect(() => {
    if (!isOpen) return;
    const w = currentProps?.pageWidth || DEFAULT_WIDTH;
    const h = currentProps?.pageHeight || DEFAULT_HEIGHT;
    const orient = currentProps?.orientation || (w > h ? 'landscape' : 'portrait');
    setPageWidth(w);
    setPageHeight(h);
    setOrientation(orient);
    setMarginTop(currentProps?.marginTop ?? DEFAULT_MARGIN);
    setMarginBottom(currentProps?.marginBottom ?? DEFAULT_MARGIN);
    setMarginLeft(currentProps?.marginLeft ?? DEFAULT_MARGIN);
    setMarginRight(currentProps?.marginRight ?? DEFAULT_MARGIN);
  }, [isOpen, currentProps]);

  const handlePageSizeChange = useCallback(
    (index: number) => {
      if (index < 0) return;
      const size = PAGE_SIZES[index];
      if (orientation === 'landscape') {
        setPageWidth(size.height);
        setPageHeight(size.width);
      } else {
        setPageWidth(size.width);
        setPageHeight(size.height);
      }
    },
    [orientation]
  );

  const handleOrientationChange = useCallback(
    (newOrientation: 'portrait' | 'landscape') => {
      if (newOrientation === orientation) return;
      setOrientation(newOrientation);
      // Swap width and height
      setPageWidth(pageHeight);
      setPageHeight(pageWidth);
    },
    [orientation, pageWidth, pageHeight]
  );

  const handleApply = useCallback(() => {
    onApply({
      pageWidth,
      pageHeight,
      orientation,
      marginTop,
      marginBottom,
      marginLeft,
      marginRight,
    });
    onClose();
  }, [
    pageWidth,
    pageHeight,
    orientation,
    marginTop,
    marginBottom,
    marginLeft,
    marginRight,
    onApply,
    onClose,
  ]);

  const handleKeyDown = useCallback(
    (e: React.KeyboardEvent) => {
      if (e.key === 'Escape') onClose();
      if (e.key === 'Enter') handleApply();
    },
    [onClose, handleApply]
  );

  if (!isOpen) return null;

  const sizeIndex = findPageSizeIndex(pageWidth, pageHeight);
  const maxMargin = unit === 'cm' ? 25.4 : 10;
  const unitLabel = t(unit === 'cm' ? 'common.units.cm' : 'common.units.inch');

  return (
    <div style={overlayStyle} onClick={onClose} onKeyDown={handleKeyDown}>
      <div
        style={dialogStyle}
        onClick={(e) => e.stopPropagation()}
        onMouseDown={(e) => e.stopPropagation()}
        role="dialog"
        aria-label={t('dialogs.pageSetup.title')}
      >
        <div style={headerStyle}>{t('dialogs.pageSetup.title')}</div>

        <div style={bodyStyle}>
          {/* Page size section */}
          <div style={sectionLabelStyle}>{t('dialogs.pageSetup.pageSize')}</div>

          <div style={rowStyle}>
            <label style={labelStyle} htmlFor={`${id}-size`}>{t('dialogs.pageSetup.sizeLabel')}</label>
            <select
              id={`${id}-size`}
              autoFocus
              style={selectStyle}
              value={sizeIndex}
              onChange={(e) => handlePageSizeChange(Number(e.target.value))}
            >
              {PAGE_SIZES.map((size, i) => (
                <option key={size.labelKey} value={i}>
                  {t(size.labelKey)}
                </option>
              ))}
              {sizeIndex < 0 && <option value={-1}>{t('dialogs.pageSetup.custom')}</option>}
            </select>
          </div>

          <div style={rowStyle}>
            <label style={labelStyle} htmlFor={`${id}-orientation`}>{t('dialogs.pageSetup.orientation')}</label>
            <select
              id={`${id}-orientation`}
              style={selectStyle}
              value={orientation}
              onChange={(e) => handleOrientationChange(e.target.value as 'portrait' | 'landscape')}
            >
              <option value="portrait">{t('dialogs.pageSetup.portrait')}</option>
              <option value="landscape">{t('dialogs.pageSetup.landscape')}</option>
            </select>
          </div>

          {/* Margins section */}
          <div style={{ ...sectionLabelStyle, marginTop: 4 }}>{t('dialogs.pageSetup.margins')}</div>

          <div style={rowStyle}>
            <label style={labelStyle} htmlFor={`${id}-top`}>{t('dialogs.pageSetup.top')}</label>
            <input
              id={`${id}-top`}
              type="number"
              style={inputStyle}
              min={0}
              max={maxMargin}
              step={0.1}
              value={twipsToUnit(marginTop, unit)}
              onChange={(e) => setMarginTop(unitToTwips(Number(e.target.value) || 0, unit))}
            />
            <span style={unitStyle}>{unitLabel}</span>
          </div>

          <div style={rowStyle}>
            <label style={labelStyle} htmlFor={`${id}-bottom`}>{t('dialogs.pageSetup.bottom')}</label>
            <input
              id={`${id}-bottom`}
              type="number"
              style={inputStyle}
              min={0}
              max={maxMargin}
              step={0.1}
              value={twipsToUnit(marginBottom, unit)}
              onChange={(e) => setMarginBottom(unitToTwips(Number(e.target.value) || 0, unit))}
            />
            <span style={unitStyle}>{unitLabel}</span>
          </div>

          <div style={rowStyle}>
            <label style={labelStyle} htmlFor={`${id}-left`}>{t('dialogs.pageSetup.left')}</label>
            <input
              id={`${id}-left`}
              type="number"
              style={inputStyle}
              min={0}
              max={maxMargin}
              step={0.1}
              value={twipsToUnit(marginLeft, unit)}
              onChange={(e) => setMarginLeft(unitToTwips(Number(e.target.value) || 0, unit))}
            />
            <span style={unitStyle}>{unitLabel}</span>
          </div>

          <div style={rowStyle}>
            <label style={labelStyle} htmlFor={`${id}-right`}>{t('dialogs.pageSetup.right')}</label>
            <input
              id={`${id}-right`}
              type="number"
              style={inputStyle}
              min={0}
              max={maxMargin}
              step={0.1}
              value={twipsToUnit(marginRight, unit)}
              onChange={(e) => setMarginRight(unitToTwips(Number(e.target.value) || 0, unit))}
            />
            <span style={unitStyle}>{unitLabel}</span>
          </div>
        </div>

        <div style={footerStyle}>
          <button type="button" style={btnStyle} onClick={onClose}>
            {t('common.cancel')}
          </button>
          <button
            type="button"
            style={{
              ...btnStyle,
              backgroundColor: 'var(--doc-primary)',
              color: 'var(--doc-on-primary)',
              borderColor: 'var(--doc-primary)',
            }}
            onClick={handleApply}
          >
            {t('common.apply')}
          </button>
        </div>
      </div>
    </div>
  );
}
