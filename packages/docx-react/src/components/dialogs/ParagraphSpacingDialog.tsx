import { useEffect, useId, useState, type CSSProperties } from 'react';
import { useTranslation } from '../../i18n';
import { getDefaultLineSpacingOptions } from '../ui/LineSpacingPicker';

export interface ParagraphSpacingData {
  before: number;
  after: number;
  lineSpacing?: number;
}

export interface ParagraphSpacingDialogProps {
  isOpen: boolean;
  onClose(): void;
  onApply(spacing: ParagraphSpacingData): void;
  currentProps?: Partial<ParagraphSpacingData>;
}

const overlayStyle: CSSProperties = {
  position: 'fixed', inset: 0, backgroundColor: 'var(--doc-overlay)',
  display: 'flex', alignItems: 'center', justifyContent: 'center', zIndex: 10000,
};
const dialogStyle: CSSProperties = {
  backgroundColor: 'var(--doc-surface)', color: 'var(--doc-text)', borderRadius: 8,
  boxShadow: '0 4px 20px var(--doc-shadow)', width: 400, maxWidth: 'calc(100% - 40px)',
};
const rowStyle: CSSProperties = { display: 'flex', alignItems: 'center', gap: 12 };
const labelStyle: CSSProperties = { width: 140, fontSize: 13, color: 'var(--doc-text-muted)' };
const inputStyle: CSSProperties = {
  flex: 1, minWidth: 0, padding: '6px 8px', border: '1px solid var(--doc-border)', borderRadius: 4, fontSize: 13,
};
const buttonStyle: CSSProperties = {
  padding: '6px 16px', border: '1px solid var(--doc-border)', borderRadius: 4,
  backgroundColor: 'var(--doc-surface)', color: 'var(--doc-text)', cursor: 'pointer', fontSize: 13,
};

export function ParagraphSpacingDialog({ isOpen, onClose, onApply, currentProps }: ParagraphSpacingDialogProps) {
  const { t, lang } = useTranslation();
  const id = useId();
  const initialBefore = currentProps?.before ?? 0;
  const initialAfter = currentProps?.after ?? 0;
  const initialLine = currentProps?.lineSpacing ?? 240;
  const [before, setBefore] = useState(String(initialBefore / 20));
  const [after, setAfter] = useState(String(initialAfter / 20));
  const [line, setLine] = useState(String(initialLine / 240));
  const [lineChanged, setLineChanged] = useState(false);
  const options = getDefaultLineSpacingOptions();
  const [multiple, setMultiple] = useState(!options.some((option) => option.twipsValue === initialLine));
  useEffect(() => {
    if (!isOpen) return;
    setBefore(String(initialBefore / 20));
    setAfter(String(initialAfter / 20));
    setLine(String(initialLine / 240));
    setMultiple(!getDefaultLineSpacingOptions().some((option) => option.twipsValue === initialLine));
    setLineChanged(false);
  }, [isOpen, initialBefore, initialAfter, initialLine]);
  if (!isOpen) return null;
  const validPoints = (value: string) => value.trim() !== '' && Number.isFinite(Number(value)) && Number(value) >= 0 && Number(value) <= 1584;
  const valid = validPoints(before) && validPoints(after) && (!lineChanged ||
    (line.trim() !== '' && Number.isFinite(Number(line)) && Number(line) > 0 && Number(line) <= 132));
  const apply = () => {
    if (!valid) return;
    onApply({ before: Math.round(Number(before) * 20), after: Math.round(Number(after) * 20),
      ...(lineChanged ? { lineSpacing: Math.round(Number(line) * 240) } : {}) });
    onClose();
  };
  return (
    <div style={overlayStyle} onClick={onClose} onMouseDown={(event) => event.stopPropagation()}
      onKeyDown={(event) => {
        if (event.key === 'Escape') { event.preventDefault(); event.stopPropagation(); onClose(); }
        if (event.key === 'Enter' && event.target instanceof HTMLInputElement && !event.nativeEvent.isComposing) {
          event.preventDefault(); event.stopPropagation(); apply();
        }
      }}>
      <form noValidate role="dialog" aria-modal="true" aria-labelledby={`${id}-title`} style={dialogStyle}
        onClick={(event) => event.stopPropagation()} onSubmit={(event) => { event.preventDefault(); apply(); }}>
        <div id={`${id}-title`} style={{ padding: '16px 20px 12px', borderBottom: '1px solid var(--doc-border)', fontSize: 16, fontWeight: 600 }}>
          {t('lineSpacing.paragraphSpacing')}
        </div>
        <div style={{ padding: '16px 20px', display: 'flex', flexDirection: 'column', gap: 14 }}>
          <div style={rowStyle}>
            <label htmlFor={`${id}-before`} style={labelStyle}>{t('dialogs.paragraphSpacing.before')}</label>
            <input id={`${id}-before`} type="number" min={0} max={1584} step={1} autoFocus required aria-describedby={`${id}-before-unit`}
              style={inputStyle} value={before} onChange={(event) => setBefore(event.target.value)} />
            <span id={`${id}-before-unit`}>{t('common.units.pt')}</span>
          </div>
          <div style={rowStyle}>
            <label htmlFor={`${id}-after`} style={labelStyle}>{t('dialogs.paragraphSpacing.after')}</label>
            <input id={`${id}-after`} type="number" min={0} max={1584} step={1} required aria-describedby={`${id}-after-unit`}
              style={inputStyle} value={after} onChange={(event) => setAfter(event.target.value)} />
            <span id={`${id}-after-unit`}>{t('common.units.pt')}</span>
          </div>
          <div style={rowStyle}>
            <label htmlFor={`${id}-line`} style={labelStyle}>{t('lineSpacing.label')}</label>
            <select id={`${id}-line`} style={inputStyle} value={multiple ? 'multiple' : line}
              onChange={(event) => {
                setMultiple(event.target.value === 'multiple');
                if (event.target.value !== 'multiple') { setLine(event.target.value); setLineChanged(true); }
              }}>
              {options.map((option) => <option key={option.twipsValue} value={String(option.value)}>
                {option.labelKey ? t(option.labelKey) : new Intl.NumberFormat(lang).format(option.value)}
              </option>)}
              <option value="multiple">{t('dialogs.paragraphSpacing.multiple')}</option>
            </select>
          </div>
          {multiple && <div style={rowStyle}>
            <label htmlFor={`${id}-multiple`} style={labelStyle}>{t('dialogs.paragraphSpacing.multiple')}</label>
            <input id={`${id}-multiple`} type="number" min={0.05} max={132} step={0.05} required
              style={inputStyle} value={line} onChange={(event) => { setLine(event.target.value); setLineChanged(true); }} />
          </div>}
        </div>
        <div style={{ padding: '12px 20px 16px', borderTop: '1px solid var(--doc-border)', display: 'flex', justifyContent: 'flex-end', gap: 8 }}>
          <button type="button" style={buttonStyle} onClick={onClose}>{t('common.cancel')}</button>
          <button type="submit" disabled={!valid} style={{ ...buttonStyle, backgroundColor: 'var(--doc-primary)', color: 'white', opacity: valid ? 1 : 0.5 }}>{t('common.apply')}</button>
        </div>
      </form>
    </div>
  );
}
