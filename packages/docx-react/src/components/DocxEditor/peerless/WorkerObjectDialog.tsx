import { useState } from 'react';
import type { PeerlessPropertyIntent } from './inputEvents';
import { useTranslation } from '../../../i18n';

export function WorkerObjectDialog({ action, target, apply, close }: {
  action: 'resizeImage' | 'moveImage' | 'resizeColumn';
  target: { handle: string; version: string; x: number; y: number; width: number; height: number };
  apply(intent: PeerlessPropertyIntent): Promise<boolean>;
  close(): void;
}) {
  const { t } = useTranslation();
  const moving = action === 'moveImage';
  const label = t(moving ? 'contextMenu.moveImage' : action === 'resizeImage' ? 'contextMenu.resizeImage' : 'contextMenu.columnWidth');
  const [width, setWidth] = useState(moving ? target.x : target.width);
  const [height, setHeight] = useState(moving ? target.y : target.height);
  const [busy, setBusy] = useState(false);
  const valid = Number.isFinite(width) && (moving || width > 0) &&
    (action === 'resizeColumn' || Number.isFinite(height) && (moving || height > 0));
  const submit = async () => {
    if (!valid || busy) return;
    setBusy(true);
    const base = { handle: target.handle, expectVersion: target.version };
    const intent: PeerlessPropertyIntent = moving ? { kind: 'moveObject', payload: { ...base, hOffsetEmu: width * 9525, vOffsetEmu: height * 9525 } } :
      { kind: 'resizeObject', payload: action === 'resizeImage' ? { ...base, widthEmu: width * 9525, heightEmu: height * 9525 } : { ...base, widthTwips: Math.round(width * 15) } };
    try { if (await apply(intent)) close(); } finally { setBusy(false); }
  };
  return <div data-docx-escape-layer="true" style={{ position: 'fixed', inset: 0, zIndex: 1000, display: 'grid', placeItems: 'center', background: '#0004' }}>
    <form role="dialog" aria-modal="true" aria-label={label} onSubmit={event => { event.preventDefault(); void submit(); }}
      onKeyDown={event => { if (event.key === 'Escape' && !busy) { event.stopPropagation(); close(); } }}
      style={{ display: 'grid', gap: 12, padding: 24, borderRadius: 8, background: 'var(--doc-bg)', color: 'var(--doc-text)' }}>
      <strong>{label}</strong>
      <label>{t(moving ? 'dialogs.objectGeometry.left' : 'dialogs.objectGeometry.width')}<input type="number" autoFocus required value={width} disabled={busy}
        onChange={event => setWidth(event.target.valueAsNumber)} /></label>
      {action !== 'resizeColumn' && <label>{t(moving ? 'dialogs.objectGeometry.top' : 'dialogs.objectGeometry.height')}<input type="number" required value={height} disabled={busy}
        onChange={event => setHeight(event.target.valueAsNumber)} /></label>}
      <div><button type="button" disabled={busy} onClick={close}>{t('common.cancel')}</button><button type="submit" disabled={!valid || busy}>{t('common.apply')}</button></div>
    </form>
  </div>;
}
