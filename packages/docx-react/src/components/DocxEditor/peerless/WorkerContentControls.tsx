import { useState, useSyncExternalStore } from 'react';
import type { YrsContentControlValue } from '@betteroffice/docx/yrs';
import { useWorkerQuery } from './useWorkerQuery';
import { applyWorkerProperty, type WorkerShellController } from './workerProperties';

export function WorkerContentControls({ controller, onError }: { controller: WorkerShellController; onError?: (error: Error) => void }) {
  const state = useSyncExternalStore(controller.subscribe, controller.getSnapshot, controller.getSnapshot);
  const query = useWorkerQuery<'contentControlHandles'>(controller, { kind: 'contentControlHandles' });
  const [busy, setBusy] = useState<string | null>(null);
  const geometry = state.view?.geometry;
  const env = controller.environment();
  if (!query || !geometry || query.value.frameEpoch !== geometry.frameEpoch || query.value.layoutRevision !== geometry.layoutRevision ||
    geometry.coversThroughSeq < (state.view?.throughSeq ?? 0)) return null;
  const change = async (handle: string, value: YrsContentControlValue | 'add' | 'remove') => {
    if (busy) return;
    setBusy(handle);
    try {
      const payload = { handle, expectVersion: query.version };
      await applyWorkerProperty(controller, query, value === 'add' ? { kind: 'addRepeatingItem', payload } :
        value === 'remove' ? { kind: 'removeRepeatingItem', payload } : { kind: 'setContentControlValue', payload: { ...payload, value } }, onError);
    } finally { setBusy(null); }
  };
  return <>{query.value.controls.flatMap(target => {
    const widget = target.widget;
    if (!widget) return [];
    const label = target.control.alias ?? target.control.tag ?? env.translate('a11y.contentControl');
    const locked = target.control.effectiveLock;
    const disabled = busy !== null || env.readOnly || env.mode !== 'editing' || !locked.known || locked.content || target.control.dataBound;
    return target.rects.flatMap((rect, index) => {
      const page = geometry.pages.find(page => page.page.pageId === rect.pageId);
      if (!page) return [];
      const repeating = widget.kind === 'repeatingItem';
      return [<div key={`${target.handle}:${index}`} data-worker-control={target.handle} data-docx-escape-layer
        style={{ position: 'absolute', left: page.bounds.x + rect.x - (repeating ? 48 : 0), top: page.bounds.y + rect.y,
          width: repeating ? 44 : rect.width, minHeight: repeating ? 22 : rect.height, display: repeating ? 'flex' : undefined,
          zIndex: 12, pointerEvents: 'none' }}>
        {widget.kind === 'checkbox' ? <input type="checkbox" aria-label={label} checked={widget.checked === true} disabled={disabled}
          style={{ width: rect.width, height: rect.height, margin: 0, pointerEvents: 'auto' }} onChange={event => { void change(target.handle, { kind: 'checkbox', checked: event.target.checked }); }} /> :
          widget.kind === 'dropdown' ? <select aria-label={label} value={widget.value ?? ''} disabled={disabled}
            style={{ minWidth: rect.width, height: Math.max(20, rect.height), pointerEvents: 'auto' }} onChange={event => { void change(target.handle, { kind: 'dropdown', value: event.target.value }); }}>
            {!widget.items?.some(item => item.value === widget.value) && <option value={widget.value ?? ''}>{widget.value}</option>}
            {widget.items?.map(item => <option key={item.value} value={item.value}>{item.displayText}</option>)}
          </select> : widget.kind === 'date' ? <input type="date" aria-label={label} value={widget.value ?? ''} disabled={disabled}
            style={{ minWidth: rect.width, height: Math.max(20, rect.height), pointerEvents: 'auto' }} onChange={event => { if (event.target.value) void change(target.handle, { kind: 'date', date: event.target.value }); }} /> :
            <><button type="button" aria-label={env.translate('a11y.addRepeatingItem')} style={{ width: 22, height: 22, pointerEvents: 'auto' }} disabled={disabled || locked.control} onClick={() => { void change(target.handle, 'add'); }}>+</button>
              <button type="button" aria-label={env.translate('a11y.removeRepeatingItem')} style={{ width: 22, height: 22, pointerEvents: 'auto' }} disabled={disabled || locked.control} onClick={() => { void change(target.handle, 'remove'); }}>−</button></>}
      </div>];
    });
  })}</>;
}
