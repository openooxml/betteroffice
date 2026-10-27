import type { PresentationCommandObjectSelection } from '@betteroffice/pptx/worker';
import { useSyncExternalStore } from 'react';
import { handleAnchor, RESIZE_HANDLES, resizeCursor } from '../interactions';
import type { PeerlessShapePreviewStore } from './usePeerlessShapeGestures';
import { useTranslation } from '../i18n';

export function PeerlessShapeOverlay({ selection, preview: previews, scale, readOnly }: {
  selection: PresentationCommandObjectSelection; preview: PeerlessShapePreviewStore; scale: number; readOnly: boolean;
}) {
  const { t } = useTranslation();
  const preview = useSyncExternalStore(previews.subscribe, previews.getSnapshot, previews.getSnapshot);
  const original = selection.bounds;
  const bounds = preview?.token === selection.token ? preview.bounds : original;
  return <>
    <div data-testid="pptx-shape-selection" style={{ position: 'absolute', pointerEvents: 'none',
      left: original.x * scale, top: original.y * scale, width: original.width * scale, height: original.height * scale,
      border: '2px solid #2563eb', boxSizing: 'border-box', transformOrigin: '0 0',
      transform: `translate(${(bounds.x - original.x) * scale}px, ${(bounds.y - original.y) * scale}px) scale(${original.width ? bounds.width / original.width : 1}, ${original.height ? bounds.height / original.height : 1})` }} />
    {!readOnly && selection.canResize && RESIZE_HANDLES.map((handle) => {
      const point = handleAnchor(bounds, handle);
      return <span key={handle} data-testid={`pptx-resize-${handle}`} data-pptx-resize={handle}
        role="img" aria-label={t(`objects.resize.${handle}`)}
        style={{ position: 'absolute', pointerEvents: 'auto', left: point.x * scale - 4, top: point.y * scale - 4,
          width: 8, height: 8, borderRadius: '50%', border: '1px solid #2563eb', background: '#fff',
          boxSizing: 'border-box', cursor: resizeCursor(handle), touchAction: 'none' }} />;
    })}
  </>;
}
