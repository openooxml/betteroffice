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
      left: bounds.x * scale, top: bounds.y * scale,
      width: Math.max(1, bounds.width * scale), height: Math.max(1, bounds.height * scale),
      border: '2px solid #2563eb', boxSizing: 'border-box' }} />
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
