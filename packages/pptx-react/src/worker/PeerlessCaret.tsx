import { findLine, indexGeometry, transformedRect } from '@betteroffice/pptx/peerless';
import type { CommandRect, SlidePrimitive } from '@betteroffice/pptx';
import { useMemo, useSyncExternalStore } from 'react';
import { useCaretBlink } from '../useCaretBlink';
import { echoLineClip, measureEchoPosition } from './localEchoPainter';
import type { PeerlessVisualStore } from './usePeerlessPresentation';

export function PeerlessCaret({ store }: { store: PeerlessVisualStore }) {
  const view = useSyncExternalStore(store.subscribePresentation, store.getSnapshot, store.getSnapshot);
  const eligible = view.focused && view.caret !== null && view.selection !== null &&
    view.selection.anchor === view.selection.focus;
  const visible = useCaretBlink(eligible, view.caret);
  const active = useMemo(() => view.activeLine ?? (view.confirmed
    ? findLine(indexGeometry(view.confirmed.displayList), view.selection) : null),
    [view.activeLine, view.confirmed, view.selection]);
  if (!view.confirmed) return null;
  let x = view.caret?.x ?? 0;
  let y = view.caret?.y ?? 0;
  let height = view.caret?.height ?? 0;
  let transform = 'none';
  if (active && view.selection) {
    const position = view.selection.focus;
    const first = active.line.caretStops[0];
    const stop = first && active.line.caretStops.reduce((best, candidate) =>
      Math.abs(candidate.position - position) < Math.abs(best.position - position)
        ? candidate : best, first);
    const localX = view.echo && view.activeLine
      ? measureEchoPosition(view.activeLine, view.echo, view.echo.caretPosition) ?? stop?.x ?? active.line.x
      : stop?.x ?? active.line.x;
    const point = (dx: number, dy: number) => transformedRect(active.transforms,
      { x: localX + dx, y: active.line.y + dy, width: 0, height: 0 });
    const origin = point(0, 0);
    const horizontal = point(1, 0);
    const vertical = point(0, 1);
    x = origin.x;
    y = origin.y;
    height = active.line.height;
    transform = `matrix(${horizontal.x - x}, ${horizontal.y - y}, ${vertical.x - x}, ${vertical.y - y}, 0, 0)`;
  }

  let overlay = <div aria-hidden="true" data-testid="pptx-peerless-caret-clip" style={{
    position: 'absolute', left: 0, top: 0, pointerEvents: 'none', overflow: 'hidden',
    width: view.confirmed.displayList.width, height: view.confirmed.displayList.height,
    clipPath: view.echo && active ? clipPolygon(active.transforms, echoLineClip(active)) : undefined,
  }}>
    <div data-testid="pptx-peerless-caret" aria-hidden="true" style={{
      position: 'absolute', pointerEvents: 'none', background: '#111827',
      left: x, top: y, width: 1, height, transform, transformOrigin: '0 0',
      visibility: eligible && visible ? 'visible' : 'hidden',
    }} />
  </div>;
  const transforms = active?.transforms ?? [];
  for (let index = transforms.length - 1; index >= 0; index -= 1) {
    const primitive = transforms[index];
    if (primitive.kind !== 'chart' && primitive.kind !== 'table') continue;
    overlay = <div key={index} aria-hidden="true" style={{
      position: 'absolute', left: 0, top: 0, pointerEvents: 'none',
      width: view.confirmed.displayList.width, height: view.confirmed.displayList.height,
      clipPath: view.echo ? clipPolygon(transforms.slice(0, index + 1), {
        x: primitive.x, y: primitive.y, width: primitive.w, height: primitive.h,
      }) : undefined,
    }}>{overlay}</div>;
  }
  return overlay;
}

function clipPolygon(transforms: readonly SlidePrimitive[], rect: CommandRect): string {
  const points = [[rect.x, rect.y], [rect.x + rect.width, rect.y],
    [rect.x + rect.width, rect.y + rect.height], [rect.x, rect.y + rect.height]];
  return `polygon(${points.map(([x, y]) => {
    const point = transformedRect(transforms, { x, y, width: 0, height: 0 });
    return `${point.x}px ${point.y}px`;
  }).join(', ')})`;
}
