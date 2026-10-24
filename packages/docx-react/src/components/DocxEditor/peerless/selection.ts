import type {
  DocxNavigationIntent,
  DocxCommandSelection,
  DocxCommandPayloads,
  DocxSelectionPointIntent,
} from '@betteroffice/docx/yrs';
import type { PeerlessFrameGeometry } from './viewStore';

const pageIdentities = new WeakMap<PeerlessFrameGeometry['pages'], ReadonlyMap<string, DocxSelectionPointIntent['page']>>();

/** @internal */
export type PeerlessSelectionIntent =
  | { kind: 'select'; payload: { selection: DocxCommandSelection } }
  | { kind: 'select'; payload: { object: { handle: string; expectVersion: string } } }
  | { kind: 'navigateCell'; payload: DocxCommandPayloads['navigateCell'] }
  | { kind: 'selectTableCells'; payload: DocxCommandPayloads['selectTableCells'] }
  | { kind: 'navigate'; payload: DocxNavigationIntent }
  | { kind: 'selectPoint'; payload: DocxSelectionPointIntent }
  | { kind: 'extendSelection'; payload: DocxSelectionPointIntent };

/** @internal Gestures use installed page geometry, never echo hit-testing. */
export type PeerlessSelectionGesture =
  | { kind: 'navigation'; navigation: DocxNavigationIntent }
  | { kind: 'point' | 'drag'; point: DocxSelectionPointIntent };

/** @internal */
export function captureSelectionIntent(
  gesture: PeerlessSelectionGesture,
  geometry: PeerlessFrameGeometry | null
): PeerlessSelectionIntent {
  if (gesture.kind === 'navigation') return { kind: 'navigate', payload: { ...gesture.navigation } };
  const point = gesture.point;
  let page = point.page;
  if (geometry) {
    let identities = pageIdentities.get(geometry.pages);
    if (!identities) {
      identities = new Map<string, DocxSelectionPointIntent['page']>(geometry.pages.map(({ page }) => [page.pageId, page]));
      pageIdentities.set(geometry.pages, identities);
    }
    page = identities.get(point.page.pageId) ?? page;
  }
  return {
    kind: gesture.kind === 'drag' ? 'extendSelection' : 'selectPoint',
    payload: {
      ...point, page: { ...page }, modifiers: { ...point.modifiers },
      frameEpoch: geometry?.frameEpoch ?? point.frameEpoch,
      layoutRevision: geometry?.layoutRevision ?? point.layoutRevision,
    },
  };
}
