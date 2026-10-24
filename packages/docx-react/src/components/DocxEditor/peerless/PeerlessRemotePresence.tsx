import { useEffect, useLayoutEffect, useRef, useState, useSyncExternalStore } from 'react';
import { resolvePresenceColor, type CollaborationPeer, type CollaborationPresence } from '@betteroffice/docx/collaboration';
import type { ResidentEngineWorkerClient, DocxPresenceGeometry } from '@betteroffice/docx/yrs';
import type { PeerlessEditSessionBinding } from '../hooks/usePeerlessEditSession';
import type { PeerlessFrameGeometry } from './viewStore';
import { simplifyRemoteSelectionRects } from '../overlays/remotePresenceGeometry';

const now = () => performance.now();
const sameIdentity = (left: PeerlessFrameGeometry['identity'], right: PeerlessFrameGeometry['identity']) =>
  left.session === right.session && left.incarnation === right.incarnation && left.epoch === right.epoch;
const FLAG_VISIBLE_MS = 3_000;

export function PeerlessRemotePresence({ binding, client, presence, pages }: {
  binding: PeerlessEditSessionBinding; client: ResidentEngineWorkerClient;
  presence: CollaborationPresence; pages: readonly [number, number];
}) {
  const snapshot = useSyncExternalStore(binding.viewStore.subscribe, binding.viewStore.getSnapshot, binding.viewStore.getServerSnapshot);
  const [remote, setRemote] = useState({ presence, peers: presence.peers });
  const peers = remote.presence === presence ? remote.peers : presence.peers;
  const input = useRef({ presence, peers, pages });
  useLayoutEffect(() => { input.current = { presence, peers, pages }; }, [presence, peers, pages[0], pages[1]]);
  const scheduleRef = useRef<(() => void) | null>(null);
  useEffect(() => {
    setRemote({ presence, peers: presence.peers });
    return presence.onPeers(peers => setRemote({ presence, peers }));
  }, [presence]);
  const [paint, setPaint] = useState<{ presence: CollaborationPresence; geometry: PeerlessFrameGeometry; reply: DocxPresenceGeometry;
    peers: readonly CollaborationPeer[]; pages: readonly [number, number] } | null>(null);
  const [flagNow, setFlagNow] = useState(now);
  useEffect(() => {
    let previousVersion = binding.viewStore.getSnapshot().version;
    let previousKey: string | undefined;
    const publish = () => {
      const view = binding.viewStore.getSnapshot();
      const selection = view.workerSelection;
      const cursor = selection && view.anchors ? { story: selection.anchor.story, anchor: view.anchors.anchor, head: view.anchors.focus } : null;
      const key = cursor ? `${cursor.story}:${Array.from(cursor.anchor)}:${Array.from(cursor.head)}` : 'null';
      const changed = view.version !== previousVersion; previousVersion = view.version;
      if (key === previousKey) return;
      previousKey = key; presence.setCursor(cursor, !changed);
    };
    publish(); const off = binding.viewStore.subscribe(publish);
    return () => { off(); presence.setCursor(null); };
  }, [binding, presence]);
  useEffect(() => {
    let disposed = false, queued = false, busy = false, dirty = false;
    const invalid = new Map<string, string>();
    const schedule = () => {
      dirty = true;
      if (disposed || queued || busy) return;
      queued = true;
      queueMicrotask(() => { queued = false; if (!disposed) void read(); });
    };
    const read = async () => {
      dirty = false;
      const view = binding.viewStore.getSnapshot();
      const geometry = view.geometry;
      if (!geometry?.documentVersion) { setPaint(null); return; }
      const request = input.current;
      const selectedPeers = request.peers;
      const window = [request.pages[0], request.pages[1]] as const;
      const liveKeys = new Set(selectedPeers.map(peer => `${peer.clientId}:${peer.clock}`));
      for (const key of invalid.keys()) if (!liveKeys.has(key)) invalid.delete(key);
      const current = () => !disposed && geometry === binding.viewStore.getSnapshot().geometry && view.version === binding.viewStore.getSnapshot().version &&
        selectedPeers === input.current.peers && request.presence === input.current.presence && window[0] === input.current.pages[0] && window[1] === input.current.pages[1];
      busy = true;
      try {
        const result = await client.documentReadAt({ kind: 'collabPresence',
          peers: selectedPeers.filter(peer => peer.inferredCursor || invalid.get(`${peer.clientId}:${peer.clock}`) !== geometry.documentVersion).slice(0, 32).map(peer => ({
            clientId: peer.clientId, clock: peer.clock, ...(peer.cursor ? { cursor: peer.cursor } : {}),
            ...(peer.inferredCursor ? { inferred: { story: peer.inferredCursor.story, paraId: peer.inferredCursor.paraId, offset: peer.inferredCursor.endOffset } } : {}),
          })), pages: [window[0], window[1]], expectVersion: geometry.documentVersion,
          expectFrameEpoch: geometry.frameEpoch, expectLayoutRevision: geometry.layoutRevision }, geometry.documentVersion);
        const reply = result.status === 'ok' ? result.value : null;
        if (!current()) return;
        if (!reply || reply.version !== geometry.documentVersion || reply.frameEpoch !== geometry.frameEpoch ||
          reply.layoutRevision !== geometry.layoutRevision) { setPaint(null); return; }
        for (const peer of reply.peers) if (peer.status === 'unresolved' && !selectedPeers.find(candidate => candidate.clientId === peer.clientId)?.inferredCursor) invalid.set(`${peer.clientId}:${peer.clock}`, reply.version);
        setPaint({ presence: request.presence, geometry, reply, peers: selectedPeers, pages: window });
      } catch { if (current()) setPaint(null); } finally { busy = false; if (dirty) schedule(); }
    };
    const offView = binding.viewStore.subscribe(schedule);
    scheduleRef.current = schedule; schedule();
    return () => { disposed = true; offView(); if (scheduleRef.current === schedule) scheduleRef.current = null; };
  }, [binding, client]);
  useEffect(() => { scheduleRef.current?.(); }, [presence, peers, pages[0], pages[1]]);
  useEffect(() => {
    const deadlines = paint ? peers.map(peer => peer.cursorMovedAt + FLAG_VISIBLE_MS).filter(deadline => deadline > now()) : [];
    if (!deadlines.length) return;
    const timer = setTimeout(() => setFlagNow(now()), Math.max(0, Math.min(...deadlines) - now()));
    return () => clearTimeout(timer);
  }, [paint, peers, flagNow]);
  // the last painted read stays up while the next one is in flight, for peers still present in the same document
  const held = snapshot.geometry && paint?.presence === presence ? paint : null;
  const shown = held && sameIdentity(held.geometry.identity, snapshot.geometry!.identity) ? held : null;
  const project = (rect: NonNullable<DocxPresenceGeometry['peers'][number]['caret']>) => {
    const page = snapshot.geometry?.pages.find(page => page.page.pageId === rect.pageId && page.page.pageIndex === rect.pageIndex);
    return page ? { left: page.bounds.x + rect.x, top: page.bounds.y + rect.y, width: rect.width, height: rect.height } : null;
  };
  return <div data-testid="remote-presence-overlay" aria-hidden="true" style={{ position: 'absolute', inset: 0, pointerEvents: 'none', overflow: 'visible', zIndex: 30 }}>
    {shown?.reply.peers.flatMap(geometry => {
      const peer = peers.find(peer => peer.clientId === geometry.clientId);
      if (!peer || geometry.status !== 'ok') return [];
      const color = resolvePresenceColor(peer.clientId, peer.user.color);
      const flagVisible = peer.cursorMovedAt > 0 && Math.max(flagNow, now()) - peer.cursorMovedAt < FLAG_VISIBLE_MS;
      const caret = geometry.caret && project(geometry.caret);
      const rects = simplifyRemoteSelectionRects(geometry.rects, { start: pages[0], end: pages[1] - 1 });
      return [
        ...rects.flatMap((rect, index) => {
          const pageId = geometry.rects.find(original => original.pageIndex === rect.pageIndex)?.pageId;
          const projected = pageId && project({ ...rect, pageId });
          return projected ? [<div key={`${peer.clientId}:range:${index}`} style={{ position: 'absolute', ...projected, boxSizing: 'border-box',
            backgroundColor: `color-mix(in srgb, ${color} 18%, transparent)`, border: `1px solid color-mix(in srgb, ${color} 28%, transparent)` }} />] : [];
        }),
        caret ? <div key={`${peer.clientId}:caret`} data-testid="remote-presence-caret" data-client-id={peer.clientId} style={{ position: 'absolute', ...caret, left: caret.left - 1, width: 2, borderRadius: 1, backgroundColor: color }}>
          <div style={{ position: 'absolute', left: 0, bottom: '100%', maxWidth: flagVisible ? 180 : 0, padding: flagVisible ? '3px 6px' : 0,
            borderRadius: '4px 4px 4px 0', overflow: 'hidden', opacity: flagVisible ? 1 : 0, color: '#FFFFFF', backgroundColor: color,
            font: '600 11px/16px system-ui, sans-serif', whiteSpace: 'nowrap', transform: flagVisible ? 'translateY(0)' : 'translateY(2px)',
            transition: 'opacity 160ms ease, transform 160ms ease, max-width 160ms ease, padding 160ms ease' }}>{peer.user.name}</div>
        </div> : null,
      ];
    })}
  </div>;
}
