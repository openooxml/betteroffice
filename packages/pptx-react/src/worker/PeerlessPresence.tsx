import { useEffect, useState, useSyncExternalStore } from 'react';
import type { PptxPresence, PptxPresencePeer } from '@betteroffice/pptx';
import { indexHitGeometry, transformedRect } from '@betteroffice/pptx/peerless';
import { groupShapePresence, limitPresence } from '../presence-rendering';
import type { PeerlessPresentation } from './usePeerlessPresentation';

export function PeerlessPresence({ run, presence }: { run: PeerlessPresentation; presence: PptxPresence }) {
  const [peers, setPeers] = useState<readonly PptxPresencePeer[]>([]);
  const visual = useSyncExternalStore(run.subscribeVisual, run.getVisualSnapshot, run.getVisualSnapshot);
  useEffect(() => {
    setPeers(presence.peers);
    return presence.onPresence(setPeers);
  }, [presence, run]);
  useEffect(() => {
    const update = () => {
      const view = run.store.getSnapshot();
      const slideId = view.confirmed?.frame.slideId;
      presence.setCursor(slideId ? { slideId, ...(view.selection?.shapeId ? { shapeId: view.selection.shapeId } : {}) } : null);
    };
    update();
    const off = run.store.subscribePresentation(update);
    return () => { off(); presence.setCursor(null); };
  }, [presence, run]);
  const frame = visual?.view.confirmed;
  const geometry = frame ? indexHitGeometry(frame.displayList) : null;
  const groups = frame && geometry ? groupShapePresence(peers, frame.frame.slideId,
    new Set(geometry.hits.flatMap((hit) => hit.primitive.shapeId ? [hit.primitive.shapeId] : []))) : null;
  const bounded = limitPresence(peers);
  return <>
    <div role="list" aria-label="Collaborators" style={{ position: 'absolute', right: 8, top: 8, zIndex: 2 }}>
      {bounded.visible.map((peer) => <span key={peer.state.clientId} role="listitem" style={{ color: peer.state.user.color }}>
        {peer.state.user.name}
      </span>)}
      {bounded.overflow > 0 && <span>+{bounded.overflow}</span>}
    </div>
    {groups?.visible.map((group) => {
      const hit = geometry!.hits.find((hit) => hit.primitive.shapeId === group.shapeId);
      if (!hit) return null;
      const bounds = transformedRect(hit.transforms, { x: hit.primitive.x, y: hit.primitive.y,
        width: hit.primitive.w, height: hit.primitive.h });
      const scale = visual!.fit.scale;
      return <div key={group.shapeId} data-testid="pptx-remote-shape" aria-hidden="true" style={{
        position: 'absolute', pointerEvents: 'none', boxSizing: 'border-box', border: `2px solid ${group.peer.state.user.color}`,
        left: bounds.x * scale, top: bounds.y * scale, width: bounds.width * scale, height: bounds.height * scale }} />;
    })}
  </>;
}
