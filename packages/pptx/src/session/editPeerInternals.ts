import type { PresentationEditPeer } from './editPeer';

export const presentationEditPeerInternals = new WeakMap<PresentationEditPeer, {
  fail(cause: unknown): void;
  whenAcknowledged(): Promise<void>;
}>();
