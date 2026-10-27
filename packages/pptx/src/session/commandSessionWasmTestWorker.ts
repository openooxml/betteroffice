import './worker';

declare const testHarness: { open(bytes: Uint8Array): Record<string, unknown> };

const open = testHarness.open;
testHarness.open = (bytes) => ({
  ...open(bytes),
  onUpdate: () => () => {},
  setUndoCaptureMode() {},
  readCommandContent: () => ({ ok: true, slides: [], stories: [], sources: [] }),
  encodeStateAsUpdate: () => new Uint8Array(),
  isProposalsAvailable: () => true,
  listProposals: () => [],
  canUndo: () => false,
  canRedo: () => false,
});
