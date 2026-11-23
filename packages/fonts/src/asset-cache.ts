type CachedAsset = { promise: Promise<ArrayBuffer> };
interface FontAssetState {
  assets: Map<string, CachedAsset>;
}
const sharedKey = Symbol.for('betteroffice.fonts.assets.v3');
type FontAssetScope = { [sharedKey]?: FontAssetState };
const state: FontAssetState = (globalThis as FontAssetScope)[sharedKey] ??= {
  assets: new Map(),
};
const { assets } = state;

export async function cachedFontAsset(key: string, read: () => Promise<ArrayBuffer>): Promise<ArrayBuffer> {
  let entry = assets.get(key);
  if (!entry) {
    const current: CachedAsset = { promise: Promise.resolve().then(read) };
    const release = () => { if (assets.get(key) === current) assets.delete(key); };
    current.promise.then(release, release);
    entry = current;
    assets.set(key, entry);
  }
  // Callers own the result and may transfer it to a decoder worker.
  return (await entry.promise).slice(0);
}
