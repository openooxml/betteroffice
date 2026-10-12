export function residentMediaImages(engine: { mediaDataUrl(token: string): string | null }) {
  const images = new Map<string, Promise<ImageBitmap | null>>();
  const resolve = (source: string): Promise<ImageBitmap | null> => {
    let image = images.get(source);
    if (!image) {
      image = (async () => {
        const url = /^media:(0|[1-9]\d*)$/.test(source) ? engine.mediaDataUrl(source) : source;
        if (!url?.startsWith('data:image/')) return null;
        return createImageBitmap(await (await fetch(url)).blob());
      })().catch(error => {
        if (error instanceof WebAssembly.RuntimeError) throw error;
        return null;
      });
      images.set(source, image);
    }
    return image;
  };
  const close = async (): Promise<void> => {
    for (const result of await Promise.allSettled(images.values())) {
      if (result.status === 'fulfilled') result.value?.close();
    }
    images.clear();
  };
  return { resolve, close };
}
