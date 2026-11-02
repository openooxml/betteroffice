import type { DefaultFontFace, DefaultFontInstallRequest } from '../../../../shared/pptxDefaultFonts';

export function createCommandFontInstallation(send: (request: DefaultFontInstallRequest, transfer: Transferable[]) => void) {
  let nextId = 0;
  const pending = new Map<number, (installed: readonly number[]) => void>();
  return {
    install(faces: readonly DefaultFontFace[]): Promise<DefaultFontFace[]> {
      return new Promise((resolve) => {
        const id = ++nextId;
        const timer = setTimeout(() => finish([]), 5000);
        const finish = (installed: readonly number[]) => {
          if (!pending.delete(id)) return;
          clearTimeout(timer);
          resolve(faces.filter((_, index) => installed.includes(index)));
        };
        pending.set(id, finish);
        const copies = faces.map((face) => ({ ...face, bytes: face.bytes.slice(0) }));
        try { send({ id, faces: copies }, copies.map((face) => face.bytes)); }
        catch { finish([]); }
      });
    },
    receive(message: unknown): boolean {
      if (!message || typeof message !== 'object' || !('kind' in message) || message.kind !== 'pptx-default-fonts-installed') return false;
      if ('id' in message && typeof message.id === 'number' && 'installed' in message && Array.isArray(message.installed)) {
        pending.get(message.id)?.(message.installed.filter((index): index is number => Number.isSafeInteger(index) && index >= 0));
      }
      return true;
    },
    dispose(): void { for (const finish of pending.values()) finish([]); },
  };
}
