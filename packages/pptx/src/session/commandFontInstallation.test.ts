import { expect, test } from 'bun:test';
import type { DefaultFontFace, DefaultFontInstallRequest } from '../../../../shared/pptxDefaultFonts';
import { createCommandFontInstallation } from './commandFontInstallation';

test('browser installation preserves coverage kinds and acknowledges only installed indices', async () => {
  let request!: DefaultFontInstallRequest;
  const installation = createCommandFontInstallation((message, transfer) => {
    request = structuredClone(message, { transfer });
  });
  const faces: DefaultFontFace[] = [
    { family: 'Calibri', bold: false, italic: false, bytes: Uint8Array.of(1).buffer },
    { family: 'Noto Sans JP', bold: false, italic: false, bytes: Uint8Array.of(2).buffer, kind: 'coverage' },
  ];
  try {
    const pending = installation.install(faces);
    expect(request.faces.map((face) => face.kind)).toEqual([undefined, 'coverage']);
    expect(request.faces[1].bytes.byteLength).toBe(1);
    expect(faces.every((face) => face.bytes.byteLength === 1)).toBe(true);
    installation.receive({ kind: 'pptx-default-fonts-installed', id: request.id, installed: [1] });
    expect(await pending).toEqual([faces[1]]);
  } finally { installation.dispose(); }
});
