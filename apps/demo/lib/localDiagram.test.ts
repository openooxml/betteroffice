import { describe, expect, test } from 'bun:test';
import { readLocalDiagram } from './localDiagram';

const bytes = Uint8Array.of(0x50, 0x4b, 0x03, 0x04);

function picked(name: string, content = bytes): { name: string; arrayBuffer(): Promise<ArrayBuffer> } {
  return { name, arrayBuffer: async () => content.buffer.slice(0) as ArrayBuffer };
}

describe('readLocalDiagram', () => {
  test('carries the picked name and bytes when the engine accepts the file', async () => {
    expect(await readLocalDiagram(picked('plan.vsdx'), async () => {})).toEqual({
      opened: { name: 'plan.vsdx', bytes },
    });
  });

  test('refuses a file the engine cannot open and says the loaded diagram stays', async () => {
    const result = await readLocalDiagram(picked('notes.docx'), async () => {
      throw new Error('could not parse VSDX: unsupported VSDX document kind Docx');
    });
    expect(result).toEqual({
      refused:
        'Could not open “notes.docx”: could not parse VSDX: unsupported VSDX document kind Docx Previous diagram kept.',
    });
  });

  test('refuses a file whose bytes cannot be read', async () => {
    const unreadable = { name: 'gone.vsdx', arrayBuffer: () => Promise.reject(new Error('NotReadableError')) };
    expect(await readLocalDiagram(unreadable, async () => {})).toEqual({
      refused: 'Could not open “gone.vsdx”: NotReadableError Previous diagram kept.',
    });
  });

  test('reports a thrown non-Error as text', async () => {
    const result = await readLocalDiagram(picked('odd.vsdx'), () => {
      throw 'wasm trap';
    });
    expect(result).toEqual({ refused: 'Could not open “odd.vsdx”: wasm trap Previous diagram kept.' });
  });
});
