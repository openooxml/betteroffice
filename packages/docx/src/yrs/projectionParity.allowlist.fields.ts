import type { ProjectionAllowance } from './projectionParity.allowlist';

export const fields: readonly ProjectionAllowance[] = [
  { kind: 'field', path: 'blocks[].content[]', reason: 'Field lowering is deferred.' },
  { kind: 'hyperlink', path: 'blocks[].content[]', reason: 'Hyperlink marks are deferred.' },
  { kind: 'tab', path: 'blocks[].content[].content[]', reason: 'Tab lowering is deferred.' },
  { kind: 'break', path: 'blocks[].content[].content[]', reason: 'Break lowering is deferred.' },
  { kind: 'pageBreak', path: 'blocks[].content[].content[]', reason: 'Page break lowering is deferred.' },
  { kind: 'noteRef', path: 'blocks[].content[].content[]', reason: 'Note reference lowering is deferred.' },
  { kind: 'sdt', path: 'blocks[].content[]', reason: 'Inline content control lowering is deferred.' },
  { kind: 'math', path: 'blocks[].content[]', reason: 'Math lowering is deferred.' },
  { kind: 'image', path: 'blocks[].content[].content[]', reason: 'Drawing lowering is deferred.' },
  { kind: 'opaqueDrawing', path: 'blocks[].content[].content[]', reason: 'Opaque drawing lowering is deferred.' },
  { kind: 'shape', path: 'blocks[].content[].content[]', reason: 'Shape lowering is deferred.' },
];
