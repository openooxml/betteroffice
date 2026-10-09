export { DocxAgentDocument, openDocx, attachDocx } from './document';
export * from './types';
export { XlsxAgentWorkbook, openXlsx, type CellEdit, type CellChange } from './xlsx';
export { PptxAgentPresentation, openPptx, attachPptx, type SlideTextChange } from './pptx';
export type { AgentProposal, PrototypeOptions } from './prototype';

export type { PptxAgentEdit, PptxAgentBatch } from './pptx-schema';
export { renderPptxSlide } from './pptx-render';
