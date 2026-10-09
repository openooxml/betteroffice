export { DocxAgentDocument, openDocx, attachDocx } from './document';
export * from './types';
export { XlsxAgentWorkbook, openXlsx, attachXlsx, type CellEdit, type CellChange, type XlsxOptions } from './xlsx';
export { PptxAgentPresentation, openPptx, attachPptx, type SlideTextChange } from './pptx';
export type { AgentProposal, PrototypeOptions } from './prototype';
export type { XlsxScalar, XlsxWriteEdit, XlsxGridEdit, XlsxSheetEdit, XlsxFormatEdit } from './xlsx-operations';
export type { PptxAgentEdit, PptxAgentBatch } from './pptx-schema';
export { renderPptxSlide } from './pptx-render';
