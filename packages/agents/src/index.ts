export { DocxAgentDocument, openDocx, attachDocx } from './document';
export * from './types';
export { XlsxAgentWorkbook, openXlsx, attachXlsx, type CellEdit, type CellChange, type XlsxOptions } from './xlsx';
export { PptxAgentPresentation, openPptx, type SlideTextChange } from './pptx';
export type { AgentProposal, PrototypeOptions } from './prototype';
export type { XlsxScalar, XlsxWriteEdit, XlsxGridEdit, XlsxSheetEdit, XlsxFormatEdit } from './xlsx-operations';
