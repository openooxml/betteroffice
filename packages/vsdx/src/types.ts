export interface Affine { a: number; b: number; c: number; d: number; e: number; f: number; }
export interface CellLocator { section?: string; rowIndex?: number; rowName?: string; cellName: string; }
export interface CellSnapshot { locator: CellLocator; name: string; formula: string | null; value: string | null; }
export interface ShapeSnapshot { id: string; sourceId: number; name: string | null; cells: CellSnapshot[]; children: ShapeSnapshot[]; }
export interface PageSnapshot { id: string; sourcePartPath: string; name: string | null; shapes: ShapeSnapshot[]; }
export interface DiagramSnapshot { pages: PageSnapshot[]; }
export interface CellFormulaReceipt { pageId: string; shapeId: string; cellName: string; before: string | null; after: string; }
export interface ShapeReceipt { pageId: string; shapeId: string; fromIndex: number | null; toIndex: number | null; }
export interface FormulaShapeDraft { sourceId: number; name?: string; cells: Array<{ locator: unknown; name: string; formula?: string }> }
export interface VsdxFontFace { family: string; bold?: boolean; italic?: boolean; bytes: Uint8Array; }
export type Paint = { kind: 'solid'; color: string } | { kind: 'gradient'; stops: Array<{ position: number; color: string }> };
export interface Stroke { color: string; width: number; dashed?: boolean; }
export interface GeometryPathCommand { type: string; [key: string]: number | string; }
export interface TextDiagnostic { category: 'Integrity' | 'Fidelity'; code: string; detail: string; }
export interface TextRun { text: string; family: string; sizeIn: number; bold: boolean; italic: boolean; underline: boolean; smallCaps: boolean; superscript: boolean; subscript: boolean; letterSpacing: number; color: string; diagnostics: TextDiagnostic[]; }
export interface TextParagraph { runs: TextRun[]; }
export interface PositionedLine { x: number; y: number; width: number; height: number; start: number; end: number; caretStops: Array<{ position: number; x: number; y: number }>; }
interface PrimitiveBase { id: string; zOrder: number; }
export interface ShapePrimitive extends PrimitiveBase { kind: 'shape'; path: GeometryPathCommand[]; fill?: Paint; stroke?: Stroke; transform?: Affine; }
export interface ImagePrimitive extends PrimitiveBase { kind: 'image'; assetId: string; x: number; y: number; width: number; height: number; transform?: Affine; }
export interface TextBoxPrimitive extends PrimitiveBase { kind: 'textBox'; x: number; y: number; width: number; height: number; paragraphs: TextParagraph[]; lines: PositionedLine[]; }
export interface PlaceholderPrimitive extends PrimitiveBase { kind: 'placeholder'; x: number; y: number; width: number; height: number; reason: string; }
export interface GroupPrimitive extends PrimitiveBase { kind: 'group'; primitives: PagePrimitive[]; transform?: Affine; }
export type PagePrimitive = ShapePrimitive | ImagePrimitive | TextBoxPrimitive | PlaceholderPrimitive | GroupPrimitive;
export interface PageDisplayList { contractVersion: 3; width: number; height: number; paintTransform: Affine; primitives: PagePrimitive[]; }
export type HitTestResult = { kind: 'shape'; shapeId: string } | { kind: 'text'; shapeId: string; position: number };
export interface HistoryResult { applied: boolean; snapshot: DiagramSnapshot; }
export type CollaborationUpdateOrigin = 'local' | 'remote';
