export interface Affine { a: number; b: number; c: number; d: number; e: number; f: number; }
export interface CellLocator { section?: string; sectionIndex?: number; rowIndex?: number; rowName?: string; cellName: string; }
export type SnapshotCellSheet = 'document' | { page: number } | { master: number };
export type SnapshotCellRow = { index: number } | { name: string };
export interface SnapshotCellLocator { sheet: SnapshotCellSheet; shapeId: number | null; section: string | null; sectionIndex?: number | null; row: SnapshotCellRow | null; cellName: string; }
export interface CellSnapshot { locator: SnapshotCellLocator; name: string; formula: string | null; value: string | null; rowType?: string; }
export interface ShapeSnapshot { id: string; sourceId: number; name: string | null; master?: number | null; cells: CellSnapshot[]; children: ShapeSnapshot[]; copySourceId?: number | null; copySourcePageId?: number | null; copyRefusal?: string | null; }
export interface PageLayer { index: number; name: string; visible: boolean; print: boolean; lock: boolean; active: boolean; color: string; status: string; }
export interface DocumentMaster { id: number; name: string | null; display: PageDisplayList | null; }
export interface ShapeTreeGlue { connectorSource: string; endpoint: string; targetSource: string; toCell: string; }
export interface FormulaShapeTreeGlue { connectorSource: string; endpoint: string; targetSource: string; toCell: string; }
export interface FormulaShapeTreeDraft { name?: string; cells: Array<{ locator: CellLocator & { rowType?: string }; name?: string; formula?: string; value?: string }>; text?: string; copySourceId?: number | null; copySourcePageId?: number | null; sourceShapeId?: string | null; sourceId?: number | null; copyRefusal?: string | null; glue?: FormulaShapeTreeGlue[]; children?: FormulaShapeTreeDraft[]; }
export type ValidationSeverity = 'error' | 'warning';
export interface RawValidationIssue { id: string; rule: string; severity: ValidationSeverity; pagePart: string; pageId: number | null; shapeId: number; otherShapeId: number | null; endpoint: string | null; row: string | null; }
export interface ValidationIssue { id: string; rule: string; severity: ValidationSeverity; pageId: string; shapeId: string; otherShapeId: string | null; endpoint: string | null; row: string | null; }
export interface PageSnapshot { id: string; sourcePartPath: string; name: string | null; shapes: ShapeSnapshot[]; }
export interface DiagramSnapshot { pages: PageSnapshot[]; }
export interface CellFormulaReceipt { pageId: string; shapeId: string; cellName: string; before: string | null; after: string; }
/** One `Property` row value to write; give a rowName or a rowIndex, not both. */
export interface ShapeDataWrite { rowName?: string; rowIndex?: number; sectionIndex?: number; formula: string; }
/** Per-row outcome of a shape-data batch; `refusal` is set when the row was not written. */
export interface ShapeDataReceipt { pageId: string; shapeId: string; rowName: string | null; rowIndex: number | null; sectionIndex: number | null; before: string | null; after: string | null; refusal: string | null; }
/** The user action a probe asks about; the cell name picks one when absent. */
export type MutationGesture = 'cellEdit' | 'moveX' | 'moveY' | 'resizeWidth' | 'resizeHeight' | 'resizeAspect' | 'rotate' | 'textEdit' | 'format' | 'delete';
/** One cell to probe, with the gesture to probe it as. */
export interface CellWriteQuery extends CellLocator { gesture?: MutationGesture; }
/** What the engine's mutation policy would do with a write to one cell, without writing it. */
export interface CellWriteProbe { cellName: string; allowed: boolean; targetCellName: string | null; refusal: 'guard' | 'lock' | 'unsupported' | null; reason: string | null; }
export interface ShapeReceipt { pageId: string; shapeId: string; fromIndex: number | null; toIndex: number | null; }
export interface TextReceipt { pageId: string; shapeId: string; before: string; after: string; }
export interface ConnectedShapeReceipt { shape: ShapeReceipt; connector: ShapeReceipt; }
export interface ConnectorRoutePoint { x: number; y: number; }
export interface ConnectorRouteReceipt { pageId: string; shapeId: string; points: number; }
export interface FormulaShapeDraft { name?: string; master?: number; cells: Array<{ locator: CellLocator & { rowType?: string }; name?: string; formula?: string; value?: string }> }
export interface ConnectorGlue { shapeId: string; toCell?: string; }
export interface VsdxFontFace { family: string; bold?: boolean; italic?: boolean; bytes: Uint8Array; }
export type Paint = { kind: 'solid'; color: string } | { kind: 'gradient'; angleDeg?: number; stops: Array<{ position: number; color: string }> };
export interface Stroke { color: string; width: number; dashed?: boolean; }
export interface GeometryPathCommand { type: string; [key: string]: number | string; }
export interface TextDiagnostic { category: 'integrity' | 'fidelity'; code: string; detail: string; }
export interface TextRun { text: string; family: string; sizeIn: number; bold: boolean; italic: boolean; underline: boolean; smallCaps: boolean; superscript: boolean; subscript: boolean; letterSpacing: number; color: string; diagnostics?: TextDiagnostic[]; }
export interface TextParagraph { runs: TextRun[]; }
export interface PositionedLine { x: number; y: number; width: number; height: number; start: number; end: number; caretStops: Array<{ position: number; x: number; y: number }>; }
interface PrimitiveBase { id: string; zOrder: number; }
export interface ShapeShadow { color: string; blurIn: number; offsetXIn: number; offsetYIn: number; }
export interface ShapePrimitive extends PrimitiveBase { kind: 'shape'; path: GeometryPathCommand[]; fill?: Paint; stroke?: Stroke; shadow?: ShapeShadow; transform?: Affine; diagnostics?: TextDiagnostic[]; }
export interface ImagePrimitive extends PrimitiveBase { kind: 'image'; assetId: string; x: number; y: number; width: number; height: number; transform?: Affine; }
export interface TextBoxPrimitive extends PrimitiveBase { kind: 'textBox'; x: number; y: number; width: number; height: number; paragraphs: TextParagraph[]; lines: PositionedLine[]; transform?: Affine; }
export interface PlaceholderPrimitive extends PrimitiveBase { kind: 'placeholder'; x: number; y: number; width: number; height: number; reason: string; }
export interface GroupPrimitive extends PrimitiveBase { kind: 'group'; primitives: PagePrimitive[]; transform?: Affine; }
export type PagePrimitive = ShapePrimitive | ImagePrimitive | TextBoxPrimitive | PlaceholderPrimitive | GroupPrimitive;
export interface PageDisplayList { contractVersion: 7; width: number; height: number; printWidth: number; printHeight: number; paintTransform: Affine; primitives: PagePrimitive[]; connectors?: ConnectorChrome[]; }
export type ConnectorEndpointGlue = 'free' | 'shape' | 'point';
export interface ConnectorChrome { id: string; begin: ConnectorEndpointGlue; end: ConnectorEndpointGlue; routable: boolean; }
export type HitTestResult = { kind: 'shape'; shapeId: string } | { kind: 'text'; shapeId: string; position: number };
export interface HistoryResult { applied: boolean; snapshot: DiagramSnapshot; }
export type CollaborationUpdateOrigin = 'local' | 'remote';
