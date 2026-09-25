pub use pptx_edit::structured::{
    AnchorScope, CellPosition, ExportComment, ExportDiagnostic, ExportDiagnosticCode, ExportError,
    ExportFailure, ExportFailureCode, ExportLink, ExportList, ExportMark, ExportNotes,
    ExportObject, ExportObjectKind, ExportParagraph, ExportPlaceholder, ExportRead, ExportRefusal,
    ExportRun, ExportRunKind, ExportSeverity, ExportShape, ExportShapeKind, ExportSlide,
    ExportStory, ExportTable, ExportTableCell, ExportTableRow, IncludedContent, MarkdownAnchor,
    PptxAnchor, PptxExportOptions, PptxExportResult, PptxMarkdownContent, PptxMarkdownOptions,
    PptxStructuredContent, ReadingOrder, SourceProvenance, TextSpan, export_outcome_json,
    snapshot_outcome_json,
};
pub use pptx_edit::{
    CaretAnchor, CommentFlavor, CommentReceipt, CommentSnapshot, DeckSnapshot, EditCtx, EditError,
    EditOrigin, ParagraphSnapshot, PresetShapeDraft, ShapeAdjustReceipt, ShapeDraft,
    ShapeFillReceipt, ShapeKind, ShapeReceipt, ShapeRect, ShapeSnapshot, ShapeStroke,
    ShapeStrokeReceipt, SlideReceipt, SlideSnapshot, StorySnapshot, TextReceipt, TextRunSnapshot,
    TextSearchMatch, TextStyle, TextStylePatch, TransformReceipt, UpdateEvent, UpdateOrigin,
    UpdateSubscription,
};
pub use pptx_edit::{
    DocumentVersion, EditApplication, EditFailure, EditFailureCode, EditHistory, EditOutcome,
    EditPreview, EditReceipt, EditRefusal, EditRequest, EditSource, EditStep, EditTarget,
    EditValidation, FillGuard, FindMatch, FindOutcome, FindRequest, FindResponse, FindScope,
    MAX_REQUEST_BYTES, OutlineGuard, ParagraphText, ReadOutcome, ReadRequest, ReadResponse,
    RectGuard, ShapeTarget, SlideTarget, StoryTarget, StoryText, TargetEdge, TextField, TextGuard,
    TextRange, TextTarget, ValidationOutcome, outcome_json, oversized_request,
};
pub use pptx_edit::{
    Proposal, ProposalAcceptance, ProposalChange, ProposalEdit, ProposalError, ProposalPreview,
    ProposalRequest, ProposalResult,
};
pub use pptx_parse::{
    BlipEffect, Bullet, Comment, CommentAuthor, GraphicFrame, GraphicFrameData, GroupShape,
    LineSpacing, MediaPart, ParagraphProperties, ParseLimits, Picture, PictureCrop, Placeholder,
    PptxError, PptxPackage, Presentation as PresentationModel, Relationship, RunProperties, Shape,
    ShapeBase, ShapeNode, ShapeTransform, Slide, SlideLayout, SlideMaster, SlideReference,
    TargetMode, TextAutofit, TextBody, TextParagraph as ModelTextParagraph,
    TextRun as ModelTextRun, TextStyleSet, ThemePart,
};
pub use pptx_render::{
    CONTRACT_VERSION, CaretStop, GradientStop, GradientType, HitTestResult, ImageCrop, ImageEffect,
    Paint, PositionedGlyph, PositionedTextLine, PositionedTextRun, Primitive, RenderError,
    RenderedSlide, Stroke, SurfaceDisplayList, TextAlign, TextAnchor,
    TextParagraph as DisplayTextParagraph, TextRun as DisplayTextRun, Transform,
};
