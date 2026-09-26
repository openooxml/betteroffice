pub use pptx_edit::{
    CaretAnchor, CommentFlavor, CommentReceipt, CommentSnapshot, DeckSnapshot, EditCtx, EditError,
    EditOrigin, InheritedGeometry, ParagraphSnapshot, PresetShapeDraft, ShapeAdjustReceipt,
    ShapeDraft, ShapeFillReceipt, ShapeKind, ShapeReceipt, ShapeRect, ShapeSnapshot, ShapeStroke,
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
    CONTRACT_VERSION, CaretStop, FontSubstitution, GradientStop, GradientType, HitTestResult,
    ImageCrop, ImageEffect, Paint, PositionedGlyph, PositionedTextLine, PositionedTextRun,
    Primitive, RenderError, RenderedSlide, Stroke, SurfaceDisplayList, TextAlign, TextAnchor,
    TextParagraph as DisplayTextParagraph, TextRun as DisplayTextRun, Transform,
};
