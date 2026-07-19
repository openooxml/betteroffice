pub use pptx_edit::{
    DeckSnapshot, EditCtx, EditError, EditOrigin, ParagraphSnapshot, ShapeDraft, ShapeKind,
    ShapeReceipt, ShapeRect, ShapeSnapshot, SlideReceipt, SlideSnapshot, StorySnapshot,
    TextReceipt, TextRunSnapshot, TextStyle, TextStylePatch, TransformReceipt, UpdateEvent,
    UpdateOrigin,
};
pub use pptx_parse::{
    Bullet, GraphicFrame, GraphicFrameData, GroupShape, MediaPart, ParagraphProperties,
    ParseLimits, Picture, PictureCrop, Placeholder, PptxError, PptxPackage,
    Presentation as PresentationModel, Relationship, RunProperties, Shape, ShapeBase, ShapeNode,
    ShapeTransform, Slide, SlideLayout, SlideMaster, SlideReference, TargetMode, TextAutofit,
    TextBody, TextParagraph as ModelTextParagraph, TextRun as ModelTextRun, TextStyleSet,
    ThemePart,
};
pub use pptx_render::{
    CONTRACT_VERSION, CaretStop, GradientStop, GradientType, HitTestResult, Paint, PositionedGlyph,
    PositionedTextLine, PositionedTextRun, Primitive, RenderError, RenderedSlide, Stroke,
    SurfaceDisplayList, TextAlign, TextAnchor, TextParagraph as DisplayTextParagraph,
    TextRun as DisplayTextRun, Transform,
};
