//! Bounded PresentationML parsing and part-preserving package writes.

mod builtin_table_styles;
mod chart;
mod comment_patch;
mod comments;
mod custom_geometry;
mod drawing;
mod error;
mod inventory;
mod model;
mod notes;
mod package;
mod relationships;
mod table_style;
mod theme;
mod write;
mod xml;

pub use builtin_table_styles::builtin_table_style;
pub use comments::{
    Comment, CommentAuthor, CommentAuthorWrite, CommentFlavor, CommentSlide, CommentWrite,
    CommentsWrite,
};
pub use error::PptxError;
pub use inventory::{
    MediaKind, NotesSource, OmittedElement, OmittedInline, SlideSource, SourceShape,
};
pub use model::*;
pub use package::{
    effective_color_map, master_for_layout, parse_pptx, parse_pptx_with_limits,
    parse_pptx_without_connectors, slide_theme, write_pptx,
};
pub use relationships::{Relationship, TargetMode, relationship_types};
pub use write::{
    DeckWrite, InheritedTransform, NotesWrite, ParagraphWrite, PictureAdd, RunWrite, ShapeAdd,
    ShapePatch, ShapeWrite, SlideWrite, TextTarget, TextWrite, is_supported_image_content_type,
    write_pptx_with_edits,
};
pub use xml::ParseLimits;
