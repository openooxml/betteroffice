//! Typed facade for opening, editing, laying out, and saving DOCX files.

mod document;
mod error;
#[cfg(feature = "raster")]
mod render;
mod types;

pub use document::Document;
#[cfg(feature = "raster")]
pub use docx_raster::{ImageScope, NoteKind, RenderedPage};
pub use error::Error;
#[cfg(feature = "raster")]
pub use render::{MAX_PIXMAP_DIM, MAX_PIXMAP_PIXELS};
pub use types::{DocumentModel, DocumentStructure, LayoutResult, SaveOptions};

pub use docx_edit::{
    EditCtx, EditError, EditOrigin, EditingDoc, FormatPolicy, Loc, LocRange, OpError, Receipt,
    StoryRange, TextView,
};
pub use docx_layout::display_list::{DisplayList, DisplayPage, Primitive};
pub use docx_layout::types::{Input as LayoutInput, Layout, LayoutOptions, MeasuredBlock, Page};
pub use docx_parse::xml::ParseLimits;
pub use docx_parse::{
    BlockContent, DocumentBody, HeaderFooter, InlineNode, Paragraph, ParagraphContent, Run,
    RunContent, Section, SectionProperties, Table, TableCell, TableRow, get_paragraph_text,
};

pub type Result<T> = std::result::Result<T, Error>;
