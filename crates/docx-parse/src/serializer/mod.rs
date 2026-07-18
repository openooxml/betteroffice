//! Safe, deterministic XML serialization primitives and DOCX writer families.

pub mod context;
pub mod foundation;
pub mod numbering;
pub mod paragraph;
pub mod parts;
mod raw;
pub mod run;
pub mod s10;
pub mod s11;
pub mod s12;
pub mod s13;
pub mod sdt;
pub mod section;
pub mod table;
pub mod watermark;
pub mod xml_writer;

pub use context::SerializerContext;
pub use foundation::{
    BorderSide, serialize_border, serialize_conditional_format_style, serialize_table_grid,
};
pub use numbering::serialize_numbering_xml;
pub use paragraph::{
    serialize_inline_sdt, serialize_paragraph, serialize_paragraph_content,
    serialize_paragraph_formatting, synthesize_sdt_properties,
};
pub use parts::{
    CommentParaInfo, serialize_comments_extended_part, serialize_comments_extensible_part,
    serialize_comments_ids_part, serialize_comments_part, serialize_comments_with_info,
    serialize_document_body, serialize_document_part, serialize_endnotes_part,
    serialize_footnotes_part, serialize_header_footer_part,
};
pub use run::{
    serialize_drawing_content, serialize_run, serialize_shape_content, serialize_text_formatting,
};
pub use s10::{
    CanonicalXmlAttribute, CanonicalXmlEvent, S10SerializeRequest, S10SerializeResponse,
    SerializerDeterminism, canonical_xml_events, serialize_s10_wire,
};
pub use s11::{S11SerializeRequest, S11SerializeResponse, serialize_s11_wire};
pub use s12::{S12SerializeRequest, S12SerializeResponse, serialize_s12_wire};
pub use s13::{
    S13SaveOptions, S13SaveRequest, S13SelectiveSave, build_patched_document_xml,
    update_core_properties, write_docx_s13,
};
pub use sdt::{serialize_block_content, serialize_block_sdt};
pub use section::serialize_section_properties;
pub use table::{
    serialize_table, serialize_table_cell, serialize_table_cell_formatting,
    serialize_table_formatting, serialize_table_row, serialize_table_row_formatting,
};
pub use watermark::serialize_watermark;
pub use xml_writer::{XmlWriter, escape_xml, int_attr, js_number};
