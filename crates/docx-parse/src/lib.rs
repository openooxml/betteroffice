// Ported from openooxml/docx, which did not gate on clippy style lints;
// burning these down is tracked follow-up work, not a merge blocker.
#![allow(
    clippy::cloned_ref_to_slice_refs,
    clippy::collapsible_if,
    clippy::collapsible_match,
    clippy::doc_lazy_continuation,
    clippy::excessive_precision,
    clippy::field_reassign_with_default,
    clippy::if_same_then_else,
    clippy::inconsistent_digit_grouping,
    clippy::items_after_test_module,
    clippy::large_enum_variant,
    clippy::manual_contains,
    clippy::manual_is_multiple_of,
    clippy::manual_pattern_char_comparison,
    clippy::manual_repeat_n,
    clippy::manual_unwrap_or,
    clippy::map_clone,
    clippy::needless_lifetimes,
    clippy::obfuscated_if_else,
    clippy::too_many_arguments,
    clippy::trim_split_whitespace,
    clippy::type_complexity,
    clippy::unnecessary_filter_map,
    clippy::unnecessary_lazy_evaluations,
    clippy::unnecessary_sort_by
)]

//! Safe DOCX parsing foundation.
//!
//! S0 freezes the cross-language canonical contract in [`canonical`]. S1 adds
//! bounded XML and relationship parsing on top of `ooxml-opc`'s existing
//! ZIP/OPC trust boundary.

pub mod block;
pub mod borders;
pub mod canonical;
pub mod chart;
pub mod comments;
pub mod document;
pub mod drawingml;
pub mod fonts;
pub mod formatting;
pub mod header_footer;
pub mod image;
pub mod inline;
pub mod media;
pub mod notes;
pub mod numbering;
pub mod paragraph;
pub mod relationships;
pub mod s2;
pub mod s3;
pub mod s4;
pub mod s5;
pub mod s6;
pub mod s7;
pub mod s8;
pub mod s9;
pub mod scalars;
pub mod section;
pub mod serializer;
pub mod settings;
pub mod shape;
pub mod smart_art;
pub mod styles;
pub mod table;
pub mod tabs;
pub mod text_box;
pub mod theme;
pub mod vml;
pub mod wrap;
pub mod xml;

use wasm_bindgen::prelude::*;

pub use block::{BlockContent, BlockSdt, StoryParser};
pub use borders::{
    BorderSpec, Borders, parse_border_spec, parse_paragraph_borders, parse_table_borders,
};
pub use comments::{Comment, parse_comments, remove_orphan_comment_ranges};
pub use document::{
    DocumentBody, Section, extract_all_template_variables, extract_template_variables,
    get_paragraph_text, is_empty_paragraph, parse_document_body,
};
pub use fonts::{FontEmbed, FontInfo, FontTable, parse_font_table};
pub use formatting::{
    CellMargins, ConditionalFormatStyle, FloatingTableProperties, FontFamily, NumberingProperties,
    ParagraphFormatting, ParagraphFrame, RunLanguage, SpacingExplicit, TableCellFormatting,
    TableFormatting, TableLook, TableMeasurement, TableRowFormatting, TextFill, TextFormatting,
    TextGlow, TextGradientStop, TextModernEffects, TextOutline, TextReflection, TextShadow,
    merge_paragraph_formatting, merge_text_formatting, parse_paragraph_properties,
    parse_run_properties, parse_table_cell_properties, parse_table_properties,
    parse_table_row_properties,
};
pub use header_footer::{
    HeaderFooter, normalize_header_footer_type, parse_header_footer, parse_related_header_footers,
    select_for_page,
};
pub use inline::{
    BookmarkEnd, BookmarkStart, ComplexField, FieldFormData, FieldSwitch, Hyperlink, InlineNode,
    InlineSdt, MathEquation, ParsedFieldInstruction, Run, RunContent, RunProjection, SdtProperties,
    SimpleField, StructuredFieldContent, StructuredFieldTree, parse_bookmark_end,
    parse_bookmark_start, parse_field_instruction, parse_field_type, parse_hyperlink,
    parse_inline_container, parse_run, parse_sdt_properties, parse_simple_field, sanitize_href,
};
pub use notes::{
    Note, NoteProperties, NoteSeparatorReference, parse_endnote_properties,
    parse_footnote_properties, parse_note_type, parse_notes,
};
pub use numbering::{
    AbstractNumbering, LegacyLevel, LevelOverride, ListLevel, ListRendering, NumberingDefinitions,
    NumberingInstance, NumberingMap, compute_list_rendering, format_number, get_bullet_character,
    is_bullet_level, pad_decimal, parse_numbering, render_list_marker,
};
pub use paragraph::{
    CommentRange, DrawingContext, HexIdAllocator, Paragraph, ParagraphContent,
    ParagraphPropertyChange, RangeEnd, RangeStart, TrackedChangeInfo, TrackedInline,
    convert_bullet_to_unicode, paragraph_starts_with_rendered_page_break,
    parse_document_paragraph_properties, parse_paragraph,
};
pub use relationships::{
    Relationship, RelationshipMap, RelationshipPart, RelationshipTarget, TargetMode, WireEnvelope,
    filter_by_type, footers, get_relationship_type_name, headers, hyperlinks, images,
    is_external_hyperlink, is_footer_relationship, is_header_relationship, is_image_relationship,
    parse_docx_relationship_parts, parse_relationships, relationship_types, resolve_relationship,
    resolve_relationship_target, resolve_relative_path, resolve_target,
};
pub use s2::{
    ElementLeaf, S2Projection, S2WireEnvelope, S2XmlPart, parse_docx_s2_projection,
    parse_docx_s2_wire, project_xml_part,
};
pub use s3::{
    NumberingResolution, ResolvedNumberingLevel, S3Projection, S3WireEnvelope, StyleDefaults,
    parse_docx_s3_projection, parse_docx_s3_wire,
};
pub use s4::{
    DrawingLeaf, S4Projection, S4WireEnvelope, S4XmlPart, parse_docx_s4_projection,
    parse_docx_s4_wire, project_s4_xml_part,
};
pub use s5::{
    S5Projection, S5WireEnvelope, S5XmlPart, parse_docx_s5_projection, parse_docx_s5_wire,
    project_xml_part as project_s5_xml_part,
};
pub use s6::{
    S6Projection, S6WireEnvelope, parse_docx_s6_projection, parse_docx_s6_wire, s6_wire_envelope,
};
pub use s7::{
    S7Projection, S7WireEnvelope, parse_docx_s7_projection, parse_docx_s7_wire, s7_wire_envelope,
};
pub use s8::{
    S8Projection, S8WireEnvelope, parse_docx_s8_projection, parse_docx_s8_wire, s8_wire_envelope,
};
pub use s9::{
    BinaryPartWire, S9DocumentBodyWire, S9DocumentWire, S9PackageWire, S9ParseOptions,
    S9SectionWire, S9WireEnvelope, parse_docx_s9_wire,
};
pub use scalars::{
    ColorValue, RunScalarProperties, ShadingProperties, UnderlineValue, parse_color_value,
    parse_run_scalar_properties, parse_shading_properties, parse_underline,
};
pub use section::{
    Column, DocumentGrid, LineNumbering, PageBorders, PageNumberingProperties, SectionBackground,
    SectionProperties, StoryReference, apply_section_inheritance, default_section_properties,
    parse_section_properties,
};
pub use serializer::{
    CanonicalXmlAttribute, CanonicalXmlEvent, S10SerializeRequest, S10SerializeResponse,
    S11SerializeRequest, S11SerializeResponse, S12SerializeRequest, S12SerializeResponse,
    S13SaveOptions, S13SaveRequest, S13SelectiveSave, SerializerDeterminism, canonical_xml_events,
    serialize_s10_wire, serialize_s11_wire, serialize_s12_wire, write_docx_s13,
};
pub use settings::{
    CompatibilityFlags, DocumentSettings, RevisionView, ThemeFontLanguage, parse_settings,
};
pub use styles::{
    DocDefaults, LatentStyles, Style, StyleDefinitions, StyleMap, TableStyleConditional,
    get_default_character_style, get_default_paragraph_style, get_default_table_style,
    get_styles_by_type, parse_style_definitions, parse_styles,
};
pub use table::{
    Table, TableCell, TableCellPropertyChange, TablePropertyChange, TableRow,
    TableRowPropertyChange, TableStructuralChangeInfo, get_header_rows, get_table_column_count,
    get_table_row_count, get_table_text, has_header_row, infer_implicit_single_cell_row_spans,
    is_cell_horizontally_merged, is_cell_merge_continuation, is_cell_merge_start,
    is_floating_table, parse_conditional_format_style, parse_document_shading,
    parse_document_table_borders, parse_document_table_cell_properties, parse_document_table_look,
    parse_document_table_properties, parse_document_table_row_properties,
    parse_floating_table_properties, parse_table_cell_property_changes,
    parse_table_cell_structural_change, parse_table_grid, parse_table_measurement,
    parse_table_property_changes, parse_table_row_property_changes,
    parse_table_row_structural_change,
};
pub use tabs::{TabError, TabStop, parse_tab_stop, parse_tab_stops};
pub use theme::{
    Theme, ThemeColorScheme, ThemeFont, ThemeFontScheme, apply_theme_font_lang, get_default_theme,
    get_major_font, get_minor_font, get_theme_color, get_theme_fonts, parse_theme,
    resolve_theme_font_ref,
};
pub use xml::{
    MAX_NESTING_DEPTH, ParseBudget, ParseError, ParseLimits, ParsedColor, XmlDocument, XmlElement,
    XmlNode, namespaces, parse_xml,
};

/// Wasm control-plane entry: safe ZIP -> bounded XML -> typed relationships.
#[wasm_bindgen]
pub fn parse_docx_relationships(data: &[u8]) -> Result<String, JsValue> {
    let envelope = parse_docx_relationship_parts(data).map_err(js_error)?;
    serde_json::to_string(&envelope).map_err(|error| js_error(error.to_string()))
}

/// Focused wasm leaf used by hostile-input and facade tests.
#[wasm_bindgen]
pub fn parse_relationships_xml(xml: &[u8], part_path: &str) -> Result<String, JsValue> {
    let limits = ParseLimits::default();
    let mut budget = ParseBudget::new(&limits);
    let relationships = parse_relationships(xml, part_path, &mut budget).map_err(js_error)?;
    let part = RelationshipPart::from_map(part_path.to_owned(), relationships).map_err(js_error)?;
    serde_json::to_string(&part).map_err(|error| js_error(error.to_string()))
}

/// Legacy staged Rust S2 entry retained for ABI compatibility.
#[wasm_bindgen]
pub fn parse_docx_s2(data: &[u8]) -> Result<String, JsValue> {
    let envelope = parse_docx_s2_wire(data).map_err(js_error)?;
    serde_json::to_string(&envelope).map_err(|error| js_error(error.to_string()))
}

/// Legacy staged Rust S3 entry retained for ABI compatibility.
#[wasm_bindgen]
pub fn parse_docx_s3(data: &[u8]) -> Result<String, JsValue> {
    let envelope = parse_docx_s3_wire(data).map_err(js_error)?;
    serde_json::to_string(&envelope).map_err(|error| js_error(error.to_string()))
}

/// Legacy staged Rust S4 entry retained for ABI compatibility.
#[wasm_bindgen]
pub fn parse_docx_s4(data: &[u8]) -> Result<String, JsValue> {
    let envelope = parse_docx_s4_wire(data).map_err(js_error)?;
    serde_json::to_string(&envelope).map_err(|error| js_error(error.to_string()))
}

/// Legacy staged Rust S5 entry retained for ABI compatibility.
#[wasm_bindgen]
pub fn parse_docx_s5(data: &[u8]) -> Result<String, JsValue> {
    let envelope = parse_docx_s5_wire(data).map_err(js_error)?;
    serde_json::to_string(&envelope).map_err(|error| js_error(error.to_string()))
}

/// Legacy staged Rust S6 entry retained for ABI compatibility.
#[wasm_bindgen]
pub fn parse_docx_s6(data: &[u8]) -> Result<String, JsValue> {
    let envelope = parse_docx_s6_wire(data).map_err(js_error)?;
    serde_json::to_string(&envelope).map_err(|error| js_error(error.to_string()))
}

/// Legacy staged Rust S7 entry retained for ABI compatibility.
#[wasm_bindgen]
pub fn parse_docx_s7(data: &[u8]) -> Result<String, JsValue> {
    let envelope = parse_docx_s7_wire(data).map_err(js_error)?;
    serde_json::to_string(&envelope).map_err(|error| js_error(error.to_string()))
}

/// Legacy staged Rust S8 entry retained for ABI compatibility.
#[wasm_bindgen]
pub fn parse_docx_s8(data: &[u8]) -> Result<String, JsValue> {
    let envelope = parse_docx_s8_wire(data).map_err(js_error)?;
    serde_json::to_string(&envelope).map_err(|error| js_error(error.to_string()))
}

/// S9 production read facade: one safe package pass to the full Document wire.
#[wasm_bindgen]
pub fn parse_docx_s9(data: &[u8], options_json: &str) -> Result<String, JsValue> {
    let options = if options_json.is_empty() {
        S9ParseOptions::default()
    } else {
        serde_json::from_str(options_json).map_err(|error| js_error(error.to_string()))?
    };
    let envelope = parse_docx_s9_wire(data, options).map_err(js_error)?;
    serde_json::to_string(&envelope).map_err(|error| js_error(error.to_string()))
}

/// Legacy staged Rust S10 serializer entry retained for ABI compatibility.
#[wasm_bindgen]
pub fn serialize_docx_s10(request_json: &str) -> Result<String, JsValue> {
    let request =
        serde_json::from_str(request_json).map_err(|error| js_error(error.to_string()))?;
    let response = serialize_s10_wire(request).map_err(js_error)?;
    serde_json::to_string(&response).map_err(|error| js_error(error.to_string()))
}

/// Legacy staged Rust S11 serializer entry retained for ABI compatibility.
#[wasm_bindgen]
pub fn serialize_docx_s11(request_json: &str) -> Result<String, JsValue> {
    let request =
        serde_json::from_str(request_json).map_err(|error| js_error(error.to_string()))?;
    let response = serialize_s11_wire(request).map_err(js_error)?;
    serde_json::to_string(&response).map_err(|error| js_error(error.to_string()))
}

/// Legacy staged Rust S12 serializer entry retained for ABI compatibility.
#[wasm_bindgen]
pub fn serialize_docx_s12(request_json: &str) -> Result<String, JsValue> {
    let request =
        serde_json::from_str(request_json).map_err(|error| js_error(error.to_string()))?;
    let response = serialize_s12_wire(request).map_err(js_error)?;
    serde_json::to_string(&response).map_err(|error| js_error(error.to_string()))
}

/// S13 production-capable package writer: typed model + original package -> DOCX.
#[wasm_bindgen]
pub fn write_docx_s13_wasm(request_json: &str, original_docx: &[u8]) -> Result<Vec<u8>, JsValue> {
    let request =
        serde_json::from_str(request_json).map_err(|error| js_error(error.to_string()))?;
    write_docx_s13(request, original_docx).map_err(js_error)
}

fn js_error(error: impl ToString) -> JsValue {
    JsValue::from_str(&error.to_string())
}
