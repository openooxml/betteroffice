//! Block-story dispatch and validated block SDT replay.

use crate::block::{BlockContent, BlockSdt};
use crate::xml::ParseError;

use super::context::SerializerContext;
use super::paragraph::{serialize_paragraph, synthesize_sdt_properties};
use super::raw::validate_raw_subtree;
use super::run::append_generated;
use super::table::serialize_table;
use super::xml_writer::XmlWriter;

pub fn serialize_block_content(
    block: &BlockContent,
    context: &mut SerializerContext,
) -> Result<String, ParseError> {
    match block {
        BlockContent::Paragraph(paragraph) => serialize_paragraph(paragraph, context),
        BlockContent::Table(table) => serialize_table(table, context),
        BlockContent::BlockSdt(sdt) => serialize_block_sdt(sdt, context),
    }
}

pub fn serialize_block_sdt(
    sdt: &BlockSdt,
    context: &mut SerializerContext,
) -> Result<String, ParseError> {
    let properties = if let Some(raw) = sdt.properties.raw_properties_xml.as_deref() {
        validate_raw_subtree(raw, "w", "sdtPr")?;
        raw.to_owned()
    } else {
        // Pinned incumbent behavior: parsed block controls without captured
        // properties emit no `w:sdtPr`. Programmatic controls normally carry
        // synthesized raw XML already, but accept the explicit modeled form.
        let synthesized = synthesize_sdt_properties(&sdt.properties);
        if sdt.properties.sdt_type == "richText"
            && sdt.properties.id.is_none()
            && sdt.properties.alias.is_none()
            && sdt.properties.tag.is_none()
            && sdt.properties.lock.is_none()
            && sdt.properties.placeholder.is_none()
            && sdt.properties.showing_placeholder.is_none()
        {
            String::new()
        } else {
            synthesized
        }
    };
    let end_properties = if let Some(raw) = sdt.properties.raw_end_properties_xml.as_deref() {
        validate_raw_subtree(raw, "w", "sdtEndPr")?;
        raw.to_owned()
    } else {
        String::new()
    };
    let mut writer = XmlWriter::with_capacity(256);
    writer.start_element("w:sdt");
    append_generated(&mut writer, &properties);
    append_generated(&mut writer, &end_properties);
    writer.start_element("w:sdtContent");
    for block in &sdt.content {
        append_generated(&mut writer, &serialize_block_content(block, context)?);
    }
    writer.end_element().end_element();
    Ok(writer.finish())
}
