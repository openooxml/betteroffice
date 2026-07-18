//! Paragraph properties and inline-content serialization.

use crate::borders::Borders;
use crate::formatting::{NumberingProperties, ParagraphFormatting, ParagraphFrame};
use crate::inline::{
    BookmarkEnd, BookmarkStart, ComplexField, Hyperlink, InlineNode, InlineSdt, MathEquation,
    SdtProperties, SimpleField,
};
use crate::paragraph::{
    CommentRange, Paragraph, ParagraphContent, ParagraphPropertyChange, RangeEnd, RangeStart,
    TrackedChangeInfo, TrackedInline,
};
use crate::section::SectionProperties;
use crate::xml::ParseError;

use super::context::SerializerContext;
use super::foundation::{BorderSide, write_border};
use super::raw::{validate_math_subtree, validate_raw_subtree};
use super::run::{
    append_generated, nonempty, nonempty_trimmed, normalized_tracked_id, serialize_deleted_run,
    serialize_run, serialize_text_formatting, write_shading,
};
use super::section::serialize_section_properties;
use super::xml_writer::{XmlWriter, int_attr, js_number};

/// Serialize one complete `w:p` element.
pub fn serialize_paragraph(
    paragraph: &Paragraph,
    context: &mut SerializerContext,
) -> Result<String, ParseError> {
    context.enter_paragraph(paragraph.rendered_page_break_before == Some(true));
    let result = serialize_paragraph_inner(paragraph, context);
    context.leave_paragraph();
    result
}

fn serialize_paragraph_inner(
    paragraph: &Paragraph,
    context: &mut SerializerContext,
) -> Result<String, ParseError> {
    let mut writer = XmlWriter::with_capacity(512);
    writer.start_element("w:p");
    if let Some(value) = nonempty(paragraph.para_id.as_deref()) {
        writer.attribute("w14:paraId", value);
    }
    if let Some(value) = nonempty(paragraph.text_id.as_deref()) {
        writer.attribute("w14:textId", value);
    }
    let properties = serialize_paragraph_formatting(
        paragraph.formatting.as_ref(),
        paragraph.property_changes.as_deref(),
        paragraph.p_pr_ins.as_ref(),
        paragraph.p_pr_del.as_ref(),
        false,
        paragraph.section_properties.as_ref(),
    )?;
    append_generated(&mut writer, &properties);
    for content in &paragraph.content {
        let content = serialize_paragraph_content(content, context)?;
        append_generated(&mut writer, &content);
    }
    writer.end_element();
    Ok(writer.finish())
}

#[allow(clippy::too_many_arguments)]
pub fn serialize_paragraph_formatting(
    formatting: Option<&ParagraphFormatting>,
    property_changes: Option<&[ParagraphPropertyChange]>,
    p_pr_ins: Option<&TrackedChangeInfo>,
    p_pr_del: Option<&TrackedChangeInfo>,
    base_only: bool,
    section_properties: Option<&SectionProperties>,
) -> Result<String, ParseError> {
    let mut body = XmlWriter::with_capacity(512);
    if let Some(formatting) = formatting {
        if let Some(value) = nonempty(formatting.style_id.as_deref()) {
            empty_attr(&mut body, "w:pStyle", "w:val", value);
        }
        on_off(&mut body, "w:keepNext", formatting.keep_next);
        on_off(&mut body, "w:keepLines", formatting.keep_lines);
        on_off(
            &mut body,
            "w:contextualSpacing",
            formatting.contextual_spacing,
        );
        on_off(&mut body, "w:pageBreakBefore", formatting.page_break_before);
        write_frame(&mut body, formatting.frame.as_ref());
        on_off(&mut body, "w:widowControl", formatting.widow_control);

        if formatting.num_pr != formatting.num_pr_from_style
            || formatting.num_pr_from_style.is_none()
        {
            write_numbering(&mut body, formatting.num_pr.as_ref());
        }
        write_paragraph_borders(&mut body, formatting.borders.as_ref());
        write_shading(&mut body, formatting.shading.as_ref());
        write_tabs(&mut body, formatting.tabs.as_deref());
        on_off(
            &mut body,
            "w:suppressLineNumbers",
            formatting.suppress_line_numbers,
        );
        on_off(
            &mut body,
            "w:suppressAutoHyphens",
            formatting.suppress_auto_hyphens,
        );
        write_spacing(&mut body, formatting);
        write_indentation(&mut body, formatting);
        on_off(&mut body, "w:bidi", formatting.bidi);
        if let Some(value) = nonempty(formatting.alignment.as_deref()) {
            empty_attr(&mut body, "w:jc", "w:val", value);
        }
        if let Some(value) = formatting.outline_level {
            empty_attr(&mut body, "w:outlineLvl", "w:val", &js_number(value));
        }
    }

    if !base_only {
        write_paragraph_mark_properties(&mut body, formatting, p_pr_ins, p_pr_del);
        if let Some(properties) = section_properties {
            append_generated(&mut body, &serialize_section_properties(Some(properties)));
        }
        if let Some(change) = property_changes.and_then(|changes| changes.first()) {
            append_generated(&mut body, &serialize_paragraph_property_change(change)?);
        }
    }
    let body = body.finish();
    if body.is_empty() {
        Ok(String::new())
    } else {
        Ok(format!("<w:pPr>{body}</w:pPr>"))
    }
}

fn serialize_paragraph_property_change(
    change: &ParagraphPropertyChange,
) -> Result<String, ParseError> {
    let previous = serialize_paragraph_formatting(
        change.previous_formatting.as_ref(),
        None,
        None,
        None,
        true,
        None,
    )?;
    let previous = if previous.is_empty() {
        "<w:pPr/>".to_owned()
    } else {
        previous
    };
    let mut writer = XmlWriter::with_capacity(previous.len() + 128);
    writer
        .start_element("w:pPrChange")
        .attribute("w:id", &normalized_tracked_id(change.info.id))
        .attribute(
            "w:author",
            nonempty_trimmed(&change.info.author).unwrap_or("Unknown"),
        );
    if let Some(date) = change.info.date.as_deref().and_then(nonempty_trimmed) {
        writer.attribute("w:date", date);
    }
    append_generated(&mut writer, &previous);
    writer.end_element();
    Ok(writer.finish())
}

pub fn serialize_paragraph_content(
    content: &ParagraphContent,
    context: &mut SerializerContext,
) -> Result<String, ParseError> {
    match content {
        ParagraphContent::Inline(node) => serialize_inline_node(node, context),
        ParagraphContent::Tracked(change) => serialize_tracked_change(change, context),
        ParagraphContent::RangeStart(marker) => serialize_range_start(marker),
        ParagraphContent::RangeEnd(marker) => serialize_range_end(marker),
        ParagraphContent::CommentRange(marker) => serialize_comment_range(marker),
    }
}

pub(crate) fn serialize_inline_node(
    node: &InlineNode,
    context: &mut SerializerContext,
) -> Result<String, ParseError> {
    match node {
        InlineNode::Run(run) => serialize_run(run, context),
        InlineNode::Hyperlink(hyperlink) => serialize_hyperlink(hyperlink, context),
        InlineNode::BookmarkStart(bookmark) => Ok(serialize_bookmark_start(bookmark)),
        InlineNode::BookmarkEnd(bookmark) => Ok(serialize_bookmark_end(bookmark)),
        InlineNode::SimpleField(field) => serialize_simple_field(field, context),
        InlineNode::ComplexField(field) => serialize_complex_field(field, context),
        InlineNode::InlineSdt(sdt) => serialize_inline_sdt(sdt, context),
        InlineNode::Math(math) => serialize_math(math),
    }
}

fn serialize_bookmark_start(bookmark: &BookmarkStart) -> String {
    let mut writer = XmlWriter::with_capacity(96);
    writer
        .start_element("w:bookmarkStart")
        .attribute("w:id", &js_number(bookmark.id))
        .attribute("w:name", &bookmark.name);
    if let Some(value) = bookmark.col_first {
        writer.attribute("w:colFirst", &js_number(value));
    }
    if let Some(value) = bookmark.col_last {
        writer.attribute("w:colLast", &js_number(value));
    }
    writer.end_element();
    writer.finish()
}

fn serialize_bookmark_end(bookmark: &BookmarkEnd) -> String {
    let mut writer = XmlWriter::with_capacity(40);
    writer
        .start_element("w:bookmarkEnd")
        .attribute("w:id", &js_number(bookmark.id))
        .end_element();
    writer.finish()
}

fn serialize_hyperlink(
    hyperlink: &Hyperlink,
    context: &mut SerializerContext,
) -> Result<String, ParseError> {
    let mut children = String::new();
    for child in &hyperlink.children {
        match child {
            InlineNode::Run(run) => children.push_str(&serialize_run(run, context)?),
            InlineNode::BookmarkStart(bookmark) => {
                children.push_str(&serialize_bookmark_start(bookmark))
            }
            InlineNode::BookmarkEnd(bookmark) => {
                children.push_str(&serialize_bookmark_end(bookmark))
            }
            _ => {}
        }
    }
    let has_attributes = nonempty(hyperlink.relationship_id.as_deref()).is_some()
        || nonempty(hyperlink.anchor.as_deref()).is_some()
        || nonempty(hyperlink.tooltip.as_deref()).is_some()
        || nonempty(hyperlink.target.as_deref()).is_some()
        || hyperlink.history == Some(false)
        || nonempty(hyperlink.doc_location.as_deref()).is_some();
    if !has_attributes && nonempty(hyperlink.href.as_deref()).is_none() {
        return Ok(children);
    }
    let mut writer = XmlWriter::with_capacity(children.len() + 128);
    writer.start_element("w:hyperlink");
    optional_attr(&mut writer, "r:id", hyperlink.relationship_id.as_deref());
    optional_attr(&mut writer, "w:anchor", hyperlink.anchor.as_deref());
    optional_attr(&mut writer, "w:tooltip", hyperlink.tooltip.as_deref());
    optional_attr(&mut writer, "w:tgtFrame", hyperlink.target.as_deref());
    if hyperlink.history == Some(false) {
        writer.attribute("w:history", "0");
    }
    optional_attr(
        &mut writer,
        "w:docLocation",
        hyperlink.doc_location.as_deref(),
    );
    append_generated(&mut writer, &children);
    writer.end_element();
    Ok(writer.finish())
}

fn serialize_simple_field(
    field: &SimpleField,
    context: &mut SerializerContext,
) -> Result<String, ParseError> {
    let properties = field
        .content
        .first()
        .and_then(|run| run.formatting.as_ref())
        .map(|formatting| serialize_text_formatting(Some(formatting)))
        .unwrap_or_default();
    let mut output = String::new();
    output.push_str("<w:r>");
    output.push_str(&properties);
    output.push_str("<w:fldChar w:fldCharType=\"begin\"");
    if field.fld_lock == Some(true) {
        output.push_str(" w:fldLock=\"true\"");
    }
    output.push_str("/></w:r>");
    output.push_str("<w:r>");
    output.push_str(&properties);
    output.push_str("<w:instrText");
    if needs_preserve(&field.instruction) {
        output.push_str(" xml:space=\"preserve\"");
    }
    output.push('>');
    output.push_str(&super::xml_writer::escape_xml(&field.instruction));
    output.push_str("</w:instrText></w:r>");
    output.push_str("<w:r>");
    output.push_str(&properties);
    output.push_str("<w:fldChar w:fldCharType=\"separate\"/></w:r>");
    for run in &field.content {
        output.push_str(&serialize_run(run, context)?);
    }
    output.push_str("<w:r>");
    output.push_str(&properties);
    output.push_str("<w:fldChar w:fldCharType=\"end\"/></w:r>");
    Ok(output)
}

fn serialize_complex_field(
    field: &ComplexField,
    context: &mut SerializerContext,
) -> Result<String, ParseError> {
    let formatting = field
        .field_result
        .first()
        .and_then(|run| run.formatting.as_ref())
        .or(field.formatting.as_ref());
    let properties = serialize_text_formatting(formatting);
    let mut output = String::new();
    output.push_str("<w:r>");
    output.push_str(&properties);
    output.push_str("<w:fldChar w:fldCharType=\"begin\"");
    if field.fld_lock == Some(true) {
        output.push_str(" w:fldLock=\"true\"");
    }
    output.push_str("/></w:r>");
    if field.field_code.is_empty() {
        output.push_str("<w:r>");
        output.push_str(&properties);
        output.push_str("<w:instrText");
        if needs_preserve(&field.instruction) {
            output.push_str(" xml:space=\"preserve\"");
        }
        output.push('>');
        output.push_str(&super::xml_writer::escape_xml(&field.instruction));
        output.push_str("</w:instrText></w:r>");
    } else {
        for run in &field.field_code {
            output.push_str(&serialize_run(run, context)?);
        }
    }
    output.push_str("<w:r>");
    output.push_str(&properties);
    output.push_str("<w:fldChar w:fldCharType=\"separate\"/></w:r>");
    for run in &field.field_result {
        output.push_str(&serialize_run(run, context)?);
    }
    output.push_str("<w:r>");
    output.push_str(&properties);
    output.push_str("<w:fldChar w:fldCharType=\"end\"/></w:r>");
    Ok(output)
}

pub fn serialize_inline_sdt(
    sdt: &InlineSdt,
    context: &mut SerializerContext,
) -> Result<String, ParseError> {
    let (properties, end_properties) = serialize_sdt_properties(&sdt.properties)?;
    let mut writer = XmlWriter::with_capacity(256);
    writer.start_element("w:sdt");
    append_generated(&mut writer, &properties);
    append_generated(&mut writer, &end_properties);
    writer.start_element("w:sdtContent");
    for item in &sdt.content {
        append_generated(&mut writer, &serialize_inline_node(item, context)?);
    }
    writer.end_element().end_element();
    Ok(writer.finish())
}

pub(crate) fn serialize_sdt_properties(
    properties: &SdtProperties,
) -> Result<(String, String), ParseError> {
    let properties_xml = if let Some(raw) = properties.raw_properties_xml.as_deref() {
        validate_raw_subtree(raw, "w", "sdtPr")?;
        raw.to_owned()
    } else {
        synthesize_sdt_properties(properties)
    };
    let end_properties = if let Some(raw) = properties.raw_end_properties_xml.as_deref() {
        validate_raw_subtree(raw, "w", "sdtEndPr")?;
        raw.to_owned()
    } else {
        String::new()
    };
    Ok((properties_xml, end_properties))
}

pub fn synthesize_sdt_properties(properties: &SdtProperties) -> String {
    let mut body = XmlWriter::with_capacity(256);
    if let Some(value) = nonempty(properties.alias.as_deref()) {
        empty_attr(&mut body, "w:alias", "w:val", value);
    }
    if let Some(value) = nonempty(properties.tag.as_deref()) {
        empty_attr(&mut body, "w:tag", "w:val", value);
    }
    if let Some(value) = properties.id {
        empty_attr(&mut body, "w:id", "w:val", &js_number(value));
    }
    if let Some(value) = nonempty(properties.lock.as_deref()).filter(|value| *value != "unlocked") {
        empty_attr(&mut body, "w:lock", "w:val", value);
    }
    if let Some(value) = nonempty(properties.placeholder.as_deref()) {
        body.start_element("w:placeholder");
        empty_attr(&mut body, "w:docPart", "w:val", value);
        body.end_element();
    }
    if properties.showing_placeholder == Some(true) {
        body.start_element("w:showingPlcHdr").end_element();
    }
    match properties.sdt_type.as_str() {
        "plainText" => {
            body.start_element("w:text").end_element();
        }
        "date" => {
            body.start_element("w:date");
            if let Some(value) = nonempty(properties.date_format.as_deref()) {
                empty_attr(&mut body, "w:dateFormat", "w:val", value);
            }
            body.end_element();
        }
        "dropDownList" | "comboBox" => {
            let name = if properties.sdt_type == "dropDownList" {
                "w:dropDownList"
            } else {
                "w:comboBox"
            };
            body.start_element(name).attribute("w:lastValue", "");
            for item in properties.list_items.as_deref().unwrap_or_default() {
                body.start_element("w:listItem")
                    .attribute("w:displayText", &item.display_text)
                    .attribute("w:value", &item.value)
                    .end_element();
            }
            body.end_element();
        }
        "checkbox" => {
            body.start_element("w14:checkbox");
            body.start_element("w14:checked")
                .attribute(
                    "w14:val",
                    if properties.checked == Some(true) {
                        "1"
                    } else {
                        "0"
                    },
                )
                .end_element()
                .start_element("w14:checkedState")
                .attribute("w14:val", "2612")
                .attribute("w14:font", "MS Gothic")
                .end_element()
                .start_element("w14:uncheckedState")
                .attribute("w14:val", "2610")
                .attribute("w14:font", "MS Gothic")
                .end_element()
                .end_element();
        }
        "picture" => {
            body.start_element("w:picture").end_element();
        }
        _ => {}
    }
    format!("<w:sdtPr>{}</w:sdtPr>", body.finish())
}

fn serialize_tracked_change(
    change: &TrackedInline,
    context: &mut SerializerContext,
) -> Result<String, ParseError> {
    let (element, deletion) = match change.node_type.as_str() {
        "insertion" => ("w:ins", false),
        "deletion" => ("w:del", true),
        "moveFrom" => ("w:moveFrom", true),
        "moveTo" => ("w:moveTo", false),
        _ => return Ok(String::new()),
    };
    let mut writer = XmlWriter::with_capacity(256);
    writer
        .start_element(element)
        .attribute("w:id", &normalized_tracked_id(change.info.id))
        .attribute(
            "w:author",
            nonempty_trimmed(&change.info.author).unwrap_or("Unknown"),
        );
    if let Some(date) = change.info.date.as_deref().and_then(nonempty_trimmed) {
        writer.attribute("w:date", date);
    }
    // The incumbent always spells tracked wrappers with separate start/end
    // tags, even when filtering leaves their content empty.
    writer.text("");
    for item in &change.content {
        match item {
            InlineNode::Run(run) => append_generated(
                &mut writer,
                &if deletion {
                    serialize_deleted_run(run, context)?
                } else {
                    serialize_run(run, context)?
                },
            ),
            InlineNode::Hyperlink(hyperlink) => {
                append_generated(&mut writer, &serialize_hyperlink(hyperlink, context)?)
            }
            _ => {}
        };
    }
    writer.end_element();
    Ok(writer.finish())
}

fn serialize_range_start(marker: &RangeStart) -> Result<String, ParseError> {
    let element = match marker.node_type.as_str() {
        "moveFromRangeStart" => "w:moveFromRangeStart",
        "moveToRangeStart" => "w:moveToRangeStart",
        _ => return Ok(String::new()),
    };
    let mut writer = XmlWriter::with_capacity(80);
    writer
        .start_element(element)
        .attribute("w:id", &js_number(marker.id))
        .attribute("w:name", &marker.name)
        .end_element();
    Ok(writer.finish())
}

fn serialize_range_end(marker: &RangeEnd) -> Result<String, ParseError> {
    let element = match marker.node_type.as_str() {
        "moveFromRangeEnd" => "w:moveFromRangeEnd",
        "moveToRangeEnd" => "w:moveToRangeEnd",
        _ => return Ok(String::new()),
    };
    let mut writer = XmlWriter::with_capacity(48);
    writer
        .start_element(element)
        .attribute("w:id", &js_number(marker.id))
        .end_element();
    Ok(writer.finish())
}

fn serialize_comment_range(marker: &CommentRange) -> Result<String, ParseError> {
    let mut writer = XmlWriter::with_capacity(180);
    match marker.node_type.as_str() {
        "commentRangeStart" => {
            writer
                .start_element("w:commentRangeStart")
                .attribute("w:id", &js_number(marker.id))
                .end_element();
        }
        "commentRangeEnd" => {
            writer
                .start_element("w:commentRangeEnd")
                .attribute("w:id", &js_number(marker.id))
                .end_element()
                .start_element("w:r")
                .start_element("w:rPr")
                .start_element("w:rStyle")
                .attribute("w:val", "CommentReference")
                .end_element()
                .end_element()
                .start_element("w:commentReference")
                .attribute("w:id", &js_number(marker.id))
                .end_element()
                .end_element();
        }
        _ => {}
    }
    Ok(writer.finish())
}

fn serialize_math(math: &MathEquation) -> Result<String, ParseError> {
    if math.omml_xml.is_empty() {
        return Ok(String::new());
    }
    validate_math_subtree(&math.omml_xml)?;
    Ok(math.omml_xml.clone())
}

fn write_paragraph_mark_properties(
    writer: &mut XmlWriter,
    formatting: Option<&ParagraphFormatting>,
    insertion: Option<&TrackedChangeInfo>,
    deletion: Option<&TrackedChangeInfo>,
) {
    let formatting = formatting
        .and_then(|formatting| formatting.run_properties.as_ref())
        .map(|formatting| serialize_text_formatting(Some(formatting)))
        .unwrap_or_default();
    let inner = formatting
        .strip_prefix("<w:rPr>")
        .and_then(|value| value.strip_suffix("</w:rPr>"))
        .unwrap_or_default();
    if insertion.is_none() && deletion.is_none() && inner.is_empty() {
        return;
    }
    writer.start_element("w:rPr");
    if let Some(info) = insertion {
        write_mark_change(writer, "w:ins", info);
    }
    if let Some(info) = deletion {
        write_mark_change(writer, "w:del", info);
    }
    append_generated(writer, inner);
    writer.end_element();
}

fn write_mark_change(writer: &mut XmlWriter, element: &'static str, info: &TrackedChangeInfo) {
    writer
        .start_element(element)
        .attribute("w:id", &normalized_tracked_id(info.id))
        .attribute(
            "w:author",
            if info.author.is_empty() {
                "Unknown"
            } else {
                &info.author
            },
        );
    if let Some(date) = nonempty(info.date.as_deref()) {
        writer.attribute("w:date", date);
    }
    writer.end_element();
}

fn write_paragraph_borders(writer: &mut XmlWriter, borders: Option<&Borders>) {
    let Some(borders) = borders else {
        return;
    };
    let entries = [
        (borders.top.as_ref(), BorderSide::Top),
        (borders.left.as_ref(), BorderSide::Left),
        (borders.bottom.as_ref(), BorderSide::Bottom),
        (borders.right.as_ref(), BorderSide::Right),
        (borders.between.as_ref(), BorderSide::Between),
        (borders.bar.as_ref(), BorderSide::Bar),
    ];
    if !entries.iter().any(|(border, _)| border.is_some()) {
        return;
    }
    writer.start_element("w:pBdr");
    for (border, side) in entries {
        if let Some(border) = border {
            write_border(writer, border, side);
        }
    }
    writer.end_element();
}

fn write_tabs(writer: &mut XmlWriter, tabs: Option<&[crate::tabs::TabStop]>) {
    let Some(tabs) = tabs.filter(|tabs| !tabs.is_empty()) else {
        return;
    };
    writer.start_element("w:tabs");
    for tab in tabs {
        writer
            .start_element("w:tab")
            .attribute("w:val", &tab.alignment)
            .attribute("w:pos", &int_attr(Some(tab.position)));
        if let Some(leader) = nonempty(tab.leader.as_deref()).filter(|value| *value != "none") {
            writer.attribute("w:leader", leader);
        }
        writer.end_element();
    }
    writer.end_element();
}

fn write_spacing(writer: &mut XmlWriter, formatting: &ParagraphFormatting) {
    if formatting.space_before.is_none()
        && formatting.space_after.is_none()
        && formatting.line_spacing.is_none()
        && nonempty(formatting.line_spacing_rule.as_deref()).is_none()
        && formatting.before_autospacing != Some(true)
        && formatting.after_autospacing != Some(true)
    {
        return;
    }
    writer.start_element("w:spacing");
    optional_int(writer, "w:before", formatting.space_before);
    optional_int(writer, "w:after", formatting.space_after);
    optional_int(writer, "w:line", formatting.line_spacing);
    optional_attr(
        writer,
        "w:lineRule",
        formatting.line_spacing_rule.as_deref(),
    );
    if formatting.before_autospacing == Some(true) {
        writer.attribute("w:beforeAutospacing", "1");
    }
    if formatting.after_autospacing == Some(true) {
        writer.attribute("w:afterAutospacing", "1");
    }
    writer.end_element();
}

fn write_indentation(writer: &mut XmlWriter, formatting: &ParagraphFormatting) {
    let first = formatting.indent_first_line;
    let has_first =
        first.is_some_and(|value| formatting.hanging_indent == Some(true) || value != 0.0);
    if formatting.indent_left.is_none() && formatting.indent_right.is_none() && !has_first {
        return;
    }
    writer.start_element("w:ind");
    optional_int(writer, "w:left", formatting.indent_left);
    optional_int(writer, "w:right", formatting.indent_right);
    if let Some(first) = first {
        if formatting.hanging_indent == Some(true) {
            writer.attribute("w:hanging", &int_attr(Some(first.abs())));
        } else if first != 0.0 {
            writer.attribute("w:firstLine", &int_attr(Some(first)));
        }
    }
    writer.end_element();
}

fn write_numbering(writer: &mut XmlWriter, numbering: Option<&NumberingProperties>) {
    let Some(numbering) = numbering else {
        return;
    };
    if numbering.ilvl.is_none() && numbering.num_id.is_none() {
        return;
    }
    writer.start_element("w:numPr");
    if let Some(value) = numbering.ilvl {
        empty_attr(writer, "w:ilvl", "w:val", &int_attr(Some(value)));
    }
    if let Some(value) = numbering.num_id {
        empty_attr(writer, "w:numId", "w:val", &int_attr(Some(value)));
    }
    writer.end_element();
}

fn write_frame(writer: &mut XmlWriter, frame: Option<&ParagraphFrame>) {
    let Some(frame) = frame else {
        return;
    };
    let has_attributes = frame.width.is_some()
        || frame.height.is_some()
        || nonempty(frame.h_anchor.as_deref()).is_some()
        || nonempty(frame.v_anchor.as_deref()).is_some()
        || frame.x.is_some()
        || frame.y.is_some()
        || nonempty(frame.x_align.as_deref()).is_some()
        || nonempty(frame.y_align.as_deref()).is_some()
        || nonempty(frame.wrap.as_deref()).is_some();
    if !has_attributes {
        return;
    }
    writer.start_element("w:framePr");
    optional_int(writer, "w:w", frame.width);
    optional_int(writer, "w:h", frame.height);
    optional_attr(writer, "w:hAnchor", frame.h_anchor.as_deref());
    optional_attr(writer, "w:vAnchor", frame.v_anchor.as_deref());
    optional_int(writer, "w:x", frame.x);
    optional_int(writer, "w:y", frame.y);
    optional_attr(writer, "w:xAlign", frame.x_align.as_deref());
    optional_attr(writer, "w:yAlign", frame.y_align.as_deref());
    optional_attr(writer, "w:wrap", frame.wrap.as_deref());
    writer.end_element();
}

fn needs_preserve(value: &str) -> bool {
    value.starts_with(' ') || value.ends_with(' ') || value.contains("  ")
}

fn optional_attr(writer: &mut XmlWriter, name: &'static str, value: Option<&str>) {
    if let Some(value) = nonempty(value) {
        writer.attribute(name, value);
    }
}

fn optional_int(writer: &mut XmlWriter, name: &'static str, value: Option<f64>) {
    if value.is_some() {
        writer.attribute(name, &int_attr(value));
    }
}

fn empty_attr(writer: &mut XmlWriter, element: &'static str, name: &'static str, value: &str) {
    writer
        .start_element(element)
        .attribute(name, value)
        .end_element();
}

fn on_off(writer: &mut XmlWriter, name: &'static str, value: Option<bool>) {
    match value {
        Some(true) => {
            writer.start_element(name).end_element();
        }
        Some(false) => {
            writer
                .start_element(name)
                .attribute("w:val", "0")
                .end_element();
        }
        None => {}
    }
}

#[cfg(test)]
mod tests {
    use crate::inline::{Run, RunContent, RunType};
    use crate::serializer::s10::SerializerDeterminism;

    use super::*;

    fn context() -> SerializerContext {
        SerializerContext::new(&SerializerDeterminism {
            seed: "0".repeat(64),
            now: "2000-01-01T00:00:00.000Z".to_owned(),
        })
        .unwrap()
    }

    #[test]
    fn paragraph_bytes_pin_ids_properties_and_rendered_break_injection() {
        let paragraph = Paragraph {
            node_type: "paragraph".to_owned(),
            para_id: Some("AA&BB\"CC".to_owned()),
            text_id: None,
            formatting: Some(ParagraphFormatting {
                keep_next: Some(false),
                alignment: Some("center".to_owned()),
                ..ParagraphFormatting::default()
            }),
            property_changes: None,
            p_pr_ins: None,
            p_pr_del: None,
            content: vec![ParagraphContent::Inline(InlineNode::Run(Run {
                node_type: RunType::Run,
                formatting: None,
                property_changes: None,
                content: vec![RunContent::Text {
                    text: "hello & goodbye".to_owned(),
                    preserve_space: None,
                }],
            }))],
            list_rendering: None,
            rendered_page_break_before: Some(true),
            section_properties: None,
        };
        assert_eq!(
            serialize_paragraph(&paragraph, &mut context()).unwrap(),
            "<w:p w14:paraId=\"AA&amp;BB&quot;CC\"><w:pPr><w:keepNext w:val=\"0\"/><w:jc w:val=\"center\"/></w:pPr><w:r><w:lastRenderedPageBreak/><w:t>hello &amp; goodbye</w:t></w:r></w:p>"
        );
    }

    #[test]
    fn empty_tracked_wrappers_keep_typescript_explicit_end_tags() {
        let change = TrackedInline {
            node_type: "deletion".to_owned(),
            info: TrackedChangeInfo {
                id: 7.0,
                author: "Ada".to_owned(),
                date: None,
            },
            content: Vec::new(),
        };
        assert_eq!(
            serialize_tracked_change(&change, &mut context()).unwrap(),
            "<w:del w:id=\"7\" w:author=\"Ada\"></w:del>"
        );
    }

    #[test]
    fn rejects_unvalidated_math_and_sdt_raw_xml() {
        let mut properties = SdtProperties {
            sdt_type: "richText".to_owned(),
            id: None,
            alias: None,
            tag: None,
            lock: None,
            placeholder: None,
            showing_placeholder: None,
            date_format: None,
            list_items: None,
            checked: None,
            run_properties: None,
            temporary: None,
            label: None,
            tab_index: None,
            multi_line: None,
            date_state: None,
            list_last_value: None,
            checked_state: None,
            unchecked_state: None,
            gallery: None,
            appearance: None,
            color: None,
            control_state: None,
            repeating_section: None,
            repeating_section_item: None,
            data_binding: None,
            raw_properties_xml: Some("<w:sdtPr/><evil/>".to_owned()),
            raw_end_properties_xml: None,
        };
        assert!(serialize_sdt_properties(&properties).is_err());
        properties.raw_properties_xml = None;
        assert!(serialize_sdt_properties(&properties).is_ok());
    }
}
