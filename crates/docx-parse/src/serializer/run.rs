//! Run, text-formatting, image, and shape serializers.

use crate::drawingml::{ShapeFill, ShapeOutline};
use crate::formatting::TextFormatting;
use crate::image::{Image, ImagePosition, ImageWrap};
use crate::inline::{Run, RunContent, RunPropertyChange};
use crate::paragraph::Paragraph;
use crate::scalars::{ColorValue, ShadingProperties};
use crate::shape::Shape;
use crate::xml::ParseError;

use super::context::SerializerContext;
use super::xml_writer::{XmlWriter, int_attr, js_number};

const VALID_HIGHLIGHT_COLORS: &[&str] = &[
    "black",
    "blue",
    "cyan",
    "darkBlue",
    "darkCyan",
    "darkGray",
    "darkGreen",
    "darkMagenta",
    "darkRed",
    "darkYellow",
    "green",
    "lightGray",
    "magenta",
    "red",
    "white",
    "yellow",
];

/// Serialize the incumbent `w:rPr` subset in its exact child order.
pub fn serialize_text_formatting(formatting: Option<&TextFormatting>) -> String {
    let Some(formatting) = formatting else {
        return String::new();
    };
    let mut body = XmlWriter::with_capacity(384);

    if let Some(value) = nonempty(formatting.style_id.as_deref()) {
        empty_attr(&mut body, "w:rStyle", "w:val", value);
    }
    if let Some(fonts) = formatting.font_family.as_ref() {
        let has_attributes = [
            fonts.ascii.as_deref(),
            fonts.h_ansi.as_deref(),
            fonts.east_asia.as_deref(),
            fonts.cs.as_deref(),
            fonts.ascii_theme.as_deref(),
            fonts.h_ansi_theme.as_deref(),
            fonts.east_asia_theme.as_deref(),
            fonts.cs_theme.as_deref(),
        ]
        .into_iter()
        .any(|value| nonempty(value).is_some());
        if has_attributes {
            body.start_element("w:rFonts");
            optional_nonempty_attr(&mut body, "w:ascii", fonts.ascii.as_deref());
            optional_nonempty_attr(&mut body, "w:hAnsi", fonts.h_ansi.as_deref());
            optional_nonempty_attr(&mut body, "w:eastAsia", fonts.east_asia.as_deref());
            optional_nonempty_attr(&mut body, "w:cs", fonts.cs.as_deref());
            optional_nonempty_attr(&mut body, "w:asciiTheme", fonts.ascii_theme.as_deref());
            optional_nonempty_attr(&mut body, "w:hAnsiTheme", fonts.h_ansi_theme.as_deref());
            optional_nonempty_attr(
                &mut body,
                "w:eastAsiaTheme",
                fonts.east_asia_theme.as_deref(),
            );
            optional_nonempty_attr(&mut body, "w:csTheme", fonts.cs_theme.as_deref());
            body.end_element();
        }
    }

    on_off(&mut body, "w:b", formatting.bold);
    on_off(&mut body, "w:bCs", formatting.bold_cs);
    on_off(&mut body, "w:i", formatting.italic);
    on_off(&mut body, "w:iCs", formatting.italic_cs);
    on_off(&mut body, "w:caps", formatting.all_caps);
    on_off(&mut body, "w:smallCaps", formatting.small_caps);
    on_off(&mut body, "w:strike", formatting.strike);
    on_off(&mut body, "w:dstrike", formatting.double_strike);
    on_off(&mut body, "w:outline", formatting.outline);
    on_off(&mut body, "w:shadow", formatting.shadow);
    on_off(&mut body, "w:emboss", formatting.emboss);
    on_off(&mut body, "w:imprint", formatting.imprint);
    on_off(&mut body, "w:vanish", formatting.hidden);

    write_color_element(&mut body, formatting.color.as_ref());
    optional_integer_element(&mut body, "w:spacing", formatting.spacing);
    optional_integer_element(&mut body, "w:w", formatting.scale);
    optional_integer_element(&mut body, "w:kern", formatting.kerning);
    optional_integer_element(&mut body, "w:position", formatting.position);
    optional_integer_element(&mut body, "w:sz", formatting.font_size);
    optional_integer_element(&mut body, "w:szCs", formatting.font_size_cs);

    if let Some(highlight) = nonempty(formatting.highlight.as_deref()).filter(|v| *v != "none") {
        if VALID_HIGHLIGHT_COLORS.contains(&highlight) {
            empty_attr(&mut body, "w:highlight", "w:val", highlight);
        } else if formatting.shading.is_none() {
            let hex = highlight.strip_prefix('#').unwrap_or(highlight);
            if hex.len() == 6 && hex.bytes().all(|byte| byte.is_ascii_hexdigit()) {
                body.start_element("w:shd")
                    .attribute("w:val", "clear")
                    .attribute("w:color", "auto")
                    .attribute("w:fill", hex)
                    .end_element();
            }
        }
    }

    if let Some(underline) = formatting.underline.as_ref() {
        body.start_element("w:u")
            .attribute("w:val", &underline.style);
        if let Some(color) = underline.color.as_ref() {
            optional_nonempty_attr(&mut body, "w:color", color.rgb.as_deref());
            optional_nonempty_attr(&mut body, "w:themeColor", color.theme_color.as_deref());
            optional_nonempty_attr(&mut body, "w:themeTint", color.theme_tint.as_deref());
            optional_nonempty_attr(&mut body, "w:themeShade", color.theme_shade.as_deref());
        }
        body.end_element();
    }
    if let Some(value) = nonempty(formatting.effect.as_deref()).filter(|value| *value != "none") {
        empty_attr(&mut body, "w:effect", "w:val", value);
    }
    if let Some(value) =
        nonempty(formatting.emphasis_mark.as_deref()).filter(|value| *value != "none")
    {
        empty_attr(&mut body, "w:em", "w:val", value);
    }
    write_shading(&mut body, formatting.shading.as_ref());
    if let Some(value) = nonempty(formatting.vert_align.as_deref()).filter(|v| *v != "baseline") {
        empty_attr(&mut body, "w:vertAlign", "w:val", value);
    }
    on_off(&mut body, "w:rtl", formatting.rtl);
    on_off(&mut body, "w:cs", formatting.cs);

    let body = body.finish();
    if body.is_empty() {
        String::new()
    } else {
        format!("<w:rPr>{body}</w:rPr>")
    }
}

/// Serialize one run, consuming any active rendered-page-break markers.
pub fn serialize_run(run: &Run, context: &mut SerializerContext) -> Result<String, ParseError> {
    let mut output = String::from("<w:r>");
    for _ in 0..context.take_rendered_page_breaks() {
        output.push_str("<w:lastRenderedPageBreak/>");
    }
    output.push_str(&serialize_run_properties(
        run.formatting.as_ref(),
        run.property_changes.as_deref(),
    ));
    for content in &run.content {
        output.push_str(&serialize_run_content(content, context)?);
    }
    output.push_str("</w:r>");
    Ok(output)
}

pub(crate) fn serialize_deleted_run(
    run: &Run,
    context: &mut SerializerContext,
) -> Result<String, ParseError> {
    let xml = serialize_run(run, context)?;
    if run
        .content
        .iter()
        .any(|content| matches!(content, RunContent::Drawing { .. }))
    {
        return Ok(xml);
    }
    Ok(xml
        // Match the incumbent `/\<w:t\b/` replacement without corrupting
        // sibling names such as `<w:tab/>`.
        .replace("<w:t>", "<w:delText>")
        .replace("<w:t ", "<w:delText ")
        .replace("<w:t/", "<w:delText/")
        .replace("</w:t>", "</w:delText>")
        .replace("<w:instrText>", "<w:delInstrText>")
        .replace("<w:instrText ", "<w:delInstrText ")
        .replace("<w:instrText/", "<w:delInstrText/")
        .replace("</w:instrText>", "</w:delInstrText>"))
}

fn serialize_run_properties(
    formatting: Option<&TextFormatting>,
    changes: Option<&[RunPropertyChange]>,
) -> String {
    let current = serialize_text_formatting(formatting);
    let mut inner = strip_wrapper(&current, "<w:rPr>", "</w:rPr>").to_owned();
    for change in changes.unwrap_or_default() {
        inner.push_str(&serialize_run_property_change(change));
    }
    if inner.is_empty() {
        String::new()
    } else {
        format!("<w:rPr>{inner}</w:rPr>")
    }
}

fn serialize_run_property_change(change: &RunPropertyChange) -> String {
    let mut writer = XmlWriter::with_capacity(128);
    let id = normalized_tracked_id(change.info.id);
    let author = nonempty_trimmed(&change.info.author).unwrap_or("Unknown");
    writer
        .start_element("w:rPrChange")
        .attribute("w:id", &id)
        .attribute("w:author", author);
    if let Some(date) = change.info.date.as_deref().and_then(nonempty_trimmed) {
        writer.attribute("w:date", date);
    }
    let previous = serialize_text_formatting(change.previous_formatting.as_ref());
    if previous.is_empty() {
        writer.start_element("w:rPr").end_element();
    } else {
        append_generated(&mut writer, &previous);
    }
    writer.end_element();
    writer.finish()
}

fn serialize_run_content(
    content: &RunContent,
    context: &mut SerializerContext,
) -> Result<String, ParseError> {
    let mut writer = XmlWriter::with_capacity(96);
    match content {
        RunContent::Text {
            text,
            preserve_space,
        } => {
            writer.start_element("w:t");
            if *preserve_space == Some(true)
                || text.starts_with(' ')
                || text.ends_with(' ')
                || text.contains("  ")
            {
                writer.attribute("xml:space", "preserve");
            }
            writer.text(text).end_element();
        }
        RunContent::Tab => {
            writer.start_element("w:tab").end_element();
        }
        RunContent::Break { break_type, clear } => {
            writer.start_element("w:br");
            match break_type.as_deref() {
                Some("page") => {
                    writer.attribute("w:type", "page");
                }
                Some("column") => {
                    writer.attribute("w:type", "column");
                }
                Some("textWrapping") => {
                    writer.attribute("w:type", "textWrapping");
                    if let Some(clear) = nonempty(clear.as_deref()).filter(|value| *value != "none")
                    {
                        writer.attribute("w:clear", clear);
                    }
                }
                _ => {}
            }
            writer.end_element();
        }
        RunContent::Symbol { font, char } => {
            writer
                .start_element("w:sym")
                .attribute("w:font", font)
                .attribute("w:char", char)
                .end_element();
        }
        RunContent::FootnoteRef { id, .. } => {
            writer
                .start_element("w:footnoteReference")
                .attribute("w:id", &js_number(*id))
                .end_element();
        }
        RunContent::EndnoteRef { id, .. } => {
            writer
                .start_element("w:endnoteReference")
                .attribute("w:id", &js_number(*id))
                .end_element();
        }
        RunContent::FootnoteRefMark => {
            writer.start_element("w:footnoteRef").end_element();
        }
        RunContent::EndnoteRefMark => {
            writer.start_element("w:endnoteRef").end_element();
        }
        RunContent::Separator => {
            writer.start_element("w:separator").end_element();
        }
        RunContent::ContinuationSeparator => {
            writer
                .start_element("w:continuationSeparator")
                .end_element();
        }
        RunContent::FieldChar {
            char_type,
            fld_lock,
            dirty,
            ..
        } => {
            writer
                .start_element("w:fldChar")
                .attribute("w:fldCharType", char_type);
            if *fld_lock == Some(true) {
                writer.attribute("w:fldLock", "true");
            }
            if *dirty == Some(true) {
                writer.attribute("w:dirty", "true");
            }
            writer.end_element();
        }
        RunContent::InstrText { text } => {
            writer.start_element("w:instrText");
            if text.starts_with(' ') || text.ends_with(' ') || text.contains("  ") {
                writer.attribute("xml:space", "preserve");
            }
            writer.text(text).end_element();
        }
        RunContent::SoftHyphen => {
            writer.start_element("w:softHyphen").end_element();
        }
        RunContent::NoBreakHyphen => {
            writer.start_element("w:noBreakHyphen").end_element();
        }
        RunContent::Drawing { image } => return serialize_drawing_content(image, context),
        RunContent::Shape { shape } => return serialize_shape_content(shape, context),
        RunContent::CommentReference { .. }
        | RunContent::Chart { .. }
        | RunContent::OpaqueDrawing { .. } => {}
    }
    Ok(writer.finish())
}

/// Serialize one image as incumbent-compatible WordprocessingDrawing XML.
pub fn serialize_drawing_content(
    image: &Image,
    context: &mut SerializerContext,
) -> Result<String, ParseError> {
    let floating = image.wrap.wrap_type != "inline";
    let id = drawing_id(image.id.as_deref(), context);
    let name = image
        .title
        .as_deref()
        .filter(|value| !value.is_empty())
        .or_else(|| image.filename.as_deref().filter(|value| !value.is_empty()))
        .map(str::to_owned)
        .unwrap_or_else(|| format!("Picture {id}"));
    let graphic = serialize_picture_graphic(image, &id);
    let mut writer = XmlWriter::with_capacity(1024 + graphic.len());
    writer.start_element("w:drawing");
    if floating {
        writer
            .start_element("wp:anchor")
            .attribute("distT", &int_attr(image.wrap.dist_t))
            .attribute("distB", &int_attr(image.wrap.dist_b))
            .attribute("distL", &int_attr(image.wrap.dist_l))
            .attribute("distR", &int_attr(image.wrap.dist_r))
            .attribute("simplePos", "0")
            .attribute("relativeHeight", "251658240")
            .attribute(
                "behindDoc",
                if image.wrap.wrap_type == "behind" {
                    "1"
                } else {
                    "0"
                },
            )
            .attribute("locked", "0")
            .attribute(
                "layoutInCell",
                if image.layout_in_cell == Some(false) {
                    "0"
                } else {
                    "1"
                },
            )
            .attribute(
                "allowOverlap",
                if image.allow_overlap == Some(false) {
                    "0"
                } else {
                    "1"
                },
            );
        writer
            .start_element("wp:simplePos")
            .attribute("x", "0")
            .attribute("y", "0")
            .end_element();
        write_position(&mut writer, image.position.as_ref());
        write_extent(&mut writer, image.size.width, image.size.height);
        write_effect_extent(&mut writer, image.padding.as_ref());
        write_wrap(&mut writer, &image.wrap);
    } else {
        writer
            .start_element("wp:inline")
            .attribute("distT", &int_attr(image.wrap.dist_t))
            .attribute("distB", &int_attr(image.wrap.dist_b))
            .attribute("distL", &int_attr(image.wrap.dist_l))
            .attribute("distR", &int_attr(image.wrap.dist_r));
        write_extent(&mut writer, image.size.width, image.size.height);
        write_effect_extent(&mut writer, image.padding.as_ref());
    }
    writer
        .start_element("wp:docPr")
        .attribute("id", &id)
        .attribute("name", &name);
    if let Some(alt) = nonempty(image.alt.as_deref()) {
        writer.attribute("descr", alt);
    }
    if image.decorative == Some(true) {
        writer.attribute("hidden", "1");
    }
    writer.end_element();
    writer.start_element("wp:cNvGraphicFramePr");
    writer
        .start_element("a:graphicFrameLocks")
        .attribute(
            "xmlns:a",
            "http://schemas.openxmlformats.org/drawingml/2006/main",
        )
        .attribute("noChangeAspect", "1")
        .end_element();
    writer.end_element();
    append_generated(&mut writer, &graphic);
    writer.end_element().end_element();
    Ok(writer.finish())
}

/// Serialize one WordprocessingShape, including recursively typed textbox
/// paragraphs stored behind the model's cycle-breaking JSON values.
pub fn serialize_shape_content(
    shape: &Shape,
    context: &mut SerializerContext,
) -> Result<String, ParseError> {
    let is_text_box = shape.shape_type == "textBox";
    let floating = shape
        .wrap
        .as_ref()
        .is_some_and(|wrap| wrap.wrap_type != "inline");
    let id = drawing_id(shape.id.as_deref(), context);
    let name = shape
        .name
        .as_deref()
        .filter(|value| !value.is_empty())
        .map(str::to_owned)
        .unwrap_or_else(|| format!("{} {id}", if is_text_box { "TextBox" } else { "Shape" }));
    let wrap = shape.wrap.as_ref();
    let mut graphic = XmlWriter::with_capacity(1024);
    graphic
        .start_element("a:graphic")
        .attribute(
            "xmlns:a",
            "http://schemas.openxmlformats.org/drawingml/2006/main",
        )
        .start_element("a:graphicData")
        .attribute(
            "uri",
            "http://schemas.microsoft.com/office/word/2010/wordprocessingShape",
        )
        .start_element("wps:wsp")
        .start_element("wps:cNvSpPr");
    if is_text_box {
        graphic.attribute("txBox", "1");
    }
    graphic.end_element().start_element("wps:spPr");
    write_transform(
        &mut graphic,
        shape.transform.as_ref(),
        shape.size.width,
        shape.size.height,
    );
    graphic
        .start_element("a:prstGeom")
        .attribute(
            "prst",
            if is_text_box {
                "rect"
            } else {
                &shape.shape_type
            },
        )
        .start_element("a:avLst")
        .end_element()
        .end_element();
    write_fill(&mut graphic, shape.fill.as_ref());
    write_outline(&mut graphic, shape.outline.as_ref());
    graphic.end_element();

    if let Some(text_body) = shape.text_body.as_ref() {
        let mut body_properties = XmlWriter::with_capacity(160);
        body_properties
            .start_element("wps:bodyPr")
            .attribute("rot", "0")
            .attribute("vert", "horz");
        if let Some(anchor) = nonempty(text_body.anchor.as_deref()) {
            body_properties.attribute("anchor", if anchor == "middle" { "ctr" } else { anchor });
        }
        if text_body.anchor_center == Some(true) {
            body_properties.attribute("anchorCtr", "1");
        }
        if let Some(margins) = text_body.margins.as_ref() {
            optional_int_attr(&mut body_properties, "lIns", margins.left);
            optional_int_attr(&mut body_properties, "tIns", margins.top);
            optional_int_attr(&mut body_properties, "rIns", margins.right);
            optional_int_attr(&mut body_properties, "bIns", margins.bottom);
        }
        body_properties.end_element();
        if is_text_box {
            graphic
                .start_element("wps:txbx")
                .start_element("w:txbxContent");
            for value in &text_body.content {
                let paragraph: Paragraph =
                    serde_json::from_value(value.clone()).map_err(|error| {
                        ParseError::Canonical(format!(
                            "shape text body contains an invalid paragraph: {error}"
                        ))
                    })?;
                let paragraph = super::paragraph::serialize_paragraph(&paragraph, context)?;
                append_generated(&mut graphic, &paragraph);
            }
            graphic.end_element().end_element();
        }
        append_generated(&mut graphic, &body_properties.finish());
    }
    graphic.end_element().end_element().end_element();
    let graphic = graphic.finish();

    let mut writer = XmlWriter::with_capacity(graphic.len() + 512);
    writer.start_element("w:drawing");
    if floating {
        let wrap = wrap.expect("floating shapes always have wrap properties");
        writer
            .start_element("wp:anchor")
            .attribute("distT", &int_attr(wrap.dist_t))
            .attribute("distB", &int_attr(wrap.dist_b))
            .attribute("distL", &int_attr(wrap.dist_l))
            .attribute("distR", &int_attr(wrap.dist_r))
            .attribute("simplePos", "0")
            .attribute("relativeHeight", "251658240")
            .attribute(
                "behindDoc",
                if wrap.wrap_type == "behind" { "1" } else { "0" },
            )
            .attribute("locked", "0")
            .attribute("layoutInCell", "1")
            .attribute("allowOverlap", "1");
        writer
            .start_element("wp:simplePos")
            .attribute("x", "0")
            .attribute("y", "0")
            .end_element();
        write_position(&mut writer, shape.position.as_ref());
        write_extent(&mut writer, shape.size.width, shape.size.height);
        write_zero_effect_extent(&mut writer);
        write_wrap(&mut writer, wrap);
    } else {
        writer
            .start_element("wp:inline")
            .attribute("distT", &int_attr(wrap.and_then(|value| value.dist_t)))
            .attribute("distB", &int_attr(wrap.and_then(|value| value.dist_b)))
            .attribute("distL", &int_attr(wrap.and_then(|value| value.dist_l)))
            .attribute("distR", &int_attr(wrap.and_then(|value| value.dist_r)));
        write_extent(&mut writer, shape.size.width, shape.size.height);
        write_zero_effect_extent(&mut writer);
    }
    writer
        .start_element("wp:docPr")
        .attribute("id", &id)
        .attribute("name", &name)
        .end_element()
        .start_element("wp:cNvGraphicFramePr")
        .end_element();
    append_generated(&mut writer, &graphic);
    writer.end_element().end_element();
    Ok(writer.finish())
}

fn serialize_picture_graphic(image: &Image, id: &str) -> String {
    let relationship_id = if image.relationship_id.is_empty() {
        "rId1"
    } else {
        &image.relationship_id
    };
    let name = image
        .filename
        .as_deref()
        .filter(|value| !value.is_empty())
        .map(str::to_owned)
        .unwrap_or_else(|| format!("image{id}"));
    let mut writer = XmlWriter::with_capacity(768);
    writer
        .start_element("a:graphic")
        .attribute(
            "xmlns:a",
            "http://schemas.openxmlformats.org/drawingml/2006/main",
        )
        .start_element("a:graphicData")
        .attribute(
            "uri",
            "http://schemas.openxmlformats.org/drawingml/2006/picture",
        )
        .start_element("pic:pic")
        .attribute(
            "xmlns:pic",
            "http://schemas.openxmlformats.org/drawingml/2006/picture",
        )
        .start_element("pic:nvPicPr")
        .start_element("pic:cNvPr")
        .attribute("id", id)
        .attribute("name", &name);
    if let Some(alt) = nonempty(image.alt.as_deref()) {
        writer.attribute("descr", alt);
    }
    writer
        .end_element()
        .start_element("pic:cNvPicPr")
        .end_element()
        .end_element()
        .start_element("pic:blipFill")
        .start_element("a:blip")
        .attribute("r:embed", relationship_id);
    if image.opacity.is_some_and(|opacity| opacity < 1.0) {
        let amount = image.opacity.unwrap_or_default().clamp(0.0, 1.0) * 100_000.0;
        writer
            .start_element("a:alphaModFix")
            .attribute("amt", &int_attr(Some(amount)))
            .end_element();
    }
    writer.end_element();
    if let Some(crop) = image.crop.as_ref() {
        let has_crop = [crop.left, crop.top, crop.right, crop.bottom]
            .into_iter()
            .flatten()
            .any(|value| value != 0.0);
        if has_crop {
            writer.start_element("a:srcRect");
            optional_truthy_scaled_attr(&mut writer, "l", crop.left, 100_000.0);
            optional_truthy_scaled_attr(&mut writer, "t", crop.top, 100_000.0);
            optional_truthy_scaled_attr(&mut writer, "r", crop.right, 100_000.0);
            optional_truthy_scaled_attr(&mut writer, "b", crop.bottom, 100_000.0);
            writer.end_element();
        }
    }
    writer
        .start_element("a:stretch")
        .start_element("a:fillRect")
        .end_element()
        .end_element()
        .end_element()
        .start_element("pic:spPr");
    write_transform(
        &mut writer,
        image.transform.as_ref(),
        image.size.width,
        image.size.height,
    );
    writer
        .start_element("a:prstGeom")
        .attribute("prst", "rect")
        .start_element("a:avLst")
        .end_element()
        .end_element();
    if let Some(outline) = image.outline.as_ref() {
        write_outline(&mut writer, Some(outline));
    }
    writer
        .end_element()
        .end_element()
        .end_element()
        .end_element();
    writer.finish()
}

fn write_position(writer: &mut XmlWriter, position: Option<&ImagePosition>) {
    if let Some(position) = position {
        write_axis(
            writer,
            "wp:positionH",
            &position.horizontal.relative_to,
            position.horizontal.alignment.as_deref(),
            position.horizontal.pos_offset,
        );
        write_axis(
            writer,
            "wp:positionV",
            &position.vertical.relative_to,
            position.vertical.alignment.as_deref(),
            position.vertical.pos_offset,
        );
    } else {
        write_axis(writer, "wp:positionH", "column", None, Some(0.0));
        write_axis(writer, "wp:positionV", "paragraph", None, Some(0.0));
    }
}

fn write_axis(
    writer: &mut XmlWriter,
    element: &'static str,
    relative_to: &str,
    alignment: Option<&str>,
    offset: Option<f64>,
) {
    writer
        .start_element(element)
        .attribute("relativeFrom", relative_to);
    if let Some(alignment) = nonempty(alignment) {
        writer
            .start_element("wp:align")
            .text(alignment)
            .end_element();
    } else {
        writer
            .start_element("wp:posOffset")
            .text(&int_attr(offset))
            .end_element();
    }
    writer.end_element();
}

fn write_wrap(writer: &mut XmlWriter, wrap: &ImageWrap) {
    let text = nonempty(wrap.wrap_text.as_deref()).unwrap_or("bothSides");
    match wrap.wrap_type.as_str() {
        "square" => {
            writer
                .start_element("wp:wrapSquare")
                .attribute("wrapText", text)
                .end_element();
        }
        "tight" => write_polygon_wrap(writer, "wp:wrapTight", text),
        "through" => write_polygon_wrap(writer, "wp:wrapThrough", text),
        "topAndBottom" => {
            writer.start_element("wp:wrapTopAndBottom").end_element();
        }
        _ => {
            writer.start_element("wp:wrapNone").end_element();
        }
    }
}

fn write_polygon_wrap(writer: &mut XmlWriter, name: &'static str, text: &str) {
    writer
        .start_element(name)
        .attribute("wrapText", text)
        .start_element("wp:wrapPolygon")
        .attribute("edited", "0")
        .start_element("wp:start")
        .attribute("x", "0")
        .attribute("y", "0")
        .end_element();
    for (x, y) in [(0, 21_600), (21_600, 21_600), (21_600, 0), (0, 0)] {
        writer
            .start_element("wp:lineTo")
            .attribute("x", &x.to_string())
            .attribute("y", &y.to_string())
            .end_element();
    }
    writer.end_element().end_element();
}

fn write_transform(
    writer: &mut XmlWriter,
    transform: Option<&crate::drawingml::Transform2D>,
    width: f64,
    height: f64,
) {
    writer.start_element("a:xfrm");
    if let Some(rotation) = transform
        .and_then(|value| value.rotation)
        .filter(|value| *value != 0.0)
    {
        writer.attribute("rot", &int_attr(Some(rotation * 60_000.0)));
    }
    if transform.is_some_and(|value| value.flip_h == Some(true)) {
        writer.attribute("flipH", "1");
    }
    if transform.is_some_and(|value| value.flip_v == Some(true)) {
        writer.attribute("flipV", "1");
    }
    writer
        .start_element("a:off")
        .attribute("x", "0")
        .attribute("y", "0")
        .end_element()
        .start_element("a:ext")
        .attribute("cx", &int_attr(Some(width)))
        .attribute("cy", &int_attr(Some(height)))
        .end_element()
        .end_element();
}

fn write_fill(writer: &mut XmlWriter, fill: Option<&ShapeFill>) {
    let Some(fill) = fill else {
        writer.start_element("a:noFill").end_element();
        return;
    };
    match fill.fill_type.as_str() {
        "none" => {
            writer.start_element("a:noFill").end_element();
        }
        "solid" if fill.color.is_some() => {
            writer.start_element("a:solidFill");
            write_drawing_color(writer, fill.color.as_ref());
            writer.end_element();
        }
        "gradient" if fill.gradient.is_some() => {
            let gradient = fill.gradient.as_ref().unwrap();
            writer.start_element("a:gradFill").start_element("a:gsLst");
            for stop in &gradient.stops {
                writer
                    .start_element("a:gs")
                    .attribute("pos", &js_number(stop.position));
                write_drawing_color(writer, Some(&stop.color));
                writer.end_element();
            }
            writer.end_element();
            if gradient.gradient_type == "linear" {
                writer
                    .start_element("a:lin")
                    .attribute(
                        "ang",
                        &js_number(gradient.angle.unwrap_or_default() * 60_000.0),
                    )
                    .attribute("scaled", "1")
                    .end_element();
            }
            writer.end_element();
        }
        _ => {}
    }
}

fn write_outline(writer: &mut XmlWriter, outline: Option<&ShapeOutline>) {
    let Some(outline) = outline else {
        return;
    };
    let has_parts = outline.color.is_some()
        || outline
            .style
            .as_deref()
            .is_some_and(|style| !style.is_empty() && style != "solid")
        || outline.head_end.is_some()
        || outline.tail_end.is_some();
    if outline.width.is_none() && nonempty(outline.cap.as_deref()).is_none() && !has_parts {
        return;
    }
    writer.start_element("a:ln");
    if let Some(width) = outline.width {
        writer.attribute("w", &js_number(width));
    }
    optional_nonempty_attr(writer, "cap", outline.cap.as_deref());
    if let Some(color) = outline.color.as_ref() {
        writer.start_element("a:solidFill");
        write_drawing_color(writer, Some(color));
        writer.end_element();
    }
    if let Some(style) = nonempty(outline.style.as_deref()).filter(|style| *style != "solid") {
        writer
            .start_element("a:prstDash")
            .attribute("val", style)
            .end_element();
    }
    if let Some(end) = outline.head_end.as_ref() {
        writer
            .start_element("a:headEnd")
            .attribute("type", &end.end_type);
        optional_nonempty_attr(writer, "w", end.width.as_deref());
        optional_nonempty_attr(writer, "len", end.length.as_deref());
        writer.end_element();
    }
    if let Some(end) = outline.tail_end.as_ref() {
        writer
            .start_element("a:tailEnd")
            .attribute("type", &end.end_type);
        optional_nonempty_attr(writer, "w", end.width.as_deref());
        optional_nonempty_attr(writer, "len", end.length.as_deref());
        writer.end_element();
    }
    writer.end_element();
}

fn write_drawing_color(writer: &mut XmlWriter, color: Option<&ColorValue>) {
    let Some(color) = color else {
        return;
    };
    if let Some(rgb) = nonempty(color.rgb.as_deref()) {
        writer
            .start_element("a:srgbClr")
            .attribute("val", &rgb.replacen('#', "", 1))
            .end_element();
    } else if let Some(theme) = nonempty(color.theme_color.as_deref()) {
        writer.start_element("a:schemeClr").attribute("val", theme);
        if let Some(tint) = nonempty(color.theme_tint.as_deref()) {
            writer
                .start_element("a:tint")
                .attribute("val", tint)
                .end_element();
        } else if let Some(shade) = nonempty(color.theme_shade.as_deref()) {
            writer
                .start_element("a:shade")
                .attribute("val", shade)
                .end_element();
        }
        writer.end_element();
    }
}

pub(crate) fn write_shading(writer: &mut XmlWriter, shading: Option<&ShadingProperties>) {
    let Some(shading) = shading else {
        return;
    };
    writer
        .start_element("w:shd")
        .attribute("w:val", shading.pattern.as_deref().unwrap_or("clear"));
    if let Some(rgb) = shading
        .color
        .as_ref()
        .and_then(|color| color.rgb.as_deref())
    {
        writer.attribute("w:color", rgb);
    } else if shading.color.as_ref().and_then(|color| color.auto) == Some(true) {
        writer.attribute("w:color", "auto");
    }
    if let Some(rgb) = shading.fill.as_ref().and_then(|color| color.rgb.as_deref()) {
        writer.attribute("w:fill", rgb);
    } else if shading.fill.as_ref().and_then(|color| color.auto) == Some(true) {
        writer.attribute("w:fill", "auto");
    }
    if let Some(fill) = shading.fill.as_ref() {
        optional_nonempty_attr(writer, "w:themeFill", fill.theme_color.as_deref());
        optional_nonempty_attr(writer, "w:themeFillTint", fill.theme_tint.as_deref());
        optional_nonempty_attr(writer, "w:themeFillShade", fill.theme_shade.as_deref());
    }
    writer.end_element();
}

fn write_color_element(writer: &mut XmlWriter, color: Option<&ColorValue>) {
    let Some(color) = color else {
        return;
    };
    if color.auto != Some(true)
        && nonempty(color.rgb.as_deref()).is_none()
        && nonempty(color.theme_color.as_deref()).is_none()
        && nonempty(color.theme_tint.as_deref()).is_none()
        && nonempty(color.theme_shade.as_deref()).is_none()
    {
        return;
    }
    writer.start_element("w:color");
    if color.auto == Some(true) {
        writer.attribute("w:val", "auto");
    } else if let Some(rgb) = nonempty(color.rgb.as_deref()) {
        writer.attribute("w:val", rgb);
    }
    optional_nonempty_attr(writer, "w:themeColor", color.theme_color.as_deref());
    optional_nonempty_attr(writer, "w:themeTint", color.theme_tint.as_deref());
    optional_nonempty_attr(writer, "w:themeShade", color.theme_shade.as_deref());
    writer.end_element();
}

fn write_extent(writer: &mut XmlWriter, width: f64, height: f64) {
    writer
        .start_element("wp:extent")
        .attribute("cx", &int_attr(Some(width)))
        .attribute("cy", &int_attr(Some(height)))
        .end_element();
}

fn write_effect_extent(writer: &mut XmlWriter, padding: Option<&crate::image::ImagePadding>) {
    writer
        .start_element("wp:effectExtent")
        .attribute("l", &int_attr(padding.and_then(|value| value.left)))
        .attribute("t", &int_attr(padding.and_then(|value| value.top)))
        .attribute("r", &int_attr(padding.and_then(|value| value.right)))
        .attribute("b", &int_attr(padding.and_then(|value| value.bottom)))
        .end_element();
}

fn write_zero_effect_extent(writer: &mut XmlWriter) {
    writer
        .start_element("wp:effectExtent")
        .attribute("l", "0")
        .attribute("t", "0")
        .attribute("r", "0")
        .attribute("b", "0")
        .end_element();
}

fn drawing_id(source: Option<&str>, context: &mut SerializerContext) -> String {
    nonempty(source)
        .map(str::to_owned)
        .unwrap_or_else(|| context.allocate_drawing_id())
}

pub(crate) fn normalized_tracked_id(id: f64) -> String {
    if id.is_finite() && id >= 0.0 && id.fract() == 0.0 {
        js_number(id)
    } else {
        "0".to_owned()
    }
}

pub(crate) fn nonempty(value: Option<&str>) -> Option<&str> {
    value.filter(|value| !value.is_empty())
}

pub(crate) fn nonempty_trimmed(value: &str) -> Option<&str> {
    let value = value.trim();
    (!value.is_empty()).then_some(value)
}

pub(crate) fn append_generated(writer: &mut XmlWriter, xml: &str) {
    writer.append_serialized(xml);
}

fn strip_wrapper<'a>(value: &'a str, start: &str, end: &str) -> &'a str {
    value
        .strip_prefix(start)
        .and_then(|value| value.strip_suffix(end))
        .unwrap_or_default()
}

fn optional_nonempty_attr(writer: &mut XmlWriter, name: &'static str, value: Option<&str>) {
    if let Some(value) = nonempty(value) {
        writer.attribute(name, value);
    }
}

fn optional_int_attr(writer: &mut XmlWriter, name: &'static str, value: Option<f64>) {
    if value.is_some() {
        writer.attribute(name, &int_attr(value));
    }
}

fn optional_truthy_scaled_attr(
    writer: &mut XmlWriter,
    name: &'static str,
    value: Option<f64>,
    scale: f64,
) {
    if let Some(value) = value.filter(|value| *value != 0.0) {
        writer.attribute(name, &int_attr(Some(value * scale)));
    }
}

fn optional_integer_element(writer: &mut XmlWriter, name: &'static str, value: Option<f64>) {
    if value.is_some() {
        writer
            .start_element(name)
            .attribute("w:val", &int_attr(value))
            .end_element();
    }
}

fn empty_attr(writer: &mut XmlWriter, element: &'static str, attr: &'static str, value: &str) {
    writer
        .start_element(element)
        .attribute(attr, value)
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
    use crate::formatting::{FontFamily, TextFormatting};
    use crate::inline::{Run, RunType};
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
    fn run_bytes_match_typescript_and_escape_all_model_strings() {
        let run = Run {
            node_type: RunType::Run,
            formatting: Some(TextFormatting {
                style_id: Some("bad\"/><evil&".to_owned()),
                font_family: Some(FontFamily {
                    ascii: Some("A&B".to_owned()),
                    ..FontFamily::default()
                }),
                bold: Some(false),
                italic: Some(true),
                font_size: Some(23.5),
                ..TextFormatting::default()
            }),
            property_changes: None,
            content: vec![RunContent::Text {
                text: " <hello> & ".to_owned(),
                preserve_space: None,
            }],
        };
        assert_eq!(
            serialize_run(&run, &mut context()).unwrap(),
            "<w:r><w:rPr><w:rStyle w:val=\"bad&quot;/&gt;&lt;evil&amp;\"/><w:rFonts w:ascii=\"A&amp;B\"/><w:b w:val=\"0\"/><w:i/><w:sz w:val=\"24\"/></w:rPr><w:t xml:space=\"preserve\"> &lt;hello&gt; &amp; </w:t></w:r>"
        );
    }

    #[test]
    fn generated_drawing_ids_consume_the_seeded_allocator() {
        let mut first = context();
        let mut second = context();
        let first_id = first.allocate_drawing_id();
        assert_eq!(first_id, second.allocate_drawing_id());
        let next_id = first.allocate_drawing_id();
        assert_eq!(next_id, second.allocate_drawing_id());
        assert_ne!(first_id, next_id);
    }

    #[test]
    fn deleted_text_rewrite_does_not_corrupt_tab_elements() {
        let run = Run {
            node_type: RunType::Run,
            formatting: None,
            property_changes: None,
            content: vec![
                RunContent::Tab,
                RunContent::Text {
                    text: " deleted ".to_owned(),
                    preserve_space: None,
                },
            ],
        };
        assert_eq!(
            serialize_deleted_run(&run, &mut context()).unwrap(),
            "<w:r><w:tab/><w:delText xml:space=\"preserve\"> deleted </w:delText></w:r>"
        );
    }
}
