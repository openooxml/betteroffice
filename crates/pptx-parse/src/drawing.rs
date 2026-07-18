use ooxml_drawingml::{ColorValue, GradientFill, GradientStop, LineEnd, ShapeFill, ShapeOutline};

use crate::PptxError;
use crate::model::*;
use crate::relationships::Relationship;
use crate::xml::{ParseBudget, XmlElement};

const MAX_SAFE_EMU: i64 = 1_000_000_000_000_000;

pub(crate) struct CommonSlideData {
    pub name: Option<String>,
    pub background: Option<ShapeFill>,
    pub shapes: Vec<ShapeNode>,
}

pub(crate) fn common_slide_data(
    root: &XmlElement,
    relationships: &[Relationship],
    part: &str,
    budget: &mut ParseBudget<'_>,
) -> Result<CommonSlideData, PptxError> {
    let common = root.child("cSld");
    let name = common
        .and_then(|value| value.attribute("name"))
        .map(str::to_owned);
    let background = common
        .and_then(|value| value.child("bg"))
        .and_then(parse_background);
    let shapes = if let Some(tree) = common.and_then(|value| value.child("spTree")) {
        parse_shape_children(tree, relationships, part, budget)?
    } else {
        Vec::new()
    };
    Ok(CommonSlideData {
        name,
        background,
        shapes,
    })
}

pub(crate) fn parse_text_styles(root: &XmlElement) -> TextStyleSet {
    let Some(styles) = root.child("txStyles") else {
        return TextStyleSet::default();
    };
    TextStyleSet {
        title: parse_style_levels(styles.child("titleStyle")),
        body: parse_style_levels(styles.child("bodyStyle")),
        other: parse_style_levels(styles.child("otherStyle")),
    }
}

fn parse_style_levels(element: Option<&XmlElement>) -> Vec<ParagraphProperties> {
    let Some(element) = element else {
        return Vec::new();
    };
    let mut levels = vec![ParagraphProperties::default(); 9];
    let mut found = false;
    for child in element.child_elements() {
        let name = child.local_name();
        let Some(level) = name
            .strip_prefix("lvl")
            .and_then(|value| value.strip_suffix("pPr"))
            .and_then(|value| value.parse::<usize>().ok())
        else {
            continue;
        };
        if (1..=9).contains(&level) {
            levels[level - 1] = parse_paragraph_properties(Some(child));
            found = true;
        }
    }
    if found { levels } else { Vec::new() }
}

fn parse_shape_children(
    parent: &XmlElement,
    relationships: &[Relationship],
    part: &str,
    budget: &mut ParseBudget<'_>,
) -> Result<Vec<ShapeNode>, PptxError> {
    let mut shapes = Vec::new();
    for child in parent.child_elements() {
        let shape = match child.local_name() {
            "sp" => Some(ShapeNode::Shape(parse_shape(child, part, budget)?)),
            "pic" => Some(ShapeNode::Picture(parse_picture(
                child,
                relationships,
                part,
                budget,
            )?)),
            "graphicFrame" => Some(ShapeNode::GraphicFrame(parse_graphic_frame(
                child,
                relationships,
                part,
                budget,
            )?)),
            "grpSp" => Some(ShapeNode::Group(parse_group(
                child,
                relationships,
                part,
                budget,
            )?)),
            _ => None,
        };
        if let Some(shape) = shape {
            shapes.push(shape);
        }
    }
    Ok(shapes)
}

fn parse_shape(
    element: &XmlElement,
    part: &str,
    budget: &mut ParseBudget<'_>,
) -> Result<Shape, PptxError> {
    budget.charge_shape(part)?;
    let properties = element.child("spPr");
    Ok(Shape {
        base: parse_base(
            element.child("nvSpPr"),
            properties.and_then(|value| value.child("xfrm")),
        ),
        geometry: parse_geometry(properties),
        fill: properties.and_then(parse_fill),
        outline: properties.and_then(parse_outline),
        text: element
            .child("txBody")
            .map(|body| parse_text_body(body, part, budget))
            .transpose()?,
    })
}

fn parse_picture(
    element: &XmlElement,
    relationships: &[Relationship],
    part: &str,
    budget: &mut ParseBudget<'_>,
) -> Result<Picture, PptxError> {
    budget.charge_shape(part)?;
    let properties = element.child("spPr");
    let blip_fill = element.child("blipFill");
    let relationship_id = blip_fill
        .and_then(|value| value.child("blip"))
        .and_then(|value| {
            value
                .attribute("r:embed")
                .or_else(|| value.attribute_local("embed"))
        })
        .map(str::to_owned);
    let media_part_path = relationship_id
        .as_deref()
        .and_then(|id| relationship_target(relationships, id));
    Ok(Picture {
        base: parse_base(
            element.child("nvPicPr"),
            properties.and_then(|value| value.child("xfrm")),
        ),
        relationship_id,
        media_part_path,
        crop: parse_crop(blip_fill.and_then(|value| value.child("srcRect"))),
        fill: properties.and_then(parse_fill),
        outline: properties.and_then(parse_outline),
    })
}

fn parse_graphic_frame(
    element: &XmlElement,
    relationships: &[Relationship],
    part: &str,
    budget: &mut ParseBudget<'_>,
) -> Result<GraphicFrame, PptxError> {
    budget.charge_shape(part)?;
    let data = element
        .child("graphic")
        .and_then(|value| value.child("graphicData"));
    let frame_data = if let Some(table) = data.and_then(|value| value.child("tbl")) {
        let mut rows = Vec::new();
        for row in table.children_named("tr") {
            let mut cells = Vec::new();
            for cell in row.children_named("tc") {
                cells.push(
                    cell.child("txBody")
                        .map(|body| parse_text_body(body, part, budget))
                        .transpose()?
                        .unwrap_or_default(),
                );
            }
            rows.push(cells);
        }
        GraphicFrameData::Table { rows }
    } else if let Some(chart) =
        data.and_then(|value| value.descendants_named("chart").first().copied())
    {
        let relationship_id = chart
            .attribute("r:id")
            .or_else(|| chart.attribute_local("id"))
            .unwrap_or_default()
            .to_owned();
        GraphicFrameData::Chart {
            part_path: relationship_target(relationships, &relationship_id),
            relationship_id,
        }
    } else if let Some(ids) =
        data.and_then(|value| value.descendants_named("relIds").first().copied())
    {
        let relationship_ids = ids
            .attributes
            .iter()
            .filter(|(key, _)| key.starts_with("r:"))
            .map(|(_, value)| value.clone())
            .collect();
        GraphicFrameData::Diagram { relationship_ids }
    } else {
        GraphicFrameData::Unknown {
            uri: data
                .and_then(|value| value.attribute("uri"))
                .map(str::to_owned),
        }
    };
    Ok(GraphicFrame {
        base: parse_base(element.child("nvGraphicFramePr"), element.child("xfrm")),
        data: frame_data,
    })
}

fn parse_group(
    element: &XmlElement,
    relationships: &[Relationship],
    part: &str,
    budget: &mut ParseBudget<'_>,
) -> Result<GroupShape, PptxError> {
    budget.charge_shape(part)?;
    Ok(GroupShape {
        base: parse_base(
            element.child("nvGrpSpPr"),
            element
                .child("grpSpPr")
                .and_then(|value| value.child("xfrm")),
        ),
        children: parse_shape_children(element, relationships, part, budget)?,
    })
}

fn parse_base(non_visual: Option<&XmlElement>, transform: Option<&XmlElement>) -> ShapeBase {
    let common = non_visual.and_then(|value| value.child("cNvPr"));
    let placeholder = non_visual
        .and_then(|value| value.child("nvPr"))
        .and_then(|value| value.child("ph"))
        .map(parse_placeholder);
    ShapeBase {
        id: common
            .and_then(|value| value.attribute("id"))
            .and_then(|value| value.parse().ok())
            .unwrap_or_default(),
        name: common
            .and_then(|value| value.attribute("name"))
            .unwrap_or_default()
            .to_owned(),
        description: common
            .and_then(|value| value.attribute("descr"))
            .map(str::to_owned),
        hidden: common
            .and_then(|value| value.attribute("hidden"))
            .is_some_and(parse_bool),
        placeholder,
        transform: parse_transform(transform),
    }
}

fn parse_placeholder(element: &XmlElement) -> Placeholder {
    Placeholder {
        placeholder_type: element.attribute("type").map(str::to_owned),
        index: element
            .attribute("idx")
            .and_then(|value| value.parse().ok()),
        orientation: element.attribute("orient").map(str::to_owned),
        size: element.attribute("sz").map(str::to_owned),
    }
}

fn parse_transform(element: Option<&XmlElement>) -> ShapeTransform {
    let Some(element) = element else {
        return ShapeTransform::default();
    };
    let offset = element.child("off");
    let extent = element.child("ext");
    let child_offset = element.child("chOff");
    let child_extent = element.child("chExt");
    ShapeTransform {
        x: numeric_attribute(offset, "x").unwrap_or_default(),
        y: numeric_attribute(offset, "y").unwrap_or_default(),
        width: numeric_attribute(extent, "cx").unwrap_or_default(),
        height: numeric_attribute(extent, "cy").unwrap_or_default(),
        rotation_deg: element
            .attribute("rot")
            .and_then(|value| value.parse::<f64>().ok())
            .filter(|value| value.is_finite())
            .map(|value| value / 60_000.0)
            .unwrap_or_default(),
        flip_h: element.attribute("flipH").is_some_and(parse_bool),
        flip_v: element.attribute("flipV").is_some_and(parse_bool),
        child_x: numeric_attribute(child_offset, "x"),
        child_y: numeric_attribute(child_offset, "y"),
        child_width: numeric_attribute(child_extent, "cx"),
        child_height: numeric_attribute(child_extent, "cy"),
    }
}

fn parse_geometry(properties: Option<&XmlElement>) -> String {
    properties
        .and_then(|value| value.child("prstGeom"))
        .and_then(|value| value.attribute("prst"))
        .map(str::to_owned)
        .or_else(|| {
            properties
                .and_then(|value| value.child("custGeom"))
                .map(|_| "custom".to_owned())
        })
        .unwrap_or_else(|| "rect".to_owned())
}

fn parse_background(element: &XmlElement) -> Option<ShapeFill> {
    if let Some(properties) = element.child("bgPr") {
        return parse_fill(properties);
    }
    element.child("bgRef").map(|reference| ShapeFill {
        fill_type: "theme".to_owned(),
        color: parse_color_container(reference),
        gradient: None,
    })
}

fn parse_fill(element: &XmlElement) -> Option<ShapeFill> {
    if element.child("noFill").is_some() {
        return Some(ShapeFill::named("none"));
    }
    if let Some(fill) = element.child("solidFill") {
        return Some(ShapeFill {
            fill_type: "solid".to_owned(),
            color: parse_color_container(fill),
            gradient: None,
        });
    }
    if let Some(fill) = element.child("gradFill") {
        return Some(parse_gradient_fill(fill));
    }
    if element.child("blipFill").is_some() {
        return Some(ShapeFill::named("picture"));
    }
    None
}

fn parse_gradient_fill(element: &XmlElement) -> ShapeFill {
    let linear = element.child("lin");
    let path = element.child("path");
    let gradient_type = match path.and_then(|value| value.attribute("path")) {
        Some("circle") => "radial",
        Some("rect") => "rectangular",
        Some(_) => "path",
        None => "linear",
    };
    let stops = element
        .child("gsLst")
        .into_iter()
        .flat_map(|list| list.children_named("gs"))
        .filter_map(|stop| {
            Some(GradientStop {
                position: stop
                    .attribute("pos")?
                    .parse::<f64>()
                    .ok()
                    .filter(|value| value.is_finite() && (0.0..=100_000.0).contains(value))?,
                color: parse_color_container(stop)?,
            })
        })
        .collect();
    ShapeFill {
        fill_type: "gradient".to_owned(),
        color: None,
        gradient: Some(GradientFill {
            gradient_type: gradient_type.to_owned(),
            angle: linear
                .and_then(|value| value.attribute("ang"))
                .and_then(|value| value.parse::<f64>().ok())
                .filter(|value| value.is_finite())
                .map(|value| value / 60_000.0),
            stops,
        }),
    }
}

fn parse_outline(element: &XmlElement) -> Option<ShapeOutline> {
    let line = element.child("ln")?;
    if line.child("noFill").is_some() {
        return None;
    }
    Some(ShapeOutline {
        width: line.attribute("w").and_then(|value| value.parse().ok()),
        color: line.child("solidFill").and_then(parse_color_container),
        style: line
            .child("prstDash")
            .and_then(|value| value.attribute("val"))
            .map(str::to_owned),
        cap: line.attribute("cap").map(str::to_owned),
        join: line
            .child_elements()
            .find(|value| matches!(value.local_name(), "round" | "bevel" | "miter"))
            .map(|value| value.local_name().to_owned()),
        head_end: line.child("headEnd").map(parse_line_end),
        tail_end: line.child("tailEnd").map(parse_line_end),
    })
}

fn parse_line_end(element: &XmlElement) -> LineEnd {
    LineEnd {
        end_type: element.attribute("type").unwrap_or("none").to_owned(),
        width: element.attribute("w").map(str::to_owned),
        length: element.attribute("len").map(str::to_owned),
    }
}

fn parse_color_container(element: &XmlElement) -> Option<ColorValue> {
    let color = element.child_elements().find(|value| {
        matches!(
            value.local_name(),
            "srgbClr" | "schemeClr" | "sysClr" | "prstClr"
        )
    })?;
    let mut parsed = match color.local_name() {
        "srgbClr" => ColorValue {
            rgb: color.attribute("val").map(str::to_owned),
            ..ColorValue::default()
        },
        "schemeClr" => ColorValue {
            theme_color: color.attribute("val").map(normalize_scheme_color),
            ..ColorValue::default()
        },
        "sysClr" => ColorValue {
            rgb: color
                .attribute("lastClr")
                .or_else(|| system_color(color.attribute("val")))
                .map(str::to_owned),
            ..ColorValue::default()
        },
        "prstClr" => ColorValue {
            rgb: color
                .attribute("val")
                .and_then(preset_color)
                .map(str::to_owned),
            ..ColorValue::default()
        },
        _ => return None,
    };
    parsed.theme_tint = color.child("tint").and_then(color_modifier);
    parsed.theme_shade = color.child("shade").and_then(color_modifier);
    Some(parsed)
}

fn normalize_scheme_color(value: &str) -> String {
    match value {
        "tx1" => "text1",
        "tx2" => "text2",
        "bg1" => "background1",
        "bg2" => "background2",
        value => value,
    }
    .to_owned()
}

fn system_color(value: Option<&str>) -> Option<&'static str> {
    match value? {
        "windowText" | "menuText" | "captionText" | "btnText" => Some("000000"),
        "window" | "menu" | "btnFace" | "btnHighlight" | "highlightText" => Some("FFFFFF"),
        "highlight" => Some("0078D7"),
        "grayText" => Some("808080"),
        _ => None,
    }
}

fn preset_color(value: &str) -> Option<&'static str> {
    match value {
        "black" => Some("000000"),
        "white" => Some("FFFFFF"),
        "red" => Some("FF0000"),
        "green" => Some("008000"),
        "blue" => Some("0000FF"),
        "yellow" => Some("FFFF00"),
        "cyan" => Some("00FFFF"),
        "magenta" => Some("FF00FF"),
        _ => None,
    }
}

fn color_modifier(element: &XmlElement) -> Option<String> {
    let value = element.attribute("val")?.parse::<f64>().ok()?;
    if !value.is_finite() || !(0.0..=100_000.0).contains(&value) {
        return None;
    }
    Some(format!(
        "{:02X}",
        (value / 100_000.0 * 255.0).round() as i64
    ))
}

fn parse_crop(element: Option<&XmlElement>) -> PictureCrop {
    let Some(element) = element else {
        return PictureCrop::default();
    };
    PictureCrop {
        left: integer_attribute(element, "l").unwrap_or_default(),
        top: integer_attribute(element, "t").unwrap_or_default(),
        right: integer_attribute(element, "r").unwrap_or_default(),
        bottom: integer_attribute(element, "b").unwrap_or_default(),
    }
}

pub(crate) fn parse_text_body(
    element: &XmlElement,
    part: &str,
    budget: &mut ParseBudget<'_>,
) -> Result<TextBody, PptxError> {
    let body_properties = element.child("bodyPr");
    let mut paragraphs = Vec::new();
    for paragraph in element.children_named("p") {
        budget.charge_paragraph(part)?;
        paragraphs.push(parse_text_paragraph(paragraph, part, budget)?);
    }
    Ok(TextBody {
        anchor: body_properties
            .and_then(|value| value.attribute("anchor"))
            .map(str::to_owned),
        vertical: body_properties
            .and_then(|value| value.attribute("vert"))
            .map(str::to_owned),
        autofit: body_properties.and_then(parse_text_autofit),
        inset_left: numeric_attribute(body_properties, "lIns"),
        inset_top: numeric_attribute(body_properties, "tIns"),
        inset_right: numeric_attribute(body_properties, "rIns"),
        inset_bottom: numeric_attribute(body_properties, "bIns"),
        paragraphs,
    })
}

fn parse_text_autofit(body_properties: &XmlElement) -> Option<TextAutofit> {
    if body_properties.child("noAutofit").is_some() {
        return Some(TextAutofit::None);
    }
    if body_properties.child("spAutoFit").is_some() {
        return Some(TextAutofit::Shape);
    }
    body_properties
        .child("normAutofit")
        .map(|autofit| TextAutofit::Normal {
            font_scale: percentage_attribute(autofit, "fontScale"),
            line_space_reduction: percentage_attribute(autofit, "lnSpcReduction"),
        })
}

fn percentage_attribute(element: &XmlElement, name: &str) -> Option<f64> {
    element
        .attribute(name)?
        .parse::<f64>()
        .ok()
        .filter(|value| value.is_finite() && (0.0..=100_000.0).contains(value))
        .map(|value| value / 100_000.0)
}

fn parse_text_paragraph(
    element: &XmlElement,
    part: &str,
    budget: &mut ParseBudget<'_>,
) -> Result<TextParagraph, PptxError> {
    let mut runs = Vec::new();
    for child in element.child_elements() {
        match child.local_name() {
            "r" | "fld" => {
                budget.charge_run(part)?;
                runs.push(parse_text_run(child));
            }
            "br" => {
                budget.charge_run(part)?;
                runs.push(TextRun {
                    text: "\n".to_owned(),
                    properties: parse_run_properties(child.child("rPr")),
                    field_id: None,
                    field_type: None,
                    line_break: true,
                });
            }
            _ => {}
        }
    }
    Ok(TextParagraph {
        properties: parse_paragraph_properties(element.child("pPr")),
        runs,
        end_properties: element
            .child("endParaRPr")
            .map(|value| parse_run_properties(Some(value))),
    })
}

fn parse_paragraph_properties(element: Option<&XmlElement>) -> ParagraphProperties {
    let Some(element) = element else {
        return ParagraphProperties::default();
    };
    let bullet = if element.child("buNone").is_some() {
        Some(Bullet::None)
    } else if let Some(character) = element.child("buChar") {
        character.attribute("char").map(|value| Bullet::Character {
            value: value.to_owned(),
        })
    } else {
        element.child("buAutoNum").map(|value| Bullet::AutoNumber {
            scheme: value.attribute("type").unwrap_or("arabicPeriod").to_owned(),
            start_at: value
                .attribute("startAt")
                .and_then(|value| value.parse().ok())
                .unwrap_or(1),
        })
    };
    ParagraphProperties {
        alignment: element.attribute("algn").map(str::to_owned),
        level: element
            .attribute("lvl")
            .and_then(|value| value.parse().ok())
            .unwrap_or_default(),
        margin_left: numeric_attribute(Some(element), "marL"),
        indent: numeric_attribute(Some(element), "indent"),
        bullet,
        default_run: element
            .child("defRPr")
            .map(|value| parse_run_properties(Some(value))),
    }
}

fn parse_text_run(element: &XmlElement) -> TextRun {
    TextRun {
        text: element
            .child("t")
            .map(XmlElement::text_content)
            .unwrap_or_default(),
        properties: parse_run_properties(element.child("rPr")),
        field_id: (element.local_name() == "fld")
            .then(|| element.attribute("id").map(str::to_owned))
            .flatten(),
        field_type: (element.local_name() == "fld")
            .then(|| element.attribute("type").map(str::to_owned))
            .flatten(),
        line_break: false,
    }
}

fn parse_run_properties(element: Option<&XmlElement>) -> RunProperties {
    let Some(element) = element else {
        return RunProperties::default();
    };
    RunProperties {
        font_size_pt: element
            .attribute("sz")
            .and_then(|value| value.parse::<f64>().ok())
            .filter(|value| value.is_finite())
            .map(|value| value / 100.0),
        bold: element.attribute("b").map(parse_bool),
        italic: element.attribute("i").map(parse_bool),
        underline: element.attribute("u").map(str::to_owned),
        font_family: element
            .child("latin")
            .and_then(|value| value.attribute("typeface"))
            .map(str::to_owned),
        color: element.child("solidFill").and_then(parse_color_container),
        language: element.attribute("lang").map(str::to_owned),
        hyperlink_relationship_id: element
            .child("hlinkClick")
            .and_then(|value| {
                value
                    .attribute("r:id")
                    .or_else(|| value.attribute_local("id"))
            })
            .map(str::to_owned),
    }
}

fn relationship_target(relationships: &[Relationship], id: &str) -> Option<String> {
    relationships
        .iter()
        .find(|relationship| relationship.id == id)
        .and_then(|relationship| relationship.resolved_target.clone())
}

fn numeric_attribute(element: Option<&XmlElement>, name: &str) -> Option<i64> {
    let value = element?.attribute(name)?.parse::<i64>().ok()?;
    (value.unsigned_abs() <= MAX_SAFE_EMU as u64).then_some(value)
}

fn integer_attribute(element: &XmlElement, name: &str) -> Option<i32> {
    element.attribute(name)?.parse().ok()
}

fn parse_bool(value: &str) -> bool {
    matches!(value, "1" | "true" | "on")
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::ParseLimits;
    use crate::xml::parse_xml;

    #[test]
    fn parses_text_formatting_and_nested_shape_types() {
        let limits = ParseLimits::default();
        let mut budget = ParseBudget::new(&limits);
        let root = parse_xml(
            br#"<p:sld><p:cSld name="Test"><p:spTree><p:sp><p:nvSpPr><p:cNvPr id="2" name="Title"/><p:nvPr><p:ph type="title"/></p:nvPr></p:nvSpPr><p:spPr><a:xfrm><a:off x="1" y="2"/><a:ext cx="3" cy="4"/></a:xfrm><a:prstGeom prst="roundRect"/><a:solidFill><a:schemeClr val="accent1"/></a:solidFill></p:spPr><p:txBody><a:bodyPr anchor="ctr"><a:normAutofit fontScale="85000" lnSpcReduction="12000"/></a:bodyPr><a:p><a:pPr algn="ctr"/><a:r><a:rPr sz="2400" b="1"><a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill><a:latin typeface="Aptos"/></a:rPr><a:t>Hello</a:t></a:r></a:p></p:txBody></p:sp></p:spTree></p:cSld></p:sld>"#,
            "ppt/slides/slide1.xml",
            &mut budget,
        )
        .unwrap();
        let data = common_slide_data(&root, &[], "ppt/slides/slide1.xml", &mut budget).unwrap();
        let ShapeNode::Shape(shape) = &data.shapes[0] else {
            panic!("expected shape");
        };
        assert_eq!(shape.geometry, "roundRect");
        assert_eq!(shape.base.transform.width, 3);
        assert_eq!(
            shape.text.as_ref().unwrap().autofit,
            Some(TextAutofit::Normal {
                font_scale: Some(0.85),
                line_space_reduction: Some(0.12),
            })
        );
        assert_eq!(
            shape.text.as_ref().unwrap().paragraphs[0].runs[0].text,
            "Hello"
        );
        assert_eq!(
            shape.text.as_ref().unwrap().paragraphs[0].runs[0]
                .properties
                .font_size_pt,
            Some(24.0)
        );
    }
}
