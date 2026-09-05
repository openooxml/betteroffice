use std::collections::BTreeMap;

use ooxml_drawingml::{
    ColorValue, GradientFill, GradientStop, LineEnd, ShapeFill, ShapeOutline, ShapeStyle,
    StyleReference,
};

use crate::PptxError;
use crate::custom_geometry::parse_custom_geometry;
use crate::model::*;
use crate::relationships::Relationship;
use crate::xml::{ParseBudget, XmlElement};

const MAX_SAFE_EMU: i64 = 1_000_000_000_000_000;
const ANGLE_UNITS_PER_DEGREE: f64 = 60_000.0;
const ADJUSTMENT_SCALE: f64 = 100_000.0;

#[derive(Clone, Copy)]
struct GuideValue {
    value: f64,
    extent_power: f64,
}

impl GuideValue {
    fn scalar(value: f64) -> Self {
        Self {
            value,
            extent_power: 0.0,
        }
    }

    fn extent(value: f64) -> Self {
        Self {
            value,
            extent_power: 1.0,
        }
    }
}

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
    elements: ShapeElements,
) -> Result<CommonSlideData, PptxError> {
    let common = root.child("cSld");
    let name = common
        .and_then(|value| value.attribute("name"))
        .map(str::to_owned);
    let background = common
        .and_then(|value| value.child("bg"))
        .and_then(parse_background);
    let mut shapes = if let Some(tree) = common.and_then(|value| value.child("spTree")) {
        parse_shape_children(tree, relationships, part, budget, elements)?
    } else {
        Vec::new()
    };
    resolve_group_fill(&mut shapes, None);
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
    elements: ShapeElements,
) -> Result<Vec<ShapeNode>, PptxError> {
    let mut shapes = Vec::new();
    for child in parent.child_elements() {
        let shape = match child.local_name() {
            "cxnSp" if elements == ShapeElements::WithoutConnectors => None,
            "sp" | "cxnSp" => Some(ShapeNode::Shape(parse_shape(child, part, budget)?)),
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
                elements,
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
    let transform = properties.and_then(|value| value.child("xfrm"));
    Ok(Shape {
        paths: properties
            .filter(|value| value.child("prstGeom").is_none())
            .and_then(|value| value.child("custGeom"))
            .and_then(|custom| parse_custom_geometry(custom, parse_shape_extent(transform)))
            .unwrap_or_default(),
        base: parse_base(
            element
                .child("nvSpPr")
                .or_else(|| element.child("nvCxnSpPr")),
            transform,
        ),
        geometry: parse_geometry(properties),
        adjust_values: parse_adjust_values(properties, parse_shape_extent(transform)),
        fill: properties.and_then(parse_fill),
        outline: properties.and_then(parse_outline),
        style: parse_shape_style(element.child("style"), properties).map(Box::new),
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
    let transform = properties.and_then(|value| value.child("xfrm"));
    Ok(Picture {
        base: parse_base(element.child("nvPicPr"), transform),
        relationship_id,
        media_part_path,
        crop: parse_crop(blip_fill.and_then(|value| value.child("srcRect"))),
        geometry: parse_geometry(properties),
        adjust_values: parse_adjust_values(properties, parse_shape_extent(transform)),
        fill: properties.and_then(parse_fill),
        outline: properties.and_then(parse_outline),
        style: parse_shape_style(element.child("style"), properties).map(Box::new),
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
    elements: ShapeElements,
) -> Result<GroupShape, PptxError> {
    budget.charge_shape(part)?;
    let own_fill = element.child("grpSpPr").and_then(parse_fill);
    let mut children = parse_shape_children(element, relationships, part, budget, elements)?;
    if let Some(fill) = own_fill
        .as_ref()
        .filter(|fill| fill.fill_type != GROUP_FILL)
    {
        resolve_group_fill(&mut children, Some(fill));
    }
    Ok(GroupShape {
        base: parse_base(
            element.child("nvGrpSpPr"),
            element
                .child("grpSpPr")
                .and_then(|value| value.child("xfrm")),
        ),
        children,
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

fn parse_shape_extent(transform: Option<&XmlElement>) -> Option<(f64, f64)> {
    let extent = transform?.child("ext")?;
    let width = numeric_attribute(Some(extent), "cx")?;
    let height = numeric_attribute(Some(extent), "cy")?;
    (width > 0 && height > 0).then_some((width as f64, height as f64))
}

fn parse_adjust_values(
    properties: Option<&XmlElement>,
    extent: Option<(f64, f64)>,
) -> BTreeMap<String, f64> {
    let Some(adjustment_list) = properties
        .and_then(|value| value.child("prstGeom"))
        .and_then(|value| value.child("avLst"))
    else {
        return BTreeMap::new();
    };
    let mut values = extent.map(standard_guide_values).unwrap_or_default();
    let mut adjustments = BTreeMap::new();
    for guide in adjustment_list
        .child_elements()
        .filter(|value| value.local_name() == "gd")
    {
        let (Some(name), Some(formula)) = (guide.attribute("name"), guide.attribute("fmla")) else {
            continue;
        };
        let Some(value) = evaluate_guide_formula(formula, &values) else {
            continue;
        };
        values.insert(name.to_owned(), value);
        let denominator = if value.extent_power == 0.0 {
            ADJUSTMENT_SCALE
        } else {
            let Some((width, height)) = extent else {
                continue;
            };
            width.min(height).powf(value.extent_power)
        };
        let adjustment = value.value / denominator;
        if adjustment.is_finite() {
            adjustments.insert(name.to_owned(), adjustment);
        }
    }
    adjustments
}

fn standard_guide_values((width, height): (f64, f64)) -> BTreeMap<String, GuideValue> {
    let short = width.min(height);
    let long = width.max(height);
    let mut values = BTreeMap::from([
        ("w".to_owned(), GuideValue::extent(width)),
        ("h".to_owned(), GuideValue::extent(height)),
        ("ss".to_owned(), GuideValue::extent(short)),
        ("ls".to_owned(), GuideValue::extent(long)),
        ("hc".to_owned(), GuideValue::extent(width / 2.0)),
        ("vc".to_owned(), GuideValue::extent(height / 2.0)),
        ("l".to_owned(), GuideValue::extent(0.0)),
        ("t".to_owned(), GuideValue::extent(0.0)),
        ("r".to_owned(), GuideValue::extent(width)),
        ("b".to_owned(), GuideValue::extent(height)),
    ]);
    for divisor in [2, 3, 4, 5, 6, 8, 10, 12, 32] {
        values.insert(
            format!("wd{divisor}"),
            GuideValue::extent(width / divisor as f64),
        );
    }
    for divisor in [2, 3, 4, 5, 6, 8] {
        values.insert(
            format!("hd{divisor}"),
            GuideValue::extent(height / divisor as f64),
        );
    }
    for divisor in [2, 4, 6, 8, 16, 32] {
        values.insert(
            format!("ssd{divisor}"),
            GuideValue::extent(short / divisor as f64),
        );
    }
    let circle = 360.0 * ANGLE_UNITS_PER_DEGREE;
    for (name, numerator, denominator) in [
        ("cd2", 1.0, 2.0),
        ("cd4", 1.0, 4.0),
        ("cd8", 1.0, 8.0),
        ("3cd4", 3.0, 4.0),
        ("3cd8", 3.0, 8.0),
        ("5cd8", 5.0, 8.0),
        ("7cd8", 7.0, 8.0),
    ] {
        values.insert(
            name.to_owned(),
            GuideValue::scalar(circle * numerator / denominator),
        );
    }
    values
}

fn evaluate_guide_formula(
    formula: &str,
    values: &BTreeMap<String, GuideValue>,
) -> Option<GuideValue> {
    let mut tokens = formula.split_whitespace();
    let operator = tokens.next()?;
    let operands = tokens
        .map(|token| guide_operand(token, values))
        .collect::<Option<Vec<_>>>()?;
    let result = match (operator, operands.as_slice()) {
        ("val", [x]) => *x,
        ("*/", [x, y, z]) if z.value != 0.0 => GuideValue {
            value: x.value * y.value / z.value,
            extent_power: x.extent_power + y.extent_power - z.extent_power,
        },
        ("+-", [x, y, z]) => GuideValue {
            value: x.value + y.value - z.value,
            extent_power: additive_extent_power(&[*x, *y, *z]),
        },
        ("+/", [x, y, z]) if z.value != 0.0 => GuideValue {
            value: (x.value + y.value) / z.value,
            extent_power: additive_extent_power(&[*x, *y]) - z.extent_power,
        },
        ("?:", [x, y, z]) => {
            if x.value > 0.0 {
                *y
            } else {
                *z
            }
        }
        ("abs", [x]) => GuideValue {
            value: x.value.abs(),
            ..*x
        },
        ("at2", [x, y]) => {
            GuideValue::scalar(y.value.atan2(x.value).to_degrees() * ANGLE_UNITS_PER_DEGREE)
        }
        ("cat2", [x, y, z]) => GuideValue {
            value: x.value * z.value.atan2(y.value).cos(),
            ..*x
        },
        ("cos", [x, y]) => GuideValue {
            value: x.value * (y.value / ANGLE_UNITS_PER_DEGREE).to_radians().cos(),
            ..*x
        },
        ("max", [x, y]) => {
            if x.value >= y.value {
                *x
            } else {
                *y
            }
        }
        ("min", [x, y]) => {
            if x.value <= y.value {
                *x
            } else {
                *y
            }
        }
        ("mod", [x, y, z]) => GuideValue {
            value: x.value.hypot(y.value).hypot(z.value),
            extent_power: additive_extent_power(&[*x, *y, *z]),
        },
        ("pin", [x, y, z]) => {
            if y.value < x.value {
                *x
            } else if y.value > z.value {
                *z
            } else {
                *y
            }
        }
        ("sat2", [x, y, z]) => GuideValue {
            value: x.value * z.value.atan2(y.value).sin(),
            ..*x
        },
        ("sin", [x, y]) => GuideValue {
            value: x.value * (y.value / ANGLE_UNITS_PER_DEGREE).to_radians().sin(),
            ..*x
        },
        ("sqrt", [x]) if x.value >= 0.0 => GuideValue {
            value: x.value.sqrt(),
            extent_power: x.extent_power / 2.0,
        },
        ("tan", [x, y]) => GuideValue {
            value: x.value * (y.value / ANGLE_UNITS_PER_DEGREE).to_radians().tan(),
            ..*x
        },
        _ => return None,
    };
    result.value.is_finite().then_some(result)
}

fn additive_extent_power(values: &[GuideValue]) -> f64 {
    values
        .iter()
        .find_map(|value| (value.extent_power != 0.0).then_some(value.extent_power))
        .unwrap_or_default()
}

fn guide_operand(token: &str, values: &BTreeMap<String, GuideValue>) -> Option<GuideValue> {
    token
        .parse::<f64>()
        .ok()
        .filter(|value| value.is_finite())
        .map(GuideValue::scalar)
        .or_else(|| values.get(token).copied())
}

fn parse_background(element: &XmlElement) -> Option<ShapeFill> {
    if let Some(properties) = element.child("bgPr") {
        return parse_fill(properties).filter(|fill| fill.fill_type != GROUP_FILL);
    }
    element.child("bgRef").map(|reference| ShapeFill {
        fill_type: "theme".to_owned(),
        color: parse_color_container(reference),
        gradient: None,
    })
}

/// Parses theme references and explicit fill barriers.
fn parse_shape_style(
    element: Option<&XmlElement>,
    properties: Option<&XmlElement>,
) -> Option<ShapeStyle> {
    let reference = |name| {
        element
            .and_then(|element| element.child(name))
            .map(|reference| StyleReference {
                index: reference
                    .attribute("idx")
                    .and_then(|value| value.parse().ok())
                    .unwrap_or_default(),
                color: parse_color_container(reference),
            })
    };
    let unsupported = |element: Option<&XmlElement>, allowed: &[&str]| {
        element.is_some_and(|element| {
            element
                .child_elements()
                .any(|child| is_fill_element(child) && !allowed.contains(&child.local_name()))
        })
    };
    let style = ShapeStyle {
        font_color: element
            .and_then(|value| value.child("fontRef"))
            .and_then(parse_color_container),
        fill: reference("fillRef"),
        line: reference("lnRef"),
        fill_disabled: unsupported(
            properties,
            &["solidFill", "gradFill", "blipFill", "noFill", "grpFill"],
        ),
        line_disabled: unsupported(
            properties.and_then(|value| value.child("ln")),
            &["solidFill"],
        ),
    };
    (!style.is_empty()).then_some(style)
}

fn is_fill_element(element: &XmlElement) -> bool {
    matches!(
        element.local_name(),
        "noFill" | "solidFill" | "gradFill" | "blipFill" | "pattFill" | "grpFill"
    )
}

fn parse_fill(element: &XmlElement) -> Option<ShapeFill> {
    element.child_elements().find_map(parse_fill_element)
}

/// Parses a shape or theme fill.
pub(crate) fn parse_fill_element(element: &XmlElement) -> Option<ShapeFill> {
    match element.local_name() {
        "noFill" => Some(ShapeFill::named("none")),
        "solidFill" => Some(ShapeFill {
            fill_type: "solid".to_owned(),
            color: parse_color_container(element),
            gradient: None,
        }),
        "gradFill" => Some(parse_gradient_fill(element)),
        "blipFill" => Some(ShapeFill::named("picture")),
        "grpFill" => Some(ShapeFill::named(GROUP_FILL)),
        _ => None,
    }
}

/// Marker left by `<a:grpFill/>`, standing until an ancestor group resolves it or the tree
/// root clears it.
const GROUP_FILL: &str = "group";

fn fill_slot(node: &mut ShapeNode) -> Option<&mut Option<ShapeFill>> {
    match node {
        ShapeNode::Shape(shape) => Some(&mut shape.fill),
        ShapeNode::Picture(picture) => Some(&mut picture.fill),
        ShapeNode::GraphicFrame(_) | ShapeNode::Group(_) => None,
    }
}

/// Hands `fill` to every descendant still waiting on `<a:grpFill/>`, or clears the marker when
/// there is no group left to inherit from.
fn resolve_group_fill(children: &mut [ShapeNode], fill: Option<&ShapeFill>) {
    for child in children {
        if let Some(slot) = fill_slot(child)
            && slot
                .as_ref()
                .is_some_and(|current| current.fill_type == GROUP_FILL)
        {
            *slot = fill.cloned();
        }
        if let ShapeNode::Group(group) = child {
            resolve_group_fill(&mut group.children, fill);
        }
    }
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
    parse_outline_element(line)
}

pub(crate) fn parse_outline_element(line: &XmlElement) -> Option<ShapeOutline> {
    if line.local_name() != "ln" {
        return None;
    }
    if line.child("noFill").is_some() {
        return Some(ShapeOutline::default());
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

pub(crate) fn parse_color_container(element: &XmlElement) -> Option<ColorValue> {
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
    parsed.luminance_modulation = color.child("lumMod").and_then(color_fraction);
    parsed.luminance_offset = color.child("lumOff").and_then(color_fraction);
    parsed.saturation_modulation = color.child("satMod").and_then(color_fraction);
    parsed.alpha = color.child("alpha").and_then(color_fraction);
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

/// A `ST_Percentage` in thousandths of a percent, as a fraction.
fn color_fraction(element: &XmlElement) -> Option<f64> {
    let value = element.attribute("val")?.trim();
    let value = value
        .strip_suffix('%')
        .map(|percent| percent.parse::<f64>().map(|value| value * 1_000.0))
        .unwrap_or_else(|| value.parse::<f64>())
        .ok()?;
    (value.is_finite() && (0.0..=1_000_000.0).contains(&value)).then_some(value / 100_000.0)
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

pub(crate) fn parse_run_properties(element: Option<&XmlElement>) -> RunProperties {
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
    use ooxml_drawingml::GeometryPathCommand;

    #[test]
    fn a_custom_geometry_becomes_a_path_normalised_to_the_shape() {
        let limits = ParseLimits::default();
        let mut budget = ParseBudget::new(&limits);
        let root = parse_xml(
            br#"<p:sld><p:cSld><p:spTree><p:sp><p:nvSpPr><p:cNvPr id="2" name="Freeform"/><p:nvPr/></p:nvSpPr><p:spPr><a:custGeom><a:pathLst><a:path w="200" h="100"><a:moveTo><a:pt x="0" y="0"/></a:moveTo><a:lnTo><a:pt x="200" y="50"/></a:lnTo><a:cubicBezTo><a:pt x="150" y="100"/><a:pt x="50" y="100"/><a:pt x="0" y="50"/></a:cubicBezTo><a:close/></a:path></a:pathLst></a:custGeom></p:spPr></p:sp><p:sp><p:nvSpPr><p:cNvPr id="3" name="Guided"/><p:nvPr/></p:nvSpPr><p:spPr><a:custGeom><a:gdLst><a:gd name="x1" fmla="*/ w 1 2"/></a:gdLst><a:pathLst><a:path w="10" h="10"><a:moveTo><a:pt x="0" y="0"/></a:moveTo><a:lnTo><a:pt x="x1" y="10"/></a:lnTo></a:path></a:pathLst></a:custGeom></p:spPr></p:sp></p:spTree></p:cSld></p:sld>"#,
            "ppt/slides/slide1.xml",
            &mut budget,
        )
        .unwrap();
        let data = common_slide_data(
            &root,
            &[],
            "ppt/slides/slide1.xml",
            &mut budget,
            ShapeElements::WithConnectors,
        )
        .unwrap();

        let ShapeNode::Shape(freeform) = &data.shapes[0] else {
            panic!("expected a shape");
        };
        assert_eq!(freeform.geometry, "custom");
        let path = &freeform.paths[0].commands;
        assert_eq!(
            path[0],
            GeometryPathCommand::Move { x: 0.0, y: 0.0 },
            "coordinates are fractions of the path box, not raw units"
        );
        assert_eq!(path[1], GeometryPathCommand::Line { x: 1.0, y: 0.5 });
        assert_eq!(
            path[2],
            GeometryPathCommand::Cubic {
                cp1x: 0.75,
                cp1y: 1.0,
                cp2x: 0.25,
                cp2y: 1.0,
                x: 0.0,
                y: 0.5
            }
        );
        assert_eq!(path[3], GeometryPathCommand::Close);

        let ShapeNode::Shape(guided) = &data.shapes[1] else {
            panic!("expected a shape");
        };
        assert!(guided.paths.is_empty());
    }

    #[test]
    fn a_shape_style_supplies_the_default_text_colour() {
        let limits = ParseLimits::default();
        let mut budget = ParseBudget::new(&limits);
        let root = parse_xml(
            br#"<p:sld><p:cSld><p:spTree><p:sp><p:nvSpPr><p:cNvPr id="2" name="Styled"/><p:nvPr/></p:nvSpPr><p:spPr/><p:style><a:lnRef idx="2"><a:schemeClr val="accent1"/></a:lnRef><a:fillRef idx="1"><a:schemeClr val="accent1"/></a:fillRef><a:effectRef idx="0"><a:schemeClr val="accent1"/></a:effectRef><a:fontRef idx="minor"><a:schemeClr val="lt1"/></a:fontRef></p:style></p:sp><p:sp><p:nvSpPr><p:cNvPr id="3" name="Plain"/><p:nvPr/></p:nvSpPr><p:spPr/></p:sp></p:spTree></p:cSld></p:sld>"#,
            "ppt/slides/slide1.xml",
            &mut budget,
        )
        .unwrap();
        let data = common_slide_data(
            &root,
            &[],
            "ppt/slides/slide1.xml",
            &mut budget,
            ShapeElements::WithConnectors,
        )
        .unwrap();

        let ShapeNode::Shape(styled) = &data.shapes[0] else {
            panic!("expected a shape");
        };
        let color = styled
            .style
            .as_ref()
            .and_then(|style| style.font_color.as_ref())
            .expect("fontRef supplies a colour");
        assert_eq!(color.theme_color.as_deref(), Some("lt1"));

        let ShapeNode::Shape(plain) = &data.shapes[1] else {
            panic!("expected a shape");
        };
        assert!(plain.style.is_none());
    }

    #[test]
    fn a_font_ref_without_a_colour_leaves_the_text_colour_unset() {
        let limits = ParseLimits::default();
        let mut budget = ParseBudget::new(&limits);
        let root = parse_xml(
            br#"<p:sld><p:cSld><p:spTree><p:sp><p:nvSpPr><p:cNvPr id="2" name="Bare"/><p:nvPr/></p:nvSpPr><p:spPr/><p:style><a:lnRef idx="2"><a:schemeClr val="accent1"/></a:lnRef><a:fillRef idx="1"><a:schemeClr val="accent1"/></a:fillRef><a:effectRef idx="0"><a:schemeClr val="accent1"/></a:effectRef><a:fontRef idx="minor"/></p:style></p:sp></p:spTree></p:cSld></p:sld>"#,
            "ppt/slides/slide1.xml",
            &mut budget,
        )
        .unwrap();
        let data = common_slide_data(
            &root,
            &[],
            "ppt/slides/slide1.xml",
            &mut budget,
            ShapeElements::WithConnectors,
        )
        .unwrap();

        let ShapeNode::Shape(shape) = &data.shapes[0] else {
            panic!("expected a shape");
        };
        let style = shape.style.as_ref().unwrap();
        assert!(style.font_color.is_none());
        assert_eq!(style.fill.as_ref().unwrap().index, 1);
        assert_eq!(style.line.as_ref().unwrap().index, 2);
    }

    #[test]
    fn a_connector_parses_as_a_shape_and_keeps_its_place_in_the_tree() {
        let limits = ParseLimits::default();
        let mut budget = ParseBudget::new(&limits);
        let root = parse_xml(
            br#"<p:sld><p:cSld><p:spTree><p:sp><p:nvSpPr><p:cNvPr id="2" name="First"/><p:nvPr/></p:nvSpPr><p:spPr/></p:sp><p:cxnSp><p:nvCxnSpPr><p:cNvPr id="3" name="Straight Connector 2"/><p:cNvCxnSpPr/><p:nvPr/></p:nvCxnSpPr><p:spPr><a:xfrm><a:off x="100" y="200"/><a:ext cx="0" cy="500"/></a:xfrm><a:prstGeom prst="line"><a:avLst/></a:prstGeom><a:ln w="19050"><a:solidFill><a:srgbClr val="FF0000"/></a:solidFill></a:ln></p:spPr></p:cxnSp><p:sp><p:nvSpPr><p:cNvPr id="4" name="Third"/><p:nvPr/></p:nvSpPr><p:spPr/></p:sp></p:spTree></p:cSld></p:sld>"#,
            "ppt/slides/slide1.xml",
            &mut budget,
        )
        .unwrap();
        let data = common_slide_data(
            &root,
            &[],
            "ppt/slides/slide1.xml",
            &mut budget,
            ShapeElements::WithConnectors,
        )
        .unwrap();

        assert_eq!(data.shapes.len(), 3);
        let ShapeNode::Shape(connector) = &data.shapes[1] else {
            panic!("expected the connector to parse as a shape");
        };
        assert_eq!(connector.base.id, 3);
        assert_eq!(connector.base.name, "Straight Connector 2");
        assert_eq!(connector.geometry, "line");
        assert_eq!(connector.base.transform.width, 0);
        assert_eq!(connector.base.transform.height, 500);
        let outline = connector
            .outline
            .as_ref()
            .expect("connector has an outline");
        assert_eq!(
            outline.color.as_ref().unwrap().rgb.as_deref(),
            Some("FF0000")
        );
    }

    #[test]
    fn group_fill_reaches_every_shape_that_asks_for_it() {
        let limits = ParseLimits::default();
        let mut budget = ParseBudget::new(&limits);
        let root = parse_xml(
            br#"<p:sld><p:cSld><p:spTree><p:grpSp><p:nvGrpSpPr><p:cNvPr id="10" name="Outer"/></p:nvGrpSpPr><p:grpSpPr><a:solidFill><a:schemeClr val="accent1"/></a:solidFill></p:grpSpPr><p:sp><p:nvSpPr><p:cNvPr id="11" name="Inherits"/><p:nvPr/></p:nvSpPr><p:spPr><a:grpFill/></p:spPr></p:sp><p:sp><p:nvSpPr><p:cNvPr id="12" name="Explicit none"/><p:nvPr/></p:nvSpPr><p:spPr><a:noFill/></p:spPr></p:sp><p:grpSp><p:nvGrpSpPr><p:cNvPr id="13" name="Middle"/></p:nvGrpSpPr><p:grpSpPr><a:grpFill/></p:grpSpPr><p:sp><p:nvSpPr><p:cNvPr id="14" name="Two levels up"/><p:nvPr/></p:nvSpPr><p:spPr><a:grpFill/></p:spPr></p:sp></p:grpSp><p:grpSp><p:nvGrpSpPr><p:cNvPr id="15" name="Own fill"/></p:nvGrpSpPr><p:grpSpPr><a:solidFill><a:srgbClr val="FF0000"/></a:solidFill></p:grpSpPr><p:sp><p:nvSpPr><p:cNvPr id="16" name="Nearest wins"/><p:nvPr/></p:nvSpPr><p:spPr><a:grpFill/></p:spPr></p:sp></p:grpSp></p:grpSp></p:spTree></p:cSld></p:sld>"#,
            "ppt/slides/slide1.xml",
            &mut budget,
        )
        .unwrap();
        let data = common_slide_data(
            &root,
            &[],
            "ppt/slides/slide1.xml",
            &mut budget,
            ShapeElements::WithConnectors,
        )
        .unwrap();
        let ShapeNode::Group(outer) = &data.shapes[0] else {
            panic!("expected a group");
        };

        let fill_of = |node: &ShapeNode| match node {
            ShapeNode::Shape(shape) => shape.fill.clone(),
            _ => panic!("expected a shape"),
        };

        // A direct child takes the group's own fill.
        let inherited = fill_of(&outer.children[0]).expect("the child should have a fill");
        assert_eq!(inherited.fill_type, "solid");
        assert_eq!(
            inherited.color.as_ref().unwrap().theme_color.as_deref(),
            Some("accent1")
        );

        // An explicit noFill is left alone.
        assert_eq!(fill_of(&outer.children[1]).unwrap().fill_type, "none");

        // A group that inherits passes the fill down to its own children.
        let ShapeNode::Group(middle) = &outer.children[2] else {
            panic!("expected a nested group");
        };
        assert_eq!(fill_of(&middle.children[0]), Some(inherited));

        // The nearest ancestor that declares a fill wins.
        let ShapeNode::Group(own) = &outer.children[3] else {
            panic!("expected a nested group");
        };
        let nearest = fill_of(&own.children[0]).expect("the child should have a fill");
        assert_eq!(
            nearest.color.as_ref().unwrap().rgb.as_deref(),
            Some("FF0000")
        );
    }

    #[test]
    fn a_group_fill_with_no_group_to_inherit_from_is_no_fill() {
        let limits = ParseLimits::default();
        let mut budget = ParseBudget::new(&limits);
        let root = parse_xml(
            br#"<p:sld><p:cSld><p:bg><p:bgPr><a:grpFill/></p:bgPr></p:bg><p:spTree><p:sp><p:nvSpPr><p:cNvPr id="2" name="Title"/><p:nvPr><p:ph type="title"/></p:nvPr></p:nvSpPr><p:spPr><a:grpFill/></p:spPr></p:sp><p:grpSp><p:nvGrpSpPr><p:cNvPr id="3" name="Unfilled"/></p:nvGrpSpPr><p:grpSpPr><a:grpFill/></p:grpSpPr><p:pic><p:nvPicPr><p:cNvPr id="4" name="Nested"/></p:nvPicPr><p:spPr><a:grpFill/></p:spPr></p:pic></p:grpSp></p:spTree></p:cSld></p:sld>"#,
            "ppt/slides/slide1.xml",
            &mut budget,
        )
        .unwrap();
        let data = common_slide_data(
            &root,
            &[],
            "ppt/slides/slide1.xml",
            &mut budget,
            ShapeElements::WithConnectors,
        )
        .unwrap();

        assert_eq!(data.background, None);
        let ShapeNode::Shape(placeholder) = &data.shapes[0] else {
            panic!("expected a shape");
        };
        assert_eq!(placeholder.fill, None);
        let ShapeNode::Group(group) = &data.shapes[1] else {
            panic!("expected a group");
        };
        let ShapeNode::Picture(nested) = &group.children[0] else {
            panic!("expected a picture");
        };
        assert_eq!(nested.fill, None);
    }

    #[test]
    fn parses_text_formatting_and_nested_shape_types() {
        let limits = ParseLimits::default();
        let mut budget = ParseBudget::new(&limits);
        let root = parse_xml(
            br#"<p:sld><p:cSld name="Test"><p:spTree><p:sp><p:nvSpPr><p:cNvPr id="2" name="Title"/><p:nvPr><p:ph type="title"/></p:nvPr></p:nvSpPr><p:spPr><a:xfrm><a:off x="1" y="2"/><a:ext cx="3" cy="4"/></a:xfrm><a:prstGeom prst="roundRect"><a:avLst><a:gd name="adj" fmla="val 20000"/></a:avLst></a:prstGeom><a:solidFill><a:schemeClr val="accent1"/></a:solidFill></p:spPr><p:txBody><a:bodyPr anchor="ctr"><a:normAutofit fontScale="85000" lnSpcReduction="12000"/></a:bodyPr><a:p><a:pPr algn="ctr"/><a:r><a:rPr sz="2400" b="1"><a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill><a:latin typeface="Aptos"/></a:rPr><a:t>Hello</a:t></a:r></a:p></p:txBody></p:sp></p:spTree></p:cSld></p:sld>"#,
            "ppt/slides/slide1.xml",
            &mut budget,
        )
        .unwrap();
        let data = common_slide_data(
            &root,
            &[],
            "ppt/slides/slide1.xml",
            &mut budget,
            ShapeElements::WithConnectors,
        )
        .unwrap();
        let ShapeNode::Shape(shape) = &data.shapes[0] else {
            panic!("expected shape");
        };
        assert_eq!(shape.geometry, "roundRect");
        assert_eq!(shape.adjust_values.get("adj"), Some(&0.2));
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

    #[test]
    fn reads_a_picture_source_crop_and_its_own_geometry() {
        let limits = ParseLimits::default();
        let mut budget = ParseBudget::new(&limits);
        let root = parse_xml(
            br#"<p:sld><p:cSld><p:spTree><p:pic><p:nvPicPr><p:cNvPr id="7" name="Screenshot"/><p:nvPr/></p:nvPicPr><p:blipFill><a:blip r:embed="rId2"/><a:srcRect t="251" b="16720"/></p:blipFill><p:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="400" cy="200"/></a:xfrm><a:prstGeom prst="ellipse"><a:avLst/></a:prstGeom></p:spPr></p:pic></p:spTree></p:cSld></p:sld>"#,
            "ppt/slides/slide1.xml",
            &mut budget,
        )
        .unwrap();
        let data = common_slide_data(
            &root,
            &[],
            "ppt/slides/slide1.xml",
            &mut budget,
            ShapeElements::WithConnectors,
        )
        .unwrap();
        let ShapeNode::Picture(picture) = &data.shapes[0] else {
            panic!("expected picture");
        };
        assert_eq!(picture.crop.top, 251);
        assert_eq!(picture.crop.bottom, 16_720);
        assert_eq!(picture.crop.left, 0);
        assert_eq!(picture.geometry, "ellipse");
    }

    #[test]
    fn a_picture_without_a_preset_geometry_reads_as_a_rectangle() {
        let limits = ParseLimits::default();
        let mut budget = ParseBudget::new(&limits);
        let root = parse_xml(
            br#"<p:sld><p:cSld><p:spTree><p:pic><p:nvPicPr><p:cNvPr id="8" name="Photo"/><p:nvPr/></p:nvPicPr><p:blipFill><a:blip r:embed="rId3"/></p:blipFill><p:spPr/></p:pic></p:spTree></p:cSld></p:sld>"#,
            "ppt/slides/slide1.xml",
            &mut budget,
        )
        .unwrap();
        let data = common_slide_data(
            &root,
            &[],
            "ppt/slides/slide1.xml",
            &mut budget,
            ShapeElements::WithConnectors,
        )
        .unwrap();
        let ShapeNode::Picture(picture) = &data.shapes[0] else {
            panic!("expected picture");
        };
        assert_eq!(picture.crop, PictureCrop::default());
        assert_eq!(picture.geometry, "rect");
    }

    #[test]
    fn a_rectangular_picture_serializes_as_it_did_before_pictures_had_a_geometry() {
        let limits = ParseLimits::default();
        let mut budget = ParseBudget::new(&limits);
        let root = parse_xml(
            br#"<p:sld><p:cSld><p:spTree><p:pic><p:nvPicPr><p:cNvPr id="8" name="Photo"/><p:nvPr/></p:nvPicPr><p:blipFill><a:blip r:embed="rId3"/></p:blipFill><p:spPr><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr></p:pic><p:pic><p:nvPicPr><p:cNvPr id="9" name="Portrait"/><p:nvPr/></p:nvPicPr><p:blipFill><a:blip r:embed="rId3"/></p:blipFill><p:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="400" cy="200"/></a:xfrm><a:prstGeom prst="roundRect"><a:avLst><a:gd name="adj" fmla="val 25000"/></a:avLst></a:prstGeom></p:spPr></p:pic></p:spTree></p:cSld></p:sld>"#,
            "ppt/slides/slide1.xml",
            &mut budget,
        )
        .unwrap();
        let data = common_slide_data(
            &root,
            &[],
            "ppt/slides/slide1.xml",
            &mut budget,
            ShapeElements::WithConnectors,
        )
        .unwrap();
        let (ShapeNode::Picture(rectangle), ShapeNode::Picture(rounded)) =
            (&data.shapes[0], &data.shapes[1])
        else {
            panic!("expected two pictures");
        };
        let json = serde_json::to_value(rectangle).unwrap();
        assert!(json.get("geometry").is_none());
        assert!(json.get("adjustValues").is_none());
        assert_eq!(serde_json::from_value::<Picture>(json).unwrap(), *rectangle);
        let json = serde_json::to_value(rounded).unwrap();
        assert_eq!(json["geometry"], "roundRect");
        assert_eq!(json["adjustValues"]["adj"], 0.25);
    }

    #[test]
    fn evaluates_literal_arithmetic_and_reference_adjustment_formulas() {
        let properties = adjustment_properties(
            r#"<a:gd name="adj" fmla="val 20000"/>
               <a:gd name="adj1" fmla="*/ 30000 2 3"/>
               <a:gd name="adj2" fmla="+- adj1 15000 5000"/>
               <a:gd name="adj3" fmla="+/ adj2 10000 2"/>"#,
        );
        let adjustments = parse_adjust_values(Some(&properties), None);

        assert_eq!(adjustments.get("adj"), Some(&0.2));
        assert_eq!(adjustments.get("adj1"), Some(&0.2));
        assert_eq!(adjustments.get("adj2"), Some(&0.3));
        assert_eq!(adjustments.get("adj3"), Some(&0.2));
    }

    #[test]
    fn seeds_standard_geometry_guides_from_extent() {
        let values = standard_guide_values((4_000_000.0, 1_000_000.0));

        for (name, expected) in [
            ("w", 4_000_000.0),
            ("h", 1_000_000.0),
            ("ss", 1_000_000.0),
            ("ls", 4_000_000.0),
            ("hc", 2_000_000.0),
            ("vc", 500_000.0),
            ("l", 0.0),
            ("t", 0.0),
            ("r", 4_000_000.0),
            ("b", 1_000_000.0),
        ] {
            assert_guide_value(&values, name, expected, 1.0);
        }
        for divisor in [2, 3, 4, 5, 6, 8, 10, 12, 32] {
            assert_guide_value(
                &values,
                &format!("wd{divisor}"),
                4_000_000.0 / divisor as f64,
                1.0,
            );
        }
        for divisor in [2, 3, 4, 5, 6, 8] {
            assert_guide_value(
                &values,
                &format!("hd{divisor}"),
                1_000_000.0 / divisor as f64,
                1.0,
            );
        }
        for divisor in [2, 4, 6, 8, 16, 32] {
            assert_guide_value(
                &values,
                &format!("ssd{divisor}"),
                1_000_000.0 / divisor as f64,
                1.0,
            );
        }
        for (name, expected) in [
            ("cd2", 10_800_000.0),
            ("cd4", 5_400_000.0),
            ("cd8", 2_700_000.0),
            ("3cd4", 16_200_000.0),
            ("3cd8", 8_100_000.0),
            ("5cd8", 13_500_000.0),
            ("7cd8", 18_900_000.0),
        ] {
            assert_guide_value(&values, name, expected, 0.0);
        }
    }

    #[test]
    fn evaluates_standard_and_mixed_guide_formulas() {
        let properties = adjustment_properties(
            r#"<a:gd name="fromW" fmla="*/ w 1 16"/>
               <a:gd name="fromH" fmla="*/ h 1 4"/>
               <a:gd name="fromSs" fmla="*/ ss 1 4"/>
               <a:gd name="fromLs" fmla="*/ ls 1 16"/>
               <a:gd name="fromHc" fmla="*/ hc 1 8"/>
               <a:gd name="fromVc" fmla="*/ vc 1 2"/>
               <a:gd name="fromL" fmla="+- l 250000 0"/>
               <a:gd name="fromT" fmla="+- t 250000 0"/>
               <a:gd name="fromR" fmla="*/ r 1 16"/>
               <a:gd name="fromB" fmla="*/ b 1 4"/>
               <a:gd name="fromWd2" fmla="*/ wd2 1 8"/>
               <a:gd name="fromHd2" fmla="*/ hd2 1 2"/>
               <a:gd name="fromSsd4" fmla="val ssd4"/>
               <a:gd name="fromAngle" fmla="*/ cd4 1 54"/>
               <a:gd name="mixed" fmla="+- w 10000 h"/>
               <a:gd name="cancelled" fmla="*/ w 10000 w"/>"#,
        );
        let adjustments = parse_adjust_values(Some(&properties), Some((4_000_000.0, 1_000_000.0)));

        for name in [
            "fromW", "fromH", "fromSs", "fromLs", "fromHc", "fromVc", "fromL", "fromT", "fromR",
            "fromB", "fromWd2", "fromHd2", "fromSsd4",
        ] {
            assert_eq!(adjustments.get(name), Some(&0.25));
        }
        assert_eq!(adjustments.get("fromAngle"), Some(&1.0));
        assert_eq!(adjustments.get("mixed"), Some(&3.01));
        assert_eq!(adjustments.get("cancelled"), Some(&0.1));
    }

    #[test]
    fn extent_relative_adjustment_matches_equivalent_literal() {
        let extent = Some((4_000_000.0, 1_000_000.0));
        let relative = adjustment_properties(r#"<a:gd name="adj" fmla="*/ ss 25000 100000"/>"#);
        let literal = adjustment_properties(r#"<a:gd name="adj" fmla="val 25000"/>"#);

        assert_eq!(
            parse_adjust_values(Some(&relative), extent),
            parse_adjust_values(Some(&literal), extent)
        );
        assert_eq!(
            parse_adjust_values(Some(&relative), extent).get("adj"),
            Some(&0.25)
        );
    }

    #[test]
    fn leaves_extent_guides_absent_without_an_extent() {
        let properties = adjustment_properties(
            r#"<a:gd name="adj" fmla="*/ ss 25000 100000"/>
               <a:gd name="adj1" fmla="val 20000"/>"#,
        );
        let adjustments = parse_adjust_values(Some(&properties), None);

        assert!(!adjustments.contains_key("adj"));
        assert_eq!(adjustments.get("adj1"), Some(&0.2));
    }

    #[test]
    fn selection_operators_keep_the_chosen_operand_units() {
        let extent = Some((4_000_000.0, 1_000_000.0));
        let properties = adjustment_properties(
            r#"<a:gd name="adj" fmla="?: 1 50000 w"/>
               <a:gd name="adj1" fmla="max 25000 ssd4"/>
               <a:gd name="adj2" fmla="pin 10000 ss 30000"/>"#,
        );
        let adjustments = parse_adjust_values(Some(&properties), extent);

        assert_eq!(adjustments.get("adj"), Some(&0.5));
        assert_eq!(adjustments.get("adj1"), Some(&0.25));
        assert_eq!(adjustments.get("adj2"), Some(&0.3));
    }

    #[test]
    fn leaves_unresolvable_adjustments_absent_for_preset_fallbacks() {
        let properties = adjustment_properties(
            r#"<a:gd name="adj" fmla="*/ missing 2 3"/>
               <a:gd name="adj1" fmla="unknown 20000"/>"#,
        );

        assert!(parse_adjust_values(Some(&properties), None).is_empty());
    }

    fn adjustment_properties(guides: &str) -> XmlElement {
        let limits = ParseLimits::default();
        let mut budget = ParseBudget::new(&limits);
        parse_xml(
            format!(
                r#"<p:spPr><a:prstGeom prst="roundRect"><a:avLst>{guides}</a:avLst></a:prstGeom></p:spPr>"#
            )
            .as_bytes(),
            "ppt/slides/slide1.xml",
            &mut budget,
        )
        .unwrap()
    }

    fn assert_guide_value(
        values: &BTreeMap<String, GuideValue>,
        name: &str,
        expected: f64,
        expected_power: f64,
    ) {
        let actual = values.get(name).unwrap();
        assert_eq!(actual.value, expected);
        assert_eq!(actual.extent_power, expected_power);
    }
}
