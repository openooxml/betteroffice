use std::f64::consts::PI;

use docx_layout::types::{
    BlockId, LineBreakRun, ParagraphAttrs, ParagraphBlock, Run, RunFormatting, ShapeBlock, TabRun,
    TextRun,
};
use serde_json::{Map, Value, json};

use super::RenderEnv;

const EMU_PER_INCH: f64 = 914_400.0;
const PIXELS_PER_INCH: f64 = 96.0;

pub(super) fn lower_shape_json(
    shape: &Value,
    pm_start: u64,
    env: &RenderEnv,
) -> Option<ShapeBlock> {
    lower_shape(shape, format!("shape:{pm_start}"), env, Some(pm_start))
}

fn lower_shape(
    shape: &Value,
    block_id: String,
    env: &RenderEnv,
    pm_start: Option<u64>,
) -> Option<ShapeBlock> {
    let shape_type = string(shape, "shapeType")
        .filter(|value| !value.is_empty())
        .unwrap_or_else(|| "rect".to_owned());
    let geometry_path = array(shape, "geometryPath")
        .filter(|path| !path.is_empty())
        .cloned()
        .or_else(|| preset_geometry(&shape_type))?;
    let size = object(shape, "size");
    let width = size
        .and_then(|value| number_in(value, "width"))
        .filter(|value| *value != 0.0)
        .map(emu_to_pixels)
        .unwrap_or(100.0);
    let height = size
        .and_then(|value| number_in(value, "height"))
        .filter(|value| *value != 0.0)
        .map(emu_to_pixels)
        .unwrap_or(80.0);
    let (width, height) = constrain_to_page(width, height, env);
    let offset = object(shape, "offset");
    let children = array(shape, "children")
        .into_iter()
        .flatten()
        .enumerate()
        .filter_map(|(index, child)| {
            lower_shape(child, format!("{block_id}:child:{index}"), env, None)
        })
        .collect();
    let (doc_start, doc_end) = pm_start
        .map(|start| (Some(start as f64), Some((start + 1) as f64)))
        .unwrap_or((None, None));
    let inner_text = shape_inner_text(shape, &block_id);

    Some(ShapeBlock {
        sdt_groups: None,
        id: BlockId::Str(block_id),
        shape_type,
        geometry_path,
        fill: shape_fill(shape),
        stroke: shape_stroke(shape),
        transform: shape_transform(shape),
        width,
        height,
        x: offset
            .and_then(|value| number_in(value, "x"))
            .map(emu_to_pixels),
        y: offset
            .and_then(|value| number_in(value, "y"))
            .map(emu_to_pixels),
        inner_text,
        inner_measures: None,
        children,
        scene: field(shape, "scene").cloned(),
        effects: array(shape, "effects").cloned(),
        text_body_properties: field(shape, "textBodyProperties").cloned(),
        position: None,
        wrap_type: None,
        wrap_text: None,
        relative_height: None,
        behind_doc: None,
        decorative: None,
        title: None,
        description: None,
        doc_start,
        doc_end,
        pm_start: doc_start,
        pm_end: doc_end,
    })
}

fn shape_fill(shape: &Value) -> Option<Value> {
    if let Some(paint) = object(shape, "fillPaint")
        && let Some(kind) = string_in(paint, "kind")
    {
        match kind.as_str() {
            "pattern" => {
                let mut fill = Map::new();
                fill.insert("type".to_owned(), Value::String("pattern".to_owned()));
                insert_color(
                    &mut fill,
                    "color",
                    paint
                        .get("foregroundColor")
                        .filter(|value| !value.is_null())
                        .or_else(|| paint.get("color")),
                );
                insert_clone(&mut fill, "patternPreset", paint.get("patternPreset"));
                insert_color(&mut fill, "foregroundColor", paint.get("foregroundColor"));
                insert_color(&mut fill, "backgroundColor", paint.get("backgroundColor"));
                return Some(Value::Object(fill));
            }
            "picture" => {
                let mut fill = Map::new();
                fill.insert("type".to_owned(), Value::String("picture".to_owned()));
                insert_color(&mut fill, "color", paint.get("color"));
                if let Some(picture) = paint.get("picture") {
                    insert_clone(
                        &mut fill,
                        "pictureRelId",
                        object_value(picture).and_then(|value| value.get("rId")),
                    );
                    if let Some(src) = object_value(picture)
                        .and_then(|value| value.get("src"))
                        .and_then(Value::as_str)
                        .filter(|value| value.starts_with("data:") || value.starts_with("blob:"))
                    {
                        fill.insert("pictureSrc".to_owned(), Value::String(src.to_owned()));
                    }
                }
                for (source, target) in [
                    ("srcRect", "pictureSrcRect"),
                    ("fillMode", "pictureFillMode"),
                    ("tile", "pictureTile"),
                    ("stretchRect", "pictureStretchRect"),
                    ("pictureOpacity", "pictureOpacity"),
                ] {
                    insert_clone(&mut fill, target, paint.get(source));
                }
                return Some(Value::Object(fill));
            }
            "theme" => {
                let mut fill = Map::new();
                fill.insert("type".to_owned(), Value::String("solid".to_owned()));
                insert_color(&mut fill, "color", paint.get("color"));
                insert_clone(&mut fill, "themeRefIndex", paint.get("themeRefIndex"));
                return Some(Value::Object(fill));
            }
            _ => {}
        }
    }

    let fill = object(shape, "fill")?;
    let fill_type = string_in(fill, "type")?;
    let mut output = Map::new();
    output.insert("type".to_owned(), Value::String(fill_type.clone()));
    if fill_type == "none" {
        return Some(Value::Object(output));
    }
    insert_color(&mut output, "color", fill.get("color"));
    if fill_type == "gradient"
        && let Some(gradient) = fill.get("gradient").and_then(object_value)
    {
        insert_clone(&mut output, "gradientType", gradient.get("type"));
        insert_clone(&mut output, "gradientAngle", gradient.get("angle"));
        if let Some(stops) = gradient.get("stops").and_then(Value::as_array) {
            output.insert(
                "gradientStops".to_owned(),
                Value::Array(
                    stops
                        .iter()
                        .map(|stop| {
                            let mut output = Map::new();
                            if let Some(stop) = object_value(stop) {
                                insert_clone(&mut output, "position", stop.get("position"));
                                output.insert(
                                    "color".to_owned(),
                                    resolve_shape_color(stop.get("color"))
                                        .map(Value::String)
                                        .unwrap_or_else(|| Value::String("#000000".to_owned())),
                                );
                            }
                            Value::Object(output)
                        })
                        .collect(),
                ),
            );
        }
    }
    Some(Value::Object(output))
}

fn shape_stroke(shape: &Value) -> Option<Value> {
    let outline = object(shape, "outline")?;
    let mut stroke = Map::new();
    insert_color(&mut stroke, "color", outline.get("color"));
    if let Some(width) = number_in(outline, "width") {
        stroke.insert("width".to_owned(), Value::from(emu_to_pixels(width)));
    }
    insert_clone(&mut stroke, "dash", outline.get("style"));
    Some(Value::Object(stroke))
}

fn shape_transform(shape: &Value) -> Option<Value> {
    let transform = object(shape, "transform")?;
    let mut output = Map::new();
    insert_clone(&mut output, "rotation", transform.get("rotation"));
    if bool_in(transform, "flipH") == Some(true) {
        output.insert("flipH".to_owned(), Value::Bool(true));
    }
    if bool_in(transform, "flipV") == Some(true) {
        output.insert("flipV".to_owned(), Value::Bool(true));
    }
    (!output.is_empty()).then_some(Value::Object(output))
}

fn shape_inner_text(shape: &Value, block_id: &str) -> Option<Vec<ParagraphBlock>> {
    let text_body = object(shape, "textBody")?;
    let content = text_body.get("content")?.as_array()?;
    Some(
        content
            .iter()
            .enumerate()
            .map(|(index, paragraph)| shape_paragraph(paragraph, format!("{block_id}:p:{index}")))
            .collect(),
    )
}

fn shape_paragraph(paragraph: &Value, block_id: String) -> ParagraphBlock {
    let alignment = object(paragraph, "formatting")
        .and_then(|value| string_in(value, "alignment"))
        .and_then(|value| match value.as_str() {
            "both" => Some("justify".to_owned()),
            "left" | "center" | "right" => Some(value),
            _ => None,
        });
    let runs = array(paragraph, "content")
        .into_iter()
        .flatten()
        .flat_map(shape_paragraph_content_runs)
        .collect();
    ParagraphBlock {
        sdt_groups: None,
        id: BlockId::Str(block_id),
        para_id: string(paragraph, "paraId"),
        runs,
        attrs: Some(ParagraphAttrs {
            alignment,
            ..ParagraphAttrs::default()
        }),
        pm_start: None,
        pm_end: None,
    }
}

fn shape_paragraph_content_runs(content: &Value) -> Vec<Run> {
    match string(content, "type").as_deref() {
        Some("run") => shape_document_runs(content),
        Some("hyperlink") => array(content, "children")
            .into_iter()
            .flatten()
            .filter(|child| string(child, "type").as_deref() == Some("run"))
            .flat_map(shape_document_runs)
            .collect(),
        _ => Vec::new(),
    }
}

fn shape_document_runs(run: &Value) -> Vec<Run> {
    let formatting = shape_run_formatting(field(run, "formatting"));
    array(run, "content")
        .into_iter()
        .flatten()
        .filter_map(|content| match string(content, "type").as_deref() {
            Some("text") => string(content, "text")
                .filter(|text| !text.is_empty())
                .map(|text| {
                    Run::Text(TextRun {
                        fmt: formatting.clone(),
                        text,
                        pm_start: None,
                        pm_end: None,
                        inline_sdt_widget: None,
                    })
                }),
            Some("tab") => Some(Run::Tab(TabRun {
                fmt: formatting.clone(),
                pm_start: None,
                pm_end: None,
                width: None,
                leader_glyphs: None,
            })),
            Some("break") => Some(Run::LineBreak(LineBreakRun {
                pm_start: None,
                pm_end: None,
            })),
            _ => None,
        })
        .collect()
}

fn shape_run_formatting(source: Option<&Value>) -> RunFormatting {
    let Some(source) = source.and_then(object_value) else {
        return RunFormatting::default();
    };
    let mut output = Map::new();
    for key in ["bold", "italic", "strike"] {
        if bool_in(source, key) == Some(true) {
            output.insert(key.to_owned(), Value::Bool(true));
        }
    }
    for (source_key, target_key) in [
        ("boldCs", "boldCs"),
        ("italicCs", "italicCs"),
        ("cs", "complexScript"),
    ] {
        insert_clone(&mut output, target_key, source.get(source_key));
    }
    if source.get("underline").is_some_and(js_truthy) {
        output.insert("underline".to_owned(), Value::Bool(true));
    }
    if let Some(color) = resolve_shape_color(source.get("color")) {
        output.insert("color".to_owned(), Value::String(color));
    }
    if let Some(size) = number_in(source, "fontSize") {
        output.insert("fontSize".to_owned(), Value::from(size / 2.0));
    }
    if let Some(size) = number_in(source, "fontSizeCs") {
        output.insert("fontSizeCs".to_owned(), Value::from(size / 2.0));
    }
    if let Some(fonts) = source.get("fontFamily").and_then(object_value) {
        output.insert("fontSlots".to_owned(), Value::Object(fonts.clone()));
        if let Some(family) = ["ascii", "hAnsi", "eastAsia", "cs"]
            .iter()
            .find_map(|key| fonts.get(*key).and_then(Value::as_str))
            .filter(|value| !value.is_empty())
        {
            output.insert("fontFamily".to_owned(), Value::String(family.to_owned()));
        }
    }
    insert_clone(&mut output, "language", source.get("language"));
    serde_json::from_value(Value::Object(output)).unwrap_or_default()
}

fn resolve_shape_color(value: Option<&Value>) -> Option<String> {
    let value = value?.as_object()?;
    if let Some(rgb) = value
        .get("rgb")
        .and_then(Value::as_str)
        .filter(|value| !value.is_empty())
    {
        return Some(format!("#{rgb}"));
    }
    let slot = value
        .get("themeColor")
        .and_then(Value::as_str)
        .filter(|value| !value.is_empty())?;
    let color = match slot.to_ascii_lowercase().as_str() {
        "dk1" | "dark1" | "text1" | "tx1" => "000000",
        "lt1" | "light1" | "background1" | "bg1" => "FFFFFF",
        "dk2" | "dark2" | "text2" | "tx2" => "44546A",
        "lt2" | "light2" | "background2" | "bg2" => "E7E6E6",
        "accent1" => "4472C4",
        "accent2" => "ED7D31",
        "accent3" => "A5A5A5",
        "accent4" => "FFC000",
        "accent5" => "5B9BD5",
        "accent6" => "70AD47",
        "hlink" | "hyperlink" => "0563C1",
        "folhlink" | "followedhyperlink" => "954F72",
        _ => return None,
    };
    Some(format!("#{color}"))
}

fn insert_color(output: &mut Map<String, Value>, key: &str, value: Option<&Value>) {
    if let Some(color) = resolve_shape_color(value) {
        output.insert(key.to_owned(), Value::String(color));
    }
}

fn insert_clone(output: &mut Map<String, Value>, key: &str, value: Option<&Value>) {
    if let Some(value) = value.filter(|value| !value.is_null()) {
        output.insert(key.to_owned(), value.clone());
    }
}

fn field<'a>(value: &'a Value, key: &str) -> Option<&'a Value> {
    value.as_object()?.get(key).filter(|value| !value.is_null())
}

fn object<'a>(value: &'a Value, key: &str) -> Option<&'a Map<String, Value>> {
    field(value, key)?.as_object()
}

fn object_value(value: &Value) -> Option<&Map<String, Value>> {
    value.as_object()
}

fn array<'a>(value: &'a Value, key: &str) -> Option<&'a Vec<Value>> {
    field(value, key)?.as_array()
}

fn string(value: &Value, key: &str) -> Option<String> {
    value.as_object().and_then(|value| string_in(value, key))
}

fn string_in(value: &Map<String, Value>, key: &str) -> Option<String> {
    value.get(key)?.as_str().map(str::to_owned)
}

fn number_in(value: &Map<String, Value>, key: &str) -> Option<f64> {
    value.get(key)?.as_f64().filter(|value| value.is_finite())
}

fn bool_in(value: &Map<String, Value>, key: &str) -> Option<bool> {
    value.get(key)?.as_bool()
}

fn js_truthy(value: &Value) -> bool {
    match value {
        Value::Null => false,
        Value::Bool(value) => *value,
        Value::Number(value) => value
            .as_f64()
            .is_some_and(|value| value != 0.0 && !value.is_nan()),
        Value::String(value) => !value.is_empty(),
        Value::Array(_) | Value::Object(_) => true,
    }
}

fn emu_to_pixels(value: f64) -> f64 {
    value / EMU_PER_INCH * PIXELS_PER_INCH
}

fn constrain_to_page(width: f64, height: f64, env: &RenderEnv) -> (f64, f64) {
    let Some(limit) = env.page_content_height.filter(|limit| *limit > 0.0) else {
        return (width, height);
    };
    if height <= limit {
        return (width, height);
    }
    ((width * (limit / height)).round(), limit)
}

fn command(command_type: &str, points: &[(&str, f64)]) -> Value {
    let mut output = Map::new();
    output.insert("type".to_owned(), Value::String(command_type.to_owned()));
    for (key, value) in points {
        output.insert((*key).to_owned(), Value::from(*value));
    }
    Value::Object(output)
}

fn close() -> Value {
    json!({"type": "close"})
}

fn polygon(points: &[(f64, f64)]) -> Vec<Value> {
    let Some((first, rest)) = points.split_first() else {
        return Vec::new();
    };
    let mut output = vec![command("move", &[("x", first.0), ("y", first.1)])];
    output.extend(
        rest.iter()
            .map(|point| command("line", &[("x", point.0), ("y", point.1)])),
    );
    output.push(close());
    output
}

fn regular_polygon(sides: usize) -> Vec<Value> {
    let points: Vec<_> = (0..sides)
        .map(|index| {
            let angle = -PI / 2.0 + index as f64 * PI * 2.0 / sides as f64;
            (0.5 + angle.cos() * 0.5, 0.5 + angle.sin() * 0.5)
        })
        .collect();
    polygon(&points)
}

fn star(points: usize) -> Vec<Value> {
    let vertices: Vec<_> = (0..points * 2)
        .map(|index| {
            let angle = -PI / 2.0 + index as f64 * PI / points as f64;
            let radius = if index % 2 == 0 { 0.5 } else { 0.225 };
            (0.5 + angle.cos() * radius, 0.5 + angle.sin() * radius)
        })
        .collect();
    polygon(&vertices)
}

fn arrow(direction: &str) -> Vec<Value> {
    let mut right = polygon(&[
        (0.0, 0.25),
        (0.5, 0.25),
        (0.5, 0.0),
        (1.0, 0.5),
        (0.5, 1.0),
        (0.5, 0.75),
        (0.0, 0.75),
    ]);
    for command in &mut right {
        let Some(command) = command.as_object_mut() else {
            continue;
        };
        if !matches!(
            command.get("type").and_then(Value::as_str),
            Some("move" | "line")
        ) {
            continue;
        }
        let x = command.get("x").and_then(Value::as_f64).unwrap_or(0.0);
        let y = command.get("y").and_then(Value::as_f64).unwrap_or(0.0);
        let (x, y) = match direction {
            "left" => (1.0 - x, y),
            "up" => (y, 1.0 - x),
            "down" => (y, x),
            _ => (x, y),
        };
        command.insert("x".to_owned(), Value::from(x));
        command.insert("y".to_owned(), Value::from(y));
    }
    right
}

fn bent_connector(segments: usize) -> Vec<Value> {
    if segments <= 2 {
        return vec![
            command("move", &[("x", 0.0), ("y", 0.0)]),
            command("line", &[("x", 0.5), ("y", 0.0)]),
            command("line", &[("x", 0.5), ("y", 1.0)]),
            command("line", &[("x", 1.0), ("y", 1.0)]),
        ];
    }
    let mut output = vec![command("move", &[("x", 0.0), ("y", 0.0)])];
    for index in 1..segments {
        let fraction = index as f64 / segments as f64;
        let point = if index % 2 == 1 {
            (
                if index == 1 { 0.5 } else { fraction },
                (index - 1) as f64 / segments as f64,
            )
        } else {
            ((index - 1) as f64 / segments as f64, fraction)
        };
        output.push(command("line", &[("x", point.0), ("y", point.1)]));
    }
    output.push(command("line", &[("x", 1.0), ("y", 1.0)]));
    output
}

fn curved_connector(segments: usize) -> Vec<Value> {
    if segments <= 2 {
        return vec![
            command("move", &[("x", 0.0), ("y", 0.0)]),
            command(
                "cubic",
                &[
                    ("cp1x", 0.5),
                    ("cp1y", 0.0),
                    ("cp2x", 0.5),
                    ("cp2y", 1.0),
                    ("x", 1.0),
                    ("y", 1.0),
                ],
            ),
        ];
    }
    let mut output = vec![command("move", &[("x", 0.0), ("y", 0.0)])];
    for index in 0..segments - 1 {
        let start = index as f64 / (segments - 1) as f64;
        let end = (index + 1) as f64 / (segments - 1) as f64;
        let middle = start + (end - start) * 0.5;
        output.push(command(
            "cubic",
            &[
                ("cp1x", middle),
                ("cp1y", start),
                ("cp2x", middle),
                ("cp2y", end),
                ("x", end),
                ("y", end),
            ],
        ));
    }
    output
}

fn preset_geometry(shape_type: &str) -> Option<Vec<Value>> {
    Some(match shape_type {
        "rect"
        | "flowChartProcess"
        | "flowChartAlternateProcess"
        | "flowChartPredefinedProcess"
        | "flowChartInternalStorage"
        | "flowChartPreparation"
        | "flowChartManualOperation"
        | "flowChartMagneticTape"
        | "flowChartMagneticDisk"
        | "flowChartMagneticDrum"
        | "flowChartDisplay"
        | "textBox" => polygon(&[(0.0, 0.0), (1.0, 0.0), (1.0, 1.0), (0.0, 1.0)]),
        "roundRect" | "flowChartTerminator" => {
            let radius = 1.0 / 6.0;
            vec![
                command("move", &[("x", radius), ("y", 0.0)]),
                command("line", &[("x", 1.0 - radius), ("y", 0.0)]),
                command(
                    "quad",
                    &[("cpx", 1.0), ("cpy", 0.0), ("x", 1.0), ("y", radius)],
                ),
                command("line", &[("x", 1.0), ("y", 1.0 - radius)]),
                command(
                    "quad",
                    &[("cpx", 1.0), ("cpy", 1.0), ("x", 1.0 - radius), ("y", 1.0)],
                ),
                command("line", &[("x", radius), ("y", 1.0)]),
                command(
                    "quad",
                    &[("cpx", 0.0), ("cpy", 1.0), ("x", 0.0), ("y", 1.0 - radius)],
                ),
                command("line", &[("x", 0.0), ("y", radius)]),
                command(
                    "quad",
                    &[("cpx", 0.0), ("cpy", 0.0), ("x", radius), ("y", 0.0)],
                ),
                close(),
            ]
        }
        "ellipse" | "flowChartConnector" => {
            let kappa = 0.552_284_749_830_793_6 / 2.0;
            vec![
                command("move", &[("x", 1.0), ("y", 0.5)]),
                command(
                    "cubic",
                    &[
                        ("cp1x", 1.0),
                        ("cp1y", 0.5 + kappa),
                        ("cp2x", 0.5 + kappa),
                        ("cp2y", 1.0),
                        ("x", 0.5),
                        ("y", 1.0),
                    ],
                ),
                command(
                    "cubic",
                    &[
                        ("cp1x", 0.5 - kappa),
                        ("cp1y", 1.0),
                        ("cp2x", 0.0),
                        ("cp2y", 0.5 + kappa),
                        ("x", 0.0),
                        ("y", 0.5),
                    ],
                ),
                command(
                    "cubic",
                    &[
                        ("cp1x", 0.0),
                        ("cp1y", 0.5 - kappa),
                        ("cp2x", 0.5 - kappa),
                        ("cp2y", 0.0),
                        ("x", 0.5),
                        ("y", 0.0),
                    ],
                ),
                command(
                    "cubic",
                    &[
                        ("cp1x", 0.5 + kappa),
                        ("cp1y", 0.0),
                        ("cp2x", 1.0),
                        ("cp2y", 0.5 - kappa),
                        ("x", 1.0),
                        ("y", 0.5),
                    ],
                ),
                close(),
            ]
        }
        "line" | "straightConnector1" => vec![
            command("move", &[("x", 0.0), ("y", 0.0)]),
            command("line", &[("x", 1.0), ("y", 1.0)]),
        ],
        "triangle" | "isosTriangle" => polygon(&[(0.5, 0.0), (1.0, 1.0), (0.0, 1.0)]),
        "rtTriangle" => polygon(&[(0.0, 0.0), (1.0, 1.0), (0.0, 1.0)]),
        "diamond" | "flowChartDecision" => {
            polygon(&[(0.5, 0.0), (1.0, 0.5), (0.5, 1.0), (0.0, 0.5)])
        }
        "parallelogram" | "flowChartInputOutput" | "flowChartManualInput" => {
            polygon(&[(0.25, 0.0), (1.0, 0.0), (0.75, 1.0), (0.0, 1.0)])
        }
        "trapezoid" => polygon(&[(0.2, 0.0), (0.8, 0.0), (1.0, 1.0), (0.0, 1.0)]),
        "pentagon" | "flowChartOffpageConnector" => regular_polygon(5),
        "hexagon" => regular_polygon(6),
        "heptagon" => regular_polygon(7),
        "octagon" => regular_polygon(8),
        "decagon" => regular_polygon(10),
        "dodecagon" => regular_polygon(12),
        "star4" => star(4),
        "star5" => star(5),
        "star6" => star(6),
        "star7" => star(7),
        "star8" => star(8),
        "star10" => star(10),
        "star12" => star(12),
        "star16" => star(16),
        "star24" => star(24),
        "star32" => star(32),
        "bentConnector2" => bent_connector(2),
        "bentConnector3" => bent_connector(3),
        "bentConnector4" => bent_connector(4),
        "bentConnector5" => bent_connector(5),
        "curvedConnector2" => curved_connector(2),
        "curvedConnector3" => curved_connector(3),
        "curvedConnector4" => curved_connector(4),
        "curvedConnector5" => curved_connector(5),
        "rightArrow" => arrow("right"),
        "leftArrow" => arrow("left"),
        "upArrow" => arrow("up"),
        "downArrow" => arrow("down"),
        "leftRightArrow" => polygon(&[
            (0.0, 0.5),
            (0.25, 0.0),
            (0.25, 0.25),
            (0.75, 0.25),
            (0.75, 0.0),
            (1.0, 0.5),
            (0.75, 1.0),
            (0.75, 0.75),
            (0.25, 0.75),
            (0.25, 1.0),
        ]),
        "upDownArrow" => polygon(&[
            (0.5, 0.0),
            (1.0, 0.25),
            (0.75, 0.25),
            (0.75, 0.75),
            (1.0, 0.75),
            (0.5, 1.0),
            (0.0, 0.75),
            (0.25, 0.75),
            (0.25, 0.25),
            (0.0, 0.25),
        ]),
        "chevron" => polygon(&[
            (0.0, 0.0),
            (0.65, 0.0),
            (1.0, 0.5),
            (0.65, 1.0),
            (0.0, 1.0),
            (0.35, 0.5),
        ]),
        "homePlate" => polygon(&[(0.0, 0.0), (0.75, 0.0), (1.0, 0.5), (0.75, 1.0), (0.0, 1.0)]),
        _ => return None,
    })
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn lowers_shape_paint_text_children_and_geometry_without_layout_json() {
        let shape = json!({
            "type": "shape",
            "shapeType": "chevron",
            "size": {"width": 914400, "height": 457200},
            "offset": {"x": 91440, "y": 182880},
            "fillPaint": {
                "kind": "picture",
                "picture": {"rId": "rId5", "src": "data:image/png;base64,AA=="},
                "fillMode": "tile",
                "pictureOpacity": 0.5
            },
            "outline": {"color": {"themeColor": "accent2"}, "width": 9525, "style": "dash"},
            "transform": {"rotation": 12, "flipH": true},
            "textBody": {"content": [{
                "paraId": "p1",
                "formatting": {"alignment": "both"},
                "content": [{
                    "type": "run",
                    "formatting": {"bold": true, "fontSize": 24, "color": {"rgb": "112233"}},
                    "content": [{"type": "text", "text": "Shape"}, {"type": "tab"}]
                }]
            }]},
            "children": [{"type": "shape", "shapeType": "ellipse", "size": {"width": 91440, "height": 91440}}],
            "scene": {"version": 1},
            "effects": [{"kind": "glow"}],
            "textBodyProperties": {"anchor": "middle"}
        });

        let block = lower_shape_json(&shape, 7, &RenderEnv::default()).unwrap();

        assert_eq!(block.shape_type, "chevron");
        assert_eq!((block.width, block.height), (96.0, 48.0));
        assert!((block.x.unwrap() - 9.6).abs() < 1e-10);
        assert!((block.y.unwrap() - 19.2).abs() < 1e-10);
        assert_eq!(block.children.len(), 1);
        assert_eq!(
            block.inner_text.as_ref().unwrap()[0]
                .attrs
                .as_ref()
                .unwrap()
                .alignment
                .as_deref(),
            Some("justify")
        );
        assert_eq!(block.inner_text.as_ref().unwrap()[0].runs.len(), 2);
        assert_eq!(block.fill.as_ref().unwrap()["pictureRelId"], "rId5");
        assert_eq!(block.stroke.as_ref().unwrap()["color"], "#ED7D31");
        assert_eq!(block.transform.as_ref().unwrap()["flipH"], true);
        assert_eq!(block.scene.as_ref().unwrap()["version"], 1);
        assert_eq!(block.pm_start, Some(7.0));
    }

    #[test]
    fn rejects_unknown_preset_without_authored_geometry() {
        assert!(
            lower_shape_json(
                &json!({"shapeType": "unsupported", "size": {"width": 1, "height": 1}}),
                0,
                &RenderEnv::default()
            )
            .is_none()
        );
    }
}
