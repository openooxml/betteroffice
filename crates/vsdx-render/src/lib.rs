//! VSDX display-list compilation. Scene data remains in Visio inches until paint.

mod display_list;
mod layout;
mod paint;

pub use display_list::*;
pub use layout::{PIXELS_PER_INCH, final_paint_transform, to_canvas, to_canvas_length};

use std::collections::BTreeMap;

use thiserror::Error;
use vsdx_eval::{Evaluation, PageShapeReferences, Value, evaluate_cell_with_shape_package_theme};
use vsdx_parse::{ParseLimits, Shape, VsdxPackage};
use vsdx_resolve::{Lookup, ResolvedShape, Resolver, realize_geometry};

const MAX_RECURSION_DEPTH: usize = 64;

struct RichParagraph {
    runs: Vec<TextRun>,
    align: i32,
    before: f32,
    after: f32,
    left: f32,
    right: f32,
    first: f32,
    line_spacing: Option<f32>,
}

#[derive(Clone, Copy, Debug, PartialEq)]
struct TabStop {
    position: f32,
    alignment: i32,
    default_interval: Option<f32>,
}

fn rich_paragraphs(
    renderer: &Renderer,
    package: &VsdxPackage,
    references: Option<&PageShapeReferences>,
    resolved: &ResolvedShape,
    shape_id: u32,
    tokens: &[vsdx_resolve::ResolvedTextToken],
) -> Vec<RichParagraph> {
    let mut character = TextRun {
        text: String::new(),
        family: "sans-serif".into(),
        size_in: 1.0 / 6.0,
        bold: false,
        italic: false,
        color: "currentColor".into(),
        underline: false,
        small_caps: false,
        superscript: false,
        subscript: false,
        letter_spacing: 0.0,
        case: 0,
        diagnostics: Vec::new(),
        tab: None,
        diagnosed_face: None,
    };
    let mut paragraphs = vec![RichParagraph {
        runs: Vec::new(),
        align: 0,
        before: 0.0,
        after: 0.0,
        left: 0.0,
        right: 0.0,
        first: 0.0,
        line_spacing: None,
    }];
    character = character_run(
        renderer,
        package,
        references,
        resolved,
        shape_id,
        &row_properties(resolved, "Character", 0),
        &character,
    );
    for token in tokens {
        match token {
            vsdx_resolve::ResolvedTextToken::Literal(value) => {
                append_run(&mut paragraphs, &character, value.clone())
            }
            vsdx_resolve::ResolvedTextToken::CharacterRun { index, properties } => {
                if properties.is_empty() {
                    append_diagnostic(
                        &mut character,
                        "unresolved-character-row",
                        format!("unresolved Character row {index}"),
                    );
                } else {
                    character = character_run(
                        renderer, package, references, resolved, shape_id, properties, &character,
                    );
                }
            }
            vsdx_resolve::ResolvedTextToken::ParagraphRun { properties, .. } => {
                let paragraph = RichParagraph {
                    runs: Vec::new(),
                    align: property_number(properties, "HorzAlign").unwrap_or(0.0) as i32,
                    before: property_number(properties, "SpBefore").unwrap_or(0.0) as f32,
                    after: property_number(properties, "SpAfter").unwrap_or(0.0) as f32,
                    left: property_number(properties, "IndLeft").unwrap_or(0.0) as f32,
                    right: property_number(properties, "IndRight").unwrap_or(0.0) as f32,
                    first: property_number(properties, "IndFirst").unwrap_or(0.0) as f32,
                    line_spacing: property_number(properties, "SpLine").map(|value| value as f32),
                };
                paragraphs.push(paragraph);
            }
            vsdx_resolve::ResolvedTextToken::Tab { properties, .. } => {
                append_run(&mut paragraphs, &character, "\t".into());
                if let Some(run) = paragraphs
                    .last_mut()
                    .and_then(|paragraph| paragraph.runs.last_mut())
                {
                    let position = property_number(properties, "Position");
                    let alignment = property_number(properties, "Alignment").unwrap_or(0.0) as i32;
                    let default_interval =
                        property_number(&resolved.cells, "DefaultTabStop").unwrap_or(0.5) as f32;
                    if position.is_none() {
                        append_diagnostic(
                            run,
                            "missing-tab-position",
                            format!(
                                "tab stop missing Position; used renderer policy DefaultTabStop {default_interval} in"
                            ),
                        );
                    }
                    run.tab = Some(TabStop {
                        position: position.unwrap_or(0.0) as f32,
                        alignment,
                        default_interval: position.is_none().then_some(default_interval),
                    });
                }
            }
            vsdx_resolve::ResolvedTextToken::Field { properties, .. } => {
                let mut run = character.clone();
                match field_value(package, resolved, properties) {
                    Ok(value) => run.text = value,
                    Err(reason) => {
                        run.text = "[unresolved field]".into();
                        append_diagnostic(&mut run, "unresolvable-field", reason);
                    }
                }
                paragraphs
                    .last_mut()
                    .expect("paragraph exists")
                    .runs
                    .push(run);
            }
        }
    }
    paragraphs
}

fn field_value(
    package: &VsdxPackage,
    resolved: &ResolvedShape,
    properties: &std::collections::BTreeMap<String, Lookup>,
) -> Result<String, String> {
    let Some(Lookup::Found(value)) = properties.get("Value") else {
        return Err("unresolvable field: no Value cell".into());
    };
    if let Some(display) = value
        .cell
        .value
        .as_deref()
        .filter(|value| !value.is_empty())
    {
        return Ok(display.into());
    }
    let Some(formula) = value.cell.formula.as_deref() else {
        return Err("unresolvable field: no cached Value".into());
    };
    match evaluate_cell_with_shape_package_theme(
        "Field.Value",
        formula,
        resolved,
        &ParseLimits::default(),
        resolved,
        package,
    ) {
        Evaluation::Evaluated(result) => match result.value {
            Value::Number(value) if value.number.is_finite() => Ok(value.number.to_string()),
            Value::Number(_) => Err("unresolvable field: non-finite Value".into()),
            Value::Color(_) => Err("unresolvable field: Value is a colour".into()),
        },
        Evaluation::Unsupported(reason) => Err(format!("unresolvable field: {reason}")),
        Evaluation::Error(error) => Err(format!("unresolvable field: {}", error.message)),
    }
}

fn append_run(paragraphs: &mut [RichParagraph], character: &TextRun, text: String) {
    let mut run = character.clone();
    run.text = match character_case(character, &text) {
        Ok(text) => text,
        Err(diagnostic) => {
            append_diagnostic(&mut run, "unresolvable-character-case", diagnostic);
            text
        }
    };
    run.tab = None;
    paragraphs
        .last_mut()
        .expect("paragraph exists")
        .runs
        .push(run);
}

fn character_run(
    renderer: &Renderer,
    package: &VsdxPackage,
    references: Option<&PageShapeReferences>,
    resolved: &ResolvedShape,
    shape_id: u32,
    properties: &std::collections::BTreeMap<String, Lookup>,
    current: &TextRun,
) -> TextRun {
    let mut run = current.clone();
    if let Some(font) = property_value(properties, "Font") {
        run.family = package
            .face_names
            .iter()
            .find(|face| {
                face.attributes
                    .iter()
                    .any(|(key, value)| key == "ID" && value == font)
            })
            .and_then(|face| {
                face.attributes
                    .iter()
                    .find(|(key, _)| key == "Name")
                    .map(|(_, value)| value.clone())
            })
            .unwrap_or_else(|| font.into());
    }
    if let Some(size) = property_number(properties, "Size") {
        run.size_in = size as f32;
    }
    if properties.contains_key("Color") {
        match text_colour(package, references, resolved, shape_id, properties) {
            Ok(color) => run.color = color,
            Err(reason) => append_diagnostic(
                &mut run,
                "unresolvable-text-colour",
                format!("unresolvable text colour: {reason}"),
            ),
        }
    }
    if let Some(style) = property_number(properties, "Style") {
        let style = style as i32;
        run.bold = style & 1 != 0;
        run.italic = style & 2 != 0;
        run.underline = style & 4 != 0;
        run.small_caps = style & 8 != 0;
    }
    if let Some(pos) = property_number(properties, "Pos") {
        match pos as i32 {
            0 => (run.superscript, run.subscript) = (false, false),
            1 => (run.superscript, run.subscript) = (true, false),
            2 => (run.superscript, run.subscript) = (false, true),
            _ => append_diagnostic(
                &mut run,
                "unresolvable-character-pos",
                format!("unresolvable Character.Pos value {pos}"),
            ),
        }
    }
    if let Some(case) = property_number(properties, "Case") {
        run.case = case as i32;
    }
    if let Some(letter_spacing) = property_number(properties, "Letterspace") {
        run.letter_spacing = letter_spacing as f32 / 1440.0;
    }
    let exact_font =
        renderer
            .registered_fonts
            .contains_key(&(run.family.clone(), run.bold, run.italic));
    if !exact_font && run.diagnosed_face.as_deref() != Some(run.family.as_str()) {
        let family = run.family.clone();
        let (code, detail) = if renderer.font_for(&family, run.bold, run.italic).is_some() {
            (
                "font-substituted",
                format!("font substituted for '{family}'"),
            )
        } else {
            ("unregistered-font", format!("unregistered font '{family}'"))
        };
        append_diagnostic(&mut run, code, detail);
        run.diagnosed_face = Some(family);
    }
    run
}

fn row_properties(
    resolved: &ResolvedShape,
    section: &str,
    index: u32,
) -> std::collections::BTreeMap<String, Lookup> {
    resolved
        .sections
        .get(section)
        .and_then(|section| section.rows.get(&format!("IX:{index}")))
        .map(|row| row.cells.clone())
        .unwrap_or_default()
}

fn append_diagnostic(run: &mut TextRun, code: &'static str, detail: impl Into<String>) {
    run.diagnostics.push(Diagnostic::for_code(code, detail));
}

fn text_colour(
    package: &VsdxPackage,
    references: Option<&PageShapeReferences>,
    resolved: &ResolvedShape,
    shape_id: u32,
    properties: &std::collections::BTreeMap<String, Lookup>,
) -> Result<String, String> {
    let Some(Lookup::Found(cell)) = properties.get("Color") else {
        return Err("missing Color cell".into());
    };
    let formula = cell
        .cell
        .formula
        .as_deref()
        .or(cell.cell.value.as_deref())
        .ok_or_else(|| "missing Color value".to_string())?;
    let references = references.ok_or_else(|| "unavailable colour references".to_string())?;
    match evaluate_cell_with_shape_package_theme(
        "Char.Color",
        formula,
        &references.for_shape(shape_id),
        &ParseLimits::default(),
        resolved,
        package,
    ) {
        Evaluation::Evaluated(result) => match result.value {
            Value::Color(color) => Ok(format!(
                "#{:02X}{:02X}{:02X}",
                color.red, color.green, color.blue
            )),
            Value::Number(value) => palette_colour(package, value.number)
                .ok_or_else(|| "Color evaluated to an unresolvable palette index".into()),
        },
        Evaluation::Unsupported(reason) => Err(reason),
        Evaluation::Error(error) => Err(error.message),
    }
}

fn palette_colour(package: &VsdxPackage, index: f64) -> Option<String> {
    let index = (index.fract() == 0.0).then_some(index as i64)?;
    let record = package.colors.iter().find(|record| {
        record.attributes.iter().any(|(name, value)| {
            matches!(name.as_str(), "IX" | "Index") && value.parse::<i64>().ok() == Some(index)
        })
    })?;
    let value = record
        .attributes
        .iter()
        .find(|(name, _)| matches!(name.as_str(), "RGB" | "Color" | "Value"))?
        .1
        .trim_start_matches('#');
    (value.len() == 6 && value.chars().all(|ch| ch.is_ascii_hexdigit()))
        .then(|| format!("#{value}"))
}

fn character_case(run: &TextRun, text: &str) -> Result<String, String> {
    match run.case {
        0 => Ok(text.into()),
        1 => Ok(text.to_uppercase()),
        2 => Ok(title_case(text)),
        value => Err(format!("unresolvable Character.Case value {value}")),
    }
}

fn title_case(text: &str) -> String {
    let mut start = true;
    text.chars()
        .flat_map(|ch| {
            let output = if start {
                ch.to_uppercase().collect::<String>()
            } else {
                ch.to_lowercase().collect()
            };
            start = !ch.is_alphanumeric();
            output.chars().collect::<Vec<_>>()
        })
        .collect()
}

fn property_value<'a>(
    properties: &'a std::collections::BTreeMap<String, Lookup>,
    name: &str,
) -> Option<&'a str> {
    match properties.get(name)? {
        Lookup::Found(cell) => cell.cell.value.as_deref(),
        Lookup::Deleted | Lookup::Absent => None,
    }
}

fn property_number(
    properties: &std::collections::BTreeMap<String, Lookup>,
    name: &str,
) -> Option<f64> {
    property_value(properties, name)?
        .parse()
        .ok()
        .filter(|value: &f64| value.is_finite())
}

#[derive(Clone, Debug)]
pub struct RenderLimits {
    pub max_shapes: usize,
    pub max_text_bytes: usize,
    pub max_text_paragraphs: usize,
    pub max_text_lines: usize,
    pub max_text_runs: usize,
    pub max_fonts: usize,
    pub max_font_bytes: usize,
    pub max_display_list_bytes: usize,
}
impl Default for RenderLimits {
    fn default() -> Self {
        Self {
            max_shapes: 100_000,
            max_text_bytes: 16 * 1024 * 1024,
            max_text_paragraphs: 100_000,
            max_text_lines: 1_000_000,
            max_text_runs: 1_000_000,
            max_fonts: 256,
            max_font_bytes: 64 * 1024 * 1024,
            max_display_list_bytes: 64 * 1024 * 1024,
        }
    }
}
#[derive(Debug, Error)]
pub enum RenderError {
    #[error("page not found: {0}")]
    MissingPage(String),
    #[error("render budget exceeded: {0}")]
    Budget(&'static str),
    #[error("invalid font: {0}")]
    Font(String),
    #[error("resolve failed: {0}")]
    Resolve(#[from] vsdx_resolve::ResolveError),
    #[error("invalid page dimensions: {0}")]
    PageDimensions(String),
}

pub struct Renderer {
    limits: RenderLimits,
    fonts: ooxml_text::FontStore,
    font_bytes: usize,
    registered_fonts: BTreeMap<(String, bool, bool), ooxml_text::FontId>,
}
impl Default for Renderer {
    fn default() -> Self {
        Self::new(RenderLimits::default())
    }
}
impl Renderer {
    pub fn new(limits: RenderLimits) -> Self {
        Self {
            limits,
            fonts: ooxml_text::FontStore::new(),
            font_bytes: 0,
            registered_fonts: BTreeMap::new(),
        }
    }
    pub fn register_font(
        &mut self,
        family: impl Into<String>,
        bold: bool,
        italic: bool,
        bytes: Vec<u8>,
    ) -> Result<(), RenderError> {
        let key = (family.into(), bold, italic);
        if !self.registered_fonts.contains_key(&key)
            && self.registered_fonts.len() >= self.limits.max_fonts
        {
            return Err(RenderError::Budget("fonts"));
        }
        let font_bytes = self
            .font_bytes
            .checked_add(bytes.len())
            .ok_or(RenderError::Budget("font bytes"))?;
        if font_bytes > self.limits.max_font_bytes {
            return Err(RenderError::Budget("font bytes"));
        }
        let id = self
            .fonts
            .register(bytes)
            .map_err(|error| RenderError::Font(error.to_string()))?;
        self.font_bytes = font_bytes;
        self.registered_fonts.insert(key, id);
        Ok(())
    }
    pub fn layout_page(
        &self,
        package: &VsdxPackage,
        page_part: &str,
    ) -> Result<VsdxDisplayList, RenderError> {
        let page = package
            .page_contents
            .get(page_part)
            .ok_or_else(|| RenderError::MissingPage(page_part.into()))?;
        let resolver = Resolver::new(package);
        let references = PageShapeReferences::new(&resolver, page_part).ok();
        let page_height = page_dimension(&resolver, package, page_part, "PageHeight")
            .ok_or_else(|| RenderError::PageDimensions("PageHeight is unavailable".into()))?;
        let page_width = page_dimension(&resolver, package, page_part, "PageWidth")
            .ok_or_else(|| RenderError::PageDimensions("PageWidth is unavailable".into()))?;
        if page_height <= 0.0
            || page_width <= 0.0
            || !(page_height as f32).is_finite()
            || !(page_width as f32).is_finite()
            || !(page_height as f32 * PIXELS_PER_INCH).is_finite()
            || !(page_width as f32 * PIXELS_PER_INCH).is_finite()
        {
            return Err(RenderError::PageDimensions(
                "dimensions must be positive finite canvas values".into(),
            ));
        }
        let mut state = State {
            count: 0,
            z_order: 0,
            text_bytes: 0,
            text_paragraphs: 0,
            text_lines: 0,
            text_runs: 0,
            primitives: Vec::new(),
        };
        for shape in page.shapes() {
            self.layout_shape(
                package,
                &resolver,
                references.as_ref(),
                page_part,
                shape,
                0,
                &mut state,
            )?;
        }
        for primitive in &mut state.primitives {
            bake_group_transform(primitive, Affine::identity());
        }
        let list = VsdxDisplayList {
            contract_version: CONTRACT_VERSION,
            width: page_width as f32 * PIXELS_PER_INCH,
            height: page_height as f32 * PIXELS_PER_INCH,
            paint_transform: final_paint_transform(page_height as f32),
            primitives: state.primitives,
        };
        if !display_list_finite(&list) {
            return Err(RenderError::PageDimensions(
                "non-finite display list".into(),
            ));
        }
        if serde_json::to_vec(&list).map_or(true, |bytes| {
            bytes.len() > self.limits.max_display_list_bytes
        }) {
            return Err(RenderError::Budget("display-list bytes"));
        }
        Ok(list)
    }
    #[allow(clippy::too_many_arguments)]
    fn layout_shape(
        &self,
        package: &VsdxPackage,
        resolver: &Resolver<'_>,
        references: Option<&PageShapeReferences>,
        page_part: &str,
        shape: &Shape,
        depth: usize,
        state: &mut State,
    ) -> Result<(), RenderError> {
        if depth >= MAX_RECURSION_DEPTH {
            return self.placeholder(shape, state, "group nesting depth exceeded");
        }
        state.count += 1;
        if state.count > self.limits.max_shapes {
            return Err(RenderError::Budget("shapes"));
        }
        let z_order = state.next_z();
        let id = format!("{page_part}:{}", shape.id);
        let resolved = resolver.resolve_shape(page_part, shape.id)?;
        if resolved.deleted {
            return Ok(());
        }
        if paint::number(&resolved, "NoShow").is_some_and(|value| value != 0.0) {
            return Ok(());
        }
        let Some(bounds) = bounds(package, references, &resolved, shape.id) else {
            return self.placeholder(shape, state, "unresolvable transform");
        };
        if !bounds_finite(bounds) {
            return self.placeholder_at(
                id,
                z_order,
                Bounds::default(),
                state,
                "overflowing transform",
            );
        }
        let geometry = resolved.sections.get("Geometry").map(realize_geometry);
        let child_shapes = shape.shapes().collect::<Vec<_>>();
        if !child_shapes.is_empty() {
            let transform = group_transform(
                bounds,
                child_coordinate_extent(
                    package,
                    resolver,
                    references,
                    page_part,
                    child_shapes.iter().copied(),
                ),
            );
            let start = state.primitives.len();
            for child in child_shapes {
                self.layout_shape(
                    package,
                    resolver,
                    references,
                    page_part,
                    child,
                    depth + 1,
                    state,
                )?;
            }
            let children = state.primitives.split_off(start);
            let z_order = state.next_z();
            state.primitives.push(Primitive::Group {
                id,
                z_order,
                primitives: children,
                transform,
            });
            return Ok(());
        }
        if let Some(data) = shape.foreign_data() {
            let asset_id = data.relationship_id.as_deref().and_then(|relationship_id| {
                package
                    .relationships
                    .get(page_part)?
                    .iter()
                    .find_map(|relationship| {
                        (relationship.id == relationship_id)
                            .then_some(relationship.resolved_target.as_deref())
                            .flatten()
                    })
            });
            if let Some(asset_id) = asset_id {
                if package.part_bytes(asset_id).is_none() {
                    return self.placeholder_at(
                        id,
                        z_order,
                        bounds,
                        state,
                        "dangling ForeignData image target",
                    );
                }
                state.primitives.push(Primitive::Image {
                    id,
                    z_order,
                    asset_id: asset_id.into(),
                    x: 0.0,
                    y: 0.0,
                    width: bounds.width as f32,
                    height: bounds.height as f32,
                    transform: bounds_affine(bounds, (0.0, 0.0), (1.0, 1.0)),
                });
                return Ok(());
            }
            return self.placeholder_at(
                id,
                z_order,
                bounds,
                state,
                "unsupported ForeignData image",
            );
        }
        let Some(geometry) = geometry else {
            return self.placeholder_at(
                id,
                z_order,
                bounds,
                state,
                "shape has no Geometry section",
            );
        };
        if !geometry.issues.is_empty() || geometry.commands.is_empty() {
            return self.placeholder_at(
                id,
                z_order,
                bounds,
                state,
                &format!("unsupported geometry: {:?}", geometry.issues),
            );
        }
        let path = geometry
            .commands
            .into_iter()
            .map(|mut command| {
                transform_command(&mut command, bounds);
                command
            })
            .collect::<Vec<_>>();
        let (fill, stroke) = match paint::paint(package, references, &resolved, shape.id) {
            Ok(paint) => paint,
            Err(reason) => {
                return self.placeholder_at(
                    id,
                    z_order,
                    bounds,
                    state,
                    &format!("unresolvable colour: {reason}"),
                );
            }
        };
        state.primitives.push(Primitive::Shape {
            id: id.clone(),
            z_order,
            path,
            fill,
            stroke,
            transform: Affine::identity(),
        });
        self.text(
            package, resolver, references, page_part, shape, &resolved, id, bounds, state,
        )?;
        Ok(())
    }
    #[allow(clippy::too_many_arguments)]
    fn text(
        &self,
        package: &VsdxPackage,
        resolver: &Resolver<'_>,
        references: Option<&PageShapeReferences>,
        page_part: &str,
        shape: &Shape,
        resolved: &ResolvedShape,
        id: String,
        bounds: Bounds,
        state: &mut State,
    ) -> Result<(), RenderError> {
        if matches!(resolved.cell("Char.Size"), Some(Lookup::Found(cell)) if cell.cell.value.as_deref().and_then(|value| value.parse::<f32>().ok()).is_some_and(|value| !value.is_finite()))
        {
            return Err(RenderError::PageDimensions(
                "non-finite text metrics".into(),
            ));
        }
        let page = package
            .page_part_ids
            .get(page_part)
            .and_then(|id| package.page_sheets.get(id))
            .or_else(|| package.page_contents.get(page_part))
            .ok_or_else(|| RenderError::MissingPage(page_part.into()))?;
        let tokens = resolver.resolve_text_in_context(shape, page, resolved)?;
        let mut paragraphs =
            rich_paragraphs(self, package, references, resolved, shape.id, &tokens);
        if paragraphs.iter().all(|paragraph| paragraph.runs.is_empty()) {
            return Ok(());
        }
        for paragraph in &mut paragraphs {
            if paragraph.align == 3
                && let Some(run) = paragraph.runs.first_mut()
            {
                append_diagnostic(run, "justify-fallback", "justify falls back to left");
            }
        }
        let text = paragraphs
            .iter()
            .flat_map(|paragraph| paragraph.runs.iter())
            .map(|run| run.text.as_str())
            .collect::<String>();
        let text_bytes = state
            .text_bytes
            .checked_add(text.len())
            .ok_or(RenderError::Budget("text bytes"))?;
        if text_bytes > self.limits.max_text_bytes {
            return Err(RenderError::Budget("text bytes"));
        }
        if state
            .text_paragraphs
            .checked_add(paragraphs.len())
            .ok_or(RenderError::Budget("text paragraphs"))?
            > self.limits.max_text_paragraphs
        {
            return Err(RenderError::Budget("text paragraphs"));
        }
        if state.text_lines >= self.limits.max_text_lines {
            return Err(RenderError::Budget("text lines"));
        }
        if state.text_runs >= self.limits.max_text_runs {
            return Err(RenderError::Budget("text runs"));
        }
        state.text_bytes = text_bytes;
        state.text_paragraphs += paragraphs.len();
        state.text_runs += paragraphs
            .iter()
            .map(|paragraph| paragraph.runs.len())
            .sum::<usize>();
        if state.text_runs > self.limits.max_text_runs {
            return Err(RenderError::Budget("text runs"));
        }
        let left = paint::number(resolved, "LeftMargin").unwrap_or(0.0) as f32;
        let right = paint::number(resolved, "RightMargin").unwrap_or(0.0) as f32;
        let top = paint::number(resolved, "TopMargin").unwrap_or(0.0) as f32;
        let bottom = paint::number(resolved, "BottomMargin").unwrap_or(0.0) as f32;
        let block_width = paint::number(resolved, "TxtWidth").unwrap_or(bounds.width) as f32;
        let block_height = paint::number(resolved, "TxtHeight").unwrap_or(bounds.height) as f32;
        let x = bounds.x as f32 + left;
        let y = bounds.y as f32 + top;
        let available_width = (block_width - left - right).max(0.0);
        let mut lines = Vec::new();
        let mut cursor_y = y;
        let mut offset = 0u32;
        for paragraph in &paragraphs {
            if paragraph.runs.is_empty() {
                continue;
            }
            cursor_y += paragraph.before;
            let width = (available_width - paragraph.left - paragraph.right).max(0.0);
            let mut laid_out =
                self.wrap_paragraph(paragraph, width, x + paragraph.left, cursor_y, offset);
            if let Some(first) = laid_out.first_mut() {
                first.x += paragraph.first;
                for stop in &mut first.caret_stops {
                    stop.x += paragraph.first;
                }
            }
            if let Some(spacing) = paragraph.line_spacing {
                let solid = laid_out.first().map_or(0.2, |line| line.height / 1.2);
                let line_height = if spacing > 0.0 {
                    spacing
                } else if spacing < 0.0 {
                    solid * -spacing / 100.0
                } else {
                    solid
                };
                for (index, line) in laid_out.iter_mut().enumerate() {
                    let y = cursor_y + index as f32 * line_height;
                    line.y = y;
                    line.height = line_height;
                    for stop in &mut line.caret_stops {
                        stop.y = y;
                    }
                }
            }
            offset += paragraph
                .runs
                .iter()
                .map(|run| run.text.len() as u32)
                .sum::<u32>();
            cursor_y += laid_out.iter().map(|line| line.height).sum::<f32>() + paragraph.after;
            lines.extend(laid_out);
        }
        if state
            .text_lines
            .checked_add(lines.len())
            .ok_or(RenderError::Budget("text lines"))?
            > self.limits.max_text_lines
        {
            return Err(RenderError::Budget("text lines"));
        }
        state.text_lines += lines.len();
        let used_height = cursor_y - y;
        let vertical = paint::number(resolved, "VerticalAlign").unwrap_or(0.0) as i32;
        let dy = match vertical {
            1 => (block_height - top - bottom - used_height) * 0.5,
            2 => block_height - top - bottom - used_height,
            _ => 0.0,
        }
        .max(0.0);
        for line in &mut lines {
            line.y += dy;
            for stop in &mut line.caret_stops {
                stop.y += dy;
            }
        }
        let z_order = state.next_z();
        state.primitives.push(Primitive::TextBox {
            id,
            z_order,
            x,
            y,
            width: available_width,
            height: (block_height - top - bottom).max(0.0),
            paragraphs: paragraphs
                .into_iter()
                .map(|paragraph| TextParagraph {
                    runs: paragraph.runs,
                })
                .collect(),
            lines,
            transform: shape_transform(bounds),
        });
        Ok(())
    }
    fn wrap_paragraph(
        &self,
        paragraph: &RichParagraph,
        available_width: f32,
        x: f32,
        y: f32,
        offset: u32,
    ) -> Vec<PositionedLine> {
        let text = paragraph
            .runs
            .iter()
            .map(|run| run.text.as_str())
            .collect::<String>();
        let breaks = ooxml_text::break_opportunities(&text);
        let mut widths = Vec::new();
        let mut height = 0.0f32;
        for (index, ch) in text.char_indices() {
            let run = paragraph
                .runs
                .iter()
                .scan(0usize, |end, run| {
                    *end += run.text.len();
                    Some((*end > index).then_some(run))
                })
                .find_map(|run| run);
            let run_end = paragraph
                .runs
                .iter()
                .scan(0usize, |end, run| {
                    *end += run.text.len();
                    Some((*end > index).then_some(*end))
                })
                .find_map(|end| end);
            let (mut width, line_height) = run
                .map(|run| {
                    (
                        self.measure(run, &ch.to_string())
                            + if run_end > Some(index + ch.len_utf8()) {
                                run.letter_spacing
                            } else {
                                0.0
                            },
                        run.size_in * 1.2,
                    )
                })
                .unwrap_or((0.0, 0.2));
            if ch == '\t'
                && let Some(tab) = run.and_then(|run| run.tab)
            {
                let cursor = widths.iter().map(|(_, _, width)| *width).sum::<f32>();
                let position = tab.default_interval.map_or(tab.position, |interval| {
                    ((cursor / interval).floor() + 1.0) * interval
                });
                let following = text[index + ch.len_utf8()..]
                    .split('\t')
                    .next()
                    .unwrap_or_default();
                let following_width =
                    self.measure_text(paragraph, index + ch.len_utf8(), following);
                let before_decimal = following
                    .split_once('.')
                    .map_or(following, |(before, _)| before);
                let decimal_width =
                    self.measure_text(paragraph, index + ch.len_utf8(), before_decimal);
                let aligned = match tab.alignment {
                    1 => position - following_width * 0.5,
                    2 => position - following_width,
                    3 => position - decimal_width,
                    _ => position,
                };
                width = (aligned - cursor).max(0.0);
            }
            widths.push((index, ch.len_utf8(), width));
            height = height.max(line_height);
        }
        let height = height.max(0.2);
        let mut result = Vec::new();
        let mut start = 0usize;
        let mut cursor = 0.0;
        let mut last_break = None;
        let mut mandatory_break = None;
        for (position, length, width) in widths.iter().copied() {
            if mandatory_break == Some(position) {
                result.push(self.line(
                    &text,
                    &widths,
                    start,
                    position,
                    x,
                    y + result.len() as f32 * height,
                    height,
                    offset,
                    paragraph.align,
                    available_width,
                ));
                start = position;
                cursor = 0.0;
                last_break = None;
                mandatory_break = None;
            }
            if position > start && cursor + width > available_width && available_width > 0.0 {
                let end = last_break.unwrap_or(position).max(start + 1);
                result.push(self.line(
                    &text,
                    &widths,
                    start,
                    end,
                    x,
                    y + result.len() as f32 * height,
                    height,
                    offset,
                    paragraph.align,
                    available_width,
                ));
                start = end;
                cursor = widths
                    .iter()
                    .filter(|(p, _, _)| *p >= start && *p < position)
                    .map(|(_, _, w)| *w)
                    .sum();
                last_break = None;
            }
            cursor += width;
            let next = position + length;
            if let Some(opportunity) = breaks.iter().find(|item| item.byte_index == next) {
                if opportunity.mandatory && next < text.len() {
                    mandatory_break = Some(next);
                } else {
                    last_break = Some(next);
                }
            }
        }
        if start < text.len() || result.is_empty() {
            result.push(self.line(
                &text,
                &widths,
                start,
                text.len(),
                x,
                y + result.len() as f32 * height,
                height,
                offset,
                paragraph.align,
                available_width,
            ));
        }
        result
    }
    #[allow(clippy::too_many_arguments)]
    fn line(
        &self,
        _text: &str,
        widths: &[(usize, usize, f32)],
        start: usize,
        end: usize,
        x: f32,
        y: f32,
        height: f32,
        offset: u32,
        align: i32,
        available: f32,
    ) -> PositionedLine {
        let width = widths
            .iter()
            .filter(|(position, _, _)| *position >= start && *position < end)
            .map(|(_, _, width)| *width)
            .sum::<f32>();
        let line_x = x + match align {
            1 => (available - width) * 0.5,
            2 => available - width,
            _ => 0.0,
        };
        let mut advance = 0.0;
        let mut stops = Vec::new();
        for (position, length, char_width) in widths
            .iter()
            .copied()
            .filter(|(position, _, _)| *position >= start && *position < end)
        {
            stops.push(CaretStop {
                position: offset + position as u32,
                x: line_x + advance,
                y,
            });
            advance += char_width;
            if position + length == end {
                stops.push(CaretStop {
                    position: offset + end as u32,
                    x: line_x + advance,
                    y,
                });
            }
        }
        PositionedLine {
            x: line_x,
            y,
            width,
            height,
            start: offset + start as u32,
            end: offset + end as u32,
            caret_stops: stops,
        }
    }
    /// Uses the effective face, then its style variants, then Arial and sans-serif.
    fn font_for(&self, family: &str, bold: bool, italic: bool) -> Option<ooxml_text::FontId> {
        std::iter::once(family)
            .chain(["Arial", "sans-serif"])
            .flat_map(|family| {
                [
                    (bold, italic),
                    (bold, false),
                    (false, italic),
                    (false, false),
                ]
                .into_iter()
                .map(move |(bold, italic)| (family.into(), bold, italic))
            })
            .find_map(|key| self.registered_fonts.get(&key))
            .copied()
    }
    fn measure(&self, run: &TextRun, text: &str) -> f32 {
        self.font_for(&run.family, run.bold, run.italic)
            .and_then(|font| ooxml_text::shape(&self.fonts, font, text, run.size_in, &[]).ok())
            .map(|glyphs| glyphs.iter().map(|glyph| glyph.x_advance).sum())
            .unwrap_or(text.chars().count() as f32 * run.size_in * 0.5)
    }
    fn measure_text(&self, paragraph: &RichParagraph, start: usize, text: &str) -> f32 {
        text.char_indices()
            .map(|(relative, ch)| {
                let index = start + relative;
                paragraph
                    .runs
                    .iter()
                    .scan(0usize, |end, run| {
                        *end += run.text.len();
                        Some((*end > index).then_some(run))
                    })
                    .find_map(|run| run)
                    .map_or(0.0, |run| {
                        let end = paragraph
                            .runs
                            .iter()
                            .scan(0usize, |end, run| {
                                *end += run.text.len();
                                Some((*end > index).then_some(*end))
                            })
                            .find_map(|end| end)
                            .unwrap_or(index + ch.len_utf8());
                        self.measure(run, &ch.to_string())
                            + (end > index + ch.len_utf8()) as u8 as f32 * run.letter_spacing
                    })
            })
            .sum()
    }
    fn placeholder(
        &self,
        shape: &Shape,
        state: &mut State,
        reason: &str,
    ) -> Result<(), RenderError> {
        self.placeholder_at(
            format!("shape:{}", shape.id),
            state.next_z(),
            Bounds::default(),
            state,
            reason,
        )
    }
    fn placeholder_at(
        &self,
        id: String,
        z_order: u32,
        bounds: Bounds,
        state: &mut State,
        reason: &str,
    ) -> Result<(), RenderError> {
        state.primitives.push(Primitive::Placeholder {
            id,
            z_order,
            x: bounds.x as f32,
            y: bounds.y as f32,
            width: bounds.width as f32,
            height: bounds.height as f32,
            reason: reason.into(),
        });
        Ok(())
    }
}
fn bake_group_transform(primitive: &mut Primitive, matrix: Affine) {
    match primitive {
        Primitive::Shape {
            path, transform, ..
        } => {
            for command in path {
                transform_affine(command, matrix);
            }
            *transform = Affine::identity();
        }
        Primitive::Image { transform, .. } => *transform = matrix.compose(*transform),
        Primitive::TextBox { transform, .. } => *transform = matrix.compose(*transform),
        Primitive::Placeholder {
            x,
            y,
            width,
            height,
            ..
        } => transform_rect(x, y, width, height, matrix),
        Primitive::Group {
            primitives,
            transform,
            ..
        } => {
            let matrix = matrix.compose(*transform);
            for child in primitives {
                bake_group_transform(child, matrix);
            }
            *transform = Affine::identity();
        }
    }
}
fn transform_rect(x: &mut f32, y: &mut f32, width: &mut f32, height: &mut f32, matrix: Affine) {
    let corners = [
        matrix.apply_point(*x, *y),
        matrix.apply_point(*x + *width, *y),
        matrix.apply_point(*x, *y + *height),
        matrix.apply_point(*x + *width, *y + *height),
    ];
    let (min_x, max_x) = corners
        .iter()
        .map(|(x, _)| *x)
        .fold((f32::INFINITY, f32::NEG_INFINITY), |(min, max), value| {
            (min.min(value), max.max(value))
        });
    let (min_y, max_y) = corners
        .iter()
        .map(|(_, y)| *y)
        .fold((f32::INFINITY, f32::NEG_INFINITY), |(min, max), value| {
            (min.min(value), max.max(value))
        });
    *x = min_x;
    *y = min_y;
    *width = max_x - min_x;
    *height = max_y - min_y;
}
fn transform_command(command: &mut ooxml_drawingml::GeometryPathCommand, bounds: Bounds) {
    transform_affine(command, bounds_affine(bounds, (0.0, 0.0), (1.0, 1.0)));
}
fn transform_affine(command: &mut ooxml_drawingml::GeometryPathCommand, matrix: Affine) {
    use ooxml_drawingml::GeometryPathCommand::*;
    let point = |x: &mut f64, y: &mut f64| {
        let (transformed_x, transformed_y) = matrix.apply_point(*x as f32, *y as f32);
        (*x, *y) = (transformed_x as f64, transformed_y as f64);
    };
    match command {
        Move { x, y } | Line { x, y } => point(x, y),
        Quad { cpx, cpy, x, y } => {
            point(cpx, cpy);
            point(x, y);
        }
        Cubic {
            cp1x,
            cp1y,
            cp2x,
            cp2y,
            x,
            y,
        } => {
            point(cp1x, cp1y);
            point(cp2x, cp2y);
            point(x, y);
        }
        Close => {}
    }
}
fn bounds_affine(group: Bounds, origin: (f64, f64), scale: (f64, f64)) -> Affine {
    let loc_x = group.loc_pin_x;
    let loc_y = group.loc_pin_y;
    let pin_x = group.x + group.loc_pin_x;
    let pin_y = group.y + group.loc_pin_y;
    let (sin, cos) = group.angle.sin_cos();
    let (origin_x, origin_y) = origin;
    let (scale_x, scale_y) = scale;
    let sx = scale_x * if group.flip_x { -1.0 } else { 1.0 };
    let sy = scale_y * if group.flip_y { -1.0 } else { 1.0 };
    Affine {
        a: (cos * sx) as f32,
        b: (sin * sx) as f32,
        c: (-sin * sy) as f32,
        d: (cos * sy) as f32,
        e: (pin_x - cos * sx * (loc_x + origin_x) + sin * sy * (loc_y + origin_y)) as f32,
        f: (pin_y - sin * sx * (loc_x + origin_x) - cos * sy * (loc_y + origin_y)) as f32,
    }
}
/// Visio documents the lower-left local origin, but not how to derive a group child scale.
/// Use the conservative selection-extent ratio and translate that documented origin.
fn group_transform(group: Bounds, child_extent: Option<ChildCoordinateExtent>) -> Affine {
    let (origin, scale) = child_extent
        .map(|extent| {
            (
                (extent.x, extent.y),
                (group.width / extent.width, group.height / extent.height),
            )
        })
        .unwrap_or(((0.0, 0.0), (1.0, 1.0)));
    bounds_affine(group, origin, scale)
}
/// Maps a shape's own absolute, pre-rotation bounds into its rotated position, the same
/// transform already baked into its geometry path, so geometry and text agree.
fn shape_transform(bounds: Bounds) -> Affine {
    bounds_affine(bounds, (bounds.x, bounds.y), (1.0, 1.0))
}
struct ChildCoordinateExtent {
    x: f64,
    y: f64,
    width: f64,
    height: f64,
}
fn child_coordinate_extent<'a>(
    package: &VsdxPackage,
    resolver: &Resolver<'_>,
    references: Option<&PageShapeReferences>,
    page_part: &str,
    shapes: impl Iterator<Item = &'a Shape>,
) -> Option<ChildCoordinateExtent> {
    let bounds = shapes
        .filter_map(|shape| {
            resolver
                .resolve_shape(page_part, shape.id)
                .ok()
                .and_then(|resolved| bounds(package, references, &resolved, shape.id))
        })
        .collect::<Vec<_>>();
    let min_x = bounds.iter().map(|bounds| bounds.x).reduce(f64::min)?;
    let min_y = bounds.iter().map(|bounds| bounds.y).reduce(f64::min)?;
    let max_x = bounds
        .iter()
        .map(|bounds| bounds.x + bounds.width)
        .reduce(f64::max)?;
    let max_y = bounds
        .iter()
        .map(|bounds| bounds.y + bounds.height)
        .reduce(f64::max)?;
    let width = max_x - min_x;
    let height = max_y - min_y;
    (width > 0.0 && height > 0.0).then_some(ChildCoordinateExtent {
        x: min_x,
        y: min_y,
        width,
        height,
    })
}
struct State {
    count: usize,
    z_order: u32,
    text_bytes: usize,
    text_paragraphs: usize,
    text_lines: usize,
    text_runs: usize,
    primitives: Vec<Primitive>,
}
impl State {
    fn next_z(&mut self) -> u32 {
        let z = self.z_order;
        self.z_order += 1;
        z
    }
}
#[derive(Clone, Copy, Default)]
struct Bounds {
    x: f64,
    y: f64,
    width: f64,
    height: f64,
    loc_pin_x: f64,
    loc_pin_y: f64,
    angle: f64,
    flip_x: bool,
    flip_y: bool,
}
fn bounds_finite(bounds: Bounds) -> bool {
    [
        bounds.x,
        bounds.y,
        bounds.width,
        bounds.height,
        bounds.loc_pin_x,
        bounds.loc_pin_y,
        bounds.angle,
    ]
    .into_iter()
    .all(|value| value.is_finite() && (value as f32).is_finite())
}
fn page_dimension(
    resolver: &Resolver<'_>,
    package: &VsdxPackage,
    page: &str,
    name: &str,
) -> Option<f64> {
    let page_id = package.page_part_ids.get(page)?;
    let sheet = package.page_sheets.get(page_id)?;
    let resolved = resolver.resolve_sheet(sheet).ok()?;
    let Lookup::Found(cell) = resolved.cell(name)? else {
        return None;
    };
    if let Some(formula) = cell.cell.formula.as_deref()
        && let Evaluation::Evaluated(result) = evaluate_cell_with_shape_package_theme(
            name,
            formula,
            &resolved,
            &ParseLimits::default(),
            &resolved,
            package,
        )
        && let Value::Number(number) = result.value
        && number.number.is_finite()
    {
        return Some(number.number);
    }
    cell.cell
        .value
        .as_deref()?
        .parse()
        .ok()
        .filter(|value: &f64| value.is_finite())
}
fn bounds(
    package: &VsdxPackage,
    references: Option<&PageShapeReferences>,
    shape: &ResolvedShape,
    shape_id: u32,
) -> Option<Bounds> {
    let value = |name| evaluated(package, references, shape, shape_id, name);
    let width = value("Width")?;
    let height = value("Height")?;
    let pin_x = value("PinX")?;
    let pin_y = value("PinY")?;
    let loc_pin_x = value("LocPinX").unwrap_or(width / 2.0);
    let loc_pin_y = value("LocPinY").unwrap_or(height / 2.0);
    Some(Bounds {
        x: pin_x - loc_pin_x,
        y: pin_y - loc_pin_y,
        width,
        height,
        loc_pin_x,
        loc_pin_y,
        angle: value("Angle").unwrap_or(0.0),
        flip_x: value("FlipX").unwrap_or(0.0) != 0.0,
        flip_y: value("FlipY").unwrap_or(0.0) != 0.0,
    })
}
fn evaluated(
    package: &VsdxPackage,
    references: Option<&PageShapeReferences>,
    shape: &ResolvedShape,
    shape_id: u32,
    name: &str,
) -> Option<f64> {
    let Lookup::Found(cell) = shape.cell(name)? else {
        return None;
    };
    if let (Some(formula), Some(references)) = (cell.cell.formula.as_deref(), references)
        && let Evaluation::Evaluated(result) = evaluate_cell_with_shape_package_theme(
            name,
            formula,
            &references.for_shape(shape_id),
            &ParseLimits::default(),
            shape,
            package,
        )
        && let Value::Number(number) = result.value
        && number.number.is_finite()
    {
        return Some(number.number);
    }
    cell.cell
        .value
        .as_deref()?
        .parse()
        .ok()
        .filter(|value: &f64| value.is_finite())
}
fn display_list_finite(list: &VsdxDisplayList) -> bool {
    list.width.is_finite()
        && list.height.is_finite()
        && [
            list.paint_transform.a,
            list.paint_transform.b,
            list.paint_transform.c,
            list.paint_transform.d,
            list.paint_transform.e,
            list.paint_transform.f,
        ]
        .into_iter()
        .all(f32::is_finite)
        && primitives_finite(&list.primitives)
}
fn primitives_finite(primitives: &[Primitive]) -> bool {
    primitives.iter().all(|primitive| match primitive {
        Primitive::Shape {
            path,
            transform,
            fill,
            stroke,
            ..
        } => {
            transform.is_finite()
                && path.iter().all(command_finite)
                && paint_finite(fill)
                && stroke
                    .as_ref()
                    .is_none_or(|stroke| stroke.width.is_finite())
        }
        Primitive::Image {
            x,
            y,
            width,
            height,
            transform,
            ..
        } => [*x, *y, *width, *height].into_iter().all(f32::is_finite) && transform.is_finite(),
        Primitive::TextBox {
            x,
            y,
            width,
            height,
            paragraphs,
            lines,
            transform,
            ..
        } => {
            [*x, *y, *width, *height].into_iter().all(f32::is_finite)
                && transform.is_finite()
                && lines.iter().all(|line| {
                    [line.x, line.y, line.width, line.height]
                        .into_iter()
                        .all(f32::is_finite)
                        && line
                            .caret_stops
                            .iter()
                            .all(|stop| stop.x.is_finite() && stop.y.is_finite())
                })
                && paragraphs
                    .iter()
                    .all(|paragraph| paragraph.runs.iter().all(|run| run.size_in.is_finite()))
        }
        Primitive::Placeholder {
            x,
            y,
            width,
            height,
            ..
        } => [*x, *y, *width, *height].into_iter().all(f32::is_finite),
        Primitive::Group {
            transform,
            primitives,
            ..
        } => transform.is_finite() && primitives_finite(primitives),
    })
}
fn paint_finite(paint: &Option<Paint>) -> bool {
    match paint {
        Some(Paint::Gradient { stops }) => stops.iter().all(|stop| stop.position.is_finite()),
        _ => true,
    }
}
fn command_finite(command: &ooxml_drawingml::GeometryPathCommand) -> bool {
    use ooxml_drawingml::GeometryPathCommand::*;
    match command {
        Move { x, y } | Line { x, y } => x.is_finite() && y.is_finite(),
        Quad { cpx, cpy, x, y } => {
            cpx.is_finite() && cpy.is_finite() && x.is_finite() && y.is_finite()
        }
        Cubic {
            cp1x,
            cp1y,
            cp2x,
            cp2y,
            x,
            y,
        } => [*cp1x, *cp1y, *cp2x, *cp2y, *x, *y]
            .into_iter()
            .all(f64::is_finite),
        Close => true,
    }
}

#[derive(Clone, Debug, PartialEq)]
pub enum HitTestResult {
    Shape { shape_id: String },
    Text { shape_id: String, position: u32 },
}
pub fn hit_test(list: &VsdxDisplayList, x: f32, y: f32) -> Option<HitTestResult> {
    let inverse = Affine {
        a: list.paint_transform.a,
        b: list.paint_transform.b,
        c: list.paint_transform.c,
        d: list.paint_transform.d,
        e: list.paint_transform.e,
        f: list.paint_transform.f,
    }
    .invert()?;
    let (x, y) = inverse.apply_point(x, y);
    fn collect<'a>(primitives: &'a [Primitive], output: &mut Vec<&'a Primitive>) {
        for primitive in primitives {
            output.push(primitive);
            if let Primitive::Group { primitives, .. } = primitive {
                collect(primitives, output);
            }
        }
    }
    let mut primitives = Vec::new();
    collect(&list.primitives, &mut primitives);
    primitives.sort_by_key(|primitive| std::cmp::Reverse(z_order(primitive)));
    for primitive in primitives {
        match primitive {
            Primitive::TextBox {
                id,
                x: left,
                y: top,
                width,
                height,
                lines,
                transform,
                ..
            } => {
                if let Some(position) =
                    text_caret_at((*left, *top, *width, *height), lines, *transform, (x, y))
                {
                    return Some(HitTestResult::Text {
                        shape_id: id.clone(),
                        position,
                    });
                }
            }
            Primitive::Shape {
                id,
                path,
                fill,
                stroke,
                transform,
                ..
            } if path_hit(
                path,
                *transform,
                fill.is_some(),
                stroke.as_ref().map_or(0.0, |stroke| stroke.width),
                x,
                y,
            ) =>
            {
                return Some(HitTestResult::Shape {
                    shape_id: id.clone(),
                });
            }
            Primitive::Image {
                id,
                x: left,
                y: top,
                width,
                height,
                transform,
                ..
            } if point_in_transformed_rect(*left, *top, *width, *height, *transform, x, y) => {
                return Some(HitTestResult::Shape {
                    shape_id: id.clone(),
                });
            }
            Primitive::Placeholder {
                id,
                x: left,
                y: top,
                width,
                height,
                ..
            } if x >= *left && x <= left + width && y >= *top && y <= top + height => {
                return Some(HitTestResult::Shape {
                    shape_id: id.clone(),
                });
            }
            _ => {}
        }
    }
    None
}

fn z_order(primitive: &Primitive) -> u32 {
    match primitive {
        Primitive::Shape { z_order, .. }
        | Primitive::Image { z_order, .. }
        | Primitive::TextBox { z_order, .. }
        | Primitive::Placeholder { z_order, .. }
        | Primitive::Group { z_order, .. } => *z_order,
    }
}

/// Probes a text box in its own local space so a rotated box rejects points that only fall
/// inside its axis-aligned bounds, and compares caret stops in the space they were laid out in.
fn text_caret_at(
    (left, top, width, height): (f32, f32, f32, f32),
    lines: &[PositionedLine],
    transform: Affine,
    (x, y): (f32, f32),
) -> Option<u32> {
    let (x, y) = transform.invert()?.apply_point(x, y);
    if x < left || x > left + width || y < top || y > top + height {
        return None;
    }
    lines
        .iter()
        .min_by(|a, b| (a.y - y).abs().total_cmp(&(b.y - y).abs()))?
        .caret_stops
        .iter()
        .min_by(|a, b| (a.x - x).abs().total_cmp(&(b.x - x).abs()))
        .map(|stop| stop.position)
}
fn point_in_transformed_rect(
    left: f32,
    top: f32,
    width: f32,
    height: f32,
    transform: Affine,
    x: f32,
    y: f32,
) -> bool {
    let Some(inverse) = transform.invert() else {
        return false;
    };
    let (x, y) = inverse.apply_point(x, y);
    x >= left && x <= left + width && y >= top && y <= top + height
}
fn path_hit(
    path: &[ooxml_drawingml::GeometryPathCommand],
    transform: Affine,
    fill: bool,
    stroke_width: f32,
    x: f32,
    y: f32,
) -> bool {
    let Some(inverse) = transform.invert() else {
        return false;
    };
    let (x, y) = inverse.apply_point(x, y);
    let points = flatten_path(path);
    if points.is_empty() {
        return false;
    }
    let mut inside = false;
    for index in 0..points.len() {
        let (ax, ay) = points[index];
        let (bx, by) = points[(index + 1) % points.len()];
        if (ay > y) != (by > y) && x < (bx - ax) * (y - ay) / (by - ay) + ax {
            inside = !inside;
        }
    }
    let closed = path
        .iter()
        .any(|command| matches!(command, ooxml_drawingml::GeometryPathCommand::Close));
    fill && inside
        || stroke_width > 0.0
            && points.windows(2).any(|segment| {
                point_segment_distance(x, y, segment[0], segment[1]) <= stroke_width / 2.0
            })
        || stroke_width > 0.0
            && closed
            && points.len() > 1
            && point_segment_distance(x, y, points[points.len() - 1], points[0])
                <= stroke_width / 2.0
}
fn point_segment_distance(x: f32, y: f32, (ax, ay): (f32, f32), (bx, by): (f32, f32)) -> f32 {
    let dx = bx - ax;
    let dy = by - ay;
    let length = dx * dx + dy * dy;
    let t = if length == 0.0 {
        0.0
    } else {
        (((x - ax) * dx + (y - ay) * dy) / length).clamp(0.0, 1.0)
    };
    ((x - (ax + t * dx)).powi(2) + (y - (ay + t * dy)).powi(2)).sqrt()
}
fn flatten_path(path: &[ooxml_drawingml::GeometryPathCommand]) -> Vec<(f32, f32)> {
    use ooxml_drawingml::GeometryPathCommand::*;
    let mut points = Vec::new();
    let mut current = (0.0, 0.0);
    for command in path {
        match *command {
            Move { x, y } => {
                current = (x as f32, y as f32);
                points.push(current);
            }
            Line { x, y } => {
                current = (x as f32, y as f32);
                points.push(current);
            }
            Quad { cpx, cpy, x, y } => {
                let start = current;
                for step in 1..=16 {
                    let t = step as f32 / 16.0;
                    let u = 1.0 - t;
                    points.push((
                        u * u * start.0 + 2.0 * u * t * cpx as f32 + t * t * x as f32,
                        u * u * start.1 + 2.0 * u * t * cpy as f32 + t * t * y as f32,
                    ));
                }
                current = (x as f32, y as f32);
            }
            Cubic {
                cp1x,
                cp1y,
                cp2x,
                cp2y,
                x,
                y,
            } => {
                let start = current;
                for step in 1..=16 {
                    let t = step as f32 / 16.0;
                    let u = 1.0 - t;
                    points.push((
                        u.powi(3) * start.0
                            + 3.0 * u * u * t * cp1x as f32
                            + 3.0 * u * t * t * cp2x as f32
                            + t.powi(3) * x as f32,
                        u.powi(3) * start.1
                            + 3.0 * u * u * t * cp1y as f32
                            + 3.0 * u * t * t * cp2y as f32
                            + t.powi(3) * y as f32,
                    ));
                }
                current = (x as f32, y as f32);
            }
            Close => {}
        }
    }
    points
}

#[cfg(test)]
mod tests {
    use super::*;
    use ooxml_drawingml::GeometryPathCommand;
    use vsdx_parse::{
        Cell, ForeignData, Row, RowChild, Section, SectionChild, ShapeChild, ShapesChild, Sheet,
        SheetChild, TextToken,
    };

    fn package(shapes: Vec<Shape>) -> VsdxPackage {
        let mut package: VsdxPackage = serde_json::from_value(serde_json::json!({
            "documentPartPath": "", "pagesPartPath": null, "mastersPartPath": null,
            "pagePartPaths": ["page"], "masterPartPaths": [], "themePartPaths": [],
            "windowsPartPath": null, "relationships": {}, "documentSheet": null,
            "styleSheets": [], "colors": [], "faceNames": [], "pageSheets": {},
            "masterSheets": {}, "pagePartIds": {"page": 1}, "masterPartIds": {},
            "pageContents": {}, "masterContents": {}
        }))
        .unwrap();
        package.page_sheets.insert(
            1,
            Sheet {
                id: None,
                children: vec![
                    SheetChild::Cell(cell("PageWidth", "10")),
                    SheetChild::Cell(cell("PageHeight", "8")),
                ],
                other_attrs: vec![],
            },
        );
        package.page_contents.insert(
            "page".into(),
            Sheet {
                id: None,
                children: vec![SheetChild::Shapes(
                    shapes.into_iter().map(ShapesChild::Shape).collect(),
                )],
                other_attrs: vec![],
            },
        );
        package
    }

    fn cell(name: &str, value: &str) -> Cell {
        Cell {
            name: name.into(),
            formula: None,
            value: Some(value.into()),
            unit: None,
            del: false,
            other_attrs: vec![],
        }
    }

    fn formula(name: &str, value: &str) -> Cell {
        Cell {
            name: name.into(),
            formula: Some(value.into()),
            value: None,
            unit: None,
            del: false,
            other_attrs: vec![],
        }
    }

    fn row(index: u32, row_type: &str, cells: Vec<Cell>) -> Row {
        Row {
            index: Some(index),
            name: None,
            local_name: None,
            row_type: Some(row_type.into()),
            del: false,
            children: cells.into_iter().map(RowChild::Cell).collect(),
            other_attrs: vec![],
        }
    }

    fn rectangle() -> Section {
        Section {
            name: "Geometry".into(),
            index: None,
            del: false,
            children: vec![
                row(0, "MoveTo", vec![cell("X", "0"), cell("Y", "0")]),
                row(1, "LineTo", vec![cell("X", "1"), cell("Y", "0")]),
                row(2, "LineTo", vec![cell("X", "1"), cell("Y", "1")]),
                row(3, "LineTo", vec![cell("X", "0"), cell("Y", "1")]),
            ]
            .into_iter()
            .map(SectionChild::Row)
            .collect(),
            other_attrs: vec![],
        }
    }

    fn shape(id: u32, pin_x: f64, pin_y: f64) -> Shape {
        Shape {
            id,
            name: None,
            name_u: None,
            shape_type: None,
            master: None,
            master_shape: None,
            line_style: None,
            fill_style: None,
            text_style: None,
            del: false,
            other_attrs: vec![],
            children: vec![
                ShapeChild::Cell(cell("Width", "1")),
                ShapeChild::Cell(cell("Height", "1")),
                ShapeChild::Cell(cell("PinX", &pin_x.to_string())),
                ShapeChild::Cell(cell("PinY", &pin_y.to_string())),
                ShapeChild::Cell(cell("LocPinX", "0")),
                ShapeChild::Cell(cell("LocPinY", "0")),
                ShapeChild::Cell(cell("FillPattern", "1")),
                ShapeChild::Cell(formula("FillForegnd", "RGB(1,2,3)")),
                ShapeChild::Cell(cell("LinePattern", "1")),
                ShapeChild::Cell(formula("LineColor", "RGB(4,5,6)")),
                ShapeChild::Cell(cell("LineWeight", "0.02")),
                ShapeChild::Section(rectangle()),
            ],
        }
    }

    fn render(shapes: Vec<Shape>) -> VsdxDisplayList {
        let package = package(shapes);
        Renderer::default().layout_page(&package, "page").unwrap()
    }

    fn text_shape(tokens: Vec<TextToken>) -> Shape {
        let mut shape = shape(1, 1.0, 1.0);
        shape.children.push(ShapeChild::Text(tokens));
        shape
    }

    fn text_box(list: &VsdxDisplayList) -> &Primitive {
        list.primitives
            .iter()
            .find(|primitive| matches!(primitive, Primitive::TextBox { .. }))
            .unwrap()
    }

    fn text_section(name: &str, rows: Vec<Row>) -> ShapeChild {
        ShapeChild::Section(Section {
            name: name.into(),
            index: None,
            del: false,
            children: rows.into_iter().map(SectionChild::Row).collect(),
            other_attrs: vec![],
        })
    }

    fn shape_primitive(list: &VsdxDisplayList, id: u32) -> &Primitive {
        list.primitives.iter().find(|primitive| matches!(primitive, Primitive::Shape { id: actual, .. } if actual == &format!("page:{id}"))).unwrap()
    }

    fn group(id: u32, pin_x: f64, pin_y: f64, children: Vec<Shape>) -> Shape {
        let mut group = shape(id, pin_x, pin_y);
        group
            .children
            .retain(|child| !matches!(child, ShapeChild::Section(_)));
        group.children.push(ShapeChild::Shapes(
            children.into_iter().map(ShapesChild::Shape).collect(),
        ));
        group
    }

    fn with_cell(shape: &mut Shape, name: &str, value: &str) {
        shape.children.retain(
            |child| !matches!(child, ShapeChild::Cell(Cell { name: actual, .. }) if actual == name),
        );
        shape.children.push(ShapeChild::Cell(cell(name, value)));
    }

    fn assert_point_close(actual: (f32, f32), expected: (f32, f32)) {
        assert!(
            (actual.0 - expected.0).abs() < 1e-5,
            "{actual:?} != {expected:?}"
        );
        assert!(
            (actual.1 - expected.1).abs() < 1e-5,
            "{actual:?} != {expected:?}"
        );
    }
    #[test]
    fn flip_is_only_final_paint_transform() {
        let transform = final_paint_transform(10.0);
        assert_eq!(to_canvas(transform, 0.0, 0.0), (0.0, 960.0));
        assert_eq!(to_canvas(transform, 0.0, 10.0), (0.0, 0.0));
    }

    #[test]
    fn renderer_keeps_inches_until_the_final_canvas_transform() {
        let list = render(vec![shape(1, 2.0, 3.0)]);
        assert_eq!(to_canvas(list.paint_transform, 0.0, 0.0), (0.0, 768.0));
        assert_eq!(to_canvas(list.paint_transform, 0.0, 8.0), (0.0, 0.0));
        let Primitive::Shape { path, .. } = shape_primitive(&list, 1) else {
            unreachable!()
        };
        assert!(matches!(
            path[0],
            GeometryPathCommand::Move { x: 2.0, y: 3.0 }
        ));
    }

    #[test]
    fn font_size_round_trips_in_inches_and_paints_in_pixels() {
        let mut shape = text_shape(vec![TextToken::Literal("a".into())]);
        shape.children.push(text_section(
            "Character",
            vec![row(0, "", vec![cell("Size", "0.25")])],
        ));
        let list = render(vec![shape]);
        let Primitive::TextBox { paragraphs, .. } = text_box(&list) else {
            unreachable!()
        };
        assert_eq!(paragraphs[0].runs[0].size_in, 0.25);
        assert_eq!(
            to_canvas_length(list.paint_transform, paragraphs[0].runs[0].size_in),
            24.0
        );
    }

    #[test]
    fn tabs_advance_to_effective_stops_and_align_following_text() {
        for (alignment, expected_start) in [(0, 4.0), (1, 3.0), (2, 2.0)] {
            let mut shape = text_shape(vec![
                TextToken::CharacterRun(0),
                TextToken::Literal("a".into()),
                TextToken::Tab(0),
                TextToken::Literal("bc".into()),
            ]);
            shape.children.push(text_section(
                "Character",
                vec![row(0, "", vec![cell("Size", "2")])],
            ));
            shape.children.push(text_section(
                "Tabs",
                vec![row(
                    0,
                    "",
                    vec![
                        cell("Position", "4"),
                        cell("Alignment", &alignment.to_string()),
                    ],
                )],
            ));
            with_cell(&mut shape, "Width", "10");
            let list = render(vec![shape]);
            let Primitive::TextBox { lines, .. } = text_box(&list) else {
                unreachable!()
            };
            let start = lines[0]
                .caret_stops
                .iter()
                .find(|stop| stop.position == 2)
                .unwrap();
            assert!((start.x - (1.0 + expected_start)).abs() < 0.001);
        }
    }

    #[test]
    fn undefined_tab_uses_actual_default_interval_with_renderer_policy_diagnostic() {
        let mut shape = text_shape(vec![
            TextToken::CharacterRun(0),
            TextToken::Literal("a".into()),
            TextToken::Tab(3),
            TextToken::Literal("b".into()),
        ]);
        shape.children.push(text_section(
            "Tabs",
            vec![row(0, "", vec![cell("Position", "0.25")])],
        ));
        with_cell(&mut shape, "DefaultTabStop", "0.75");
        let list = render(vec![shape]);
        let Primitive::TextBox {
            paragraphs, lines, ..
        } = text_box(&list)
        else {
            unreachable!()
        };
        assert!(paragraphs[0].runs[1].diagnostics.iter().any(|diagnostic| {
            diagnostic.code == "missing-tab-position"
                && diagnostic
                    .detail
                    .contains("renderer policy DefaultTabStop 0.75 in")
        }));
        assert!((lines[0].caret_stops[2].x - 1.75).abs() < 0.001);
    }

    #[test]
    fn fields_render_cached_values_or_diagnostic_runs_in_order() {
        let mut shape = text_shape(vec![
            TextToken::Literal(" before ".into()),
            TextToken::Field(0),
            TextToken::Literal(" middle ".into()),
            TextToken::Field(1),
            TextToken::Literal(" after ".into()),
        ]);
        shape.children.push(text_section(
            "Field",
            vec![
                row(0, "", vec![cell("Value", "resolved")]),
                row(1, "", vec![cell("Format", "not a value")]),
            ],
        ));
        let list = render(vec![shape]);
        let Primitive::TextBox { paragraphs, .. } = text_box(&list) else {
            unreachable!()
        };
        let runs = &paragraphs[0].runs;
        assert_eq!(
            runs.iter().map(|run| run.text.as_str()).collect::<String>(),
            " before resolved middle [unresolved field] after "
        );
        assert!(runs[1].diagnostics.iter().any(|diagnostic| {
            diagnostic.code == "unregistered-font"
                && diagnostic.detail == "unregistered font 'sans-serif'"
                && diagnostic.category == DiagnosticCategory::Fidelity
        }));
        assert!(
            runs[3]
                .diagnostics
                .iter()
                .any(|diagnostic| diagnostic.code == "unresolvable-field"
                    && diagnostic.detail.contains("no Value cell")
                    && diagnostic.category == DiagnosticCategory::Integrity)
        );
    }

    #[test]
    fn word_wraps_at_uax14_opportunities_in_scene_inches() {
        let mut shape = text_shape(vec![TextToken::Literal("a b".into())]);
        with_cell(&mut shape, "TxtWidth", "0.18");
        let list = render(vec![shape]);
        let Primitive::TextBox { lines, .. } = text_box(&list) else {
            unreachable!()
        };
        assert_eq!(lines.len(), 2);
        assert_eq!((lines[0].start, lines[0].end), (0, 2));
        assert_eq!((lines[1].start, lines[1].end), (2, 3));
    }

    #[test]
    fn paragraph_alignment_positions_lines_in_scene_inches() {
        for (align, expected_x) in [(0, 1.0), (1, 1.458_333_4), (2, 1.916_666_6)] {
            let mut shape = text_shape(vec![
                TextToken::ParagraphRun(0),
                TextToken::Literal("a".into()),
            ]);
            with_cell(&mut shape, "TxtWidth", "1");
            shape.children.push(ShapeChild::Section(Section {
                name: "Paragraph".into(),
                index: None,
                del: false,
                children: vec![SectionChild::Row(row(
                    0,
                    "",
                    vec![cell("HorzAlign", &align.to_string())],
                ))],
                other_attrs: vec![],
            }));
            let list = render(vec![shape]);
            let Primitive::TextBox { lines, .. } = text_box(&list) else {
                unreachable!()
            };
            assert!((lines[0].x - expected_x).abs() < 1e-5);
        }
    }

    #[test]
    fn paragraph_indents_and_line_spacing_change_line_geometry() {
        let mut shape = text_shape(vec![
            TextToken::ParagraphRun(0),
            TextToken::Literal("a\nb".into()),
        ]);
        with_cell(&mut shape, "TxtWidth", "1");
        shape.children.push(ShapeChild::Section(Section {
            name: "Paragraph".into(),
            index: None,
            del: false,
            children: vec![SectionChild::Row(row(
                0,
                "",
                vec![
                    cell("IndLeft", "0.2"),
                    cell("IndRight", "0.3"),
                    cell("IndFirst", "0.1"),
                    cell("SpLine", "-200"),
                ],
            ))],
            other_attrs: vec![],
        }));
        let list = render(vec![shape]);
        let Primitive::TextBox { lines, .. } = text_box(&list) else {
            unreachable!()
        };
        assert_eq!(lines.len(), 2);
        assert!((lines[0].x - 1.3).abs() < 1e-5);
        assert!((lines[1].x - 1.2).abs() < 1e-5);
        assert!((lines[1].y - lines[0].y - 1.0 / 3.0).abs() < 1e-5);
    }

    #[test]
    fn character_defaults_case_position_and_tracking_apply_before_text() {
        let mut shape = text_shape(vec![TextToken::Literal("hello".into())]);
        shape.children.push(text_section(
            "Character",
            vec![row(
                0,
                "",
                vec![
                    cell("Case", "1"),
                    cell("Pos", "1"),
                    cell("Letterspace", "20"),
                    cell("Style", "12"),
                ],
            )],
        ));
        let list = render(vec![shape]);
        let Primitive::TextBox { paragraphs, .. } = text_box(&list) else {
            unreachable!()
        };
        let run = &paragraphs[0].runs[0];
        assert_eq!(run.text, "HELLO");
        assert!(run.superscript && run.underline && run.small_caps);
        assert!((run.letter_spacing - 1.0 / 72.0).abs() < 1e-5);
    }

    #[test]
    fn vertical_alignment_and_margins_inset_text_in_scene_inches() {
        for (align, expected_y) in [(0, 1.1), (1, 1.4), (2, 1.7)] {
            let mut shape = text_shape(vec![TextToken::Literal("a".into())]);
            let vertical = align.to_string();
            for (name, value) in [
                ("TxtWidth", "1"),
                ("TxtHeight", "1"),
                ("LeftMargin", "0.1"),
                ("RightMargin", "0.2"),
                ("TopMargin", "0.1"),
                ("BottomMargin", "0.1"),
                ("VerticalAlign", vertical.as_str()),
            ] {
                with_cell(&mut shape, name, value);
            }
            let list = render(vec![shape]);
            let Primitive::TextBox {
                x,
                y,
                width,
                height,
                lines,
                ..
            } = text_box(&list)
            else {
                unreachable!()
            };
            assert_point_close((*x, *y), (1.1, 1.1));
            assert_point_close((*width, *height), (0.7, 0.8));
            assert!((lines[0].y - expected_y).abs() < 1e-5);
            assert!((lines[0].height - 0.2).abs() < 1e-5);
        }
    }

    #[test]
    fn unregistered_font_records_a_diagnostic() {
        let mut shape = text_shape(vec![
            TextToken::CharacterRun(0),
            TextToken::Literal("a".into()),
        ]);
        shape.children.push(ShapeChild::Section(Section {
            name: "Character".into(),
            index: None,
            del: false,
            children: vec![SectionChild::Row(row(0, "", vec![cell("Font", "99")]))],
            other_attrs: vec![],
        }));
        let list = render(vec![shape]);
        let Primitive::TextBox { paragraphs, .. } = text_box(&list) else {
            unreachable!()
        };
        assert!(paragraphs[0].runs[0].diagnostics.iter().any(|diagnostic| {
            diagnostic.code == "unregistered-font" && diagnostic.detail == "unregistered font '99'"
        }));
    }

    #[test]
    fn unregistered_default_font_records_a_diagnostic() {
        let list = render(vec![text_shape(vec![TextToken::Literal("a".into())])]);
        let Primitive::TextBox { paragraphs, .. } = text_box(&list) else {
            unreachable!()
        };
        assert!(paragraphs[0].runs[0].diagnostics.iter().any(|diagnostic| {
            diagnostic.code == "unregistered-font"
                && diagnostic.detail == "unregistered font 'sans-serif'"
        }));
    }

    #[test]
    fn text_runs_append_structured_diagnostics_without_losing_integrity() {
        let mut run = TextRun {
            text: String::new(),
            family: "sans-serif".into(),
            size_in: 1.0 / 6.0,
            bold: false,
            italic: false,
            underline: false,
            small_caps: false,
            superscript: false,
            subscript: false,
            letter_spacing: 0.0,
            case: 0,
            color: "currentColor".into(),
            diagnostics: Vec::new(),
            tab: None,
            diagnosed_face: None,
        };
        append_diagnostic(&mut run, "missing-tab-position", "missing tab position");
        append_diagnostic(&mut run, "unresolvable-text-colour", "unresolvable colour");
        assert_eq!(
            run.diagnostics
                .iter()
                .map(|diagnostic| diagnostic.category)
                .collect::<Vec<_>>(),
            [DiagnosticCategory::Fidelity, DiagnosticCategory::Integrity]
        );
    }

    #[test]
    fn text_diagnostics_preserve_categories_when_a_tab_and_case_fail() {
        let mut shape = text_shape(vec![TextToken::Literal("a".into()), TextToken::Tab(3)]);
        shape.children.push(text_section(
            "Character",
            vec![row(
                0,
                "",
                vec![cell("Color", "not-a-colour"), cell("Case", "9")],
            )],
        ));
        let list = render(vec![shape]);
        let Primitive::TextBox { paragraphs, .. } = text_box(&list) else {
            unreachable!()
        };
        let literal = &paragraphs[0].runs[0].diagnostics;
        assert!(
            literal
                .iter()
                .any(|item| item.category == DiagnosticCategory::Integrity)
        );
        assert!(
            literal
                .iter()
                .any(|item| item.category == DiagnosticCategory::Fidelity)
        );
        let tab = &paragraphs[0].runs[1].diagnostics;
        assert!(
            tab.iter()
                .any(|item| item.code == "unresolvable-text-colour")
        );
        assert!(tab.iter().any(|item| item.code == "missing-tab-position"));
    }

    #[test]
    fn sequential_unavailable_faces_each_record_a_diagnostic() {
        let mut shape = text_shape(vec![
            TextToken::CharacterRun(0),
            TextToken::Literal("a".into()),
            TextToken::CharacterRun(1),
            TextToken::Literal("b".into()),
        ]);
        shape.children.push(text_section(
            "Character",
            vec![
                row(0, "", vec![cell("Font", "99")]),
                row(1, "", vec![cell("Font", "98")]),
            ],
        ));
        let mut renderer = Renderer::default();
        renderer
            .register_font(
                "sans-serif",
                false,
                false,
                include_bytes!("../../ooxml-text/tests/fonts/LiberationSans-Regular.ttf").to_vec(),
            )
            .unwrap();
        let list = renderer.layout_page(&package(vec![shape]), "page").unwrap();
        let Primitive::TextBox { paragraphs, .. } = text_box(&list) else {
            unreachable!()
        };
        let details = paragraphs[0]
            .runs
            .iter()
            .flat_map(|run| &run.diagnostics)
            .map(|diagnostic| diagnostic.detail.as_str())
            .collect::<Vec<_>>();
        assert!(
            details.contains(&"font substituted for '99'"),
            "{details:?}"
        );
        assert!(
            details.contains(&"font substituted for '98'"),
            "{details:?}"
        );
    }

    #[test]
    fn palette_colours_and_unknown_codes_fail_safe() {
        let mut package = package(Vec::new());
        package.colors = serde_json::from_value(serde_json::json!([
            {"name":"ColorEntry","attributes":[["IX","3"],["RGB","010203"]],"children":[]}
        ]))
        .unwrap();
        assert_eq!(palette_colour(&package, 3.0), Some("#010203".into()));
        assert_eq!(
            Diagnostic::for_code("future-diagnostic", "detail").category,
            DiagnosticCategory::Integrity
        );
    }

    #[test]
    fn group_child_extent_origin_maps_to_the_group_lower_left_corner() {
        let matrix = group_transform(
            Bounds {
                x: 10.0,
                y: 20.0,
                width: 8.0,
                height: 12.0,
                ..Default::default()
            },
            Some(ChildCoordinateExtent {
                x: 2.0,
                y: 3.0,
                width: 4.0,
                height: 6.0,
            }),
        );
        assert_point_close(matrix.apply_point(2.0, 3.0), (10.0, 20.0));
        assert_point_close(matrix.apply_point(6.0, 9.0), (18.0, 32.0));
    }

    #[test]
    fn shape_transform_rotates_and_flips_about_its_local_pin() {
        let mut rotated = shape(1, 3.0, 4.0);
        rotated.children.push(ShapeChild::Cell(cell(
            "Angle",
            &std::f64::consts::FRAC_PI_2.to_string(),
        )));
        let list = render(vec![rotated]);
        let Primitive::Shape {
            transform, path, ..
        } = shape_primitive(&list, 1)
        else {
            unreachable!()
        };
        assert_eq!(*transform, Affine::identity());
        assert!(
            matches!(path[1], GeometryPathCommand::Line { x, y } if (x - 3.0).abs() < 1e-9 && (y - 5.0).abs() < 1e-9)
        );

        for (id, flip_x, flip_y) in [(2, true, false), (3, false, true)] {
            let mut flipped = shape(id, 3.0, 4.0);
            flipped.children.push(ShapeChild::Cell(cell(
                "FlipX",
                if flip_x { "1" } else { "0" },
            )));
            flipped.children.push(ShapeChild::Cell(cell(
                "FlipY",
                if flip_y { "1" } else { "0" },
            )));
            let list = render(vec![flipped]);
            let Primitive::Shape {
                path, transform, ..
            } = shape_primitive(&list, id)
            else {
                unreachable!()
            };
            assert_eq!(*transform, Affine::identity());
            let expected = if flip_x { (2.0, 5.0) } else { (4.0, 3.0) };
            assert!(
                matches!(path[2], GeometryPathCommand::Line { x, y } if (x - expected.0).abs() < 1e-9 && (y - expected.1).abs() < 1e-9)
            );
        }
    }

    #[test]
    fn group_composes_translation_rotation_and_flips_for_child_geometry() {
        let child = shape(2, 1.0, 0.0);
        let mut group = shape(1, 10.0, 20.0);
        group
            .children
            .retain(|child| !matches!(child, ShapeChild::Section(_)));
        group.children.extend([
            ShapeChild::Cell(cell("Angle", &std::f64::consts::FRAC_PI_2.to_string())),
            ShapeChild::Cell(cell("FlipX", "1")),
            ShapeChild::Shapes(vec![ShapesChild::Shape(child)]),
        ]);
        let list = render(vec![group]);
        let Primitive::Group { primitives, .. } = &list.primitives[0] else {
            unreachable!()
        };
        let Primitive::Shape {
            path, transform, ..
        } = &primitives[0]
        else {
            unreachable!()
        };
        assert_eq!(*transform, Affine::identity());
        assert!(
            matches!(path[0], GeometryPathCommand::Move { x, y } if (x - 10.0).abs() < 1e-9 && (y - 20.0).abs() < 1e-9)
        );
    }

    #[test]
    fn forty_five_degree_group_keeps_text_in_a_local_box_under_the_group_transform() {
        let mut child = shape(2, 1.0, 2.0);
        child
            .children
            .push(ShapeChild::Text(vec![TextToken::Literal("ab".into())]));
        let mut parent = group(1, 10.0, 20.0, vec![child]);
        with_cell(
            &mut parent,
            "Angle",
            &std::f64::consts::FRAC_PI_4.to_string(),
        );
        let list = render(vec![parent]);
        let Primitive::Group { primitives, .. } = &list.primitives[0] else {
            unreachable!()
        };
        let Primitive::TextBox {
            x,
            y,
            width,
            height,
            lines,
            transform,
            ..
        } = &primitives[1]
        else {
            unreachable!()
        };
        let matrix = Affine {
            a: std::f32::consts::FRAC_1_SQRT_2,
            b: std::f32::consts::FRAC_1_SQRT_2,
            c: -std::f32::consts::FRAC_1_SQRT_2,
            d: std::f32::consts::FRAC_1_SQRT_2,
            e: 10.707_107,
            f: 17.878_68,
        };
        assert_point_close((transform.a, transform.b), (matrix.a, matrix.b));
        assert_point_close((transform.c, transform.d), (matrix.c, matrix.d));
        assert_point_close((transform.e, transform.f), (matrix.e, matrix.f));
        assert_point_close((*x, *y), (1.0, 2.0));
        assert_point_close((*width, *height), (1.0, 1.0));
        assert_point_close(
            transform.apply_point(lines[0].x, lines[0].y),
            matrix.apply_point(1.0, 2.0),
        );
        assert_point_close(
            transform.apply_point(lines[0].caret_stops[1].x, lines[0].caret_stops[1].y),
            matrix.apply_point(1.0 + 0.166_666_67 * 0.5, 2.0),
        );
    }

    #[test]
    fn rotated_text_hit_testing_rejects_axis_aligned_bounds_and_follows_caret_geometry() {
        let mut child = shape(2, 1.0, 2.0);
        child
            .children
            .push(ShapeChild::Text(vec![TextToken::Literal("ab".into())]));
        let mut parent = group(1, 4.0, 4.0, vec![child]);
        with_cell(
            &mut parent,
            "Angle",
            &std::f64::consts::FRAC_PI_4.to_string(),
        );
        let list = render(vec![parent]);
        let Primitive::Group { primitives, .. } = &list.primitives[0] else {
            unreachable!()
        };
        let Primitive::TextBox {
            id,
            x,
            y,
            width,
            height,
            lines,
            transform,
            ..
        } = &primitives[1]
        else {
            unreachable!()
        };
        let canvas = |(x, y): (f32, f32)| hit_test(&list, x * 96.0, 768.0 - y * 96.0);
        let (mut left, mut top, mut bounds_width, mut bounds_height) = (*x, *y, *width, *height);
        transform_rect(
            &mut left,
            &mut top,
            &mut bounds_width,
            &mut bounds_height,
            *transform,
        );
        let outside = (left + bounds_width * 0.05, top + bounds_height * 0.05);
        assert!(
            outside.0 >= left
                && outside.0 <= left + bounds_width
                && outside.1 >= top
                && outside.1 <= top + bounds_height
        );
        assert_eq!(canvas(outside), None);
        let stop = lines[0].caret_stops[1];
        assert_eq!(
            canvas(transform.apply_point(stop.x, stop.y + height * 0.25)),
            Some(HitTestResult::Text {
                shape_id: id.clone(),
                position: stop.position,
            })
        );
    }

    #[test]
    fn standalone_rotated_shape_rotates_its_text_with_its_own_transform() {
        let mut rotated = shape(1, 4.0, 4.0);
        rotated
            .children
            .push(ShapeChild::Text(vec![TextToken::Literal("ab".into())]));
        with_cell(
            &mut rotated,
            "Angle",
            &std::f64::consts::FRAC_PI_4.to_string(),
        );
        let list = render(vec![rotated]);
        let Primitive::TextBox {
            id,
            x,
            y,
            width,
            height,
            lines,
            transform,
            ..
        } = text_box(&list)
        else {
            unreachable!()
        };
        let matrix = Affine {
            a: std::f32::consts::FRAC_1_SQRT_2,
            b: std::f32::consts::FRAC_1_SQRT_2,
            c: -std::f32::consts::FRAC_1_SQRT_2,
            d: std::f32::consts::FRAC_1_SQRT_2,
            e: 4.0,
            f: 4.0 - 4.0 * std::f32::consts::SQRT_2,
        };
        assert_point_close((transform.a, transform.b), (matrix.a, matrix.b));
        assert_point_close((transform.c, transform.d), (matrix.c, matrix.d));
        assert_point_close((transform.e, transform.f), (matrix.e, matrix.f));
        assert_point_close((*x, *y), (4.0, 4.0));
        assert_point_close((*width, *height), (1.0, 1.0));
        assert_point_close(
            transform.apply_point(lines[0].x, lines[0].y),
            matrix.apply_point(4.0, 4.0),
        );

        let canvas = |(x, y): (f32, f32)| hit_test(&list, x * 96.0, 768.0 - y * 96.0);
        let (mut left, mut top, mut bounds_width, mut bounds_height) = (*x, *y, *width, *height);
        transform_rect(
            &mut left,
            &mut top,
            &mut bounds_width,
            &mut bounds_height,
            *transform,
        );
        let outside = (left + bounds_width * 0.05, top + bounds_height * 0.05);
        assert!(
            outside.0 >= left
                && outside.0 <= left + bounds_width
                && outside.1 >= top
                && outside.1 <= top + bounds_height
        );
        assert_eq!(canvas(outside), None);
        let stop = lines[0].caret_stops[1];
        assert_eq!(
            canvas(transform.apply_point(stop.x, stop.y + height * 0.25)),
            Some(HitTestResult::Text {
                shape_id: id.clone(),
                position: stop.position,
            })
        );
    }

    #[test]
    fn forty_five_degree_group_preserves_image_orientation_and_all_corners() {
        let mut image = shape(2, 1.0, 2.0);
        image.children.push(ShapeChild::ForeignData(ForeignData {
            foreign_type: None,
            compression_type: None,
            relationship_id: Some("image".into()),
            other_attrs: vec![],
        }));
        let mut parent = group(1, 10.0, 20.0, vec![image]);
        with_cell(
            &mut parent,
            "Angle",
            &std::f64::consts::FRAC_PI_4.to_string(),
        );
        let mut package = package(vec![parent]);
        package.relationships.insert(
            "page".into(),
            vec![vsdx_parse::Relationship {
                id: "image".into(),
                relationship_type: "image".into(),
                target: "image.png".into(),
                target_mode: Default::default(),
                resolved_target: Some("image.png".into()),
            }],
        );
        package.add_part("image.png", vec![0]);
        let list = Renderer::default().layout_page(&package, "page").unwrap();
        let Primitive::Group { primitives, .. } = &list.primitives[0] else {
            unreachable!()
        };
        let Primitive::Image {
            x,
            y,
            width,
            height,
            transform,
            ..
        } = &primitives[0]
        else {
            unreachable!()
        };
        assert_eq!((*x, *y, *width, *height), (0.0, 0.0, 1.0, 1.0));
        assert!((transform.b - std::f32::consts::FRAC_1_SQRT_2).abs() < 1e-5);
        let corners = [(0.0, 0.0), (1.0, 0.0), (0.0, 1.0), (1.0, 1.0)]
            .map(|(x, y)| transform.apply_point(x, y));
        let expected = [(1.0, 2.0), (2.0, 2.0), (1.0, 3.0), (2.0, 3.0)].map(|(x, y)| {
            Affine {
                a: std::f32::consts::FRAC_1_SQRT_2,
                b: std::f32::consts::FRAC_1_SQRT_2,
                c: -std::f32::consts::FRAC_1_SQRT_2,
                d: std::f32::consts::FRAC_1_SQRT_2,
                e: 10.707_107,
                f: 17.878_68,
            }
            .apply_point(x, y)
        });
        for (actual, expected) in corners.into_iter().zip(expected) {
            assert_point_close(actual, expected);
        }
    }

    #[test]
    fn nested_rotated_flipped_groups_compose_by_matrix_multiplication() {
        let child = shape(3, 1.0, 0.0);
        let mut inner = group(2, 2.0, 3.0, vec![child]);
        with_cell(
            &mut inner,
            "Angle",
            &std::f64::consts::FRAC_PI_2.to_string(),
        );
        with_cell(&mut inner, "FlipX", "1");
        let mut outer = group(1, 10.0, 20.0, vec![inner]);
        with_cell(
            &mut outer,
            "Angle",
            &std::f64::consts::FRAC_PI_2.to_string(),
        );
        with_cell(&mut outer, "FlipY", "1");
        let list = render(vec![outer]);
        let Primitive::Group { primitives, .. } = &list.primitives[0] else {
            unreachable!()
        };
        let Primitive::Group { primitives, .. } = &primitives[0] else {
            unreachable!()
        };
        let Primitive::Shape { path, .. } = &primitives[0] else {
            unreachable!()
        };
        let GeometryPathCommand::Move { x, y } = path[0] else {
            unreachable!()
        };
        let outer = Affine {
            a: 0.0,
            b: 1.0,
            c: 1.0,
            d: 0.0,
            e: 7.0,
            f: 18.0,
        };
        let inner = Affine {
            a: 0.0,
            b: -1.0,
            c: -1.0,
            d: 0.0,
            e: 2.0,
            f: 4.0,
        };
        let expected = outer.compose(inner).apply_point(1.0, 0.0);
        assert_point_close((x as f32, y as f32), expected);
    }

    #[test]
    fn group_non_uniformly_scales_children_before_rotation_and_flip() {
        let child = shape(2, 0.0, 0.0);
        let mut group = group(1, 10.0, 20.0, vec![child]);
        with_cell(&mut group, "Width", "4");
        with_cell(&mut group, "Height", "3");
        with_cell(
            &mut group,
            "Angle",
            &std::f64::consts::FRAC_PI_2.to_string(),
        );
        with_cell(&mut group, "FlipX", "1");
        let list = render(vec![group]);
        let Primitive::Group { primitives, .. } = &list.primitives[0] else {
            unreachable!()
        };
        let Primitive::Shape { path, .. } = &primitives[0] else {
            unreachable!()
        };
        let GeometryPathCommand::Line { x, y } = path[1] else {
            unreachable!()
        };
        assert_point_close((x as f32, y as f32), (10.0, 16.0));
    }

    #[test]
    fn affine_inversion_round_trips_and_degenerate_matrices_are_safe() {
        let matrix = Affine {
            a: -1.5,
            b: -2.0,
            c: -0.5,
            d: 3.0,
            e: 12.0,
            f: -7.0,
        };
        let inverse = matrix.invert().unwrap();
        assert_point_close(
            inverse.apply_point(
                matrix.apply_point(3.25, -8.5).0,
                matrix.apply_point(3.25, -8.5).1,
            ),
            (3.25, -8.5),
        );
        let degenerate = Affine {
            a: 0.0,
            ..Affine::identity()
        };
        assert_eq!(degenerate.invert(), None);
        assert!(!point_in_transformed_rect(
            0.0, 0.0, 1.0, 1.0, degenerate, 0.5, 0.5
        ));
    }

    #[test]
    fn hit_testing_honors_geometry_transform_order_groups_and_empty_canvas() {
        let mut rotated = shape(2, 3.0, 3.0);
        rotated.children.push(ShapeChild::Cell(cell(
            "Angle",
            &std::f64::consts::FRAC_PI_2.to_string(),
        )));
        let group = Shape {
            id: 3,
            name: None,
            name_u: None,
            shape_type: None,
            master: None,
            master_shape: None,
            line_style: None,
            fill_style: None,
            text_style: None,
            del: false,
            other_attrs: vec![],
            children: vec![
                ShapeChild::Cell(cell("Width", "1")),
                ShapeChild::Cell(cell("Height", "1")),
                ShapeChild::Cell(cell("PinX", "6")),
                ShapeChild::Cell(cell("PinY", "1")),
                ShapeChild::Shapes(vec![ShapesChild::Shape(shape(4, 0.0, 0.0))]),
            ],
        };
        let list = render(vec![shape(1, 1.0, 1.0), rotated, shape(5, 1.0, 1.0), group]);
        assert_eq!(
            hit_test(&list, 96.0 * 1.5, 768.0 - 96.0 * 1.5),
            Some(HitTestResult::Shape {
                shape_id: "page:5".into()
            })
        );
        assert_eq!(hit_test(&list, 96.0 * 2.1, 768.0 - 96.0 * 2.1), None);
        assert_eq!(
            hit_test(&list, 96.0 * 5.5, 768.0 - 96.0 * 0.5),
            Some(HitTestResult::Shape {
                shape_id: "page:4".into()
            })
        );
        assert_eq!(hit_test(&list, 1.0, 1.0), None);
    }

    #[test]
    fn hit_testing_uses_rotated_quads_curves_strokes_and_z_order() {
        let canvas = VsdxDisplayList {
            contract_version: CONTRACT_VERSION,
            width: 100.0,
            height: 100.0,
            paint_transform: final_paint_transform(1.0),
            primitives: vec![
                Primitive::Shape {
                    id: "bottom".into(),
                    z_order: 10,
                    path: vec![
                        GeometryPathCommand::Move { x: 0.0, y: 0.0 },
                        GeometryPathCommand::Line { x: 1.0, y: 0.0 },
                        GeometryPathCommand::Line { x: 1.0, y: 1.0 },
                        GeometryPathCommand::Line { x: 0.0, y: 1.0 },
                    ],
                    fill: Some(Paint::Solid {
                        color: "#000".into(),
                    }),
                    stroke: None,
                    transform: Affine::identity(),
                },
                Primitive::Shape {
                    id: "top".into(),
                    z_order: 20,
                    path: vec![
                        GeometryPathCommand::Move { x: 0.0, y: 0.0 },
                        GeometryPathCommand::Line { x: 1.0, y: 0.0 },
                        GeometryPathCommand::Line { x: 1.0, y: 1.0 },
                        GeometryPathCommand::Line { x: 0.0, y: 1.0 },
                    ],
                    fill: Some(Paint::Solid {
                        color: "#000".into(),
                    }),
                    stroke: None,
                    transform: Affine::identity(),
                },
                Primitive::Image {
                    id: "rotated-image".into(),
                    z_order: 1,
                    asset_id: "i".into(),
                    x: 0.0,
                    y: 0.0,
                    width: 2.0,
                    height: 2.0,
                    transform: Affine {
                        a: std::f32::consts::FRAC_1_SQRT_2,
                        b: std::f32::consts::FRAC_1_SQRT_2,
                        c: -std::f32::consts::FRAC_1_SQRT_2,
                        d: std::f32::consts::FRAC_1_SQRT_2,
                        e: 5.0,
                        f: 0.0,
                    },
                },
                Primitive::Shape {
                    id: "curve".into(),
                    z_order: 2,
                    path: vec![
                        GeometryPathCommand::Move { x: 7.0, y: 0.0 },
                        GeometryPathCommand::Cubic {
                            cp1x: 9.0,
                            cp1y: 0.0,
                            cp2x: 9.0,
                            cp2y: 2.0,
                            x: 7.0,
                            y: 2.0,
                        },
                        GeometryPathCommand::Cubic {
                            cp1x: 5.0,
                            cp1y: 2.0,
                            cp2x: 5.0,
                            cp2y: 0.0,
                            x: 7.0,
                            y: 0.0,
                        },
                    ],
                    fill: Some(Paint::Solid {
                        color: "#000".into(),
                    }),
                    stroke: None,
                    transform: Affine::identity(),
                },
                Primitive::Shape {
                    id: "stroke".into(),
                    z_order: 3,
                    path: vec![
                        GeometryPathCommand::Move { x: 0.0, y: 3.0 },
                        GeometryPathCommand::Line { x: 2.0, y: 3.0 },
                    ],
                    fill: None,
                    stroke: Some(Stroke {
                        color: "#000".into(),
                        width: 0.2,
                        dashed: false,
                    }),
                    transform: Affine::identity(),
                },
            ],
        };
        let hit = |x: f32, y: f32| hit_test(&canvas, x * 96.0, 96.0 - y * 96.0);
        assert_eq!(
            hit(0.5, 0.5),
            Some(HitTestResult::Shape {
                shape_id: "top".into()
            })
        );
        assert_eq!(
            hit(5.0, 1.0),
            Some(HitTestResult::Shape {
                shape_id: "rotated-image".into()
            })
        );
        assert_eq!(hit(3.7, 0.1), None);
        assert_eq!(
            hit(7.0, 1.0),
            Some(HitTestResult::Shape {
                shape_id: "curve".into()
            })
        );
        assert_eq!(hit(9.5, 1.0), None);
        assert_eq!(
            hit(1.0, 3.08),
            Some(HitTestResult::Shape {
                shape_id: "stroke".into()
            })
        );
    }

    #[test]
    fn hit_testing_returns_a_rotated_group_child() {
        let child = shape(2, 1.0, 0.0);
        let mut parent = group(1, 5.0, 5.0, vec![child]);
        with_cell(
            &mut parent,
            "Angle",
            &std::f64::consts::FRAC_PI_2.to_string(),
        );
        let list = render(vec![parent]);
        assert_eq!(
            hit_test(&list, 96.0 * 4.5, 768.0 - 96.0 * 5.5),
            Some(HitTestResult::Shape {
                shape_id: "page:2".into()
            })
        );
    }

    #[test]
    fn paint_images_visibility_and_z_order_follow_the_render_contract() {
        let mut hidden = shape(1, 1.0, 1.0);
        hidden.children.push(ShapeChild::Cell(cell("NoShow", "1")));
        let mut deleted = shape(2, 2.0, 1.0);
        deleted.del = true;
        let mut text = shape(3, 3.0, 1.0);
        text.children
            .push(ShapeChild::Text(vec![TextToken::Literal("text".into())]));
        let mut image = shape(4, 4.0, 1.0);
        image.children.push(ShapeChild::ForeignData(ForeignData {
            foreign_type: Some("Bitmap".into()),
            compression_type: None,
            relationship_id: Some("image".into()),
            other_attrs: vec![],
        }));
        let mut package = package(vec![hidden, deleted, text, image]);
        package.relationships.insert(
            "page".into(),
            vec![vsdx_parse::Relationship {
                id: "image".into(),
                relationship_type: "image".into(),
                target: "media/image1.png".into(),
                target_mode: Default::default(),
                resolved_target: Some("visio/media/image1.png".into()),
            }],
        );
        package.add_part("visio/media/image1.png", vec![0]);
        let list = Renderer::default().layout_page(&package, "page").unwrap();
        assert!(
            matches!(&list.primitives[0], Primitive::Shape { fill: Some(Paint::Solid { color }), stroke: Some(Stroke { color: line, width, dashed: false }), .. } if color == "#010203" && line == "#040506" && *width == 0.02)
        );
        assert!(matches!(&list.primitives[1], Primitive::TextBox { .. }));
        assert!(
            matches!(&list.primitives[2], Primitive::Image { asset_id, .. } if asset_id == "visio/media/image1.png")
        );
        let z_orders = list
            .primitives
            .iter()
            .map(|primitive| match primitive {
                Primitive::Shape { z_order, .. }
                | Primitive::Image { z_order, .. }
                | Primitive::TextBox { z_order, .. }
                | Primitive::Placeholder { z_order, .. }
                | Primitive::Group { z_order, .. } => *z_order,
            })
            .collect::<std::collections::BTreeSet<_>>();
        assert_eq!(z_orders.len(), 3);
        assert_eq!(z_orders.into_iter().collect::<Vec<_>>(), vec![2, 3, 4]);
    }

    #[test]
    fn unresolved_foreign_data_and_colours_become_placeholders() {
        let mut image = shape(1, 1.0, 1.0);
        image.children.push(ShapeChild::ForeignData(ForeignData {
            foreign_type: None,
            compression_type: None,
            relationship_id: Some("missing".into()),
            other_attrs: vec![],
        }));
        let mut colour = shape(2, 2.0, 1.0);
        colour.children.retain(
            |child| !matches!(child, ShapeChild::Cell(Cell { name, .. }) if name == "FillForegnd"),
        );
        let list = render(vec![image, colour]);
        assert!(
            matches!(&list.primitives[0], Primitive::Placeholder { reason, .. } if reason == "unsupported ForeignData image")
        );
        assert!(
            matches!(&list.primitives[1], Primitive::Placeholder { reason, .. } if reason.starts_with("unresolvable colour:"))
        );
    }

    #[test]
    fn dangling_media_and_non_finite_stroke_width_become_placeholders() {
        let mut image = shape(1, 1.0, 1.0);
        image.children.push(ShapeChild::ForeignData(ForeignData {
            foreign_type: None,
            compression_type: None,
            relationship_id: Some("image".into()),
            other_attrs: vec![],
        }));
        let mut line = shape(2, 2.0, 1.0);
        with_cell(&mut line, "LineWeight", "1e100");
        let mut package = package(vec![image, line]);
        package.relationships.insert(
            "page".into(),
            vec![vsdx_parse::Relationship {
                id: "image".into(),
                relationship_type: "image".into(),
                target: "missing.png".into(),
                target_mode: Default::default(),
                resolved_target: Some("missing.png".into()),
            }],
        );
        let list = Renderer::default().layout_page(&package, "page").unwrap();
        assert!(
            matches!(&list.primitives[0], Primitive::Placeholder { reason, .. } if reason.contains("dangling ForeignData image target"))
        );
        assert!(
            matches!(&list.primitives[1], Primitive::Placeholder { reason, .. } if reason.contains("non-finite stroke width"))
        );
    }

    #[test]
    fn unsupported_colour_reason_preserves_evaluator_detail() {
        let mut coloured = shape(1, 1.0, 1.0);
        coloured.children.retain(
            |child| !matches!(child, ShapeChild::Cell(Cell { name, .. }) if name == "FillForegnd"),
        );
        coloured
            .children
            .push(ShapeChild::Cell(formula("FillForegnd", "THEMEVAL(999)")));
        let list = render(vec![coloured]);
        assert!(
            matches!(&list.primitives[0], Primitive::Placeholder { reason, .. } if reason.starts_with("unresolvable colour:") && reason.len() > "unresolvable colour:".len())
        );
    }

    #[test]
    fn invalid_page_dimensions_are_rejected_before_pixel_conversion() {
        for (width, height) in [("0", "8"), ("-1", "8"), ("1e100", "8")] {
            let mut page = package(vec![]);
            let sheet = page.page_sheets.get_mut(&1).unwrap();
            sheet.children = vec![
                SheetChild::Cell(cell("PageWidth", width)),
                SheetChild::Cell(cell("PageHeight", height)),
            ];
            assert!(
                matches!(
                    Renderer::default().layout_page(&page, "page"),
                    Err(RenderError::PageDimensions(_))
                ),
                "{width} x {height}"
            );
        }
    }

    #[test]
    fn non_finite_text_run_size_and_line_height_are_rejected() {
        let mut text = shape(1, 1.0, 1.0);
        with_cell(&mut text, "Char.Size", "1e100");
        text.children
            .push(ShapeChild::Text(vec![TextToken::Literal("x".into())]));
        assert!(matches!(
            Renderer::default().layout_page(&package(vec![text]), "page"),
            Err(RenderError::PageDimensions(_))
        ));
    }

    #[test]
    fn every_renderer_budget_rejects_cleanly_and_font_rejection_is_atomic() {
        let page = package(vec![shape(1, 1.0, 1.0)]);
        let mut limits = RenderLimits {
            max_shapes: 0,
            ..RenderLimits::default()
        };
        assert!(matches!(
            Renderer::new(limits.clone()).layout_page(&page, "page"),
            Err(RenderError::Budget("shapes"))
        ));
        let mut text_shape = shape(1, 1.0, 1.0);
        text_shape
            .children
            .push(ShapeChild::Text(vec![TextToken::Literal("x".into())]));
        let text_page = package(vec![text_shape]);
        for (name, set) in [
            ("text bytes", 0usize),
            ("text paragraphs", 0),
            ("text lines", 0),
            ("text runs", 0),
        ] {
            limits = RenderLimits::default();
            match name {
                "text bytes" => limits.max_text_bytes = set,
                "text paragraphs" => limits.max_text_paragraphs = set,
                "text lines" => limits.max_text_lines = set,
                _ => limits.max_text_runs = set,
            }
            assert!(
                matches!(Renderer::new(limits).layout_page(&text_page, "page"), Err(RenderError::Budget(actual)) if actual == name)
            );
        }
        limits = RenderLimits {
            max_display_list_bytes: 0,
            ..RenderLimits::default()
        };
        assert!(matches!(
            Renderer::new(limits).layout_page(&page, "page"),
            Err(RenderError::Budget("display-list bytes"))
        ));
        for limits in [
            RenderLimits {
                max_fonts: 0,
                ..RenderLimits::default()
            },
            RenderLimits {
                max_font_bytes: 0,
                ..RenderLimits::default()
            },
        ] {
            let mut renderer = Renderer::new(limits);
            assert!(renderer.register_font("x", false, false, vec![1]).is_err());
            assert_eq!(renderer.font_bytes, 0);
        }
    }

    #[test]
    fn pathological_group_nesting_is_placeholdered_at_the_depth_limit() {
        let mut nested = shape(65, 0.0, 0.0);
        for id in (1..65).rev() {
            let mut group = shape(id, 0.0, 0.0);
            group
                .children
                .retain(|child| !matches!(child, ShapeChild::Section(_)));
            group
                .children
                .push(ShapeChild::Shapes(vec![ShapesChild::Shape(nested)]));
            nested = group;
        }
        let list = render(vec![nested]);
        fn reasons(primitives: &[Primitive]) -> Vec<&str> {
            primitives
                .iter()
                .flat_map(|primitive| match primitive {
                    Primitive::Placeholder { reason, .. } => vec![reason.as_str()],
                    Primitive::Group { primitives, .. } => reasons(primitives),
                    _ => vec![],
                })
                .collect()
        }
        assert_eq!(reasons(&list.primitives), ["group nesting depth exceeded"]);
    }

    #[test]
    fn overflowing_coordinates_become_a_finite_placeholder() {
        let mut overflowing = shape(1, 1.0, 1.0);
        overflowing.children.retain(
            |child| !matches!(child, ShapeChild::Cell(Cell { name, .. }) if name == "PinX"),
        );
        overflowing
            .children
            .push(ShapeChild::Cell(cell("PinX", "1e308")));
        let list = render(vec![overflowing]);
        assert!(matches!(
            list.primitives.as_slice(),
            [Primitive::Placeholder { .. }]
        ));
        assert!(display_list_finite(&list));
    }
    #[test]
    fn rejects_unknown_contract() {
        let list = VsdxDisplayList {
            contract_version: CONTRACT_VERSION + 1,
            width: 0.0,
            height: 0.0,
            paint_transform: final_paint_transform(0.0),
            primitives: vec![],
        };
        assert!(list.validate().is_err());
    }

    #[test]
    fn display_list_decode_rejects_unsupported_contract() {
        for version in [2, CONTRACT_VERSION - 1] {
            let payload = format!(
                r#"{{"contractVersion":{version},"width":0.0,"height":0.0,"paintTransform":{{"a":1.0,"b":0.0,"c":0.0,"d":1.0,"e":0.0,"f":0.0}},"primitives":[]}}"#
            );
            let error = serde_json::from_str::<VsdxDisplayList>(&payload).unwrap_err();
            assert!(
                error.to_string().contains(&format!(
                    "unsupported VSDX display-list contract version {version}"
                )),
                "{error}"
            );
        }
    }

    #[test]
    fn decoded_diagnostics_own_codes_and_honour_wire_categories() {
        let diagnostics = (0..10_000)
            .map(|index| {
                serde_json::from_value::<Diagnostic>(serde_json::json!({
                    "category": "fidelity",
                    "code": format!("unknown-{index}"),
                    "detail": "fixture",
                }))
                .unwrap()
            })
            .collect::<Vec<_>>();
        let _: String = diagnostics[0].code.clone();
        assert_eq!(diagnostics[9_999].code, "unknown-9999");
        assert!(diagnostics.iter().all(|diagnostic| {
            diagnostic.category == DiagnosticCategory::Fidelity && !diagnostic.category_defaulted
        }));
    }

    #[test]
    fn decoded_diagnostic_missing_or_invalid_category_defaults_to_integrity() {
        for payload in [
            serde_json::json!({"code":"unknown","detail":"fixture"}),
            serde_json::json!({"category":"future","code":"unknown","detail":"fixture"}),
        ] {
            let diagnostic = serde_json::from_value::<Diagnostic>(payload).unwrap();
            assert_eq!(diagnostic.category, DiagnosticCategory::Integrity);
            assert!(diagnostic.category_defaulted);
        }
    }

    #[test]
    fn closed_path_stroke_hits_its_implicit_closing_segment() {
        let path = vec![
            GeometryPathCommand::Move { x: 0.0, y: 0.0 },
            GeometryPathCommand::Line { x: 1.0, y: 0.0 },
            GeometryPathCommand::Line { x: 1.0, y: 1.0 },
            GeometryPathCommand::Close,
        ];
        assert!(path_hit(&path, Affine::identity(), false, 0.1, 0.5, 0.5));
    }

    #[test]
    fn foundation_display_list_matches_golden_contract() {
        let package = vsdx_parse::parse_vsdx(include_bytes!(
            "../../vsdx-parse/tests/fixtures/foundation.vsdx"
        ))
        .unwrap();
        let list = Renderer::default()
            .layout_page(&package, &package.page_part_paths[0])
            .unwrap();
        let actual = serde_json::to_value(list).unwrap();
        let expected: serde_json::Value =
            serde_json::from_str(include_str!("../tests/golden/foundation.json")).unwrap();
        assert_eq!(actual, expected);
    }

    #[test]
    fn nested_groups_fixture_preserves_scaled_affine_content_and_replay_order() {
        let package = vsdx_parse::parse_vsdx(include_bytes!(
            "../../vsdx-parse/tests/fixtures/nested-groups.vsdx"
        ))
        .unwrap();
        let list = Renderer::default()
            .layout_page(&package, &package.page_part_paths[0])
            .unwrap();
        let Primitive::Group {
            primitives: outer, ..
        } = &list.primitives[0]
        else {
            unreachable!()
        };
        let Primitive::Group {
            primitives: inner, ..
        } = &outer[0]
        else {
            unreachable!()
        };
        assert!(matches!(&inner[0], Primitive::Shape { id, .. } if id.ends_with(":3")));
        assert!(
            matches!(&inner[1], Primitive::TextBox { id, lines, .. } if id.ends_with(":3") && lines.len() == 2)
        );
        assert!(
            matches!(&inner[2], Primitive::Image { id, asset_id, .. } if id.ends_with(":4") && asset_id == "visio/media/image1.png")
        );
        assert!(
            matches!(&inner[3], Primitive::Placeholder { id, reason, .. } if id.ends_with(":5") && reason.starts_with("unsupported geometry:"))
        );
        let Primitive::Shape { path, .. } = &inner[0] else {
            unreachable!()
        };
        let GeometryPathCommand::Move { x, y } = path[0] else {
            unreachable!()
        };
        // These values were calculated by hand from the composed outer and inner affine,
        // M = [[-0.94190204, 1.8844845, 10.021151], [-1.1970047, 0.27151108, 11.018823]].
        // Each expected AABB is min/max(M(corner)) for the original source rectangle;
        // the image corners and text caret are direct substitutions into that same affine.
        assert_point_close((x as f32, y as f32), (10.021151, 11.018823));
        let Primitive::TextBox {
            x,
            y,
            width,
            height,
            lines,
            transform,
            ..
        } = &inner[1]
        else {
            unreachable!()
        };
        let (mut left, mut top, mut text_width, mut text_height) = (*x, *y, *width, *height);
        transform_rect(
            &mut left,
            &mut top,
            &mut text_width,
            &mut text_height,
            *transform,
        );
        assert_point_close((left, top), (9.079249, 9.821818));
        assert_point_close((text_width, text_height), (2.8263865, 1.4685154));
        assert_point_close(
            transform.apply_point(lines[0].x, lines[0].y),
            (10.021151, 11.018823),
        );
        assert_point_close(
            transform.apply_point(lines[0].caret_stops[1].x, lines[0].caret_stops[1].y),
            (9.942658, 10.919072),
        );
        let Primitive::Image {
            x,
            y,
            width,
            height,
            transform,
            ..
        } = &inner[2]
        else {
            unreachable!()
        };
        let corners = [
            transform.apply_point(*x, *y),
            transform.apply_point(*x + *width, *y),
            transform.apply_point(*x, *y + *height),
            transform.apply_point(*x + *width, *y + *height),
        ];
        assert_point_close(corners[0], (10.0218315, 8.896324));
        assert_point_close(corners[1], (9.079929, 7.6993194));
        assert_point_close(corners[2], (13.7908, 9.439346));
        assert_point_close(corners[3], (12.848899, 8.242341));
        let Primitive::Placeholder {
            x,
            y,
            width,
            height,
            ..
        } = &inner[3]
        else {
            unreachable!()
        };
        assert_point_close((*x, *y), (13.7908, 9.439346));
        assert_point_close((*width, *height), (2.8263874, 1.4685163));
        let z_orders = inner.iter().map(z_order).collect::<Vec<_>>();
        assert_eq!(z_orders, vec![2, 3, 4, 5]);
        assert_eq!(
            hit_test(&list, 10.021151 * 96.0, (11.0 - 11.018823) * 96.0),
            Some(HitTestResult::Text {
                shape_id: "visio/pages/page1.xml:3".into(),
                position: 0,
            })
        );
    }

    #[test]
    fn corpus_smoke_reports_painted_and_placeholdered_shapes() {
        let Ok(directory) = std::env::var("VSDX_CORPUS_DIR") else {
            eprintln!(
                "warning: skipping VSDX corpus renderer smoke test; VSDX_CORPUS_DIR is unset"
            );
            return;
        };
        let mut painted = 0usize;
        let mut placeholders = 0usize;
        let mut groups = 0usize;
        let mut reasons = BTreeMap::new();
        let mut text = TextCorpusStats::default();
        let mut renderer = Renderer::default();
        renderer
            .register_font(
                "sans-serif",
                false,
                false,
                include_bytes!("../../ooxml-text/tests/fonts/LiberationSans-Regular.ttf").to_vec(),
            )
            .unwrap();
        for file in ["lichtsysteme.vsdx", "soundplan.vsdx"] {
            let path = std::path::Path::new(&directory).join(file);
            let package = vsdx_parse::parse_vsdx(&std::fs::read(path).unwrap()).unwrap();
            for page in &package.page_part_paths {
                let list = renderer.layout_page(&package, page).unwrap();
                let json = serde_json::to_string(&list).unwrap();
                assert!(!json.contains("NaN") && !json.contains("Infinity"));
                let expected = package.page_contents[page]
                    .shapes()
                    .flat_map(|shape| shape_ids(page, shape))
                    .collect::<std::collections::BTreeSet<_>>();
                let mut page_shapes = std::collections::BTreeSet::new();
                let mut page_text = std::collections::BTreeSet::new();
                let mut page_images = std::collections::BTreeSet::new();
                let mut page_placeholders = std::collections::BTreeSet::new();
                let mut page_groups = std::collections::BTreeSet::new();
                let mut image_shapes = std::collections::BTreeSet::new();
                expected_subcontent(
                    package.page_contents[page].shapes(),
                    page,
                    &mut image_shapes,
                );
                count_primitives(
                    &list.primitives,
                    &mut page_shapes,
                    &mut page_text,
                    &mut page_images,
                    &mut page_placeholders,
                    &mut page_groups,
                    &mut reasons,
                );
                assert!(page_shapes.is_disjoint(&page_placeholders));
                assert!(page_shapes.is_disjoint(&page_groups));
                assert!(page_placeholders.is_disjoint(&page_groups));
                let actual = page_shapes
                    .union(&page_text)
                    .cloned()
                    .collect::<std::collections::BTreeSet<_>>();
                let actual = actual
                    .union(&page_images)
                    .cloned()
                    .collect::<std::collections::BTreeSet<_>>();
                let actual = actual
                    .union(&page_placeholders)
                    .cloned()
                    .collect::<std::collections::BTreeSet<_>>();
                let actual = actual
                    .union(&page_groups)
                    .cloned()
                    .collect::<std::collections::BTreeSet<_>>();
                assert_eq!(actual, expected);
                assert_eq!(page_images, image_shapes);
                let mut resolved_text = std::collections::BTreeSet::new();
                assert_resolved_text(
                    &renderer,
                    &package,
                    page,
                    &list.primitives,
                    &mut resolved_text,
                    &mut text,
                );
                assert_eq!(page_text, resolved_text);
                painted += page_shapes.len() + page_text.len() + page_images.len();
                placeholders += page_placeholders.len();
                groups += page_groups.len();
            }
        }
        eprintln!(
            "VSDX corpus render: painted={painted} placeholdered={placeholders} group={groups} placeholder reasons={reasons:?}",
        );
        eprintln!(
            "VSDX corpus text: shapes={} zero-integrity-diagnostics={} integrity diagnostics={} integrity reasons={:?} fidelity diagnostics={} fidelity reasons={:?} paragraphs={} runs={} marker-only={} field-only={} style-inherited={} master-inherited={}",
            text.shapes,
            text.zero_integrity_diagnostics,
            text.integrity_diagnostics,
            text.integrity_reasons,
            text.fidelity_diagnostics,
            text.fidelity_reasons,
            text.paragraphs,
            text.runs,
            text.marker_only,
            text.field_only,
            text.style_inherited,
            text.master_inherited,
        );
        assert_eq!(text.integrity_diagnostics, 0, "text-integrity diagnostics");
    }

    #[test]
    fn text_accounting_fixture_covers_resolved_text_sources() {
        let package = vsdx_parse::parse_vsdx(include_bytes!(
            "../../vsdx-parse/tests/fixtures/text-accounting.vsdx"
        ))
        .unwrap();
        let page = &package.page_part_paths[0];
        let list = Renderer::default().layout_page(&package, page).unwrap();
        let mut rendered = std::collections::BTreeSet::new();
        let mut stats = TextCorpusStats::default();
        assert_resolved_text(
            &Renderer::default(),
            &package,
            page,
            &list.primitives,
            &mut rendered,
            &mut stats,
        );
        assert_eq!(stats.marker_only, 1);
        assert_eq!(stats.field_only, 1);
        assert_eq!(stats.style_inherited, 1);
        assert_eq!(stats.master_inherited, 1);
        let mut text_by_id = BTreeMap::new();
        text_boxes_by_id(&list.primitives, &mut text_by_id);
        let text = |id: &str| match text_by_id[&format!("{page}:{id}")] {
            Primitive::TextBox { paragraphs, .. } => paragraphs,
            _ => unreachable!(),
        };
        assert_eq!(text("4")[0].runs[0].text, "master text");
        assert_eq!(text("2")[0].runs[0].text, "field value");
        let inherited = &text("3")[0].runs[0];
        assert_eq!(inherited.family, "Calibri");
        assert_eq!(inherited.size_in, 0.25);
        assert_eq!(inherited.color, "#010203");
    }

    fn accounting_oracle_fixture() -> (VsdxDisplayList, Vec<Vec<ExpectedTextRun>>) {
        let run = |text: &str, category| TextRun {
            text: text.into(),
            family: "sans-serif".into(),
            size_in: 1.0 / 6.0,
            bold: false,
            italic: false,
            underline: false,
            small_caps: false,
            superscript: false,
            subscript: false,
            letter_spacing: 0.0,
            case: 0,
            color: "currentColor".into(),
            diagnostics: vec![Diagnostic::new(category, "fixture", "fixture")],
            tab: None,
            diagnosed_face: None,
        };
        let paragraphs = vec![
            TextParagraph {
                runs: vec![
                    run("first", DiagnosticCategory::Integrity),
                    run("second", DiagnosticCategory::Fidelity),
                ],
            },
            TextParagraph {
                runs: vec![run("third", DiagnosticCategory::Integrity)],
            },
        ];
        let expected = paragraphs
            .iter()
            .map(|paragraph| {
                paragraph
                    .runs
                    .iter()
                    .map(|run| ExpectedTextRun {
                        text: run.text.clone(),
                        diagnostic_categories: run
                            .diagnostics
                            .iter()
                            .map(|item| item.category)
                            .collect(),
                    })
                    .collect()
            })
            .collect();
        (
            VsdxDisplayList {
                contract_version: CONTRACT_VERSION,
                width: 0.0,
                height: 0.0,
                paint_transform: final_paint_transform(0.0),
                primitives: vec![Primitive::TextBox {
                    id: "fixture".into(),
                    z_order: 0,
                    x: 0.0,
                    y: 0.0,
                    width: 0.0,
                    height: 0.0,
                    paragraphs,
                    lines: Vec::new(),
                    transform: Affine::identity(),
                }],
            },
            expected,
        )
    }

    fn fixture_paragraphs(list: &VsdxDisplayList) -> &[TextParagraph] {
        let Primitive::TextBox { paragraphs, .. } = &list.primitives[0] else {
            unreachable!()
        };
        paragraphs
    }

    fn fixture_paragraphs_mut(list: &mut VsdxDisplayList) -> &mut Vec<TextParagraph> {
        let Primitive::TextBox { paragraphs, .. } = &mut list.primitives[0] else {
            unreachable!()
        };
        paragraphs
    }

    #[test]
    fn accounting_oracle_rejects_a_dropped_run() {
        let (mut list, expected) = accounting_oracle_fixture();
        fixture_paragraphs_mut(&mut list)[0].runs.pop();
        assert!(
            assert_display_text_matches("fixture", fixture_paragraphs(&list), &expected).is_err()
        );
    }

    #[test]
    fn accounting_oracle_rejects_a_dropped_paragraph() {
        let (mut list, expected) = accounting_oracle_fixture();
        fixture_paragraphs_mut(&mut list).pop();
        assert!(
            assert_display_text_matches("fixture", fixture_paragraphs(&list), &expected).is_err()
        );
    }

    #[test]
    fn accounting_oracle_rejects_reordered_runs() {
        let (mut list, expected) = accounting_oracle_fixture();
        fixture_paragraphs_mut(&mut list)[0].runs.swap(0, 1);
        assert!(
            assert_display_text_matches("fixture", fixture_paragraphs(&list), &expected).is_err()
        );
    }

    #[test]
    fn accounting_oracle_rejects_a_swapped_diagnostic_category() {
        let (mut list, expected) = accounting_oracle_fixture();
        fixture_paragraphs_mut(&mut list)[0].runs[0].diagnostics[0].category =
            DiagnosticCategory::Fidelity;
        assert!(
            assert_display_text_matches("fixture", fixture_paragraphs(&list), &expected).is_err()
        );
    }

    fn count_primitives(
        primitives: &[Primitive],
        shapes: &mut std::collections::BTreeSet<String>,
        text: &mut std::collections::BTreeSet<String>,
        images: &mut std::collections::BTreeSet<String>,
        placeholders: &mut std::collections::BTreeSet<String>,
        groups: &mut std::collections::BTreeSet<String>,
        reasons: &mut BTreeMap<String, usize>,
    ) {
        for primitive in primitives {
            match primitive {
                Primitive::Shape { id, .. } => {
                    shapes.insert(id.clone());
                }
                Primitive::TextBox { id, .. } => {
                    text.insert(id.clone());
                }
                Primitive::Image { id, .. } => {
                    images.insert(id.clone());
                }
                Primitive::Placeholder { id, reason, .. } => {
                    placeholders.insert(id.clone());
                    *reasons.entry(reason.clone()).or_default() += 1;
                }
                Primitive::Group { id, primitives, .. } => {
                    groups.insert(id.clone());
                    count_primitives(
                        primitives,
                        shapes,
                        text,
                        images,
                        placeholders,
                        groups,
                        reasons,
                    )
                }
            }
        }
    }

    fn shape_ids(page: &str, shape: &Shape) -> Vec<String> {
        std::iter::once(format!("{page}:{}", shape.id))
            .chain(shape.shapes().flat_map(|child| shape_ids(page, child)))
            .collect()
    }

    fn expected_subcontent<'a>(
        shapes: impl Iterator<Item = &'a Shape>,
        page: &str,
        images: &mut std::collections::BTreeSet<String>,
    ) {
        for shape in shapes {
            let id = format!("{page}:{}", shape.id);
            if shape.foreign_data().is_some() {
                images.insert(id);
            }
            expected_subcontent(shape.shapes(), page, images);
        }
    }

    fn assert_resolved_text(
        renderer: &Renderer,
        package: &VsdxPackage,
        page_part: &str,
        primitives: &[Primitive],
        rendered: &mut std::collections::BTreeSet<String>,
        stats: &mut TextCorpusStats,
    ) {
        let contents = &package.page_contents[page_part];
        let page = package
            .page_part_ids
            .get(page_part)
            .and_then(|id| package.page_sheets.get(id))
            .unwrap_or(contents);
        let resolver = Resolver::new(package);
        let references = PageShapeReferences::new(&resolver, page_part).ok();
        let mut text_boxes = BTreeMap::new();
        text_boxes_by_id(primitives, &mut text_boxes);
        for shape in contents.shapes() {
            assert_shape_text(
                renderer,
                package,
                &resolver,
                references.as_ref(),
                page,
                page_part,
                shape,
                &text_boxes,
                rendered,
                stats,
            );
        }
    }

    #[allow(clippy::too_many_arguments)]
    fn assert_shape_text(
        renderer: &Renderer,
        package: &VsdxPackage,
        resolver: &Resolver<'_>,
        references: Option<&PageShapeReferences>,
        page: &Sheet,
        page_part: &str,
        shape: &Shape,
        text_boxes: &BTreeMap<String, &Primitive>,
        rendered: &mut std::collections::BTreeSet<String>,
        stats: &mut TextCorpusStats,
    ) {
        let id = format!("{page_part}:{}", shape.id);
        let resolved = resolver.resolve_shape(page_part, shape.id).unwrap();
        if resolved.deleted || paint::number(&resolved, "NoShow").is_some_and(|value| value != 0.0)
        {
            return;
        }
        let tokens = resolver
            .resolve_text_in_context(shape, page, &resolved)
            .unwrap();
        if tokens.iter().all(|token| {
            matches!(
                token,
                vsdx_resolve::ResolvedTextToken::CharacterRun { .. }
                    | vsdx_resolve::ResolvedTextToken::ParagraphRun { .. }
            )
        }) && !tokens.is_empty()
        {
            stats.marker_only += 1;
        }
        if tokens.iter().all(|token| {
            matches!(
                token,
                vsdx_resolve::ResolvedTextToken::CharacterRun { .. }
                    | vsdx_resolve::ResolvedTextToken::ParagraphRun { .. }
                    | vsdx_resolve::ResolvedTextToken::Field { .. }
            )
        }) && tokens
            .iter()
            .any(|token| matches!(token, vsdx_resolve::ResolvedTextToken::Field { .. }))
        {
            stats.field_only += 1;
        }
        if !tokens.is_empty() && shape.text().is_none() {
            stats.master_inherited += 1;
        }
        if tokens.iter().any(|token| match token {
            vsdx_resolve::ResolvedTextToken::CharacterRun { properties, .. }
            | vsdx_resolve::ResolvedTextToken::ParagraphRun { properties, .. }
            | vsdx_resolve::ResolvedTextToken::Tab { properties, .. }
            | vsdx_resolve::ResolvedTextToken::Field { properties, .. } => properties.values().any(
                |value| matches!(value, Lookup::Found(value) if value.provenance == vsdx_resolve::Provenance::StyleText),
            ),
            vsdx_resolve::ResolvedTextToken::Literal(_) => false,
        }) {
            stats.style_inherited += 1;
        }
        let expected =
            expected_text_runs(renderer, package, references, &resolved, shape.id, &tokens);
        let expected_runs = expected.iter().map(Vec::len).sum::<usize>();
        if expected_runs > 0 {
            stats.paragraphs += expected.len();
            stats.runs += expected_runs;
        }
        match text_boxes.get(&id) {
            Some(Primitive::TextBox { paragraphs, .. }) => {
                rendered.insert(id.clone());
                assert_display_text_matches(&id, paragraphs, &expected).unwrap();
                stats.shapes += 1;
                let integrity_diagnostics = paragraphs
                    .iter()
                    .flat_map(|paragraph| &paragraph.runs)
                    .flat_map(|run| &run.diagnostics)
                    .filter(|diagnostic| diagnostic.category == DiagnosticCategory::Integrity)
                    .count();
                if integrity_diagnostics == 0 {
                    stats.zero_integrity_diagnostics += 1;
                }
                for diagnostic in paragraphs
                    .iter()
                    .flat_map(|paragraph| &paragraph.runs)
                    .flat_map(|run| &run.diagnostics)
                {
                    match diagnostic.category {
                        DiagnosticCategory::Integrity => {
                            stats.integrity_diagnostics += 1;
                            *stats
                                .integrity_reasons
                                .entry(format!("{}: {}", diagnostic.code, diagnostic.detail))
                                .or_default() += 1;
                        }
                        DiagnosticCategory::Fidelity => {
                            stats.fidelity_diagnostics += 1;
                            *stats
                                .fidelity_reasons
                                .entry(format!("{}: {}", diagnostic.code, diagnostic.detail))
                                .or_default() += 1;
                        }
                    }
                }
            }
            None => assert_eq!(expected_runs, 0, "{id}: resolved text was not rendered"),
            Some(_) => unreachable!(),
        }
        for child in shape.shapes() {
            assert_shape_text(
                renderer, package, resolver, references, page, page_part, child, text_boxes,
                rendered, stats,
            );
        }
    }

    #[derive(Debug, PartialEq)]
    struct ExpectedTextRun {
        text: String,
        diagnostic_categories: Vec<DiagnosticCategory>,
    }

    #[derive(Default)]
    struct OracleCharacter {
        family: String,
        bold: bool,
        italic: bool,
        case: i32,
        diagnostics: Vec<DiagnosticCategory>,
        diagnosed_face: Option<String>,
    }

    fn oracle_value<'a>(
        properties: &'a std::collections::BTreeMap<String, Lookup>,
        name: &str,
    ) -> Option<&'a str> {
        match properties.get(name)? {
            Lookup::Found(cell) => cell.cell.value.as_deref(),
            Lookup::Deleted | Lookup::Absent => None,
        }
    }

    fn oracle_number(
        properties: &std::collections::BTreeMap<String, Lookup>,
        name: &str,
    ) -> Option<f64> {
        oracle_value(properties, name)?
            .parse()
            .ok()
            .filter(|value: &f64| value.is_finite())
    }

    fn oracle_row(
        resolved: &ResolvedShape,
        section: &str,
        index: u32,
    ) -> std::collections::BTreeMap<String, Lookup> {
        resolved
            .sections
            .get(section)
            .and_then(|section| section.rows.get(&format!("IX:{index}")))
            .map(|row| row.cells.clone())
            .unwrap_or_default()
    }

    fn oracle_case(case: i32, text: &str) -> Result<String, ()> {
        match case {
            0 => Ok(text.into()),
            1 => Ok(text.to_uppercase()),
            2 => {
                let mut start = true;
                Ok(text
                    .chars()
                    .flat_map(|ch| {
                        let output = if start {
                            ch.to_uppercase().collect::<String>()
                        } else {
                            ch.to_lowercase().collect()
                        };
                        start = !ch.is_alphanumeric();
                        output.chars().collect::<Vec<_>>()
                    })
                    .collect())
            }
            _ => Err(()),
        }
    }

    fn oracle_character(
        renderer: &Renderer,
        package: &VsdxPackage,
        properties: &std::collections::BTreeMap<String, Lookup>,
        mut character: OracleCharacter,
    ) -> OracleCharacter {
        if character.family.is_empty() {
            character.family = "sans-serif".into();
        }
        if let Some(font) = oracle_value(properties, "Font") {
            character.family = package
                .face_names
                .iter()
                .find(|face| {
                    face.attributes
                        .iter()
                        .any(|(key, value)| key == "ID" && value == font)
                })
                .and_then(|face| {
                    face.attributes
                        .iter()
                        .find(|(key, _)| key == "Name")
                        .map(|(_, value)| value.clone())
                })
                .unwrap_or_else(|| font.into());
        }
        if let Some(style) = oracle_number(properties, "Style") {
            let style = style as i32;
            character.bold = style & 1 != 0;
            character.italic = style & 2 != 0;
        }
        if let Some(case) = oracle_number(properties, "Case") {
            character.case = case as i32;
        }
        if let Some(pos) = oracle_number(properties, "Pos")
            && !matches!(pos as i32, 0..=2)
        {
            character.diagnostics.push(DiagnosticCategory::Fidelity);
        }
        if properties.contains_key("Color") && oracle_value(properties, "Color").is_none() {
            character.diagnostics.push(DiagnosticCategory::Integrity);
        }
        let exact = renderer.registered_fonts.contains_key(&(
            character.family.clone(),
            character.bold,
            character.italic,
        ));
        if !exact && character.diagnosed_face.as_deref() != Some(character.family.as_str()) {
            character.diagnostics.push(DiagnosticCategory::Fidelity);
            character.diagnosed_face = Some(character.family.clone());
        }
        character
    }

    fn oracle_field(properties: &std::collections::BTreeMap<String, Lookup>) -> Result<String, ()> {
        oracle_value(properties, "Value")
            .filter(|value| !value.is_empty())
            .map(str::to_owned)
            .ok_or(())
    }

    fn expected_text_runs(
        renderer: &Renderer,
        package: &VsdxPackage,
        _references: Option<&PageShapeReferences>,
        resolved: &ResolvedShape,
        _shape_id: u32,
        tokens: &[vsdx_resolve::ResolvedTextToken],
    ) -> Vec<Vec<ExpectedTextRun>> {
        let mut character = oracle_character(
            renderer,
            package,
            &oracle_row(resolved, "Character", 0),
            OracleCharacter::default(),
        );
        let mut paragraphs = vec![Vec::new()];
        let mut justified = vec![false];
        for token in tokens {
            match token {
                vsdx_resolve::ResolvedTextToken::Literal(value) => {
                    let mut categories = character.diagnostics.clone();
                    let text = match oracle_case(character.case, value) {
                        Ok(text) => text,
                        Err(_) => {
                            categories.push(DiagnosticCategory::Fidelity);
                            value.clone()
                        }
                    };
                    paragraphs.last_mut().unwrap().push(ExpectedTextRun {
                        text,
                        diagnostic_categories: categories,
                    });
                }
                vsdx_resolve::ResolvedTextToken::CharacterRun {
                    index: _,
                    properties,
                } => {
                    if properties.is_empty() {
                        character.diagnostics.push(DiagnosticCategory::Integrity);
                    } else {
                        character = oracle_character(renderer, package, properties, character);
                    }
                }
                vsdx_resolve::ResolvedTextToken::ParagraphRun { properties, .. } => {
                    paragraphs.push(Vec::new());
                    justified.push(oracle_number(properties, "HorzAlign") == Some(3.0));
                }
                vsdx_resolve::ResolvedTextToken::Tab { properties, .. } => {
                    let mut categories = character.diagnostics.clone();
                    if oracle_number(properties, "Position").is_none() {
                        categories.push(DiagnosticCategory::Fidelity);
                    }
                    paragraphs.last_mut().unwrap().push(ExpectedTextRun {
                        text: "\t".into(),
                        diagnostic_categories: categories,
                    });
                }
                vsdx_resolve::ResolvedTextToken::Field { properties, .. } => {
                    let result = oracle_field(properties);
                    let mut categories = character.diagnostics.clone();
                    if result.is_err() {
                        categories.push(DiagnosticCategory::Integrity);
                    }
                    paragraphs.last_mut().unwrap().push(ExpectedTextRun {
                        text: result.unwrap_or_else(|_| "[unresolved field]".into()),
                        diagnostic_categories: categories,
                    });
                }
            }
        }
        for (runs, justified) in paragraphs.iter_mut().zip(justified) {
            if justified && let Some(run) = runs.first_mut() {
                run.diagnostic_categories.push(DiagnosticCategory::Fidelity);
            }
        }
        paragraphs
    }

    fn assert_display_text_matches(
        id: &str,
        paragraphs: &[TextParagraph],
        expected: &[Vec<ExpectedTextRun>],
    ) -> Result<(), String> {
        if paragraphs.len() != expected.len() {
            return Err(format!("{id}: paragraph count"));
        }
        for (paragraph_index, (actual, expected)) in paragraphs.iter().zip(expected).enumerate() {
            if actual.runs.len() != expected.len() {
                return Err(format!("{id}: paragraph {paragraph_index} run count"));
            }
            for (run_index, (actual, expected)) in actual.runs.iter().zip(expected).enumerate() {
                if actual.text != expected.text {
                    return Err(format!(
                        "{id}: paragraph {paragraph_index} run {run_index} text"
                    ));
                }
                let categories = actual
                    .diagnostics
                    .iter()
                    .map(|item| item.category)
                    .collect::<Vec<_>>();
                if categories != expected.diagnostic_categories {
                    return Err(format!(
                        "{id}: paragraph {paragraph_index} run {run_index} diagnostics"
                    ));
                }
            }
        }
        Ok(())
    }

    #[derive(Default)]
    struct TextCorpusStats {
        shapes: usize,
        zero_integrity_diagnostics: usize,
        integrity_diagnostics: usize,
        fidelity_diagnostics: usize,
        paragraphs: usize,
        runs: usize,
        marker_only: usize,
        field_only: usize,
        style_inherited: usize,
        master_inherited: usize,
        integrity_reasons: BTreeMap<String, usize>,
        fidelity_reasons: BTreeMap<String, usize>,
    }

    fn text_boxes_by_id<'a>(
        primitives: &'a [Primitive],
        text_boxes: &mut BTreeMap<String, &'a Primitive>,
    ) {
        for primitive in primitives {
            match primitive {
                Primitive::TextBox { id, .. } => {
                    assert!(text_boxes.insert(id.clone(), primitive).is_none());
                }
                Primitive::Group { primitives, .. } => text_boxes_by_id(primitives, text_boxes),
                _ => {}
            }
        }
    }
}
