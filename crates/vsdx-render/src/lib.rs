//! VSDX display-list compilation. Scene data remains in Visio inches until paint.

mod display_list;
mod layout;
mod line_jumps;
mod paint;
mod pdf;

pub use display_list::*;
pub use layout::{PIXELS_PER_INCH, final_paint_transform, to_canvas, to_canvas_length};

use std::collections::{BTreeMap, HashMap};

use thiserror::Error;
use vsdx_eval::{Evaluation, PageShapeReferences, Value, evaluate_cell_with_shape_package_theme};
use vsdx_parse::{ParseLimits, Shape, VsdxPackage};
use vsdx_resolve::{Lookup, ResolvedShape, Resolver, ScenePoint, SceneTransform, realize_geometry};

const MAX_RECURSION_DEPTH: usize = 64;

#[derive(Default)]
struct LayoutCache {
    fonts: HashMap<String, [Option<Option<ooxml_text::FontId>>; 4]>,
    advances: HashMap<(Option<ooxml_text::FontId>, u32, char), f32>,
}

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

struct RunCursor<'a> {
    runs: std::slice::Iter<'a, TextRun>,
    run: Option<&'a TextRun>,
    end: usize,
}

impl<'a> RunCursor<'a> {
    fn new(runs: &'a [TextRun]) -> Self {
        Self {
            runs: runs.iter(),
            run: None,
            end: 0,
        }
    }

    fn advance_to(&mut self, index: usize) -> Option<(&'a TextRun, usize)> {
        while self.end <= index {
            self.run = self.runs.next();
            self.end += self.run?.text.len();
        }
        self.run.map(|run| (run, self.end))
    }
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
        family: "Calibri".into(),
        size_in: 10.0 / 72.0,
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
        align: 1,
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
                    align: property_number(properties, "HorzAlign").unwrap_or(1.0) as i32,
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
    layer_overrides: BTreeMap<(String, u32), bool>,
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
            layer_overrides: BTreeMap::new(),
        }
    }
    /// Session-local layer visibility, shadowing the page sheet until cleared.
    pub fn set_layer_override(&mut self, page_part: &str, index: u32, visible: bool) {
        self.layer_overrides
            .insert((page_part.to_owned(), index), visible);
    }
    pub fn clear_layer_overrides(&mut self) {
        self.layer_overrides.clear();
    }
    /// Page layers with session overrides applied, in row-index order.
    pub fn effective_page_layers(
        &self,
        package: &VsdxPackage,
        page_part: &str,
    ) -> Vec<vsdx_resolve::PageLayer> {
        let mut layers = vsdx_resolve::page_layers(package, page_part);
        for layer in &mut layers {
            if let Some(visible) = self
                .layer_overrides
                .get(&(page_part.to_owned(), layer.index))
            {
                layer.visible = *visible;
            }
        }
        layers
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
        let fallback_shapes;
        let shapes = match references.as_ref() {
            Some(references) => references.shapes(),
            None => {
                fallback_shapes = resolver.resolve_page_shapes(page_part)?;
                &fallback_shapes
            }
        };
        let connectivity = resolver.resolve_page_connectivity_with(page_part, shapes)?;
        let transforms = vsdx_resolve::scene_transforms(page, shapes, |id, shape, name| {
            evaluated(package, references.as_ref(), shape, id, name)
        });
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
            jump_overrides: Vec::new(),
        };
        let mut cache = LayoutCache::default();
        let layers = self.effective_page_layers(package, page_part);
        for shape in page.shapes() {
            self.layout_shape(
                package,
                &resolver,
                &connectivity,
                references.as_ref(),
                shapes,
                &transforms,
                &layers,
                page_part,
                shape,
                0,
                &mut state,
                &mut cache,
            )?;
        }
        line_jumps::apply_line_jumps(
            &mut state.primitives,
            &state.jump_overrides,
            &line_jumps::page_jump_settings(package, &resolver, page_part),
        );
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
        connectivity: &vsdx_resolve::PageConnectivity,
        references: Option<&PageShapeReferences>,
        shapes: &BTreeMap<u32, ResolvedShape>,
        transforms: &BTreeMap<u32, SceneTransform>,
        layers: &[vsdx_resolve::PageLayer],
        page_part: &str,
        shape: &Shape,
        depth: usize,
        state: &mut State,
        cache: &mut LayoutCache,
    ) -> Result<(), RenderError> {
        if depth >= MAX_RECURSION_DEPTH {
            return self.placeholder(page_part, shape, state, "group nesting depth exceeded");
        }
        state.count += 1;
        if state.count > self.limits.max_shapes {
            return Err(RenderError::Budget("shapes"));
        }
        let z_order = state.next_z();
        let id = format!("{page_part}:{}", shape.id);
        let fallback_resolved;
        let resolved = match shapes.get(&shape.id) {
            Some(resolved) => resolved,
            None => {
                fallback_resolved = resolver.resolve_shape(page_part, shape.id)?;
                &fallback_resolved
            }
        };
        if resolved.deleted {
            return Ok(());
        }
        if paint::number(resolved, "NoShow").is_some_and(|value| value != 0.0) {
            return Ok(());
        }
        if vsdx_resolve::shape_hidden_by_layers(resolved, layers) {
            return Ok(());
        }
        let mut sections = resolved
            .sections
            .values()
            .filter(|section| section.name == "Geometry" && !section.deleted)
            .collect::<Vec<_>>();
        if let Some(section) = sections
            .iter()
            .find(|section| !section.unsupported_controls.is_empty())
        {
            return self.placeholder_at(
                id,
                z_order,
                bounds(package, references, resolved, shape.id)
                    .filter(|bounds| bounds_finite(*bounds))
                    .unwrap_or_default(),
                state,
                &format!(
                    "unsupported Geometry section controls at IX={}: {}",
                    section.index.unwrap_or(0),
                    section.unsupported_controls.join(", ")
                ),
            );
        }
        if connectivity
            .connectors
            .get(&shape.id)
            .is_some_and(|connector| {
                connector.is_1d
                    && connector.glue.iter().all(|glue| {
                        !glue.diagnostics.iter().any(|diagnostic| {
                            matches!(
                                diagnostic,
                                vsdx_resolve::ConnectivityDiagnostic::UnsupportedFromCell { .. }
                            )
                        })
                    })
                    || (connector.begin.is_some() || connector.end.is_some())
                        && connector.glue.iter().any(|glue| {
                            glue.to
                                .as_ref()
                                .is_some_and(|target| target.connection_point.is_none())
                        })
            })
        {
            return self.layout_connector(
                package,
                resolver,
                connectivity,
                references,
                page_part,
                shape,
                id,
                z_order,
                resolved,
                &sections,
                state,
                transforms,
            );
        }
        let Some(bounds) = bounds(package, references, resolved, shape.id) else {
            return self.placeholder(page_part, shape, state, "unresolvable transform");
        };
        let Some(transform) = transforms.get(&shape.id) else {
            return self.placeholder_at(id, z_order, bounds, state, "unresolvable transform");
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
        sections.sort_by_key(|section| section.index.unwrap_or(0));
        let geometry = (!sections.is_empty()).then(|| {
            sections
                .iter()
                .map(|section| realize_geometry(section, bounds.width, bounds.height))
                .collect::<Vec<_>>()
        });
        let child_shapes = shape.shapes().collect::<Vec<_>>();
        if !child_shapes.is_empty() {
            let group_transform = affine(transform.local);
            let start = state.primitives.len();
            for child in child_shapes {
                self.layout_shape(
                    package,
                    resolver,
                    connectivity,
                    references,
                    shapes,
                    transforms,
                    layers,
                    page_part,
                    child,
                    depth + 1,
                    state,
                    cache,
                )?;
            }
            let children = state.primitives.split_off(start);
            let z_order = state.next_z();
            state.primitives.push(Primitive::Group {
                id,
                z_order,
                primitives: children,
                transform: group_transform,
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
                    transform: affine(transform.local),
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
        let drawn = geometry
            .iter()
            .filter(|realized| !realized.controls.no_show)
            .collect::<Vec<_>>();
        let issues = drawn
            .iter()
            .flat_map(|realized| realized.issues.iter().cloned())
            .collect::<Vec<_>>();
        if !issues.is_empty() {
            return self.placeholder_at(
                id,
                z_order,
                bounds,
                state,
                &format!("unsupported geometry: {issues:?}"),
            );
        }
        let paths = drawn
            .iter()
            .filter(|realized| !realized.commands.is_empty())
            .map(|realized| {
                (
                    &realized.controls,
                    realized
                        .commands
                        .iter()
                        .cloned()
                        .map(|mut command| {
                            transform_affine(&mut command, affine(transform.local));
                            command
                        })
                        .collect::<Vec<_>>(),
                )
            })
            .collect::<Vec<_>>();
        if paths.is_empty() && !drawn.is_empty() {
            return self.placeholder_at(
                id,
                z_order,
                bounds,
                state,
                "unsupported geometry: no path commands",
            );
        }
        let needs_fill = paths.iter().any(|(controls, _)| !controls.no_fill);
        let needs_stroke = paths.iter().any(|(controls, _)| !controls.no_line);
        let outcome = paint::paint(
            package,
            references,
            resolved,
            shape.id,
            needs_fill,
            needs_stroke,
        );
        let mut diagnostics = outcome.diagnostics;
        for (controls, path) in paths {
            state.primitives.push(Primitive::Shape {
                id: id.clone(),
                z_order,
                path,
                fill: if controls.no_fill {
                    None
                } else {
                    outcome.fill.clone()
                },
                stroke: if controls.no_line {
                    None
                } else {
                    outcome.stroke.clone()
                },
                transform: Affine::identity(),
                diagnostics: std::mem::take(&mut diagnostics),
            });
        }
        self.text(
            package,
            resolver,
            references,
            page_part,
            shape,
            resolved,
            id,
            bounds,
            affine(transform.local).compose(Affine {
                a: 1.0,
                b: 0.0,
                c: 0.0,
                d: 1.0,
                e: -bounds.x as f32,
                f: -bounds.y as f32,
            }),
            state,
            cache,
        )?;
        Ok(())
    }
    #[allow(clippy::too_many_arguments)]
    fn layout_connector(
        &self,
        package: &VsdxPackage,
        resolver: &Resolver<'_>,
        connectivity: &vsdx_resolve::PageConnectivity,
        references: Option<&PageShapeReferences>,
        page_part: &str,
        shape: &Shape,
        id: String,
        z_order: u32,
        resolved: &ResolvedShape,
        sections: &[&vsdx_resolve::ResolvedSection],
        state: &mut State,
        transforms: &BTreeMap<u32, SceneTransform>,
    ) -> Result<(), RenderError> {
        let visible = sections.iter().filter(|section| !section.controls.no_show);
        if !sections.is_empty() && visible.clone().next().is_none() {
            return Ok(());
        }
        let Some(connector) = connectivity.connectors.get(&shape.id) else {
            return self.placeholder(
                page_part,
                shape,
                state,
                "1D shape is missing connector state",
            );
        };
        let mut begin = connector.begin;
        let mut end = connector.end;
        for glue in &connector.glue {
            let Some(point) = glue
                .to
                .as_ref()
                .and_then(|target| target.connection_point.as_ref())
                .map(|point| point.position)
            else {
                let reason = glue.diagnostics.first().map_or_else(
                    || "connector route cannot be computed: unresolved glued endpoint".into(),
                    |diagnostic| format!("connector route cannot be computed: unresolved glued endpoint: {diagnostic:?}"),
                );
                return self.placeholder_at(id, z_order, Bounds::default(), state, &reason);
            };
            match glue.endpoint {
                vsdx_resolve::ConnectorEndpoint::Begin => begin = Some(point),
                vsdx_resolve::ConnectorEndpoint::End => end = Some(point),
            }
        }
        let (Some(begin), Some(end)) = (begin, end) else {
            return self.placeholder_at(
                id,
                z_order,
                Bounds::default(),
                state,
                "connector route cannot be computed: unresolved endpoint",
            );
        };
        if ![begin.x, begin.y, end.x, end.y]
            .into_iter()
            .all(f64::is_finite)
        {
            return self.placeholder(
                page_part,
                shape,
                state,
                "connector route cannot be computed: non-finite endpoint",
            );
        }
        let needs_fill =
            sections.is_empty() || visible.clone().any(|section| !section.controls.no_fill);
        let needs_stroke =
            sections.is_empty() || visible.clone().any(|section| !section.controls.no_line);
        let outcome = paint::paint(
            package,
            references,
            resolved,
            shape.id,
            needs_fill,
            needs_stroke,
        );
        let style = connector_route_style(package, resolver, page_part, resolved);
        let path = connector_geometry(
            package, references, resolved, shape.id, transforms, begin, end,
        )
        .unwrap_or_else(|| connector_route(begin, end, style));
        state.primitives.push(Primitive::Shape {
            id: id.clone(),
            z_order,
            path,
            fill: outcome.fill,
            stroke: outcome.stroke,
            transform: Affine::identity(),
            diagnostics: outcome.diagnostics,
        });
        state
            .jump_overrides
            .push(line_jumps::ConnectorJumpOverride {
                id,
                code: line_jumps::connector_override(package, resolved, "ConLineJumpCode"),
                style: line_jumps::connector_override(package, resolved, "ConLineJumpStyle"),
                dir_x: line_jumps::connector_override(package, resolved, "ConLineJumpDirX"),
                dir_y: line_jumps::connector_override(package, resolved, "ConLineJumpDirY"),
            });
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
        transform: Affine,
        state: &mut State,
        cache: &mut LayoutCache,
    ) -> Result<(), RenderError> {
        if matches!(resolved.cell("Char.Size"), Some(Lookup::Found(cell)) if cell.cell.value.as_deref().and_then(|value| value.parse::<f32>().ok()).is_some_and(|value| !value.is_finite()))
        {
            return Err(RenderError::PageDimensions(
                "non-finite text metrics".into(),
            ));
        }
        let lookup = package
            .page_contents
            .get(page_part)
            .ok_or_else(|| RenderError::MissingPage(page_part.into()))?;
        let tokens = resolver.resolve_text_in_context(shape, lookup, resolved)?;
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
            let mut laid_out = self.wrap_paragraph(
                paragraph,
                width,
                x + paragraph.left,
                cursor_y,
                offset,
                cache,
            );
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
        let vertical = paint::number(resolved, "VerticalAlign").unwrap_or(1.0) as i32;
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
            transform,
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
        cache: &mut LayoutCache,
    ) -> Vec<PositionedLine> {
        let text = paragraph
            .runs
            .iter()
            .map(|run| run.text.as_str())
            .collect::<String>();
        let breaks = ooxml_text::break_opportunities(&text);
        let mut widths = Vec::new();
        let mut height = 0.0f32;
        let mut runs = RunCursor::new(&paragraph.runs);
        for (index, ch) in text.char_indices() {
            let (run, run_end) = runs
                .advance_to(index)
                .map_or((None, None), |(run, end)| (Some(run), Some(end)));
            let (mut width, line_height) = run
                .map(|run| {
                    (
                        self.measure(run, ch, cache)
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
                    self.measure_text(paragraph, index + ch.len_utf8(), following, cache);
                let before_decimal = following
                    .split_once('.')
                    .map_or(following, |(before, _)| before);
                let decimal_width =
                    self.measure_text(paragraph, index + ch.len_utf8(), before_decimal, cache);
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
        let mut start_index = 0usize;
        let mut cursor = 0.0;
        let mut last_break = None;
        let mut mandatory_break = None;
        let mut break_index = 0usize;
        for (width_index, (position, length, width)) in widths.iter().copied().enumerate() {
            if mandatory_break == Some(position) {
                result.push(self.line(
                    &text,
                    &widths[start_index..width_index],
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
                start_index = width_index;
                cursor = 0.0;
                last_break = None;
                mandatory_break = None;
            }
            if position > start && cursor + width > available_width && available_width > 0.0 {
                let (break_position, break_index) = last_break.unwrap_or((position, width_index));
                let end = break_position.max(start + 1);
                result.push(self.line(
                    &text,
                    &widths[start_index..break_index],
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
                cursor = widths[break_index..width_index]
                    .iter()
                    .map(|(_, _, w)| *w)
                    .sum();
                start_index = break_index;
                last_break = None;
            }
            cursor += width;
            let next = position + length;
            while break_index < breaks.len() && breaks[break_index].byte_index < next {
                break_index += 1;
            }
            if break_index < breaks.len() && breaks[break_index].byte_index == next {
                let opportunity = &breaks[break_index];
                if opportunity.mandatory && next < text.len() {
                    mandatory_break = Some(next);
                } else {
                    last_break = Some((next, width_index + 1));
                }
            }
        }
        if start < text.len() || result.is_empty() {
            result.push(self.line(
                &text,
                &widths[start_index..],
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
        let width = widths.iter().map(|(_, _, width)| *width).sum::<f32>();
        let line_x = x + match align {
            1 => (available - width) * 0.5,
            2 => available - width,
            _ => 0.0,
        };
        let mut advance = 0.0;
        let mut stops = Vec::new();
        for (position, length, char_width) in widths.iter().copied() {
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
        for candidate_family in [family, "Arial", "sans-serif"] {
            for (candidate_bold, candidate_italic) in [
                (bold, italic),
                (bold, false),
                (false, italic),
                (false, false),
            ] {
                if let Some(font) = self.registered_fonts.iter().find_map(
                    |((registered_family, registered_bold, registered_italic), font)| {
                        (registered_family == candidate_family
                            && *registered_bold == candidate_bold
                            && *registered_italic == candidate_italic)
                            .then_some(*font)
                    },
                ) {
                    return Some(font);
                }
            }
        }
        None
    }
    fn cached_font_for(
        &self,
        run: &TextRun,
        cache: &mut LayoutCache,
    ) -> Option<ooxml_text::FontId> {
        let style = (run.bold as usize) << 1 | run.italic as usize;
        if let Some(styles) = cache.fonts.get_mut(run.family.as_str()) {
            if let Some(font) = styles[style] {
                return font;
            }
            let font = self.font_for(&run.family, run.bold, run.italic);
            styles[style] = Some(font);
            return font;
        }
        let font = self.font_for(&run.family, run.bold, run.italic);
        let mut styles = [None; 4];
        styles[style] = Some(font);
        cache.fonts.insert(run.family.clone(), styles);
        font
    }
    fn measure(&self, run: &TextRun, ch: char, cache: &mut LayoutCache) -> f32 {
        let font = self.cached_font_for(run, cache);
        let key = (font, run.size_in.to_bits(), ch);
        if let Some(width) = cache.advances.get(&key) {
            return *width;
        }
        let mut buffer = [0; 4];
        let width = font
            .and_then(|font| {
                ooxml_text::shape(
                    &self.fonts,
                    font,
                    ch.encode_utf8(&mut buffer),
                    run.size_in,
                    &[],
                )
                .ok()
            })
            .map(|glyphs| glyphs.iter().map(|glyph| glyph.x_advance).sum())
            .unwrap_or(run.size_in * 0.5);
        cache.advances.insert(key, width);
        width
    }
    fn measure_text(
        &self,
        paragraph: &RichParagraph,
        start: usize,
        text: &str,
        cache: &mut LayoutCache,
    ) -> f32 {
        let mut runs = RunCursor::new(&paragraph.runs);
        text.char_indices()
            .map(|(relative, ch)| {
                let index = start + relative;
                runs.advance_to(index).map_or(0.0, |(run, end)| {
                    self.measure(run, ch, cache)
                        + (end > index + ch.len_utf8()) as u8 as f32 * run.letter_spacing
                })
            })
            .sum()
    }
    fn placeholder(
        &self,
        page_part: &str,
        shape: &Shape,
        state: &mut State,
        reason: &str,
    ) -> Result<(), RenderError> {
        self.placeholder_at(
            format!("{page_part}:{}", shape.id),
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
fn affine(transform: vsdx_resolve::SceneAffine) -> Affine {
    Affine {
        a: transform.a as f32,
        b: transform.b as f32,
        c: transform.c as f32,
        d: transform.d as f32,
        e: transform.e as f32,
        f: transform.f as f32,
    }
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
struct State {
    count: usize,
    z_order: u32,
    text_bytes: usize,
    text_paragraphs: usize,
    text_lines: usize,
    text_runs: usize,
    primitives: Vec<Primitive>,
    jump_overrides: Vec<line_jumps::ConnectorJumpOverride>,
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
}
/// No style, straight and center-to-center run directly, vertical starters bend vertical-first, other right-angle styles bend horizontal-first.
fn connector_route(
    begin: ScenePoint,
    end: ScenePoint,
    route_style: f64,
) -> Vec<ooxml_drawingml::GeometryPathCommand> {
    use ooxml_drawingml::GeometryPathCommand::{Line, Move};
    let mut path = vec![Move {
        x: begin.x,
        y: begin.y,
    }];
    let straight = route_style == 0.0 || route_style == 2.0 || route_style == 16.0;
    if !straight && begin.x != end.x && begin.y != end.y {
        if matches!(route_style as i64, 3 | 5 | 7 | 10 | 12 | 14 | 17 | 19 | 22) {
            path.push(Line {
                x: begin.x,
                y: end.y,
            });
        } else {
            path.push(Line {
                x: end.x,
                y: begin.y,
            });
        }
    }
    path.push(Line { x: end.x, y: end.y });
    path
}
/// Per-shape ShapeRouteStyle wins, zero means the page RouteStyle, absent page means no routing intent.
fn connector_route_style(
    package: &VsdxPackage,
    resolver: &Resolver<'_>,
    page_part: &str,
    resolved: &ResolvedShape,
) -> f64 {
    if let Some(style) = paint::number(resolved, "ShapeRouteStyle")
        && style != 0.0
    {
        return style;
    }
    page_dimension(resolver, package, page_part, "RouteStyle").unwrap_or(0.0)
}
/// Filed connector waypoints in scene space, or nothing when the file route is unusable.
fn connector_geometry(
    package: &VsdxPackage,
    references: Option<&PageShapeReferences>,
    resolved: &ResolvedShape,
    shape_id: u32,
    transforms: &BTreeMap<u32, SceneTransform>,
    begin: ScenePoint,
    end: ScenePoint,
) -> Option<Vec<ooxml_drawingml::GeometryPathCommand>> {
    let transform = transforms.get(&shape_id)?;
    let size = bounds(package, references, resolved, shape_id)?;
    if ![size.width, size.height].into_iter().all(f64::is_finite) {
        return None;
    }
    let mut sections = resolved
        .sections
        .values()
        .filter(|section| section.name == "Geometry" && !section.deleted)
        .collect::<Vec<_>>();
    if sections.is_empty() {
        return None;
    }
    sections.sort_by_key(|section| section.index.unwrap_or(0));
    let mut geometry = vsdx_resolve::RealizedGeometry::default();
    for section in sections {
        let realized = realize_geometry(section, size.width, size.height);
        geometry.commands.extend(realized.commands);
        geometry.issues.extend(realized.issues);
    }
    if geometry.commands.is_empty()
        || !geometry.issues.is_empty()
        || geometry
            .commands
            .iter()
            .all(|command| matches!(command, ooxml_drawingml::GeometryPathCommand::Move { .. }))
        || geometry
            .commands
            .iter()
            .any(|command| matches!(command, ooxml_drawingml::GeometryPathCommand::Close))
    {
        return None;
    }
    if !matches!(
        geometry.commands.first(),
        Some(ooxml_drawingml::GeometryPathCommand::Move { .. })
    ) {
        geometry.commands.insert(
            0,
            ooxml_drawingml::GeometryPathCommand::Move { x: 0.0, y: 0.0 },
        );
    }
    if geometry
        .commands
        .iter()
        .filter(|command| matches!(command, ooxml_drawingml::GeometryPathCommand::Move { .. }))
        .nth(1)
        .is_some()
    {
        return None;
    }
    let matrix = affine(transform.local);
    for command in &mut geometry.commands {
        transform_affine(command, matrix);
    }
    if !geometry.commands.iter().all(command_finite) {
        return None;
    }
    let mut path = geometry.commands;
    path[0] = ooxml_drawingml::GeometryPathCommand::Move {
        x: begin.x,
        y: begin.y,
    };
    path.push(ooxml_drawingml::GeometryPathCommand::Line { x: end.x, y: end.y });
    Some(path)
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
    fn collect<'a>(
        primitives: &'a [Primitive],
        ancestors: Affine,
        output: &mut Vec<(&'a Primitive, Affine)>,
    ) {
        for primitive in primitives {
            match primitive {
                Primitive::Group {
                    primitives,
                    transform,
                    ..
                } => collect(primitives, ancestors.compose(*transform), output),
                Primitive::Shape { transform, .. }
                | Primitive::Image { transform, .. }
                | Primitive::TextBox { transform, .. } => {
                    output.push((primitive, ancestors.compose(*transform)))
                }
                Primitive::Placeholder { .. } => output.push((primitive, ancestors)),
            }
        }
    }
    let mut primitives = Vec::new();
    collect(&list.primitives, Affine::identity(), &mut primitives);
    primitives.sort_by_key(|(primitive, _)| std::cmp::Reverse(z_order(primitive)));
    for (primitive, matrix) in primitives {
        match primitive {
            Primitive::TextBox {
                id,
                x: left,
                y: top,
                width,
                height,
                lines,
                ..
            } => {
                if let Some(position) =
                    text_caret_at((*left, *top, *width, *height), lines, matrix, (x, y))
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
                ..
            } if path_hit(
                path,
                matrix,
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
                ..
            } if point_in_transformed_rect(*left, *top, *width, *height, matrix, x, y) => {
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
            } if point_in_transformed_rect(*left, *top, *width, *height, matrix, x, y) => {
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
    let mut winding = 0_i64;
    for points in flatten_path(path) {
        if stroke_width > 0.0
            && points.windows(2).any(|segment| {
                point_segment_distance(x, y, segment[0], segment[1]) <= stroke_width / 2.0
            })
        {
            return true;
        }
        for index in 0..points.len() {
            let (ax, ay) = points[index];
            let (bx, by) = points[(index + 1) % points.len()];
            let side = (bx - ax) * (y - ay) - (x - ax) * (by - ay);
            if ay <= y && by > y && side > 0.0 {
                winding += 1;
            } else if ay > y && by <= y && side < 0.0 {
                winding -= 1;
            }
        }
    }
    fill && winding != 0
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
fn flatten_path(path: &[ooxml_drawingml::GeometryPathCommand]) -> Vec<Vec<(f32, f32)>> {
    use ooxml_drawingml::GeometryPathCommand::*;
    let mut paths = Vec::new();
    let mut points = Vec::new();
    let mut current = (0.0, 0.0);
    for command in path {
        match *command {
            Move { x, y } => {
                if !points.is_empty() {
                    paths.push(std::mem::take(&mut points));
                }
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
            Close => {
                if let Some(first) = points.first().copied() {
                    points.push(first);
                    current = first;
                }
            }
        }
    }
    if !points.is_empty() {
        paths.push(points);
    }
    paths
}

#[cfg(test)]
mod tests {
    #[test]
    fn forty_five_degree_group_keeps_text_in_a_local_box_under_the_group_transform() {
        let mut child = shape(2, 1.0, 2.0);
        child
            .children
            .push(ShapeChild::Text(vec![TextToken::Literal("ab".into())]));
        child.children.push(text_section(
            "Character",
            vec![row(0, "", vec![cell("Size", "1")])],
        ));
        let mut parent = group(1, 10.0, 20.0, vec![child]);
        with_cell(
            &mut parent,
            "Angle",
            &std::f64::consts::FRAC_PI_4.to_string(),
        );
        let list = render(vec![parent]);
        let Primitive::Group {
            primitives,
            transform: group,
            ..
        } = &list.primitives[0]
        else {
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
            e: 10.0,
            f: 20.0,
        };
        assert_point_close((group.a, group.b), (matrix.a, matrix.b));
        assert_point_close((group.c, group.d), (matrix.c, matrix.d));
        assert_point_close((group.e, group.f), (matrix.e, matrix.f));
        assert_eq!(*transform, Affine::identity());
        assert_point_close((*x, *y), (1.0, 2.0));
        assert_point_close((*width, *height), (1.0, 1.0));
        let page = group.compose(*transform);
        assert_point_close(
            page.apply_point(lines[0].x, lines[0].y),
            matrix.apply_point(1.0, 2.0),
        );
        assert_point_close(
            page.apply_point(lines[0].caret_stops[1].x, lines[0].caret_stops[1].y),
            matrix.apply_point(1.5, 2.0),
        );
    }

    use super::*;
    use ooxml_drawingml::GeometryPathCommand;
    use vsdx_parse::{
        Cell, Connect, ConnectsChild, ForeignData, Row, RowChild, Section, SectionChild,
        ShapeChild, ShapesChild, Sheet, SheetChild, TextToken,
    };

    fn styled_run(text: &str, family: &str, bold: bool, italic: bool) -> TextRun {
        TextRun {
            text: text.into(),
            family: family.into(),
            size_in: 1.0 / 6.0,
            bold,
            italic,
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
        }
    }

    #[test]
    fn run_cursor_attributes_each_character_to_its_styled_run() {
        let paragraph = RichParagraph {
            runs: vec![
                styled_run("A", "Arial", false, false),
                styled_run("β", "Calibri", true, false),
                styled_run("CD", "Tahoma", false, true),
            ],
            align: 0,
            before: 0.0,
            after: 0.0,
            left: 0.0,
            right: 0.0,
            first: 0.0,
            line_spacing: None,
        };
        let mut cursor = RunCursor::new(&paragraph.runs);
        let attributed = "AβCD"
            .char_indices()
            .map(|(index, ch)| {
                let (run, end) = cursor.advance_to(index).unwrap();
                (ch, run.family.as_str(), run.bold, run.italic, end)
            })
            .collect::<Vec<_>>();

        assert_eq!(
            attributed,
            [
                ('A', "Arial", false, false, 1),
                ('β', "Calibri", true, false, 3),
                ('C', "Tahoma", false, true, 5),
                ('D', "Tahoma", false, true, 5),
            ]
        );
    }

    #[test]
    fn wrapping_preserves_boundaries_and_caret_stops_across_mandatory_breaks() {
        let mut run = styled_run("ab cd\nef gh", "sans-serif", false, false);
        run.size_in = 1.0;
        let paragraph = RichParagraph {
            runs: vec![run],
            align: 0,
            before: 0.0,
            after: 0.0,
            left: 0.0,
            right: 0.0,
            first: 0.0,
            line_spacing: None,
        };

        let mut cache = LayoutCache::default();
        let lines = Renderer::default().wrap_paragraph(&paragraph, 2.0, 0.0, 0.0, 0, &mut cache);

        assert_eq!(
            lines
                .iter()
                .map(|line| (line.start, line.end))
                .collect::<Vec<_>>(),
            [(0, 3), (3, 6), (6, 9), (9, 11)]
        );
        assert_eq!(
            lines
                .iter()
                .map(|line| {
                    line.caret_stops
                        .iter()
                        .map(|stop| stop.position)
                        .collect::<Vec<_>>()
                })
                .collect::<Vec<_>>(),
            [
                vec![0, 1, 2, 3],
                vec![3, 4, 5, 6],
                vec![6, 7, 8, 9],
                vec![9, 10, 11]
            ]
        );
    }

    #[test]
    fn measurement_cache_preserves_widths_for_distinct_run_styles() {
        let mut renderer = Renderer::default();
        for (family, bold, italic, bytes) in [
            (
                "Liberation Sans",
                false,
                false,
                include_bytes!("../../../packages/fonts/assets/LiberationSans-Regular.ttf")
                    .as_slice(),
            ),
            (
                "Liberation Sans",
                true,
                false,
                include_bytes!("../../../packages/fonts/assets/LiberationSans-Bold.ttf").as_slice(),
            ),
            (
                "Liberation Sans",
                false,
                true,
                include_bytes!("../../../packages/fonts/assets/LiberationSans-Italic.ttf")
                    .as_slice(),
            ),
            (
                "Liberation Serif",
                false,
                false,
                include_bytes!("../../../packages/fonts/assets/LiberationSerif-Regular.ttf")
                    .as_slice(),
            ),
        ] {
            renderer
                .register_font(family, bold, italic, bytes.to_vec())
                .unwrap();
        }
        let mut runs = [
            styled_run("W", "Liberation Sans", false, false),
            styled_run("W", "Liberation Sans", true, false),
            styled_run("W", "Liberation Sans", false, true),
            styled_run("W", "Liberation Serif", false, false),
            styled_run("W", "Liberation Sans", false, false),
        ];
        runs[4].size_in = 0.25;
        let expected = runs
            .iter()
            .map(|run| {
                renderer
                    .font_for(&run.family, run.bold, run.italic)
                    .and_then(|font| {
                        ooxml_text::shape(&renderer.fonts, font, "W", run.size_in, &[]).ok()
                    })
                    .map(|glyphs| glyphs.iter().map(|glyph| glyph.x_advance).sum())
                    .unwrap_or(run.size_in * 0.5)
            })
            .collect::<Vec<f32>>();
        let mut cache = LayoutCache::default();

        for (run, width) in runs.iter().zip(&expected) {
            assert_eq!(renderer.measure(run, 'W', &mut cache), *width);
        }
        for (run, width) in runs.iter().zip(&expected) {
            assert_eq!(renderer.measure(run, 'W', &mut cache), *width);
        }
        for run in &runs {
            assert_eq!(
                renderer.cached_font_for(run, &mut cache),
                renderer.font_for(&run.family, run.bold, run.italic)
            );
        }
        assert_eq!(cache.advances.len(), runs.len());
    }

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

    mod layers;
    mod section_controls;

    #[test]
    fn hit_testing_does_not_bridge_separate_subpaths() {
        use ooxml_drawingml::GeometryPathCommand::*;
        let path = [
            Move { x: 0.0, y: 0.0 },
            Line { x: 1.0, y: 0.0 },
            Move { x: 3.0, y: 0.0 },
            Line { x: 4.0, y: 0.0 },
        ];
        assert!(path_hit(&path, Affine::identity(), false, 0.1, 0.5, 0.0));
        assert!(path_hit(&path, Affine::identity(), false, 0.1, 3.5, 0.0));
        assert!(!path_hit(&path, Affine::identity(), false, 0.1, 2.0, 0.0));
    }

    #[test]
    fn renders_each_indexed_geometry_section() {
        let source = include_bytes!("../../vsdx-parse/tests/fixtures/indexed-geometry.vsdx");
        let package = vsdx_parse::parse_vsdx(source).unwrap();
        let list = Renderer::default()
            .layout_page(&package, "visio/pages/page1.xml")
            .unwrap();
        let paths = list
            .primitives
            .iter()
            .filter_map(|primitive| match primitive {
                Primitive::Shape { path, .. } => Some(path),
                _ => None,
            })
            .collect::<Vec<_>>();
        assert_eq!(paths.len(), 2);
        for path in &paths {
            assert_eq!(
                path.iter()
                    .filter(|command| matches!(
                        command,
                        ooxml_drawingml::GeometryPathCommand::Move { .. }
                    ))
                    .count(),
                1
            );
            assert_eq!(path.len(), 2);
        }
    }

    fn render(shapes: Vec<Shape>) -> VsdxDisplayList {
        let package = package(shapes);
        Renderer::default().layout_page(&package, "page").unwrap()
    }

    fn glued_connector_package(to_cell: &str) -> VsdxPackage {
        let mut connector = shape(1, 1.0, 1.0);
        connector.children.extend([
            ShapeChild::Cell(cell("OneD", "1")),
            ShapeChild::Cell(cell("BeginX", "1")),
            ShapeChild::Cell(cell("BeginY", "1")),
            ShapeChild::Cell(cell("EndX", "4")),
            ShapeChild::Cell(cell("EndY", "1")),
        ]);
        let mut target = shape(2, 5.0, 1.0);
        target.children.push(ShapeChild::Section(Section {
            name: "Connection".into(),
            index: None,
            del: false,
            children: vec![
                row(0, "Connection", vec![cell("X", "0"), cell("Y", "0.5")]),
                row(1, "Connection", vec![cell("X", "1"), cell("Y", "0.5")]),
            ]
            .into_iter()
            .map(SectionChild::Row)
            .collect(),
            other_attrs: vec![],
        }));
        let mut package = package(vec![connector, target]);
        package
            .page_contents
            .get_mut("page")
            .unwrap()
            .children
            .push(SheetChild::Connects(vec![ConnectsChild::Connect(
                Connect {
                    from_sheet: 1,
                    from_cell: Some("BeginX".into()),
                    from_part: None,
                    to_sheet: 2,
                    to_cell: Some(to_cell.into()),
                    to_part: None,
                    other_attrs: vec![],
                },
            )]));
        package
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
    fn container_and_member_paint_as_ordered_shapes() {
        let mut container = shape(1, 5.0, 4.0);
        container.children.push(ShapeChild::Cell(Cell {
            name: "Relationships".into(),
            formula: Some("SUM(DEPENDSON(1,Sheet.2!SheetRef()))".into()),
            value: Some("0".into()),
            unit: None,
            del: false,
            other_attrs: Vec::new(),
        }));
        container.children.push(ShapeChild::Section(Section {
            name: "User".into(),
            index: None,
            del: false,
            children: vec![SectionChild::Row(Row {
                index: None,
                name: Some("msvStructureType".into()),
                local_name: None,
                row_type: None,
                del: false,
                children: vec![RowChild::Cell(cell("Value", "Container"))],
                other_attrs: Vec::new(),
            })],
            other_attrs: Vec::new(),
        }));
        let member = shape(2, 5.0, 4.0);
        let list = render(vec![container, member]);
        let ids = list
            .primitives
            .iter()
            .filter_map(|primitive| match primitive {
                Primitive::Shape { id, .. } => Some(id.clone()),
                _ => None,
            })
            .collect::<Vec<_>>();
        assert_eq!(ids, vec!["page:1".to_owned(), "page:2".to_owned()]);
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
                TextToken::ParagraphRun(0),
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
            shape.children.push(text_section(
                "Paragraph",
                vec![row(0, "", vec![cell("HorzAlign", "0")])],
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
            TextToken::ParagraphRun(0),
            TextToken::CharacterRun(0),
            TextToken::Literal("a".into()),
            TextToken::Tab(3),
            TextToken::Literal("b".into()),
        ]);
        shape.children.push(text_section(
            "Tabs",
            vec![row(0, "", vec![cell("Position", "0.25")])],
        ));
        shape.children.push(text_section(
            "Paragraph",
            vec![row(0, "", vec![cell("HorzAlign", "0")])],
        ));
        with_cell(&mut shape, "DefaultTabStop", "0.75");
        let list = render(vec![shape]);
        let Primitive::TextBox {
            paragraphs, lines, ..
        } = text_box(&list)
        else {
            unreachable!()
        };
        assert!(paragraphs[1].runs[1].diagnostics.iter().any(|diagnostic| {
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
                && diagnostic.detail == "unregistered font 'Calibri'"
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
        for (align, expected_x) in [
            (None, 1.4375),
            (Some(0), 1.0),
            (Some(1), 1.4375),
            (Some(2), 1.875),
        ] {
            let mut shape = text_shape(vec![
                TextToken::ParagraphRun(0),
                TextToken::Literal("a".into()),
            ]);
            with_cell(&mut shape, "TxtWidth", "1");
            shape.children.push(text_section(
                "Character",
                vec![row(0, "", vec![cell("Size", "0.25")])],
            ));
            shape.children.push(ShapeChild::Section(Section {
                name: "Paragraph".into(),
                index: None,
                del: false,
                children: vec![SectionChild::Row(row(
                    0,
                    "",
                    align
                        .map(|value| vec![cell("HorzAlign", &value.to_string())])
                        .unwrap_or_default(),
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
        shape.children.push(text_section(
            "Character",
            vec![row(0, "", vec![cell("Size", "0.25")])],
        ));
        shape.children.push(ShapeChild::Section(Section {
            name: "Paragraph".into(),
            index: None,
            del: false,
            children: vec![SectionChild::Row(row(
                0,
                "",
                vec![
                    cell("HorzAlign", "0"),
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
        assert!((lines[1].y - lines[0].y - 0.5).abs() < 1e-5);
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
                && diagnostic.detail == "unregistered font 'Calibri'"
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
    fn renders_indexed_shape_fill_and_stroke_colours() {
        let mut indexed = shape(1, 1.0, 1.0);
        for child in &mut indexed.children {
            if let ShapeChild::Cell(value) = child
                && matches!(value.name.as_str(), "FillForegnd" | "LineColor")
            {
                *value = cell(&value.name, "3");
            }
        }
        let mut package = package(vec![indexed]);
        package.colors = serde_json::from_value(serde_json::json!([
            {"name":"ColorEntry","attributes":[["IX","3"],["RGB","010203"]],"children":[]}
        ]))
        .unwrap();
        let list = Renderer::default().layout_page(&package, "page").unwrap();
        let Primitive::Shape { fill, stroke, .. } = &list.primitives[0] else {
            panic!("indexed shape did not render");
        };
        assert_eq!(
            fill,
            &Some(Paint::Solid {
                color: "#010203".into()
            })
        );
        assert_eq!(stroke.as_ref().unwrap().color, "#010203");
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
    fn group_local_origin_maps_to_the_group_box_corner() {
        let matrix = affine(vsdx_resolve::bounds_affine(vsdx_resolve::ShapeBounds {
            x: 10.0,
            y: 20.0,
            width: 8.0,
            height: 12.0,
            loc_pin_x: 4.0,
            loc_pin_y: 6.0,
            angle: 0.0,
            flip_x: false,
            flip_y: false,
        }));
        assert_point_close(matrix.apply_point(0.0, 0.0), (10.0, 20.0));
        assert_point_close(matrix.apply_point(4.0, 6.0), (14.0, 26.0));
        assert_point_close(matrix.apply_point(8.0, 12.0), (18.0, 32.0));
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
        let Primitive::Group {
            primitives,
            transform: group,
            ..
        } = &list.primitives[0]
        else {
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
            matches!(path[0], GeometryPathCommand::Move { x, y } if (x - 1.0).abs() < 1e-9 && y.abs() < 1e-9)
        );
        assert_point_close((group.a, group.b), (0.0, -1.0));
        assert_point_close((group.c, group.d), (-1.0, 0.0));
        assert_point_close((group.e, group.f), (10.0, 20.0));
        let GeometryPathCommand::Move { x, y } = path[0] else {
            unreachable!()
        };
        assert_point_close(
            group.compose(*transform).apply_point(x as f32, y as f32),
            (10.0, 19.0),
        );
    }

    #[test]
    fn deeply_nested_groups_preserve_display_list_hierarchy() {
        let nested = group(
            1,
            1.0,
            1.0,
            vec![group(
                2,
                2.0,
                2.0,
                vec![group(
                    3,
                    3.0,
                    3.0,
                    vec![group(
                        4,
                        4.0,
                        4.0,
                        vec![group(5, 5.0, 5.0, vec![shape(6, 6.0, 6.0)])],
                    )],
                )],
            )],
        );
        let list = render(vec![nested]);
        let Primitive::Group {
            id,
            z_order,
            primitives: level_2,
            ..
        } = &list.primitives[0]
        else {
            unreachable!()
        };
        assert_eq!((id, z_order), (&"page:1".into(), &10));
        let Primitive::Group {
            id,
            z_order,
            primitives: level_3,
            ..
        } = &level_2[0]
        else {
            unreachable!()
        };
        assert_eq!((id, z_order), (&"page:2".into(), &9));
        let Primitive::Group {
            id,
            z_order,
            primitives: level_4,
            ..
        } = &level_3[0]
        else {
            unreachable!()
        };
        assert_eq!((id, z_order), (&"page:3".into(), &8));
        let Primitive::Group {
            id,
            z_order,
            primitives: level_5,
            ..
        } = &level_4[0]
        else {
            unreachable!()
        };
        assert_eq!((id, z_order), (&"page:4".into(), &7));
        let Primitive::Group {
            id,
            z_order,
            primitives: leaf,
            ..
        } = &level_5[0]
        else {
            unreachable!()
        };
        assert_eq!((id, z_order), (&"page:5".into(), &6));
        assert!(
            matches!(&leaf[0], Primitive::Shape { id, z_order, .. } if id == "page:6" && *z_order == 5)
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
        let Primitive::Group {
            primitives,
            transform: group,
            ..
        } = &list.primitives[0]
        else {
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
        let page = group.compose(*transform);
        let (mut left, mut top, mut bounds_width, mut bounds_height) = (*x, *y, *width, *height);
        transform_rect(
            &mut left,
            &mut top,
            &mut bounds_width,
            &mut bounds_height,
            page,
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
            canvas(page.apply_point(stop.x, stop.y + height * 0.25)),
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
        rotated.children.push(text_section(
            "Character",
            vec![row(0, "", vec![cell("Size", "1")])],
        ));
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
        let Primitive::Group {
            primitives,
            transform: group,
            ..
        } = &list.primitives[0]
        else {
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
        assert!((group.b - std::f32::consts::FRAC_1_SQRT_2).abs() < 1e-5);
        let page = group.compose(*transform);
        let corners =
            [(0.0, 0.0), (1.0, 0.0), (0.0, 1.0), (1.0, 1.0)].map(|(x, y)| page.apply_point(x, y));
        let expected = [(1.0, 2.0), (2.0, 2.0), (1.0, 3.0), (2.0, 3.0)].map(|(x, y)| {
            Affine {
                a: std::f32::consts::FRAC_1_SQRT_2,
                b: std::f32::consts::FRAC_1_SQRT_2,
                c: -std::f32::consts::FRAC_1_SQRT_2,
                d: std::f32::consts::FRAC_1_SQRT_2,
                e: 10.0,
                f: 20.0,
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
        let Primitive::Group {
            primitives,
            transform: outer_transform,
            ..
        } = &list.primitives[0]
        else {
            unreachable!()
        };
        let Primitive::Group {
            primitives,
            transform: inner_transform,
            ..
        } = &primitives[0]
        else {
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
            e: 10.0,
            f: 20.0,
        };
        let inner = Affine {
            a: 0.0,
            b: -1.0,
            c: -1.0,
            d: 0.0,
            e: 2.0,
            f: 3.0,
        };
        assert_point_close((outer_transform.a, outer_transform.b), (outer.a, outer.b));
        assert_point_close((outer_transform.c, outer_transform.d), (outer.c, outer.d));
        assert_point_close((outer_transform.e, outer_transform.f), (outer.e, outer.f));
        assert_point_close((inner_transform.a, inner_transform.b), (inner.a, inner.b));
        assert_point_close((inner_transform.c, inner_transform.d), (inner.c, inner.d));
        assert_point_close((inner_transform.e, inner_transform.f), (inner.e, inner.f));
        assert_point_close((x as f32, y as f32), (1.0, 0.0));
        let page = outer_transform.compose(*inner_transform);
        assert_point_close(
            page.apply_point(x as f32, y as f32),
            outer.compose(inner).apply_point(1.0, 0.0),
        );
    }

    fn group_child_page_path(width: &str, height: &str, children: Vec<Shape>) -> Vec<(f32, f32)> {
        let mut group = group(1, 10.0, 20.0, children);
        with_cell(&mut group, "Width", width);
        with_cell(&mut group, "Height", height);
        with_cell(
            &mut group,
            "Angle",
            &std::f64::consts::FRAC_PI_2.to_string(),
        );
        with_cell(&mut group, "FlipX", "1");
        let list = render(vec![group]);
        let Primitive::Group {
            primitives,
            transform: group,
            ..
        } = &list.primitives[0]
        else {
            unreachable!()
        };
        let Primitive::Shape {
            path, transform, ..
        } = &primitives[0]
        else {
            unreachable!()
        };
        let page = group.compose(*transform);
        path.iter()
            .filter_map(|command| match command {
                GeometryPathCommand::Move { x, y } | GeometryPathCommand::Line { x, y } => {
                    Some(page.apply_point(*x as f32, *y as f32))
                }
                _ => None,
            })
            .collect()
    }

    #[test]
    fn a_groups_own_size_never_scales_its_children() {
        let narrow = group_child_page_path("4", "3", vec![shape(2, 0.0, 0.0)]);
        let wide = group_child_page_path("9", "7", vec![shape(2, 0.0, 0.0)]);
        assert_eq!(narrow, wide);
        assert_point_close(narrow[1], (10.0, 19.0));
    }

    #[test]
    fn removing_a_sibling_leaves_the_remaining_child_where_it_was() {
        let both = group_child_page_path("4", "3", vec![shape(2, 0.0, 0.0), shape(3, 5.0, 6.0)]);
        let alone = group_child_page_path("4", "3", vec![shape(2, 0.0, 0.0)]);
        assert_eq!(both, alone);
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
                    diagnostics: Vec::new(),
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
                    diagnostics: Vec::new(),
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
                    diagnostics: Vec::new(),
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
                    diagnostics: Vec::new(),
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
            hit_test(&list, 96.0 * 4.5, 768.0 - 96.0 * 6.5),
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
    fn hidden_one_d_connector_emits_no_primitives() {
        let mut connector = shape(1, 1.0, 1.0);
        connector.children.extend([
            ShapeChild::Cell(cell("OneD", "1")),
            ShapeChild::Cell(cell("BeginX", "1")),
            ShapeChild::Cell(cell("BeginY", "1")),
            ShapeChild::Cell(cell("EndX", "4")),
            ShapeChild::Cell(cell("EndY", "1")),
            ShapeChild::Cell(cell("NoShow", "1")),
        ]);
        assert!(render(vec![connector]).primitives.is_empty());
    }

    #[test]
    fn dangling_glue_renders_connector_placeholder() {
        let package = glued_connector_package("Connections.X9");
        let connectivity = Resolver::new(&package)
            .resolve_page_connectivity("page")
            .unwrap();
        assert!(
            connectivity.connectors[&1].glue[0]
                .to
                .as_ref()
                .unwrap()
                .connection_point
                .is_none()
        );

        let list = Renderer::default().layout_page(&package, "page").unwrap();
        assert!(matches!(
            list.primitives.iter().find(|primitive| {
                matches!(primitive, Primitive::Placeholder { id, .. } if id == "page:1")
            }),
            Some(Primitive::Placeholder { id, .. }) if id == "page:1"
        ));
        assert!(
            !list.primitives.iter().any(
                |primitive| matches!(primitive, Primitive::Shape { id, .. } if id == "page:1")
            )
        );
    }

    #[test]
    fn resolved_glue_paints_connector() {
        let package = glued_connector_package("Connections.X1");
        let connectivity = Resolver::new(&package)
            .resolve_page_connectivity("page")
            .unwrap();
        assert!(
            connectivity.connectors[&1].glue[0]
                .to
                .as_ref()
                .unwrap()
                .connection_point
                .is_some()
        );

        let list = Renderer::default().layout_page(&package, "page").unwrap();
        assert!(
            list.primitives.iter().any(
                |primitive| matches!(primitive, Primitive::Shape { id, .. } if id == "page:1")
            )
        );
        assert!(!list.primitives.iter().any(
            |primitive| matches!(primitive, Primitive::Placeholder { id, .. } if id == "page:1")
        ));
    }

    fn free_connector(id: u32, begin: (f64, f64), end: (f64, f64)) -> Shape {
        let mut connector = shape(id, 1.0, 1.0);
        connector.children.retain(
            |child| !matches!(child, ShapeChild::Section(section) if section.name == "Geometry"),
        );
        connector.children.extend([
            ShapeChild::Cell(cell("OneD", "1")),
            ShapeChild::Cell(cell("BeginX", &begin.0.to_string())),
            ShapeChild::Cell(cell("BeginY", &begin.1.to_string())),
            ShapeChild::Cell(cell("EndX", &end.0.to_string())),
            ShapeChild::Cell(cell("EndY", &end.1.to_string())),
        ]);
        connector
    }

    fn jump_package(
        shapes: Vec<Shape>,
        page_cells: &[(&str, &str)],
        style_cells: &[(&str, &str)],
    ) -> VsdxPackage {
        let mut package = package(shapes);
        let sheet = package.page_sheets.get_mut(&1).unwrap();
        for (name, value) in page_cells {
            sheet.children.push(SheetChild::Cell(cell(name, value)));
        }
        if !style_cells.is_empty() {
            package.style_sheets.push(Sheet {
                id: Some(0),
                children: style_cells
                    .iter()
                    .map(|(name, value)| SheetChild::Cell(cell(name, value)))
                    .collect(),
                other_attrs: vec![("NameU".into(), "No Style".into())],
            });
        }
        package
    }

    fn jump_path(list: &VsdxDisplayList, id: u32) -> &[GeometryPathCommand] {
        let Primitive::Shape { path, .. } = shape_primitive(list, id) else {
            unreachable!()
        };
        path
    }

    fn quad_count(path: &[GeometryPathCommand]) -> usize {
        path.iter()
            .filter(|command| matches!(command, GeometryPathCommand::Quad { .. }))
            .count()
    }

    fn move_count(path: &[GeometryPathCommand]) -> usize {
        path.iter()
            .filter(|command| matches!(command, GeometryPathCommand::Move { .. }))
            .count()
    }

    fn crossing_shapes() -> Vec<Shape> {
        vec![
            free_connector(1, (0.0, 1.0), (4.0, 1.0)),
            free_connector(2, (2.0, 0.0), (2.0, 2.0)),
        ]
    }

    fn apex(path: &[GeometryPathCommand], axis: char, at: f64) -> Option<(f64, f64)> {
        path.iter().find_map(|command| match *command {
            GeometryPathCommand::Quad { x, y, .. }
                if (if axis == 'y' { y } else { x } - at).abs() < 1e-9 =>
            {
                Some((x, y))
            }
            _ => None,
        })
    }

    #[test]
    fn line_jump_bridges_the_more_horizontal_connector() {
        let package = jump_package(crossing_shapes(), &[("LineJumpCode", "1")], &[]);
        let list = Renderer::default().layout_page(&package, "page").unwrap();
        let horizontal = jump_path(&list, 1);
        assert_eq!(quad_count(horizontal), 2);
        assert_eq!(apex(horizontal, 'y', 1.025), Some((2.0, 1.025)));
        assert_eq!(jump_path(&list, 2).len(), 2);
    }

    #[test]
    fn line_jump_code_zero_suppresses_bridges() {
        let package = jump_package(crossing_shapes(), &[("LineJumpCode", "0")], &[]);
        let list = Renderer::default().layout_page(&package, "page").unwrap();
        assert_eq!(jump_path(&list, 1).len(), 2);
        assert_eq!(jump_path(&list, 2).len(), 2);
    }

    #[test]
    fn line_jump_page_sheet_overrides_the_stylesheet() {
        let package = jump_package(
            crossing_shapes(),
            &[("LineJumpCode", "0")],
            &[("LineJumpCode", "1")],
        );
        let list = Renderer::default().layout_page(&package, "page").unwrap();
        assert_eq!(quad_count(jump_path(&list, 1)), 0);
        assert_eq!(quad_count(jump_path(&list, 2)), 0);
    }

    #[test]
    fn line_jump_stylesheet_fallback_sizes_the_bridge() {
        let package = jump_package(
            crossing_shapes(),
            &[],
            &[
                ("LineJumpCode", "1"),
                ("LineJumpFactorX", "1"),
                ("LineToLineX", "0.2"),
            ],
        );
        let list = Renderer::default().layout_page(&package, "page").unwrap();
        assert_eq!(apex(jump_path(&list, 1), 'y', 1.1), Some((2.0, 1.1)));
    }

    #[test]
    fn line_jump_vertical_code_bridges_the_more_vertical_connector() {
        let package = jump_package(crossing_shapes(), &[("LineJumpCode", "2")], &[]);
        let list = Renderer::default().layout_page(&package, "page").unwrap();
        assert_eq!(jump_path(&list, 1).len(), 2);
        assert_eq!(apex(jump_path(&list, 2), 'x', 1.975), Some((1.975, 1.0)));
    }

    #[test]
    fn connector_never_deflects_the_bridge_to_the_other_connector() {
        let mut shapes = crossing_shapes();
        with_cell(&mut shapes[0], "ConLineJumpCode", "1");
        let package = jump_package(shapes, &[("LineJumpCode", "1")], &[]);
        let list = Renderer::default().layout_page(&package, "page").unwrap();
        assert_eq!(jump_path(&list, 1).len(), 2);
        assert_eq!(quad_count(jump_path(&list, 2)), 2);
    }

    #[test]
    fn connector_neither_suppresses_the_crossing() {
        let mut shapes = crossing_shapes();
        with_cell(&mut shapes[0], "ConLineJumpCode", "4");
        let package = jump_package(shapes, &[("LineJumpCode", "1")], &[]);
        let list = Renderer::default().layout_page(&package, "page").unwrap();
        assert_eq!(quad_count(jump_path(&list, 1)), 0);
        assert_eq!(quad_count(jump_path(&list, 2)), 0);
    }

    #[test]
    fn connector_always_forces_a_bridge_under_code_zero() {
        let mut shapes = crossing_shapes();
        with_cell(&mut shapes[0], "ConLineJumpCode", "2");
        let package = jump_package(shapes, &[("LineJumpCode", "0")], &[]);
        let list = Renderer::default().layout_page(&package, "page").unwrap();
        assert_eq!(quad_count(jump_path(&list, 1)), 2);
    }

    #[test]
    fn line_jump_gap_style_breaks_the_route() {
        let mut shapes = crossing_shapes();
        with_cell(&mut shapes[0], "ConLineJumpStyle", "2");
        let package = jump_package(shapes, &[("LineJumpCode", "1")], &[]);
        let list = Renderer::default().layout_page(&package, "page").unwrap();
        let horizontal = jump_path(&list, 1);
        assert_eq!(quad_count(horizontal), 0);
        assert_eq!(move_count(horizontal), 2);
        assert_eq!(jump_path(&list, 2).len(), 2);
    }

    #[test]
    fn line_jump_ignores_shared_endpoints() {
        let package = jump_package(
            vec![
                free_connector(1, (0.0, 1.0), (2.0, 1.0)),
                free_connector(2, (2.0, 1.0), (2.0, 3.0)),
            ],
            &[("LineJumpCode", "1")],
            &[],
        );
        let list = Renderer::default().layout_page(&package, "page").unwrap();
        assert_eq!(jump_path(&list, 1).len(), 2);
        assert_eq!(jump_path(&list, 2).len(), 2);
    }

    #[test]
    fn line_jump_display_order_codes_follow_z_order() {
        for (code, jumper) in [("4", 2), ("5", 1)] {
            let package = jump_package(crossing_shapes(), &[("LineJumpCode", code)], &[]);
            let list = Renderer::default().layout_page(&package, "page").unwrap();
            assert_eq!(quad_count(jump_path(&list, jumper)), 2, "code {code}");
            assert_eq!(
                quad_count(jump_path(&list, if jumper == 1 { 2 } else { 1 })),
                0,
                "code {code}"
            );
        }
    }

    #[test]
    fn line_jump_bridges_bent_connector_legs() {
        let mut bent = free_connector(1, (0.0, 0.0), (2.0, 2.0));
        with_cell(&mut bent, "ShapeRouteStyle", "1");
        let package = jump_package(
            vec![bent, free_connector(2, (1.0, -1.0), (1.0, 1.0))],
            &[("LineJumpCode", "1")],
            &[],
        );
        let list = Renderer::default().layout_page(&package, "page").unwrap();
        assert_eq!(apex(jump_path(&list, 1), 'y', 0.025), Some((1.0, 0.025)));
        assert_eq!(jump_path(&list, 2).len(), 2);
    }

    fn nested_shape(
        primitives: &[Primitive],
        id: &str,
        ancestors: Affine,
    ) -> Option<(Vec<GeometryPathCommand>, Affine)> {
        primitives.iter().find_map(|primitive| match primitive {
            Primitive::Shape {
                id: actual,
                path,
                transform,
                ..
            } if actual == id => Some((path.clone(), ancestors.compose(*transform))),
            Primitive::Group {
                primitives,
                transform,
                ..
            } => nested_shape(primitives, id, ancestors.compose(*transform)),
            _ => None,
        })
    }

    fn page_jump_apexes(list: &VsdxDisplayList, id: u32) -> Vec<(f32, f32)> {
        let (path, matrix) =
            nested_shape(&list.primitives, &format!("page:{id}"), Affine::identity()).unwrap();
        path.iter()
            .filter_map(|command| match *command {
                GeometryPathCommand::Quad { x, y, .. } => {
                    Some(matrix.apply_point(x as f32, y as f32))
                }
                _ => None,
            })
            .collect()
    }

    fn grouped_crossing_package(
        angle: &str,
        local: ((f64, f64), (f64, f64)),
        code: &str,
    ) -> VsdxPackage {
        let mut host = group(10, 2.0, 2.0, vec![free_connector(2, local.0, local.1)]);
        with_cell(&mut host, "Angle", angle);
        jump_package(
            vec![free_connector(1, (0.0, 1.0), (4.0, 1.0)), host],
            &[("LineJumpCode", code)],
            &[],
        )
    }

    #[test]
    fn line_jump_crosses_a_grouped_connector_in_page_space() {
        let package = grouped_crossing_package("0", ((0.0, -1.5), (0.0, 0.5)), "1");
        let list = Renderer::default().layout_page(&package, "page").unwrap();
        assert_eq!(apex(jump_path(&list, 1), 'y', 1.025), Some((2.0, 1.025)));
        assert!(page_jump_apexes(&list, 2).is_empty());
    }

    #[test]
    fn line_jump_on_a_connector_inside_a_rotated_group_lands_in_page_space() {
        let package = grouped_crossing_package(
            &std::f64::consts::FRAC_PI_2.to_string(),
            ((-1.5, 0.0), (0.5, 0.0)),
            "2",
        );
        let list = Renderer::default().layout_page(&package, "page").unwrap();
        assert_eq!(jump_path(&list, 1).len(), 2);
        let apexes = page_jump_apexes(&list, 2);
        assert_eq!(apexes.len(), 2, "{apexes:?}");
        assert_point_close(apexes[0], (1.975, 1.0));
    }

    #[test]
    fn connector_route_styles_bend_or_run_straight() {
        use GeometryPathCommand::{Line, Move};
        let begin = ScenePoint { x: 1.0, y: 1.0 };
        let end = ScenePoint { x: 4.0, y: 3.0 };
        let horizontal = vec![
            Move { x: 1.0, y: 1.0 },
            Line { x: 4.0, y: 1.0 },
            Line { x: 4.0, y: 3.0 },
        ];
        let vertical = vec![
            Move { x: 1.0, y: 1.0 },
            Line { x: 1.0, y: 3.0 },
            Line { x: 4.0, y: 3.0 },
        ];
        let direct = vec![Move { x: 1.0, y: 1.0 }, Line { x: 4.0, y: 3.0 }];
        for style in [1.0, 4.0, 6.0, 8.0, 9.0, 21.0] {
            assert_eq!(connector_route(begin, end, style), horizontal);
        }
        for style in [3.0, 5.0, 7.0, 10.0, 12.0, 14.0, 17.0, 19.0, 22.0] {
            assert_eq!(connector_route(begin, end, style), vertical);
        }
        for style in [0.0, 2.0, 16.0] {
            assert_eq!(connector_route(begin, end, style), direct);
        }
        for style in [-1.0, 23.0, 1e30, f64::NAN, f64::INFINITY, f64::NEG_INFINITY] {
            assert_eq!(connector_route(begin, end, style), horizontal);
        }
        let aligned = ScenePoint { x: 1.0, y: 3.0 };
        assert_eq!(
            connector_route(begin, aligned, 1.0),
            vec![Move { x: 1.0, y: 1.0 }, Line { x: 1.0, y: 3.0 }]
        );
    }

    fn unglued_connector(style: Option<&str>) -> Shape {
        let mut connector = shape(1, 1.0, 1.0);
        connector.children.retain(
            |child| !matches!(child, ShapeChild::Section(section) if section.name == "Geometry"),
        );
        connector.children.extend([
            ShapeChild::Cell(cell("OneD", "1")),
            ShapeChild::Cell(cell("BeginX", "1")),
            ShapeChild::Cell(cell("BeginY", "1")),
            ShapeChild::Cell(cell("EndX", "4")),
            ShapeChild::Cell(cell("EndY", "3")),
        ]);
        if let Some(style) = style {
            with_cell(&mut connector, "ShapeRouteStyle", style);
        }
        connector
    }

    fn connector_path(package: &VsdxPackage, page: &str) -> Vec<GeometryPathCommand> {
        let list = Renderer::default().layout_page(package, page).unwrap();
        list.primitives
            .iter()
            .find_map(|primitive| match primitive {
                Primitive::Shape { id, path, .. } if id == &format!("{page}:1") => {
                    Some(path.clone())
                }
                _ => None,
            })
            .unwrap()
    }

    #[test]
    fn right_angle_shape_route_style_bends_unglued_connector() {
        use GeometryPathCommand::{Line, Move};
        let list = render(vec![unglued_connector(Some("1"))]);
        let Primitive::Shape { path, .. } = shape_primitive(&list, 1) else {
            unreachable!()
        };
        assert_eq!(
            *path,
            vec![
                Move { x: 1.0, y: 1.0 },
                Line { x: 4.0, y: 1.0 },
                Line { x: 4.0, y: 3.0 },
            ]
        );
    }

    #[test]
    fn page_route_style_applies_without_shape_override() {
        use GeometryPathCommand::{Line, Move};
        let mut package = package(vec![unglued_connector(None)]);
        package
            .page_sheets
            .get_mut(&1)
            .unwrap()
            .children
            .push(SheetChild::Cell(cell("RouteStyle", "5")));
        assert_eq!(
            connector_path(&package, "page"),
            vec![
                Move { x: 1.0, y: 1.0 },
                Line { x: 1.0, y: 3.0 },
                Line { x: 4.0, y: 3.0 },
            ]
        );
    }

    #[test]
    fn one_d_shape_without_any_route_style_runs_straight() {
        use GeometryPathCommand::{Line, Move};
        let package = package(vec![unglued_connector(None)]);
        assert_eq!(
            connector_path(&package, "page"),
            vec![Move { x: 1.0, y: 1.0 }, Line { x: 4.0, y: 3.0 }]
        );
    }

    #[test]
    fn filed_connector_geometry_wins_over_synthesized_route() {
        use GeometryPathCommand::{Line, Move};
        let mut connector = unglued_connector(Some("2"));
        connector.children.push(ShapeChild::Section(Section {
            name: "Geometry".into(),
            index: None,
            del: false,
            children: vec![
                row(2, "LineTo", vec![cell("X", "0.5"), cell("Y", "0")]),
                row(3, "LineTo", vec![cell("X", "0.5"), cell("Y", "2")]),
            ]
            .into_iter()
            .map(SectionChild::Row)
            .collect(),
            other_attrs: vec![],
        }));
        let list = render(vec![connector]);
        let Primitive::Shape { path, .. } = shape_primitive(&list, 1) else {
            unreachable!()
        };
        assert_eq!(
            *path,
            vec![
                Move { x: 1.0, y: 1.0 },
                Line { x: 1.5, y: 1.0 },
                Line { x: 1.5, y: 3.0 },
                Line { x: 4.0, y: 3.0 },
            ]
        );
    }

    #[test]
    fn filed_connector_geometry_without_segments_falls_back_to_synthesized_route() {
        use GeometryPathCommand::{Line, Move};
        let mut connector = unglued_connector(Some("1"));
        connector.children.push(ShapeChild::Section(Section {
            name: "Geometry".into(),
            index: None,
            del: false,
            children: vec![row(0, "MoveTo", vec![cell("X", "0"), cell("Y", "0")])]
                .into_iter()
                .map(SectionChild::Row)
                .collect(),
            other_attrs: vec![],
        }));
        let list = render(vec![connector]);
        let Primitive::Shape { path, .. } = shape_primitive(&list, 1) else {
            unreachable!()
        };
        assert_eq!(
            *path,
            vec![
                Move { x: 1.0, y: 1.0 },
                Line { x: 4.0, y: 1.0 },
                Line { x: 4.0, y: 3.0 },
            ]
        );
    }

    #[test]
    fn closed_filed_connector_geometry_falls_back_to_synthesized_route() {
        use GeometryPathCommand::{Line, Move};
        let mut connector = unglued_connector(Some("1"));
        connector.children.push(ShapeChild::Section(Section {
            name: "Geometry".into(),
            index: None,
            del: false,
            children: vec![
                row(0, "MoveTo", vec![cell("X", "0"), cell("Y", "0")]),
                row(1, "LineTo", vec![cell("X", "0.5"), cell("Y", "0")]),
                row(2, "Close", vec![]),
            ]
            .into_iter()
            .map(SectionChild::Row)
            .collect(),
            other_attrs: vec![],
        }));
        let list = render(vec![connector]);
        let Primitive::Shape { path, .. } = shape_primitive(&list, 1) else {
            unreachable!()
        };
        assert_eq!(
            *path,
            vec![
                Move { x: 1.0, y: 1.0 },
                Line { x: 4.0, y: 1.0 },
                Line { x: 4.0, y: 3.0 },
            ]
        );
    }

    fn route_fixture_path(page: &str) -> Vec<GeometryPathCommand> {
        let source = include_bytes!("../../vsdx-parse/tests/fixtures/connector-route-style.vsdx");
        let package = vsdx_parse::parse_vsdx(source).unwrap();
        connector_path(&package, page)
    }

    #[test]
    fn shape_route_style_overrides_page_in_fixture() {
        use GeometryPathCommand::{Line, Move};
        assert_eq!(
            route_fixture_path("visio/pages/page1.xml"),
            vec![
                Move { x: 1.0, y: 1.0 },
                Line { x: 4.0, y: 1.0 },
                Line { x: 4.0, y: 3.0 },
            ]
        );
    }

    #[test]
    fn page_route_style_applies_in_fixture() {
        use GeometryPathCommand::{Line, Move};
        assert_eq!(
            route_fixture_path("visio/pages/page2.xml"),
            vec![
                Move { x: 1.0, y: 1.0 },
                Line { x: 1.0, y: 3.0 },
                Line { x: 4.0, y: 3.0 },
            ]
        );
    }

    #[test]
    fn straight_shape_route_style_runs_direct_in_fixture() {
        use GeometryPathCommand::{Line, Move};
        assert_eq!(
            route_fixture_path("visio/pages/page3.xml"),
            vec![Move { x: 1.0, y: 1.0 }, Line { x: 4.0, y: 3.0 }]
        );
    }

    #[test]
    fn two_subpath_filed_connector_geometry_falls_back_to_synthesized_route() {
        use GeometryPathCommand::{Line, Move};
        assert_eq!(
            route_fixture_path("visio/pages/page4.xml"),
            vec![
                Move { x: 1.0, y: 1.0 },
                Line { x: 4.0, y: 1.0 },
                Line { x: 4.0, y: 3.0 },
            ]
        );
    }

    #[test]
    fn unresolved_foreign_data_placeholders_while_colours_default_paint() {
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
        let Primitive::Shape {
            fill,
            stroke,
            diagnostics,
            path,
            ..
        } = &list.primitives[1]
        else {
            panic!("unresolvable fill did not render");
        };
        assert!(!path.is_empty());
        assert_eq!(
            fill,
            &Some(Paint::Solid {
                color: "#000000".into()
            })
        );
        assert_eq!(stroke.as_ref().unwrap().color, "#040506");
        assert!(
            diagnostics
                .iter()
                .any(|diagnostic| diagnostic.code == "unresolvable-fill-colour"
                    && diagnostic.detail
                        == "unresolvable fill colour: missing colour cell FillForegnd"
                    && diagnostic.category == DiagnosticCategory::Fidelity)
        );
    }

    #[test]
    fn dangling_media_placeholders_while_non_finite_stroke_width_defaults() {
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
        let Primitive::Shape {
            fill,
            stroke,
            diagnostics,
            ..
        } = &list.primitives[1]
        else {
            panic!("non-finite stroke width did not render");
        };
        assert_eq!(
            fill,
            &Some(Paint::Solid {
                color: "#010203".into()
            })
        );
        let stroke = stroke.as_ref().unwrap();
        assert_eq!(stroke.color, "#040506");
        assert_eq!(stroke.width, 0.01);
        assert_eq!(diagnostics.len(), 1);
        assert_eq!(diagnostics[0].code, "unresolvable-stroke-width");
        assert_eq!(
            diagnostics[0].detail,
            "unresolvable stroke width: non-finite LineWeight"
        );
        assert_eq!(diagnostics[0].category, DiagnosticCategory::Fidelity);
    }

    #[test]
    fn unresolvable_stroke_defaults_stroke_and_keeps_fill() {
        let mut lined = shape(1, 1.0, 1.0);
        lined.children.retain(
            |child| !matches!(child, ShapeChild::Cell(Cell { name, .. }) if name == "LineColor"),
        );
        let list = render(vec![lined]);
        let Primitive::Shape {
            fill,
            stroke,
            diagnostics,
            ..
        } = &list.primitives[0]
        else {
            panic!("unresolvable stroke did not render");
        };
        assert_eq!(
            fill,
            &Some(Paint::Solid {
                color: "#010203".into()
            })
        );
        assert_eq!(
            stroke,
            &Some(Stroke {
                color: "#000000".into(),
                width: 0.02,
                dashed: false,
            })
        );
        assert_eq!(diagnostics.len(), 1);
        assert_eq!(diagnostics[0].code, "unresolvable-stroke-colour");
        assert_eq!(
            diagnostics[0].detail,
            "unresolvable stroke colour: missing colour cell LineColor"
        );
        assert_eq!(diagnostics[0].category, DiagnosticCategory::Fidelity);
    }

    #[test]
    fn defaulted_paint_diagnostics_survive_the_display_list() {
        let mut coloured = shape(1, 1.0, 1.0);
        coloured.children.retain(
            |child| !matches!(child, ShapeChild::Cell(Cell { name, .. }) if name == "FillForegnd"),
        );
        let list = render(vec![shape(2, 4.0, 1.0), coloured]);
        let json = serde_json::to_value(&list).unwrap();
        let shapes = json["primitives"].as_array().unwrap();
        assert!(
            shapes[0].get("diagnostics").is_none(),
            "resolved paint must not grow the contract"
        );
        let diagnostics = shapes[1]["diagnostics"].as_array().unwrap();
        assert_eq!(diagnostics.len(), 1);
        assert_eq!(diagnostics[0]["code"], "unresolvable-fill-colour");
        assert_eq!(diagnostics[0]["category"], "fidelity");
        assert_eq!(
            diagnostics[0]["detail"],
            "unresolvable fill colour: missing colour cell FillForegnd"
        );
        let round_tripped: VsdxDisplayList = serde_json::from_value(json).unwrap();
        assert_eq!(round_tripped, list);
    }

    #[test]
    fn unresolvable_connector_colour_keeps_route() {
        let mut package = glued_connector_package("Connections.X1");
        let SheetChild::Shapes(shapes) = package
            .page_contents
            .get_mut("page")
            .unwrap()
            .children
            .iter_mut()
            .find(|child| matches!(child, SheetChild::Shapes(_)))
            .unwrap()
        else {
            unreachable!()
        };
        let connector = shapes
            .iter_mut()
            .find_map(|child| match child {
                ShapesChild::Shape(shape) if shape.id == 1 => Some(shape),
                _ => None,
            })
            .unwrap();
        connector.children.retain(
            |child| !matches!(child, ShapeChild::Cell(Cell { name, .. }) if name == "LineColor"),
        );
        connector
            .children
            .push(ShapeChild::Cell(formula("LineColor", "THEMEVAL(999)")));
        let list = Renderer::default().layout_page(&package, "page").unwrap();
        assert!(!list.primitives.iter().any(
            |primitive| matches!(primitive, Primitive::Placeholder { id, .. } if id == "page:1")
        ));
        let Primitive::Shape {
            stroke,
            diagnostics,
            path,
            ..
        } = shape_primitive(&list, 1)
        else {
            unreachable!()
        };
        assert!(!path.is_empty());
        assert_eq!(stroke.as_ref().unwrap().color, "#000000");
        assert!(
            diagnostics
                .iter()
                .any(|diagnostic| diagnostic.code == "unresolvable-stroke-colour"
                    && diagnostic.detail.contains("THEMEVAL colour-scheme index"))
        );
    }

    #[test]
    fn unsupported_colour_reason_survives_as_a_paint_diagnostic() {
        let mut coloured = shape(1, 1.0, 1.0);
        coloured.children.retain(
            |child| !matches!(child, ShapeChild::Cell(Cell { name, .. }) if name == "FillForegnd"),
        );
        coloured
            .children
            .push(ShapeChild::Cell(formula("FillForegnd", "THEMEVAL(999)")));
        let list = render(vec![coloured]);
        let Primitive::Shape {
            fill, diagnostics, ..
        } = &list.primitives[0]
        else {
            panic!("unsupported fill colour did not render");
        };
        assert_eq!(
            fill,
            &Some(Paint::Solid {
                color: "#000000".into()
            })
        );
        assert!(
            diagnostics.iter().any(|diagnostic| diagnostic.code == "unresolvable-fill-colour"
                && diagnostic.detail
                    == "unresolvable fill colour: THEMEVAL colour-scheme index must be 1 through 8")
        );
    }

    /// Theme fills resolve through the package theme instead of placeholdering.
    #[test]
    fn themed_fills_paint_from_the_package_theme() {
        let mut named = shape(1, 1.0, 1.0);
        named.children.retain(
            |child| !matches!(child, ShapeChild::Cell(Cell { name, .. }) if name == "FillForegnd"),
        );
        named.children.push(ShapeChild::Cell(formula(
            "FillForegnd",
            "THEMEVAL(\"AccentColor1\")",
        )));
        let mut host = shape(2, 2.0, 1.0);
        host.children.retain(
            |child| !matches!(child, ShapeChild::Cell(Cell { name, .. }) if name == "FillForegnd"),
        );
        host.children
            .push(ShapeChild::Cell(formula("FillForegnd", "THEMEVAL()")));
        let mut package = package(vec![named, host]);
        let mut theme = ooxml_drawingml::Theme::default();
        theme.color_scheme.accent1 = "112233".into();
        package.themes.insert(1, theme);
        let list = Renderer::default().layout_page(&package, "page").unwrap();
        assert_eq!(list.primitives.len(), 2);
        for primitive in &list.primitives {
            assert!(
                matches!(primitive, Primitive::Shape { fill: Some(Paint::Solid { color }), .. } if color == "#112233"),
                "{primitive:?}"
            );
        }
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
    fn missing_page_sheet_preserves_page_dimension_error() {
        let mut page = package(vec![shape(1, 1.0, 1.0)]);
        page.page_sheets.remove(&1);
        assert!(matches!(
            Renderer::default().layout_page(&page, "page"),
            Err(RenderError::PageDimensions(message)) if message == "PageHeight is unavailable"
        ));
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
        let payload = r#"{"contractVersion":2,"width":0.0,"height":0.0,"paintTransform":{"a":1.0,"b":0.0,"c":0.0,"d":1.0,"e":0.0,"f":0.0},"primitives":[]}"#;
        assert!(serde_json::from_str::<VsdxDisplayList>(payload).is_err());
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
    fn relative_geometry_rows_are_scaled_and_transformed_in_the_display_list() {
        let package = vsdx_parse::parse_vsdx(include_bytes!(
            "../../vsdx-parse/tests/fixtures/geometry-relative-rows.vsdx"
        ))
        .unwrap();
        let list = Renderer::default()
            .layout_page(&package, &package.page_part_paths[0])
            .unwrap();
        let path_for = |shape_id| {
            let Primitive::Shape { path, .. } = list
                .primitives
                .iter()
                .find(|primitive| matches!(primitive, Primitive::Shape { id, .. } if id.ends_with(shape_id)))
                .unwrap()
            else {
                unreachable!()
            };
            path
        };
        assert_eq!(
            path_for(":1").as_slice(),
            &[
                GeometryPathCommand::Move { x: 5.0, y: 3.5 },
                GeometryPathCommand::Line { x: 9.0, y: 3.5 },
                GeometryPathCommand::Line { x: 9.0, y: 6.5 },
                GeometryPathCommand::Line { x: 5.0, y: 6.5 },
            ]
        );
        assert_eq!(
            path_for(":2").as_slice(),
            &[
                GeometryPathCommand::Move { x: 12.0, y: 3.5 },
                GeometryPathCommand::Line { x: 16.0, y: 3.5 },
                GeometryPathCommand::Move { x: 12.0, y: 6.5 },
                GeometryPathCommand::Line { x: 16.0, y: 6.5 },
            ]
        );
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
            primitives: outer,
            transform: outer_transform,
            ..
        } = &list.primitives[0]
        else {
            unreachable!()
        };
        let Primitive::Group {
            primitives: inner,
            transform: inner_transform,
            ..
        } = &outer[0]
        else {
            unreachable!()
        };
        let group = outer_transform.compose(*inner_transform);
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
        // The composed group affine M = outer * inner was calculated by hand as
        // [[-0.258819, 0.9659258, 8.224235], [-0.9659258, -0.258819, 10.3686084]].
        // Children stay in group-local coordinates; every page-space expectation below is
        // a direct substitution into M.
        assert_point_close((group.a, group.b), (-0.258819, -0.9659258));
        assert_point_close((group.c, group.d), (0.9659258, -0.258819));
        assert_point_close((group.e, group.f), (8.224235, 10.368608));
        assert_point_close((x as f32, y as f32), (0.0, 0.0));
        assert_point_close(group.apply_point(x as f32, y as f32), (8.224235, 10.368608));
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
        // The text box's own x/y/width/height stay in the shape's local, pre-transform frame;
        // composing the group chain with the primitive's own transform carries the local text
        // box, its lines and its caret stops into scene coordinates.
        let page = group.compose(*transform);
        assert_point_close((*x, *y), (0.0, 0.0));
        assert_point_close((*width, *height), (1.0, 1.0));
        assert_eq!(*transform, Affine::identity());
        assert_point_close(
            page.apply_point(lines[0].x, lines[0].y),
            (8.429537, 9.975697),
        );
        assert_point_close(
            page.apply_point(lines[0].caret_stops[1].x, lines[0].caret_stops[1].y),
            (8.411563, 9.908618),
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
        let page = group.compose(*transform);
        let corners = [
            page.apply_point(*x, *y),
            page.apply_point(*x + *width, *y),
            page.apply_point(*x, *y + *height),
            page.apply_point(*x + *width, *y + *height),
        ];
        assert_point_close(corners[0], (8.672523, 8.177938));
        assert_point_close(corners[1], (8.413704, 7.212012));
        assert_point_close(corners[2], (10.604374, 7.6603));
        assert_point_close(corners[3], (10.345555, 6.694374));
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
        assert_point_close((*x, *y), (1.0, 3.0));
        assert_point_close((*width, *height), (1.0, 1.0));
        let z_orders = inner.iter().map(z_order).collect::<Vec<_>>();
        assert_eq!(z_orders, vec![2, 3, 4, 5]);
        let inside = group.apply_point(0.05, 0.05);
        assert_eq!(
            hit_test(&list, inside.0 * 96.0, (11.0 - inside.1) * 96.0),
            Some(HitTestResult::Text {
                shape_id: "visio/pages/page1.xml:3".into(),
                position: 0,
            })
        );
    }

    #[test]
    fn corpus_dangling_glue_renders_placeholders() {
        let Ok(directory) = std::env::var("VSDX_CORPUS_DIR") else {
            eprintln!(
                "warning: skipping corpus dangling-glue render test; VSDX_CORPUS_DIR is unset"
            );
            return;
        };
        let mut dangling = 0;
        let mut expected = std::collections::BTreeSet::new();
        let mut placeholders = std::collections::BTreeSet::new();
        let mut painted = std::collections::BTreeSet::new();
        for file in ["lichtsysteme.vsdx", "soundplan.vsdx"] {
            let package = vsdx_parse::parse_vsdx(
                &std::fs::read(std::path::Path::new(&directory).join(file)).unwrap(),
            )
            .unwrap();
            let resolver = Resolver::new(&package);
            let renderer = Renderer::default();
            for page in &package.page_part_paths {
                let connectivity = resolver.resolve_page_connectivity(page).unwrap();
                for connector in connectivity.connectors.values() {
                    let unresolved = connector
                        .glue
                        .iter()
                        .filter(|glue| {
                            glue.to
                                .as_ref()
                                .and_then(|target| target.connection_point.as_ref())
                                .is_none()
                        })
                        .count();
                    dangling += unresolved;
                    if unresolved != 0 {
                        expected.insert(format!("{page}:{}", connector.shape_id));
                    }
                }
                collect_connector_primitives(
                    &renderer.layout_page(&package, page).unwrap().primitives,
                    &mut placeholders,
                    &mut painted,
                );
            }
        }
        assert_eq!(dangling, 0, "corpus dangling-glue record count changed");
        assert_eq!(placeholders, expected);
        assert!(painted.is_disjoint(&expected));
    }

    fn collect_connector_primitives(
        primitives: &[Primitive],
        placeholders: &mut std::collections::BTreeSet<String>,
        painted: &mut std::collections::BTreeSet<String>,
    ) {
        for primitive in primitives {
            match primitive {
                Primitive::Placeholder { id, reason, .. } => {
                    if reason.starts_with("connector route cannot be computed:") {
                        assert!(!reason.is_empty());
                        placeholders.insert(id.clone());
                    }
                }
                Primitive::Shape { id, .. } => {
                    painted.insert(id.clone());
                }
                Primitive::Group { primitives, .. } => {
                    collect_connector_primitives(primitives, placeholders, painted);
                }
                Primitive::Image { .. } | Primitive::TextBox { .. } => {}
            }
        }
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
                let expected = expected
                    .difference(&layer_hidden_shape_ids(&package, page))
                    .cloned()
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
    fn group_subshapes_render_master_geometry_and_text_with_page_formatting() {
        let package = vsdx_parse::parse_vsdx(include_bytes!(
            "../../vsdx-parse/tests/fixtures/group-master-shape.vsdx"
        ))
        .unwrap();
        let page = &package.page_part_paths[0];
        let renderer = Renderer::default();
        let list = renderer.layout_page(&package, page).unwrap();
        let mut boxes = BTreeMap::new();
        text_boxes_by_id(&list.primitives, &mut boxes);
        assert_eq!(boxes.len(), 2);
        for id in [2, 4] {
            let Primitive::TextBox {
                width,
                height,
                paragraphs,
                ..
            } = boxes[&format!("{page}:{id}")]
            else {
                panic!("missing text box")
            };
            assert_eq!((*width, *height), (2.0, 1.0));
            let runs = paragraphs
                .iter()
                .flat_map(|paragraph| &paragraph.runs)
                .collect::<Vec<_>>();
            assert_eq!(runs.len(), 1);
            assert_eq!(runs[0].text, "group label");
            assert_eq!(runs[0].size_in, 0.25);
        }
        let mut rendered = std::collections::BTreeSet::new();
        assert_resolved_text(
            &renderer,
            &package,
            page,
            &list.primitives,
            &mut rendered,
            &mut TextCorpusStats::default(),
        );
        assert_eq!(rendered.len(), 2);
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

    fn layer_hidden_shape_ids(
        package: &VsdxPackage,
        page: &str,
    ) -> std::collections::BTreeSet<String> {
        let resolver = Resolver::new(package);
        let layers = vsdx_resolve::page_layers(package, page);
        let mut hidden = std::collections::BTreeSet::new();
        for shape in package.page_contents[page].shapes() {
            collect_layer_hidden(&resolver, page, &layers, shape, false, &mut hidden);
        }
        hidden
    }

    fn collect_layer_hidden(
        resolver: &Resolver<'_>,
        page: &str,
        layers: &[vsdx_resolve::PageLayer],
        shape: &Shape,
        ancestor_hidden: bool,
        hidden: &mut std::collections::BTreeSet<String>,
    ) {
        let hidden_here = ancestor_hidden
            || resolver
                .resolve_shape(page, shape.id)
                .is_ok_and(|resolved| vsdx_resolve::shape_hidden_by_layers(&resolved, layers));
        if hidden_here {
            hidden.insert(format!("{page}:{}", shape.id));
        }
        for child in shape.shapes() {
            collect_layer_hidden(resolver, page, layers, child, hidden_here, hidden);
        }
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
                contents,
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
        lookup: &Sheet,
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
        if vsdx_resolve::shape_hidden_by_layers(
            &resolved,
            &vsdx_resolve::page_layers(package, page_part),
        ) {
            return;
        }
        let tokens = resolver
            .resolve_text_in_context(shape, lookup, &resolved)
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
                renderer, package, resolver, references, lookup, page_part, child, text_boxes,
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
