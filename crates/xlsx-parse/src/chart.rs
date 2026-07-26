//! charts as first-class package content: discovery through the drawing layer,
//! the `xdr:` anchor model, the `c:f` references the workbook owns, an adapter
//! that lets the shared DrawingML parser read a chart part, and in-place
//! patching of both parts when a structural edit moves what they name.

use std::ops::Range;

use ooxml_drawingml::chart::{ChartSpace, ChartXml, parse_chart_space};
use ooxml_drawingml::{
    Theme as DrawingTheme, parse_color_value, resolve_color_value_to_hex_with_theme,
};
use xlsx_model::addr::{MAX_COLS, MAX_ROWS};
use xlsx_model::styles::Theme;
use xlsx_model::{
    AnchorCell, AnchorEditAs, AnchorExtent, AnchorPos, ChartAnchor, ChartRef, ChartRefKind,
    SheetChart,
};

use crate::tree::{Element, parse_tree, splice};
use crate::xml::{find_part, resolve_part_path};
use crate::{MAX_CHART_ANCHORS, MAX_CHART_REFS, MAX_DEPTH, ParseError};

/// relationship references inside a part (`r:id`).
const NS_RELATIONSHIPS: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
/// the `.rels` part vocabulary itself.
const NS_PACKAGE_RELATIONSHIPS: &str =
    "http://schemas.openxmlformats.org/package/2006/relationships";
/// classic `c:chartSpace`.
const NS_CHART: &str = "http://schemas.openxmlformats.org/drawingml/2006/chart";
/// `cx:chartSpace`, whose reference syntax and caches are their own vocabulary.
const NS_CHART_EX: &str = "http://schemas.microsoft.com/office/drawing/2014/chartex";
/// the worksheet drawing vocabulary that carries the anchors.
const NS_SPREADSHEET_DRAWING: &str =
    "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing";

/// relationship types followed out of a sheet or drawing part.
const TYPE_DRAWING: &str = "drawing";
const TYPE_CHART: &str = "chart";

/// Parse a chart part into the shared model, resolving `a:schemeClr` through
/// the workbook theme. `None` when the part carries no recognized plot.
pub fn chart_space(part: &[u8], theme: &Theme) -> Option<ChartSpace> {
    let root = parse_tree(part).ok()?;
    let drawing_theme = drawing_theme(theme);
    parse_chart_space(&ChartNode::new(&root, &drawing_theme))
}

/// A chart element with the workbook theme bound to it, so scheme colors
/// resolve against this workbook rather than the Office defaults.
struct ChartNode<'a> {
    element: &'a Element,
    theme: &'a DrawingTheme,
    children: Vec<ChartNode<'a>>,
}

impl<'a> ChartNode<'a> {
    fn new(element: &'a Element, theme: &'a DrawingTheme) -> Self {
        Self {
            element,
            theme,
            children: element
                .child_elements()
                .map(|child| ChartNode::new(child, theme))
                .collect(),
        }
    }
}

impl ChartXml for ChartNode<'_> {
    fn local_name(&self) -> &str {
        self.element.local_name()
    }

    fn attribute(&self, prefix: Option<&str>, name: &str) -> Option<&str> {
        self.element.attribute(prefix, name)
    }

    fn child_elements(&self) -> impl Iterator<Item = &Self> {
        self.children.iter()
    }

    fn descendant_text(&self) -> String {
        self.element.text_content()
    }

    fn solid_fill_hex(&self) -> Option<String> {
        let color = self
            .children
            .iter()
            .find_map(|child| match child.local_name() {
                "srgbClr" => Some(parse_color_value(
                    child.attribute(None, "val"),
                    None,
                    None,
                    None,
                )),
                "schemeClr" => Some(parse_color_value(
                    None,
                    child.attribute(None, "val").map(scheme_slot),
                    None,
                    None,
                )),
                _ => None,
            })?;
        resolve_color_value_to_hex_with_theme(Some(&color), Some(self.theme))
    }
}

/// `tx1`/`bg1` and friends name the slots the theme stores as `dk1`/`lt1`.
fn scheme_slot(slot: &str) -> &str {
    match slot {
        "tx1" => "dk1",
        "bg1" => "lt1",
        "tx2" => "dk2",
        "bg2" => "lt2",
        other => other,
    }
}

fn drawing_theme(theme: &Theme) -> DrawingTheme {
    const SLOTS: [&str; 12] = [
        "dk1", "lt1", "dk2", "lt2", "accent1", "accent2", "accent3", "accent4", "accent5",
        "accent6", "hlink", "folHlink",
    ];
    let mut drawing = DrawingTheme::default();
    for (slot, value) in SLOTS.iter().zip(&theme.colors) {
        drawing
            .color_scheme
            .set(slot, value.trim_start_matches('#').to_ascii_uppercase());
    }
    drawing
}

/// The charts a sheet anchors, followed from its drawing relationships. Both
/// worksheets and chartsheets carry drawings, so both are walked. A chart whose
/// part is missing is skipped rather than failing the parse.
pub(crate) fn parse_sheet_charts(
    parts: &[(String, Vec<u8>)],
    sheet_path: &str,
) -> Result<Vec<SheetChart>, ParseError> {
    let mut charts = Vec::new();
    for (drawing_path, drawing_xml) in sheet_drawings(parts, sheet_path)? {
        let drawing_rels = find_part(parts, &relationship_part_path(&drawing_path))
            .map(parse_relationships)
            .transpose()?
            .unwrap_or_default();
        let drawing_dir = directory_of(&drawing_path).to_owned();
        for (index, anchor) in read_anchors(&parse_tree(drawing_xml)?)?
            .into_iter()
            .enumerate()
        {
            let Some(part) = anchor
                .chart_rel
                .as_deref()
                .and_then(|id| relationship_target(&drawing_rels, id, TYPE_CHART))
                .map(|target| resolve_part_path(&drawing_dir, target))
            else {
                continue;
            };
            let Some(chart_xml) = find_part(parts, &part) else {
                continue;
            };
            let root = parse_tree(chart_xml)?;
            if !root.is(NS_CHART, "chartSpace") {
                continue;
            }
            if charts.len() >= MAX_CHART_ANCHORS {
                return Err(ParseError::TooManyCharts);
            }
            charts.push(SheetChart {
                part,
                drawing: drawing_path.clone(),
                anchor_index: index,
                anchor: anchor.anchor,
                refs: chart_refs(&root)?,
            });
        }
    }
    Ok(charts)
}

/// Every drawing part a sheet relates to, with its bytes.
fn sheet_drawings<'a>(
    parts: &'a [(String, Vec<u8>)],
    sheet_path: &str,
) -> Result<Vec<(String, &'a [u8])>, ParseError> {
    let Some(rels) = find_part(parts, &relationship_part_path(sheet_path)) else {
        return Ok(Vec::new());
    };
    let sheet_dir = directory_of(sheet_path);
    let mut drawings = Vec::new();
    for (_, _, target) in parse_relationships(rels)?
        .iter()
        .filter(|(_, kind, _)| type_is(kind, TYPE_DRAWING))
    {
        let path = resolve_part_path(sheet_dir, target);
        if let Some(bytes) = find_part(parts, &path) {
            drawings.push((path, bytes));
        }
    }
    Ok(drawings)
}

struct DrawingAnchor {
    anchor: ChartAnchor,
    chart_rel: Option<String>,
}

/// Every `xdr:` anchor in a drawing part, in document order, so an index into
/// this list keeps naming the same anchor across a save.
fn read_anchors(root: &Element) -> Result<Vec<DrawingAnchor>, ParseError> {
    if !root.is(NS_SPREADSHEET_DRAWING, "wsDr") {
        return Ok(Vec::new());
    }
    let mut anchors = Vec::new();
    for child in root.child_elements().filter(is_anchor) {
        let anchor = match child.local_name() {
            "twoCellAnchor" => ChartAnchor::TwoCell {
                from: anchor_cell(child.child("from")),
                to: anchor_cell(child.child("to")),
                edit_as: child
                    .attribute(None, "editAs")
                    .and_then(AnchorEditAs::from_sml)
                    .unwrap_or_default(),
            },
            "oneCellAnchor" => ChartAnchor::OneCell {
                from: anchor_cell(child.child("from")),
                extent: anchor_extent(child.child("ext")),
            },
            _ => ChartAnchor::Absolute {
                pos: anchor_pos(child.child("pos")),
                extent: anchor_extent(child.child("ext")),
            },
        };
        if anchors.len() >= MAX_CHART_ANCHORS {
            return Err(ParseError::TooManyCharts);
        }
        anchors.push(DrawingAnchor {
            anchor,
            chart_rel: chart_relationship_id(child, 0),
        });
    }
    Ok(anchors)
}

fn is_anchor(element: &&Element) -> bool {
    element.namespace() == Some(NS_SPREADSHEET_DRAWING)
        && matches!(
            element.local_name(),
            "twoCellAnchor" | "oneCellAnchor" | "absoluteAnchor"
        )
}

/// `c:chart/@r:id` under a graphic frame, searched depth-capped by expanded
/// name so an alternate prefix still resolves.
fn chart_relationship_id(element: &Element, depth: usize) -> Option<String> {
    if depth > MAX_DEPTH {
        return None;
    }
    if element.is(NS_CHART, "chart") {
        return element
            .attribute_ns(NS_RELATIONSHIPS, "id")
            .map(str::to_owned);
    }
    element
        .child_elements()
        .find_map(|child| chart_relationship_id(child, depth + 1))
}

fn anchor_cell(element: Option<&Element>) -> AnchorCell {
    let Some(element) = element else {
        return AnchorCell::default();
    };
    AnchorCell {
        col: child_index(element, "col", MAX_COLS),
        col_off: child_number(element, "colOff"),
        row: child_index(element, "row", MAX_ROWS),
        row_off: child_number(element, "rowOff"),
    }
}

fn anchor_extent(element: Option<&Element>) -> AnchorExtent {
    let Some(element) = element else {
        return AnchorExtent::default();
    };
    AnchorExtent {
        cx: attribute_number(element, "cx"),
        cy: attribute_number(element, "cy"),
    }
}

fn anchor_pos(element: Option<&Element>) -> AnchorPos {
    let Some(element) = element else {
        return AnchorPos::default();
    };
    AnchorPos {
        x: attribute_number(element, "x"),
        y: attribute_number(element, "y"),
    }
}

/// a grid index, clamped into the sheet so a hostile drawing can never reach
/// an out-of-range address.
fn child_index(element: &Element, local: &str, limit: u32) -> u32 {
    element
        .child(local)
        .map(Element::text_content)
        .and_then(|text| text.trim().parse::<u32>().ok())
        .unwrap_or(0)
        .min(limit.saturating_sub(1))
}

fn child_number(element: &Element, local: &str) -> i64 {
    element
        .child(local)
        .map(Element::text_content)
        .and_then(|text| text.trim().parse::<i64>().ok())
        .unwrap_or(0)
}

fn attribute_number(element: &Element, name: &str) -> i64 {
    element
        .attribute(None, name)
        .and_then(|value| value.trim().parse::<i64>().ok())
        .unwrap_or(0)
}

/// One `c:f` found in a chart part: what it references and where its content
/// sits, so the same walk serves both the model and the writer.
struct RefSite {
    kind: ChartRefKind,
    formula: String,
    span: Option<Range<usize>>,
}

fn chart_refs(root: &Element) -> Result<Vec<ChartRef>, ParseError> {
    Ok(ref_sites(root)?
        .into_iter()
        .map(|site| ChartRef {
            kind: site.kind,
            formula: site.formula,
        })
        .collect())
}

fn ref_sites(root: &Element) -> Result<Vec<RefSite>, ParseError> {
    let mut sites = Vec::new();
    walk_refs(root, ChartRefKind::Other, 0, &mut sites)?;
    Ok(sites)
}

fn walk_refs(
    element: &Element,
    inherited: ChartRefKind,
    depth: usize,
    out: &mut Vec<RefSite>,
) -> Result<(), ParseError> {
    if depth > MAX_DEPTH {
        return Err(ParseError::DepthExceeded);
    }
    let kind = slot_kind(element.local_name()).unwrap_or(inherited);
    if element.is(NS_CHART, "f") {
        if out.len() >= MAX_CHART_REFS {
            return Err(ParseError::TooManyCharts);
        }
        out.push(RefSite {
            kind,
            formula: element.text_content().trim().to_owned(),
            span: element.splice_target(),
        });
        return Ok(());
    }
    for child in element.child_elements() {
        walk_refs(child, kind, depth + 1, out)?;
    }
    Ok(())
}

fn slot_kind(local: &str) -> Option<ChartRefKind> {
    match local {
        "tx" => Some(ChartRefKind::SeriesName),
        "cat" | "xVal" => Some(ChartRefKind::Categories),
        "val" | "yVal" => Some(ChartRefKind::Values),
        "bubbleSize" => Some(ChartRefKind::BubbleSize),
        "title" => Some(ChartRefKind::Title),
        "dLbls" => Some(ChartRefKind::DataLabels),
        _ => None,
    }
}

/// Write `refs` back into their `c:f` elements, leaving every other byte
/// alone. Refuses when the part no longer holds the references the model was
/// built from, rather than writing them into the wrong slots.
pub(crate) fn patch_chart_refs(part: &[u8], refs: &[ChartRef]) -> Result<Vec<u8>, ParseError> {
    let sites = ref_sites(&parse_tree(part)?)?;
    if sites.len() != refs.len() {
        return Err(ParseError::UnsupportedEdit(format!(
            "chart part holds {} references but the model carries {}",
            sites.len(),
            refs.len()
        )));
    }
    let mut edits = Vec::new();
    for (site, reference) in sites.iter().zip(refs) {
        if site.formula == reference.formula {
            continue;
        }
        let Some(span) = site.span.clone() else {
            return Err(ParseError::UnsupportedEdit(
                "a self-closing c:f cannot take a rewritten reference".into(),
            ));
        };
        edits.push((span, reference.formula.clone()));
    }
    if edits.is_empty() {
        return Ok(part.to_vec());
    }
    splice(part, &edits)
}

/// Write moved anchors back into their drawing part. Only `col` and `row`
/// move; offsets, extents and every other byte stay as authored.
pub(crate) fn patch_drawing_anchors(
    part: &[u8],
    anchors: &[(usize, ChartAnchor)],
) -> Result<Vec<u8>, ParseError> {
    let root = parse_tree(part)?;
    let indexed = root.child_elements().filter(is_anchor).collect::<Vec<_>>();
    let mut edits: Vec<(Range<usize>, String)> = Vec::new();
    for (index, anchor) in anchors {
        let Some(element) = indexed.get(*index) else {
            return Err(ParseError::UnsupportedEdit(
                "a chart anchor no longer exists in its drawing part".into(),
            ));
        };
        match anchor {
            ChartAnchor::TwoCell { from, to, .. } => {
                push_cell_edits(element.child("from"), *from, &mut edits)?;
                push_cell_edits(element.child("to"), *to, &mut edits)?;
            }
            ChartAnchor::OneCell { from, .. } => {
                push_cell_edits(element.child("from"), *from, &mut edits)?;
            }
            ChartAnchor::Absolute { .. } => {}
        }
    }
    if edits.is_empty() {
        return Ok(part.to_vec());
    }
    edits.sort_by_key(|(span, _)| span.start);
    splice(part, &edits)
}

fn push_cell_edits(
    element: Option<&Element>,
    cell: AnchorCell,
    out: &mut Vec<(Range<usize>, String)>,
) -> Result<(), ParseError> {
    let Some(element) = element else {
        return Ok(());
    };
    for (local, value) in [("col", cell.col), ("row", cell.row)] {
        let Some(child) = element.child(local) else {
            continue;
        };
        let value = value.to_string();
        if child.text_content().trim() == value {
            continue;
        }
        let span = child.splice_target().ok_or_else(|| {
            ParseError::UnsupportedEdit("a self-closing anchor index cannot be moved".into())
        })?;
        out.push((span, value));
    }
    Ok(())
}

/// content types that carry workbook references in a chart vocabulary. the
/// style and colour-style parts under `xl/charts/` carry none.
const CHART_CONTENT_TYPES: [&str; 2] = [
    "application/vnd.openxmlformats-officedocument.drawingml.chart+xml",
    "application/vnd.ms-office.chartex+xml",
];

/// A chart part in the package that the model does not fully cover: one no
/// sheet claims, one that is not a classic `c:chartSpace`, or one carrying a
/// reference form the remapper cannot rewrite. Structural edits are refused
/// while such a part is present, because moving cells would strand it.
pub(crate) fn unmodelled_chart_part(
    parts: &[(String, Vec<u8>)],
    content_types: &[(String, String)],
    workbook: &xlsx_model::Workbook,
) -> Result<Option<String>, ParseError> {
    let modelled = workbook
        .sheets
        .iter()
        .flat_map(|sheet| sheet.charts.iter())
        .map(|chart| normalize_part_path(&chart.part))
        .collect::<std::collections::HashSet<_>>();
    for (path, bytes) in parts {
        if !is_chart_part(path, content_types) {
            continue;
        }
        if !modelled.contains(normalize_part_path(path)) {
            return Ok(Some(path.clone()));
        }
        if unsupported_reference_form(&parse_tree(bytes)?, 0) {
            return Ok(Some(path.clone()));
        }
    }
    Ok(None)
}

fn normalize_part_path(path: &str) -> &str {
    path.trim_start_matches('/')
}

/// Whether a part holds a chart, by its declared content type when the package
/// declares one and by the conventional layout otherwise.
fn is_chart_part(path: &str, content_types: &[(String, String)]) -> bool {
    let normalized = normalize_part_path(path).to_ascii_lowercase();
    if let Some((_, content_type)) = content_types
        .iter()
        .find(|(declared, _)| *declared == normalized)
    {
        return CHART_CONTENT_TYPES
            .iter()
            .any(|known| content_type.eq_ignore_ascii_case(known));
    }
    normalized.starts_with("xl/charts/chart")
        && normalized.ends_with(".xml")
        && !normalized.contains("/_rels/")
}

/// Whether a chart part carries a reference this crate cannot rewrite. Only a
/// classic `c:chartSpace` whose references are all `c:f` is covered; ChartEx,
/// pivot-sourced and externally-cached charts, and the `sqref` extension
/// references filtered series carry, are not.
fn unsupported_reference_form(element: &Element, depth: usize) -> bool {
    if depth == 0 && !element.is(NS_CHART, "chartSpace") {
        return true;
    }
    if depth > MAX_DEPTH {
        return true;
    }
    let unsupported = element.namespace() == Some(NS_CHART_EX)
        || element.local_name() == "sqref"
        || element.is(NS_CHART, "pivotSource")
        || element.is(NS_CHART, "externalData")
        || (element.local_name() == "f" && element.namespace() != Some(NS_CHART));
    unsupported
        || element
            .child_elements()
            .any(|child| unsupported_reference_form(child, depth + 1))
}

pub(crate) fn relationship_part_path(path: &str) -> String {
    match path.rsplit_once('/') {
        Some((directory, file)) => format!("{directory}/_rels/{file}.rels"),
        None => format!("_rels/{path}.rels"),
    }
}

fn directory_of(path: &str) -> &str {
    path.rsplit_once('/')
        .map(|(directory, _)| directory)
        .unwrap_or("")
}

fn type_is(kind: &str, suffix: &str) -> bool {
    kind.rsplit('/').next() == Some(suffix)
}

fn relationship_target<'a>(
    rels: &'a [(String, String, String)],
    id: &str,
    suffix: &str,
) -> Option<&'a str> {
    rels.iter()
        .find(|(rel_id, kind, _)| rel_id == id && type_is(kind, suffix))
        .map(|(_, _, target)| target.as_str())
}

/// `(Id, Type, Target)` for every internal relationship in a `.rels` part.
fn parse_relationships(data: &[u8]) -> Result<Vec<(String, String, String)>, ParseError> {
    let root = parse_tree(data)?;
    Ok(root
        .child_elements()
        .filter(|child| child.is(NS_PACKAGE_RELATIONSHIPS, "Relationship"))
        .filter(|child| {
            !child
                .attribute_local("TargetMode")
                .is_some_and(|mode| mode.eq_ignore_ascii_case("external"))
        })
        .filter_map(|child| {
            Some((
                child.attribute_local("Id")?.to_owned(),
                child.attribute_local("Type").unwrap_or_default().to_owned(),
                child.attribute_local("Target")?.to_owned(),
            ))
        })
        .collect())
}
