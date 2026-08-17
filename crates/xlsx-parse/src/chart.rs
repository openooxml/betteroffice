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
    AnchorCell, AnchorEditAs, AnchorExtent, AnchorPos, CellRef, CellValue, ChartAnchor, ChartRef,
    ChartRefKind, ErrorValue, SheetChart,
};

use crate::package::PartContentType;
use crate::tree::{
    Edit, Element, Part, Replacement, escape_text, names_are_resolvable, parse_tree,
};
use crate::xml::{find_part, resolve_part_path};
use crate::{MAX_CHART_ANCHORS, MAX_CHART_REFS, MAX_DEPTH, ParseError};

/// relationship references inside a part (`r:id`).
pub(crate) const NS_RELATIONSHIPS: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
/// the `.rels` part vocabulary itself.
pub(crate) const NS_PACKAGE_RELATIONSHIPS: &str =
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
    if has_3d_plot(&root, 0) {
        return None;
    }
    let drawing_theme = drawing_theme(theme);
    parse_chart_space(&ChartNode::new(&root, &drawing_theme))
}

/// Upper bound on the references one part may have projected, so a hostile
/// part cannot turn one frame into thousands of resolutions.
const MAX_PROJECTED_REFERENCES: usize = 256;

/// A chart part rendered against the current workbook. Every cache this crate
/// can resolve safely is rebuilt from live cells so an ordinary edit reaches
/// the chart; every other cache keeps the values the file was authored with.
/// This changes only what is drawn — a save still writes the source bytes back
/// untouched.
///
/// This parses the part twice per frame: once here to find the cache sites,
/// once in [`chart_space`] to read the spliced result. Memoizing it needs a key
/// over everything the output depends on: the part bytes, the theme, the owner
/// sheet name, the value of every cell the references name, and the workbook's
/// set of sheet names — a rename alone flips the output, because a reference
/// into a sheet the workbook no longer holds walks no cells at all. So a memo
/// has to be invalidated by cell writes and by sheet renames, not by a
/// workbook-wide revision counter, or a chart stops following the grid again.
pub fn preserved_chart_space(
    part: &[u8],
    workbook: &xlsx_model::Workbook,
    owner: &str,
    theme: &Theme,
) -> Option<ChartSpace> {
    match refreshed_chart_part(part, workbook, owner) {
        Some(refreshed) => chart_space(&refreshed, theme).or_else(|| chart_space(part, theme)),
        None => chart_space(part, theme),
    }
}

/// The part with its resolvable caches spliced up to date, or `None` when
/// nothing could be refreshed and the source bytes already say it best.
/// Anything that puts the reference vocabulary itself in doubt — ChartEx, a
/// pivot or external source, a filtered-series `sqref` — declines the whole
/// part rather than refreshing the references beside it.
fn refreshed_chart_part(
    part: &[u8],
    workbook: &xlsx_model::Workbook,
    owner: &str,
) -> Option<Vec<u8>> {
    if names_a_foreign_vocabulary(part) {
        return None;
    }
    let source = Part::decode(part).ok()?;
    let root = source.tree().ok()?;
    if !root.is(NS_CHART, "chartSpace") || unsupported_reference_form(&root, 0) {
        return None;
    }
    let sites = ref_sites(&root).ok()?;
    if sites.len() > MAX_PROJECTED_REFERENCES {
        return None;
    }
    let mut rebuild = CacheRebuild::displayed();
    let mut declined = std::collections::HashSet::new();
    let mut edits: Vec<(Option<usize>, Edit)> = Vec::new();
    for site in &sites {
        let Some(cache) = &site.cache else {
            continue;
        };
        let refreshed = cache
            .span
            .clone()
            .zip(regenerated_cache(cache, &site.formula, workbook, owner, &mut rebuild).ok());
        match refreshed {
            Some((span, markup)) => {
                edits.push((site.series, (span, Replacement::Markup(markup))));
            }
            None => {
                declined.extend(site.series);
            }
        }
    }
    // A renderer pairs a series' values with its categories slot by slot, so
    // refreshing one of them while the other keeps authored values would draw a
    // pairing no version of the file ever held. A series is projected whole or
    // left alone.
    let mut edits = edits
        .into_iter()
        .filter(|(series, _)| series.is_none_or(|series| !declined.contains(&series)))
        .map(|(_, edit)| edit)
        .collect::<Vec<_>>();
    if edits.is_empty() {
        return None;
    }
    edits.sort_by_key(|(span, _)| span.start);
    source.splice(&edits).ok()
}

/// Whether the raw bytes name a reference vocabulary this crate does not read,
/// checked before the part is decoded so a pivot or ChartEx chart pays nothing
/// per frame. Only ever causes a fall back to authored values, so a false match
/// costs freshness rather than correctness; [`unsupported_reference_form`] is
/// still the authority once a part is decoded.
fn names_a_foreign_vocabulary(part: &[u8]) -> bool {
    let Ok(text) = std::str::from_utf8(part) else {
        return false;
    };
    ["pivotSource", "externalData", "sqref", "chartex"]
        .iter()
        .any(|needle| text.contains(needle))
}

fn has_3d_plot(element: &Element, depth: usize) -> bool {
    element.local_name().ends_with("3DChart")
        || (depth < MAX_DEPTH
            && element
                .child_elements()
                .any(|child| has_3d_plot(child, depth + 1)))
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
/// worksheets and chartsheets carry drawings, so both are walked. A drawing or
/// chart part that is missing, or that this crate cannot read, is skipped
/// rather than failing the parse: the workbook opens without that chart, and
/// the part itself is still carried through a save untouched.
///
/// A drawing is skipped whole, never anchor by anchor, because `anchor_index`
/// names a position in the part's anchor list and a partial list would rename
/// the anchors a later save patches.
///
/// Every part skipped is appended to `declined`: this is the only place that
/// knows one was, and a part no save rewrites has to veto the structural edits
/// that would move what it names. A chart target is recorded as the drawing
/// relationship resolved it, which catches one outside the conventional layout
/// and one the package does not hold — neither of which a walk over the parts
/// can see.
pub(crate) fn parse_sheet_charts(
    parts: &[(String, Vec<u8>)],
    sheet_path: &str,
    declined: &mut Vec<String>,
) -> Result<Vec<SheetChart>, ParseError> {
    let mut charts = Vec::new();
    for (drawing_path, drawing_xml) in sheet_drawings(parts, sheet_path)? {
        let drawing_rels = find_part(parts, &relationship_part_path(&drawing_path))
            .map(parse_relationships)
            .transpose()?
            .unwrap_or_default();
        let drawing_dir = directory_of(&drawing_path).to_owned();
        let Ok(drawing_root) = parse_tree(drawing_xml) else {
            declined.push(drawing_path);
            continue;
        };
        let Ok(anchors) = read_anchors(&drawing_root) else {
            declined.push(drawing_path);
            continue;
        };
        for (index, anchor) in anchors.into_iter().enumerate() {
            let rel_id = match &anchor.chart {
                AnchorChart::None => continue,
                AnchorChart::Unrelated => {
                    declined.push(drawing_path.clone());
                    continue;
                }
                AnchorChart::Related(id) => id.as_str(),
            };
            let Some(part) = relationship_target(&drawing_rels, rel_id, TYPE_CHART)
                .map(|target| resolve_part_path(&drawing_dir, target))
            else {
                declined.push(drawing_path.clone());
                continue;
            };
            let Some(chart_xml) = find_part(parts, &part) else {
                declined.push(part);
                continue;
            };
            let Ok(root) = parse_tree(chart_xml) else {
                declined.push(part);
                continue;
            };
            if !root.is(NS_CHART, "chartSpace") {
                declined.push(part);
                continue;
            }
            let Ok(refs) = chart_refs(&root) else {
                declined.push(part);
                continue;
            };
            if charts.len() >= MAX_CHART_ANCHORS {
                return Err(ParseError::TooManyCharts);
            }
            charts.push(SheetChart {
                part,
                drawing: drawing_path.clone(),
                anchor_index: index,
                anchor: anchor.anchor,
                refs,
            });
        }
    }
    Ok(charts)
}

/// Every drawing part a sheet relates to, with its bytes. A part two
/// relationships both name is one drawing, and following it twice would emit
/// every anchor twice, so it is walked once.
fn sheet_drawings<'a>(
    parts: &'a [(String, Vec<u8>)],
    sheet_path: &str,
) -> Result<Vec<(String, &'a [u8])>, ParseError> {
    let Some(rels) = find_part(parts, &relationship_part_path(sheet_path)) else {
        return Ok(Vec::new());
    };
    let sheet_dir = directory_of(sheet_path);
    let mut drawings: Vec<(String, &[u8])> = Vec::new();
    for (_, _, target) in parse_relationships(rels)?
        .iter()
        .filter(|(_, kind, _)| type_is(kind, TYPE_DRAWING))
    {
        let path = resolve_part_path(sheet_dir, target);
        if drawings.iter().any(|(walked, _)| *walked == path) {
            continue;
        }
        if let Some(bytes) = find_part(parts, &path) {
            drawings.push((path, bytes));
        }
    }
    Ok(drawings)
}

struct DrawingAnchor {
    anchor: ChartAnchor,
    chart: AnchorChart,
}

/// What an anchor's graphic frame holds: no chart, a chart naming a
/// relationship, or a chart naming none — which is a frame this crate cannot
/// model and a save cannot move.
enum AnchorChart {
    None,
    Related(String),
    Unrelated,
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
                from: anchor_cell(child.child("from"))?,
                to: anchor_cell(child.child("to"))?,
                edit_as: match child.attribute(None, "editAs") {
                    Some(value) => AnchorEditAs::from_sml(value).ok_or_else(|| {
                        ParseError::Malformed(format!("invalid chart editAs value {value:?}"))
                    })?,
                    None => AnchorEditAs::default(),
                },
            },
            "oneCellAnchor" => ChartAnchor::OneCell {
                from: anchor_cell(child.child("from"))?,
                extent: anchor_extent(child.child("ext"))?,
            },
            _ => ChartAnchor::Absolute {
                pos: anchor_pos(child.child("pos"))?,
                extent: anchor_extent(child.child("ext"))?,
            },
        };
        if anchors.len() >= MAX_CHART_ANCHORS {
            return Err(ParseError::TooManyCharts);
        }
        anchors.push(DrawingAnchor {
            anchor,
            chart: anchor_chart(child, 0),
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

/// The `c:chart` under a graphic frame, searched depth-capped by expanded name
/// so an alternate prefix still resolves, and whether it names a relationship.
fn anchor_chart(element: &Element, depth: usize) -> AnchorChart {
    if depth > MAX_DEPTH {
        return AnchorChart::None;
    }
    if element.is(NS_CHART, "chart") {
        return match element.attribute_ns(NS_RELATIONSHIPS, "id") {
            Some(id) => AnchorChart::Related(id.to_owned()),
            None => AnchorChart::Unrelated,
        };
    }
    for child in element.child_elements() {
        match anchor_chart(child, depth + 1) {
            AnchorChart::None => {}
            found => return found,
        }
    }
    AnchorChart::None
}

fn anchor_cell(element: Option<&Element>) -> Result<AnchorCell, ParseError> {
    let element =
        element.ok_or_else(|| ParseError::Malformed("chart anchor has no cell".into()))?;
    Ok(AnchorCell {
        col: child_index(element, "col", MAX_COLS)?,
        col_off: child_number(element, "colOff")?,
        row: child_index(element, "row", MAX_ROWS)?,
        row_off: child_number(element, "rowOff")?,
    })
}

fn anchor_extent(element: Option<&Element>) -> Result<AnchorExtent, ParseError> {
    let element =
        element.ok_or_else(|| ParseError::Malformed("chart anchor has no extent".into()))?;
    Ok(AnchorExtent {
        cx: attribute_number(element, "cx")?,
        cy: attribute_number(element, "cy")?,
    })
}

fn anchor_pos(element: Option<&Element>) -> Result<AnchorPos, ParseError> {
    let element = element
        .ok_or_else(|| ParseError::Malformed("absolute chart anchor has no position".into()))?;
    Ok(AnchorPos {
        x: attribute_number(element, "x")?,
        y: attribute_number(element, "y")?,
    })
}

/// Reads an in-grid chart anchor index.
fn child_index(element: &Element, local: &str, limit: u32) -> Result<u32, ParseError> {
    let value = element
        .child(local)
        .map(Element::text_content)
        .and_then(|text| text.trim().parse::<u32>().ok())
        .filter(|value| *value < limit)
        .ok_or_else(|| ParseError::Malformed(format!("invalid chart anchor {local}")))?;
    Ok(value)
}

fn child_number(element: &Element, local: &str) -> Result<i64, ParseError> {
    let Some(child) = element.child(local) else {
        return Ok(0);
    };
    child
        .text_content()
        .trim()
        .parse::<i64>()
        .map_err(|_| ParseError::Malformed(format!("invalid chart anchor {local}")))
}

fn attribute_number(element: &Element, name: &str) -> Result<i64, ParseError> {
    element
        .attribute(None, name)
        .and_then(|value| value.trim().parse::<i64>().ok())
        .ok_or_else(|| ParseError::Malformed(format!("invalid chart anchor {name}")))
}

/// The cache sitting beside a `c:f`: the points a consumer reads when it does
/// not recalculate. A reference that moves must take its cache with it.
struct CacheSite {
    /// `numCache`, `strCache` or `multiLvlStrCache`.
    local: String,
    /// the authored prefix, so regenerated markup keeps the part's binding.
    prefix: String,
    span: Option<Range<usize>>,
    format_code: Option<String>,
    /// whether the cache carries content [`regenerated_cache`] does not
    /// re-emit, which it would therefore delete.
    unmodelled_content: bool,
    /// the points the file was authored with, so a rebuild that would empty a
    /// populated cache can be recognized and declined.
    authored_points: usize,
}

/// One `c:f` found in a chart part: what it references and where its content
/// sits, so the same walk serves both the model and the writer.
struct RefSite {
    kind: ChartRefKind,
    formula: String,
    span: Option<Range<usize>>,
    cache: Option<CacheSite>,
    /// the `c:ser` this reference sits in, by document order. a series is the
    /// unit a renderer pairs values and categories within, so it is the unit a
    /// projection has to accept or decline whole.
    series: Option<usize>,
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
    let mut series = Series {
        current: None,
        next: 0,
    };
    walk_refs(root, ChartRefKind::Other, &mut series, 0, &mut sites)?;
    Ok(sites)
}

/// Which `c:ser` a walk is inside, and the ordinal the next one takes.
struct Series {
    current: Option<usize>,
    next: usize,
}

fn walk_refs(
    element: &Element,
    inherited: ChartRefKind,
    series: &mut Series,
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
            cache: None,
            series: series.current,
        });
        return Ok(());
    }
    let enclosing = series.current;
    if element.is(NS_CHART, "ser") {
        series.current = Some(series.next);
        series.next += 1;
    }
    let before = out.len();
    for child in element.child_elements() {
        walk_refs(child, kind, series, depth + 1, out)?;
    }
    series.current = enclosing;
    if out.len() == before + 1
        && element
            .child(FORMULA_LOCAL)
            .is_some_and(|child| child.is(NS_CHART, FORMULA_LOCAL))
        && let Some(cache) = cache_site(element)
    {
        out[before].cache = Some(cache);
    }
    Ok(())
}

const FORMULA_LOCAL: &str = "f";

fn cache_site(reference: &Element) -> Option<CacheSite> {
    let cache = reference.child_elements().find(|child| {
        child.namespace() == Some(NS_CHART)
            && matches!(
                child.local_name(),
                "numCache" | "strCache" | "multiLvlStrCache"
            )
    })?;
    Some(CacheSite {
        local: cache.local_name().to_owned(),
        prefix: cache
            .name
            .rsplit_once(':')
            .map(|(prefix, _)| format!("{prefix}:"))
            .unwrap_or_default(),
        span: cache.splice_target(),
        format_code: cache
            .child("formatCode")
            .map(|element| element.text_content()),
        unmodelled_content: !cache_is_fully_modelled(cache),
        authored_points: cache
            .child_elements()
            .filter(|child| child.is(NS_CHART, "pt"))
            .count(),
    })
}

/// Whether every byte of a cache's content is something the regenerator can
/// write back. `extLst` and a per-point `formatCode` are legal cache content
/// this crate does not model, so a cache carrying either is left alone rather
/// than rebuilt without it.
fn cache_is_fully_modelled(cache: &Element) -> bool {
    cache.child_elements().all(|child| {
        child.namespace() == Some(NS_CHART)
            && match child.local_name() {
                "formatCode" => child.attributes.is_empty(),
                "ptCount" => child.attributes.iter().all(|value| value.name == "val"),
                "pt" => {
                    child.attributes.iter().all(|value| value.name == "idx")
                        && child
                            .child_elements()
                            .all(|value| value.is(NS_CHART, "v") && value.attributes.is_empty())
                }
                _ => false,
            }
    })
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

/// Upper bound on the points one regenerated cache may carry. Excel plots
/// 32,000 per series; a reference longer than this is refused rather than
/// turned into megabytes of markup.
pub(crate) const MAX_CACHE_POINTS: u32 = 65_536;

/// Write `refs` back into their `c:f` elements and regenerate the caches
/// beside them from `workbook`, leaving every other byte alone. Refuses when
/// the part no longer holds the references the model was built from, or when a
/// moved reference's cache cannot be rebuilt, rather than leaving a reference
/// pointing at one range while its cache still holds another's values.
pub(crate) fn patch_chart_refs(
    part: &[u8],
    refs: &[ChartRef],
    workbook: &xlsx_model::Workbook,
    owner: &str,
) -> Result<Vec<u8>, ParseError> {
    let source = Part::decode(part)?;
    let sites = ref_sites(&source.tree()?)?;
    if sites.len() != refs.len() {
        return Err(ParseError::UnsupportedEdit(format!(
            "chart part holds {} references but the model carries {}",
            sites.len(),
            refs.len()
        )));
    }
    let mut edits: Vec<Edit> = Vec::new();
    for (site, reference) in sites.iter().zip(refs) {
        if site.formula == reference.formula {
            continue;
        }
        let Some(span) = site.span.clone() else {
            return Err(ParseError::UnsupportedEdit(
                "a self-closing c:f cannot take a rewritten reference".into(),
            ));
        };
        edits.push((span, Replacement::Text(reference.formula.clone())));
        let Some(cache) = &site.cache else {
            continue;
        };
        let Some(cache_span) = cache.span.clone() else {
            return Err(ParseError::UnsupportedEdit(
                "a self-closing chart cache cannot be regenerated".into(),
            ));
        };
        edits.push((
            cache_span,
            Replacement::Markup(regenerated_cache(
                cache,
                &reference.formula,
                workbook,
                owner,
                &mut CacheRebuild::persisted(),
            )?),
        ));
    }
    if edits.is_empty() {
        return Ok(part.to_vec());
    }
    edits.sort_by_key(|(span, _)| span.start);
    source.splice(&edits)
}

/// What a rebuilt cache must be able to stand behind, and how many points the
/// caller still allows.
struct CacheRebuild {
    fidelity: CacheFidelity,
    /// the most cells any one reference may resolve to.
    limit: usize,
    /// what is left for the whole part, where the caller bounds it.
    budget: usize,
}

/// Whether a rebuilt cache is written back into a package or only read by a
/// renderer. A save must express exactly what the authored cache did; a render
/// only has to put a readable value on the screen.
#[derive(Clone, Copy, PartialEq, Eq)]
enum CacheFidelity {
    Persisted,
    Displayed,
}

impl CacheRebuild {
    /// A save: bounded per reference by what one cache may already carry, the
    /// way this crate has always written one back.
    fn persisted() -> Self {
        Self {
            fidelity: CacheFidelity::Persisted,
            limit: MAX_CACHE_POINTS as usize,
            budget: usize::MAX,
        }
    }

    /// A render: bounded across the whole part, because it runs every frame.
    fn displayed() -> Self {
        Self {
            fidelity: CacheFidelity::Displayed,
            limit: crate::MAX_PROJECTED_CACHE_POINTS,
            budget: crate::MAX_PROJECTED_CACHE_POINTS,
        }
    }
}

/// The content of a cache rebuilt from the current workbook. Refused when the
/// authored cache holds anything the rebuild would not write back.
fn regenerated_cache(
    cache: &CacheSite,
    formula: &str,
    workbook: &xlsx_model::Workbook,
    owner: &str,
    rebuild: &mut CacheRebuild,
) -> Result<String, ParseError> {
    if cache.local == "multiLvlStrCache" {
        return Err(ParseError::UnsupportedEdit(
            "a multi-level category cache cannot be regenerated".into(),
        ));
    }
    if cache.unmodelled_content {
        return Err(ParseError::UnsupportedEdit(
            "a chart cache carrying content this crate does not model cannot be regenerated".into(),
        ));
    }
    let numeric = cache.local == "numCache";
    let cells = resolve_reference(formula, workbook, owner, rebuild.limit.min(rebuild.budget))?;
    rebuild.budget -= cells.len();
    // A populated cache over cells the grid reads as empty — including a
    // reference that names none at all, such as `#REF!` — is the file speaking
    // for data the workbook does not hold; emptying the chart would be a worse
    // answer than a stale one. A save is regenerating because the reference
    // moved, so there the new emptiness is the truth.
    if rebuild.fidelity == CacheFidelity::Displayed
        && cache.authored_points > 0
        && cells.iter().all(|cell| *cell == CellValue::Empty)
    {
        return Err(ParseError::UnsupportedEdit(
            "a chart cache over cells the grid reads as empty keeps its authored values".into(),
        ));
    }
    let points = cells
        .iter()
        .map(|value| cache_point(value, numeric, rebuild.fidelity))
        .collect::<Result<Vec<_>, _>>()?;
    // A cache omits the points a consumer cannot read, and the shared parser
    // collects what remains positionally rather than by `@idx`. A hole before
    // the last surviving point would therefore slide every later value into an
    // earlier slot, under another point's category. A trailing hole shifts
    // nothing, so only an interior one declines.
    if rebuild.fidelity == CacheFidelity::Displayed
        && let Some(last) = points.iter().rposition(Option::is_some)
        && points[..last].iter().any(Option::is_none)
    {
        return Err(ParseError::UnsupportedEdit(
            "a chart cache with a gap between points keeps its authored values".into(),
        ));
    }
    let prefix = &cache.prefix;
    let mut out = String::new();
    if let Some(format_code) = &cache.format_code {
        out.push_str(&format!(
            "<{prefix}formatCode>{}</{prefix}formatCode>",
            escape_text(format_code)?
        ));
    }
    out.push_str(&format!("<{prefix}ptCount val=\"{}\"/>", cells.len()));
    for (index, text) in points.iter().enumerate() {
        let Some(text) = text else {
            continue;
        };
        out.push_str(&format!(
            "<{prefix}pt idx=\"{index}\"><{prefix}v>{}</{prefix}v></{prefix}pt>",
            escape_text(text)?
        ));
    }
    Ok(out)
}

/// One cached point, or `None` for a cell a cache omits. A save cannot put a
/// value that is not text into a string cache without applying the cell's
/// number format, so it refuses instead of guessing; a render only has to show
/// a readable value.
fn cache_point(
    value: &CellValue,
    numeric: bool,
    fidelity: CacheFidelity,
) -> Result<Option<String>, ParseError> {
    Ok(match (value, numeric) {
        (CellValue::Empty, _) => None,
        (CellValue::Number { value }, true) => Some(format_number(*value)),
        (_, true) => None,
        (CellValue::Text { value }, false) => Some(value.clone()),
        (_, false) if fidelity == CacheFidelity::Persisted => {
            return Err(ParseError::UnsupportedEdit(
                "a string cache over non-text cells cannot be regenerated".into(),
            ));
        }
        (CellValue::Number { value }, false) => Some(format_number(*value)),
        (CellValue::Bool { value }, false) => {
            Some(if *value { "TRUE" } else { "FALSE" }.to_owned())
        }
        (CellValue::Error { value }, false) => Some(value.as_str().to_owned()),
    })
}

fn format_number(value: f64) -> String {
    if value.is_finite() {
        format!("{value}")
    } else {
        "0".to_owned()
    }
}

/// The cells one chart reference names, in the order a cache lists them, up to
/// `limit` of them. Only a single contiguous one-dimensional area on a sheet
/// this workbook holds can be resolved; `#REF!` resolves to nothing, which is
/// the correct empty cache. Every other form — a union, an external book, a
/// defined name, a two-dimensional area — is refused here, so a caller needs no
/// predicate of its own.
fn resolve_reference(
    formula: &str,
    workbook: &xlsx_model::Workbook,
    owner: &str,
    limit: usize,
) -> Result<Vec<CellValue>, ParseError> {
    let trimmed = formula.trim();
    if trimmed.is_empty() || trimmed == ErrorValue::Ref.as_str() {
        return Ok(Vec::new());
    }
    let refused =
        |reason: &str| ParseError::UnsupportedEdit(format!("chart reference {trimmed} {reason}"));
    let (sheet_name, area) =
        split_qualifier(trimmed).ok_or_else(|| refused("is not a single same-workbook area"))?;
    let sheet_name = sheet_name.unwrap_or_else(|| owner.to_owned());
    let Some(sheet) = workbook
        .sheets
        .iter()
        .find(|sheet| sheet.name.eq_ignore_ascii_case(&sheet_name))
    else {
        return Err(refused("names a sheet this workbook does not hold"));
    };
    let (start, end) = parse_area(area).ok_or_else(|| refused("is not a cell range"))?;
    let (rows, cols) = (end.row - start.row + 1, end.col - start.col + 1);
    if rows > 1 && cols > 1 {
        return Err(refused("spans more than one row and column"));
    }
    let count = u64::from(rows) * u64::from(cols);
    if count > limit as u64 {
        return Err(refused("would hold more points than a chart can carry"));
    }
    let mut values = Vec::with_capacity(count as usize);
    for row in start.row..=end.row {
        for col in start.col..=end.col {
            values.push(
                sheet
                    .cell(CellRef::new(row, col))
                    .map(|cell| cell.value.clone())
                    .unwrap_or_default(),
            );
        }
    }
    Ok(values)
}

/// Splits `Sheet!$A$1:$A$5` into its sheet name and its area. `None` when the
/// reference carries anything else: a union, an external book, a defined name.
fn split_qualifier(source: &str) -> Option<(Option<String>, &str)> {
    if source.contains(',') || source.contains('(') || source.contains('[') {
        return None;
    }
    let Some(rest) = source.strip_prefix('\'') else {
        return match source.split_once('!') {
            Some((name, area)) if !name.is_empty() && !name.contains(':') => {
                Some((Some(name.to_owned()), area))
            }
            Some(_) => None,
            None => Some((None, source)),
        };
    };
    let mut name = String::new();
    let mut cursor = 0;
    let bytes = rest.as_bytes();
    while cursor < bytes.len() {
        if bytes[cursor] == b'\'' {
            if bytes.get(cursor + 1) == Some(&b'\'') {
                name.push('\'');
                cursor += 2;
                continue;
            }
            let area = rest.get(cursor + 1..)?.strip_prefix('!')?;
            return (!name.contains(':')).then_some((Some(name), area));
        }
        let character = rest[cursor..].chars().next()?;
        name.push(character);
        cursor += character.len_utf8();
    }
    None
}

/// `$A$1:$A$5` or `$A$1` as an inclusive corner pair.
pub(crate) fn parse_area(source: &str) -> Option<(CellRef, CellRef)> {
    let (start, end) = match source.split_once(':') {
        Some((start, end)) => (start, end),
        None => (source, source),
    };
    let start = CellRef::parse_a1(start).ok()?;
    let end = CellRef::parse_a1(end).ok()?;
    Some((
        CellRef::new(start.row.min(end.row), start.col.min(end.col)),
        CellRef::new(start.row.max(end.row), start.col.max(end.col)),
    ))
}

/// Write moved anchors back into their drawing part. A grid-anchored marker is
/// written whole — `col`, `colOff`, `row` and `rowOff` — so a two-cell anchor
/// may be re-placed and resized by moving its corners. An anchor whose kind,
/// `editAs` mode, one-cell extent or absolute position also changed is refused,
/// because none of that is written back.
pub(crate) fn patch_drawing_anchors(
    part: &[u8],
    anchors: &[(usize, ChartAnchor)],
) -> Result<Vec<u8>, ParseError> {
    let source = Part::decode(part)?;
    let root = source.tree()?;
    let indexed = anchor_elements(&root);
    let authored = read_anchors(&root)?;
    let mut edits: Vec<Edit> = Vec::new();
    for (index, anchor) in anchors {
        let (Some(element), Some(authored)) = (indexed.get(*index), authored.get(*index)) else {
            return Err(ParseError::UnsupportedEdit(
                "a chart anchor no longer exists in its drawing part".into(),
            ));
        };
        if !only_grid_position_moved(&authored.anchor, anchor) {
            return Err(ParseError::UnsupportedEdit(
                "a chart anchor changed more than the grid position a save writes back".into(),
            ));
        }
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
    source.splice(&edits)
}

/// The anchor elements [`read_anchors`] modelled, in the same order, so an
/// index selects the same anchor in both.
fn anchor_elements(root: &Element) -> Vec<&Element> {
    if !root.is(NS_SPREADSHEET_DRAWING, "wsDr") {
        return Vec::new();
    }
    root.child_elements().filter(is_anchor).collect()
}

/// Whether the model's anchor differs from the authored one only in what
/// [`push_cell_edits`] writes: the markers, cell and offset alike. The corners
/// themselves are unconstrained, so a two-cell anchor may also have been
/// resized. What is refused is what no marker carries — a different anchor
/// kind, a different `editAs` mode, a one-cell extent (which lives in
/// `xdr:ext`) and an absolute position (which lives in attributes).
fn only_grid_position_moved(authored: &ChartAnchor, moved: &ChartAnchor) -> bool {
    match (authored, moved) {
        (
            ChartAnchor::TwoCell { edit_as, .. },
            ChartAnchor::TwoCell {
                edit_as: moved_edit_as,
                ..
            },
        ) => edit_as == moved_edit_as,
        (
            ChartAnchor::OneCell { extent, .. },
            ChartAnchor::OneCell {
                extent: moved_extent,
                ..
            },
        ) => extent == moved_extent,
        (ChartAnchor::Absolute { .. }, ChartAnchor::Absolute { .. }) => authored == moved,
        _ => false,
    }
}

/// Writes a moved marker back. A child the drawing already carries is patched
/// in place, so the rest of the part survives byte for byte. A drawing that
/// omitted `colOff`/`rowOff` — which [`child_number`] reads as zero — has no
/// span to splice a non-zero offset into, so the whole marker is regenerated
/// rather than dropping the offset silently.
fn push_cell_edits(
    element: Option<&Element>,
    cell: AnchorCell,
    out: &mut Vec<Edit>,
) -> Result<(), ParseError> {
    let Some(element) = element else {
        return Ok(());
    };
    let mut patches = Vec::new();
    let mut moved = false;
    let mut absent = false;
    for (local, value) in marker_values(cell) {
        let Some(child) = element.child(local) else {
            moved |= value != 0;
            absent |= value != 0;
            continue;
        };
        let authored = child
            .text_content()
            .trim()
            .parse::<i64>()
            .map_err(|_| ParseError::Malformed(format!("invalid chart anchor {local}")))?;
        if authored == value {
            continue;
        }
        moved = true;
        let span = child.splice_target().ok_or_else(|| {
            ParseError::UnsupportedEdit("a self-closing anchor marker cannot be moved".into())
        })?;
        patches.push((span, Replacement::Text(value.to_string())));
    }
    if !moved {
        return Ok(());
    }
    if absent {
        let span = element.splice_target().ok_or_else(|| {
            ParseError::UnsupportedEdit("a self-closing chart anchor cannot be moved".into())
        })?;
        out.push((span, Replacement::Markup(marker_markup(element, cell))));
        return Ok(());
    }
    out.append(&mut patches);
    Ok(())
}

/// The four elements `CT_Marker` sequences, in schema order.
fn marker_values(cell: AnchorCell) -> [(&'static str, i64); 4] {
    [
        ("col", i64::from(cell.col)),
        ("colOff", cell.col_off),
        ("row", i64::from(cell.row)),
        ("rowOff", cell.row_off),
    ]
}

/// A whole marker body in schema order, carrying the prefix the drawing binds
/// the marker itself with.
fn marker_markup(element: &Element, cell: AnchorCell) -> String {
    let prefix = element
        .name
        .rsplit_once(':')
        .map(|(prefix, _)| format!("{prefix}:"))
        .unwrap_or_default();
    marker_values(cell)
        .iter()
        .map(|(local, value)| format!("<{prefix}{local}>{value}</{prefix}{local}>"))
        .collect()
}

/// content types that carry workbook references in a chart vocabulary. the
/// style and colour-style parts under `xl/charts/` carry none.
pub(crate) const CHART_CONTENT_TYPES: [&str; 2] = [
    "application/vnd.openxmlformats-officedocument.drawingml.chart+xml",
    "application/vnd.ms-office.chartex+xml",
];

/// A chart part in the package that the model does not fully cover: one no
/// sheet claims, one that is not a classic `c:chartSpace`, one carrying a
/// reference form the remapper cannot rewrite, or one whose cache sits beside
/// a reference this crate cannot rebuild it from. Structural edits that would
/// move what such a part names are refused, because nothing rewrites it.
///
/// Each part comes with the sheet an unqualified reference in it resolves
/// against: the one sheet claiming it, absent when no sheet or several do.
pub(crate) fn unmodelled_chart_parts(
    parts: &[(String, Vec<u8>)],
    content_types: &[PartContentType],
    workbook: &xlsx_model::Workbook,
) -> Result<Vec<UnmodelledChart>, ParseError> {
    let mut claims: std::collections::HashMap<&str, Vec<&str>> = std::collections::HashMap::new();
    for sheet in &workbook.sheets {
        for chart in &sheet.charts {
            claims
                .entry(normalize_part_path(&chart.part))
                .or_default()
                .push(sheet.name.as_str());
        }
    }
    let mut unmodelled = Vec::new();
    for (path, bytes) in parts {
        if !is_chart_part(path, content_types) {
            continue;
        }
        let owners = claims.get(normalize_part_path(path));
        let claimed = owners.is_some();
        if claimed {
            let root = parse_tree(bytes)?;
            if names_are_resolvable(&root)
                && !unsupported_reference_form(&root, 0)
                && !holds_an_unrebuildable_cache(&root)?
            {
                continue;
            }
        }
        unmodelled.push(UnmodelledChart {
            path: path.clone(),
            owner: owners
                .filter(|owners| owners.len() == 1)
                .map(|owners| owners[0].to_owned()),
            claimed,
        });
    }
    Ok(unmodelled)
}

/// A chart part no save rewrites, with the sheet its unqualified references
/// resolve against.
pub(crate) struct UnmodelledChart {
    pub(crate) path: String,
    pub(crate) owner: Option<String>,
    /// whether a sheet anchors it. A part nothing anchors is one this crate
    /// has not read the shape of, so what it names is not resolved from it.
    pub(crate) claimed: bool,
}

/// The areas a chart part names, or `None` when one of its references is not a
/// form this crate reads and every cell has to be assumed named. `owner` is the
/// sheet an unqualified reference resolves against.
pub(crate) fn chart_reference_areas(
    part: &[u8],
    owner: Option<&str>,
) -> Result<Option<Vec<(String, CellRef)>>, ParseError> {
    let root = parse_tree(part)?;
    if !names_are_resolvable(&root) || unsupported_reference_form(&root, 0) {
        return Ok(None);
    }
    let mut areas = Vec::new();
    for site in ref_sites(&root)? {
        let formula = site.formula.trim();
        if formula.is_empty() || formula == ErrorValue::Ref.as_str() {
            continue;
        }
        let Some((qualifier, area)) = split_qualifier(formula) else {
            return Ok(None);
        };
        let (Some(sheet), Some((_, end))) = (
            qualifier.or_else(|| owner.map(str::to_owned)),
            parse_area(area),
        ) else {
            return Ok(None);
        };
        areas.push((sheet, end));
    }
    Ok(Some(areas))
}

/// Whether a cache sits beside a reference this crate could not rebuild it
/// from. Such a reference survives a structural edit unchanged — a defined
/// name resolves somewhere else afterwards, and the remapper never touches the
/// literal formula — so the save would keep a cache holding pre-edit values.
fn holds_an_unrebuildable_cache(root: &Element) -> Result<bool, ParseError> {
    Ok(ref_sites(root)?.iter().any(|site| {
        site.cache.as_ref().is_some_and(|cache| {
            cache.local == "multiLvlStrCache" || !is_direct_one_dimensional_range(&site.formula)
        })
    }))
}

/// Whether a reference is the one form [`resolve_reference`] reads: a single
/// contiguous one-dimensional area, optionally sheet-qualified. An empty
/// reference and `#REF!` name nothing, which is the empty cache.
fn is_direct_one_dimensional_range(formula: &str) -> bool {
    let trimmed = formula.trim();
    if trimmed.is_empty() || trimmed == ErrorValue::Ref.as_str() {
        return true;
    }
    let Some((_, area)) = split_qualifier(trimmed) else {
        return false;
    };
    let Some((start, end)) = parse_area(area) else {
        return false;
    };
    let (rows, cols) = (end.row - start.row + 1, end.col - start.col + 1);
    (rows == 1 || cols == 1) && u64::from(rows) * u64::from(cols) <= u64::from(MAX_CACHE_POINTS)
}

fn normalize_part_path(path: &str) -> &str {
    path.trim_start_matches('/')
}

/// Whether a part holds a chart. Monotonic over the blanket veto the way
/// pivot discovery is: the conventional layout always answers, and conforming
/// typing may only add parts beside it. Typing that cannot be read must never
/// take a chart away, because a chart nobody looks at vetoes nothing.
fn is_chart_part(path: &str, content_types: &[PartContentType]) -> bool {
    let normalized = normalize_part_path(path).to_ascii_lowercase();
    (normalized.starts_with("xl/charts/chart")
        && normalized.ends_with(".xml")
        && !normalized.contains("/_rels/"))
        || content_types.iter().any(|part| {
            part.path == normalized
                && CHART_CONTENT_TYPES
                    .iter()
                    .any(|known| part.content_type.eq_ignore_ascii_case(known))
        })
}

/// Whether every anchor in a drawing claims at most one chart. The claim is
/// followed by taking the first `c:chart` descendant of an anchor, so a second
/// one lets two anchors swap which chart each is taken to hold — both still
/// claimed once, so nothing looks unknown, while an unqualified reference
/// resolves against the wrong sheet.
pub(crate) fn drawing_claims_are_unambiguous(root: &Element) -> bool {
    if !root.is(NS_SPREADSHEET_DRAWING, "wsDr") {
        return true;
    }
    root.child_elements()
        .filter(is_anchor)
        .all(|anchor| chart_claims(anchor, 0) <= 1)
}

/// How many claims an anchor could be read as carrying. A `c:chart` is one; so
/// is a second relationships-namespace `id` on it, since the claim is followed
/// by taking the first such attribute and either could answer.
fn chart_claims(element: &Element, depth: usize) -> usize {
    if depth > MAX_DEPTH {
        return usize::MAX;
    }
    let here = if element.is(NS_CHART, "chart") {
        element
            .attributes
            .iter()
            .filter(|attribute| {
                attribute.local_name() == "id"
                    && attribute.namespace.as_deref() == Some(NS_RELATIONSHIPS)
            })
            .count()
            .max(1)
    } else {
        0
    };
    here + element
        .child_elements()
        .map(|child| chart_claims(child, depth + 1))
        .sum::<usize>()
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

pub(crate) fn directory_of(path: &str) -> &str {
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

/// `(Id, Type, Target)` for every internal relationship in a `.rels` part. The
/// `.rels` vocabulary has one element name, so an undeclared namespace is read
/// rather than treated as a different part shape.
pub(crate) fn parse_relationships(
    data: &[u8],
) -> Result<Vec<(String, String, String)>, ParseError> {
    let root = parse_tree(data)?;
    Ok(root
        .child_elements()
        .filter(|child| {
            child.local_name() == "Relationship"
                && matches!(child.namespace(), None | Some(NS_PACKAGE_RELATIONSHIPS))
        })
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
