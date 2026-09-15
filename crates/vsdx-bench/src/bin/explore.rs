use std::collections::{BTreeMap, BTreeSet};
use std::fs;
use std::path::PathBuf;

use serde::Serialize;
use vsdx_eval::{
    DocumentReferences, Evaluation, Expr, PageShapeReferences, evaluate_cell_with_package_theme,
    evaluate_cell_with_shape_package_theme, parse as eval_parse,
};
use vsdx_parse::{
    Cell, ParseLimits, Row, Section, Shape, ShapeChild, ShapesChild, Sheet, VsdxError, parse_vsdx,
    write_vsdx,
};
use vsdx_render::{Primitive, RenderLimits, Renderer};
use vsdx_resolve::{Lookup, ResolvedShape, Resolver};

const FONT_BYTES: &[u8] =
    include_bytes!("../../../ooxml-text/tests/fonts/LiberationSans-Regular.ttf");
const KNOWN_CELLS: [&str; 4] = ["PinX", "PinY", "Width", "Height"];
const VISIBILITY_CONTROLS: [&str; 3] = ["NoFill", "NoLine", "NoShow"];
const HISTOGRAM_CAP: usize = 20;
/// Standard Visio ShapeSheet function names; anything else buckets as unknown.
const KNOWN_FUNCTIONS: [&str; 213] = [
    "ABS",
    "ACOS",
    "AND",
    "ANG360",
    "ANGLEALONGPATH",
    "ANGLETOLOC",
    "ANGLETOPAR",
    "ARG",
    "ASIN",
    "ATAN",
    "ATAN2",
    "BITAND",
    "BITNOT",
    "BITOR",
    "BITXOR",
    "BKGPAGENAME",
    "BLEND",
    "BLOB",
    "BLUE",
    "BOUND",
    "BOUNDINGBOXDIST",
    "BOUNDINGBOXRECT",
    "CALLOUTCOUNT",
    "CALLOUTTARGETREF",
    "CALLTHIS",
    "CATEGORY",
    "CEILING",
    "CHAR",
    "COMPANY",
    "CONTAINERCOUNT",
    "CONTAINERMEMBERCOUNT",
    "CONTAINERSHEETREF",
    "COS",
    "COSH",
    "CREATOR",
    "CY",
    "DATA1",
    "DATA2",
    "DATA3",
    "DATE",
    "DATETIME",
    "DATEVALUE",
    "DAY",
    "DAYOFYEAR",
    "DECIMALSEP",
    "DEFAULTEVENT",
    "DEG",
    "DEPENDSON",
    "DESCRIPTION",
    "DIRECTORY",
    "DISTTOPATH",
    "DOCCREATION",
    "DOCLASTEDIT",
    "DOCLASTPRINT",
    "DOCLASTSAVE",
    "DOCMD",
    "DOOLEVERB",
    "EVALCELL",
    "EVALTEXT",
    "FIELDPICTURE",
    "FILENAME",
    "FIND",
    "FLOOR",
    "FONT",
    "FONTTOID",
    "FORMAT",
    "FORMATEX",
    "FORMULAEXISTS",
    "GETREF",
    "GETVAL",
    "GOTOPAGE",
    "GRAVITY",
    "GREEN",
    "GUARD",
    "HASCATEGORY",
    "HELP",
    "HOUR",
    "HSL",
    "HUE",
    "HUEDIFF",
    "HYPERLINK",
    "HYPERLINKBASE",
    "ID",
    "IF",
    "IFERROR",
    "INDEX",
    "INT",
    "INTERSECTX",
    "INTERSECTY",
    "INTUP",
    "IS1D",
    "ISERR",
    "ISERRNA",
    "ISERROR",
    "ISERRVALUE",
    "ISTHEMED",
    "KEYWORDS",
    "LANGUAGE",
    "LEFT",
    "LEN",
    "LISTMEMBERCOUNT",
    "LISTORDER",
    "LISTSEP",
    "LISTSHEETREF",
    "LN",
    "LOC",
    "LOCTOLOC",
    "LOCTOPAR",
    "LOCALFORMULAEXISTS",
    "LOG10",
    "LOOKUP",
    "LOWER",
    "LUM",
    "LUMDIFF",
    "MAGNITUDE",
    "MANAGER",
    "MASTERNAME",
    "MAX",
    "MID",
    "MIN",
    "MINUTE",
    "MOD",
    "MODULUS",
    "MONTH",
    "MSOSHADE",
    "MSOTINT",
    "NA",
    "NAME",
    "NEARESTPOINTONPATH",
    "NOT",
    "NOW",
    "NURBS",
    "OPENFILE",
    "OPENGROUPWIN",
    "OPENSHEETWIN",
    "OPENTEXTWIN",
    "OR",
    "PAGECOUNT",
    "PAGENAME",
    "PAGENUMBER",
    "PAR",
    "PATHLENGTH",
    "PATHSEGMENT",
    "PI",
    "PLAYSOUND",
    "PNT",
    "PNTX",
    "PNTY",
    "POINTALONGPATH",
    "POLYLINE",
    "POW",
    "QUEUEMARKEREVENT",
    "RAD",
    "RAND",
    "RECTSECT",
    "RED",
    "REF",
    "REPLACE",
    "REPT",
    "REWIDEN",
    "RGB",
    "RIGHT",
    "ROUND",
    "RUNADDON",
    "RUNADDONWARGS",
    "RUNMACRO",
    "SAT",
    "SATDIFF",
    "SECOND",
    "SEGMENTCOUNT",
    "SETATREF",
    "SETATREFEVAL",
    "SETATREFEXPR",
    "SETF",
    "SHADE",
    "SHAPETEXT",
    "SHEETREF",
    "SIGN",
    "SIN",
    "SINH",
    "SQRT",
    "STRSAME",
    "STRSAMEEX",
    "SUBJECT",
    "SUBSTITUTE",
    "SUM",
    "TAN",
    "TANH",
    "TEXTHEIGHT",
    "TEXTWIDTH",
    "THEME",
    "THEMECBV",
    "THEMEGUARD",
    "THEMERESTORE",
    "THEMEVAL",
    "TIME",
    "TIMEVALUE",
    "TINT",
    "TITLE",
    "TONE",
    "TRIM",
    "TRUNC",
    "TYPE",
    "TYPEDESC",
    "UNICHAR",
    "UPPER",
    "USE",
    "USERUI",
    "VERSION",
    "WEEKDAY",
    "YEAR",
    "NO FORMULA",
    "_XFTRIGGER",
];
/// Standard Visio Geometry row types; anything else buckets as unknown.
const KNOWN_GEOMETRY_ROWS: [&str; 17] = [
    "ArcTo",
    "CubBezTo",
    "Ellipse",
    "EllipticalArcTo",
    "InfiniteLine",
    "LineTo",
    "MoveTo",
    "NURBSTo",
    "PolylineTo",
    "QuadBezTo",
    "RelCubBezTo",
    "RelEllipticalArcTo",
    "RelLineTo",
    "RelMoveTo",
    "RelQuadBezTo",
    "SplineKnot",
    "SplineStart",
];

#[derive(Default, Serialize)]
#[serde(rename_all = "camelCase")]
struct FileSurvey {
    name: String,
    parse_ok: bool,
    parse_error: Option<String>,
    part_count: usize,
    page_count: usize,
    roundtrip_identical: Option<bool>,
    shape_count: usize,
    master_shape_count: usize,
    resolve_absent: BTreeMap<String, usize>,
    evaluated: usize,
    unsupported_known: usize,
    unsupported_other: usize,
    error: usize,
    total: usize,
    unsupported_constructs: BTreeMap<String, usize>,
    unsupported_other_kinds: BTreeMap<String, usize>,
    error_kinds: BTreeMap<String, usize>,
    painted_only_shapes: usize,
    placeholder_shapes: usize,
    hidden: usize,
    unrendered: usize,
    primitives_emitted: usize,
    primitives_painted: usize,
    primitives_placeholdered: usize,
    placeholder_reasons: BTreeMap<String, usize>,
    render_page_errors: usize,
    geometry_rows: BTreeMap<String, usize>,
    master_geometry_rows: BTreeMap<String, usize>,
    visibility_carriers: BTreeMap<String, usize>,
    master_visibility_carriers: BTreeMap<String, usize>,
    visibility_placeholders: BTreeMap<String, usize>,
}

#[derive(Default, Serialize)]
#[serde(rename_all = "camelCase")]
struct Aggregate {
    files: usize,
    parse_ok: usize,
    parse_failed: usize,
    roundtrip_identical: usize,
    roundtrip_differ: usize,
    roundtrip_not_attempted: usize,
    pages: usize,
    shapes: usize,
    master_shapes: usize,
    resolve_absent: BTreeMap<String, usize>,
    evaluated: usize,
    unsupported_known: usize,
    unsupported_other: usize,
    error: usize,
    total: usize,
    unsupported_constructs: BTreeMap<String, usize>,
    unsupported_other_kinds: BTreeMap<String, usize>,
    error_kinds: BTreeMap<String, usize>,
    painted_only_shapes: usize,
    placeholder_shapes: usize,
    hidden: usize,
    unrendered: usize,
    primitives_emitted: usize,
    primitives_painted: usize,
    primitives_placeholdered: usize,
    placeholder_reasons: BTreeMap<String, usize>,
    render_page_errors: usize,
    geometry_rows: BTreeMap<String, usize>,
    master_geometry_rows: BTreeMap<String, usize>,
    visibility_carriers: BTreeMap<String, usize>,
    master_visibility_carriers: BTreeMap<String, usize>,
    visibility_placeholders: BTreeMap<String, usize>,
}

#[derive(Serialize)]
#[serde(rename_all = "camelCase")]
struct ExploreOutput {
    files: Vec<FileSurvey>,
    aggregate: Aggregate,
}

fn main() -> Result<(), Box<dyn std::error::Error>> {
    let args = std::env::args().skip(1).collect::<Vec<_>>();
    if args.iter().any(|arg| arg != "--json-only") {
        return Err("usage: explore [--json-only]".into());
    }
    let json_only = !args.is_empty();
    let Some(directory) = std::env::var_os("VSDX_EXPLORE_DIR") else {
        eprintln!("VSDX_EXPLORE_DIR is unset; nothing to survey");
        return Ok(());
    };
    let directory = PathBuf::from(directory);
    let entries = fs::read_dir(&directory).map_err(|error| {
        format!(
            "cannot read VSDX_EXPLORE_DIR {}: {error}",
            directory.display()
        )
    })?;
    let mut paths = Vec::new();
    for entry in entries {
        let Ok(entry) = entry else { continue };
        let path = entry.path();
        let is_file = entry.file_type().is_ok_and(|kind| kind.is_file());
        if !is_file {
            continue;
        }
        let wanted = path
            .extension()
            .and_then(|extension| extension.to_str())
            .is_some_and(|extension| {
                extension.eq_ignore_ascii_case("vsdx") || extension.eq_ignore_ascii_case("vstx")
            });
        if wanted {
            paths.push(path);
        }
    }
    paths.sort();
    let mut files = Vec::with_capacity(paths.len());
    for path in &paths {
        let name = path
            .file_name()
            .map(|name| name.to_string_lossy().into_owned())
            .unwrap_or_else(|| path.display().to_string());
        match fs::read(path) {
            Ok(bytes) => files.push(survey(&bytes, &name)),
            Err(error) => files.push(FileSurvey {
                name,
                parse_error: Some(error.kind().to_string()),
                ..Default::default()
            }),
        }
    }
    for file in &mut files {
        cap_file_histograms(file);
    }
    let mut aggregate = aggregate(&files);
    cap_aggregate_histograms(&mut aggregate);
    if !json_only {
        print_summary(&files, &aggregate);
    }
    println!(
        "{}",
        serde_json::to_string(&ExploreOutput { files, aggregate })?
    );
    Ok(())
}

fn survey(bytes: &[u8], name: &str) -> FileSurvey {
    let mut survey = FileSurvey {
        name: name.to_owned(),
        part_count: zip_part_count(bytes).unwrap_or(0),
        ..Default::default()
    };
    let package = match parse_vsdx(bytes) {
        Ok(package) => package,
        Err(error) => {
            survey.parse_error = Some(parse_error_kind(&error));
            return survey;
        }
    };
    survey.parse_ok = true;
    survey.part_count = zip_part_count(bytes).unwrap_or_else(|| logical_part_count(&package));
    survey.page_count = package.page_part_paths.len();
    survey.roundtrip_identical = Some(match write_vsdx(&package) {
        Ok(out) => parse_vsdx(&out).is_ok_and(|reparsed| reparsed == package),
        Err(_) => false,
    });
    survey.shape_count = package.page_contents.values().map(count_shapes).sum();
    survey.master_shape_count = package.master_contents.values().map(count_shapes).sum();
    survey.resolve_absent = resolve_absent(&package);
    let measurement = measure_formulas(&package);
    survey.evaluated = measurement.evaluated;
    survey.unsupported_known = measurement.unsupported_known;
    survey.unsupported_other = measurement.unsupported_other;
    survey.error = measurement.error;
    survey.total = measurement.total;
    survey.unsupported_constructs = measurement.unsupported_constructs;
    survey.unsupported_other_kinds = measurement.unsupported_other_kinds;
    survey.error_kinds = measurement.error_kinds;
    let render = render_pages(&package);
    survey.painted_only_shapes = render.shapes_painted_only;
    survey.placeholder_shapes = render.shapes_placeholdered;
    survey.hidden = count_hidden(&package, &render.placeholder_ids, &render.failed_pages);
    survey.unrendered = render.unrendered;
    survey.primitives_emitted = render.primitives_emitted;
    survey.primitives_painted = render.primitives_painted;
    survey.primitives_placeholdered = render.primitives_placeholdered;
    survey.placeholder_reasons = render.reasons;
    survey.render_page_errors = render.page_errors;
    let geometry = count_geometry(&package);
    survey.geometry_rows = geometry.page_rows;
    survey.master_geometry_rows = geometry.master_rows;
    let visibility = count_visibility(&package, &render.placeholder_ids);
    survey.visibility_carriers = visibility.page_carriers;
    survey.master_visibility_carriers = visibility.master_carriers;
    survey.visibility_placeholders = visibility.degraded;
    survey
}

/// Parse failure kind; drops embedded paths and content.
fn parse_error_kind(error: &VsdxError) -> String {
    match error {
        VsdxError::Container(_) => "container",
        VsdxError::MissingPart(_) => "missing part",
        VsdxError::UnsupportedDocumentKind(_) => "unsupported document kind",
        VsdxError::ConflictingMainDocumentRelationships(_) => {
            "conflicting main document relationships"
        }
        VsdxError::MalformedXml { .. } => "malformed xml",
        VsdxError::UnsafeXml { .. } => "unsafe xml",
        VsdxError::ResourceLimit { .. } => "resource limit exceeded",
        VsdxError::InvalidRelationship { .. } => "invalid relationship",
        VsdxError::InvalidSpan => "invalid span",
        VsdxError::InvalidXmlCharacter => "invalid xml character",
        VsdxError::PatchLimit { .. } => "patch limit exceeded",
        VsdxError::InvalidCellEdit { .. } => "invalid cell edit",
    }
    .to_owned()
}

fn aggregate(files: &[FileSurvey]) -> Aggregate {
    let mut total = Aggregate {
        files: files.len(),
        ..Default::default()
    };
    for file in files {
        if file.parse_ok {
            total.parse_ok += 1;
        } else {
            total.parse_failed += 1;
        }
        match file.roundtrip_identical {
            Some(true) => total.roundtrip_identical += 1,
            Some(false) => total.roundtrip_differ += 1,
            None => total.roundtrip_not_attempted += 1,
        }
        total.pages += file.page_count;
        total.shapes += file.shape_count;
        total.master_shapes += file.master_shape_count;
        merge(&mut total.resolve_absent, &file.resolve_absent);
        total.evaluated += file.evaluated;
        total.unsupported_known += file.unsupported_known;
        total.unsupported_other += file.unsupported_other;
        total.error += file.error;
        total.total += file.total;
        merge(
            &mut total.unsupported_constructs,
            &file.unsupported_constructs,
        );
        merge(
            &mut total.unsupported_other_kinds,
            &file.unsupported_other_kinds,
        );
        merge(&mut total.error_kinds, &file.error_kinds);
        total.painted_only_shapes += file.painted_only_shapes;
        total.placeholder_shapes += file.placeholder_shapes;
        total.hidden += file.hidden;
        total.unrendered += file.unrendered;
        total.primitives_emitted += file.primitives_emitted;
        total.primitives_painted += file.primitives_painted;
        total.primitives_placeholdered += file.primitives_placeholdered;
        merge(&mut total.placeholder_reasons, &file.placeholder_reasons);
        total.render_page_errors += file.render_page_errors;
        merge(&mut total.geometry_rows, &file.geometry_rows);
        merge(&mut total.master_geometry_rows, &file.master_geometry_rows);
        merge(&mut total.visibility_carriers, &file.visibility_carriers);
        merge(
            &mut total.master_visibility_carriers,
            &file.master_visibility_carriers,
        );
        merge(
            &mut total.visibility_placeholders,
            &file.visibility_placeholders,
        );
    }
    total
}

fn merge(into: &mut BTreeMap<String, usize>, from: &BTreeMap<String, usize>) {
    for (key, value) in from {
        *into.entry(key.clone()).or_default() += value;
    }
}

fn cap_file_histograms(file: &mut FileSurvey) {
    cap_map(&mut file.resolve_absent);
    cap_map(&mut file.unsupported_constructs);
    cap_map(&mut file.unsupported_other_kinds);
    cap_map(&mut file.error_kinds);
    cap_map(&mut file.placeholder_reasons);
    cap_map(&mut file.geometry_rows);
    cap_map(&mut file.master_geometry_rows);
    cap_map(&mut file.visibility_carriers);
    cap_map(&mut file.master_visibility_carriers);
    cap_map(&mut file.visibility_placeholders);
}

fn cap_aggregate_histograms(total: &mut Aggregate) {
    cap_map(&mut total.resolve_absent);
    cap_map(&mut total.unsupported_constructs);
    cap_map(&mut total.unsupported_other_kinds);
    cap_map(&mut total.error_kinds);
    cap_map(&mut total.placeholder_reasons);
    cap_map(&mut total.geometry_rows);
    cap_map(&mut total.master_geometry_rows);
    cap_map(&mut total.visibility_carriers);
    cap_map(&mut total.master_visibility_carriers);
    cap_map(&mut total.visibility_placeholders);
}

fn cap_map(map: &mut BTreeMap<String, usize>) {
    if map.len() <= HISTOGRAM_CAP {
        return;
    }
    let keep = top(map, HISTOGRAM_CAP)
        .into_iter()
        .map(|(key, _)| key)
        .collect::<BTreeSet<_>>();
    map.retain(|key, _| keep.contains(key));
}

fn print_summary(files: &[FileSurvey], total: &Aggregate) {
    for file in files {
        let roundtrip = match file.roundtrip_identical {
            Some(true) => "identical",
            Some(false) => "differ",
            None => "not attempted",
        };
        eprintln!(
            "{} parse={} pages={} shapes={} roundtrip={} formulas={}/{} shapes_painted_only={} shapes_placeholdered={} hidden={} unrendered={} primitives={}/{}/{}",
            file.name,
            if file.parse_ok { "ok" } else { "error" },
            file.page_count,
            file.shape_count,
            roundtrip,
            file.evaluated,
            file.total,
            file.painted_only_shapes,
            file.placeholder_shapes,
            file.hidden,
            file.unrendered,
            file.primitives_emitted,
            file.primitives_painted,
            file.primitives_placeholdered,
        );
        if let Some(error) = &file.parse_error {
            eprintln!("  parse error: {error}");
        }
    }
    eprintln!(
        "aggregate files={} parse_ok={} parse_failed={} pages={} shapes={} masters={} roundtrip_identical={} roundtrip_differ={} roundtrip_not_attempted={}",
        total.files,
        total.parse_ok,
        total.parse_failed,
        total.pages,
        total.shapes,
        total.master_shapes,
        total.roundtrip_identical,
        total.roundtrip_differ,
        total.roundtrip_not_attempted,
    );
    eprintln!(
        "aggregate formulas evaluated={} unsupported_known={} unsupported_other={} error={} total={}",
        total.evaluated, total.unsupported_known, total.unsupported_other, total.error, total.total,
    );
    eprintln!(
        "aggregate render shapes_painted_only={} shapes_placeholdered={} hidden={} unrendered={} page_errors={}",
        total.painted_only_shapes,
        total.placeholder_shapes,
        total.hidden,
        total.unrendered,
        total.render_page_errors,
    );
    eprintln!(
        "aggregate primitives emitted={} painted={} placeholdered={}",
        total.primitives_emitted, total.primitives_painted, total.primitives_placeholdered,
    );
    eprintln!(
        "aggregate geometry page NURBSTo={} SplineStart={} SplineKnot={} master NURBSTo={} SplineStart={} SplineKnot={}",
        row_count(&total.geometry_rows, "NURBSTo"),
        row_count(&total.geometry_rows, "SplineStart"),
        row_count(&total.geometry_rows, "SplineKnot"),
        row_count(&total.master_geometry_rows, "NURBSTo"),
        row_count(&total.master_geometry_rows, "SplineStart"),
        row_count(&total.master_geometry_rows, "SplineKnot"),
    );
    eprintln!(
        "aggregate visibility page_carriers={:?} master_carriers={:?} degraded={:?}",
        total.visibility_carriers, total.master_visibility_carriers, total.visibility_placeholders,
    );
    eprintln!(
        "aggregate resolve absent={:?}",
        top(&total.resolve_absent, 20)
    );
    eprintln!(
        "top unsupported constructs: {:?}",
        top(&total.unsupported_constructs, 20)
    );
    eprintln!(
        "top other unsupported kinds: {:?}",
        top(&total.unsupported_other_kinds, 20)
    );
    eprintln!("top error kinds: {:?}", top(&total.error_kinds, 20));
    eprintln!(
        "top placeholder reasons: {:?}",
        top(&total.placeholder_reasons, 20)
    );
    eprintln!("top geometry rows: {:?}", top(&total.geometry_rows, 20));
    eprintln!(
        "top master geometry rows: {:?}",
        top(&total.master_geometry_rows, 20)
    );
}

fn row_count(rows: &BTreeMap<String, usize>, row_type: &str) -> usize {
    rows.get(row_type).copied().unwrap_or(0)
}

fn top(map: &BTreeMap<String, usize>, limit: usize) -> Vec<(String, usize)> {
    let mut entries = map
        .iter()
        .map(|(key, value)| (key.clone(), *value))
        .collect::<Vec<_>>();
    entries.sort_by(|left, right| right.1.cmp(&left.1).then_with(|| left.0.cmp(&right.0)));
    entries.truncate(limit);
    entries
}

fn zip_part_count(bytes: &[u8]) -> Option<usize> {
    const EOCD: [u8; 4] = [0x50, 0x4b, 0x05, 0x06];
    const MIN_EOCD: usize = 22;
    if bytes.len() < MIN_EOCD {
        return None;
    }
    let start = bytes.len().saturating_sub(65_557 + MIN_EOCD);
    let mut index = bytes.len() - MIN_EOCD;
    loop {
        if bytes[index..].starts_with(&EOCD) {
            let total = u16::from_le_bytes([bytes[index + 10], bytes[index + 11]]) as usize;
            return Some(total);
        }
        if index == start {
            return None;
        }
        index -= 1;
    }
}

fn logical_part_count(package: &vsdx_parse::VsdxPackage) -> usize {
    let mut count = 1;
    if package.pages_part_path.is_some() {
        count += 1;
    }
    if package.masters_part_path.is_some() {
        count += 1;
    }
    if package.windows_part_path.is_some() {
        count += 1;
    }
    count
        + package.page_part_paths.len()
        + package.master_part_paths.len()
        + package.theme_part_paths.len()
}

fn count_shapes(sheet: &Sheet) -> usize {
    sheet.shapes().map(count_shape).sum()
}

fn count_shape(shape: &Shape) -> usize {
    1 + shape.shapes().map(count_shape).sum::<usize>()
}

fn resolve_absent(package: &vsdx_parse::VsdxPackage) -> BTreeMap<String, usize> {
    let mut absent = BTreeMap::new();
    let resolver = Resolver::new(package);
    for page in &package.page_part_paths {
        let Ok(shapes) = resolver.resolve_page_shapes(page) else {
            continue;
        };
        for shape in shapes.values() {
            for cell in KNOWN_CELLS {
                if matches!(shape.cell(cell), Some(Lookup::Absent)) {
                    *absent.entry(cell.to_owned()).or_default() += 1;
                }
            }
        }
    }
    absent
}

/// Hidden shapes on rendered pages; failed pages count as unrendered elsewhere.
fn count_hidden(
    package: &vsdx_parse::VsdxPackage,
    placeholder_ids: &BTreeSet<String>,
    failed_pages: &BTreeSet<String>,
) -> usize {
    let resolver = Resolver::new(package);
    let mut hidden = 0;
    for page in &package.page_part_paths {
        if failed_pages.contains(page) {
            continue;
        }
        let Ok(shapes) = resolver.resolve_page_shapes(page) else {
            continue;
        };
        let Some(sheet) = package.page_contents.get(page) else {
            continue;
        };
        for shape in sheet.shapes() {
            hidden += hidden_subtree(&shapes, &resolver, page, shape, placeholder_ids);
        }
    }
    hidden
}

fn hidden_subtree(
    shapes: &BTreeMap<u32, ResolvedShape>,
    resolver: &Resolver<'_>,
    page: &str,
    shape: &Shape,
    placeholder_ids: &BTreeSet<String>,
) -> usize {
    if shape_hidden(shapes, resolver, page, shape) {
        return count_shape(shape);
    }
    if placeholder_ids.contains(&format!("{page}:{}", shape.id)) {
        return shape.shapes().map(count_shape).sum();
    }
    shape
        .shapes()
        .map(|child| hidden_subtree(shapes, resolver, page, child, placeholder_ids))
        .sum()
}

fn shape_hidden(
    shapes: &BTreeMap<u32, ResolvedShape>,
    resolver: &Resolver<'_>,
    page: &str,
    shape: &Shape,
) -> bool {
    if let Some(resolved) = shapes.get(&shape.id) {
        return is_hidden(resolved);
    }
    resolver
        .resolve_shape(page, shape.id)
        .is_ok_and(|resolved| is_hidden(&resolved))
}

/// Deleted or nonzero-NoShow shapes.
fn is_hidden(resolved: &ResolvedShape) -> bool {
    if resolved.deleted {
        return true;
    }
    let Some(Lookup::Found(cell)) = resolved.cell("NoShow") else {
        return false;
    };
    cell.cell
        .value
        .as_deref()
        .and_then(|value| value.parse::<f64>().ok())
        .is_some_and(|value| value.is_finite() && value != 0.0)
}

struct Measurement {
    evaluated: usize,
    unsupported_known: usize,
    unsupported_other: usize,
    error: usize,
    total: usize,
    unsupported_constructs: BTreeMap<String, usize>,
    unsupported_other_kinds: BTreeMap<String, usize>,
    error_kinds: BTreeMap<String, usize>,
}

fn measure_formulas(package: &vsdx_parse::VsdxPackage) -> Measurement {
    let mut measurement = Measurement {
        evaluated: 0,
        unsupported_known: 0,
        unsupported_other: 0,
        error: 0,
        total: 0,
        unsupported_constructs: BTreeMap::new(),
        unsupported_other_kinds: BTreeMap::new(),
        error_kinds: BTreeMap::new(),
    };
    let limits = ParseLimits::default();
    let resolver = Resolver::new(package);
    let document = package
        .document_sheet
        .as_ref()
        .and_then(|sheet| resolver.resolve_sheet(sheet).ok());
    for sheet in package
        .document_sheet
        .iter()
        .chain(package.style_sheets.iter())
        .chain(package.page_sheets.values())
        .chain(package.master_sheets.values())
    {
        let Ok(resolved) = resolver.resolve_sheet(sheet) else {
            continue;
        };
        let refs = DocumentReferences::new(&resolved, document.as_ref());
        for (name, cell) in sheet_formula_cells(sheet) {
            let Some(formula) = cell.formula.as_deref() else {
                continue;
            };
            let evaluation =
                evaluate_cell_with_package_theme(&name, formula, &refs, &limits, package);
            record_formula(formula, evaluation, &mut measurement);
        }
    }
    for (page, sheet) in &package.page_contents {
        let owned = sheet_references(sheet);
        let refs = DocumentReferences::new(&owned, document.as_ref());
        for (name, cell) in sheet_formula_cells(sheet) {
            let Some(formula) = cell.formula.as_deref() else {
                continue;
            };
            let evaluation =
                evaluate_cell_with_package_theme(&name, formula, &refs, &limits, package);
            record_formula(formula, evaluation, &mut measurement);
        }
        let Ok(page_refs) = PageShapeReferences::new(&resolver, page) else {
            continue;
        };
        for shape in shapes_in(sheet) {
            let refs = page_refs.for_shape(shape.id);
            let Some(resolved) = page_refs.shape(shape.id) else {
                continue;
            };
            for (name, cell) in shape_formula_cells(shape) {
                let Some(formula) = cell.formula.as_deref() else {
                    continue;
                };
                let evaluation = evaluate_cell_with_shape_package_theme(
                    &name, formula, &refs, &limits, resolved, package,
                );
                record_formula(formula, evaluation, &mut measurement);
            }
        }
    }
    for sheet in package.master_contents.values() {
        let owned = sheet_references(sheet);
        let refs = DocumentReferences::new(&owned, document.as_ref());
        for (name, cell) in sheet_formula_cells(sheet) {
            let Some(formula) = cell.formula.as_deref() else {
                continue;
            };
            let evaluation =
                evaluate_cell_with_package_theme(&name, formula, &refs, &limits, package);
            record_formula(formula, evaluation, &mut measurement);
        }
        for shape in shapes_in(sheet) {
            let Ok(resolved) = resolver.resolve_shape_in_sheet(shape, sheet) else {
                continue;
            };
            for (name, cell) in shape_formula_cells(shape) {
                let Some(formula) = cell.formula.as_deref() else {
                    continue;
                };
                let evaluation = evaluate_cell_with_shape_package_theme(
                    &name,
                    formula,
                    &DocumentReferences::new(&resolved, document.as_ref()),
                    &limits,
                    &resolved,
                    package,
                );
                record_formula(formula, evaluation, &mut measurement);
            }
        }
    }
    measurement
}

fn record_formula(formula: &str, evaluation: Evaluation, measurement: &mut Measurement) {
    measurement.total += 1;
    if let Ok(expression) = eval_parse(formula, &ParseLimits::default())
        && has_unsupported(&expression)
    {
        collect_unsupported(&expression, &mut measurement.unsupported_constructs);
    }
    match evaluation {
        Evaluation::Evaluated(_) => measurement.evaluated += 1,
        Evaluation::Unsupported(reason) => {
            if is_known_deferred_reason(&reason) {
                measurement.unsupported_known += 1;
            } else {
                measurement.unsupported_other += 1;
                *measurement
                    .unsupported_other_kinds
                    .entry(classify_unsupported(&reason))
                    .or_default() += 1;
            }
        }
        Evaluation::Error(error) => {
            measurement.error += 1;
            *measurement
                .error_kinds
                .entry(classify_error(&error.message))
                .or_default() += 1;
        }
    }
}

fn is_known_deferred_call(name: &str) -> bool {
    !matches!(
        name.to_ascii_uppercase().as_str(),
        "IF" | "AND"
            | "OR"
            | "NOT"
            | "MIN"
            | "MAX"
            | "ABS"
            | "INT"
            | "ROUND"
            | "CEILING"
            | "FLOOR"
            | "SQRT"
            | "SIN"
            | "COS"
            | "TAN"
            | "ATAN2"
            | "PI"
            | "MOD"
            | "SUM"
            | "TRUNC"
            | "SIGN"
            | "RGB"
            | "TINT"
            | "MSOTINT"
            | "SAT"
            | "THEMEVAL"
            | "THEMEGUARD"
            | "_XFTRIGGER"
            | "GUARD"
    )
}

fn is_known_deferred_reason(reason: &str) -> bool {
    if reason == "Inh has no concrete inherited value" {
        return true;
    }
    let name = reason
        .strip_prefix("unsupported function ")
        .or_else(|| reason.strip_suffix(" is not implemented"))
        .or_else(|| reason.strip_suffix(" is outside the phase-4 evaluator"));
    name.is_some_and(is_known_deferred_call)
}

fn has_unsupported(expression: &Expr) -> bool {
    match expression {
        Expr::Call(name, args) => is_known_deferred_call(name) || args.iter().any(has_unsupported),
        Expr::Unary(value) => has_unsupported(value),
        Expr::Binary(left, _, right) => has_unsupported(left) || has_unsupported(right),
        _ => false,
    }
}

fn collect_unsupported(expression: &Expr, counts: &mut BTreeMap<String, usize>) {
    match expression {
        Expr::Call(name, args) => {
            if has_unsupported(&Expr::Call(name.clone(), Vec::new())) {
                *counts.entry(fold_call_name(name)).or_default() += 1;
            }
            for argument in args {
                collect_unsupported(argument, counts);
            }
        }
        Expr::Unary(value) => collect_unsupported(value, counts),
        Expr::Binary(left, _, right) => {
            collect_unsupported(left, counts);
            collect_unsupported(right, counts);
        }
        _ => {}
    }
}

/// Cross-sheet calls fold to `<sheet-ref>`; unknown names to `<unknown-function>`.
fn fold_call_name(name: &str) -> String {
    let upper = name.to_ascii_uppercase();
    let is_cross_sheet = upper.contains('!')
        && upper.split('!').next().is_some_and(|scope| {
            scope == "THEDOC"
                || scope == "THEPAGE"
                || (scope.starts_with("SHEET.")
                    && scope[6..].chars().all(|cell| cell.is_ascii_digit()))
        });
    if is_cross_sheet {
        return "<sheet-ref>".to_owned();
    }
    if KNOWN_FUNCTIONS.contains(&upper.as_str()) {
        return upper;
    }
    "<unknown-function>".to_owned()
}

/// Evaluator error class; drops document-derived tails.
fn classify_error(message: &str) -> String {
    if let Some(name) = message.strip_prefix("unresolved reference ") {
        if name.starts_with("Sheet.")
            || matches!(name.split_once('!'), Some(("ThePage" | "TheDoc", _)))
        {
            return "unresolved cross-sheet reference".to_owned();
        }
        return "unresolved cell reference".to_owned();
    }
    if message == "colour used where a numeric value is required"
        || message == "numeric value used where a colour is required"
    {
        return "type error".to_owned();
    }
    if message == "missing argument" || message.contains(" requires ") {
        return "arity error".to_owned();
    }
    if message.contains("unit")
        || message.contains("dimensional")
        || message.contains("trigonometric argument")
    {
        return "unit/dimension error".to_owned();
    }
    if message.contains("limit exceeded") {
        return "budget/depth/step exceeded".to_owned();
    }
    "other".to_owned()
}

/// Unsupported-reason class; unknown functions fold to a fixed bucket.
fn classify_unsupported(reason: &str) -> String {
    if is_static_unsupported_reason(reason) {
        return reason.to_owned();
    }
    if let Some(name) = reason.strip_prefix("unsupported function ") {
        return format!("unsupported function {}", fold_call_name(name));
    }
    if let Some(name) = reason.strip_suffix(" is not implemented") {
        return format!("not implemented: {}", fold_call_name(name));
    }
    if let Some(name) = reason.strip_suffix(" is outside the phase-4 evaluator") {
        return format!("outside phase-4 evaluator: {}", fold_call_name(name));
    }
    "other unsupported".to_owned()
}

fn is_static_unsupported_reason(reason: &str) -> bool {
    matches!(
        reason,
        "Inh has no concrete inherited value"
            | "event cell is outside the display evaluation profile"
            | "TheText requires phase-4b text layout"
            | "string values are not display numbers"
            | "SQRT of dimensional values is not implemented"
            | "THEMEVAL colour-scheme index must be 1 through 8"
            | "THEMEVAL requires a string or integer theme value"
            | "THEMEVAL host-cell lookup requires theme-cell context"
            | "THEMEVAL has no resolvable theme"
            | "unresolvable THEMEVAL value"
            | "cell value is not a supported display literal"
            | "cell value has an unsupported unit"
            | "non-finite result"
            | "SETATREFEXPR/SETATREFEVAL transformations are not implemented"
            | "SETATREF set_expression handling is not implemented"
            | "SETATREF requires a cell-reference first argument"
    )
}

fn sheet_formula_cells(sheet: &Sheet) -> Vec<(String, &Cell)> {
    let mut values = sheet
        .cells()
        .filter(|cell| cell.formula.is_some())
        .map(|cell| (cell.name.clone(), cell))
        .collect::<Vec<_>>();
    for section in sheet.sections() {
        for row in section.rows() {
            values.extend(
                row.cells()
                    .filter(|cell| cell.formula.is_some())
                    .map(|cell| (section_cell_name(section, row, cell), cell)),
            );
        }
    }
    values
}

fn shape_formula_cells(shape: &Shape) -> Vec<(String, &Cell)> {
    let mut values = shape
        .cells()
        .filter(|cell| cell.formula.is_some())
        .map(|cell| (cell.name.clone(), cell))
        .collect::<Vec<_>>();
    for section in shape.sections() {
        for row in section.rows() {
            values.extend(
                row.cells()
                    .filter(|cell| cell.formula.is_some())
                    .map(|cell| (section_cell_name(section, row, cell), cell)),
            );
        }
    }
    values
}

fn section_cell_name(section: &Section, row: &Row, cell: &Cell) -> String {
    row.name.as_ref().map_or_else(
        || format!("{}.{}", section.name, cell.name),
        |row| format!("{}.{}.{}", section.name, row, cell.name),
    )
}

fn shapes_in(sheet: &Sheet) -> Vec<&Shape> {
    let mut values = Vec::new();
    for shape in sheet.shapes() {
        collect_shapes(shape, &mut values);
    }
    values
}

fn collect_shapes<'a>(shape: &'a Shape, values: &mut Vec<&'a Shape>) {
    values.push(shape);
    for child in &shape.children {
        if let ShapeChild::Shapes(children) = child {
            for child in children {
                if let ShapesChild::Shape(shape) = child {
                    collect_shapes(shape, values);
                }
            }
        }
    }
}

fn sheet_references(sheet: &Sheet) -> BTreeMap<String, String> {
    let mut refs = references(sheet.cells().map(|cell| (cell.name.clone(), cell)));
    for section in sheet.sections() {
        for row in section.rows() {
            for cell in row.cells() {
                refs.extend(references(std::iter::once((
                    section_cell_name(section, row, cell),
                    cell,
                ))));
            }
        }
    }
    refs
}

fn references<'a>(cells: impl Iterator<Item = (String, &'a Cell)>) -> BTreeMap<String, String> {
    cells
        .filter_map(|(name, cell)| cell.formula.as_ref().map(|formula| (name, formula.clone())))
        .collect()
}

struct RenderCounts {
    shapes_painted_only: usize,
    shapes_placeholdered: usize,
    primitives_emitted: usize,
    primitives_painted: usize,
    primitives_placeholdered: usize,
    unrendered: usize,
    reasons: BTreeMap<String, usize>,
    painted_ids: BTreeSet<String>,
    placeholder_ids: BTreeSet<String>,
    failed_pages: BTreeSet<String>,
    page_errors: usize,
}

/// Shape-level render buckets; failed pages count as unrendered.
fn render_pages(package: &vsdx_parse::VsdxPackage) -> RenderCounts {
    render_pages_with_limits(package, RenderLimits::default())
}

/// Same buckets under custom limits.
fn render_pages_with_limits(
    package: &vsdx_parse::VsdxPackage,
    limits: RenderLimits,
) -> RenderCounts {
    let mut counts = RenderCounts {
        shapes_painted_only: 0,
        shapes_placeholdered: 0,
        primitives_emitted: 0,
        primitives_painted: 0,
        primitives_placeholdered: 0,
        unrendered: 0,
        reasons: BTreeMap::new(),
        painted_ids: BTreeSet::new(),
        placeholder_ids: BTreeSet::new(),
        failed_pages: BTreeSet::new(),
        page_errors: 0,
    };
    let mut renderer = Renderer::new(limits);
    if renderer
        .register_font("sans-serif", false, false, FONT_BYTES.to_vec())
        .is_err()
    {
        counts.page_errors = package.page_part_paths.len();
        for page in &package.page_part_paths {
            counts.failed_pages.insert(page.clone());
            counts.unrendered += page_shape_count(package, page);
        }
        return counts;
    }
    for page in &package.page_part_paths {
        match renderer.layout_page(package, page) {
            Ok(list) => tally_primitives(
                &list.primitives,
                &mut counts.primitives_painted,
                &mut counts.primitives_placeholdered,
                &mut counts.reasons,
                &mut counts.painted_ids,
                &mut counts.placeholder_ids,
            ),
            Err(_) => {
                counts.page_errors += 1;
                counts.failed_pages.insert(page.clone());
                counts.unrendered += page_shape_count(package, page);
            }
        }
    }
    counts.primitives_emitted = counts.primitives_painted + counts.primitives_placeholdered;
    counts.shapes_placeholdered = counts.placeholder_ids.len();
    counts.shapes_painted_only = counts
        .painted_ids
        .difference(&counts.placeholder_ids)
        .count();
    counts
}

/// Page shape count including nested children.
fn page_shape_count(package: &vsdx_parse::VsdxPackage, page: &str) -> usize {
    package
        .page_contents
        .get(page)
        .map(count_shapes)
        .unwrap_or(0)
}

/// Per-bucket shape ids plus raw primitive counts; text boxes touch neither.
fn tally_primitives(
    primitives: &[Primitive],
    painted: &mut usize,
    placeholdered: &mut usize,
    reasons: &mut BTreeMap<String, usize>,
    painted_ids: &mut BTreeSet<String>,
    placeholder_ids: &mut BTreeSet<String>,
) {
    for primitive in primitives {
        match primitive {
            Primitive::Shape { id, .. } | Primitive::Image { id, .. } => {
                *painted += 1;
                painted_ids.insert(id.clone());
            }
            Primitive::Placeholder { id, reason, .. } => {
                *placeholdered += 1;
                *reasons
                    .entry(classify_placeholder_reason(reason))
                    .or_default() += 1;
                placeholder_ids.insert(id.clone());
            }
            Primitive::Group { id, primitives, .. } => {
                *painted += 1;
                painted_ids.insert(id.clone());
                tally_primitives(
                    primitives,
                    painted,
                    placeholdered,
                    reasons,
                    painted_ids,
                    placeholder_ids,
                );
            }
            Primitive::TextBox { .. } => {}
        }
    }
}

/// Placeholder-reason class; fixed vocabulary only.
fn classify_placeholder_reason(reason: &str) -> String {
    if let Some(detail) = reason.strip_prefix("unsupported Geometry section controls") {
        return section_control_class(detail);
    }
    if let Some(detail) = reason.strip_prefix("unsupported geometry") {
        let buckets = geometry_issue_buckets(detail);
        if buckets.is_empty() {
            return "unsupported geometry".to_owned();
        }
        return buckets.join(", ");
    }
    if reason.starts_with("unresolvable colour") {
        return "unresolvable colour".to_owned();
    }
    if reason.starts_with("connector route cannot be computed") {
        return "connector route cannot be computed".to_owned();
    }
    if matches!(
        reason,
        "group nesting depth exceeded"
            | "unresolvable transform"
            | "overflowing transform"
            | "dangling ForeignData image target"
            | "unsupported ForeignData image"
            | "shape has no Geometry section"
            | "1D shape is missing connector state"
            | "non-finite stroke width"
    ) {
        return reason.to_owned();
    }
    if reason.starts_with("missing colour")
        || reason.starts_with("missing Color")
        || reason.starts_with("unavailable colour")
        || reason.contains("palette index")
    {
        return "unresolvable colour".to_owned();
    }
    let error_class = classify_error(reason);
    if error_class != "other" {
        return format!("unresolvable colour: {error_class}");
    }
    let unsupported_class = classify_unsupported(reason);
    if unsupported_class != "other unsupported" {
        return format!("unresolvable colour: {unsupported_class}");
    }
    "other placeholder".to_owned()
}

/// Fired visibility controls; drops the section index.
fn section_control_class(detail: &str) -> String {
    let mut controls = VISIBILITY_CONTROLS
        .into_iter()
        .filter(|control| detail.contains(*control))
        .collect::<Vec<_>>();
    controls.sort_unstable();
    if controls.is_empty() {
        return "unsupported Geometry section controls".to_owned();
    }
    format!(
        "unsupported Geometry section controls: {}",
        controls.join(", ")
    )
}

/// Known row types pass through; anything else folds to unknown.
fn fold_row_type(row_type: &str) -> &str {
    if KNOWN_GEOMETRY_ROWS.contains(&row_type) {
        row_type
    } else {
        "<unknown-row-type>"
    }
}

/// Geometry-issue buckets; keeps the kind, drops cell names.
fn geometry_issue_buckets(detail: &str) -> Vec<String> {
    let mut buckets = BTreeSet::new();
    collect_row_type_buckets(detail, &mut buckets);
    collect_geometry_cell_buckets(detail, "MissingCell", "missing cell in", &mut buckets);
    collect_geometry_cell_buckets(
        detail,
        "UnevaluatedCell",
        "unevaluated cell in",
        &mut buckets,
    );
    collect_geometry_control_buckets(detail, &mut buckets);
    buckets.into_iter().collect()
}

/// `UnsupportedRowType` entries by whitelisted row type.
fn collect_row_type_buckets(detail: &str, buckets: &mut BTreeSet<String>) {
    const MARKER: &str = "UnsupportedRowType(\"";
    let mut rest = detail;
    while let Some(found) = rest.find(MARKER) {
        rest = &rest[found + MARKER.len()..];
        let end = rest.find('"').unwrap_or(rest.len());
        buckets.insert(format!(
            "unsupported geometry: unimplemented row type {}",
            fold_row_type(&rest[..end])
        ));
    }
}

/// `MissingCell`/`UnevaluatedCell` entries by kind and row type.
fn collect_geometry_cell_buckets(
    detail: &str,
    variant: &str,
    kind: &str,
    buckets: &mut BTreeSet<String>,
) {
    const FIELD: &str = "row_type: \"";
    let mut rest = detail;
    while let Some(found) = rest.find(variant) {
        rest = &rest[found + variant.len()..];
        let Some(field) = rest.find(FIELD) else {
            continue;
        };
        if rest[..field].contains('}') {
            continue;
        }
        rest = &rest[field + FIELD.len()..];
        let end = rest.find('"').unwrap_or(rest.len());
        buckets.insert(format!(
            "unsupported geometry: {kind} {}",
            fold_row_type(&rest[..end])
        ));
    }
}

/// `UnsupportedSectionControl` entries by whitelisted control.
fn collect_geometry_control_buckets(detail: &str, buckets: &mut BTreeSet<String>) {
    const MARKER: &str = "UnsupportedSectionControl(\"";
    let mut rest = detail;
    while let Some(found) = rest.find(MARKER) {
        rest = &rest[found + MARKER.len()..];
        let end = rest.find('"').unwrap_or(rest.len());
        let control = &rest[..end];
        buckets.insert(format!(
            "unsupported geometry: section control {}",
            if VISIBILITY_CONTROLS.contains(&control) {
                control
            } else {
                "<unknown-control>"
            }
        ));
    }
}

struct GeometryCounts {
    page_rows: BTreeMap<String, usize>,
    master_rows: BTreeMap<String, usize>,
}

fn count_geometry(package: &vsdx_parse::VsdxPackage) -> GeometryCounts {
    let mut counts = GeometryCounts {
        page_rows: BTreeMap::new(),
        master_rows: BTreeMap::new(),
    };
    for sheet in package.page_contents.values() {
        count_geometry_rows(sheet, &mut counts.page_rows);
    }
    for sheet in package.master_contents.values() {
        count_geometry_rows(sheet, &mut counts.master_rows);
    }
    counts
}

fn count_geometry_rows(sheet: &Sheet, rows: &mut BTreeMap<String, usize>) {
    for shape in shapes_in(sheet) {
        for section in shape.sections() {
            if section.name != "Geometry" {
                continue;
            }
            for row in section.rows() {
                if row.del {
                    continue;
                }
                *rows
                    .entry(
                        row.row_type
                            .as_deref()
                            .map_or("<unknown-row-type>", fold_row_type)
                            .to_owned(),
                    )
                    .or_default() += 1;
            }
        }
    }
}

struct VisibilityCounts {
    page_carriers: BTreeMap<String, usize>,
    master_carriers: BTreeMap<String, usize>,
    degraded: BTreeMap<String, usize>,
}

fn count_visibility(
    package: &vsdx_parse::VsdxPackage,
    placeholder_ids: &BTreeSet<String>,
) -> VisibilityCounts {
    let mut counts = VisibilityCounts {
        page_carriers: BTreeMap::new(),
        master_carriers: BTreeMap::new(),
        degraded: BTreeMap::new(),
    };
    let resolver = Resolver::new(package);
    for page in &package.page_part_paths {
        let Ok(shapes) = resolver.resolve_page_shapes(page) else {
            continue;
        };
        for (id, resolved) in &shapes {
            let controls = section_controls(resolved);
            for control in &controls {
                *counts.page_carriers.entry(control.clone()).or_default() += 1;
            }
            if placeholder_ids.contains(&format!("{page}:{id}")) {
                for control in &controls {
                    *counts.degraded.entry(control.clone()).or_default() += 1;
                }
            }
        }
    }
    for sheet in package.master_contents.values() {
        for shape in shapes_in(sheet) {
            let Ok(resolved) = resolver.resolve_shape_in_sheet(shape, sheet) else {
                continue;
            };
            for control in section_controls(&resolved) {
                *counts.master_carriers.entry(control).or_default() += 1;
            }
        }
    }
    counts
}

/// Geometry section controls after master inheritance.
fn section_controls(resolved: &ResolvedShape) -> Vec<String> {
    let mut controls = Vec::new();
    for control in VISIBILITY_CONTROLS {
        let carries = resolved.sections.values().any(|section| {
            section.name == "Geometry"
                && !section.deleted
                && section
                    .unsupported_controls
                    .iter()
                    .any(|name| name == control)
        });
        if carries {
            controls.push(control.to_owned());
        }
    }
    controls
}

#[cfg(test)]
mod tests {
    use super::{
        classify_error, classify_placeholder_reason, classify_unsupported, collect_unsupported,
        count_hidden, count_shapes, fold_call_name, fold_row_type, measure_formulas,
        render_pages_with_limits, survey, tally_primitives,
    };
    use std::collections::{BTreeMap, BTreeSet};
    use vsdx_render::{Affine, Primitive};

    #[test]
    fn foundation_fixture_survey_is_consistent() {
        let bytes = include_bytes!("../../../vsdx-parse/tests/fixtures/foundation.vsdx");
        let file = survey(bytes, "foundation.vsdx");
        assert!(file.parse_ok);
        assert_eq!(file.roundtrip_identical, Some(true));
        assert_eq!(
            file.evaluated + file.unsupported_known + file.unsupported_other + file.error,
            file.total
        );
        assert_eq!(
            file.painted_only_shapes + file.placeholder_shapes + file.hidden + file.unrendered,
            file.shape_count
        );
        assert_eq!(
            file.primitives_emitted,
            file.primitives_painted + file.primitives_placeholdered
        );
    }

    #[test]
    fn failed_page_shapes_still_reconcile() {
        let bytes = include_bytes!("../../../vsdx-parse/tests/fixtures/foundation.vsdx");
        let package = vsdx_parse::parse_vsdx(bytes).expect("parse foundation fixture");
        let render = render_pages_with_limits(
            &package,
            vsdx_render::RenderLimits {
                max_shapes: 0,
                ..vsdx_render::RenderLimits::default()
            },
        );
        assert!(render.page_errors > 0);
        assert!(render.unrendered > 0);
        let shape_count: usize = package.page_contents.values().map(count_shapes).sum();
        let hidden = count_hidden(&package, &render.placeholder_ids, &render.failed_pages);
        assert_eq!(
            render.shapes_painted_only + render.shapes_placeholdered + hidden + render.unrendered,
            shape_count
        );
        assert_eq!(
            render.primitives_emitted,
            render.primitives_painted + render.primitives_placeholdered
        );
    }

    #[test]
    fn multi_section_shape_reconciles_at_shape_level() {
        let primitives = vec![
            Primitive::Shape {
                id: "page:1".to_owned(),
                z_order: 0,
                path: Vec::new(),
                fill: None,
                stroke: None,
                transform: Affine::identity(),
            },
            Primitive::Placeholder {
                id: "page:1".to_owned(),
                z_order: 1,
                x: 0.0,
                y: 0.0,
                width: 1.0,
                height: 1.0,
                reason: "shape has no Geometry section".to_owned(),
            },
            Primitive::Shape {
                id: "page:2".to_owned(),
                z_order: 2,
                path: Vec::new(),
                fill: None,
                stroke: None,
                transform: Affine::identity(),
            },
        ];
        let mut painted = 0;
        let mut placeholdered = 0;
        let mut reasons = BTreeMap::new();
        let mut painted_ids = BTreeSet::new();
        let mut placeholder_ids = BTreeSet::new();
        tally_primitives(
            &primitives,
            &mut painted,
            &mut placeholdered,
            &mut reasons,
            &mut painted_ids,
            &mut placeholder_ids,
        );
        let shapes_placeholdered = placeholder_ids.len();
        let shapes_painted_only = painted_ids.difference(&placeholder_ids).count();
        assert_eq!((painted, placeholdered), (2, 1));
        assert_eq!((shapes_painted_only, shapes_placeholdered), (1, 1));
        assert_eq!(shapes_painted_only + shapes_placeholdered, 2);
        assert_ne!(painted + placeholdered, 2);
    }

    #[test]
    fn unknown_function_and_row_names_fold_into_fixed_buckets() {
        assert_eq!(fold_call_name("SUM"), "SUM");
        assert_eq!(fold_call_name("sum"), "SUM");
        assert_eq!(fold_call_name("ACOS"), "ACOS");
        assert_eq!(fold_call_name("Sheet.1!Width"), "<sheet-ref>");
        assert_eq!(fold_call_name("EVILFUNC"), "<unknown-function>");
        assert_eq!(fold_row_type("LineTo"), "LineTo");
        assert_eq!(fold_row_type("CubBezTo"), "CubBezTo");
        assert_eq!(fold_row_type("QuadBezTo"), "QuadBezTo");
        assert_eq!(fold_row_type("RelCubBezTo"), "RelCubBezTo");
        assert_eq!(fold_row_type("RelQuadBezTo"), "RelQuadBezTo");
        assert_eq!(fold_row_type("RelEllipticalArcTo"), "RelEllipticalArcTo");
        assert_eq!(fold_row_type("EVILTYPE"), "<unknown-row-type>");
    }

    #[test]
    fn standard_but_unimplemented_rows_stay_legible() {
        assert_eq!(
            classify_placeholder_reason(
                "unsupported geometry: [UnsupportedRowType(\"RelCubBezTo\")]"
            ),
            "unsupported geometry: unimplemented row type RelCubBezTo"
        );
        assert_eq!(
            classify_placeholder_reason("unsupported geometry: [UnsupportedRowType(\"CubBezTo\")]"),
            "unsupported geometry: unimplemented row type CubBezTo"
        );
        assert_eq!(
            classify_placeholder_reason(
                "unsupported geometry: [UnsupportedRowType(\"QuadBezTo\")]"
            ),
            "unsupported geometry: unimplemented row type QuadBezTo"
        );
    }

    #[test]
    fn geometry_placeholder_buckets_keep_the_failure_kind() {
        assert_eq!(
            classify_placeholder_reason("unsupported geometry: [UnsupportedRowType(\"NURBSTo\")]"),
            "unsupported geometry: unimplemented row type NURBSTo"
        );
        assert_eq!(
            classify_placeholder_reason(
                "unsupported geometry: [MissingCell { row_type: \"LineTo\", cell: \"A\" }]"
            ),
            "unsupported geometry: missing cell in LineTo"
        );
        assert_eq!(
            classify_placeholder_reason(
                "unsupported geometry: [UnevaluatedCell { row_type: \"LineTo\", cell: \"A\" }]"
            ),
            "unsupported geometry: unevaluated cell in LineTo"
        );
        assert_eq!(
            classify_placeholder_reason(
                "unsupported geometry: [MissingCell { row_type: \"EVILTYPE\", cell: \"A\" }]"
            ),
            "unsupported geometry: missing cell in <unknown-row-type>"
        );
    }

    #[test]
    fn adversarial_names_fold_into_fixed_buckets() {
        let expression = vsdx_eval::parse("EVILFUNC(1)", &vsdx_parse::ParseLimits::default())
            .expect("parse author-defined call");
        let mut counts = BTreeMap::new();
        collect_unsupported(&expression, &mut counts);
        assert_eq!(counts.keys().collect::<Vec<_>>(), ["<unknown-function>"]);
        assert_eq!(
            classify_unsupported("unsupported function EVILFUNC"),
            "unsupported function <unknown-function>"
        );
        assert_eq!(
            classify_unsupported("EVILFUNC is not implemented"),
            "not implemented: <unknown-function>"
        );
        assert_eq!(
            classify_unsupported("EVILFUNC is outside the phase-4 evaluator"),
            "outside phase-4 evaluator: <unknown-function>"
        );
        assert_eq!(
            classify_error("unresolved reference SecretCell"),
            "unresolved cell reference"
        );
        assert_eq!(
            classify_placeholder_reason("unsupported geometry: [UnsupportedRowType(\"EVILTYPE\")]"),
            "unsupported geometry: unimplemented row type <unknown-row-type>"
        );
    }

    #[test]
    fn harness_formula_totals_match_eval_pinned_corpus() {
        let Some(directory) = std::env::var_os("VSDX_CORPUS_DIR") else {
            eprintln!("SKIPPED HARNESS FORMULA PIN: VSDX_CORPUS_DIR is unset");
            return;
        };
        let directory = std::path::PathBuf::from(directory);
        let mut evaluated = 0;
        let mut unsupported_known = 0;
        let mut unsupported_other = 0;
        let mut error = 0;
        let mut total = 0;
        for file in ["lichtsysteme.vsdx", "soundplan.vsdx"] {
            let bytes = std::fs::read(directory.join(file)).expect("read corpus file");
            let package = vsdx_parse::parse_vsdx(&bytes).expect("parse corpus package");
            let measurement = measure_formulas(&package);
            evaluated += measurement.evaluated;
            unsupported_known += measurement.unsupported_known;
            unsupported_other += measurement.unsupported_other;
            error += measurement.error;
            total += measurement.total;
        }
        assert_eq!(
            (
                evaluated,
                unsupported_known,
                unsupported_other,
                error,
                total
            ),
            (3679, 1892, 1256, 165, 6992)
        );
    }
}
