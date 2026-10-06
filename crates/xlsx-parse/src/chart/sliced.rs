use super::*;
use crate::tree::{
    Unqualified, Vocabulary, charge_bytes, copy_text, names_are_resolvable_sliced,
    owned_local_name, parse_tree_sliced, text_content_sliced,
};
use crate::xml::find_part_sliced;
use ooxml_opc::WorkBudget;

pub(crate) async fn parse_sheet_charts_sliced(
    parts: &[(String, Vec<u8>)],
    sheet_path: &str,
    declined: &mut Vec<String>,
    work: &WorkBudget,
) -> Result<Vec<SheetChart>, ParseError> {
    let mut charts = Vec::new();
    for (drawing_path, drawing_xml) in sheet_drawings_sliced(parts, sheet_path, work).await? {
        let drawing_rels =
            match find_part_sliced(parts, &relationship_part_path(&drawing_path), work).await {
                Some(bytes) => parse_relationships_sliced(bytes, work).await?,
                None => Vec::new(),
            };
        let drawing_dir = directory_of(&drawing_path).to_owned();
        let Ok(drawing_root) = parse_tree_sliced(drawing_xml, work).await else {
            declined.push(drawing_path);
            continue;
        };
        let Ok(anchors) = read_anchors_sliced(&drawing_root, work).await else {
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
            work.step().await;
            let Some(part) = relationship_target_sliced(&drawing_rels, rel_id, TYPE_CHART, work)
                .await
                .map(|target| resolve_part_path(&drawing_dir, target))
            else {
                declined.push(drawing_path.clone());
                continue;
            };
            let Some(chart_xml) = find_part_sliced(parts, &part, work).await else {
                declined.push(part);
                continue;
            };
            let Ok(root) = parse_tree_sliced(chart_xml, work).await else {
                declined.push(part);
                continue;
            };
            if !root.is(NS_CHART, "chartSpace") {
                declined.push(part);
                continue;
            }
            let Ok(refs) = chart_refs_sliced(&root, work).await else {
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

async fn sheet_drawings_sliced<'a>(
    parts: &'a [(String, Vec<u8>)],
    sheet_path: &str,
    work: &WorkBudget,
) -> Result<Vec<(String, &'a [u8])>, ParseError> {
    let Some(rels) = find_part_sliced(parts, &relationship_part_path(sheet_path), work).await
    else {
        return Ok(Vec::new());
    };
    let mut drawings: Vec<(String, &[u8])> = Vec::new();
    for (_, kind, target) in parse_relationships_sliced(rels, work).await? {
        work.step().await;
        if !type_is(&kind, TYPE_DRAWING) {
            continue;
        }
        let path = resolve_part_path(directory_of(sheet_path), &target);
        let mut found = false;
        for (walked, _) in &drawings {
            work.step().await;
            if *walked == path {
                found = true;
                break;
            }
        }
        if !found && let Some(bytes) = find_part_sliced(parts, &path, work).await {
            drawings.push((path, bytes));
        }
    }
    Ok(drawings)
}

async fn child_sliced<'a>(
    element: &'a Element,
    local: &str,
    work: &WorkBudget,
) -> Option<&'a Element> {
    for child in element.child_elements() {
        work.step().await;
        if child.local_name() == local {
            return Some(child);
        }
    }
    None
}

async fn read_anchors_sliced(
    root: &Element,
    work: &WorkBudget,
) -> Result<Vec<DrawingAnchor>, ParseError> {
    if !root.is(NS_SPREADSHEET_DRAWING, "wsDr") {
        return Ok(Vec::new());
    }
    let mut anchors = Vec::new();
    for child in root.child_elements() {
        work.step().await;
        if !is_anchor(&child) {
            continue;
        }
        let anchor = anchor_geometry_sliced(child, work).await?;
        if anchors.len() >= MAX_CHART_ANCHORS {
            return Err(ParseError::TooManyCharts);
        }
        anchors.push(DrawingAnchor {
            anchor,
            chart: anchor_chart_sliced(child, 0, work).await,
        });
    }
    Ok(anchors)
}

async fn anchor_geometry_sliced(
    anchor: &Element,
    work: &WorkBudget,
) -> Result<ChartAnchor, ParseError> {
    Ok(match anchor.local_name() {
        "twoCellAnchor" => ChartAnchor::TwoCell {
            from: anchor_cell_sliced(child_sliced(anchor, "from", work).await, work).await?,
            to: anchor_cell_sliced(child_sliced(anchor, "to", work).await, work).await?,
            edit_as: match attribute_sliced(anchor, "editAs", work).await {
                Some(value) => AnchorEditAs::from_sml(value).ok_or_else(|| {
                    ParseError::Malformed(format!("invalid chart editAs value {value:?}"))
                })?,
                None => AnchorEditAs::default(),
            },
        },
        "oneCellAnchor" => ChartAnchor::OneCell {
            from: anchor_cell_sliced(child_sliced(anchor, "from", work).await, work).await?,
            extent: anchor_extent_sliced(child_sliced(anchor, "ext", work).await, work).await?,
        },
        _ => ChartAnchor::Absolute {
            pos: anchor_pos_sliced(child_sliced(anchor, "pos", work).await, work).await?,
            extent: anchor_extent_sliced(child_sliced(anchor, "ext", work).await, work).await?,
        },
    })
}

async fn attribute_sliced<'a>(
    element: &'a Element,
    local: &str,
    work: &WorkBudget,
) -> Option<&'a str> {
    for attribute in &element.attributes {
        work.step().await;
        if attribute.name == local {
            return Some(&attribute.value);
        }
    }
    None
}

async fn anchor_chart_sliced(element: &Element, depth: usize, work: &WorkBudget) -> AnchorChart {
    work.step().await;
    if depth > MAX_DEPTH {
        return AnchorChart::None;
    }
    if element.is(NS_CHART, "chart") {
        for attribute in &element.attributes {
            work.step().await;
            let vocabulary = match attribute.namespace.as_deref() {
                Some(namespace) => Vocabulary::Bound(namespace),
                None => Vocabulary::Absent,
            };
            if owned_local_name(
                &attribute.name,
                vocabulary,
                &[NS_RELATIONSHIPS],
                Unqualified::Foreign,
            ) == Some("id")
            {
                return AnchorChart::Related(crate::tree::copy_text(&attribute.value, work).await);
            }
        }
        return AnchorChart::Unrelated;
    }
    for child in element.child_elements() {
        match Box::pin(anchor_chart_sliced(child, depth + 1, work)).await {
            AnchorChart::None => {}
            found => return found,
        }
    }
    AnchorChart::None
}

async fn child_number_sliced(
    element: &Element,
    local: &str,
    work: &WorkBudget,
) -> Result<i64, ParseError> {
    let Some(child) = child_sliced(element, local, work).await else {
        return Ok(0);
    };
    text_content_sliced(child, work)
        .await
        .trim()
        .parse::<i64>()
        .map_err(|_| ParseError::Malformed(format!("invalid chart anchor {local}")))
}

async fn child_index_sliced(
    element: &Element,
    local: &str,
    limit: u32,
    work: &WorkBudget,
) -> Result<u32, ParseError> {
    let value = match child_sliced(element, local, work).await {
        Some(child) => text_content_sliced(child, work)
            .await
            .trim()
            .parse::<u32>()
            .ok(),
        None => None,
    };
    value
        .filter(|value| *value < limit)
        .ok_or_else(|| ParseError::Malformed(format!("invalid chart anchor {local}")))
}

async fn attribute_number_sliced(
    element: &Element,
    name: &str,
    work: &WorkBudget,
) -> Result<i64, ParseError> {
    attribute_sliced(element, name, work)
        .await
        .and_then(|value| value.trim().parse::<i64>().ok())
        .ok_or_else(|| ParseError::Malformed(format!("invalid chart anchor {name}")))
}

async fn anchor_cell_sliced(
    element: Option<&Element>,
    work: &WorkBudget,
) -> Result<AnchorCell, ParseError> {
    let element =
        element.ok_or_else(|| ParseError::Malformed("chart anchor has no cell".into()))?;
    Ok(AnchorCell {
        col: child_index_sliced(element, "col", MAX_COLS, work).await?,
        col_off: child_number_sliced(element, "colOff", work).await?,
        row: child_index_sliced(element, "row", MAX_ROWS, work).await?,
        row_off: child_number_sliced(element, "rowOff", work).await?,
    })
}

async fn anchor_extent_sliced(
    element: Option<&Element>,
    work: &WorkBudget,
) -> Result<AnchorExtent, ParseError> {
    let element =
        element.ok_or_else(|| ParseError::Malformed("chart anchor has no extent".into()))?;
    Ok(AnchorExtent {
        cx: attribute_number_sliced(element, "cx", work).await?,
        cy: attribute_number_sliced(element, "cy", work).await?,
    })
}

async fn anchor_pos_sliced(
    element: Option<&Element>,
    work: &WorkBudget,
) -> Result<AnchorPos, ParseError> {
    let element = element
        .ok_or_else(|| ParseError::Malformed("absolute chart anchor has no position".into()))?;
    Ok(AnchorPos {
        x: attribute_number_sliced(element, "x", work).await?,
        y: attribute_number_sliced(element, "y", work).await?,
    })
}

async fn chart_refs_sliced(root: &Element, work: &WorkBudget) -> Result<Vec<ChartRef>, ParseError> {
    let mut refs = Vec::new();
    for site in ref_sites_sliced(root, work).await? {
        work.step().await;
        refs.push(ChartRef {
            kind: site.kind,
            formula: site.formula,
        });
    }
    Ok(refs)
}

async fn ref_sites_sliced(root: &Element, work: &WorkBudget) -> Result<Vec<RefSite>, ParseError> {
    let mut sites = Vec::new();
    let mut series = Series {
        current: None,
        next: 0,
    };
    walk_refs_sliced(root, ChartRefKind::Other, &mut series, 0, &mut sites, work).await?;
    Ok(sites)
}

async fn walk_refs_sliced(
    element: &Element,
    inherited: ChartRefKind,
    series: &mut Series,
    depth: usize,
    out: &mut Vec<RefSite>,
    work: &WorkBudget,
) -> Result<(), ParseError> {
    work.step().await;
    if depth > MAX_DEPTH {
        return Err(ParseError::DepthExceeded);
    }
    let kind = slot_kind(element.local_name()).unwrap_or(inherited);
    if element.is(NS_CHART, "f") {
        if out.len() >= MAX_CHART_REFS {
            return Err(ParseError::TooManyCharts);
        }
        let formula = text_content_sliced(element, work).await;
        charge_bytes(formula.len(), work).await;
        let formula = copy_text(formula.trim(), work).await;
        out.push(RefSite {
            kind,
            formula,
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
        Box::pin(walk_refs_sliced(child, kind, series, depth + 1, out, work)).await?;
    }
    series.current = enclosing;
    if out.len() == before + 1
        && child_sliced(element, FORMULA_LOCAL, work)
            .await
            .is_some_and(|child| child.is(NS_CHART, FORMULA_LOCAL))
        && let Some(cache) = cache_site_sliced(element, work).await
    {
        out[before].cache = Some(cache);
    }
    Ok(())
}

async fn cache_site_sliced(reference: &Element, work: &WorkBudget) -> Option<CacheSite> {
    let mut cache = None;
    for child in reference.child_elements() {
        work.step().await;
        if child.namespace() == Some(NS_CHART)
            && matches!(
                child.local_name(),
                "numCache" | "strCache" | "multiLvlStrCache"
            )
        {
            cache = Some(child);
            break;
        }
    }
    let cache = cache?;
    let format_code = match child_sliced(cache, "formatCode", work).await {
        Some(element) => Some(text_content_sliced(element, work).await),
        None => None,
    };
    let mut authored_points = 0;
    for child in cache.child_elements() {
        work.step().await;
        authored_points += usize::from(child.is(NS_CHART, "pt"));
    }
    Some(CacheSite {
        local: cache.local_name().to_owned(),
        prefix: cache
            .name
            .rsplit_once(':')
            .map(|(prefix, _)| format!("{prefix}:"))
            .unwrap_or_default(),
        span: cache.splice_target(),
        format_code,
        unmodelled_content: !cache_is_fully_modelled_sliced(cache, work).await,
        authored_points,
    })
}

async fn cache_is_fully_modelled_sliced(cache: &Element, work: &WorkBudget) -> bool {
    for child in cache.child_elements() {
        work.step().await;
        if child.namespace() != Some(NS_CHART) {
            return false;
        }
        let expected = match child.local_name() {
            "formatCode" => {
                if !child.attributes.is_empty() {
                    return false;
                }
                continue;
            }
            "ptCount" => "val",
            "pt" => "idx",
            _ => return false,
        };
        for attribute in &child.attributes {
            work.step().await;
            if attribute.name != expected {
                return false;
            }
        }
        if expected == "idx" {
            for value in child.child_elements() {
                work.step().await;
                if !value.is(NS_CHART, "v") || !value.attributes.is_empty() {
                    return false;
                }
            }
        }
    }
    true
}

pub(crate) async fn unmodelled_chart_parts_sliced(
    parts: &[(String, Vec<u8>)],
    content_types: &[PartContentType],
    workbook: &xlsx_model::Workbook,
    work: &WorkBudget,
) -> Result<Vec<UnmodelledChart>, ParseError> {
    let mut claims: std::collections::HashMap<&str, Vec<&str>> = std::collections::HashMap::new();
    for sheet in &workbook.sheets {
        for chart in &sheet.charts {
            work.step().await;
            claims
                .entry(normalize_part_path(&chart.part))
                .or_default()
                .push(sheet.name.as_str());
        }
    }
    let mut unmodelled = Vec::new();
    for (path, bytes) in parts {
        work.step().await;
        if !is_chart_part_sliced(path, content_types, work).await {
            continue;
        }
        let owners = claims.get(normalize_part_path(path));
        let claimed = owners.is_some();
        if claimed {
            let root = parse_tree_sliced(bytes, work).await?;
            if names_are_resolvable_sliced(&root, work).await
                && !unsupported_reference_form_sliced(&root, 0, work).await
                && !holds_an_unrebuildable_cache_sliced(&root, work).await?
            {
                continue;
            }
        }
        let owner = match owners.filter(|owners| owners.len() == 1) {
            Some(owners) => Some(copy_text(owners[0], work).await),
            None => None,
        };
        unmodelled.push(UnmodelledChart {
            path: copy_text(path, work).await,
            owner,
            claimed,
        });
    }
    Ok(unmodelled)
}
pub(crate) async fn chart_reference_areas_sliced(
    part: &[u8],
    owner: Option<&str>,
    work: &WorkBudget,
) -> Result<Option<Vec<(String, CellRef)>>, ParseError> {
    let root = parse_tree_sliced(part, work).await?;
    if !names_are_resolvable_sliced(&root, work).await
        || unsupported_reference_form_sliced(&root, 0, work).await
    {
        return Ok(None);
    }
    let mut areas = Vec::new();
    for site in ref_sites_sliced(&root, work).await? {
        work.step().await;
        charge_bytes(site.formula.len(), work).await;
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

async fn holds_an_unrebuildable_cache_sliced(
    root: &Element,
    work: &WorkBudget,
) -> Result<bool, ParseError> {
    for site in ref_sites_sliced(root, work).await? {
        work.step().await;
        charge_bytes(site.formula.len(), work).await;
        if site.cache.as_ref().is_some_and(|cache| {
            cache.local == "multiLvlStrCache" || !is_direct_one_dimensional_range(&site.formula)
        }) {
            return Ok(true);
        }
    }
    Ok(false)
}

async fn is_chart_part_sliced(
    path: &str,
    content_types: &[PartContentType],
    work: &WorkBudget,
) -> bool {
    charge_bytes(path.len(), work).await;
    let normalized = normalize_part_path(path).to_ascii_lowercase();
    if normalized.starts_with("xl/charts/chart")
        && normalized.ends_with(".xml")
        && !normalized.contains("/_rels/")
    {
        return true;
    }
    for part in content_types {
        work.step().await;
        if part.path == normalized
            && CHART_CONTENT_TYPES
                .iter()
                .any(|known| part.content_type.eq_ignore_ascii_case(known))
        {
            return true;
        }
    }
    false
}

pub(crate) async fn drawing_claims_are_unambiguous_sliced(
    root: &Element,
    work: &WorkBudget,
) -> bool {
    if !root.is(NS_SPREADSHEET_DRAWING, "wsDr") {
        return true;
    }
    for anchor in root.child_elements() {
        work.step().await;
        if is_anchor(&anchor) && chart_claims_sliced(anchor, 0, work).await > 1 {
            return false;
        }
    }
    true
}

async fn chart_claims_sliced(element: &Element, depth: usize, work: &WorkBudget) -> usize {
    work.step().await;
    if depth > MAX_DEPTH {
        return usize::MAX;
    }
    let mut count: usize = 0;
    if element.is(NS_CHART, "chart") {
        for attribute in &element.attributes {
            work.step().await;
            if attribute.local_name() == "id"
                && attribute.namespace.as_deref() == Some(NS_RELATIONSHIPS)
            {
                count += 1;
            }
        }
        count = count.max(1);
    }
    for child in element.child_elements() {
        count += Box::pin(chart_claims_sliced(child, depth + 1, work)).await;
    }
    count
}

async fn unsupported_reference_form_sliced(
    element: &Element,
    depth: usize,
    work: &WorkBudget,
) -> bool {
    work.step().await;
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
    if unsupported {
        return true;
    }
    for child in element.child_elements() {
        if Box::pin(unsupported_reference_form_sliced(child, depth + 1, work)).await {
            return true;
        }
    }
    false
}

async fn relationship_target_sliced<'a>(
    rels: &'a [(String, String, String)],
    id: &str,
    suffix: &str,
    work: &WorkBudget,
) -> Option<&'a str> {
    for (rel_id, kind, target) in rels {
        work.step().await;
        if rel_id == id && type_is(kind, suffix) {
            return Some(target);
        }
    }
    None
}

pub(crate) async fn parse_relationships_sliced(
    data: &[u8],
    work: &WorkBudget,
) -> Result<Vec<(String, String, String)>, ParseError> {
    let root = parse_tree_sliced(data, work).await?;
    let mut rels = Vec::new();
    for child in root.child_elements() {
        work.step().await;
        if child.local_name() != "Relationship"
            || !matches!(child.namespace(), None | Some(NS_PACKAGE_RELATIONSHIPS))
        {
            continue;
        }
        let mut id = None;
        let mut kind = None;
        let mut target = None;
        let mut mode = None;
        for attribute in &child.attributes {
            work.step().await;
            match attribute.local_name() {
                "Id" if id.is_none() => id = Some(attribute.value.as_str()),
                "Type" if kind.is_none() => kind = Some(attribute.value.as_str()),
                "Target" if target.is_none() => target = Some(attribute.value.as_str()),
                "TargetMode" if mode.is_none() => mode = Some(attribute.value.as_str()),
                _ => {}
            }
        }
        if mode.is_some_and(|mode| mode.eq_ignore_ascii_case("external")) {
            continue;
        }
        if let (Some(id), Some(target)) = (id, target) {
            rels.push((
                crate::tree::copy_text(id, work).await,
                crate::tree::copy_text(kind.unwrap_or_default(), work).await,
                crate::tree::copy_text(target, work).await,
            ));
        }
    }
    Ok(rels)
}
