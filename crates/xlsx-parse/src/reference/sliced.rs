use super::*;
use crate::chart::{
    chart_reference_areas_sliced, drawing_claims_are_unambiguous_sliced,
    parse_relationships_sliced, unmodelled_chart_parts_sliced,
};
use crate::tree::{
    charge_bytes, copy_text, names_are_resolvable_sliced, parse_tree_sliced, retire_tree,
};
use crate::xml::find_part_sliced;
use ooxml_opc::WorkBudget;

pub(crate) async fn unpatchable_references_sliced(
    parts: &[(String, Vec<u8>)],
    content_types: &[PartContentType],
    workbook: &Workbook,
    sheet_paths: &[String],
    declined: &[String],
    work: &WorkBudget,
) -> Result<Vec<UnpatchableReference>, ParseError> {
    let mut references = Vec::new();
    if package_metadata_conforms_sliced(parts, content_types, sheet_paths, work).await {
        references =
            pivot_references_sliced(parts, content_types, workbook, sheet_paths, work).await;
        for chart in unmodelled_chart_parts_sliced(parts, content_types, workbook, work).await? {
            work.step().await;
            let bytes = if chart.claimed {
                find_part_sliced(parts, &chart.path, work).await
            } else {
                None
            };
            let areas = match bytes {
                Some(bytes) => {
                    chart_reference_areas_sliced(bytes, chart.owner.as_deref(), work).await?
                }
                None => None,
            };
            references.push(bound_sliced(chart.path, areas, workbook, work).await);
        }
    } else if package_bears_references_sliced(parts, work).await {
        for (path, _) in parts {
            work.step().await;
            references.push(bound_sliced(copy_text(path, work).await, None, workbook, work).await);
        }
    }
    for path in declined {
        let mut found = false;
        for reference in &mut references {
            work.step().await;
            if reference.part.as_str() == path {
                reference.areas = None;
                found = true;
                break;
            }
        }
        if !found {
            references.push(bound_sliced(copy_text(path, work).await, None, workbook, work).await);
        }
    }
    Ok(references)
}

async fn package_bears_references_sliced(parts: &[(String, Vec<u8>)], work: &WorkBudget) -> bool {
    for (path, _) in parts {
        work.step().await;
        charge_bytes(path.len(), work).await;
        let key = part_key(path);
        if ["xl/pivottables/", "xl/pivotcache/", "xl/charts/"]
            .iter()
            .any(|prefix| key.starts_with(prefix))
        {
            return true;
        }
    }
    let Some(bytes) = find_part_sliced(parts, "[Content_Types].xml", work).await else {
        return false;
    };
    let Ok(root) = parse_tree_sliced(bytes, work).await else {
        return false;
    };
    let result = names_a_reference_bearing_type_sliced(&root, 0, work).await;
    retire_tree(root, work).await;
    result
}

async fn names_a_reference_bearing_type_sliced(
    element: &Element,
    depth: usize,
    work: &WorkBudget,
) -> bool {
    work.step().await;
    if depth > MAX_DEPTH {
        return false;
    }
    for attribute in &element.attributes {
        work.step().await;
        charge_bytes(attribute.value.len(), work).await;
        if PIVOT_CONTENT_TYPES
            .iter()
            .chain(CHART_CONTENT_TYPES.iter())
            .any(|known| attribute.value.eq_ignore_ascii_case(known))
        {
            return true;
        }
    }
    for child in element.child_elements() {
        if Box::pin(names_a_reference_bearing_type_sliced(
            child,
            depth + 1,
            work,
        ))
        .await
        {
            return true;
        }
    }
    false
}

async fn bound_sliced(
    part: String,
    areas: Option<Vec<NamedArea>>,
    workbook: &Workbook,
    work: &WorkBudget,
) -> UnpatchableReference {
    let mut bound = None;
    if let Some(areas) = areas {
        let mut resolved = Vec::new();
        let mut valid = true;
        for (name, end) in areas {
            let mut count = 0;
            for sheet in &workbook.sheets {
                work.step().await;
                charge_bytes(name.len().min(sheet.name.len()), work).await;
                count += usize::from(sheet.name.eq_ignore_ascii_case(&name));
            }
            if count != 1 {
                valid = false;
                break;
            }
            resolved.push(ReferenceArea { sheet: name, end });
        }
        if valid {
            bound = Some(resolved);
        }
    }
    UnpatchableReference { part, areas: bound }
}

async fn pivot_references_sliced(
    parts: &[(String, Vec<u8>)],
    content_types: &[PartContentType],
    workbook: &Workbook,
    sheet_paths: &[String],
    work: &WorkBudget,
) -> Vec<UnpatchableReference> {
    let mut hosts: HashMap<String, Vec<String>> = HashMap::new();
    for (sheet, path) in workbook.sheets.iter().zip(sheet_paths) {
        work.step().await;
        for table in related_parts_sliced(parts, path, &REL_PIVOT_TABLE, work).await {
            work.step().await;
            hosts
                .entry(table)
                .or_default()
                .push(copy_text(&sheet.name, work).await);
        }
    }
    let mut pivot_parts = Vec::new();
    for (path, bytes) in parts {
        work.step().await;
        if is_pivot_part_sliced(path, content_types, work).await {
            pivot_parts.push((path, bytes, root_local_name_sliced(bytes, work).await));
        }
    }
    let mut cache_areas: HashMap<String, Option<Vec<NamedArea>>> = HashMap::new();
    let mut record_owners: HashMap<String, Option<String>> = HashMap::new();
    for (path, bytes, name) in &pivot_parts {
        work.step().await;
        if name.as_deref() != Some("pivotCacheDefinition") {
            continue;
        }
        let root = tree_sliced(bytes, work).await;
        let areas = match &root {
            Some(root) => source_areas_sliced(root, work).await,
            None => None,
        };
        cache_areas.insert(part_key(path), areas);
        if let Some(root) = &root {
            let mut id = None;
            for attribute in &root.attributes {
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
                    id = Some(attribute.value.as_str());
                    break;
                }
            }
            if let Some(id) = id
                && let Some(records) =
                    related_part_sliced(parts, path, id, &REL_PIVOT_CACHE_RECORDS, work).await
            {
                record_owners
                    .entry(records)
                    .and_modify(|owner| *owner = None)
                    .or_insert_with(|| Some(part_key(path)));
            }
        }
        if let Some(root) = root {
            retire_tree(root, work).await;
        }
    }
    let mut references = Vec::new();
    for (path, bytes, name) in pivot_parts {
        work.step().await;
        let areas = match name.as_deref() {
            Some("pivotCacheDefinition") => {
                clone_areas(
                    cache_areas.get(&part_key(path)).and_then(Option::as_ref),
                    work,
                )
                .await
            }
            Some("pivotCacheRecords") => {
                let areas = record_owners
                    .get(&part_key(path))
                    .and_then(Option::as_ref)
                    .and_then(|owner| cache_areas.get(owner))
                    .and_then(Option::as_ref);
                clone_areas(areas, work).await
            }
            Some("pivotTableDefinition") => match tree_sliced(bytes, work).await {
                Some(root) => {
                    let areas =
                        location_areas_sliced(&root, hosts.get(&part_key(path)), work).await;
                    retire_tree(root, work).await;
                    areas
                }
                None => None,
            },
            _ => None,
        };
        references.push(bound_sliced(copy_text(path, work).await, areas, workbook, work).await);
    }
    references
}

async fn clone_areas(areas: Option<&Vec<NamedArea>>, work: &WorkBudget) -> Option<Vec<NamedArea>> {
    let mut result = Vec::new();
    for (name, end) in areas? {
        work.step().await;
        result.push((copy_text(name, work).await, *end));
    }
    Some(result)
}

async fn root_local_name_sliced(bytes: &[u8], work: &WorkBudget) -> Option<String> {
    charge_bytes(bytes.len(), work).await;
    let mut reader = NsReader::from_reader(bytes);
    let config = reader.config_mut();
    config.expand_empty_elements = true;
    config.check_end_names = true;
    let mut buf = Vec::new();
    let mut depth = 0usize;
    let mut root = None;
    loop {
        work.step().await;
        let (namespace, event) = reader.read_resolved_event_into(&mut buf).ok()?;
        match event {
            Event::Start(start) => {
                depth += 1;
                if depth > MAX_DEPTH {
                    return None;
                }
                let cleared = declares_empty_default_sliced(&start, work).await?;
                if depth == 1 {
                    if root.is_some() {
                        return None;
                    }
                    root = Some(ours_local_name(namespace, &start, cleared)?);
                }
            }
            Event::End(_) => depth = depth.checked_sub(1)?,
            Event::Text(text) if depth == 0 => {
                let mut offset = 0;
                while offset < text.len() {
                    let count = work.take((text.len() - offset).div_ceil(64).min(256)).await * 64;
                    let end = (offset + count).min(text.len());
                    if !text[offset..end].iter().all(u8::is_ascii_whitespace) {
                        return None;
                    }
                    offset = end;
                }
            }
            Event::CData(_) | Event::GeneralRef(_) if depth == 0 => return None,
            Event::Eof => break,
            _ => {}
        }
        buf.clear();
    }
    (depth == 0).then_some(root).flatten()
}

async fn tree_sliced(bytes: &[u8], work: &WorkBudget) -> Option<Element> {
    let root = parse_tree_sliced(bytes, work).await.ok()?;
    if carries_alternate_content_sliced(&root, work).await {
        retire_tree(root, work).await;
        None
    } else {
        Some(root)
    }
}

async fn declares_empty_default_sliced(start: &BytesStart<'_>, work: &WorkBudget) -> Option<bool> {
    let mut cleared = false;
    for attribute in start.attributes() {
        work.step().await;
        let attribute = attribute.ok()?;
        cleared |= attribute.key.as_ref() == b"xmlns" && attribute.value.as_ref().is_empty();
    }
    Some(cleared)
}

async fn carries_alternate_content_sliced(element: &Element, work: &WorkBudget) -> bool {
    work.step().await;
    if element.local_name() == "AlternateContent" {
        return true;
    }
    for child in element.child_elements() {
        if Box::pin(carries_alternate_content_sliced(child, work)).await {
            return true;
        }
    }
    false
}

async fn source_areas_sliced(root: &Element, work: &WorkBudget) -> Option<Vec<NamedArea>> {
    let source = sole_child_sliced(root, "cacheSource", work).await?;
    let mut children = Vec::new();
    for child in source.child_elements() {
        work.step().await;
        children.push(child);
    }
    let [only] = children[..] else {
        return None;
    };
    match sole_attribute_sliced(source, "type", work)
        .await?
        .unwrap_or("worksheet")
    {
        "worksheet" if only.answers_to(&OURS, "worksheetSource") => {
            Some(vec![named_area_sliced(only, work).await?])
        }
        "consolidation" if only.answers_to(&OURS, "consolidation") => {
            consolidated_areas_sliced(only, work).await
        }
        _ => None,
    }
}

async fn consolidated_areas_sliced(
    consolidation: &Element,
    work: &WorkBudget,
) -> Option<Vec<NamedArea>> {
    for child in consolidation.child_elements() {
        work.step().await;
        if !child.answers_to(&OURS, "pages") && !child.answers_to(&OURS, "rangeSets") {
            return None;
        }
    }
    let sets = sole_child_sliced(consolidation, "rangeSets", work).await?;
    let mut areas = Vec::new();
    for set in sets.child_elements() {
        work.step().await;
        if !set.answers_to(&OURS, "rangeSet") {
            return None;
        }
        areas.push(named_area_sliced(set, work).await?);
    }
    (!areas.is_empty()).then_some(areas)
}

async fn named_area_sliced(element: &Element, work: &WorkBudget) -> Option<NamedArea> {
    for attribute in &element.attributes {
        work.step().await;
        if matches!(attribute.local_name(), "name" | "id") {
            return None;
        }
    }
    let sheet = copy_text(sole_attribute_sliced(element, "sheet", work).await??, work).await;
    let reference = sole_attribute_sliced(element, "ref", work).await??;
    charge_bytes(reference.len(), work).await;
    let (_, end) = parse_area(reference)?;
    Some((sheet, end))
}

async fn location_areas_sliced(
    root: &Element,
    hosts: Option<&Vec<String>>,
    work: &WorkBudget,
) -> Option<Vec<NamedArea>> {
    let location = sole_child_sliced(root, "location", work).await?;
    let reference = sole_attribute_sliced(location, "ref", work).await??;
    charge_bytes(reference.len(), work).await;
    let (_, end) = parse_area(reference)?;
    let end = CellRef::new(
        end.row
            .saturating_add(page_count_sliced(location, "rowPageCount", work).await?)
            .min(MAX_ROWS - 1),
        end.col
            .saturating_add(page_count_sliced(location, "colPageCount", work).await?)
            .min(MAX_COLS - 1),
    );
    let hosts = hosts?;
    if hosts.is_empty() {
        return None;
    }
    let mut areas = Vec::new();
    for host in hosts {
        work.step().await;
        areas.push((copy_text(host, work).await, end));
    }
    Some(areas)
}

async fn page_count_sliced(location: &Element, name: &str, work: &WorkBudget) -> Option<u32> {
    match sole_attribute_sliced(location, name, work).await? {
        None => Some(0),
        Some(value) => {
            charge_bytes(value.len(), work).await;
            value.trim().parse().ok()
        }
    }
}

async fn sole_attribute_sliced<'a>(
    element: &'a Element,
    name: &str,
    work: &WorkBudget,
) -> Option<Option<&'a str>> {
    let mut only = None;
    for attribute in &element.attributes {
        work.step().await;
        let vocabulary = match attribute.namespace.as_deref() {
            Some(namespace) => Vocabulary::Bound(namespace),
            None => Vocabulary::Absent,
        };
        if owned_local_name(&attribute.name, vocabulary, &OURS, Unqualified::Owned) == Some(name) {
            if only.is_some() {
                return None;
            }
            only = Some(attribute.value.as_str());
        }
    }
    Some(only)
}

async fn sole_child_sliced<'a>(
    element: &'a Element,
    local: &str,
    work: &WorkBudget,
) -> Option<&'a Element> {
    let mut only = None;
    for child in element.child_elements() {
        work.step().await;
        if child.answers_to(&OURS, local) {
            if only.is_some() {
                return None;
            }
            only = Some(child);
        }
    }
    only
}

async fn related_sliced(
    parts: &[(String, Vec<u8>)],
    path: &str,
    work: &WorkBudget,
) -> Vec<(String, String, String)> {
    let Some(rels) = find_part_sliced(parts, &relationship_part_path(path), work).await else {
        return Vec::new();
    };
    let mut related = Vec::new();
    for (id, kind, target) in parse_relationships_sliced(rels, work)
        .await
        .unwrap_or_default()
    {
        work.step().await;
        charge_bytes(target.len() + path.len(), work).await;
        related.push((
            id,
            kind,
            part_key(&resolve_part_path(directory_of(path), &target)),
        ));
    }
    related
}

async fn related_part_sliced(
    parts: &[(String, Vec<u8>)],
    path: &str,
    id: &str,
    types: &[&str],
    work: &WorkBudget,
) -> Option<String> {
    for (rel_id, kind, target) in related_sliced(parts, path, work).await {
        work.step().await;
        if rel_id == id && types.contains(&kind.as_str()) {
            return Some(target);
        }
    }
    None
}

async fn related_parts_sliced(
    parts: &[(String, Vec<u8>)],
    path: &str,
    types: &[&str],
    work: &WorkBudget,
) -> Vec<String> {
    let mut targets = Vec::new();
    for (_, kind, target) in related_sliced(parts, path, work).await {
        work.step().await;
        if types.contains(&kind.as_str()) {
            targets.push(target);
        }
    }
    targets
}

async fn narrowing_drawings_sliced(
    parts: &[(String, Vec<u8>)],
    sheet_paths: &[String],
    work: &WorkBudget,
) -> Vec<String> {
    let mut drawings = Vec::new();
    for sheet in sheet_paths {
        work.step().await;
        for drawing in related_parts_sliced(parts, sheet, &REL_DRAWING, work).await {
            work.step().await;
            drawings.push(drawing);
        }
    }
    drawings
}

async fn package_metadata_conforms_sliced(
    parts: &[(String, Vec<u8>)],
    content_types: &[PartContentType],
    sheet_paths: &[String],
    work: &WorkBudget,
) -> bool {
    if !content_types_conform_sliced(parts, work).await {
        return false;
    }
    let mut owners = vec!["xl/workbook.xml".to_owned()];
    for sheet in sheet_paths {
        work.step().await;
        owners.push(copy_text(sheet, work).await);
    }
    for drawing in narrowing_drawings_sliced(parts, sheet_paths, work).await {
        work.step().await;
        owners.push(drawing);
    }
    for (path, _) in parts {
        work.step().await;
        if is_pivot_part_sliced(path, content_types, work).await {
            owners.push(copy_text(path, work).await);
        }
    }
    for owner in owners {
        work.step().await;
        if !relationships_conform_sliced(parts, &owner, work).await {
            return false;
        }
    }
    for drawing in narrowing_drawings_sliced(parts, sheet_paths, work).await {
        work.step().await;
        if let Some(bytes) = find_part_sliced(parts, &drawing, work).await {
            let Ok(root) = parse_tree_sliced(bytes, work).await else {
                return false;
            };
            let conforms = names_are_resolvable_sliced(&root, work).await
                && drawing_claims_are_unambiguous_sliced(&root, work).await;
            retire_tree(root, work).await;
            if !conforms {
                return false;
            }
        }
    }
    sheet_relationships_are_unambiguous_sliced(parts, work).await
}

async fn sole_and_ours_sliced(
    element: &Element,
    namespaces: &[&str],
    name: &str,
    work: &WorkBudget,
) -> bool {
    let mut named = 0;
    let mut ours = 0;
    for attribute in &element.attributes {
        work.step().await;
        if attribute.local_name() == name {
            named += 1;
        }
        let vocabulary = match attribute.namespace.as_deref() {
            Some(namespace) => Vocabulary::Bound(namespace),
            None => Vocabulary::Absent,
        };
        if owned_local_name(&attribute.name, vocabulary, namespaces, Unqualified::Owned)
            == Some(name)
        {
            ours += 1;
        }
    }
    named <= 1 && ours == named
}

async fn attribute_local_sliced<'a>(
    element: &'a Element,
    name: &str,
    work: &WorkBudget,
) -> Option<&'a str> {
    for attribute in &element.attributes {
        work.step().await;
        if attribute.local_name() == name {
            return Some(attribute.value.as_str());
        }
    }
    None
}

async fn content_types_conform_sliced(parts: &[(String, Vec<u8>)], work: &WorkBudget) -> bool {
    let Some(bytes) = find_part_sliced(parts, "[Content_Types].xml", work).await else {
        return false;
    };
    let Ok(root) = parse_tree_sliced(bytes, work).await else {
        return false;
    };
    let result = content_types_root_conforms_sliced(&root, work).await;
    retire_tree(root, work).await;
    result
}

async fn content_types_root_conforms_sliced(root: &Element, work: &WorkBudget) -> bool {
    if !root.answers_to(&CONTENT_TYPES, "Types") {
        return false;
    }
    let (mut overrides, mut defaults) = (HashSet::new(), HashSet::new());
    for entry in root.child_elements() {
        work.step().await;
        if entry.child_elements().next().is_some() {
            return false;
        }
        for name in ["PartName", "ContentType", "Extension"] {
            if !sole_and_ours_sliced(entry, &CONTENT_TYPES, name, work).await {
                return false;
            }
        }
        if !attribute_local_sliced(entry, "ContentType", work)
            .await
            .is_some_and(|kind| !kind.trim().is_empty())
        {
            return false;
        }
        if entry.answers_to(&CONTENT_TYPES, "Override") {
            let Some(part) = attribute_local_sliced(entry, "PartName", work).await else {
                return false;
            };
            charge_bytes(part.len(), work).await;
            if part.len() <= 1 || !part.starts_with('/') || !overrides.insert(part_key(part)) {
                return false;
            }
        } else if entry.answers_to(&CONTENT_TYPES, "Default") {
            let Some(extension) = attribute_local_sliced(entry, "Extension", work).await else {
                return false;
            };
            charge_bytes(extension.len(), work).await;
            if extension.is_empty()
                || extension.contains('/')
                || !defaults.insert(extension.to_ascii_lowercase())
            {
                return false;
            }
        } else {
            return false;
        }
    }
    true
}

async fn relationships_conform_sliced(
    parts: &[(String, Vec<u8>)],
    owner: &str,
    work: &WorkBudget,
) -> bool {
    let Some(bytes) = find_part_sliced(parts, &relationship_part_path(owner), work).await else {
        return true;
    };
    let Ok(root) = parse_tree_sliced(bytes, work).await else {
        return false;
    };
    let result = relationships_root_conforms_sliced(&root, work).await;
    retire_tree(root, work).await;
    result
}

async fn relationships_root_conforms_sliced(root: &Element, work: &WorkBudget) -> bool {
    if root.namespace() != Some(NS_PACKAGE_RELATIONSHIPS) || root.local_name() != "Relationships" {
        return false;
    }
    let mut ids = HashSet::new();
    for entry in root.child_elements() {
        work.step().await;
        if entry.namespace() != Some(NS_PACKAGE_RELATIONSHIPS)
            || entry.local_name() != "Relationship"
            || entry.child_elements().next().is_some()
        {
            return false;
        }
        for name in ["Id", "Type", "Target", "TargetMode"] {
            if !sole_and_ours_sliced(entry, &RELATIONSHIPS, name, work).await {
                return false;
            }
        }
        if !attribute_local_sliced(entry, "Type", work)
            .await
            .is_some_and(followed_type_is_exact)
        {
            return false;
        }
        for name in ["Type", "Target"] {
            if !attribute_local_sliced(entry, name, work)
                .await
                .is_some_and(|value| !value.is_empty())
            {
                return false;
            }
        }
        if attribute_local_sliced(entry, "TargetMode", work)
            .await
            .is_some_and(|mode| !matches!(mode, "Internal" | "External"))
        {
            return false;
        }
        let Some(id) = attribute_local_sliced(entry, "Id", work).await else {
            return false;
        };
        if id.is_empty() || !ids.insert(copy_text(id, work).await) {
            return false;
        }
    }
    true
}

async fn sheet_relationships_are_unambiguous_sliced(
    parts: &[(String, Vec<u8>)],
    work: &WorkBudget,
) -> bool {
    let Some(bytes) = find_part_sliced(parts, "xl/workbook.xml", work).await else {
        return false;
    };
    let Ok(root) = parse_tree_sliced(bytes, work).await else {
        return false;
    };
    let result = sheet_relationships_root_are_unambiguous_sliced(&root, work).await;
    retire_tree(root, work).await;
    result
}

async fn sheet_relationships_root_are_unambiguous_sliced(
    root: &Element,
    work: &WorkBudget,
) -> bool {
    let Some(sheets) = sole_child_sliced(root, "sheets", work).await else {
        return false;
    };
    let mut canonical = Vec::new();
    for sheet in sheets.child_elements() {
        work.step().await;
        if sheet.local_name() == "sheet" {
            canonical.push(sheet);
        }
    }
    if named_sheet_elements_sliced(root, 0, work).await != canonical.len() {
        return false;
    }
    for sheet in canonical {
        work.step().await;
        if !sheet.answers_to(&OURS, "sheet") {
            return false;
        }
        let mut named = 0;
        let mut related = false;
        for attribute in &sheet.attributes {
            work.step().await;
            if attribute.local_name() == "id" {
                named += 1;
            }
            let vocabulary = match attribute.namespace.as_deref() {
                Some(namespace) => Vocabulary::Bound(namespace),
                None => Vocabulary::Absent,
            };
            if owned_local_name(
                &attribute.name,
                vocabulary,
                &[NS_RELATIONSHIPS],
                Unqualified::Owned,
            ) == Some("id")
            {
                related = true;
            }
        }
        if named != 1
            || !related
            || !sole_and_ours_sliced(sheet, &OURS, "name", work).await
            || !sole_and_ours_sliced(sheet, &OURS, "sheetId", work).await
        {
            return false;
        }
    }
    true
}

async fn named_sheet_elements_sliced(element: &Element, depth: usize, work: &WorkBudget) -> usize {
    work.step().await;
    if depth > MAX_DEPTH {
        return usize::MAX;
    }
    let mut count = usize::from(element.local_name() == "sheet");
    for child in element.child_elements() {
        count += Box::pin(named_sheet_elements_sliced(child, depth + 1, work)).await;
    }
    count
}

async fn is_pivot_part_sliced(
    path: &str,
    content_types: &[PartContentType],
    work: &WorkBudget,
) -> bool {
    charge_bytes(path.len(), work).await;
    let key = part_key(path);
    if !key.contains("/_rels/")
        && PIVOT_DIRECTORIES
            .iter()
            .any(|prefix| key.starts_with(prefix))
    {
        return true;
    }
    for part in content_types {
        work.step().await;
        if part.path == key
            && PIVOT_CONTENT_TYPES
                .iter()
                .any(|known| part.content_type.eq_ignore_ascii_case(known))
        {
            return true;
        }
    }
    false
}
