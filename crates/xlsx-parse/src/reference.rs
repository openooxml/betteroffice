//! What the preserved parts no save rewrites — pivot caches, pivot tables and
//! the charts the model does not cover — actually name, so a structural edit is
//! refused only when it would move one of those cells rather than whenever such
//! a part is anywhere in the package.
//!
//! Every resolution here is all or nothing. A source this crate does not read
//! whole, a sheet the workbook does not hold exactly one of, a relationships
//! part that will not parse: each leaves the part naming every cell of every
//! sheet, which refuses everything, rather than a narrow area that would let a
//! stranding edit through.

use std::collections::{HashMap, HashSet};

use quick_xml::NsReader;
use quick_xml::events::{BytesStart, Event};
use quick_xml::name::ResolveResult;
use xlsx_model::addr::{MAX_COLS, MAX_ROWS};
use xlsx_model::{CellRef, Workbook};

use crate::chart::{
    CHART_CONTENT_TYPES, NS_PACKAGE_RELATIONSHIPS, NS_RELATIONSHIPS, chart_reference_areas,
    directory_of, drawing_claims_are_unambiguous, parse_area, parse_relationships,
    relationship_part_path, unmodelled_chart_parts,
};
use crate::package::PartContentType;
use crate::tree::{
    Element, Unqualified, Vocabulary, names_are_resolvable, owned_local_name, parse_tree,
};
use crate::write::{NS_MAIN, NS_STRICT_MAIN};
use crate::xml::{find_part, resolve_part_path};
use crate::{MAX_DEPTH, ParseError};

/// the spreadsheetml vocabulary a pivot part is read in. Transitional and
/// Strict spell the same names in different namespaces and this crate reads
/// both; an unprefixed name with no namespace at all is read as ours too,
/// because a part that declares none is still the vocabulary it looks like.
const OURS: [&str; 2] = [NS_MAIN, NS_STRICT_MAIN];

/// the package-relationships vocabulary a `.rels` part is written in.
const RELATIONSHIPS: [&str; 1] = [NS_PACKAGE_RELATIONSHIPS];

/// the vocabulary `[Content_Types].xml` is written in.
const CONTENT_TYPES: [&str; 1] = ["http://schemas.openxmlformats.org/package/2006/content-types"];

/// the relationship types the pivot path follows, in both the Transitional and
/// Strict spellings a package may use.
const REL_PIVOT_CACHE_RECORDS: [&str; 2] = [
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotCacheRecords",
    "http://purl.oclc.org/ooxml/officeDocument/relationships/pivotCacheRecords",
];
const REL_DRAWING: [&str; 2] = [
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing",
    "http://purl.oclc.org/ooxml/officeDocument/relationships/drawing",
];
const REL_CHART: [&str; 2] = [
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart",
    "http://purl.oclc.org/ooxml/officeDocument/relationships/chart",
];
const REL_PIVOT_TABLE: [&str; 2] = [
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotTable",
    "http://purl.oclc.org/ooxml/officeDocument/relationships/pivotTable",
];

/// A preserved part naming sheets and cells neither this crate nor the model
/// rewrites.
#[derive(Clone, Debug)]
pub struct UnpatchableReference {
    part: String,
    /// the areas it names, or `None` when it did not resolve whole and every
    /// cell of every sheet has to be assumed named.
    areas: Option<Vec<ReferenceArea>>,
}

/// One area a part names, kept as its far corner: an insert or delete at or
/// before that corner moves what the area covers, and one past it does not.
#[derive(Clone, Debug)]
struct ReferenceArea {
    sheet: String,
    end: CellRef,
}

impl UnpatchableReference {
    /// The package path of the part.
    pub fn part(&self) -> &str {
        &self.part
    }

    /// Whether renaming, dropping or reordering `sheet` would strand it.
    pub fn names_sheet(&self, sheet: &str) -> bool {
        self.strands(|area| area.sheet.eq_ignore_ascii_case(sheet))
    }

    /// Whether inserting or deleting rows at `at` on `sheet` would move a cell
    /// it names. Everything from `at` down moves, so an area ending above it is
    /// left where it was.
    pub fn moved_by_rows(&self, sheet: &str, at: u32) -> bool {
        self.strands(|area| area.sheet.eq_ignore_ascii_case(sheet) && area.end.row >= at)
    }

    /// The column-wise counterpart of [`Self::moved_by_rows`].
    pub fn moved_by_cols(&self, sheet: &str, at: u32) -> bool {
        self.strands(|area| area.sheet.eq_ignore_ascii_case(sheet) && area.end.col >= at)
    }

    fn strands(&self, disturbed: impl Fn(&ReferenceArea) -> bool) -> bool {
        match &self.areas {
            None => true,
            Some(areas) => areas.iter().any(disturbed),
        }
    }
}

/// An area before it is bound to the workbook: a sheet name as the part spells
/// it, and the far corner of what it covers there.
type NamedArea = (String, CellRef);

/// Every preserved part a save cannot rewrite, with what each one names.
/// `sheet_paths` is the source part of each of `workbook`'s sheets, in order.
/// `declined` are the drawing and chart parts the parse could not read, which
/// only it knows about.
pub(crate) fn unpatchable_references(
    parts: &[(String, Vec<u8>)],
    content_types: &[PartContentType],
    workbook: &Workbook,
    sheet_paths: &[String],
    declined: &[String],
) -> Result<Vec<UnpatchableReference>, ParseError> {
    let mut references = if package_metadata_conforms(parts, content_types, sheet_paths) {
        let mut references = pivot_references(parts, content_types, workbook, sheet_paths);
        for chart in unmodelled_chart_parts(parts, content_types, workbook)? {
            let areas = match chart
                .claimed
                .then(|| find_part(parts, &chart.path))
                .flatten()
            {
                Some(bytes) => chart_reference_areas(bytes, chart.owner.as_deref())?,
                None => None,
            };
            references.push(bound(chart.path, areas, workbook));
        }
        references
    } else if package_bears_references(parts) {
        // Discovery must not be a function of a reader this path has just
        // decided it cannot trust, so here it stops being a function of any
        // reader: every part names everything. Nothing can drop out of a set
        // that is the whole package, whatever a future reader does with it.
        parts
            .iter()
            .map(|(path, _)| bound(path.clone(), None, workbook))
            .collect()
    } else {
        // A package holding nothing reference-bearing has nothing to strand,
        // and gets the veto it always got: none.
        Vec::new()
    };
    // A part the parse declined names cells nothing here resolved, so it names
    // every cell of every sheet, the way a chart part no sheet anchors does.
    // Out here rather than in the branch above because a decline vetoes even in
    // a package that bears no other reference, and because a walk over the parts
    // cannot see a target that is missing or unconventionally placed.
    for path in declined {
        match references
            .iter_mut()
            .find(|reference| reference.part.as_str() == path)
        {
            Some(reference) => reference.areas = None,
            None => references.push(bound(path.clone(), None, workbook)),
        }
    }
    Ok(references)
}

/// Whether the package holds anything a save could strand. Read as loosely as
/// it can be: a conventional directory, or a content type named by any
/// attribute anywhere in the typing part. Over-reading a trigger only ever
/// vetoes more, and reading it this loosely is what stops a decoy attribute
/// from hiding a real part behind a reader that takes the first match.
fn package_bears_references(parts: &[(String, Vec<u8>)]) -> bool {
    const DIRECTORIES: [&str; 3] = ["xl/pivottables/", "xl/pivotcache/", "xl/charts/"];
    if parts.iter().any(|(path, _)| {
        let key = part_key(path);
        DIRECTORIES.iter().any(|prefix| key.starts_with(prefix))
    }) {
        return true;
    }
    let Some(bytes) = find_part(parts, "[Content_Types].xml") else {
        return false;
    };
    parse_tree(bytes).is_ok_and(|root| names_a_reference_bearing_type(&root, 0))
}

fn names_a_reference_bearing_type(element: &Element, depth: usize) -> bool {
    depth <= MAX_DEPTH
        && (element.attributes.iter().any(|attribute| {
            PIVOT_CONTENT_TYPES
                .iter()
                .chain(CHART_CONTENT_TYPES.iter())
                .any(|known| attribute.value.eq_ignore_ascii_case(known))
        }) || element
            .child_elements()
            .any(|child| names_a_reference_bearing_type(child, depth + 1)))
}

/// Binds resolved areas to the workbook. A sheet it does not hold exactly one
/// of is one this crate cannot reason about — a reference into it would rebind
/// the moment some other sheet took that name — so the part is left naming
/// everything.
fn bound(part: String, areas: Option<Vec<NamedArea>>, workbook: &Workbook) -> UnpatchableReference {
    let areas = areas
        .filter(|areas| {
            areas
                .iter()
                .all(|(sheet, _)| names_one_sheet(workbook, sheet))
        })
        .map(|areas| {
            areas
                .into_iter()
                .map(|(sheet, end)| ReferenceArea { sheet, end })
                .collect()
        });
    UnpatchableReference { part, areas }
}

fn names_one_sheet(workbook: &Workbook, sheet: &str) -> bool {
    workbook
        .sheets
        .iter()
        .filter(|candidate| candidate.name.eq_ignore_ascii_case(sheet))
        .count()
        == 1
}

/// The directories a pivot part conventionally lives in, which answer for a
/// package that types its parts by extension alone.
const PIVOT_DIRECTORIES: [&str; 2] = ["xl/pivottables/", "xl/pivotcache/"];

/// content types that carry workbook references in a pivot vocabulary. A save
/// rewrites neither the ranges a cache is built from nor the grid a pivot table
/// is laid out on.
const PIVOT_CONTENT_TYPES: [&str; 3] = [
    "application/vnd.openxmlformats-officedocument.spreadsheetml.pivotTable+xml",
    "application/vnd.openxmlformats-officedocument.spreadsheetml.pivotCacheDefinition+xml",
    "application/vnd.openxmlformats-officedocument.spreadsheetml.pivotCacheRecords+xml",
];

fn pivot_references(
    parts: &[(String, Vec<u8>)],
    content_types: &[PartContentType],
    workbook: &Workbook,
    sheet_paths: &[String],
) -> Vec<UnpatchableReference> {
    let hosts = pivot_table_hosts(parts, workbook, sheet_paths);
    let pivot_parts = parts
        .iter()
        .filter(|(path, _)| is_pivot_part(path, content_types))
        .map(|(path, bytes)| (path, bytes, root_local_name(bytes)))
        .collect::<Vec<_>>();

    let mut cache_areas: HashMap<String, Option<Vec<NamedArea>>> = HashMap::new();
    let mut record_owners: HashMap<String, Option<String>> = HashMap::new();
    for (path, bytes) in pivot_parts
        .iter()
        .filter(|(_, _, root)| root.as_deref() == Some("pivotCacheDefinition"))
        .map(|(path, bytes, _)| (path, bytes))
    {
        let root = tree(bytes);
        cache_areas.insert(part_key(path), root.as_ref().and_then(source_areas));
        if let Some(records) = root
            .as_ref()
            .and_then(|root| root.attribute_ns(NS_RELATIONSHIPS, "id"))
            .and_then(|id| related_part(parts, path, id, &REL_PIVOT_CACHE_RECORDS))
        {
            record_owners
                .entry(records)
                .and_modify(|owner| *owner = None)
                .or_insert_with(|| Some(part_key(path)));
        }
    }

    let mut references = Vec::new();
    for (path, bytes, root) in pivot_parts {
        let areas = match root.as_deref() {
            Some("pivotCacheDefinition") => cache_areas.get(&part_key(path)).cloned().flatten(),
            Some("pivotCacheRecords") => record_owners
                .get(&part_key(path))
                .and_then(Option::as_ref)
                .and_then(|owner| cache_areas.get(owner))
                .cloned()
                .flatten(),
            Some("pivotTableDefinition") => tree(bytes)
                .as_ref()
                .and_then(|root| location_areas(root, hosts.get(&part_key(path)))),
            _ => None,
        };
        references.push(bound(path.clone(), areas, workbook));
    }
    references
}

/// The local name of a part's root element when the root is one of ours and the
/// part parses whole. Streamed rather than built into a tree, so a cached-
/// records part running to megabytes costs one pass and no memory — but the
/// whole pass, because the first tag says nothing about whether the rest of the
/// document closes. A foreign root, a root behind a prefix the part never
/// declared, an element left open, or anything beside the root answers to no
/// name, which leaves the part naming everything.
fn root_local_name(bytes: &[u8]) -> Option<String> {
    let mut reader = NsReader::from_reader(bytes);
    let config = reader.config_mut();
    config.expand_empty_elements = true;
    config.check_end_names = true;
    let mut buf = Vec::new();
    let mut depth = 0usize;
    let mut root = None;
    loop {
        let (namespace, event) = reader.read_resolved_event_into(&mut buf).ok()?;
        match event {
            Event::Start(start) => {
                depth += 1;
                if depth > MAX_DEPTH {
                    return None;
                }
                let cleared = declares_empty_default(&start)?;
                if depth == 1 {
                    if root.is_some() {
                        return None;
                    }
                    root = Some(ours_local_name(namespace, &start, cleared)?);
                }
            }
            Event::End(_) => depth = depth.checked_sub(1)?,
            Event::Text(text) if depth == 0 => {
                if !text.iter().all(u8::is_ascii_whitespace) {
                    return None;
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

/// The local name of a start tag whose expanded name is in our vocabulary,
/// decided by the same rule the tree readers use. `NsReader` reports a default
/// declared away as `Unbound`, exactly as it reports a part that declared none,
/// so the tag's own `xmlns` is what tells those two apart here.
fn ours_local_name(
    namespace: ResolveResult<'_>,
    start: &BytesStart<'_>,
    cleared: bool,
) -> Option<String> {
    let resolved = String::from_utf8_lossy(match namespace {
        ResolveResult::Bound(ref namespace) => namespace.as_ref(),
        _ => b"",
    })
    .into_owned();
    let vocabulary = match namespace {
        ResolveResult::Bound(_) => Vocabulary::Bound(&resolved),
        ResolveResult::Unbound if cleared => Vocabulary::Cleared,
        ResolveResult::Unbound | ResolveResult::Unknown(_) => Vocabulary::Absent,
    };
    let qname = String::from_utf8(start.name().as_ref().to_vec()).ok()?;
    owned_local_name(&qname, vocabulary, &OURS, Unqualified::Owned).map(str::to_owned)
}

/// Whether a start tag carries `xmlns=""`, which takes it out of every
/// vocabulary rather than leaving it in none by omission. `None` when its
/// attributes do not parse — quick-xml reads those lazily, so a tag only proves
/// itself well formed once they are walked.
fn declares_empty_default(start: &BytesStart<'_>) -> Option<bool> {
    let mut cleared = false;
    for attribute in start.attributes() {
        let attribute = attribute.ok()?;
        cleared |= attribute.key.as_ref() == b"xmlns" && attribute.value.as_ref().is_empty();
    }
    Some(cleared)
}

/// A pivot part as an element tree, or nothing when this crate declines to read
/// it whole: too large or too malformed, or carrying a markup-compatibility
/// choice it does not make. Declining is treated as naming every cell rather
/// than failing the open of a workbook that used to open.
fn tree(bytes: &[u8]) -> Option<Element> {
    parse_tree(bytes)
        .ok()
        .filter(|root| !carries_alternate_content(root))
}

/// The worksheet ranges a pivot cache is built from. Only a source read whole
/// resolves: an unhandled `type`, a child this crate does not model — an
/// `extLst` carrying the real source among them — or a `rangeSet` it cannot
/// read all of leaves the cache naming everything.
fn source_areas(root: &Element) -> Option<Vec<NamedArea>> {
    let source = sole_child(root, "cacheSource")?;
    let children = source.child_elements().collect::<Vec<_>>();
    let [only] = children[..] else {
        return None;
    };
    match sole_attribute(source, "type")?.unwrap_or("worksheet") {
        "worksheet" if only.answers_to(&OURS, "worksheetSource") => Some(vec![named_area(only)?]),
        "consolidation" if only.answers_to(&OURS, "consolidation") => consolidated_areas(only),
        _ => None,
    }
}

fn consolidated_areas(consolidation: &Element) -> Option<Vec<NamedArea>> {
    if consolidation
        .child_elements()
        .any(|child| !child.answers_to(&OURS, "pages") && !child.answers_to(&OURS, "rangeSets"))
    {
        return None;
    }
    let areas = sole_child(consolidation, "rangeSets")?
        .child_elements()
        .map(|set| {
            set.answers_to(&OURS, "rangeSet")
                .then(|| named_area(set))
                .flatten()
        })
        .collect::<Option<Vec<_>>>()?;
    (!areas.is_empty()).then_some(areas)
}

/// A `sheet`/`ref` pair. Refused when anything else could be naming the source
/// instead: the schema allows exactly one of `ref` and `name`, and an `r:id`
/// puts the range in another workbook.
fn named_area(element: &Element) -> Option<NamedArea> {
    if element.any_attribute_named("name") || element.any_attribute_named("id") {
        return None;
    }
    let sheet = sole_attribute(element, "sheet")??.to_owned();
    let (_, end) = parse_area(sole_attribute(element, "ref")??)?;
    Some((sheet, end))
}

/// The grid a pivot table occupies on each sheet hosting it. `ref` covers the
/// body; the filter area sits beside it, so its page counts are padded onto the
/// far corner rather than trusting `ref` to be the whole block.
fn location_areas(root: &Element, hosts: Option<&Vec<String>>) -> Option<Vec<NamedArea>> {
    let location = sole_child(root, "location")?;
    let (_, end) = parse_area(sole_attribute(location, "ref")??)?;
    let end = CellRef::new(
        end.row
            .saturating_add(page_count(location, "rowPageCount")?)
            .min(MAX_ROWS - 1),
        end.col
            .saturating_add(page_count(location, "colPageCount")?)
            .min(MAX_COLS - 1),
    );
    let hosts = hosts?;
    (!hosts.is_empty()).then(|| hosts.iter().map(|host| (host.clone(), end)).collect())
}

/// How far the filter area reaches past `ref`. Absent is none; anything that is
/// not a count leaves the pivot table unresolved.
fn page_count(location: &Element, name: &str) -> Option<u32> {
    match sole_attribute(location, name)? {
        None => Some(0),
        Some(value) => value.trim().parse().ok(),
    }
}

/// An attribute read by a name only one of ours may answer to. The outer `None`
/// is ambiguity — several carry that name, so a lookup would be choosing by
/// authoring order — and the inner `None` is absence, which an optional
/// attribute may default. A foreign attribute is not one of ours and so reads
/// as absent, exactly as it does to a consumer that drops it.
fn sole_attribute<'a>(element: &'a Element, name: &'a str) -> Option<Option<&'a str>> {
    let mut candidates = element.attributes_in(&OURS, name);
    let Some(only) = candidates.next() else {
        return Some(None);
    };
    candidates
        .next()
        .is_none()
        .then_some(Some(only.value.as_str()))
}

/// The one child answering to a name, or `None` when none of ours does or
/// several do. A name only a foreign node answers to is a name nothing answers
/// to, which for a node the part cannot be read without is unresolved.
fn sole_child<'a>(element: &'a Element, local: &'a str) -> Option<&'a Element> {
    let mut candidates = element.children_in(&OURS, local);
    let only = candidates.next()?;
    candidates.next().is_none().then_some(only)
}

/// Whether a part hands a consumer a choice this crate does not make. Choosing
/// between `mc:Choice` and `mc:Fallback` is markup-compatibility processing,
/// and a branch can supply the very node a lookup is after, so a pivot part
/// carrying one is read as naming everything rather than as whichever branch
/// happens to sit first. Matched by local name alone, because a branch behind a
/// prefix the part never declared still hides whatever it wraps.
fn carries_alternate_content(element: &Element) -> bool {
    element.local_name() == "AlternateContent"
        || element.child_elements().any(carries_alternate_content)
}

/// The sheets each pivot table part is laid out on, followed from the worksheet
/// relationships that anchor it. A part two sheets anchor is laid out on both.
fn pivot_table_hosts(
    parts: &[(String, Vec<u8>)],
    workbook: &Workbook,
    sheet_paths: &[String],
) -> HashMap<String, Vec<String>> {
    let mut hosts: HashMap<String, Vec<String>> = HashMap::new();
    for (sheet, path) in workbook.sheets.iter().zip(sheet_paths) {
        for table in related_parts(parts, path, &REL_PIVOT_TABLE) {
            hosts.entry(table).or_default().push(sheet.name.clone());
        }
    }
    hosts
}

/// The part a relationship id on `path` points at, under an exact type. Only
/// reached once the package metadata has been found conforming, so the shared
/// reader's loose matching has nothing left to be fooled by and the type can be
/// required exactly rather than by its final segment.
fn related_part(
    parts: &[(String, Vec<u8>)],
    path: &str,
    id: &str,
    types: &[&str],
) -> Option<String> {
    related(parts, path)
        .into_iter()
        .find(|(rel_id, kind, _)| rel_id == id && types.contains(&kind.as_str()))
        .map(|(_, _, target)| target)
}

/// Every part `path` relates to under an exact relationship type.
fn related_parts(parts: &[(String, Vec<u8>)], path: &str, types: &[&str]) -> Vec<String> {
    related(parts, path)
        .into_iter()
        .filter(|(_, kind, _)| types.contains(&kind.as_str()))
        .map(|(_, _, target)| target)
        .collect()
}

/// The relationships declared for `path`, with targets resolved to lookup keys.
fn related(parts: &[(String, Vec<u8>)], path: &str) -> Vec<(String, String, String)> {
    let Some(rels) = find_part(parts, &relationship_part_path(path)) else {
        return Vec::new();
    };
    let directory = directory_of(path).to_owned();
    parse_relationships(rels)
        .unwrap_or_default()
        .into_iter()
        .map(|(id, kind, target)| (id, kind, part_key(&resolve_part_path(&directory, &target))))
        .collect()
}

/// Whether the package metadata this path reads is exactly what OPC specifies.
/// Narrowing the veto means trusting readers that match names loosely, and this
/// crate cannot reconstruct that trust one hole at a time. So it asks once, of
/// the whole package: does the content-type part conform, and does every
/// relationships part this path follows conform? A no returns the workbook to
/// the blanket veto that shipped before narrowing existed — safe by
/// construction, because it is the behaviour already in people's hands.
fn package_metadata_conforms(
    parts: &[(String, Vec<u8>)],
    content_types: &[PartContentType],
    sheet_paths: &[String],
) -> bool {
    content_types_conform(parts)
        && narrowing_relationship_owners(parts, content_types, sheet_paths)
            .iter()
            .all(|owner| relationships_conform(parts, owner))
        && narrowing_drawings(parts, sheet_paths)
            .iter()
            .filter_map(|drawing| find_part(parts, drawing))
            .all(|bytes| {
                parse_tree(bytes).is_ok_and(|root| {
                    names_are_resolvable(&root) && drawing_claims_are_unambiguous(&root)
                })
            })
        && sheet_relationships_are_unambiguous(parts)
}

/// Every part whose relationships the narrowing decision reads, transitively.
/// Stated in one place, with why each is here, so a change that makes this path
/// follow something new shows up as a change to this list rather than as a hole
/// nobody notices. A part whose relationships are absent is not in doubt; one
/// whose relationships are read is only as good as they are.
fn narrowing_relationship_owners(
    parts: &[(String, Vec<u8>)],
    content_types: &[PartContentType],
    sheet_paths: &[String],
) -> Vec<String> {
    // the workbook, whose relationships resolve every `<sheet r:id>` to its
    // part, and so decide which sheet hosts a pivot table and which sheet
    // claims a chart
    let mut owners = vec!["xl/workbook.xml".to_owned()];
    for sheet in sheet_paths {
        // a sheet's relationships anchor the pivot tables laid out on it and
        // the drawings that claim its charts
        owners.push(sheet.clone());
    }
    // a drawing's relationships name the chart parts its sheet claims, which is
    // what makes a chart modelled and gives an unqualified reference its owner
    owners.extend(narrowing_drawings(parts, sheet_paths));
    // a cache definition's relationships name the records part that inherits
    // its source range
    owners.extend(
        parts
            .iter()
            .map(|(path, _)| path.clone())
            .filter(|path| is_pivot_part(path, content_types)),
    );
    owners
}

/// The drawing parts the narrowing decision reads, whose own markup decides
/// which sheet claims which chart.
fn narrowing_drawings(parts: &[(String, Vec<u8>)], sheet_paths: &[String]) -> Vec<String> {
    sheet_paths
        .iter()
        .flat_map(|sheet| related_parts(parts, sheet, &REL_DRAWING))
        .collect()
}

/// Whether every attribute a reader could take for `name` is the one this path
/// would read: at most one carries that local name, and it is ours.
fn sole_and_ours(element: &Element, namespaces: &[&str], name: &str) -> bool {
    element.attributes_named(name) <= 1
        && element.attributes_in(namespaces, name).count() == element.attributes_named(name)
}

/// Whether `[Content_Types].xml` holds only what OPC allows: a `Types` root
/// over `Default` and `Override` entries, each carrying its required names once
/// and unambiguously, no extension or part name typed twice.
fn content_types_conform(parts: &[(String, Vec<u8>)]) -> bool {
    const NAMES: [&str; 3] = ["PartName", "ContentType", "Extension"];
    let Some(bytes) = find_part(parts, "[Content_Types].xml") else {
        return false;
    };
    let Ok(root) = parse_tree(bytes) else {
        return false;
    };
    if !root.answers_to(&CONTENT_TYPES, "Types") {
        return false;
    }
    let (mut overrides, mut defaults) = (HashSet::new(), HashSet::new());
    root.child_elements().all(|entry| {
        entry.child_elements().next().is_none()
            && NAMES
                .iter()
                .all(|name| sole_and_ours(entry, &CONTENT_TYPES, name))
            && entry
                .attribute_local("ContentType")
                .is_some_and(|kind| !kind.trim().is_empty())
            && if entry.answers_to(&CONTENT_TYPES, "Override") {
                entry
                    .attribute_local("PartName")
                    .is_some_and(|part| part.len() > 1 && part.starts_with('/'))
                    .then(|| entry.attribute_local("PartName"))
                    .flatten()
                    .is_some_and(|part| overrides.insert(part_key(part)))
            } else if entry.answers_to(&CONTENT_TYPES, "Default") {
                entry
                    .attribute_local("Extension")
                    .filter(|extension| !extension.is_empty() && !extension.contains('/'))
                    .is_some_and(|extension| defaults.insert(extension.to_ascii_lowercase()))
            } else {
                false
            }
    })
}

/// Whether the relationships declared for a part hold only what OPC allows: a
/// `Relationships` root over `Relationship` entries, each carrying its required
/// names once and unambiguously, no id declared twice.
fn relationships_conform(parts: &[(String, Vec<u8>)], owner: &str) -> bool {
    const NAMES: [&str; 4] = ["Id", "Type", "Target", "TargetMode"];
    let Some(bytes) = find_part(parts, &relationship_part_path(owner)) else {
        return true;
    };
    let Ok(root) = parse_tree(bytes) else {
        return false;
    };
    if root.namespace() != Some(NS_PACKAGE_RELATIONSHIPS) || root.local_name() != "Relationships" {
        return false;
    }
    let mut ids = HashSet::new();
    root.child_elements().all(|entry| {
        entry.namespace() == Some(NS_PACKAGE_RELATIONSHIPS)
            && entry.local_name() == "Relationship"
            && entry.child_elements().next().is_none()
            && NAMES
                .iter()
                .all(|name| sole_and_ours(entry, &RELATIONSHIPS, name))
            && entry
                .attribute_local("Type")
                .is_some_and(followed_type_is_exact)
            && ["Type", "Target"].iter().all(|name| {
                entry
                    .attribute_local(name)
                    .is_some_and(|value| !value.is_empty())
            })
            && entry
                .attribute_local("TargetMode")
                .is_none_or(|mode| matches!(mode, "Internal" | "External"))
            && entry
                .attribute_local("Id")
                .filter(|id| !id.is_empty())
                .is_some_and(|id| ids.insert(id.to_owned()))
    })
}

/// Whether each `<sheet>` in the workbook names its part unambiguously. The
/// entry is read by taking the first attribute whose local name is `id`, so a
/// foreign one before the real `r:id` swaps which part a sheet resolves to —
/// leaving every sheet claimed while a pivot table is attributed to the wrong
/// one, which is a narrow reading of a workbook nobody can read that way.
fn sheet_relationships_are_unambiguous(parts: &[(String, Vec<u8>)]) -> bool {
    let Some(bytes) = find_part(parts, "xl/workbook.xml") else {
        return false;
    };
    let Ok(root) = parse_tree(bytes) else {
        return false;
    };
    let Some(sheets) = sole_child(&root, "sheets") else {
        return false;
    };
    let canonical = sheets
        .child_elements()
        .filter(|sheet| sheet.local_name() == "sheet")
        .collect::<Vec<_>>();
    // The model is built by a streaming reader that takes every `sheet` by
    // local name wherever it sits, while the sheet-to-part mapping records only
    // these children. One more anywhere else shifts the two out of step, and a
    // pivot table is then attributed to whichever sheet moved into its slot.
    if named_sheet_elements(&root, 0) != canonical.len() {
        return false;
    }
    canonical.iter().all(|sheet| {
        sheet.answers_to(&OURS, "sheet")
            && sheet.attributes_named("id") == 1
            && sheet
                .attributes_in(&[NS_RELATIONSHIPS], "id")
                .next()
                .is_some()
            && sole_and_ours(sheet, &OURS, "name")
            && sole_and_ours(sheet, &OURS, "sheetId")
    })
}

/// How many elements the streaming model reader would take for a sheet: local
/// name `sheet`, at any depth.
fn named_sheet_elements(element: &Element, depth: usize) -> usize {
    if depth > MAX_DEPTH {
        return usize::MAX;
    }
    usize::from(element.local_name() == "sheet")
        + element
            .child_elements()
            .map(|child| named_sheet_elements(child, depth + 1))
            .sum::<usize>()
}

/// Whether a relationship type is one the shared readers follow. Those match
/// on the segment after the last slash, so a type merely ending in one of them
/// steers chart ownership or a pivot claim without being the standard type this
/// path enumerated. Such a package is not one to read relationships from.
fn followed_type_is_exact(kind: &str) -> bool {
    const FOLLOWED: [(&str, &[&str]); 4] = [
        ("drawing", &REL_DRAWING),
        ("chart", &REL_CHART),
        ("pivotTable", &REL_PIVOT_TABLE),
        ("pivotCacheRecords", &REL_PIVOT_CACHE_RECORDS),
    ];
    FOLLOWED.iter().all(|(segment, standard)| {
        kind.rsplit('/').next() != Some(segment) || standard.contains(&kind)
    })
}

/// Whether a part holds a pivot cache or pivot table.
///
/// Discovery is monotonic over the blanket veto: a part the conventional
/// directories name is always looked at, and conforming metadata may only ADD
/// parts beside them. Never a choice between the two — typing this path cannot
/// read must not be able to take a part away, because a part nobody looks at
/// vetoes nothing, which is the one outcome worse than a spurious refusal.
fn is_pivot_part(path: &str, content_types: &[PartContentType]) -> bool {
    let key = part_key(path);
    // A `.rels` part names parts, never cells. The blanket scan swept the ones
    // under these directories in, but only ever to veto a workbook that holds
    // the pivot part they describe — which is itself conventional, so leaving
    // them out here cannot lose a veto. An explicitly typed part is a pivot
    // part wherever it sits.
    (!key.contains("/_rels/")
        && PIVOT_DIRECTORIES
            .iter()
            .any(|prefix| key.starts_with(prefix)))
        || content_types.iter().any(|part| {
            part.path == key
                && PIVOT_CONTENT_TYPES
                    .iter()
                    .any(|known| part.content_type.eq_ignore_ascii_case(known))
        })
}

fn part_key(path: &str) -> String {
    path.trim_start_matches('/').to_ascii_lowercase()
}
