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

use std::collections::HashMap;

use quick_xml::events::Event;
use xlsx_model::addr::{MAX_COLS, MAX_ROWS};
use xlsx_model::{CellRef, Workbook};

use crate::ParseError;
use crate::chart::{
    NS_RELATIONSHIPS, chart_reference_areas, directory_of, parse_area, parse_relationships,
    relationship_part_path, unmodelled_chart_parts,
};
use crate::package::PartContentType;
use crate::tree::{Element, parse_tree};
use crate::write::NS_MAIN;
use crate::xml::{find_part, local_name, next_event, reader, resolve_part_path};

/// the markup-compatibility vocabulary, whose branches stand in for the element
/// they wrap.
const NS_MCE: &str = "http://schemas.openxmlformats.org/markup-compatibility/2006";

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
pub(crate) fn unpatchable_references(
    parts: &[(String, Vec<u8>)],
    content_types: &[PartContentType],
    workbook: &Workbook,
    sheet_paths: &[String],
) -> Result<Vec<UnpatchableReference>, ParseError> {
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
    Ok(references)
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
    let mut record_owners: HashMap<String, String> = HashMap::new();
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
            .and_then(|id| related_part(parts, path, id))
        {
            record_owners.insert(records, part_key(path));
        }
    }

    let mut references = Vec::new();
    for (path, bytes, root) in pivot_parts {
        let areas = match root.as_deref() {
            Some("pivotCacheDefinition") => cache_areas.get(&part_key(path)).cloned().flatten(),
            Some("pivotCacheRecords") => record_owners
                .get(&part_key(path))
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

/// The local name of a part's root element, read without building a tree so a
/// cached-records part running to megabytes costs nothing to classify.
fn root_local_name(bytes: &[u8]) -> Option<String> {
    let mut reader = reader(bytes);
    let mut buf = Vec::new();
    let mut depth = 0;
    loop {
        match next_event(&mut reader, &mut buf, &mut depth).ok()? {
            Event::Start(start) => return String::from_utf8(local_name(&start)).ok(),
            Event::Eof => return None,
            _ => {}
        }
    }
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
        "worksheet" if only.answers_to(NS_MAIN, "worksheetSource") => Some(vec![named_area(only)?]),
        "consolidation" if only.answers_to(NS_MAIN, "consolidation") => consolidated_areas(only),
        _ => None,
    }
}

fn consolidated_areas(consolidation: &Element) -> Option<Vec<NamedArea>> {
    if consolidation
        .child_elements()
        .any(|child| !child.answers_to(NS_MAIN, "pages") && !child.answers_to(NS_MAIN, "rangeSets"))
    {
        return None;
    }
    let areas = sole_child(consolidation, "rangeSets")?
        .child_elements()
        .map(|set| {
            set.answers_to(NS_MAIN, "rangeSet")
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
    if element.attributes_in(NS_MAIN, "name").next().is_some()
        || element.attributes_in(NS_MAIN, "id").next().is_some()
    {
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
    let mut candidates = element.attributes_in(NS_MAIN, name);
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
    let mut candidates = element.children_in(NS_MAIN, local);
    let only = candidates.next()?;
    candidates.next().is_none().then_some(only)
}

/// Whether a part hands a consumer a choice this crate does not make. Choosing
/// between `mc:Choice` and `mc:Fallback` is markup-compatibility processing,
/// and a branch can supply the very node a lookup is after, so a pivot part
/// carrying one is read as naming everything rather than as whichever branch
/// happens to sit first.
fn carries_alternate_content(element: &Element) -> bool {
    element.is(NS_MCE, "AlternateContent")
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
        for table in related_parts(parts, path, "pivotTable") {
            hosts.entry(table).or_default().push(sheet.name.clone());
        }
    }
    hosts
}

/// The part a relationship id on `path` points at, by lookup key.
fn related_part(parts: &[(String, Vec<u8>)], path: &str, id: &str) -> Option<String> {
    let rels = find_part(parts, &relationship_part_path(path))?;
    let directory = directory_of(path).to_owned();
    parse_relationships(rels)
        .ok()?
        .iter()
        .find(|(rel_id, _, _)| rel_id == id)
        .map(|(_, _, target)| part_key(&resolve_part_path(&directory, target)))
}

/// Every part `path` relates to under a relationship type, by lookup key. A
/// relationships part that will not parse relates nothing, which leaves what
/// depended on it unresolved rather than failing the open.
fn related_parts(parts: &[(String, Vec<u8>)], path: &str, relationship: &str) -> Vec<String> {
    let Some(rels) = find_part(parts, &relationship_part_path(path)) else {
        return Vec::new();
    };
    let directory = directory_of(path).to_owned();
    parse_relationships(rels)
        .unwrap_or_default()
        .iter()
        .filter(|(_, kind, _)| kind.rsplit('/').next() == Some(relationship))
        .map(|(_, _, target)| part_key(&resolve_part_path(&directory, target)))
        .collect()
}

/// Whether a part holds a pivot cache or pivot table. Typed the way OPC types
/// anything: an `Override` is authoritative, and the conventional directories
/// answer only for a part a `Default` extension mapping left untyped. A pivot
/// part written outside those directories is still found, because a veto that
/// reasons per part gives a part it never discovers no cover at all.
fn is_pivot_part(path: &str, content_types: &[PartContentType]) -> bool {
    let key = part_key(path);
    if key.contains("/_rels/") {
        return false;
    }
    match content_types.iter().find(|part| part.path == key) {
        Some(part)
            if PIVOT_CONTENT_TYPES
                .iter()
                .any(|known| part.content_type.eq_ignore_ascii_case(known)) =>
        {
            true
        }
        Some(part) if part.overridden => false,
        _ => PIVOT_DIRECTORIES
            .iter()
            .any(|prefix| key.starts_with(prefix)),
    }
}

fn part_key(path: &str) -> String {
    path.trim_start_matches('/').to_ascii_lowercase()
}
