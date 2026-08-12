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
    NS_PACKAGE_RELATIONSHIPS, NS_RELATIONSHIPS, chart_reference_areas, directory_of, parse_area,
    parse_relationships, relationship_part_path, unmodelled_chart_parts,
};
use crate::package::PartContentType;
use crate::tree::{Element, Unqualified, Vocabulary, owned_local_name, parse_tree};
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
    let typing = content_type_typing(parts);
    let hosts = pivot_table_hosts(parts, workbook, sheet_paths);
    let pivot_parts = parts
        .iter()
        .filter(|(path, _)| is_pivot_part(path, content_types, &typing))
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
        let areas = matches!(typing, Typing::Trusted).then_some(areas).flatten();
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
        for table in related_parts(parts, path, "pivotTable") {
            hosts.entry(table).or_default().push(sheet.name.clone());
        }
    }
    hosts
}

/// The relationships declared for `path`, when they are ones this path may be
/// read from. [`parse_relationships`] is shared code that matches `Id`, `Type`
/// and `Target` by local name and takes the first of each, so a foreign
/// attribute could answer for a real one and a repeated id could shadow the
/// relationship it names. Rather than reach into that reader, the pivot path
/// declines the whole part when it holds either, which costs a refusal and
/// never a wrong claim.
/// Whether every attribute a reader could take for `name` is the one this path
/// would read: at most one carries that local name, and it is ours. Counting
/// only the ours-namespaced ones would pass a foreign attribute sitting beside
/// a real one, which is the one the reader takes first.
fn sole_and_ours(element: &Element, namespaces: &[&str], name: &str) -> bool {
    element.attributes_named(name) <= 1
        && element.attributes_in(namespaces, name).count() == element.attributes_named(name)
}

fn trusted_relationships<'a>(parts: &'a [(String, Vec<u8>)], path: &str) -> Option<&'a [u8]> {
    const NAMES: [&str; 4] = ["Id", "Type", "Target", "TargetMode"];
    let rels = find_part(parts, &relationship_part_path(path))?;
    let root = parse_tree(rels).ok()?;
    let mut ids = HashSet::new();
    for child in root
        .child_elements()
        .filter(|child| child.local_name() == "Relationship")
    {
        if !child.answers_to(&RELATIONSHIPS, "Relationship") {
            return None;
        }
        for name in NAMES {
            if !sole_and_ours(child, &RELATIONSHIPS, name) {
                return None;
            }
        }
        if let Some(id) = child.attributes_in(&RELATIONSHIPS, "Id").next()
            && !ids.insert(id.value.as_str())
        {
            return None;
        }
    }
    Some(rels)
}

/// The part a relationship id on `path` points at, by lookup key.
fn related_part(parts: &[(String, Vec<u8>)], path: &str, id: &str) -> Option<String> {
    let rels = trusted_relationships(parts, path)?;
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
    let Some(rels) = trusted_relationships(parts, path) else {
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

/// How far `[Content_Types].xml` may be trusted to type a pivot part.
enum Typing {
    /// every entry reads unambiguously, so an override may rule a part in or
    /// out the way OPC says it does
    Trusted,
    /// something in it cannot be read as ours, so an override may only add a
    /// part to look at, never take one away, and nothing it typed resolves
    Untrusted(HashSet<String>),
}

/// How far the content types may be trusted, and failing that, every part any
/// override could have been naming as a pivot. OPC typing is shared code that
/// matches `Override`, `PartName` and `ContentType` by local name at every
/// depth and takes the first of each, so a foreign one could answer for a real
/// one. This gate walks what that reader walks — the whole tree, root included
/// — rather than an approximation of it.
fn content_type_typing(parts: &[(String, Vec<u8>)]) -> Typing {
    const NAMES: [&str; 3] = ["PartName", "ContentType", "Extension"];
    let Some(bytes) = find_part(parts, "[Content_Types].xml") else {
        return Typing::Trusted;
    };
    let Ok(root) = parse_tree(bytes) else {
        return Typing::Untrusted(HashSet::new());
    };
    let mut entries = Vec::new();
    collect_content_type_entries(&root, 0, &mut entries);
    let mut typed = HashSet::new();
    let trusted = entries.iter().all(|(depth, entry)| {
        *depth == 1
            && (entry.answers_to(&CONTENT_TYPES, "Override")
                || entry.answers_to(&CONTENT_TYPES, "Default"))
            && NAMES
                .iter()
                .all(|name| sole_and_ours(entry, &CONTENT_TYPES, name))
    }) && entries
        .iter()
        .filter(|(_, entry)| entry.local_name() == "Override")
        .all(|(_, entry)| match entry.attribute_local("PartName") {
            Some(part) => typed.insert(part_key(part)),
            None => true,
        });
    if trusted {
        return Typing::Trusted;
    }
    Typing::Untrusted(pivot_candidates(&entries))
}

/// Every element the OPC reader would take for an entry, with its depth:
/// matched by local name at any depth, the root included, exactly as
/// [`crate::package`] reads them. The depth comes back because that reader
/// honours an entry the schema only allows directly under `Types`, and one
/// sitting anywhere else is a reading no conforming consumer shares.
fn collect_content_type_entries<'a>(
    element: &'a Element,
    depth: usize,
    out: &mut Vec<(usize, &'a Element)>,
) {
    if matches!(element.local_name(), "Override" | "Default") {
        out.push((depth, element));
    }
    for child in element.child_elements() {
        collect_content_type_entries(child, depth + 1, out);
    }
}

/// Every part an override could have been typing as a pivot, however the entry
/// spells its names. Used only when the typing is not trustworthy, where the
/// answer has to be inclusive: a part left undiscovered vetoes nothing.
fn pivot_candidates(entries: &[(usize, &Element)]) -> HashSet<String> {
    entries
        .iter()
        .map(|(_, entry)| entry)
        .filter(|entry| entry.local_name() == "Override")
        .filter(|entry| {
            entry.attributes.iter().any(|attribute| {
                attribute.local_name() == "ContentType"
                    && PIVOT_CONTENT_TYPES
                        .iter()
                        .any(|known| attribute.value.eq_ignore_ascii_case(known))
            })
        })
        .flat_map(|entry| {
            entry
                .attributes
                .iter()
                .filter(|attribute| attribute.local_name() == "PartName")
                .map(|attribute| part_key(&attribute.value))
        })
        .collect()
}

/// Whether a part holds a pivot cache or pivot table. Typed the way OPC types
/// anything: an `Override` is authoritative, and the conventional directories
/// answer only for a part a `Default` extension mapping left untyped. A pivot
/// part written outside those directories is still found, because a veto that
/// reasons per part gives a part it never discovers no cover at all. Typing
/// this path cannot trust never rules a part out; it only stops ruling one in.
fn is_pivot_part(path: &str, content_types: &[PartContentType], typing: &Typing) -> bool {
    let key = part_key(path);
    if key.contains("/_rels/") {
        return false;
    }
    let conventional = || {
        PIVOT_DIRECTORIES
            .iter()
            .any(|prefix| key.starts_with(prefix))
    };
    let typed = match typing {
        Typing::Trusted => content_types.iter().find(|part| part.path == key),
        Typing::Untrusted(candidates) => {
            return conventional() || candidates.contains(&key);
        }
    };
    match typed {
        Some(part)
            if PIVOT_CONTENT_TYPES
                .iter()
                .any(|known| part.content_type.eq_ignore_ascii_case(known)) =>
        {
            true
        }
        Some(part) if part.overridden => false,
        _ => conventional(),
    }
}

fn part_key(path: &str) -> String {
    path.trim_start_matches('/').to_ascii_lowercase()
}
