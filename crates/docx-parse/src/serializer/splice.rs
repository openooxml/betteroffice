//! Story parts written as their source XML with only the paragraphs a save changed re-serialized.

use std::borrow::Cow;
use std::collections::{BTreeMap, BTreeSet, HashMap};
use std::ops::Range;

use crate::paragraph_identity::{
    MC_NAMESPACE, ParagraphOccurrence, Tag, attribute, paragraph_occurrences, patch_paragraph_ids,
    tags, unescaped,
};

use super::context::RecordedParagraphs;
use super::paragraph_ids::{S13SpliceAnchor, S13SplicedPart};
use super::s13::element_span;

const W_NAMESPACE: &str = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

/// What the XML written for an unchanged paragraph must agree on with its source XML for the
/// source to stand in for it: the same comment markers in order, at the same text offsets, the
/// same permission markers, and no revision or note ID the source lacks. Source IDs the written XML lacks are kept, as are
/// relationship IDs: the save keeps every source relationship.
#[derive(Debug, Default, PartialEq, Eq)]
struct Marks {
    comments: Vec<(String, String, usize)>,
    permissions: Vec<(String, String)>,
    ids: BTreeMap<(&'static str, String), usize>,
}

impl Marks {
    fn admit(&self, written: &Marks) -> bool {
        self.comments == written.comments
            && self.permissions == written.permissions
            && written
                .ids
                .iter()
                .all(|(id, count)| self.ids.get(id).is_some_and(|source| source >= count))
    }
}

/// Run content that takes a position in the text, besides `w:t` and `w:delText` characters.
pub(crate) const POSITIONED: [&str; 19] = [
    "w:tab",
    "w:ptab",
    "w:br",
    "w:cr",
    "w:sym",
    "w:noBreakHyphen",
    "w:softHyphen",
    "w:drawing",
    "w:pict",
    "w:object",
    "w:fldChar",
    "w:footnoteReference",
    "w:endnoteReference",
    "w:footnoteRef",
    "w:endnoteRef",
    "w:separator",
    "w:continuationSeparator",
    "w:annotationRef",
    "m:oMath",
];

fn marks(xml: &str) -> Option<Marks> {
    let mut marks = Marks::default();
    let mut offset = 0;
    let mut in_text = 0usize;
    let mut in_properties = 0usize;
    let mut previous = 0;
    for tag in tags(xml)? {
        if in_text > 0 {
            offset += unescaped(xml, previous..tag.range.start)?.chars().count();
        }
        previous = tag.range.end;
        let text = matches!(tag.name, "w:t" | "w:delText");
        let properties = tag.name.ends_with("Pr");
        if tag.end {
            if text {
                in_text = in_text.saturating_sub(1);
            } else if properties {
                in_properties = in_properties.saturating_sub(1);
            }
            continue;
        }
        if properties {
            in_properties += usize::from(!tag.empty);
            continue;
        }
        if in_properties == 0 && POSITIONED.contains(&tag.name) {
            offset += 1;
        }
        let id = || {
            attribute(&tag, "w:id")
                .and_then(|range| unescaped(xml, range))
                .unwrap_or_default()
        };
        let kind = match tag.name {
            _ if text => {
                in_text += usize::from(!tag.empty);
                continue;
            }
            "w:commentRangeStart" | "w:commentRangeEnd" | "w:commentReference" => {
                marks.comments.push((tag.name.to_owned(), id(), offset));
                continue;
            }
            "w:permStart" | "w:permEnd" => {
                marks.permissions.push((tag.name.to_owned(), id()));
                continue;
            }
            "w:ins" | "w:del" | "w:moveFrom" | "w:moveTo" => "revision",
            name if name.ends_with("Change") => "revision",
            "w:footnoteReference" => "footnote",
            "w:endnoteReference" => "endnote",
            _ => continue,
        };
        *marks.ids.entry((kind, id())).or_insert(0) += 1;
    }
    marks.permissions.sort();
    Some(marks)
}

/// How many comment markers of each kind and ID `xml` holds.
fn comment_counts(xml: &str) -> Option<BTreeMap<(String, String), isize>> {
    let mut counts = BTreeMap::new();
    for (name, id, _) in marks(xml)?.comments {
        *counts.entry((name, id)).or_insert(0) += 1;
    }
    Some(counts)
}

fn subtract(
    total: &mut BTreeMap<(String, String), isize>,
    part: BTreeMap<(String, String), isize>,
) {
    for (key, count) in part {
        *total.entry(key).or_insert(0) -= count;
    }
    total.retain(|_, count| *count != 0);
}

/// The namespace prefixes `fragment` uses outside the scope of its own declarations of them.
pub(crate) fn required_prefixes(fragment: &str) -> Option<BTreeSet<&str>> {
    let mut required = BTreeSet::new();
    let mut scopes: Vec<Vec<&str>> = Vec::new();
    for tag in tags(fragment)? {
        if tag.end {
            scopes.pop();
            continue;
        }
        scopes.push(
            tag.attributes
                .iter()
                .filter_map(|(key, _)| key.strip_prefix("xmlns:"))
                .collect(),
        );
        let names = std::iter::once(tag.name).chain(
            tag.attributes
                .iter()
                .map(|(key, _)| *key)
                .filter(|key| *key != "xmlns" && !key.starts_with("xmlns:")),
        );
        for prefix in names.filter_map(|name| name.split_once(':').map(|(prefix, _)| prefix)) {
            if prefix != "xml" && !scopes.iter().any(|scope| scope.contains(&prefix)) {
                required.insert(prefix);
            }
        }
        if tag.empty {
            scopes.pop();
        }
    }
    Some(required)
}

/// The root element's start tag.
pub(crate) fn root_tag(xml: &str) -> Option<Tag<'_>> {
    let bytes = xml.as_bytes();
    let mut cursor = 0;
    loop {
        let start = cursor + xml[cursor..].find('<')?;
        let rest = &xml[start..];
        let terminator = if rest.starts_with("<?") {
            "?>"
        } else if rest.starts_with("<!--") {
            "-->"
        } else if rest.starts_with("<!") {
            ">"
        } else {
            let mut quote = None;
            for (offset, byte) in bytes[start + 1..].iter().enumerate() {
                match (quote, *byte) {
                    (None, b'"' | b'\'') => quote = Some(*byte),
                    (Some(current), byte) if current == byte => quote = None,
                    (None, b'>') => return tags(&xml[..start + offset + 2])?.pop(),
                    _ => {}
                }
            }
            return None;
        };
        cursor = start + rest.find(terminator)? + terminator.len();
    }
}

fn bindings<'a>(xml: &'a str, tag: &Tag<'a>) -> HashMap<&'a str, &'a str> {
    tag.attributes
        .iter()
        .filter_map(|(key, range)| Some((key.strip_prefix("xmlns:")?, &xml[range.clone()])))
        .collect()
}

pub(crate) fn root_with_prefixes<'s>(
    source: &'s str,
    serialized: &str,
    prefixes: &BTreeSet<&str>,
) -> Option<Cow<'s, str>> {
    let root = root_tag(source)?;
    let bound = bindings(source, &root);
    let target = root_tag(serialized)?;
    let wanted = bindings(serialized, &target);
    let mut declarations = String::new();
    for prefix in prefixes {
        if !bound.contains_key(prefix)
            && let Some(uri) = wanted.get(prefix)
        {
            declarations.push_str(&format!(" xmlns:{prefix}=\"{uri}\""));
        }
    }
    let mut edits = Vec::new();
    if prefixes.contains("w14") && !bound.contains_key("w14") {
        match bound.get("mc") {
            Some(uri) if uri.trim() != MC_NAMESPACE => return None,
            Some(_) => {}
            None => {
                if !prefixes.contains("mc") {
                    declarations.push_str(&format!(" xmlns:mc=\"{MC_NAMESPACE}\""));
                }
            }
        }
        match attribute(&root, "mc:Ignorable") {
            Some(range) => {
                let value = unescaped(source, range.clone())?;
                if !value.split_whitespace().any(|prefix| prefix == "w14") {
                    let suffix = if value.is_empty() { "w14" } else { " w14" };
                    edits.push((range.end..range.end, suffix.to_owned()));
                }
            }
            None => declarations.push_str(" mc:Ignorable=\"w14\""),
        }
    }
    if !declarations.is_empty() {
        let at = root.range.end - if root.empty { 2 } else { 1 };
        edits.push((at..at, declarations));
    }
    if edits.is_empty() {
        return Some(Cow::Borrowed(source));
    }
    edits.sort_by_key(|(range, _)| range.start);
    let mut patched = String::with_capacity(source.len());
    let mut cursor = 0;
    for (range, text) in edits {
        patched.push_str(&source[cursor..range.start]);
        patched.push_str(&text);
        cursor = range.end;
    }
    patched.push_str(&source[cursor..]);
    Some(Cow::Owned(patched))
}

/// Whether `source`'s root binds every prefix in `prefixes` to the namespace `serialized`'s
/// root binds it to, so that paragraphs written with them need no new declaration.
pub(crate) fn root_binds(
    source: &str,
    serialized: &str,
    prefixes: &BTreeSet<&str>,
    allow_missing: bool,
) -> Option<bool> {
    let root = root_tag(source)?;
    let bound = bindings(source, &root);
    let target = root_tag(serialized)?;
    let wanted = bindings(serialized, &target);
    Some(
        prefixes
            .iter()
            .all(|prefix| match (bound.get(prefix), wanted.get(prefix)) {
                (Some(existing), Some(uri)) => existing.trim() == uri.trim(),
                (None, _) => allow_missing,
                _ => false,
            }),
    )
}

/// Where the addressed paragraphs sit in a source part: each one's parent element and, for
/// the anchors of inserted paragraphs, the tags right before and right after it.
struct Layout<'s> {
    parents: HashMap<u32, &'s str>,
    edges: HashMap<u32, Edges<'s>>,
}

/// The tags around a paragraph, each as `(name, whether it is an end tag)`.
#[derive(Default)]
struct Edges<'s> {
    before: Option<(&'s str, bool)>,
    after: Option<(&'s str, bool)>,
    field_after: bool,
}

impl Edges<'_> {
    /// Whether a paragraph goes in at `anchor` exactly where the model has it: with only a
    /// paragraph, a table, the section or the edge of its story or cell on the far side.
    fn admit(&self, anchor: S13SpliceAnchor) -> bool {
        match anchor {
            S13SpliceAnchor::After(_) if self.field_after => false,
            S13SpliceAnchor::After(_) => match self.after {
                Some((name, true)) => PARENTS.contains(&name),
                Some((name, false)) => matches!(name, "w:p" | "w:tbl" | "w:sectPr"),
                None => false,
            },
            S13SpliceAnchor::Before(_) => match self.before {
                Some((name, false)) => PARENTS.contains(&name),
                Some((name, true)) => name == "w:tcPr",
                None => false,
            },
        }
    }
}

/// The [`Layout`] of the paragraphs at `spans` in `source`, with edges for those in `anchors`.
/// Each addressed paragraph's parent element is `w:fldChar` for one that starts
/// inside a complex field, whose code or result it belongs to.
/// `None` unless every namespace prefix in `source` keeps one binding and every namespace one
/// prefix, `w` is bound to WordprocessingML and no default namespace is declared, so that
/// element names identify elements, and every field character is a begin, separate or end
/// that closes an open field.
fn layout<'s>(
    source: &'s str,
    spans: &BTreeMap<u32, Range<usize>>,
    anchors: &BTreeSet<u32>,
) -> Option<Layout<'s>> {
    let starts: HashMap<usize, u32> = spans
        .iter()
        .map(|(ordinal, span)| (span.start, *ordinal))
        .collect();
    let mut ends: Vec<(usize, u32)> = anchors
        .iter()
        .map(|ordinal| Some((spans.get(ordinal)?.end, *ordinal)))
        .collect::<Option<_>>()?;
    ends.sort_unstable();
    let mut ends = ends.into_iter().peekable();
    let mut prefixes: HashMap<&str, &str> = HashMap::new();
    let mut namespaces: HashMap<&str, &str> = HashMap::new();
    let mut stack: Vec<&str> = Vec::new();
    let mut fields = 0usize;
    let mut layout = Layout {
        parents: HashMap::new(),
        edges: HashMap::new(),
    };
    let mut previous = None;
    for tag in tags(source)? {
        while let Some((_, ordinal)) = ends.next_if(|(end, _)| *end <= tag.range.start) {
            let edges = layout.edges.entry(ordinal).or_default();
            edges.after = Some((tag.name, tag.end));
            edges.field_after = fields > 0;
        }
        let last = previous.replace((tag.name, tag.end));
        if tag.end {
            stack.pop();
            continue;
        }
        if tag
            .name
            .rsplit_once(':')
            .map_or(tag.name, |(_, local)| local)
            == "fldChar"
        {
            match &*unescaped(source, attribute(&tag, "w:fldCharType")?)? {
                "begin" => fields += 1,
                "separate" => {}
                "end" => fields = fields.checked_sub(1)?,
                _ => return None,
            }
        }
        for (key, range) in &tag.attributes {
            if *key == "xmlns" {
                return None;
            }
            let Some(prefix) = key.strip_prefix("xmlns:") else {
                continue;
            };
            let uri = source[range.clone()].trim();
            if uri.contains('&')
                || *prefixes.entry(prefix).or_insert(uri) != uri
                || *namespaces.entry(uri).or_insert(prefix) != prefix
            {
                return None;
            }
        }
        if let Some(ordinal) = starts.get(&tag.range.start) {
            layout.parents.insert(
                *ordinal,
                if fields > 0 {
                    "w:fldChar"
                } else {
                    *stack.last()?
                },
            );
            if anchors.contains(ordinal) {
                layout.edges.entry(*ordinal).or_default().before = last;
            }
        }
        if !tag.empty {
            stack.push(tag.name);
        }
    }
    (prefixes.get("w") == Some(&W_NAMESPACE)).then_some(layout)
}

/// The elements a rewritten paragraph may sit in: story roots and table cells, not content
/// controls or other wrappers.
const PARENTS: [&str; 6] = [
    "w:body",
    "w:tc",
    "w:hdr",
    "w:ftr",
    "w:footnote",
    "w:endnote",
];

/// Range ends by kind, ID and whether the end opens its range, with how often each occurs.
type Ends<'x> = BTreeMap<(&'x str, String, bool), usize>;

struct Ranges<'x> {
    all: Ends<'x>,
    unpaired: Ends<'x>,
    field_or_move: bool,
}

/// Whether `xml`, one paragraph, holds no equations or content controls and pairs only within
/// itself, so that rewriting it leaves every other paragraph's markup whole.
fn simple(xml: &str) -> Option<bool> {
    Some(unwrapped(xml)? && pairs_within(xml)?)
}

/// Whether `xml` holds no equations or content controls.
fn unwrapped(xml: &str) -> Option<bool> {
    Some(!tags(xml)?.iter().any(|tag| {
        !tag.end
            && matches!(
                tag.name
                    .rsplit_once(':')
                    .map_or(tag.name, |(_, local)| local),
                "oMath" | "oMathPara" | "sdt"
            )
    }))
}

/// Whether rewriting `source`, one paragraph, as `written` leaves every other paragraph's markup
/// whole: neither holds equations or content controls, and both hold the same ends of ranges
/// whose other end lies outside the paragraph. Proofing marks are left out of the comparison,
/// as the model drops them and Word recomputes them.
fn keeps_ranges(source: &str, written: &str) -> Option<bool> {
    fn crossing(xml: &str) -> Option<(Ends<'_>, Ends<'_>)> {
        let mut ranges = ranges(xml)?;
        if ranges.field_or_move {
            return None;
        }
        ranges
            .unpaired
            .retain(|(kind, _, _), _| !matches!(*kind, "spell" | "gram"));
        ranges.all.retain(|(kind, id, _), _| {
            ranges.unpaired.contains_key(&(*kind, id.clone(), true))
                || ranges.unpaired.contains_key(&(*kind, id.clone(), false))
        });
        Some((ranges.unpaired, ranges.all))
    }
    Some(unwrapped(source)? && unwrapped(written)? && crossing(source)? == crossing(written)?)
}

/// Whether `xml`, one paragraph, holds both ends of everything in it that can pair across
/// paragraphs, each opened before it is closed, and no field character or move. `None` where
/// [`unpaired`] is.
fn pairs_within(xml: &str) -> Option<bool> {
    Some(unpaired(xml)?.is_empty())
}

/// The ends in `xml`, one paragraph, of bookmarks, permissions, comment and custom XML ranges and
/// proofing marks whose other end lies outside it: an end that closes what `xml` has not opened,
/// or opens what it does not close. `None` on a field character or move, whose partners lie
/// elsewhere, on a range marker without an ID and on a proofing mark without a type.
fn unpaired(xml: &str) -> Option<Ends<'_>> {
    let ranges = ranges(xml)?;
    (!ranges.field_or_move).then_some(ranges.unpaired)
}

fn ranges(xml: &str) -> Option<Ranges<'_>> {
    let mut open: HashMap<(&str, String), usize> = HashMap::new();
    let mut ends = Ends::new();
    let mut all = Ends::new();
    let mut field_or_move = false;
    for tag in tags(xml)? {
        if tag.end {
            continue;
        }
        let local = tag
            .name
            .rsplit_once(':')
            .map_or(tag.name, |(_, local)| local);
        if local == "fldChar" || local.starts_with("move") || local.starts_with("customXmlMove") {
            field_or_move = true;
        }
        let (kind, id, opens) = if local == "proofErr" {
            let (kind, opens) = match &*unescaped(xml, attribute(&tag, "w:type")?)? {
                "spellStart" => ("spell", true),
                "spellEnd" => ("spell", false),
                "gramStart" => ("gram", true),
                "gramEnd" => ("gram", false),
                _ => return None,
            };
            (kind, String::new(), opens)
        } else {
            let (kind, opens) = match local {
                "bookmarkStart" => ("bookmark", true),
                "bookmarkEnd" => ("bookmark", false),
                "permStart" => ("perm", true),
                "permEnd" => ("perm", false),
                name => match (
                    name.strip_suffix("RangeStart"),
                    name.strip_suffix("RangeEnd"),
                ) {
                    (Some(kind), _) => (kind, true),
                    (_, Some(kind)) => (kind, false),
                    _ => continue,
                },
            };
            (kind, unescaped(xml, attribute(&tag, "w:id")?)?, opens)
        };
        *all.entry((kind, id.clone(), opens)).or_insert(0) += 1;
        let count = open.entry((kind, id.clone())).or_insert(0);
        if opens {
            *count += 1;
        } else if *count == 0 {
            *ends.entry((kind, id, false)).or_insert(0) += 1;
        } else {
            *count -= 1;
        }
    }
    for ((kind, id), count) in open {
        if count > 0 {
            ends.insert((kind, id, true), count);
        }
    }
    Some(Ranges {
        all,
        unpaired: ends,
        field_or_move,
    })
}

fn unique_ranges<'x>(fragments: impl IntoIterator<Item = &'x str>) -> Option<bool> {
    let mut ends = Ends::new();
    for fragment in fragments {
        for ((kind, id, opens), count) in ranges(fragment)?.all {
            if matches!(kind, "spell" | "gram") {
                continue;
            }
            let total = ends.entry((kind, id, opens)).or_insert(0);
            *total += count;
            if *total > 1 {
                return Some(false);
            }
        }
    }
    Some(true)
}

/// The comment markers `written`, the XML written for an unchanged paragraph, adds to `source`,
/// the paragraph's source XML, each as the offset in `source` it goes at and its markup in
/// `written`: a range start or end as its tag, a reference as the run that holds only it. Each
/// goes right after the marker it follows in `written`, or after that marker's reference run.
/// `None` unless `written` holds `source`'s comment markers in order with only such markers
/// added, each after a marker that is an empty range tag or alone in its run.
fn added_comment_ranges<'w>(source: &str, written: &'w str) -> Option<Vec<(usize, &'w str)>> {
    fn markers<'x>(xml: &'x str, tags: &[Tag<'x>]) -> Option<Vec<(usize, String)>> {
        tags.iter()
            .enumerate()
            .filter(|(_, tag)| {
                !tag.end
                    && matches!(
                        tag.name,
                        "w:commentRangeStart" | "w:commentRangeEnd" | "w:commentReference"
                    )
            })
            .map(|(index, tag)| Some((index, unescaped(xml, attribute(tag, "w:id")?)?)))
            .collect()
    }
    /// The span of the marker at `index`: its empty range tag, or the run it is alone in.
    fn markup(tags: &[Tag<'_>], index: usize) -> Option<Range<usize>> {
        let tag = &tags[index];
        if tag.name != "w:commentReference" {
            return tag.empty.then(|| tag.range.clone());
        }
        let close = tags
            .get(index + 1)
            .filter(|tag| tag.end && tag.name == "w:r")?;
        let mut at = index.checked_sub(1)?;
        if tags[at].end && tags[at].name == "w:rPr" {
            let mut depth = 0usize;
            loop {
                if tags[at].end {
                    depth += 1;
                } else if !tags[at].empty {
                    depth = depth.checked_sub(1)?;
                }
                if depth == 0 {
                    break;
                }
                at = at.checked_sub(1)?;
            }
            at = at.checked_sub(1)?;
        }
        let open = &tags[at];
        (open.name == "w:r" && !open.end && !open.empty).then(|| open.range.start..close.range.end)
    }
    let source_tags = tags(source)?;
    let written_tags = tags(written)?;
    let source_markers = markers(source, &source_tags)?;
    let mut held = source_markers.iter().peekable();
    let mut after = None;
    let mut added = Vec::new();
    for (index, id) in markers(written, &written_tags)? {
        let name = written_tags[index].name;
        if let Some((kept, _)) =
            held.next_if(|(kept, kept_id)| source_tags[*kept].name == name && *kept_id == id)
        {
            after = markup(&source_tags, *kept).map(|span| span.end);
            continue;
        }
        added.push((after?, &written[markup(&written_tags, index)?]));
    }
    (held.next().is_none() && !added.is_empty()).then_some(added)
}

/// `source` with each tag in `added` inserted at its offset.
fn with_insertions(source: &str, added: &[(usize, &str)]) -> String {
    let mut output =
        String::with_capacity(source.len() + added.iter().map(|(_, tag)| tag.len()).sum::<usize>());
    let mut cursor = 0;
    for &(at, tag) in added {
        output.push_str(&source[cursor..at]);
        output.push_str(tag);
        cursor = at;
    }
    output.push_str(&source[cursor..]);
    output
}

pub(crate) fn paragraph_span(xml: &str, occurrence: &ParagraphOccurrence) -> Option<Range<usize>> {
    let start = occurrence.tag.start;
    let relative = element_span(&xml[start..], &[])?;
    Some(start + relative.start..start + relative.end)
}

/// The paragraphs written without a `sourceOrdinal`, each with its anchor from `anchors`, which
/// lists them in writing order; `None` unless each was written right after its `After` anchor
/// or right before its `Before` anchor, besides others with the same anchor.
fn placed<'r>(
    recorded: &'r RecordedParagraphs,
    anchors: &[S13SpliceAnchor],
) -> Option<Vec<(S13SpliceAnchor, &'r str)>> {
    if recorded.inserted.len() != anchors.len() {
        return None;
    }
    let mut next = anchors.iter();
    let entries: Vec<Result<u32, S13SpliceAnchor>> = recorded
        .order
        .iter()
        .map(|entry| match entry {
            Some(ordinal) => Some(Ok(*ordinal)),
            None => next.next().map(|anchor| Err(*anchor)),
        })
        .collect::<Option<_>>()?;
    for (index, entry) in entries.iter().enumerate() {
        let Err(anchor) = *entry else {
            continue;
        };
        let neighbour = match anchor {
            S13SpliceAnchor::After(_) => index.checked_sub(1).and_then(|at| entries.get(at)),
            S13SpliceAnchor::Before(_) => entries.get(index + 1),
        };
        match neighbour {
            Some(Ok(ordinal)) if *ordinal == anchor.ordinal() => {}
            Some(Err(other)) if *other == anchor => {}
            _ => return None,
        }
    }
    Some(
        anchors
            .iter()
            .copied()
            .zip(recorded.inserted.iter().map(String::as_str))
            .collect(),
    )
}

struct RowSpan {
    span: Range<usize>,
    table: usize,
    cells: usize,
    paragraphs: Vec<usize>,
}

fn row_spans(xml: &str) -> Option<Vec<RowSpan>> {
    let paragraphs = paragraph_occurrences(xml)?;
    let mut tables = Vec::new();
    let mut opened = Vec::new();
    let mut rows = Vec::new();
    let mut table_index = 0;
    for tag in tags(xml)? {
        match (tag.name, tag.end, tag.empty) {
            ("w:tbl", false, false) => {
                tables.push(table_index);
                table_index += 1;
            }
            ("w:tbl", true, _) => {
                tables.pop()?;
            }
            ("w:tr", false, false) => opened.push((tag.range.start, *tables.last()?, 0)),
            ("w:tc", false, _) => opened.last_mut()?.2 += 1,
            ("w:tr", true, _) => {
                let (start, table, cells) = opened.pop()?;
                let span = start..tag.range.end;
                rows.push(RowSpan {
                    paragraphs: paragraphs
                        .iter()
                        .enumerate()
                        .filter_map(|(at, paragraph)| {
                            span.contains(&paragraph.tag.start).then_some(at)
                        })
                        .collect(),
                    span,
                    table,
                    cells,
                });
            }
            _ => {}
        }
    }
    rows.sort_by_key(|row| row.span.start);
    Some(rows)
}

/// The entry of `recorded.order` each `w:p` of `serialized` was written as: `None` for one
/// inside another paragraph or for the empty paragraph of a cell without content. `None` unless
/// the paragraphs outside any other are exactly those `recorded` lists, in its order.
fn recorded_entries(serialized: &str, recorded: &RecordedParagraphs) -> Option<Vec<Option<usize>>> {
    let mut placeholders = recorded.placeholders.iter().copied().peekable();
    let mut entries = Vec::new();
    let mut next = 0;
    let mut end = 0;
    for occurrence in paragraph_occurrences(serialized)? {
        if occurrence.tag.start < end {
            entries.push(None);
            continue;
        }
        end = paragraph_span(serialized, &occurrence)?.end;
        if placeholders.next_if_eq(&next).is_some() {
            if serialized.get(occurrence.tag) != Some("<w:p/>") {
                return None;
            }
            entries.push(None);
        } else if next < recorded.order.len() {
            entries.push(Some(next));
            next += 1;
        } else {
            return None;
        }
    }
    if next != recorded.order.len() || placeholders.next().is_some() {
        return None;
    }
    Some(entries)
}

fn unchanged_placeholder_tables(
    source_rows: &[RowSpan],
    target_rows: &[RowSpan],
    entries: &[Option<usize>],
    recorded: &RecordedParagraphs,
    removed: &BTreeSet<u32>,
) -> Option<()> {
    let tables: BTreeSet<usize> = target_rows
        .iter()
        .filter(|row| {
            !row.paragraphs.is_empty() && row.paragraphs.iter().all(|at| entries[*at].is_none())
        })
        .map(|row| row.table)
        .chain(
            source_rows
                .iter()
                .filter(|row| {
                    !row.paragraphs.is_empty()
                        && row.paragraphs.iter().all(|at| {
                            !recorded.written.contains_key(&(*at as u32))
                                && !removed.contains(&(*at as u32))
                        })
                })
                .map(|row| row.table),
        )
        .collect();
    for table in tables {
        let source: Vec<_> = source_rows
            .iter()
            .filter(|row| row.table == table)
            .collect();
        let target: Vec<_> = target_rows
            .iter()
            .filter(|row| row.table == table)
            .collect();
        if source.len() != target.len() {
            return None;
        }
        let mut identified = false;
        for (source, target) in source.iter().zip(&target) {
            if source.cells != target.cells
                || source.paragraphs.len() != target.paragraphs.len()
                || source
                    .paragraphs
                    .iter()
                    .any(|at| removed.contains(&(*at as u32)))
            {
                return None;
            }
            let expected: Vec<_> = source
                .paragraphs
                .iter()
                .map(|at| *at as u32)
                .filter(|ordinal| recorded.written.contains_key(ordinal))
                .collect();
            let actual = target
                .paragraphs
                .iter()
                .filter_map(|at| entries[*at])
                .map(|at| recorded.order[at])
                .collect::<Option<Vec<_>>>()?;
            if actual != expected {
                return None;
            }
            identified |= !expected.is_empty();
        }
        if !identified {
            return None;
        }
    }
    Some(())
}

#[derive(Default)]
struct RowEdits<'a> {
    edits: Vec<(Range<usize>, u8, &'a str)>,
    inserted: BTreeSet<usize>,
    removed: BTreeSet<u32>,
}

fn row_edits<'a>(
    source: &str,
    serialized: &'a str,
    recorded: &RecordedParagraphs,
    inserted: &[(S13SpliceAnchor, &str)],
    removed: &BTreeSet<u32>,
) -> Option<RowEdits<'a>> {
    if inserted.is_empty() && removed.is_empty() {
        return Some(RowEdits::default());
    }
    let source_rows = row_spans(source)?;
    let mut target_rows = row_spans(serialized)?;
    let mut edits = RowEdits::default();
    if source_rows.is_empty() {
        return Some(edits);
    }
    let entries = recorded_entries(serialized, recorded)?;
    unchanged_placeholder_tables(&source_rows, &target_rows, &entries, recorded, removed)?;
    for row in &mut target_rows {
        row.paragraphs = row
            .paragraphs
            .iter()
            .filter_map(|at| entries[*at])
            .collect();
    }
    let mut insertion_at = 0;
    let insertion_indices: Vec<Option<usize>> = recorded
        .order
        .iter()
        .map(|ordinal| {
            ordinal.is_none().then(|| {
                let at = insertion_at;
                insertion_at += 1;
                at
            })
        })
        .collect();
    for row in &target_rows {
        if row.paragraphs.is_empty()
            || row
                .paragraphs
                .iter()
                .any(|at| recorded.order[*at].is_some())
        {
            continue;
        }
        let indices = row
            .paragraphs
            .iter()
            .map(|at| insertion_indices[*at])
            .collect::<Option<Vec<_>>>()?;
        let anchor = inserted.get(*indices.first()?)?.0;
        if indices
            .iter()
            .any(|at| inserted.get(*at).map(|(anchor, _)| *anchor) != Some(anchor))
        {
            return None;
        }
        let source_row = source_rows
            .iter()
            .find(|row| row.paragraphs.contains(&(anchor.ordinal() as usize)))?;
        if source_row.table != row.table || !pairs_within(&serialized[row.span.clone()])? {
            return None;
        }
        let position = match anchor {
            S13SpliceAnchor::Before(_) => source_row.span.start,
            S13SpliceAnchor::After(_) => source_row.span.end,
        };
        edits.inserted.extend(indices);
        edits
            .edits
            .push((position..position, 1, &serialized[row.span.clone()]));
    }
    for row in &source_rows {
        let dropped = |at: &usize| removed.contains(&(*at as u32));
        if !row.paragraphs.iter().any(dropped)
            || row
                .paragraphs
                .iter()
                .any(|at| recorded.written.contains_key(&(*at as u32)))
        {
            continue;
        }
        if !row.paragraphs.iter().all(dropped) {
            return None;
        }
        if edits
            .edits
            .iter()
            .any(|(span, _, _)| span.start <= row.span.start && span.end >= row.span.end)
        {
            continue;
        }
        if target_rows
            .iter()
            .filter(|target| target.table == row.table)
            .any(|target| {
                target
                    .paragraphs
                    .iter()
                    .filter_map(|at| recorded.order[*at])
                    .any(|ordinal| row.paragraphs.contains(&(ordinal as usize)))
            })
            || !pairs_within(&source[row.span.clone()])?
        {
            return None;
        }
        edits
            .removed
            .extend(row.paragraphs.iter().map(|at| *at as u32));
        edits.edits.push((row.span.clone(), 2, ""));
    }
    Some(edits)
}

/// `source` with the paragraphs `part` addresses rewritten from `recorded` where they changed
/// or their written XML no longer agrees with their source on comments, revisions, notes or
/// relationships, the removed ones dropped and the new ones inserted at their anchors, every
/// other byte kept except paragraph IDs and the comment markers an unchanged paragraph gains,
/// which go into its source XML (see [`added_comment_ranges`]): the written IDs, and `assignments` for source
/// paragraphs no model paragraph is written from. `None` when the part cannot be spliced, so
/// the caller writes `serialized` whole: among others when a paragraph to remove or insert is
/// not [`simple`], one to rewrite does not [`keeps_ranges`] from its source to its written XML,
/// one of them sits in anything but a story root or table cell, and when the result would need a
/// namespace declaration the source does not have.
pub(crate) fn splice_story_part(
    source: &str,
    serialized: &str,
    recorded: &RecordedParagraphs,
    part: &S13SplicedPart,
    assignments: &BTreeMap<u32, String>,
) -> Option<String> {
    if !crate::xml::reads_as_written(source.as_bytes()) {
        return None;
    }
    let written = &recorded.written;
    let removed: BTreeSet<u32> = part.removed.iter().copied().collect();
    if written.len() != part.paragraphs.len()
        || part
            .paragraphs
            .iter()
            .any(|ordinal| !written.contains_key(ordinal))
        || removed.len() != part.removed.len()
        || removed.iter().any(|ordinal| written.contains_key(ordinal))
    {
        return None;
    }
    let inserted = placed(recorded, &part.inserted)?;
    let changed: BTreeSet<u32> = part.changed.iter().copied().collect();
    let occurrences = paragraph_occurrences(source)?;
    let rows = row_edits(source, serialized, recorded, &inserted, &removed)?;
    let mut spans = BTreeMap::new();
    for &ordinal in written.keys().chain(&removed) {
        let occurrence = occurrences
            .get(ordinal as usize)
            .filter(|occurrence| occurrence.ordinal == ordinal)?;
        spans.insert(ordinal, paragraph_span(source, occurrence)?);
    }
    let mut outside_source = comment_counts(source)?;
    let mut outside_written = comment_counts(serialized)?;
    let mut end = 0;
    for (ordinal, span) in &spans {
        if span.start < end {
            return None;
        }
        end = span.end;
        subtract(&mut outside_source, comment_counts(&source[span.clone()])?);
        if let Some(xml) = written.get(ordinal) {
            subtract(&mut outside_written, comment_counts(xml)?);
        }
    }
    for (_, xml) in &inserted {
        subtract(&mut outside_written, comment_counts(xml)?);
    }
    if outside_source != outside_written {
        return None;
    }
    let mut replaced = BTreeSet::new();
    let mut ranged = BTreeSet::new();
    for (&ordinal, xml) in written {
        if changed.contains(&ordinal) {
            replaced.insert(ordinal);
            continue;
        }
        let paragraph = &source[spans[&ordinal].clone()];
        let wanted = marks(xml)?;
        if marks(paragraph)?.admit(&wanted) {
            continue;
        }
        let rangeable = added_comment_ranges(paragraph, xml).is_some_and(|added| {
            marks(&with_insertions(paragraph, &added)).is_some_and(|patched| patched.admit(&wanted))
        });
        if rangeable {
            ranged.insert(ordinal);
        } else {
            replaced.insert(ordinal);
        }
    }
    let anchors: BTreeSet<u32> = inserted
        .iter()
        .map(|(anchor, _)| anchor.ordinal())
        .collect();
    let layout = layout(source, &spans, &anchors)?;
    let in_story = |ordinal: &u32| {
        layout
            .parents
            .get(ordinal)
            .is_some_and(|parent| PARENTS.contains(parent))
    };
    for ordinal in replaced.iter().chain(&removed) {
        let paragraph = &source[spans[ordinal].clone()];
        let admitted = match written.get(ordinal) {
            Some(xml) => keeps_ranges(paragraph, xml)?,
            None => simple(paragraph)?,
        };
        if !in_story(ordinal) || !admitted {
            return None;
        }
    }
    for (at, (anchor, xml)) in inserted.iter().enumerate() {
        if rows.inserted.contains(&at) {
            continue;
        }
        if !in_story(&anchor.ordinal())
            || !layout
                .edges
                .get(&anchor.ordinal())
                .is_some_and(|edges| edges.admit(*anchor))
            || !simple(xml)?
        {
            return None;
        }
    }
    let mut prefixes = BTreeSet::new();
    for ordinal in &replaced {
        prefixes.extend(required_prefixes(&written[ordinal])?);
    }
    for (_, xml) in &inserted {
        prefixes.extend(required_prefixes(xml)?);
    }
    for (_, _, xml) in &rows.edits {
        prefixes.extend(required_prefixes(xml)?);
    }
    for ordinal in &ranged {
        for (_, tag) in added_comment_ranges(&source[spans[ordinal].clone()], &written[ordinal])? {
            prefixes.extend(required_prefixes(tag)?);
        }
    }
    let root = root_tag(source)?;
    let bound = bindings(source, &root);
    if prefixes
        .iter()
        .any(|prefix| !bound.contains_key(prefix) && *prefix != "w14")
    {
        return None;
    }
    let rooted = root_with_prefixes(source, serialized, &prefixes)?;
    let source = rooted.as_ref();
    if !root_binds(source, serialized, &prefixes, false)? {
        return None;
    }
    let occurrences = match &rooted {
        Cow::Borrowed(_) => occurrences,
        Cow::Owned(patched) => paragraph_occurrences(patched)?,
    };
    let mut ids = assignments.clone();
    for (&ordinal, xml) in written {
        let occurrence = &occurrences[ordinal as usize];
        if replaced.contains(&ordinal) {
            continue;
        }
        if let Some(id) = paragraph_occurrences(xml)?
            .into_iter()
            .next()?
            .para_id
            .filter(|id| occurrence.para_id.as_ref() != Some(id))
        {
            ids.insert(ordinal, id);
        }
    }
    let source: Cow<str> = if ids.is_empty() {
        Cow::Borrowed(source)
    } else {
        let patched = patch_paragraph_ids(source, &ids, true)?;
        if patched.matches("xmlns").count() != source.matches("xmlns").count() {
            return None;
        }
        Cow::Owned(patched)
    };
    let occurrences = match &source {
        Cow::Borrowed(_) => occurrences,
        Cow::Owned(patched) => paragraph_occurrences(patched)?,
    };
    let span = |ordinal: u32| paragraph_span(&source, occurrences.get(ordinal as usize)?);
    let rows = row_edits(&source, serialized, recorded, &inserted, &removed)?;
    let mut edits: Vec<(Range<usize>, u8, &str)> = rows.edits;
    for &ordinal in replaced.iter().chain(&removed) {
        if rows.removed.contains(&ordinal) {
            continue;
        }
        let xml = written.get(&ordinal).map_or("", String::as_str);
        edits.push((span(ordinal)?, 2, xml));
    }
    for &ordinal in &ranged {
        let range = span(ordinal)?;
        for (at, tag) in added_comment_ranges(&source[range.clone()], &written[&ordinal])? {
            edits.push((range.start + at..range.start + at, 0, tag));
        }
    }
    for (at, (anchor, xml)) in inserted.iter().enumerate() {
        if rows.inserted.contains(&at) {
            continue;
        }
        let range = span(anchor.ordinal())?;
        edits.push(match anchor {
            S13SpliceAnchor::After(_) => (range.end..range.end, 0, xml),
            S13SpliceAnchor::Before(_) => (range.start..range.start, 1, xml),
        });
    }
    edits.sort_by_key(|(range, rank, _)| (range.start, *rank));
    let mut fragments = Vec::with_capacity(edits.len() * 2 + 1);
    let mut cursor = 0;
    for (range, _, text) in edits {
        if range.start < cursor {
            return None;
        }
        fragments.push(&source[cursor..range.start]);
        fragments.push(text);
        cursor = range.end;
    }
    fragments.push(&source[cursor..]);
    unique_ranges(fragments.iter().copied())?.then(|| fragments.concat())
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::serializer::{S13SaveRequest, write_docx_s13};
    use serde_json::json;
    use sha2::{Digest, Sha256};

    const ROOT: &str = "<w:document xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\" xmlns:w14=\"http://schemas.microsoft.com/office/word/2010/wordml\" xmlns:mc=\"http://schemas.openxmlformats.org/markup-compatibility/2006\" mc:Ignorable=\"w14\">";
    const W_ONLY_ROOT: &str =
        "<w:document xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\">";
    const SERIALIZED_ROOT: &str = "<w:document xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\" xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\" xmlns:w14=\"http://schemas.microsoft.com/office/word/2010/wordml\" xmlns:mc=\"http://schemas.openxmlformats.org/markup-compatibility/2006\" mc:Ignorable=\"w14\">";

    /// `W_ONLY_ROOT` with only the declarations a paragraph id needs appended.
    fn patched_w_only_root() -> String {
        format!(
            "{} xmlns:w14=\"http://schemas.microsoft.com/office/word/2010/wordml\" xmlns:mc=\"{MC_NAMESPACE}\" mc:Ignorable=\"w14\">",
            W_ONLY_ROOT.strip_suffix('>').expect("root tag")
        )
    }

    /// Splices `source` with `paragraphs` as the written XML of their ordinals, serialized
    /// between `before` and `after`.
    fn splice_around(
        source: &str,
        paragraphs: &[(u32, &str)],
        changed: &[u32],
        assignments: &[(u32, &str)],
        (before, after): (&str, &str),
    ) -> Option<String> {
        let written: Vec<(Option<u32>, &str)> = paragraphs
            .iter()
            .map(|(ordinal, xml)| (Some(*ordinal), *xml))
            .collect();
        let edit = Edit {
            changed,
            ..Edit::default()
        };
        splice_written(source, &written, edit, assignments, (before, after))
    }

    #[derive(Clone, Copy, Default)]
    struct Edit<'a> {
        changed: &'a [u32],
        inserted: &'a [S13SpliceAnchor],
        removed: &'a [u32],
    }

    /// Splices `source` with `written` as the paragraphs the serializer wrote, in order: by
    /// ordinal, or `None` for one the source lacks.
    fn splice_written(
        source: &str,
        written: &[(Option<u32>, &str)],
        edit: Edit,
        assignments: &[(u32, &str)],
        (before, after): (&str, &str),
    ) -> Option<String> {
        let mut recorded = RecordedParagraphs::default();
        for (ordinal, xml) in written {
            match ordinal {
                Some(ordinal) => {
                    recorded.written.insert(*ordinal, (*xml).to_owned());
                }
                None => recorded.inserted.push((*xml).to_owned()),
            }
            recorded.order.push(*ordinal);
        }
        let body: String = written.iter().map(|(_, xml)| *xml).collect();
        let serialized =
            format!("{SERIALIZED_ROOT}<w:body>{before}{body}{after}</w:body></w:document>");
        let part = S13SplicedPart {
            part: "word/document.xml".to_owned(),
            sha256: String::new(),
            paragraphs: written.iter().filter_map(|(ordinal, _)| *ordinal).collect(),
            changed: edit.changed.to_vec(),
            inserted: edit.inserted.to_vec(),
            removed: edit.removed.to_vec(),
        };
        let assignments = assignments
            .iter()
            .map(|(ordinal, id)| (*ordinal, (*id).to_owned()))
            .collect();
        splice_story_part(source, &serialized, &recorded, &part, &assignments)
    }

    fn splice(source: &str, paragraphs: &[(u32, &str)], changed: &[u32]) -> Option<String> {
        splice_around(source, paragraphs, changed, &[], ("", ""))
    }

    fn body(paragraphs: &str) -> String {
        format!("{ROOT}<w:body>{paragraphs}</w:body></w:document>")
    }

    /// A package holding only `xml` as its main document part.
    fn package(xml: &str) -> Vec<u8> {
        ooxml_opc::rezip_parts(&[
            (
                "[Content_Types].xml".to_owned(),
                br#"<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>"#.to_vec(),
            ),
            (
                "_rels/.rels".to_owned(),
                br#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="doc" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>"#.to_vec(),
            ),
            ("word/document.xml".to_owned(), xml.as_bytes().to_vec()),
        ])
        .unwrap()
    }

    /// The main document part of the package `bytes`.
    fn document_part(bytes: &[u8]) -> Vec<u8> {
        ooxml_opc::unzip_parts(bytes)
            .unwrap()
            .into_iter()
            .find(|(path, _)| path == "word/document.xml")
            .unwrap()
            .1
    }

    fn assert_whole_part_fallback(source: &str, written: &[(Option<u32>, &str)], edit: Edit) {
        assert_eq!(splice_written(source, written, edit, &[], ("", "")), None);
        let original = package(source);
        let target = body(&written.iter().map(|(_, xml)| *xml).collect::<String>());
        let parsed = crate::parse_docx_s9_wire(&package(&target), Default::default())
            .unwrap()
            .document
            .package;
        let mut content = serde_json::to_value(&parsed.document.content).unwrap();
        let paragraphs = content.as_array_mut().unwrap();
        assert_eq!(paragraphs.len(), written.len());
        for (paragraph, (ordinal, _)) in paragraphs.iter_mut().zip(written) {
            if let Some(ordinal) = ordinal {
                paragraph["sourceOrdinal"] = json!(ordinal);
            }
        }
        let mut request: S13SaveRequest = serde_json::from_value(json!({
            "determinism": {"seed": "0".repeat(64), "now": "2000-01-01T00:00:00.000Z"},
            "document": {"content": content},
            "options": {"updateModifiedDate": false},
            "paragraphIds": {"splicedParts": [{
                "part": "word/document.xml",
                "sha256": format!("{:x}", Sha256::digest(source.as_bytes())),
                "paragraphs": written.iter().filter_map(|(ordinal, _)| *ordinal).collect::<Vec<_>>(),
                "changed": edit.changed,
                "inserted": edit.inserted,
                "removed": edit.removed,
            }]},
        }))
        .unwrap();
        let saved = write_docx_s13(request.clone(), &original).unwrap();
        request
            .paragraph_ids
            .as_mut()
            .unwrap()
            .spliced_parts
            .clear();
        let whole = write_docx_s13(request, &original).unwrap();
        assert_eq!(document_part(&saved), document_part(&whole));
        assert_ne!(document_part(&saved), source.as_bytes());
    }

    #[test]
    fn structural_edits_inside_a_complex_field_fall_back_to_the_whole_part() {
        let opening = "<w:p><w:r><w:fldChar w:fldCharType=\"begin\"/></w:r><w:r><w:instrText>TOC</w:instrText></w:r><w:r><w:fldChar w:fldCharType=\"separate\"/></w:r></w:p>";
        let middle = "<w:p><w:r><w:t>Entry</w:t></w:r></w:p>";
        let closing = "<w:p><w:r><w:fldChar w:fldCharType=\"end\"/></w:r></w:p>";
        let new = "<w:p><w:r><w:t>New entry</w:t></w:r></w:p>";
        let source = body(&format!("{opening}{middle}{closing}"));
        assert!(crate::xml::reads_as_written(source.as_bytes()));
        for (written, edit) in [
            (
                vec![
                    (Some(0), opening),
                    (Some(1), middle),
                    (None, new),
                    (Some(2), closing),
                ],
                Edit {
                    inserted: &[S13SpliceAnchor::After(1)],
                    ..Edit::default()
                },
            ),
            (
                vec![(Some(0), opening), (Some(2), closing)],
                Edit {
                    removed: &[1],
                    ..Edit::default()
                },
            ),
            (
                vec![
                    (Some(0), opening),
                    (None, new),
                    (Some(1), middle),
                    (Some(2), closing),
                ],
                Edit {
                    inserted: &[S13SpliceAnchor::After(0)],
                    ..Edit::default()
                },
            ),
        ] {
            assert_whole_part_fallback(&source, &written, edit);
        }
    }

    #[test]
    fn structural_edits_of_xml_not_read_as_written_fall_back_to_the_whole_part() {
        let first = "<w:p><w:r><w:t>A</w:t></w:r></w:p>";
        let second = "<w:p><w:r><w:t>B</w:t></w:r></w:p>";
        let new = "<w:p><w:r><w:t>New</w:t></w:r></w:p>";
        for source in [
            body(&format!("{first}{second}<!-- note -->")),
            body(&format!("{first}{second}<?pi x?>")),
            body(&format!("{first}{second}")).replace("<w:t>A</w:t>", "<w:t>A & B</w:t>"),
        ] {
            assert!(!crate::xml::reads_as_written(source.as_bytes()));
            assert_whole_part_fallback(
                &source,
                &[(Some(0), first), (None, new), (Some(1), second)],
                Edit {
                    inserted: &[S13SpliceAnchor::After(0)],
                    ..Edit::default()
                },
            );
            assert_whole_part_fallback(
                &source,
                &[(Some(1), second)],
                Edit {
                    removed: &[0],
                    ..Edit::default()
                },
            );
        }
    }

    #[test]
    fn insertions_at_inadmissible_edges_fall_back_to_the_whole_part() {
        let first = "<w:p><w:r><w:t>A</w:t></w:r></w:p>";
        let second = "<w:p><w:r><w:t>B</w:t></w:r></w:p>";
        let new = "<w:p><w:r><w:t>New</w:t></w:r></w:p>";
        let source = body(&format!(
            "{first}<w:bookmarkStart w:id=\"1\" w:name=\"mark\"/><w:bookmarkEnd w:id=\"1\"/>{second}"
        ));
        assert!(crate::xml::reads_as_written(source.as_bytes()));
        for anchor in [S13SpliceAnchor::After(0), S13SpliceAnchor::Before(1)] {
            assert_whole_part_fallback(
                &source,
                &[(Some(0), first), (None, new), (Some(1), second)],
                Edit {
                    inserted: &[anchor],
                    ..Edit::default()
                },
            );
        }
    }

    #[test]
    fn insertions_with_ambiguous_neighbours_fall_back_to_the_whole_part() {
        let first = "<w:p><w:r><w:t>A</w:t></w:r></w:p>";
        let second = "<w:p><w:r><w:t>B</w:t></w:r></w:p>";
        let new = "<w:p><w:r><w:t>New</w:t></w:r></w:p>";
        let source = body(&format!("{first}{second}"));
        for anchor in [S13SpliceAnchor::After(0), S13SpliceAnchor::Before(0)] {
            assert_whole_part_fallback(
                &source,
                &[(Some(0), first), (Some(1), second), (None, new)],
                Edit {
                    inserted: &[anchor],
                    ..Edit::default()
                },
            );
        }
        assert_whole_part_fallback(
            &source,
            &[
                (Some(0), first),
                (None, new),
                (None, new),
                (Some(1), second),
            ],
            Edit {
                inserted: &[S13SpliceAnchor::After(0), S13SpliceAnchor::Before(1)],
                ..Edit::default()
            },
        );
    }

    #[test]
    fn inserts_new_paragraphs_at_their_anchors_and_drops_removed_ones() {
        let first =
            "<w:p w14:paraId=\"0000000A\"><w:r><w:t>first half second half</w:t></w:r></w:p>";
        let middle = "<w:p w14:paraId=\"0000000B\"><w:r><w:t>removed</w:t></w:r></w:p>";
        let last =
            "<w:p w14:paraId=\"0000000C\" w:rsidR=\"00AB12CD\"><w:r><w:t>kept</w:t></w:r></w:p>";
        let source = body(&format!("{first}\n{middle}\n  {last}"));
        let split = "<w:p w14:paraId=\"0000000A\"><w:r><w:t>first half</w:t></w:r></w:p>";
        let new = "<w:p w14:paraId=\"0000000D\"><w:r><w:t>second half</w:t></w:r></w:p>";
        let written_last = "<w:p w14:paraId=\"0000000C\"><w:r><w:t>kept</w:t></w:r></w:p>";
        let edit = Edit {
            changed: &[0],
            inserted: &[S13SpliceAnchor::After(0)],
            removed: &[1],
        };
        assert_eq!(
            splice_written(
                &source,
                &[(Some(0), split), (None, new), (Some(2), written_last)],
                edit,
                &[],
                ("", "")
            ),
            Some(body(&format!("{split}{new}\n\n  {last}")))
        );
    }

    #[test]
    fn a_paragraph_inserted_first_in_a_cell_goes_after_the_cell_properties() {
        let cell = "<w:p w14:paraId=\"0000000A\"><w:r><w:t>cell</w:t></w:r></w:p>";
        let source = body(&format!(
            "<w:tbl><w:tr><w:tc><w:tcPr><w:tcW w:w=\"10\" w:type=\"dxa\"/></w:tcPr>{cell}</w:tc></w:tr></w:tbl><w:p w14:paraId=\"0000000B\"/>"
        ));
        let new = "<w:p w14:paraId=\"0000000C\"><w:r><w:t>new</w:t></w:r></w:p>";
        let edit = Edit {
            inserted: &[S13SpliceAnchor::Before(0)],
            ..Edit::default()
        };
        let spliced = splice_written(
            &source,
            &[
                (None, new),
                (Some(0), cell),
                (Some(1), "<w:p w14:paraId=\"0000000B\"/>"),
            ],
            edit,
            &[],
            ("", ""),
        );
        assert_eq!(spliced, Some(source.replace(cell, &format!("{new}{cell}"))));
    }

    /// The `word/document.xml` a save writes for `source` once `edit` changed its model: spliced
    /// as `spliced` describes, or whole without it.
    fn saved_document(
        source: &str,
        edit: impl FnOnce(&mut Vec<serde_json::Value>),
        spliced: Option<serde_json::Value>,
    ) -> String {
        let original = package(source);
        let options = crate::S9ParseOptions {
            source_ordinals: true,
            ..Default::default()
        };
        let parsed = crate::parse_docx_s9_wire(&original, options)
            .unwrap()
            .document
            .package;
        let mut content = serde_json::to_value(&parsed.document.content).unwrap();
        edit(content.as_array_mut().unwrap());
        let mut request = json!({
            "determinism": {"seed": "0".repeat(64), "now": "2000-01-01T00:00:00.000Z"},
            "document": {"content": content},
            "options": {"updateModifiedDate": false},
        });
        if let Some(mut part) = spliced {
            part["part"] = json!("word/document.xml");
            part["sha256"] = json!(format!("{:x}", Sha256::digest(source.as_bytes())));
            request["paragraphIds"] = json!({"splicedParts": [part]});
        }
        let saved = write_docx_s13(serde_json::from_value(request).unwrap(), &original).unwrap();
        let (_, xml) = ooxml_opc::unzip_parts(&saved)
            .unwrap()
            .into_iter()
            .find(|(path, _)| path == "word/document.xml")
            .unwrap();
        String::from_utf8(xml).unwrap()
    }

    /// A body paragraph, then a table whose first column merges its two rows, then a paragraph:
    /// ordinals 0, 1 to 4 in reading order with 3 in the continuation cell, and 5.
    fn merged_table_body() -> String {
        let cell = |properties: &str, paragraph: &str| {
            format!(
                "<w:tc><w:tcPr><w:tcW w:w=\"10\" w:type=\"dxa\"/>{properties}</w:tcPr>{paragraph}</w:tc>"
            )
        };
        let paragraph = |id: &str, text: &str| {
            format!("<w:p w14:paraId=\"{id}\"><w:r><w:t>{text}</w:t></w:r></w:p>")
        };
        body(&format!(
            "{}<w:tbl><w:tblGrid><w:gridCol w:w=\"10\"/><w:gridCol w:w=\"10\"/></w:tblGrid><w:tr>{}{}</w:tr><w:tr>{}{}</w:tr></w:tbl>{}",
            paragraph("0000000A", "first"),
            cell(
                "<w:vMerge w:val=\"restart\"/>",
                &paragraph("0000000B", "merged")
            ),
            cell("", &paragraph("0000000C", "right")),
            cell("<w:vMerge/>", "<w:p w14:paraId=\"0000000D\"/>"),
            cell("", &paragraph("0000000E", "below")),
            paragraph("0000000F", "last"),
        ))
    }

    /// The model of [`merged_table_body`] as a save projects it: the continuation cell's
    /// paragraph is not read, so that cell has no content.
    fn without_continued_content(content: &mut [serde_json::Value]) {
        content[1]["rows"][1]["cells"][0]["content"] = json!([]);
    }

    /// `source`, a one-run paragraph, with no `sourceOrdinal`, the paragraph ID `id` and `text`.
    fn new_paragraph(source: &serde_json::Value, id: &str, text: &str) -> serde_json::Value {
        let mut paragraph = source.clone();
        paragraph.as_object_mut().unwrap().remove("sourceOrdinal");
        paragraph["paraId"] = json!(id);
        *paragraph.pointer_mut("/content/0/content/0/text").unwrap() = json!(text);
        paragraph
    }

    fn single_column_merged_table_body() -> String {
        body(concat!(
            "<w:p w14:paraId=\"0000000A\"><w:r><w:t>first</w:t></w:r></w:p>",
            "<w:tbl><w:tblGrid><w:gridCol w:w=\"10\"/></w:tblGrid>",
            "<w:tr><w:tc><w:tcPr><w:vMerge w:val=\"restart\"/></w:tcPr>",
            "<w:p w14:paraId=\"0000000B\"><w:r><w:t>merged</w:t></w:r></w:p>",
            "</w:tc></w:tr>",
            "<w:tr><w:tc><w:tcPr><w:vMerge/></w:tcPr>",
            "<w:p w14:paraId=\"0000000C\"/></w:tc></w:tr></w:tbl>",
            "<w:p w14:paraId=\"0000000D\"><w:r><w:t>last</w:t></w:r></w:p>",
        ))
    }

    #[test]
    fn body_insertion_keeps_a_placeholder_only_row_byte_identical() {
        let source = single_column_merged_table_body();
        let saved = saved_document(
            &source,
            |content| {
                without_continued_content(content);
                let new = new_paragraph(&content[0], "00000010", "new");
                content.insert(1, new);
            },
            Some(json!({
                "paragraphs": [0, 1, 3],
                "inserted": [S13SpliceAnchor::After(0)],
            })),
        );
        let (head, tail) = source.split_at(source.find("<w:tbl>").unwrap());
        assert!(saved.starts_with(head) && saved.ends_with(tail), "{saved}");
        let new = &saved[head.len()..saved.len() - tail.len()];
        assert!(new.starts_with("<w:p ") && new.ends_with("</w:p>"), "{new}");
        assert!(new.contains("<w:t>new</w:t>"), "{new}");
    }

    #[test]
    fn row_edits_next_to_placeholder_only_rows_write_the_whole_part() {
        let source = single_column_merged_table_body();
        for insert in [true, false] {
            let edit = |content: &mut Vec<serde_json::Value>| {
                without_continued_content(content);
                if insert {
                    let mut row = content[1]["rows"][0].clone();
                    row["cells"][0]["content"][0] =
                        new_paragraph(&row["cells"][0]["content"][0], "00000010", "added");
                    content[1]["rows"].as_array_mut().unwrap().insert(1, row);
                } else {
                    content[1]["rows"].as_array_mut().unwrap().remove(0);
                }
            };
            let part = if insert {
                json!({"paragraphs": [0, 1, 3], "inserted": [S13SpliceAnchor::After(1)]})
            } else {
                json!({"paragraphs": [0, 3], "removed": [1]})
            };
            let saved = saved_document(&source, edit, Some(part));
            assert_eq!(saved, saved_document(&source, edit, None));
            assert_eq!(saved.contains("<w:t>added</w:t>"), insert);
        }
    }

    #[test]
    fn adding_or_removing_placeholder_only_rows_writes_the_whole_part() {
        let source = single_column_merged_table_body();
        for insert in [true, false] {
            let edit = |content: &mut Vec<serde_json::Value>| {
                without_continued_content(content);
                let rows = content[1]["rows"].as_array_mut().unwrap();
                if insert {
                    rows.push(rows[1].clone());
                } else {
                    rows.remove(1);
                }
                let new = new_paragraph(&content[0], "00000010", "new");
                content.insert(1, new);
            };
            let saved = saved_document(
                &source,
                edit,
                Some(json!({
                    "paragraphs": [0, 1, 3],
                    "inserted": [S13SpliceAnchor::After(0)],
                })),
            );
            assert_eq!(saved, saved_document(&source, edit, None));
            assert_eq!(saved.matches("<w:tr>").count(), if insert { 3 } else { 1 });
        }
    }

    #[test]
    fn placeholder_only_row_cell_or_paragraph_count_changes_write_the_whole_part() {
        let source = single_column_merged_table_body().replace(
            "<w:p w14:paraId=\"0000000C\"/>",
            "<w:p w14:paraId=\"0000000C\"/><w:p w14:paraId=\"0000000E\"/>",
        );
        for add_cell in [false, true] {
            let edit = |content: &mut Vec<serde_json::Value>| {
                without_continued_content(content);
                if add_cell {
                    let cells = content[1]["rows"][1]["cells"].as_array_mut().unwrap();
                    cells.push(cells[0].clone());
                }
                let new = new_paragraph(&content[0], "00000010", "new");
                content.insert(1, new);
            };
            let saved = saved_document(
                &source,
                edit,
                Some(json!({
                    "paragraphs": [0, 1, 4],
                    "inserted": [S13SpliceAnchor::After(0)],
                })),
            );
            assert_eq!(saved, saved_document(&source, edit, None));
            assert_eq!(
                saved.matches("<w:tc>").count(),
                if add_cell { 3 } else { 2 }
            );
        }
    }

    #[test]
    fn paragraphs_and_rows_inserted_around_cells_without_content_splice() {
        let source = merged_table_body();
        let saved = saved_document(
            &source,
            |content| {
                without_continued_content(content);
                let new = new_paragraph(&content[0], "00000010", "new");
                content.insert(1, new);
            },
            Some(json!({
                "paragraphs": [0, 1, 2, 4, 5],
                "inserted": [S13SpliceAnchor::After(0)],
            })),
        );
        let (head, tail) = source.split_at(source.find("<w:tbl>").unwrap());
        assert!(saved.starts_with(head) && saved.ends_with(tail), "{saved}");
        let new = &saved[head.len()..saved.len() - tail.len()];
        assert!(new.starts_with("<w:p ") && new.ends_with("</w:p>"), "{new}");
        assert!(new.contains("<w:t>new</w:t>"), "{new}");

        let saved = saved_document(
            &source,
            |content| {
                without_continued_content(content);
                let mut row = content[1]["rows"][1].clone();
                let added = new_paragraph(&row["cells"][1]["content"][0], "00000010", "added");
                row["cells"][1]["content"][0] = added;
                content[1]["rows"].as_array_mut().unwrap().insert(1, row);
            },
            Some(json!({
                "paragraphs": [0, 1, 2, 4, 5],
                "inserted": [S13SpliceAnchor::After(2)],
            })),
        );
        let (head, tail) = source.split_at(source.find("</w:tr>").unwrap() + "</w:tr>".len());
        assert!(saved.starts_with(head) && saved.ends_with(tail), "{saved}");
        let row = &saved[head.len()..saved.len() - tail.len()];
        assert!(
            row.starts_with("<w:tr") && row.ends_with("</w:tr>"),
            "{row}"
        );
        assert!(
            row.contains("<w:p/>") && row.contains("<w:t>added</w:t>"),
            "{row}"
        );
    }

    #[test]
    fn removing_a_row_with_a_paragraph_the_model_does_not_read_writes_the_whole_part() {
        let source = merged_table_body();
        let remove_row = |content: &mut Vec<serde_json::Value>| {
            content[1]["rows"].as_array_mut().unwrap().remove(1);
        };
        let spliced = saved_document(
            &source,
            remove_row,
            Some(json!({"paragraphs": [0, 1, 2, 5], "removed": [4]})),
        );
        assert_eq!(spliced, saved_document(&source, remove_row, None));
        assert!(!spliced.contains("below"), "{spliced}");
    }

    #[test]
    fn refuses_insertions_next_to_range_markers_or_away_from_their_anchor() {
        let first = "<w:p w14:paraId=\"0000000A\"><w:r><w:t>a</w:t></w:r></w:p>";
        let second = "<w:p w14:paraId=\"0000000B\"><w:r><w:t>b</w:t></w:r></w:p>";
        let new = "<w:p w14:paraId=\"0000000C\"><w:r><w:t>new</w:t></w:r></w:p>";
        let after_first = Edit {
            inserted: &[S13SpliceAnchor::After(0)],
            ..Edit::default()
        };
        let marked = body(&format!(
            "<w:bookmarkStart w:id=\"1\" w:name=\"b\"/>{first}<w:bookmarkEnd w:id=\"1\"/>{second}"
        ));
        assert_eq!(
            splice_written(
                &marked,
                &[(Some(0), first), (None, new), (Some(1), second)],
                after_first,
                &[],
                ("", "")
            ),
            None
        );
        let plain = body(&format!("{first}{second}"));
        let after_first = || Edit {
            inserted: &[S13SpliceAnchor::After(0)],
            ..Edit::default()
        };
        assert!(
            splice_written(
                &plain,
                &[(Some(0), first), (None, new), (Some(1), second)],
                after_first(),
                &[],
                ("", "")
            )
            .is_some()
        );
        assert_eq!(
            splice_written(
                &plain,
                &[(Some(0), first), (Some(1), second), (None, new)],
                after_first(),
                &[],
                ("", "")
            ),
            None
        );
        assert_eq!(
            splice_written(
                &plain,
                &[(Some(0), first), (None, new), (Some(1), second)],
                Edit::default(),
                &[],
                ("", "")
            ),
            None
        );
    }

    #[test]
    fn refuses_removing_or_inserting_a_paragraph_with_one_end_of_a_range() {
        let opening =
            "<w:p><w:bookmarkStart w:id=\"1\" w:name=\"b\"/><w:r><w:t>a</w:t></w:r></w:p>";
        let middle = "<w:p><w:r><w:t>b</w:t></w:r></w:p>";
        let closing = "<w:p><w:r><w:t>c</w:t></w:r><w:bookmarkEnd w:id=\"1\"/></w:p>";
        let source = body(&format!("{opening}{middle}{closing}"));
        let removed = Edit {
            removed: &[0],
            ..Edit::default()
        };
        assert_eq!(
            splice_written(
                &source,
                &[(Some(1), middle), (Some(2), closing)],
                removed,
                &[],
                ("", "")
            ),
            None
        );
        let removed = Edit {
            removed: &[1],
            ..Edit::default()
        };
        assert_eq!(
            splice_written(
                &source,
                &[(Some(0), opening), (Some(2), closing)],
                removed,
                &[],
                ("", "")
            ),
            Some(body(&format!("{opening}{closing}")))
        );
        let new = "<w:p><w:bookmarkStart w:id=\"2\" w:name=\"n\"/><w:r><w:t>new</w:t></w:r></w:p>";
        let inserted = Edit {
            inserted: &[S13SpliceAnchor::After(1)],
            ..Edit::default()
        };
        assert_eq!(
            splice_written(
                &source,
                &[
                    (Some(0), opening),
                    (Some(1), middle),
                    (None, new),
                    (Some(2), closing)
                ],
                inserted,
                &[],
                ("", "")
            ),
            None
        );
    }

    #[test]
    fn keeps_unchanged_paragraphs_and_rewrites_changed_ones() {
        let source = concat!(
            "<?xml version=\"1.0\"?>\n<w:document xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\" xmlns:w14=\"http://schemas.microsoft.com/office/word/2010/wordml\"><w:body>",
            "<w:p w14:paraId=\"0000000A\" w:rsidR=\"00AB12CD\"><w:r><w:t>keep</w:t></w:r></w:p>",
            "\n  ",
            "<w:p w14:paraId=\"0000000B\"><w:r><w:t>old</w:t></w:r></w:p>",
            "<w:sectPr><w:pgSz w:w=\"11906\"/></w:sectPr></w:body></w:document>"
        );
        let new = "<w:p w14:paraId=\"0000000B\"><w:r><w:t>new</w:t></w:r></w:p>";
        let spliced = splice(
            source,
            &[
                (
                    0,
                    "<w:p w14:paraId=\"0000000A\"><w:pPr><w:spacing w:after=\"200\"/></w:pPr><w:r><w:t>keep</w:t></w:r></w:p>",
                ),
                (1, new),
            ],
            &[1],
        );
        assert_eq!(
            spliced,
            Some(source.replace(
                "<w:p w14:paraId=\"0000000B\"><w:r><w:t>old</w:t></w:r></w:p>",
                new
            ))
        );
    }

    #[test]
    fn rewrites_an_unchanged_paragraph_whose_comment_markers_moved() {
        let source = format!(
            "{ROOT}<w:body><w:p w14:paraId=\"0000000A\"><w:commentRangeStart w:id=\"1\"/><w:r><w:t>ab</w:t></w:r><w:commentRangeEnd w:id=\"1\"/></w:p></w:body></w:document>"
        );
        let moved = "<w:p w14:paraId=\"0000000A\"><w:r><w:t>a</w:t></w:r><w:commentRangeStart w:id=\"1\"/><w:r><w:t>b</w:t></w:r><w:commentRangeEnd w:id=\"1\"/></w:p>";
        assert!(splice(&source, &[(0, moved)], &[]).is_some_and(|xml| xml.contains(moved)));
        let same = "<w:p w14:paraId=\"0000000A\"><w:commentRangeStart w:id=\"1\"/><w:r><w:t>a</w:t><w:t>b</w:t></w:r><w:commentRangeEnd w:id=\"1\"/></w:p>";
        assert_eq!(splice(&source, &[(0, same)], &[]), Some(source.clone()));
    }

    #[test]
    fn rewrites_an_unchanged_paragraph_whose_revision_or_note_ids_differ() {
        let root = ROOT.replace(
            "<w:document ",
            "<w:document xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\" ",
        );
        let source = format!(
            "{root}<w:body><w:p><w:hyperlink r:id=\"rId4\"><w:ins w:id=\"7\"><w:r><w:t>x</w:t></w:r></w:ins></w:hyperlink><w:r><w:footnoteReference w:id=\"2\"/></w:r></w:p></w:body></w:document>"
        );
        for other in [
            "<w:p><w:hyperlink r:id=\"rId4\"><w:ins w:id=\"8\"><w:r><w:t>x</w:t></w:r></w:ins></w:hyperlink><w:r><w:footnoteReference w:id=\"2\"/></w:r></w:p>",
            "<w:p><w:hyperlink r:id=\"rId4\"><w:ins w:id=\"7\"><w:r><w:t>x</w:t></w:r></w:ins></w:hyperlink><w:r><w:footnoteReference w:id=\"3\"/></w:r></w:p>",
        ] {
            let spliced = splice(&source, &[(0, other)], &[]);
            assert!(spliced.is_some_and(|xml| xml.contains(other)), "{other}");
        }
        let same = "<w:p><w:hyperlink r:id=\"rId9\"><w:ins w:id=\"7\"><w:r><w:t>x</w:t></w:r></w:ins></w:hyperlink><w:r><w:footnoteReference w:id=\"2\"/></w:r></w:p>";
        assert_eq!(splice(&source, &[(0, same)], &[]), Some(source.clone()));
        let fewer = "<w:p><w:hyperlink r:id=\"rId4\"><w:r><w:t>x</w:t></w:r></w:hyperlink></w:p>";
        assert_eq!(splice(&source, &[(0, fewer)], &[]), Some(source.clone()));
    }

    #[test]
    fn patches_written_and_assigned_paragraph_ids_into_kept_paragraphs() {
        let source = format!(
            "{ROOT}<w:body><w:p><w:r><w:t>a</w:t></w:r></w:p><w:p><w:r><w:t>b</w:t></w:r></w:p></w:body></w:document>"
        );
        let patch = |source: &str| {
            splice_around(
                source,
                &[(
                    0,
                    "<w:p w14:paraId=\"1234ABCD\"><w:r><w:t>a</w:t></w:r></w:p>",
                )],
                &[],
                &[(1, "0BCD1234")],
                ("", "<w:p><w:r><w:t>b</w:t></w:r></w:p>"),
            )
        };
        let spliced = patch(&source).expect("spliced");
        let occurrences = paragraph_occurrences(&spliced).expect("occurrences");
        assert_eq!(occurrences[0].para_id.as_deref(), Some("1234ABCD"));
        assert_eq!(occurrences[1].para_id.as_deref(), Some("0BCD1234"));
        assert!(spliced.starts_with(ROOT));
        assert!(spliced.contains("<w:r><w:t>a</w:t></w:r></w:p><w:p "));

        let undeclared = source.replace(ROOT, W_ONLY_ROOT);
        assert_eq!(patch(&undeclared), None);
    }

    #[test]
    fn patches_the_root_when_a_rewritten_paragraph_needs_a_paragraph_id_declaration() {
        let source = concat!(
            "<w:document xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\"><w:body>",
            "<w:p><w:r><w:t>a</w:t></w:r></w:p>",
            "</w:body></w:document>"
        );
        let rewritten = "<w:p w14:paraId=\"1234ABCD\"><w:r><w:t>new</w:t></w:r></w:p>";
        assert_eq!(
            splice(source, &[(0, rewritten)], &[0]),
            Some(
                source
                    .replace(W_ONLY_ROOT, &patched_w_only_root())
                    .replace("<w:p><w:r><w:t>a</w:t></w:r></w:p>", rewritten)
            )
        );
        assert_eq!(
            splice(source, &[(0, "<w:p><w:r><w:t>new</w:t></w:r></w:p>")], &[0]),
            Some(source.replace("<w:t>a</w:t>", "<w:t>new</w:t>"))
        );
    }

    #[test]
    fn a_prefix_declared_on_one_element_does_not_cover_its_siblings() {
        assert_eq!(
            required_prefixes("<w:p><w:r><a:x xmlns:a=\"urn:a\"><a:y/></a:x><a:z/></w:r></w:p>"),
            Some(BTreeSet::from(["a", "w"]))
        );
        assert_eq!(
            required_prefixes("<w:p xmlns:w=\"urn:w\"><w:r xml:space=\"preserve\"/></w:p>"),
            Some(BTreeSet::new())
        );
    }

    #[test]
    fn a_comment_range_moved_across_tabs_rewrites_the_paragraph() {
        let source = format!(
            "{ROOT}<w:body><w:p><w:commentRangeStart w:id=\"1\"/><w:r><w:tab/><w:tab/></w:r><w:commentRangeEnd w:id=\"1\"/><w:r><w:commentReference w:id=\"1\"/></w:r></w:p></w:body></w:document>"
        );
        let narrowed = "<w:p><w:r><w:tab/></w:r><w:commentRangeStart w:id=\"1\"/><w:r><w:tab/></w:r><w:commentRangeEnd w:id=\"1\"/><w:r><w:commentReference w:id=\"1\"/></w:r></w:p>";
        assert!(splice(&source, &[(0, narrowed)], &[]).is_some_and(|xml| xml.contains(narrowed)));
        let tabs = "<w:p><w:pPr><w:tabs><w:tab w:val=\"left\" w:pos=\"720\"/></w:tabs></w:pPr><w:commentRangeStart w:id=\"1\"/><w:r><w:tab/><w:tab/></w:r><w:commentRangeEnd w:id=\"1\"/><w:r><w:commentReference w:id=\"1\"/></w:r></w:p>";
        assert_eq!(splice(&source, &[(0, tabs)], &[]), Some(source.clone()));
    }

    #[test]
    fn a_paragraph_with_an_equation_is_not_kept_stale_or_rewritten() {
        let math = "<m:oMath xmlns:m=\"http://schemas.openxmlformats.org/officeDocument/2006/math\"><m:r><m:t>x</m:t></m:r></m:oMath>";
        let source = format!(
            "{ROOT}<w:body><w:p><w:commentRangeStart w:id=\"1\"/>{math}{math}<w:commentRangeEnd w:id=\"1\"/><w:r><w:commentReference w:id=\"1\"/></w:r></w:p></w:body></w:document>"
        );
        let narrowed = format!(
            "<w:p>{math}<w:commentRangeStart w:id=\"1\"/>{math}<w:commentRangeEnd w:id=\"1\"/><w:r><w:commentReference w:id=\"1\"/></w:r></w:p>"
        );
        assert_eq!(splice(&source, &[(0, narrowed.as_str())], &[]), None);
    }

    #[test]
    fn refuses_rewriting_a_paragraph_with_field_characters_however_they_are_spelled() {
        let opening = "<w:p><w:r><w:fldChar w:fldCharType=\"begin\"/></w:r><w:r><w:instrText xml:space=\"preserve\"> TOC </w:instrText></w:r><w:r><w:fldChar w:fldCharType=\"separate\"/></w:r><w:r><w:t>First</w:t></w:r></w:p>";
        let closing = "<w:p><w:r><w:t>Second</w:t></w:r><w:r><w:fldChar w:fldCharType=\"e&#110;d\"/></w:r></w:p>";
        let after = "<w:p><w:r><w:t>After</w:t></w:r></w:p>";
        let source = format!("{ROOT}<w:body>{opening}{closing}{after}</w:body></w:document>");
        let written_opening = "<w:p><w:r><w:fldChar w:fldCharType=\"begin\"/></w:r><w:r><w:instrText xml:space=\"preserve\"> TOC </w:instrText></w:r><w:r><w:fldChar w:fldCharType=\"separate\"/></w:r><w:r><w:t>First</w:t></w:r><w:r><w:fldChar w:fldCharType=\"end\"/></w:r></w:p>";
        let written_closing = "<w:p><w:r><w:t>Second, edited</w:t></w:r></w:p>";
        assert_eq!(
            splice(
                &source,
                &[(0, written_opening), (1, written_closing), (2, after)],
                &[1]
            ),
            None
        );
        let aliased = source
            .replacen(
                "<w:body>",
                "<w:body xmlns:x=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\">",
                1,
            )
            .replace("<w:fldChar", "<x:fldChar");
        let edited_after = "<w:p><w:r><w:t>After, edited</w:t></w:r></w:p>";
        assert_eq!(
            splice(
                &aliased,
                &[
                    (0, written_opening),
                    (1, written_closing),
                    (2, edited_after)
                ],
                &[2]
            ),
            None
        );
        let field = "<w:p><w:r><w:fldChar w:fldCharType=\"begin\"/></w:r><w:r><w:instrText xml:space=\"preserve\"> PAGE </w:instrText></w:r><w:r><w:fldChar w:fldCharType=\"end\"/></w:r></w:p>";
        let source = format!("{ROOT}<w:body>{field}</w:body></w:document>");
        assert_eq!(splice(&source, &[(0, field)], &[0]), None);
    }

    #[test]
    fn ranges_across_paragraphs_keep_their_ends() {
        let opening =
            "<w:p><w:bookmarkStart w:id=\"1\" w:name=\"a\"/><w:r><w:t>a</w:t></w:r></w:p>";
        let inside = "<w:p><w:r><w:t>b</w:t></w:r></w:p>";
        let closing = "<w:p><w:r><w:t>c</w:t></w:r><w:bookmarkEnd w:id=\"1\"/></w:p>";
        let source = format!("{ROOT}<w:body>{opening}{inside}{closing}</w:body></w:document>");
        let rewrite = |xml: &str| xml.replace("</w:p>", "<w:r/></w:p>");
        let edited_inside = "<w:p><w:bookmarkStart w:id=\"2\" w:name=\"b\"/><w:r><w:t>b, edited</w:t></w:r><w:bookmarkEnd w:id=\"2\"/></w:p>";
        assert_eq!(
            splice(
                &source,
                &[
                    (0, rewrite(opening).as_str()),
                    (1, edited_inside),
                    (2, rewrite(closing).as_str())
                ],
                &[1]
            ),
            Some(format!(
                "{ROOT}<w:body>{opening}{edited_inside}{closing}</w:body></w:document>"
            ))
        );
        let edited_closing =
            "<w:p><w:r><w:t>c, edited</w:t></w:r><w:bookmarkEnd w:id=\"1\"/></w:p>";
        assert_eq!(
            splice(
                &source,
                &[(0, opening), (1, inside), (2, edited_closing)],
                &[2]
            ),
            Some(format!(
                "{ROOT}<w:body>{opening}{inside}{edited_closing}</w:body></w:document>"
            ))
        );
        let crossed = "<w:p><w:bookmarkEnd w:id=\"3\"/><w:r><w:t>b</w:t></w:r><w:bookmarkStart w:id=\"3\" w:name=\"c\"/></w:p>";
        assert_eq!(
            splice(&source, &[(0, opening), (1, crossed), (2, closing)], &[1]),
            None
        );
    }

    /// Splices an edit to either end of a pair spanning two paragraphs, `start` in the first and
    /// `end` in the second, written with the marker kept or dropped, giving each result beside
    /// the part with only that paragraph replaced. Asserts first that an edit to a third paragraph
    /// keeps both ends.
    fn rewrite_one_end(start: &str, end: &str, keep: bool) -> [(Option<String>, String); 2] {
        let opening = format!("<w:p>{start}<w:r><w:t>a</w:t></w:r></w:p>");
        let closing = format!("<w:p><w:r><w:t>b</w:t></w:r>{end}</w:p>");
        let outside = "<w:p><w:r><w:t>c</w:t></w:r></w:p>";
        let part = |paragraphs: [&str; 3]| body(&paragraphs.concat());
        let source = part([&opening, &closing, outside]);
        let edited = |xml: &str| xml.replace("</w:t>", ", edited</w:t>");
        let outside_edited = edited(outside);
        assert_eq!(
            splice(
                &source,
                &[(0, &opening), (1, &closing), (2, &outside_edited)],
                &[2]
            ),
            Some(part([&opening, &closing, &outside_edited])),
            "{start} {end}"
        );
        let drop = |xml: String, marker: &str| if keep { xml } else { xml.replace(marker, "") };
        let opening_edited = drop(edited(&opening), start);
        let closing_edited = drop(edited(&closing), end);
        [
            (
                splice(
                    &source,
                    &[(0, &opening_edited), (1, &closing), (2, outside)],
                    &[0],
                ),
                part([&opening_edited, &closing, outside]),
            ),
            (
                splice(
                    &source,
                    &[(0, &opening), (1, &closing_edited), (2, outside)],
                    &[1],
                ),
                part([&opening, &closing_edited, outside]),
            ),
        ]
    }

    /// Asserts that a paragraph holding `start` or `end` of a pair spanning two paragraphs is
    /// rewritten alone when its written XML keeps the marker, and not when it drops it.
    fn rewrites_one_end_only_with_its_marker(start: &str, end: &str) {
        for (spliced, part) in rewrite_one_end(start, end, true) {
            assert_eq!(spliced, Some(part), "{start} {end}");
        }
        for (spliced, _) in rewrite_one_end(start, end, false) {
            assert_eq!(spliced, None, "{start} {end}");
        }
    }

    #[test]
    fn a_bookmark_across_paragraphs_is_rewritten_at_one_end_that_keeps_it() {
        rewrites_one_end_only_with_its_marker(
            "<w:bookmarkStart w:id=\"1\" w:name=\"b\"/>",
            "<w:bookmarkEnd w:id=\"1\"/>",
        );
    }

    #[test]
    fn a_comment_range_across_paragraphs_is_rewritten_at_one_end_that_keeps_it() {
        rewrites_one_end_only_with_its_marker(
            "<w:commentRangeStart w:id=\"1\"/>",
            "<w:commentRangeEnd w:id=\"1\"/><w:r><w:commentReference w:id=\"1\"/></w:r>",
        );
    }

    #[test]
    fn a_permission_across_paragraphs_is_rewritten_at_one_end_that_keeps_it() {
        rewrites_one_end_only_with_its_marker(
            "<w:permStart w:id=\"1\" w:edGrp=\"everyone\"/>",
            "<w:permEnd w:id=\"1\"/>",
        );
    }

    #[test]
    fn a_custom_xml_range_across_paragraphs_is_rewritten_at_one_end_that_keeps_it() {
        rewrites_one_end_only_with_its_marker(
            "<w:customXmlInsRangeStart w:id=\"1\" w:author=\"A\"/>",
            "<w:customXmlInsRangeEnd w:id=\"1\"/>",
        );
    }

    #[test]
    fn a_new_local_pair_cannot_duplicate_a_range_in_retained_tracked_changes() {
        let opening = "<w:p><w:bookmarkStart w:id=\"1\" w:name=\"crossing\"/><w:ins w:id=\"3\" w:author=\"A\"><w:bookmarkStart w:id=\"2\" w:name=\"retained\"/><w:r><w:t>a</w:t></w:r><w:bookmarkEnd w:id=\"2\"/></w:ins></w:p>";
        let written_opening = "<w:p><w:bookmarkStart w:id=\"1\" w:name=\"crossing\"/><w:ins w:id=\"3\" w:author=\"A\"><w:r><w:t>a</w:t></w:r></w:ins></w:p>";
        let closing = "<w:p><w:r><w:t>b</w:t></w:r><w:bookmarkEnd w:id=\"1\"/></w:p>";
        let rebuilt = "<w:p><w:bookmarkStart w:id=\"2\" w:name=\"added\"/><w:r><w:t>b, edited</w:t></w:r><w:bookmarkEnd w:id=\"2\"/><w:bookmarkEnd w:id=\"1\"/></w:p>";
        assert_whole_part_fallback(
            &body(&format!("{opening}{closing}")),
            &[(Some(0), written_opening), (Some(1), rebuilt)],
            Edit {
                changed: &[1],
                ..Edit::default()
            },
        );
        let source = body(&format!("{opening}{closing}"));
        assert_eq!(splice(&source, &[(1, rebuilt)], &[1]), None);
        assert_eq!(
            splice(&source, &[(0, written_opening), (1, rebuilt)], &[0, 1]),
            Some(body(&format!("{written_opening}{rebuilt}")))
        );
    }

    #[test]
    fn an_added_same_id_pair_cannot_cancel_out_of_a_crossing_range() {
        for (start, end) in [
            (
                "<w:bookmarkStart w:id=\"1\" w:name=\"crossing\"/>",
                "<w:bookmarkEnd w:id=\"1\"/>",
            ),
            ("<w:permStart w:id=\"1\"/>", "<w:permEnd w:id=\"1\"/>"),
            (
                "<w:commentRangeStart w:id=\"1\"/>",
                "<w:commentRangeEnd w:id=\"1\"/>",
            ),
            (
                "<w:customXmlInsRangeStart w:id=\"1\"/>",
                "<w:customXmlInsRangeEnd w:id=\"1\"/>",
            ),
        ] {
            let opening = format!("<w:p>{start}<w:r><w:t>a</w:t></w:r></w:p>");
            let closing = format!("<w:p><w:r><w:t>b</w:t></w:r>{end}</w:p>");
            let source = body(&format!("{opening}{closing}"));
            for (ordinal, original) in [(0, &opening), (1, &closing)] {
                let rebuilt = original.replace("</w:p>", &format!("{start}{end}</w:p>"));
                assert_eq!(keeps_ranges(original, &rebuilt), Some(false), "{rebuilt}");
                let written = if ordinal == 0 {
                    [(0, rebuilt.as_str()), (1, closing.as_str())]
                } else {
                    [(0, opening.as_str()), (1, rebuilt.as_str())]
                };
                assert_eq!(splice(&source, &written, &[ordinal]), None, "{rebuilt}");
            }
        }
    }

    #[test]
    fn a_proofing_mark_across_paragraphs_does_not_hold_back_rewriting_one_end() {
        for keep in [true, false] {
            for (spliced, part) in rewrite_one_end(
                "<w:proofErr w:type=\"spellStart\"/>",
                "<w:proofErr w:type=\"spellEnd\"/>",
                keep,
            ) {
                assert_eq!(spliced, Some(part), "{keep}");
            }
        }
    }

    #[test]
    fn a_move_is_not_rewritten_at_either_half() {
        for (spliced, _) in rewrite_one_end(
            "<w:moveFromRangeStart w:id=\"0\" w:name=\"m\" w:author=\"A\"/><w:moveFrom w:id=\"1\" w:author=\"A\"><w:r><w:delText>x</w:delText></w:r></w:moveFrom><w:moveFromRangeEnd w:id=\"0\"/>",
            "<w:moveToRangeStart w:id=\"2\" w:name=\"m\" w:author=\"A\"/><w:moveTo w:id=\"3\" w:author=\"A\"><w:r><w:t>x</w:t></w:r></w:moveTo><w:moveToRangeEnd w:id=\"2\"/>",
            true,
        ) {
            assert_eq!(spliced, None);
        }
        let whole_move = "<w:p><w:moveFrom w:id=\"1\" w:author=\"A\"><w:r><w:delText>x</w:delText></w:r></w:moveFrom><w:moveTo w:id=\"3\" w:author=\"A\"><w:r><w:t>x</w:t></w:r></w:moveTo></w:p>";
        let source = format!("{ROOT}<w:body>{whole_move}</w:body></w:document>");
        let rewritten = "<w:p><w:del w:id=\"1\" w:author=\"A\"><w:r><w:delText>x</w:delText></w:r></w:del><w:ins w:id=\"3\" w:author=\"A\"><w:r><w:t>x, edited</w:t></w:r></w:ins></w:p>";
        assert_eq!(splice(&source, &[(0, rewritten)], &[0]), None);
    }

    #[test]
    fn an_end_written_with_another_id_is_not_rewritten_alone() {
        let opening =
            "<w:p><w:bookmarkStart w:id=\"1\" w:name=\"b\"/><w:r><w:t>a</w:t></w:r></w:p>";
        let closing = "<w:p><w:r><w:t>b</w:t></w:r><w:bookmarkEnd w:id=\"1\"/></w:p>";
        let renumbered = "<w:p><w:r><w:t>b, edited</w:t></w:r><w:bookmarkEnd w:id=\"2\"/></w:p>";
        assert_eq!(
            splice(
                &body(&format!("{opening}{closing}")),
                &[(0, opening), (1, renumbered)],
                &[1]
            ),
            None
        );
    }

    #[test]
    fn cell_paragraphs_holding_one_end_of_a_range_are_rewritten_alone() {
        let table = |cell: &str| {
            format!(
                "<w:tbl><w:tblGrid><w:gridCol w:w=\"2000\"/></w:tblGrid><w:tr><w:tc>{cell}</w:tc></w:tr></w:tbl>"
            )
        };
        let opening = "<w:p w14:paraId=\"00000001\"><w:bookmarkStart w:id=\"0\" w:name=\"mark\"/><w:r><w:t>Intro</w:t></w:r></w:p>";
        let marked = "<w:p w14:paraId=\"00000002\"><w:r><w:t>Cell</w:t></w:r><w:bookmarkEnd w:id=\"0\"/></w:p>";
        let noted = "<w:p w14:paraId=\"00000003\"><w:commentRangeStart w:id=\"0\"/><w:r><w:t>Note</w:t></w:r></w:p>";
        let closing = "<w:p w14:paraId=\"00000004\"><w:r><w:t>Tail</w:t></w:r><w:commentRangeEnd w:id=\"0\"/><w:r><w:commentReference w:id=\"0\"/></w:r></w:p>";
        let edited = "<w:p w14:paraId=\"00000002\"><w:bookmarkStart w:id=\"7\" w:name=\"own\"/><w:r><w:t>Cell, edited</w:t></w:r><w:bookmarkEnd w:id=\"7\"/><w:bookmarkEnd w:id=\"0\"/></w:p>";
        let moved = "<w:p w14:paraId=\"00000003\"><w:r><w:t>No</w:t></w:r><w:commentRangeStart w:id=\"0\"/><w:r><w:t>te</w:t></w:r></w:p>";
        let part = |cell: &str| body(&format!("{opening}{}{closing}", table(cell)));
        assert_eq!(
            splice(
                &part(&format!("{marked}{noted}")),
                &[(0, opening), (1, edited), (2, moved), (3, closing)],
                &[1]
            ),
            Some(part(&format!("{edited}{moved}")))
        );
    }

    #[test]
    fn dropping_an_unpaired_proofing_mark_keeps_every_range_marker() {
        let first = "<w:p><w:bookmarkStart w:id=\"0\" w:name=\"mark\"/><w:permStart w:id=\"5\" w:edGrp=\"everyone\"/><w:commentRangeStart w:id=\"3\"/><w:r><w:t>One</w:t></w:r></w:p>";
        let second = "<w:p><w:r><w:t>Two</w:t></w:r><w:bookmarkEnd w:id=\"0\"/><w:proofErr w:type=\"spellStart\"/><w:r><w:t>teh</w:t></w:r></w:p>";
        let third = "<w:p><w:r><w:t>wrd</w:t></w:r><w:proofErr w:type=\"spellEnd\"/><w:permEnd w:id=\"5\"/><w:commentRangeEnd w:id=\"3\"/><w:r><w:commentReference w:id=\"3\"/></w:r></w:p>";
        let rewritten = "<w:p><w:r><w:t>Two, edited</w:t></w:r><w:bookmarkEnd w:id=\"0\"/><w:r><w:t>the</w:t></w:r></w:p>";
        let source = body(&format!("{first}{second}{third}"));
        let spliced =
            splice(&source, &[(0, first), (1, rewritten), (2, third)], &[1]).expect("spliced");
        assert_eq!(spliced, body(&format!("{first}{rewritten}{third}")));
        crate::parse_docx_s9_wire(&package(&spliced), Default::default()).expect("parses");
        let markers = |xml: &str| -> Vec<String> {
            tags(xml)
                .unwrap()
                .into_iter()
                .filter(|tag| {
                    ["w:bookmark", "w:commentRange", "w:perm"]
                        .iter()
                        .any(|kind| tag.name.starts_with(kind))
                })
                .map(|tag| xml[tag.range].to_owned())
                .collect()
        };
        assert_eq!(markers(&spliced), markers(&source));
    }

    #[test]
    fn saving_an_edit_to_a_cell_paragraph_that_ends_a_bookmark_keeps_every_other_byte() {
        let marked = "<w:p w14:paraId=\"00000002\"><w:r><w:t>Cell</w:t></w:r><w:bookmarkEnd w:id=\"0\"/></w:p>";
        let source = body(&format!(
            "<w:p w14:paraId=\"00000001\"><w:bookmarkStart w:id=\"0\" w:name=\"mark\"/><w:r><w:t>Intro</w:t></w:r></w:p><w:tbl><w:tblGrid><w:gridCol w:w=\"2000\"/></w:tblGrid><w:tr><w:tc>{marked}</w:tc></w:tr></w:tbl><w:p w14:paraId=\"00000003\"><w:r><w:t>Tail</w:t></w:r></w:p><w:sectPr/>"
        ));
        let original = package(&source);
        let parsed = crate::parse_docx_s9_wire(
            &original,
            crate::S9ParseOptions {
                source_ordinals: true,
                ..Default::default()
            },
        )
        .unwrap();
        let content = serde_json::to_string(&parsed.document.package.document.content)
            .unwrap()
            .replace("\"Cell\"", "\"Cell, edited\"");
        let request: S13SaveRequest = serde_json::from_value(json!({
            "determinism": {"seed": "0".repeat(64), "now": "2000-01-01T00:00:00.000Z"},
            "document": {"content": serde_json::from_str::<serde_json::Value>(&content).unwrap()},
            "options": {"updateModifiedDate": false},
            "paragraphIds": {"splicedParts": [{
                "part": "word/document.xml",
                "sha256": format!("{:x}", Sha256::digest(source.as_bytes())),
                "paragraphs": [0, 1, 2],
                "changed": [1],
            }]},
        }))
        .unwrap();
        let saved =
            String::from_utf8(document_part(&write_docx_s13(request, &original).unwrap())).unwrap();
        let at = source.find(marked).unwrap();
        let rest = &source[at + marked.len()..];
        assert!(saved.starts_with(&source[..at]), "{saved}");
        assert!(saved.ends_with(rest), "{saved}");
        let written = &saved[at..saved.len() - rest.len()];
        assert!(
            written.contains("Cell, edited") && written.contains("<w:bookmarkEnd w:id=\"0\"/>"),
            "{written}"
        );
    }

    #[test]
    fn a_field_spanning_paragraphs_is_kept_around_an_edit_elsewhere() {
        let opening = "<w:p><w:r><w:fldChar w:fldCharType=\"begin\"/></w:r><w:r><w:instrText xml:space=\"preserve\"> TOC </w:instrText></w:r><w:r><w:fldChar w:fldCharType=\"separate\"/></w:r><w:r><w:t>First entry</w:t></w:r></w:p>";
        let closing = "<w:p><w:r><w:t>Second entry</w:t></w:r><w:r><w:fldChar w:fldCharType=\"end\"/></w:r></w:p>";
        let outside = "<w:p><w:r><w:t>Outside the field</w:t></w:r></w:p>";
        let source = format!("{ROOT}<w:body>{opening}{closing}{outside}</w:body></w:document>");
        let written_opening = "<w:p><w:r><w:fldChar w:fldCharType=\"begin\"/></w:r><w:r><w:instrText xml:space=\"preserve\"> TOC </w:instrText></w:r><w:r><w:fldChar w:fldCharType=\"separate\"/></w:r><w:r><w:t>First entry</w:t></w:r><w:r><w:fldChar w:fldCharType=\"end\"/></w:r></w:p>";
        let written_closing = "<w:p><w:r><w:t>Second entry</w:t></w:r></w:p>";
        assert_eq!(
            splice(
                &source,
                &[(0, written_opening), (1, written_closing), (2, outside)],
                &[]
            ),
            Some(source.clone())
        );
        let edited_outside = "<w:p><w:r><w:t>Outside, edited</w:t></w:r></w:p>";
        assert_eq!(
            splice(
                &source,
                &[
                    (0, written_opening),
                    (1, written_closing),
                    (2, edited_outside)
                ],
                &[2]
            ),
            Some(format!(
                "{ROOT}<w:body>{opening}{closing}{edited_outside}</w:body></w:document>"
            ))
        );
    }

    #[test]
    fn refuses_rewriting_a_paragraph_that_starts_inside_a_field() {
        let opening = "<w:p><w:r><w:fldChar w:fldCharType=\"begin\"/></w:r><w:r><w:instrText xml:space=\"preserve\"> IF 1 = 1 </w:instrText></w:r></w:p>";
        let code =
            "<w:p><w:r><w:instrText xml:space=\"preserve\">\"yes\" </w:instrText></w:r></w:p>";
        let closing = "<w:p><w:r><w:fldChar w:fldCharType=\"separate\"/></w:r><w:r><w:t>yes</w:t></w:r><w:r><w:fldChar w:fldCharType=\"e&#110;d\"/></w:r></w:p>";
        let outside = "<w:p><w:r><w:t>Outside</w:t></w:r></w:p>";
        let source =
            format!("{ROOT}<w:body>{opening}{code}{closing}{outside}</w:body></w:document>");
        let typed = "<w:p><w:r><w:t>typed</w:t></w:r></w:p>";
        assert_eq!(
            splice(
                &source,
                &[(0, opening), (1, typed), (2, closing), (3, outside)],
                &[1]
            ),
            None
        );
        let edited_outside = "<w:p><w:r><w:t>Outside, edited</w:t></w:r></w:p>";
        assert_eq!(
            splice(
                &source,
                &[(0, opening), (1, code), (2, closing), (3, edited_outside)],
                &[3]
            ),
            Some(format!(
                "{ROOT}<w:body>{opening}{code}{closing}{edited_outside}</w:body></w:document>"
            ))
        );
    }

    #[test]
    fn refuses_rewriting_a_paragraph_whose_range_ends_outside_every_paragraph_or_whose_prefix_an_ancestor_rebinds()
     {
        let paragraph = "<w:p><w:r><w:t>a</w:t></w:r><w:bookmarkEnd w:id=\"3\"/></w:p>";
        let source = format!(
            "{ROOT}<w:body><w:bookmarkStart w:id=\"3\" w:name=\"b\"/>{paragraph}</w:body></w:document>"
        );
        let edited = "<w:p><w:r><w:t>b</w:t></w:r><w:bookmarkEnd w:id=\"3\"/></w:p>";
        assert_eq!(
            splice(&source, &[(0, edited)], &[0]),
            Some(source.replace(paragraph, edited))
        );
        let dropped = "<w:p><w:r><w:t>b</w:t></w:r></w:p>";
        assert!(splice(&source, &[(0, dropped)], &[0]).is_none());
        assert_eq!(
            splice(&source, &[(0, paragraph)], &[]),
            Some(source.clone())
        );

        let rebound = format!(
            "{SERIALIZED_ROOT}<w:body xmlns:r=\"urn:other\"><w:p><w:hyperlink xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\" r:id=\"rId1\"><w:r><w:t>Link</w:t></w:r></w:hyperlink></w:p></w:body></w:document>"
        );
        let link = "<w:p><w:hyperlink r:id=\"rId1\"><w:r><w:t>Link, edited</w:t></w:r></w:hyperlink></w:p>";
        assert!(splice(&rebound, &[(0, link)], &[0]).is_none());
    }

    #[test]
    fn writes_the_whole_part_when_a_prefix_the_rewrite_needs_is_not_bound_on_the_root_alone() {
        let rel = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
        let photo = "<w:p><w:r><w:t>Photo</w:t></w:r></w:p>";
        let pictured = "<w:p><w:r><w:t>Photo</w:t></w:r><w:r><w:drawing><a:blip xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\" r:embed=\"rId9\"/></w:drawing></w:r></w:p>";
        let shadowed =
            format!("{ROOT}<w:body xmlns:r=\"urn:unused\">{photo}</w:body></w:document>");
        assert_eq!(splice(&shadowed, &[(0, pictured)], &[0]), None);
        let below_root = shadowed.replace("urn:unused", rel);
        assert_eq!(splice(&below_root, &[(0, pictured)], &[0]), None);
        let rooted = format!(
            "{}<w:body>{photo}</w:body></w:document>",
            ROOT.replace("<w:document ", &format!("<w:document xmlns:r=\"{rel}\" "))
        );
        let rebound = rooted.replace("<w:body>", "<w:body xmlns:r=\"urn:unused\">");
        assert_eq!(splice(&rebound, &[(0, pictured)], &[0]), None);
        assert_eq!(
            splice(&rooted, &[(0, pictured)], &[0]),
            Some(rooted.replace(photo, pictured))
        );
    }

    #[test]
    fn patches_the_root_when_patched_ids_and_a_rewrite_need_a_paragraph_id_declaration() {
        let source = format!(
            "{W_ONLY_ROOT}<w:body><w:p><w:r><w:t>A</w:t></w:r></w:p><w:p><w:r><w:t>B</w:t></w:r></w:p></w:body></w:document>"
        );
        let paragraphs = [
            (0, "<w:p><w:r><w:t>A</w:t></w:r></w:p>"),
            (
                1,
                "<w:p w14:paraId=\"0BCD1234\"><w:r><w:t>B, edited</w:t></w:r></w:p>",
            ),
        ];
        let assigned = [(0, "1234ABCD")];
        let declared = source.replace(W_ONLY_ROOT, ROOT);
        let spliced =
            splice_around(&declared, &paragraphs, &[1], &assigned, ("", "")).expect("spliced");
        assert!(spliced.starts_with(ROOT));
        assert_eq!(spliced.matches("xmlns:w14=").count(), 1);
        let occurrences = paragraph_occurrences(&spliced).expect("occurrences");
        assert_eq!(occurrences[0].para_id.as_deref(), Some("1234ABCD"));
        assert_eq!(occurrences[1].para_id.as_deref(), Some("0BCD1234"));
        assert_eq!(
            splice_around(&source, &paragraphs, &[1], &assigned, ("", "")),
            Some(spliced.replacen(ROOT, &patched_w_only_root(), 1))
        );
    }

    #[test]
    fn refuses_parts_it_cannot_splice() {
        let source = format!(
            "{ROOT}<w:body><w:p><w:r><w:t>a</w:t></w:r></w:p><w:p><w:r><w:t>b</w:t></w:r></w:p></w:body></w:document>"
        );
        assert!(splice(&source, &[(5, "<w:p/>")], &[]).is_none());
        let rebound = source.replace(
            "xmlns:w14=\"http://schemas.microsoft.com/office/word/2010/wordml\"",
            "xmlns:w14=\"urn:other\"",
        );
        let rewritten = "<w:p w14:paraId=\"1234ABCD\"><w:r><w:t>new</w:t></w:r></w:p>";
        assert!(splice(&rebound, &[(0, rewritten)], &[0]).is_none());
        let nested = format!(
            "{ROOT}<w:body><w:p><w:r><w:txbxContent><w:p><w:r><w:t>in</w:t></w:r></w:p></w:txbxContent></w:r></w:p></w:body></w:document>"
        );
        let outer = "<w:p><w:r><w:txbxContent><w:p><w:r><w:t>IN</w:t></w:r></w:p></w:txbxContent></w:r></w:p>";
        assert!(
            splice(
                &nested,
                &[(0, outer), (1, "<w:p><w:r><w:t>IN</w:t></w:r></w:p>")],
                &[0, 1]
            )
            .is_none()
        );
    }

    #[test]
    fn refuses_comment_markers_between_paragraphs_that_the_save_no_longer_writes() {
        let (start, end) = (
            "<w:commentRangeStart w:id=\"1\"/>",
            "<w:commentRangeEnd w:id=\"1\"/>",
        );
        let paragraph = "<w:p><w:r><w:t>a</w:t></w:r></w:p>";
        let source = format!("{ROOT}<w:body>{start}{paragraph}{end}</w:body></w:document>");
        assert!(splice(&source, &[(0, paragraph)], &[]).is_none());
        assert_eq!(
            splice_around(&source, &[(0, paragraph)], &[], &[], (start, end)),
            Some(source.clone())
        );
    }

    #[test]
    fn unchanged_paragraphs_take_added_comment_ranges_into_their_source() {
        let field = "<w:r><w:t xml:space=\"preserve\">Page </w:t><w:fldChar w:fldCharType=\"begin\"/><w:instrText xml:space=\"preserve\"> PAGE </w:instrText><w:fldChar w:fldCharType=\"separate\"/><w:t>1</w:t><w:fldChar w:fldCharType=\"end\"/></w:r>";
        let first = |reply: &str| {
            format!(
                "<w:p w14:paraId=\"00000001\"><w:commentRangeStart w:id=\"1\"/>{reply}{field}</w:p>"
            )
        };
        let second = |reply: &str| {
            format!(
                "<w:p w14:paraId=\"00000002\"><w:r><w:t>Tail</w:t></w:r><w:commentRangeEnd w:id=\"1\"/>{reply}<w:r><w:commentReference w:id=\"1\"/></w:r></w:p>"
            )
        };
        let start = "<w:commentRangeStart w:id=\"2\"/>";
        let end = "<w:commentRangeEnd w:id=\"2\"/><w:r><w:rPr><w:rStyle w:val=\"CommentReference\"/></w:rPr><w:commentReference w:id=\"2\"/></w:r>";
        let written_first = format!(
            "<w:p w14:paraId=\"00000001\"><w:commentRangeStart w:id=\"1\"/>{start}<w:r><w:t xml:space=\"preserve\">Page 1</w:t></w:r></w:p>"
        );
        let written_second = second(end).replace("<w:t>Tail</w:t>", "<w:t>Tail</w:t><w:t/>");
        assert_eq!(
            splice(
                &body(&format!("{}{}", first(""), second(""))),
                &[(0, &written_first), (1, &written_second)],
                &[]
            ),
            Some(body(&format!("{}{}", first(start), second(end))))
        );
    }

    #[test]
    fn a_comment_range_that_follows_no_source_range_rewrites_its_paragraph() {
        let source = "<w:p w14:paraId=\"00000001\"><w:r><w:t>Plain</w:t></w:r></w:p>";
        let written = "<w:p w14:paraId=\"00000001\"><w:commentRangeStart w:id=\"2\"/><w:r><w:t>Plain</w:t></w:r><w:commentRangeEnd w:id=\"2\"/></w:p>";
        assert_eq!(
            splice(&body(source), &[(0, written)], &[]),
            Some(body(written))
        );
    }
}
