//! Story parts written as their source XML with only the paragraphs a save changed re-serialized.

use std::borrow::Cow;
use std::collections::{BTreeMap, BTreeSet, HashMap};
use std::ops::Range;

use crate::paragraph_identity::{
    ParagraphOccurrence, Tag, attribute, paragraph_occurrences, patch_paragraph_ids, tags,
    unescaped,
};

use super::context::RecordedParagraphs;
use super::paragraph_ids::{S13SpliceAnchor, S13SplicedPart};
use super::s13::element_span;

const MC_NAMESPACE: &str = "http://schemas.openxmlformats.org/markup-compatibility/2006";

/// What the XML written for an unchanged paragraph must agree on with its source XML for the
/// source to stand in for it: the same comment markers in order, at the same text offsets, and
/// no revision or note ID the source lacks. Source IDs the written XML lacks are kept, as are
/// relationship IDs: the save keeps every source relationship.
#[derive(Debug, Default, PartialEq, Eq)]
struct Marks {
    comments: Vec<(String, String, usize)>,
    ids: BTreeMap<(&'static str, String), usize>,
}

impl Marks {
    fn admit(&self, written: &Marks) -> bool {
        self.comments == written.comments
            && written
                .ids
                .iter()
                .all(|(id, count)| self.ids.get(id).is_some_and(|source| source >= count))
    }
}

/// Run content that takes a position in the text, besides `w:t` and `w:delText` characters.
const POSITIONED: [&str; 19] = [
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
            "w:ins" | "w:del" | "w:moveFrom" | "w:moveTo" => "revision",
            name if name.ends_with("Change") => "revision",
            "w:footnoteReference" => "footnote",
            "w:endnoteReference" => "endnote",
            _ => continue,
        };
        *marks.ids.entry((kind, id())).or_insert(0) += 1;
    }
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
fn required_prefixes(fragment: &str) -> Option<BTreeSet<&str>> {
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
fn root_tag(xml: &str) -> Option<Tag<'_>> {
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

/// `source`'s root start tag binding every prefix in `prefixes` as `serialized`'s root does,
/// ignorable where it is ignorable there; `None` when `source` binds one of them elsewhere.
fn declared_root(
    source: &str,
    serialized: &str,
    prefixes: &BTreeSet<&str>,
) -> Option<Option<(Range<usize>, String)>> {
    let root = root_tag(source)?;
    let bound = bindings(source, &root);
    let target = root_tag(serialized)?;
    let wanted = bindings(serialized, &target);
    let target_ignorable: Vec<&str> = attribute(&target, "mc:Ignorable")
        .map(|range| serialized[range].split_whitespace().collect())
        .unwrap_or_default();
    let mut declarations = String::new();
    let mut ignorable = Vec::new();
    for prefix in prefixes {
        let uri = wanted.get(prefix)?;
        match bound.get(prefix) {
            Some(existing) if existing.trim() == uri.trim() => {}
            Some(_) => return None,
            None => {
                declarations.push_str(&format!(" xmlns:{prefix}=\"{uri}\""));
                if target_ignorable.contains(prefix) {
                    ignorable.push(*prefix);
                }
            }
        }
    }
    if declarations.is_empty() {
        return Some(None);
    }
    let mut text = source[root.range.clone()].to_owned();
    let base = root.range.start;
    if !ignorable.is_empty() {
        if let Some(range) = attribute(&root, "mc:Ignorable") {
            let value = format!("{} {}", &source[range.clone()], ignorable.join(" "));
            text.replace_range(range.start - base..range.end - base, &value);
        } else {
            match bound.get("mc") {
                Some(uri) if uri.trim() == MC_NAMESPACE => {}
                Some(_) => return None,
                None if prefixes.contains("mc") => {}
                None => declarations.push_str(&format!(" xmlns:mc=\"{MC_NAMESPACE}\"")),
            }
            declarations.push_str(&format!(" mc:Ignorable=\"{}\"", ignorable.join(" ")));
        }
    }
    let close = text.len() - 1 - usize::from(text.ends_with("/>"));
    text.insert_str(close, &declarations);
    Some(Some((root.range, text)))
}

/// How the paragraphs a splice addresses hang together in the source: the paragraphs a field,
/// bookmark, permission, move or comment range spans from its opening paragraph to its closing
/// one, which are kept or rewritten together; the paragraphs holding one end of such a range
/// whose other end lies outside every addressed paragraph, which a splice must keep; and the
/// namespace bindings an element between the root and a paragraph declares. For each paragraph
/// in `anchors`, the tags right before and right after it.
#[derive(Default)]
struct Structure {
    /// Each group as the first and last of its paragraphs in source order, overlapping groups
    /// merged.
    groups: Vec<(u32, u32)>,
    pinned: BTreeSet<u32>,
    bindings: HashMap<u32, Vec<(String, String)>>,
    edges: HashMap<u32, Edges>,
}

/// A tag as `(name, whether it is an end tag)`.
type Edge = (String, bool);

#[derive(Default)]
struct Edges {
    before: Option<Edge>,
    after: Option<Edge>,
}

/// Story containers a paragraph can open or close.
const CONTAINERS: [&str; 7] = [
    "w:body",
    "w:tc",
    "w:hdr",
    "w:ftr",
    "w:footnote",
    "w:endnote",
    "w:sdtContent",
];

impl Edges {
    /// Whether a paragraph goes in at `anchor` exactly where the model has it: next to its
    /// anchor with only a paragraph, table, section or the edge of its container beyond.
    fn admit(&self, anchor: S13SpliceAnchor) -> bool {
        match anchor {
            S13SpliceAnchor::After(_) => match &self.after {
                Some((name, true)) => CONTAINERS.contains(&name.as_str()),
                Some((name, false)) => matches!(name.as_str(), "w:p" | "w:tbl" | "w:sectPr"),
                None => false,
            },
            S13SpliceAnchor::Before(_) => match &self.before {
                Some((name, false)) => CONTAINERS.contains(&name.as_str()),
                Some((name, true)) => name == "w:tcPr",
                None => false,
            },
        }
    }
}

fn structure(
    source: &str,
    spans: &BTreeMap<u32, Range<usize>>,
    anchors: &BTreeSet<u32>,
) -> Option<Structure> {
    let order: Vec<(u32, Range<usize>)> = {
        let mut order: Vec<_> = spans.iter().map(|(o, s)| (*o, s.clone())).collect();
        order.sort_by_key(|(_, span)| span.start);
        order
    };
    let index: HashMap<u32, usize> = order
        .iter()
        .enumerate()
        .map(|(i, (o, _))| (*o, i))
        .collect();
    let mut result = Structure::default();
    let mut intervals: Vec<(usize, usize)> = Vec::new();
    let mut link = |from: Option<u32>, to: Option<u32>, result: &mut Structure| match (from, to) {
        (Some(from), Some(to)) if from != to => {
            let (from, to) = (index[&from], index[&to]);
            intervals.push((from.min(to), from.max(to)));
        }
        (Some(paragraph), None) | (None, Some(paragraph)) => {
            result.pinned.insert(paragraph);
        }
        _ => {}
    };
    let mut cursor = 0;
    let mut fields: Vec<Option<u32>> = Vec::new();
    let mut open: HashMap<(String, String), Option<u32>> = HashMap::new();
    let mut scopes: Vec<Vec<(String, String)>> = Vec::new();
    let mut previous: Option<(&str, bool)> = None;
    for tag in tags(source)? {
        while cursor < order.len() && order[cursor].1.end <= tag.range.start {
            if anchors.contains(&order[cursor].0) {
                result.edges.entry(order[cursor].0).or_default().after =
                    Some((tag.name.to_owned(), tag.end));
            }
            cursor += 1;
        }
        let last = previous.replace((tag.name, tag.end));
        if let Some((ordinal, span)) = order.get(cursor)
            && span.start == tag.range.start
            && anchors.contains(ordinal)
        {
            result.edges.entry(*ordinal).or_default().before =
                last.map(|(name, end)| (name.to_owned(), end));
        }
        let here = order
            .get(cursor)
            .filter(|(_, span)| span.start <= tag.range.start)
            .map(|(ordinal, _)| *ordinal);
        if tag.end {
            scopes.pop();
            continue;
        }
        if let Some((ordinal, span)) = order.get(cursor)
            && span.start == tag.range.start
            && scopes.len() > 1
        {
            let declared: Vec<(String, String)> = scopes[1..].iter().flatten().cloned().collect();
            if !declared.is_empty() {
                result.bindings.insert(*ordinal, declared);
            }
        }
        scopes.push(
            tag.attributes
                .iter()
                .filter_map(|(key, range)| {
                    Some((
                        key.strip_prefix("xmlns:")?.to_owned(),
                        source[range.clone()].trim().to_owned(),
                    ))
                })
                .collect(),
        );
        if tag.empty {
            scopes.pop();
        }
        let id = || {
            attribute(&tag, "w:id")
                .and_then(|range| unescaped(source, range))
                .unwrap_or_default()
        };
        if tag.name == "w:fldChar" {
            let kind = unescaped(source, attribute(&tag, "w:fldCharType")?)?;
            match kind.trim() {
                "begin" => fields.push(here),
                "end" => {
                    let begin = fields.pop()?;
                    link(begin, here, &mut result);
                }
                "separate" => {}
                _ => return None,
            }
            continue;
        }
        let (kind, opens) = match tag.name {
            "w:bookmarkStart" => ("bookmark", true),
            "w:bookmarkEnd" => ("bookmark", false),
            "w:permStart" => ("perm", true),
            "w:permEnd" => ("perm", false),
            name => match (
                name.strip_suffix("RangeStart"),
                name.strip_suffix("RangeEnd"),
            ) {
                (Some(kind), _) => (kind, true),
                (_, Some(kind)) => (kind, false),
                _ => continue,
            },
        };
        let key = (kind.to_owned(), id());
        if opens {
            open.insert(key, here);
        } else if let Some(start) = open.remove(&key) {
            link(start, here, &mut result);
        } else {
            link(None, here, &mut result);
        }
    }
    for start in fields.into_iter().chain(open.into_values()) {
        link(start, None, &mut result);
    }
    intervals.sort_unstable();
    let mut merged: Vec<(usize, usize)> = Vec::new();
    for (low, high) in intervals {
        match merged.last_mut() {
            Some(last) if low <= last.1 => last.1 = last.1.max(high),
            _ => merged.push((low, high)),
        }
    }
    result.groups = merged
        .into_iter()
        .map(|(low, high)| (order[low].0, order[high].0))
        .collect();
    Some(result)
}

fn paragraph_span(xml: &str, occurrence: &ParagraphOccurrence) -> Option<Range<usize>> {
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

/// `source` with the paragraphs `part` addresses rewritten from `recorded` where they changed
/// or their written XML no longer agrees with their source on comments, revisions, notes or
/// relationships, the removed ones dropped and the new ones inserted at their anchors, every
/// other byte kept except paragraph IDs: the written ones, and `assignments` for source
/// paragraphs no model paragraph is written from. `None` when the part cannot be spliced, so
/// the caller writes `serialized` whole.
pub(crate) fn splice_story_part(
    source: &str,
    serialized: &str,
    recorded: &RecordedParagraphs,
    part: &S13SplicedPart,
    assignments: &BTreeMap<u32, String>,
) -> Option<String> {
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
    let mut touched = removed.clone();
    for (&ordinal, xml) in written {
        if changed.contains(&ordinal)
            || !marks(&source[spans[&ordinal].clone()])?.admit(&marks(xml)?)
        {
            touched.insert(ordinal);
        }
    }
    let anchors: BTreeSet<u32> = inserted
        .iter()
        .map(|(anchor, _)| anchor.ordinal())
        .collect();
    let structure = structure(source, &spans, &anchors)?;
    for &(first, last) in &structure.groups {
        if touched.range(first..=last).next().is_some() {
            touched.extend(spans.range(first..=last).map(|(ordinal, _)| *ordinal));
        }
    }
    if touched
        .iter()
        .any(|ordinal| structure.pinned.contains(ordinal))
        || inserted.iter().any(|(anchor, _)| {
            !structure
                .edges
                .get(&anchor.ordinal())
                .is_some_and(|edges| edges.admit(*anchor))
        })
    {
        return None;
    }
    let replaced: BTreeSet<u32> = touched.difference(&removed).copied().collect();
    let serialized_root = root_tag(serialized)?;
    let serialized_bindings = bindings(serialized, &serialized_root);
    let new_paragraphs = replaced
        .iter()
        .map(|ordinal| (*ordinal, written[ordinal].as_str()))
        .chain(
            inserted
                .iter()
                .map(|(anchor, xml)| (anchor.ordinal(), *xml)),
        );
    let mut prefixes = BTreeSet::new();
    for (ordinal, xml) in new_paragraphs {
        let required = required_prefixes(xml)?;
        if let Some(declared) = structure.bindings.get(&ordinal) {
            for prefix in &required {
                if let Some((_, uri)) = declared.iter().rev().find(|(bound, _)| bound == prefix)
                    && serialized_bindings.get(prefix).map(|wanted| wanted.trim())
                        != Some(uri.as_str())
                {
                    return None;
                }
            }
        }
        prefixes.extend(required);
    }
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
        Cow::Owned(patch_paragraph_ids(source, &ids, true)?)
    };
    let occurrences = match &source {
        Cow::Borrowed(_) => occurrences,
        Cow::Owned(patched) => paragraph_occurrences(patched)?,
    };
    let span = |ordinal: u32| paragraph_span(&source, occurrences.get(ordinal as usize)?);
    let mut edits: Vec<(Range<usize>, u8, &str)> = Vec::new();
    for &ordinal in &touched {
        let xml = written.get(&ordinal).map_or("", String::as_str);
        edits.push((span(ordinal)?, 2, xml));
    }
    for (anchor, xml) in &inserted {
        let range = span(anchor.ordinal())?;
        edits.push(match anchor {
            S13SpliceAnchor::After(_) => (range.end..range.end, 0, xml),
            S13SpliceAnchor::Before(_) => (range.start..range.start, 1, xml),
        });
    }
    edits.sort_by_key(|(range, rank, _)| (range.start, *rank));
    let root = declared_root(&source, serialized, &prefixes)?;
    let mut output = String::with_capacity(source.len());
    let mut cursor = 0;
    let root = root
        .as_ref()
        .map(|(range, text)| (range.clone(), 0, text.as_str()));
    for (range, _, text) in root.into_iter().chain(edits) {
        if range.start < cursor {
            return None;
        }
        output.push_str(&source[cursor..range.start]);
        output.push_str(text);
        cursor = range.end;
    }
    output.push_str(&source[cursor..]);
    Some(output)
}

#[cfg(test)]
mod tests {
    use super::*;

    const ROOT: &str = "<w:document xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\" xmlns:w14=\"http://schemas.microsoft.com/office/word/2010/wordml\" xmlns:mc=\"http://schemas.openxmlformats.org/markup-compatibility/2006\" mc:Ignorable=\"w14\">";
    const SERIALIZED_ROOT: &str = "<w:document xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\" xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\" xmlns:w14=\"http://schemas.microsoft.com/office/word/2010/wordml\" xmlns:mc=\"http://schemas.openxmlformats.org/markup-compatibility/2006\" mc:Ignorable=\"w14\">";

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

    #[derive(Default)]
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

    #[test]
    fn inserts_new_paragraphs_at_their_anchors_and_drops_removed_ones() {
        let first =
            "<w:p w14:paraId=\"0000000A\"><w:r><w:t>first half second half</w:t></w:r></w:p>";
        let middle = "<w:p w14:paraId=\"0000000B\"><w:r><w:t>removed</w:t></w:r></w:p>";
        let last =
            "<w:p w14:paraId=\"0000000C\" w:rsidR=\"00AB12CD\"><w:r><w:t>kept</w:t></w:r></w:p>";
        let source = body(&format!("{first}\n{middle}\n<!-- gap -->{last}"));
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
            Some(body(&format!("{split}{new}\n\n<!-- gap -->{last}")))
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
    fn removing_a_paragraph_rewrites_the_paragraphs_its_ranges_reach_and_refuses_pinned_ones() {
        let opening =
            "<w:p><w:bookmarkStart w:id=\"1\" w:name=\"b\"/><w:r><w:t>a</w:t></w:r></w:p>";
        let middle = "<w:p><w:r><w:t>b</w:t></w:r></w:p>";
        let closing = "<w:p><w:r><w:t>c</w:t></w:r><w:bookmarkEnd w:id=\"1\"/></w:p>";
        let outside = "<w:p><w:r><w:t>d</w:t></w:r></w:p>";
        let source = body(&format!("{opening}{middle}{closing}{outside}"));
        let written_middle = "<w:p><w:r><w:t>b</w:t></w:r><w:r/></w:p>";
        let written_closing = "<w:p><w:r><w:t>c</w:t></w:r><w:r/></w:p>";
        let written_outside = "<w:p><w:r><w:t>d</w:t></w:r><w:r/></w:p>";
        let removed = || Edit {
            removed: &[0],
            ..Edit::default()
        };
        assert_eq!(
            splice_written(
                &source,
                &[
                    (Some(1), written_middle),
                    (Some(2), written_closing),
                    (Some(3), written_outside)
                ],
                removed(),
                &[],
                ("", "")
            ),
            Some(body(&format!("{written_middle}{written_closing}{outside}")))
        );
        let pinned = body(&format!(
            "<w:bookmarkStart w:id=\"2\" w:name=\"p\"/><w:p><w:r><w:t>a</w:t></w:r><w:bookmarkEnd w:id=\"2\"/></w:p>{outside}"
        ));
        assert_eq!(
            splice_written(&pinned, &[(Some(1), outside)], removed(), &[], ("", "")),
            None
        );
    }

    #[test]
    fn keeps_unchanged_paragraphs_and_rewrites_changed_ones() {
        let source = concat!(
            "<?xml version=\"1.0\"?>\n<w:document xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\" xmlns:w14=\"http://schemas.microsoft.com/office/word/2010/wordml\"><w:body>",
            "<w:p w14:paraId=\"0000000A\" w:rsidR=\"00AB12CD\"><w:r><w:t>keep</w:t></w:r></w:p>",
            "<!-- gap -->",
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
        let source = format!(
            "{ROOT}<w:body><w:p><w:hyperlink r:id=\"rId4\"><w:moveTo w:id=\"7\"><w:r><w:t>x</w:t></w:r></w:moveTo></w:hyperlink><w:r><w:footnoteReference w:id=\"2\"/></w:r></w:p></w:body></w:document>"
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
        let source = concat!(
            "<w:document xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\"><w:body>",
            "<w:p><w:r><w:t>a</w:t></w:r></w:p>",
            "<w:p><w:r><w:t>b</w:t></w:r></w:p>",
            "</w:body></w:document>"
        );
        let spliced = splice_around(
            source,
            &[(
                0,
                "<w:p w14:paraId=\"1234ABCD\"><w:r><w:t>a</w:t></w:r></w:p>",
            )],
            &[],
            &[(1, "0BCD1234")],
            ("", "<w:p><w:r><w:t>b</w:t></w:r></w:p>"),
        )
        .expect("spliced");
        let occurrences = paragraph_occurrences(&spliced).expect("occurrences");
        assert_eq!(occurrences[0].para_id.as_deref(), Some("1234ABCD"));
        assert_eq!(occurrences[1].para_id.as_deref(), Some("0BCD1234"));
        assert!(spliced.contains("<w:r><w:t>a</w:t></w:r></w:p><w:p "));
        assert!(
            spliced.contains("xmlns:w14=\"http://schemas.microsoft.com/office/word/2010/wordml\"")
        );
    }

    #[test]
    fn declares_the_prefixes_a_rewritten_paragraph_uses() {
        let source = concat!(
            "<w:document xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\"><w:body>",
            "<w:p><w:r><w:t>a</w:t></w:r></w:p>",
            "</w:body></w:document>"
        );
        let rewritten = "<w:p w14:paraId=\"1234ABCD\"><w:r><w:t>new</w:t></w:r></w:p>";
        assert_eq!(
            splice(source, &[(0, rewritten)], &[0]),
            Some(format!(
                "<w:document xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\" xmlns:w14=\"http://schemas.microsoft.com/office/word/2010/wordml\" xmlns:mc=\"{MC_NAMESPACE}\" mc:Ignorable=\"w14\"><w:body>{rewritten}</w:body></w:document>"
            ))
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
    fn a_comment_range_moved_across_an_equation_rewrites_the_paragraph() {
        let math = "<m:oMath xmlns:m=\"http://schemas.openxmlformats.org/officeDocument/2006/math\"><m:r><m:t>x</m:t></m:r></m:oMath>";
        let source = format!(
            "{ROOT}<w:body><w:p><w:commentRangeStart w:id=\"1\"/>{math}{math}<w:commentRangeEnd w:id=\"1\"/><w:r><w:commentReference w:id=\"1\"/></w:r></w:p></w:body></w:document>"
        );
        let narrowed = format!(
            "<w:p>{math}<w:commentRangeStart w:id=\"1\"/>{math}<w:commentRangeEnd w:id=\"1\"/><w:r><w:commentReference w:id=\"1\"/></w:r></w:p>"
        );
        assert!(
            splice(&source, &[(0, narrowed.as_str())], &[])
                .is_some_and(|xml| xml.contains(&narrowed))
        );
    }

    #[test]
    fn a_field_end_written_with_character_references_still_closes_its_field() {
        let opening = "<w:p><w:r><w:fldChar w:fldCharType=\"begin\"/></w:r><w:r><w:instrText> TOC </w:instrText></w:r><w:r><w:fldChar w:fldCharType=\"separate\"/></w:r><w:r><w:t>First</w:t></w:r></w:p>";
        let closing = "<w:p><w:r><w:t>Second</w:t></w:r><w:r><w:fldChar w:fldCharType=\"e&#110;d\"/></w:r></w:p>";
        let source = format!("{ROOT}<w:body>{opening}{closing}</w:body></w:document>");
        let written_opening = "<w:p><w:r><w:fldChar w:fldCharType=\"begin\"/></w:r><w:r><w:instrText> TOC </w:instrText></w:r><w:r><w:fldChar w:fldCharType=\"separate\"/></w:r><w:r><w:t>First</w:t></w:r><w:r><w:fldChar w:fldCharType=\"end\"/></w:r></w:p>";
        let written_closing = "<w:p><w:r><w:t>Second, edited</w:t></w:r></w:p>";
        assert_eq!(
            splice(&source, &[(0, written_opening), (1, written_closing)], &[1]),
            Some(format!(
                "{ROOT}<w:body>{written_opening}{written_closing}</w:body></w:document>"
            ))
        );
        let unknown = "<w:p><w:r><w:fldChar w:fldCharType=\"other\"/></w:r></w:p>";
        let source = format!("{ROOT}<w:body>{unknown}</w:body></w:document>");
        assert_eq!(splice(&source, &[(0, unknown)], &[]), None);
    }

    #[test]
    fn overlapping_ranges_are_rewritten_together_and_separate_ones_are_not() {
        let paragraph =
            |index: usize, marks: &str| format!("<w:p>{marks}<w:r><w:t>{index}</w:t></w:r></w:p>");
        let marks = [
            "<w:bookmarkStart w:id=\"1\" w:name=\"a\"/>",
            "",
            "<w:bookmarkEnd w:id=\"1\"/><w:bookmarkStart w:id=\"2\" w:name=\"b\"/>",
            "<w:bookmarkEnd w:id=\"2\"/>",
            "<w:bookmarkStart w:id=\"3\" w:name=\"c\"/>",
            "<w:bookmarkEnd w:id=\"3\"/>",
        ];
        let source_paragraphs: Vec<String> = marks
            .iter()
            .enumerate()
            .map(|(index, marks)| paragraph(index, marks))
            .collect();
        let written_paragraphs: Vec<String> = source_paragraphs
            .iter()
            .map(|xml| xml.replace("</w:p>", "<w:r/></w:p>"))
            .collect();
        let source = format!(
            "{ROOT}<w:body>{}</w:body></w:document>",
            source_paragraphs.concat()
        );
        let written: Vec<(u32, &str)> = written_paragraphs
            .iter()
            .enumerate()
            .map(|(index, xml)| (index as u32, xml.as_str()))
            .collect();
        let expected = format!(
            "{ROOT}<w:body>{}{}</w:body></w:document>",
            written_paragraphs[..4].concat(),
            source_paragraphs[4..].concat()
        );
        assert_eq!(splice(&source, &written, &[3]), Some(expected));
    }

    #[test]
    fn a_field_spanning_paragraphs_is_kept_or_rewritten_whole() {
        let opening = "<w:p><w:r><w:fldChar w:fldCharType=\"begin\"/></w:r><w:r><w:instrText> TOC </w:instrText></w:r><w:r><w:fldChar w:fldCharType=\"separate\"/></w:r><w:r><w:t>First entry</w:t></w:r></w:p>";
        let closing = "<w:p><w:r><w:t>Second entry</w:t></w:r><w:r><w:fldChar w:fldCharType=\"end\"/></w:r></w:p>";
        let outside = "<w:p><w:r><w:t>Outside the field</w:t></w:r></w:p>";
        let source = format!("{ROOT}<w:body>{opening}{closing}{outside}</w:body></w:document>");
        let written_opening = "<w:p><w:r><w:fldChar w:fldCharType=\"begin\"/></w:r><w:r><w:instrText> TOC </w:instrText></w:r><w:r><w:fldChar w:fldCharType=\"separate\"/></w:r><w:r><w:t>First entry</w:t></w:r><w:r><w:fldChar w:fldCharType=\"end\"/></w:r></w:p>";
        let written_closing = "<w:p><w:r><w:t>Second entry, edited</w:t></w:r></w:p>";
        let spliced = splice(
            &source,
            &[(0, written_opening), (1, written_closing), (2, outside)],
            &[1],
        );
        assert_eq!(
            spliced,
            Some(format!(
                "{ROOT}<w:body>{written_opening}{written_closing}{outside}</w:body></w:document>"
            ))
        );
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
                &[(0, written_opening), (1, closing), (2, edited_outside)],
                &[2]
            ),
            Some(format!(
                "{ROOT}<w:body>{opening}{closing}{edited_outside}</w:body></w:document>"
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
        assert!(splice(&source, &[(0, edited)], &[0]).is_none());
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
}
