//! Story parts written as their source XML with only the paragraphs a save changed re-serialized.

use std::borrow::Cow;
use std::collections::{BTreeMap, BTreeSet, HashMap};
use std::ops::Range;

use crate::paragraph_identity::{
    ParagraphOccurrence, Tag, attribute, paragraph_occurrences, patch_paragraph_ids, tags,
    unescaped,
};

use super::paragraph_ids::S13SplicedPart;
use super::s13::element_span;

const W_NAMESPACE: &str = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

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

/// Whether `source`'s root binds every prefix in `prefixes` to the namespace `serialized`'s
/// root binds it to, so that paragraphs written with them need no new declaration.
fn root_binds(source: &str, serialized: &str, prefixes: &BTreeSet<&str>) -> Option<bool> {
    let root = root_tag(source)?;
    let bound = bindings(source, &root);
    let target = root_tag(serialized)?;
    let wanted = bindings(serialized, &target);
    Some(prefixes.iter().all(|prefix| {
        matches!(
            (bound.get(prefix), wanted.get(prefix)),
            (Some(existing), Some(uri)) if existing.trim() == uri.trim()
        )
    }))
}

/// Each addressed paragraph's parent element, by ordinal, or `w:fldChar` for one that starts
/// inside a complex field, whose code or result it belongs to. `None` unless every namespace
/// prefix in `source` keeps one binding and every namespace one prefix, `w` is bound to
/// WordprocessingML, no default namespace is declared, so that element names identify elements,
/// and every field character is a begin, separate or end that closes an open field.
fn parents<'s>(
    source: &'s str,
    spans: &BTreeMap<u32, Range<usize>>,
) -> Option<HashMap<u32, &'s str>> {
    let starts: HashMap<usize, u32> = spans
        .iter()
        .map(|(ordinal, span)| (span.start, *ordinal))
        .collect();
    let mut prefixes: HashMap<&str, &str> = HashMap::new();
    let mut namespaces: HashMap<&str, &str> = HashMap::new();
    let mut stack: Vec<&str> = Vec::new();
    let mut parents = HashMap::new();
    let mut fields = 0usize;
    for tag in tags(source)? {
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
            parents.insert(
                *ordinal,
                if fields > 0 {
                    "w:fldChar"
                } else {
                    *stack.last()?
                },
            );
        }
        if !tag.empty {
            stack.push(tag.name);
        }
    }
    (prefixes.get("w") == Some(&W_NAMESPACE)).then_some(parents)
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

/// Whether `xml`, one paragraph, holds no equations or content controls and pairs only within
/// itself, so that rewriting it leaves every other paragraph's markup whole.
fn simple(xml: &str) -> Option<bool> {
    for tag in tags(xml)? {
        if !tag.end
            && matches!(
                tag.name
                    .rsplit_once(':')
                    .map_or(tag.name, |(_, local)| local),
                "oMath" | "oMathPara" | "sdt"
            )
        {
            return Some(false);
        }
    }
    pairs_within(xml)
}

/// Whether `xml`, one paragraph, holds both ends of everything in it that can pair across
/// paragraphs: bookmarks, permissions, comment and custom XML ranges and proofing marks, each
/// opened before it is closed; and no field character or move, whose partners lie elsewhere.
/// `None` on a range marker without an ID or a proofing mark without a type.
fn pairs_within(xml: &str) -> Option<bool> {
    let mut open: HashMap<(&str, String), usize> = HashMap::new();
    for tag in tags(xml)? {
        if tag.end {
            continue;
        }
        let local = tag
            .name
            .rsplit_once(':')
            .map_or(tag.name, |(_, local)| local);
        if local == "fldChar" || local.starts_with("move") || local.starts_with("customXmlMove") {
            return Some(false);
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
        let count = open.entry((kind, id)).or_insert(0);
        if opens {
            *count += 1;
        } else if *count == 0 {
            return Some(false);
        } else {
            *count -= 1;
        }
    }
    Some(open.values().all(|count| *count == 0))
}

fn paragraph_span(xml: &str, occurrence: &ParagraphOccurrence) -> Option<Range<usize>> {
    let start = occurrence.tag.start;
    let relative = element_span(&xml[start..], &[])?;
    Some(start + relative.start..start + relative.end)
}

/// `source` with the paragraphs `part` addresses rewritten from `written` where they changed
/// or their written XML no longer agrees with their source on comments, revisions, notes or
/// relationships, every other byte kept except paragraph IDs: the written ones, and
/// `assignments` for source paragraphs no model paragraph is written from. `None` when the
/// part cannot be spliced, so the caller writes `serialized` whole: among others when a
/// paragraph to rewrite is not [`simple`] in its source or written XML, or sits in anything
/// but a story root or table cell, and when the result would need a namespace declaration the
/// source does not have.
pub(crate) fn splice_story_part(
    source: &str,
    serialized: &str,
    written: &HashMap<u32, String>,
    part: &S13SplicedPart,
    assignments: &BTreeMap<u32, String>,
) -> Option<String> {
    if written.len() != part.paragraphs.len()
        || part
            .paragraphs
            .iter()
            .any(|ordinal| !written.contains_key(ordinal))
    {
        return None;
    }
    let changed: BTreeSet<u32> = part.changed.iter().copied().collect();
    let occurrences = paragraph_occurrences(source)?;
    let mut spans = BTreeMap::new();
    for &ordinal in written.keys() {
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
        subtract(&mut outside_written, comment_counts(&written[ordinal])?);
    }
    if outside_source != outside_written {
        return None;
    }
    let mut replaced = BTreeSet::new();
    for (&ordinal, xml) in written {
        if changed.contains(&ordinal)
            || !marks(&source[spans[&ordinal].clone()])?.admit(&marks(xml)?)
        {
            replaced.insert(ordinal);
        }
    }
    let parents = parents(source, &spans)?;
    for ordinal in &replaced {
        if !parents
            .get(ordinal)
            .is_some_and(|parent| PARENTS.contains(parent))
            || !simple(&source[spans[ordinal].clone()])?
            || !simple(&written[ordinal])?
        {
            return None;
        }
    }
    let mut prefixes = BTreeSet::new();
    for ordinal in &replaced {
        prefixes.extend(required_prefixes(&written[ordinal])?);
    }
    if !root_binds(source, serialized, &prefixes)? {
        return None;
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
    let mut edits: Vec<(Range<usize>, &str)> = Vec::new();
    for ordinal in replaced {
        let range = paragraph_span(&source, occurrences.get(ordinal as usize)?)?;
        edits.push((range, written[&ordinal].as_str()));
    }
    let mut output = String::with_capacity(source.len());
    let mut cursor = 0;
    for (range, text) in edits {
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
    const W_ONLY_ROOT: &str =
        "<w:document xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\">";
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
        let written: HashMap<u32, String> = paragraphs
            .iter()
            .map(|(ordinal, xml)| (*ordinal, (*xml).to_owned()))
            .collect();
        let body: String = paragraphs.iter().map(|(_, xml)| *xml).collect();
        let serialized =
            format!("{SERIALIZED_ROOT}<w:body>{before}{body}{after}</w:body></w:document>");
        let part = S13SplicedPart {
            part: "word/document.xml".to_owned(),
            sha256: String::new(),
            paragraphs: paragraphs.iter().map(|(ordinal, _)| *ordinal).collect(),
            changed: changed.to_vec(),
        };
        let assignments = assignments
            .iter()
            .map(|(ordinal, id)| (*ordinal, (*id).to_owned()))
            .collect();
        splice_story_part(source, &serialized, &written, &part, &assignments)
    }

    fn splice(source: &str, paragraphs: &[(u32, &str)], changed: &[u32]) -> Option<String> {
        splice_around(source, paragraphs, changed, &[], ("", ""))
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
    fn writes_the_whole_part_when_a_rewritten_paragraph_needs_a_root_declaration() {
        let source = concat!(
            "<w:document xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\"><w:body>",
            "<w:p><w:r><w:t>a</w:t></w:r></w:p>",
            "</w:body></w:document>"
        );
        let rewritten = "<w:p w14:paraId=\"1234ABCD\"><w:r><w:t>new</w:t></w:r></w:p>";
        assert_eq!(splice(source, &[(0, rewritten)], &[0]), None);
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
        let opening = "<w:p><w:r><w:fldChar w:fldCharType=\"begin\"/></w:r><w:r><w:instrText> TOC </w:instrText></w:r><w:r><w:fldChar w:fldCharType=\"separate\"/></w:r><w:r><w:t>First</w:t></w:r></w:p>";
        let closing = "<w:p><w:r><w:t>Second</w:t></w:r><w:r><w:fldChar w:fldCharType=\"e&#110;d\"/></w:r></w:p>";
        let after = "<w:p><w:r><w:t>After</w:t></w:r></w:p>";
        let source = format!("{ROOT}<w:body>{opening}{closing}{after}</w:body></w:document>");
        let written_opening = "<w:p><w:r><w:fldChar w:fldCharType=\"begin\"/></w:r><w:r><w:instrText> TOC </w:instrText></w:r><w:r><w:fldChar w:fldCharType=\"separate\"/></w:r><w:r><w:t>First</w:t></w:r><w:r><w:fldChar w:fldCharType=\"end\"/></w:r></w:p>";
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
        let field = "<w:p><w:r><w:fldChar w:fldCharType=\"begin\"/></w:r><w:r><w:instrText> PAGE </w:instrText></w:r><w:r><w:fldChar w:fldCharType=\"end\"/></w:r></w:p>";
        let source = format!("{ROOT}<w:body>{field}</w:body></w:document>");
        assert_eq!(splice(&source, &[(0, field)], &[0]), None);
    }

    #[test]
    fn ranges_across_paragraphs_keep_their_ends_and_refuse_rewriting_them() {
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
            None
        );
        let crossed = "<w:p><w:bookmarkEnd w:id=\"3\"/><w:r><w:t>b</w:t></w:r><w:bookmarkStart w:id=\"3\" w:name=\"c\"/></w:p>";
        assert_eq!(
            splice(&source, &[(0, opening), (1, crossed), (2, closing)], &[1]),
            None
        );
    }

    /// Asserts that a paragraph holding `start` or `end` of a pair spanning two paragraphs is not
    /// rewritten alone, and that an edit to a third paragraph keeps both.
    fn refuses_rewriting_one_end(start: &str, end: &str) {
        let opening = format!("<w:p>{start}<w:r><w:t>a</w:t></w:r></w:p>");
        let closing = format!("<w:p><w:r><w:t>b</w:t></w:r>{end}</w:p>");
        let outside = "<w:p><w:r><w:t>c</w:t></w:r></w:p>";
        let source = format!("{ROOT}<w:body>{opening}{closing}{outside}</w:body></w:document>");
        let edited = |xml: &str| xml.replace("</w:t>", ", edited</w:t>");
        let (opening_edited, closing_edited) = (edited(&opening), edited(&closing));
        for (paragraphs, changed) in [
            (
                [
                    (0, opening_edited.as_str()),
                    (1, closing.as_str()),
                    (2, outside),
                ],
                0,
            ),
            (
                [
                    (0, opening.as_str()),
                    (1, closing_edited.as_str()),
                    (2, outside),
                ],
                1,
            ),
        ] {
            assert_eq!(
                splice(&source, &paragraphs, &[changed]),
                None,
                "{start} {end}"
            );
        }
        let outside_edited = edited(outside);
        assert_eq!(
            splice(
                &source,
                &[
                    (0, opening.as_str()),
                    (1, closing.as_str()),
                    (2, outside_edited.as_str())
                ],
                &[2]
            ),
            Some(format!(
                "{ROOT}<w:body>{opening}{closing}{outside_edited}</w:body></w:document>"
            ))
        );
    }

    #[test]
    fn a_bookmark_across_paragraphs_is_not_rewritten_at_one_end() {
        refuses_rewriting_one_end(
            "<w:bookmarkStart w:id=\"1\" w:name=\"b\"/>",
            "<w:bookmarkEnd w:id=\"1\"/>",
        );
    }

    #[test]
    fn a_comment_range_across_paragraphs_is_not_rewritten_at_one_end() {
        refuses_rewriting_one_end(
            "<w:commentRangeStart w:id=\"1\"/>",
            "<w:commentRangeEnd w:id=\"1\"/><w:r><w:commentReference w:id=\"1\"/></w:r>",
        );
    }

    #[test]
    fn a_permission_across_paragraphs_is_not_rewritten_at_one_end() {
        refuses_rewriting_one_end(
            "<w:permStart w:id=\"1\" w:edGrp=\"everyone\"/>",
            "<w:permEnd w:id=\"1\"/>",
        );
    }

    #[test]
    fn a_custom_xml_range_across_paragraphs_is_not_rewritten_at_one_end() {
        refuses_rewriting_one_end(
            "<w:customXmlInsRangeStart w:id=\"1\" w:author=\"A\"/>",
            "<w:customXmlInsRangeEnd w:id=\"1\"/>",
        );
    }

    #[test]
    fn a_proofing_mark_across_paragraphs_is_not_rewritten_at_one_end() {
        refuses_rewriting_one_end(
            "<w:proofErr w:type=\"spellStart\"/>",
            "<w:proofErr w:type=\"spellEnd\"/>",
        );
    }

    #[test]
    fn a_move_is_not_rewritten_at_either_half() {
        refuses_rewriting_one_end(
            "<w:moveFromRangeStart w:id=\"0\" w:name=\"m\" w:author=\"A\"/><w:moveFrom w:id=\"1\" w:author=\"A\"><w:r><w:delText>x</w:delText></w:r></w:moveFrom><w:moveFromRangeEnd w:id=\"0\"/>",
            "<w:moveToRangeStart w:id=\"2\" w:name=\"m\" w:author=\"A\"/><w:moveTo w:id=\"3\" w:author=\"A\"><w:r><w:t>x</w:t></w:r></w:moveTo><w:moveToRangeEnd w:id=\"2\"/>",
        );
        let whole_move = "<w:p><w:moveFrom w:id=\"1\" w:author=\"A\"><w:r><w:delText>x</w:delText></w:r></w:moveFrom><w:moveTo w:id=\"3\" w:author=\"A\"><w:r><w:t>x</w:t></w:r></w:moveTo></w:p>";
        let source = format!("{ROOT}<w:body>{whole_move}</w:body></w:document>");
        let rewritten = "<w:p><w:del w:id=\"1\" w:author=\"A\"><w:r><w:delText>x</w:delText></w:r></w:del><w:ins w:id=\"3\" w:author=\"A\"><w:r><w:t>x, edited</w:t></w:r></w:ins></w:p>";
        assert_eq!(splice(&source, &[(0, rewritten)], &[0]), None);
    }

    #[test]
    fn a_field_spanning_paragraphs_is_kept_around_an_edit_elsewhere() {
        let opening = "<w:p><w:r><w:fldChar w:fldCharType=\"begin\"/></w:r><w:r><w:instrText> TOC </w:instrText></w:r><w:r><w:fldChar w:fldCharType=\"separate\"/></w:r><w:r><w:t>First entry</w:t></w:r></w:p>";
        let closing = "<w:p><w:r><w:t>Second entry</w:t></w:r><w:r><w:fldChar w:fldCharType=\"end\"/></w:r></w:p>";
        let outside = "<w:p><w:r><w:t>Outside the field</w:t></w:r></w:p>";
        let source = format!("{ROOT}<w:body>{opening}{closing}{outside}</w:body></w:document>");
        let written_opening = "<w:p><w:r><w:fldChar w:fldCharType=\"begin\"/></w:r><w:r><w:instrText> TOC </w:instrText></w:r><w:r><w:fldChar w:fldCharType=\"separate\"/></w:r><w:r><w:t>First entry</w:t></w:r><w:r><w:fldChar w:fldCharType=\"end\"/></w:r></w:p>";
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
        let opening = "<w:p><w:r><w:fldChar w:fldCharType=\"begin\"/></w:r><w:r><w:instrText> IF 1 = 1 </w:instrText></w:r></w:p>";
        let code = "<w:p><w:r><w:instrText>\"yes\" </w:instrText></w:r></w:p>";
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
    fn writes_the_whole_part_when_patched_ids_would_declare_a_namespace_a_rewrite_needs() {
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
        assert_eq!(
            splice_around(&source, &paragraphs, &[1], &assigned, ("", "")),
            None
        );
        let declared = source.replace(W_ONLY_ROOT, ROOT);
        let spliced =
            splice_around(&declared, &paragraphs, &[1], &assigned, ("", "")).expect("spliced");
        assert!(spliced.starts_with(ROOT));
        assert_eq!(spliced.matches("xmlns:w14=").count(), 1);
        let occurrences = paragraph_occurrences(&spliced).expect("occurrences");
        assert_eq!(occurrences[0].para_id.as_deref(), Some("1234ABCD"));
        assert_eq!(occurrences[1].para_id.as_deref(), Some("0BCD1234"));
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
