//! Verification of a saved comparison against both inputs.

use std::collections::{BTreeMap, HashMap};
use std::ops::Range;

use yrs::{Any, Transact};

use super::format::{Formats, RunStyles, View};
use super::options::CompareLimits;
use super::text::{self, COMPLEX_SCRIPT_BOLD, COMPLEX_SCRIPT_ITALIC, Hunk, UnitAttrs};
use super::xml::Namespaces;
use super::{
    CompareDiagnostic, CompareDiagnosticCode, CompareInput, RevisionNumbers, location, package,
};
use crate::ops::{Chunk, ChunkKind, capture_pilcrow};
use crate::structured::Severity;
use crate::target::{EditTextView, ParagraphView, StoryView, TextAtom};
use crate::{DEL, EditingDoc, INS, story_ref};

/// A body paragraph a comparison changed, and what reading it back must show.
#[derive(Clone, Debug)]
pub(crate) struct ChangedParagraph {
    pub session: usize,
    pub block: usize,
    pub path: Vec<u32>,
    pub revised_text: String,
    pub revised_atoms: Vec<TextAtom>,
    /// The revised formatting of each UTF-16 unit, in comparable form.
    pub revised_attrs: Vec<String>,
    /// Effective run formatting from source XML: the original and the revised paragraph's.
    pub original_formats: Formats,
    pub revised_formats: Formats,
    /// The paragraph's changes, as positions in the change list.
    pub changes: Range<usize>,
    pub hunks: Vec<Hunk>,
}

/// What a saved comparison must satisfy.
pub(crate) struct ComparePostconditions {
    pub original: Vec<u8>,
    pub document_part: String,
    pub author: String,
    pub date: String,
    pub paragraphs: Vec<ChangedParagraph>,
    /// The serialized revision ids of each change, in change order.
    pub change_ids: Vec<RevisionNumbers>,
}

const BODY: &str = "body";
const VERIFY_CLIENT_ID: u64 = 3;

fn mismatch(
    message: impl Into<String>,
    part: Option<&str>,
    path: Option<Vec<u32>>,
) -> Vec<CompareDiagnostic> {
    vec![CompareDiagnostic {
        code: CompareDiagnosticCode::RoundtripMismatch,
        severity: Severity::Error,
        message: message.into(),
        locations: vec![location(CompareInput::Output, part, path)],
    }]
}

/// Formatting in comparable form, without the complex-script bold and italic a comparison states
/// for its save, which a reopened document leaves implicit.
pub(crate) fn comparable(attrs: &UnitAttrs) -> String {
    let mut normalized = (**attrs).clone();
    normalized.remove(COMPLEX_SCRIPT_BOLD);
    normalized.remove(COMPLEX_SCRIPT_ITALIC);
    text::attrs_key(&normalized)
}

fn seeded(bytes: &[u8]) -> Option<(EditingDoc, StoryView, StoryView)> {
    let doc = EditingDoc::new(VERIFY_CLIENT_ID);
    let (envelope, parts) = crate::seed::parse_docx_with_parts(bytes).ok()?;
    crate::seed::seed_parsed_docx_with(&doc, envelope, Some(&parts)).ok()?;
    let (accepted, original) = {
        let txn = doc.yrs_doc().transact();
        (
            StoryView::build(&doc, &txn, BODY, EditTextView::Accepted)?,
            StoryView::build(&doc, &txn, BODY, EditTextView::Original)?,
        )
    };
    Some((doc, accepted, original))
}

/// The body chunks overlapping `start..=end`.
fn chunks_within(chunks: &[Chunk], start: u32, end: u32) -> &[Chunk] {
    let first = chunks.partition_point(|chunk| chunk.end() <= start);
    let last = chunks.partition_point(|chunk| chunk.start <= end);
    &chunks[first..last.max(first)]
}

/// Embed kinds before the paragraph's first inline unit, and its mark's properties without
/// identity and bookkeeping.
fn frame(doc: &EditingDoc, paragraph: &ParagraphView) -> (Vec<String>, String) {
    let txn = doc.yrs_doc().transact();
    let Ok(story) = story_ref(&txn, BODY) else {
        return (Vec::new(), String::new());
    };
    let chunks = doc.chunk_snapshot(BODY, &story, &txn);
    let mut leading = Vec::new();
    let mut mark = String::new();
    for chunk in chunks_within(&chunks, paragraph.start, paragraph.pilcrow) {
        match &chunk.kind {
            ChunkKind::Embed(map) if chunk.start < paragraph.node_start => leading.push(
                map.as_ref()
                    .and_then(|map| crate::map_string(map, &txn, crate::KIND_KEY))
                    .unwrap_or_default(),
            ),
            ChunkKind::Pilcrow(map) if chunk.start == paragraph.pilcrow => {
                let (_, properties) = capture_pilcrow(map, &txn);
                let kept: BTreeMap<String, Any> = properties
                    .into_iter()
                    .filter(|(key, _)| !key.starts_with('_') && key != "textId")
                    .collect();
                mark = text::attrs_key(&kept);
            }
            _ => {}
        }
    }
    (leading, mark)
}

/// Units the paragraph's stamped chunks cover in `view`'s offsets, with each stamp's revision.
fn stamped(doc: &EditingDoc, paragraph: &ParagraphView, key: &str) -> (Vec<Range<u32>>, Vec<Any>) {
    let txn = doc.yrs_doc().transact();
    let Ok(story) = story_ref(&txn, BODY) else {
        return (Vec::new(), Vec::new());
    };
    let chunks = doc.chunk_snapshot(BODY, &story, &txn);
    let mut ranges: Vec<Range<u32>> = Vec::new();
    let mut revisions = Vec::new();
    for chunk in chunks_within(&chunks, paragraph.node_start, paragraph.pilcrow) {
        let start = chunk.start.max(paragraph.node_start);
        let end = chunk.end().min(paragraph.pilcrow);
        if start >= end || !chunk.attr_active(key) {
            continue;
        }
        revisions.extend(chunk.attrs.get(key).cloned());
        let range = paragraph.offset_of_raw(start)..paragraph.offset_of_raw(end);
        match ranges.last_mut() {
            Some(last) if last.end == range.start => last.end = range.end,
            _ => ranges.push(range),
        }
    }
    (ranges, revisions)
}

fn revision_field<'a>(revision: &'a Any, field: &str) -> Option<&'a str> {
    match revision {
        Any::Map(map) => match map.get(field) {
            Some(Any::String(value)) => Some(value),
            _ => None,
        },
        _ => None,
    }
}

/// The revision ids of every `w:ins` and `w:del` below `element`, and other tracked-change
/// elements found.
fn serialized_revisions(
    element: &docx_parse::XmlElement,
    author: &str,
    date: &str,
    found: &mut Vec<(bool, Option<u32>)>,
) -> Result<(), String> {
    for child in element.child_elements() {
        let local = child.local_name();
        if child.matches_name("w", "ins") || child.matches_name("w", "del") {
            if child.attribute(Some("w"), "author") != Some(author)
                || child.attribute(Some("w"), "date") != Some(date)
            {
                return Err(format!("a w:{local} is not attributed to the comparison"));
            }
            let id = child
                .attribute(Some("w"), "id")
                .and_then(|id| id.parse::<u32>().ok());
            found.push((local == "ins", id));
        } else if matches!(
            local,
            "moveFrom"
                | "moveTo"
                | "rPrChange"
                | "pPrChange"
                | "moveFromRangeStart"
                | "moveToRangeStart"
        ) {
            return Err(format!("the paragraph holds an unexpected w:{local}"));
        }
        serialized_revisions(child, author, date, found)?;
    }
    Ok(())
}

/// The element `path` names below `root`, entering namespaces along the way; `children` holds
/// the body's children so the lookup below the body is constant.
fn element_at<'a>(
    root: &'a docx_parse::XmlElement,
    children: &[&'a docx_parse::XmlElement],
    path: &[u32],
    ns: &mut Namespaces,
) -> Option<&'a docx_parse::XmlElement> {
    ns.enter(root);
    let (body, rest) = path.split_first()?;
    let (first, rest) = rest.split_first()?;
    ns.enter(root.child_elements().nth(*body as usize)?);
    let mut current = *children.get(*first as usize)?;
    for step in rest {
        ns.enter(current);
        current = current.child_elements().nth(*step as usize)?;
    }
    Some(current)
}

/// Checks the effective run formatting of every changed paragraph in both views of the output
/// against the inputs', read from source XML rather than the editing session.
fn check_formats(
    original: &[(String, Vec<u8>)],
    output_xml: &str,
    part: &str,
    expected: &ComparePostconditions,
) -> Result<(), Vec<CompareDiagnostic>> {
    let styles = RunStyles::read(original).map_err(|message| {
        mismatch(
            format!("the original's styles cannot be read: {message}"),
            None,
            None,
        )
    })?;
    let document = package::parse(output_xml.as_bytes(), part).ok_or_else(|| {
        mismatch(
            "the saved main document part cannot be read",
            Some(part),
            None,
        )
    })?;
    let root = document
        .root()
        .ok_or_else(|| mismatch("the saved main document part is empty", Some(part), None))?;
    let body_index = expected
        .paragraphs
        .first()
        .map_or(0, |paragraph| paragraph.path[0] as usize);
    let children: Vec<&docx_parse::XmlElement> = root
        .child_elements()
        .nth(body_index)
        .map(|body| body.child_elements().collect())
        .unwrap_or_default();
    for paragraph in &expected.paragraphs {
        let fail = || {
            mismatch(
                "a changed paragraph's effective formatting differs from the input it restores",
                Some(part),
                Some(paragraph.path.clone()),
            )
        };
        let mut ns = Namespaces::default();
        let element = element_at(root, &children, &paragraph.path, &mut ns).ok_or_else(fail)?;
        let original = styles
            .paragraph(element, &mut ns, View::Original)
            .map(|read| read.0);
        let accepted = styles
            .paragraph(element, &mut ns, View::Accepted)
            .map(|read| read.0);
        if original.as_ref() != Some(&paragraph.original_formats)
            || accepted.as_ref() != Some(&paragraph.revised_formats)
        {
            return Err(fail());
        }
    }
    Ok(())
}

/// Verifies `bytes`, a saved comparison, against `expected`.
pub(crate) fn verify_saved_comparison(
    bytes: &[u8],
    expected: &ComparePostconditions,
    limits: &CompareLimits,
) -> Result<(), Vec<CompareDiagnostic>> {
    let part = expected.document_part.as_str();
    if bytes.len() > limits.max_output_bytes {
        return Err(vec![CompareDiagnostic {
            code: CompareDiagnosticCode::LimitExceeded,
            severity: Severity::Error,
            message: format!(
                "the saved document has {} bytes, more than maxOutputBytes ({})",
                bytes.len(),
                limits.max_output_bytes
            ),
            locations: Vec::new(),
        }]);
    }
    let output = ooxml_opc::unzip_parts_with_limits(bytes, limits.max_expanded_bytes as u64)
        .map_err(|message| {
            mismatch(
                format!("the saved document cannot be read: {message}"),
                None,
                None,
            )
        })?;
    let original = ooxml_opc::unzip_parts(&expected.original).map_err(|message| {
        mismatch(
            format!("the original cannot be read: {message}"),
            None,
            None,
        )
    })?;
    let output: HashMap<&str, &[u8]> = output
        .iter()
        .map(|(path, bytes)| (path.as_str(), bytes.as_slice()))
        .collect();
    if output.len() != original.len() {
        return Err(mismatch(
            "the saved document's parts differ from the original's",
            None,
            None,
        ));
    }
    for (path, bytes) in &original {
        match output.get(path.as_str()) {
            None => return Err(mismatch(format!("{path} is missing"), Some(path), None)),
            Some(saved) if path != part && saved != bytes => {
                return Err(mismatch(format!("{path} changed"), Some(path), None));
            }
            _ => {}
        }
    }
    let original_xml = original
        .iter()
        .find(|(path, _)| path == part)
        .and_then(|(_, bytes)| std::str::from_utf8(bytes).ok());
    let output_xml = output
        .get(part)
        .and_then(|bytes| std::str::from_utf8(bytes).ok());
    let (Some(original_xml), Some(output_xml)) = (original_xml, output_xml) else {
        return Err(mismatch(
            "the main document part cannot be read",
            Some(part),
            None,
        ));
    };
    let (mut cursor_original, mut cursor_output) = (0, 0);
    let mut output_spans = Vec::with_capacity(expected.paragraphs.len());
    for paragraph in &expected.paragraphs {
        let spans = docx_parse::element_span(original_xml, &paragraph.path)
            .zip(docx_parse::element_span(output_xml, &paragraph.path));
        let Some((before, after)) = spans else {
            return Err(mismatch(
                "a changed paragraph is not where the original has it",
                Some(part),
                Some(paragraph.path.clone()),
            ));
        };
        if before.start < cursor_original
            || after.start < cursor_output
            || original_xml[cursor_original..before.start] != output_xml[cursor_output..after.start]
        {
            return Err(mismatch(
                "the main document part changed outside the changed paragraphs",
                Some(part),
                Some(paragraph.path.clone()),
            ));
        }
        (cursor_original, cursor_output) = (before.end, after.end);
        output_spans.push(after);
    }
    if original_xml[cursor_original..] != output_xml[cursor_output..] {
        return Err(mismatch(
            "the main document part changed outside the changed paragraphs",
            Some(part),
            None,
        ));
    }

    check_formats(&original, output_xml, part, expected)?;
    let (Some((saved, accepted, rejected)), Some((source, source_view, _))) =
        (seeded(bytes), seeded(&expected.original))
    else {
        return Err(mismatch("the saved document cannot be opened", None, None));
    };
    let count = source_view.paragraphs.len();
    if accepted.paragraphs.len() != count || rejected.paragraphs.len() != count {
        return Err(mismatch(
            "the saved body has a different number of paragraphs",
            Some(part),
            None,
        ));
    }
    let changed: HashMap<usize, (&ChangedParagraph, &Range<usize>)> = expected
        .paragraphs
        .iter()
        .zip(&output_spans)
        .map(|(paragraph, span)| (paragraph.session, (paragraph, span)))
        .collect();
    for index in 0..count {
        let (before, now, then) = (
            &source_view.paragraphs[index],
            &accepted.paragraphs[index],
            &rejected.paragraphs[index],
        );
        let fail = |message: &str| {
            mismatch(
                message,
                Some(part),
                changed
                    .get(&index)
                    .map(|(paragraph, _)| paragraph.path.clone()),
            )
        };
        if then.text != before.text || then.atoms != before.atoms {
            return Err(fail(
                "rejecting every change does not restore the original text",
            ));
        }
        if frame(&saved, now) != frame(&source, before) {
            return Err(fail(
                "a paragraph's properties or surrounding blocks changed",
            ));
        }
        let Some((paragraph, span)) = changed.get(&index) else {
            if now.text != before.text || now.has_revisions() {
                return Err(fail("an unchanged paragraph changed"));
            }
            continue;
        };
        if now.text != paragraph.revised_text || now.atoms != paragraph.revised_atoms {
            return Err(fail(
                "accepting every change does not give the revised text",
            ));
        }
        let keys = |doc: &EditingDoc, view: &ParagraphView| -> Vec<String> {
            text::unit_attrs(doc, BODY, view)
                .iter()
                .map(comparable)
                .collect()
        };
        if keys(&saved, now) != paragraph.revised_attrs {
            return Err(fail(
                "accepting every change does not give the revised formatting",
            ));
        }
        if keys(&saved, then) != keys(&source, before) {
            return Err(fail(
                "rejecting every change does not restore the original formatting",
            ));
        }
        let (inserted, insertions) = stamped(&saved, now, INS);
        let (deleted, deletions) = stamped(&saved, then, DEL);
        let expected_ranges = |pick: fn(&Hunk) -> Range<u32>| -> Vec<Range<u32>> {
            paragraph
                .hunks
                .iter()
                .map(pick)
                .filter(|range| !range.is_empty())
                .collect()
        };
        if inserted != expected_ranges(|hunk| hunk.revised.clone())
            || deleted != expected_ranges(|hunk| hunk.original.clone())
        {
            return Err(fail("the tracked changes cover other text than planned"));
        }
        if insertions.iter().chain(&deletions).any(|revision| {
            revision_field(revision, "author") != Some(expected.author.as_str())
                || revision_field(revision, "date") != Some(expected.date.as_str())
        }) {
            return Err(fail("a tracked change is not attributed to the comparison"));
        }
        let fragment = package::parse(output_xml[(*span).clone()].as_bytes(), part);
        let Some(root) = fragment.as_ref().and_then(|document| document.root()) else {
            return Err(fail("a changed paragraph cannot be read"));
        };
        let mut found = Vec::new();
        serialized_revisions(root, &expected.author, &expected.date, &mut found)
            .map_err(|message| fail(&message))?;
        let mut serialized: Vec<(bool, Option<u32>)> = found;
        serialized.sort_unstable();
        let mut planned: Vec<(bool, Option<u32>)> = expected.change_ids[paragraph.changes.clone()]
            .iter()
            .flat_map(|ids| {
                ids.deletion
                    .map(|id| (false, Some(id)))
                    .into_iter()
                    .chain(ids.insertion.map(|id| (true, Some(id))))
            })
            .collect();
        planned.sort_unstable();
        if serialized != planned {
            return Err(fail(
                "the serialized revisions do not carry their planned ids",
            ));
        }
    }
    Ok(())
}
