//! Comparison of two DOCX packages into tracked changes. Only text in body paragraphs of
//! unchanged structure is compared; every other difference is a blocking diagnostic.

mod align;
mod format;
mod options;
mod package;
mod text;
mod verify;
mod xml;

use std::collections::{BTreeMap, HashMap, HashSet};

use serde::{Deserialize, Serialize};
use serde_json::json;
use yrs::{ReadTxn, Transact};

use self::align::{Budget, Exhausted, Issue};
use self::format::{ComplexScript, Formats, RunStyles, View};
use self::options::UnsupportedPolicy;
use self::package::InspectionBudget;
use self::xml::{Namespaces, SourceParagraph};
use crate::batch::{
    BatchStep, EditGuard, EditHistory, EditSource, EditSuggestion, RichReplacement,
};
use crate::structured::Severity;
use crate::structured::source::sha256;
use crate::target::{EditTextView, StoryView, TextAtom};
use crate::{EditError, EditResult, EditingDoc, UndoSession};

pub(crate) use options::{CompareLimits, CompareOptions, CompareOptionsWire};
pub(crate) use verify::ComparePostconditions;

const BODY: &str = "body";
/// The client id of the private session the revised package is read into.
const REVISED_CLIENT_ID: u64 = 2;

#[derive(Clone, Copy, Debug, Deserialize, Eq, PartialEq, Serialize)]
#[serde(rename_all = "camelCase")]
pub(crate) enum CompareInput {
    Original,
    Revised,
    Output,
}

/// A place in one of the inputs or the output: a part, an element-child path in it, and a
/// paragraph-local UTF-16 range.
#[derive(Clone, Debug, Deserialize, Eq, PartialEq, Serialize)]
#[serde(rename_all = "camelCase")]
pub(crate) struct CompareLocation {
    pub input: CompareInput,
    pub part: Option<String>,
    pub path: Option<Vec<u32>>,
    pub start: Option<u32>,
    pub end: Option<u32>,
}

pub(crate) fn location(
    input: CompareInput,
    part: Option<&str>,
    path: Option<Vec<u32>>,
) -> CompareLocation {
    CompareLocation {
        input,
        part: part.map(str::to_owned),
        path,
        start: None,
        end: None,
    }
}

#[derive(Clone, Copy, Debug, Deserialize, Eq, Hash, PartialEq, Serialize)]
#[serde(rename_all = "kebab-case")]
pub(crate) enum CompareDiagnosticCode {
    InvalidOptions,
    InvalidDocx,
    ExistingRevisions,
    ParagraphInsertion,
    ParagraphDeletion,
    ParagraphMove,
    AmbiguousAlignment,
    TableChange,
    StructureChange,
    FieldChange,
    ObjectChange,
    ContentControlChange,
    FormattingChange,
    UnsupportedContent,
    UnsupportedFormatting,
    OutOfScopeChange,
    OpaquePartChange,
    ProvenanceUnavailable,
    AmbiguousIdentity,
    MetadataDifference,
    LimitExceeded,
    DiagnosticsTruncated,
    BatchRefused,
    SerializationFailed,
    RoundtripMismatch,
}

#[derive(Clone, Debug, Deserialize, Eq, PartialEq, Serialize)]
#[serde(rename_all = "camelCase")]
pub(crate) struct CompareDiagnostic {
    pub code: CompareDiagnosticCode,
    pub severity: Severity,
    pub message: String,
    pub locations: Vec<CompareLocation>,
}

#[derive(Clone, Copy, Debug, Deserialize, Eq, PartialEq, Serialize)]
#[serde(rename_all = "camelCase")]
pub(crate) enum ChangeKind {
    Insertion,
    Deletion,
    Replacement,
}

/// Paragraph-local UTF-16 range of an input's body paragraph.
#[derive(Clone, Debug, Deserialize, Eq, PartialEq, Serialize)]
#[serde(rename_all = "camelCase")]
pub(crate) struct CompareTextSpan {
    pub part: String,
    pub part_sha256: String,
    pub path: Vec<u32>,
    pub start: u32,
    pub end: u32,
    pub text: String,
}

#[derive(Clone, Debug, Deserialize, Eq, PartialEq, Serialize)]
#[serde(rename_all = "camelCase")]
pub(crate) struct ComparedChange {
    pub id: String,
    pub kind: ChangeKind,
    pub original: CompareTextSpan,
    pub revised: CompareTextSpan,
}

/// The serialized ids of one session revision's deletion and insertion.
#[derive(Clone, Copy, Debug, Default, Deserialize, Eq, PartialEq, Serialize)]
#[serde(rename_all = "camelCase")]
pub(crate) struct RevisionNumbers {
    #[serde(skip_serializing_if = "Option::is_none")]
    pub insertion: Option<u32>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub deletion: Option<u32>,
}

#[derive(Clone, Debug, Deserialize, Eq, PartialEq, Serialize)]
#[serde(rename_all = "camelCase")]
pub(crate) struct SourceParagraphTarget {
    pub path: Vec<u32>,
    pub block: usize,
}

/// The main-document paragraphs a save may replace, each revision's serialized ids, and a fixed
/// seed and clock.
#[derive(Clone, Debug, Deserialize, Eq, PartialEq, Serialize)]
#[serde(rename_all = "camelCase")]
pub(crate) struct SaveManifest {
    pub part_sha256: String,
    pub paragraphs: Vec<SourceParagraphTarget>,
    pub revision_ids: BTreeMap<String, RevisionNumbers>,
    pub seed: String,
    pub now: String,
}

pub(crate) enum CompareOutcome {
    Refused(Diagnostics),
    Unchanged(Diagnostics),
    Applied(Box<CompareApplied>),
}

/// A final result as bridge JSON, `noop` marking a successful one that changed nothing: refused
/// when the complete response would exceed `max_bytes`.
fn result_json(
    ok: bool,
    noop: bool,
    changes: &[ComparedChange],
    diagnostics: &[CompareDiagnostic],
    max_bytes: usize,
) -> Result<String, serde_json::Error> {
    let json = serde_json::to_string(&match (ok, noop) {
        (true, true) => {
            json!({ "ok": true, "noop": true, "changes": changes, "diagnostics": diagnostics })
        }
        (true, false) => json!({ "ok": true, "changes": changes, "diagnostics": diagnostics }),
        (false, _) => json!({ "ok": false, "diagnostics": diagnostics }),
    })?;
    let size = json.len();
    if size > max_bytes {
        return serde_json::to_string(&json!({
            "ok": false,
            "diagnostics": [CompareDiagnostic {
                code: CompareDiagnosticCode::LimitExceeded,
                severity: Severity::Error,
                message: format!(
                    "the comparison result needs {size} bytes, more than maxResultBytes ({max_bytes})"
                ),
                locations: Vec::new(),
            }],
        }));
    }
    Ok(json)
}

impl CompareOutcome {
    /// Bridge JSON: the final result of a refusal or no-op, or the save manifest of applied
    /// changes, whose result [`CompareApplied::finish`] returns.
    pub(crate) fn to_json(&self, limits: &CompareLimits) -> Result<String, serde_json::Error> {
        match self {
            Self::Refused(diagnostics) => result_json(
                false,
                false,
                &[],
                &diagnostics.items,
                limits.max_result_bytes,
            ),
            Self::Unchanged(diagnostics) => {
                result_json(true, true, &[], &diagnostics.items, limits.max_result_bytes)
            }
            Self::Applied(applied) => serde_json::to_string(&json!({
                "ok": true,
                "noop": false,
                "save": applied.manifest,
            })),
        }
    }
}

/// Inspection ends: a blocker under `unsupported: "fail"`, or work or diagnostics ran out.
pub(crate) struct Stop;

pub(crate) struct Diagnostics {
    pub(crate) items: Vec<CompareDiagnostic>,
    max: usize,
    policy: UnsupportedPolicy,
    blocked: bool,
}

impl Diagnostics {
    fn new(max: usize, policy: UnsupportedPolicy) -> Self {
        Self {
            items: Vec::new(),
            max: max.max(1),
            policy,
            blocked: false,
        }
    }

    pub(crate) fn push(
        &mut self,
        code: CompareDiagnosticCode,
        severity: Severity,
        message: String,
        locations: Vec<CompareLocation>,
    ) -> Result<(), Stop> {
        let blocking = severity == Severity::Error;
        self.blocked |= blocking;
        if self.items.len() == self.max {
            self.blocked = true;
            if let Some(last) = self.items.last_mut() {
                *last = CompareDiagnostic {
                    code: CompareDiagnosticCode::DiagnosticsTruncated,
                    severity: Severity::Error,
                    message: format!(
                        "more than {} diagnostics; inspection stopped before it was complete",
                        self.max
                    ),
                    locations: Vec::new(),
                };
            }
            return Err(Stop);
        }
        self.items.push(CompareDiagnostic {
            code,
            severity,
            message,
            locations,
        });
        if blocking && self.policy == UnsupportedPolicy::Fail {
            return Err(Stop);
        }
        Ok(())
    }

    fn block(
        &mut self,
        code: CompareDiagnosticCode,
        message: impl Into<String>,
        locations: Vec<CompareLocation>,
    ) -> Result<(), Stop> {
        self.push(code, Severity::Error, message.into(), locations)
    }

    /// Records that a limit ended inspection.
    fn exhausted(&mut self, message: String) -> Stop {
        let _ = self.push(
            CompareDiagnosticCode::LimitExceeded,
            Severity::Error,
            message,
            Vec::new(),
        );
        self.blocked = true;
        Stop
    }

    fn budget(&mut self, exhausted: Exhausted, limits: &options::CompareLimits) -> Stop {
        self.exhausted(match exhausted {
            Exhausted::Alignment => format!(
                "paragraph alignment needs more than maxAlignmentCells ({})",
                limits.max_alignment_cells
            ),
            Exhausted::Diff => format!(
                "text comparison needs more than maxDiffCells ({})",
                limits.max_diff_cells
            ),
            Exhausted::Gap => format!(
                "more than {} consecutive paragraphs differ without an unchanged paragraph between them",
                options::MAX_GAP_PARAGRAPHS
            ),
        })
    }
}

/// One input package, inflated.
pub(crate) struct Package {
    pub input: CompareInput,
    pub parts: Vec<(String, Vec<u8>)>,
    pub content_types: HashMap<String, String>,
    pub document_part: String,
    pub document_sha256: String,
}

impl Package {
    fn name(&self) -> &'static str {
        match self.input {
            CompareInput::Original => "original",
            CompareInput::Revised => "revised",
            CompareInput::Output => "output",
        }
    }

    pub(crate) fn part(&self, path: &str) -> Option<&[u8]> {
        self.parts
            .iter()
            .find(|(name, _)| name == path)
            .map(|(_, bytes)| bytes.as_slice())
    }

    fn read(
        input: CompareInput,
        bytes: &[u8],
        limits: &options::CompareLimits,
        diagnostics: &mut Diagnostics,
    ) -> Result<Self, Stop> {
        let name = match input {
            CompareInput::Original => "original",
            CompareInput::Revised => "revised",
            CompareInput::Output => "output",
        };
        if bytes.len() > limits.max_input_bytes {
            return Err(diagnostics.exhausted(format!(
                "the {name} document has {} bytes, more than maxInputBytes ({})",
                bytes.len(),
                limits.max_input_bytes
            )));
        }
        let parts =
            match ooxml_opc::unzip_parts_with_limits(bytes, limits.max_expanded_bytes as u64) {
                Ok(parts) => parts,
                Err(message) if message.contains("inflated size exceeds") => {
                    return Err(diagnostics.exhausted(format!(
                        "the {name} document inflates to more than maxExpandedBytes ({})",
                        limits.max_expanded_bytes
                    )));
                }
                Err(message) => {
                    diagnostics.block(
                        CompareDiagnosticCode::InvalidDocx,
                        format!("the {name} document is not a readable DOCX package: {message}"),
                        vec![location(input, None, None)],
                    )?;
                    return Err(Stop);
                }
            };
        let Some(content_types) = package::content_types(&parts) else {
            diagnostics.block(
                CompareDiagnosticCode::InvalidDocx,
                format!(
                    "the {name} document's content types are unreadable or leave a part untyped"
                ),
                vec![location(input, Some("[Content_Types].xml"), None)],
            )?;
            return Err(Stop);
        };
        let limits = docx_parse::ParseLimits::default();
        let document_part = docx_parse::relationships::office_document_path(
            &parts,
            &mut docx_parse::ParseBudget::new(&limits),
        )
        .ok();
        let Some((document_part, sha)) = document_part.and_then(|path| {
            let sha = sha256(
                parts
                    .iter()
                    .find(|(name, _)| *name == path)
                    .map(|(_, bytes)| bytes)?,
            );
            Some((path, sha))
        }) else {
            diagnostics.block(
                CompareDiagnosticCode::InvalidDocx,
                format!("the {name} document has no main document part"),
                vec![location(input, None, None)],
            )?;
            return Err(Stop);
        };
        Ok(Self {
            input,
            parts,
            content_types,
            document_part,
            document_sha256: sha,
        })
    }

    fn at(&self, path: &[u32]) -> CompareLocation {
        location(self.input, Some(&self.document_part), Some(path.to_vec()))
    }

    fn span(&self, path: &[u32], range: std::ops::Range<u32>) -> CompareLocation {
        CompareLocation {
            start: Some(range.start),
            end: Some(range.end),
            ..self.at(path)
        }
    }
}

/// A top-level body block other than a paragraph.
struct OtherBlock {
    local: String,
    canonical: String,
    path: Vec<u32>,
}

impl OtherBlock {
    fn code(&self) -> CompareDiagnosticCode {
        match self.local.as_str() {
            "tbl" => CompareDiagnosticCode::TableChange,
            "sdt" => CompareDiagnosticCode::ContentControlChange,
            _ => CompareDiagnosticCode::StructureChange,
        }
    }

    fn label(&self) -> &str {
        match self.local.as_str() {
            "tbl" => "table",
            "sdt" => "content control",
            _ => "block",
        }
    }
}

struct BodyParagraph {
    /// Index among the body story's paragraphs in the session.
    session: usize,
    /// Index among the body's story blocks.
    block: usize,
    path: Vec<u32>,
    source: SourceParagraph,
    /// Effective run formatting per text unit, when the paragraph holds only text runs.
    formats: Option<Formats>,
    /// Each text unit's own complex-script bold and italic.
    complex_script: ComplexScript,
    text: String,
    atoms: Vec<TextAtom>,
}

/// An input's body: its non-paragraph blocks and the paragraphs between them.
struct Body {
    view: StoryView,
    others: Vec<OtherBlock>,
    segments: Vec<Vec<usize>>,
    paragraphs: Vec<BodyParagraph>,
    blocks: usize,
    /// The main document part with every body block as a placeholder, in canonical form.
    rest: String,
}

fn read_body(
    package: &Package,
    doc: &EditingDoc,
    styles: &RunStyles,
    diagnostics: &mut Diagnostics,
) -> Result<Body, Stop> {
    let unavailable = |diagnostics: &mut Diagnostics, message: &str| {
        let _ = diagnostics.block(
            CompareDiagnosticCode::ProvenanceUnavailable,
            message.to_owned(),
            vec![location(package.input, Some(&package.document_part), None)],
        );
        Stop
    };
    let parsed = package
        .part(&package.document_part)
        .and_then(|bytes| package::parse(bytes, &package.document_part));
    let Some(root) = parsed.as_ref().and_then(|document| document.root()) else {
        return Err(unavailable(
            diagnostics,
            "the main document part could not be read",
        ));
    };
    let mut ns = Namespaces::default();
    let root_mark = ns.enter(root);
    let Some((body_index, body)) = root
        .child_elements()
        .enumerate()
        .find(|(_, element)| ns.w_local(element) == Some("body"))
    else {
        return Err(unavailable(
            diagnostics,
            "the main document part has no body",
        ));
    };
    let body_mark = ns.enter(body);
    let view = StoryView::build(doc, &doc.yrs_doc().transact(), BODY, EditTextView::Accepted);
    let sources = doc.source_metadata().and_then(|source| {
        source
            .read()
            .provenance
            .paragraph_sources
            .get(BODY)
            .cloned()
    });
    let (Some(view), Some(sources)) = (view, sources) else {
        return Err(unavailable(
            diagnostics,
            "the body's source provenance is unavailable",
        ));
    };
    if sources.len() != view.paragraphs.len() {
        return Err(unavailable(
            diagnostics,
            "the body's paragraphs do not match their source provenance",
        ));
    }
    let mut sessions: HashMap<usize, usize> = HashMap::new();
    for (session, block) in sources.iter().enumerate() {
        if let Some(block) = block
            && sessions.insert(*block, session).is_some()
        {
            return Err(unavailable(
                diagnostics,
                "two body paragraphs claim one source paragraph",
            ));
        }
    }
    let blocks = docx_parse::story_block_elements(body);
    let mut holes = HashSet::with_capacity(blocks.len());
    let mut result = Body {
        view,
        others: Vec::new(),
        segments: vec![Vec::new()],
        paragraphs: Vec::new(),
        blocks: blocks.len(),
        rest: String::new(),
    };
    for (block, (relative, element)) in blocks.iter().enumerate() {
        let path: Vec<u32> = std::iter::once(body_index as u32)
            .chain(relative.iter().copied())
            .collect();
        holes.insert(path.clone());
        if ns.w_local(element) == Some("p") {
            let Some(&session) = sessions.get(&block) else {
                return Err(unavailable(
                    diagnostics,
                    "a source paragraph has no session paragraph",
                ));
            };
            let paragraph = &result.view.paragraphs[session];
            let (formats, complex_script) = styles
                .paragraph(element, &mut ns, View::Plain)
                .map_or((None, Vec::new()), |(formats, complex)| {
                    (Some(formats), complex)
                });
            result.paragraphs.push(BodyParagraph {
                session,
                block,
                path,
                source: xml::read_paragraph(element, &mut ns),
                formats,
                complex_script,
                text: paragraph.text.clone(),
                atoms: paragraph.atoms.clone(),
            });
            let index = result.paragraphs.len() - 1;
            if let Some(segment) = result.segments.last_mut() {
                segment.push(index);
            }
        } else {
            result.others.push(OtherBlock {
                local: element.local_name().to_owned(),
                canonical: xml::canonical_string(element, &mut ns),
                path,
            });
            result.segments.push(Vec::new());
        }
    }
    if sessions.len() != result.paragraphs.len() {
        return Err(unavailable(
            diagnostics,
            "a session paragraph has no source paragraph",
        ));
    }
    ns.leave(body_mark);
    ns.leave(root_mark);
    result.rest = xml::canonical_skeleton(root, &mut Namespaces::default(), &holes);
    Ok(result)
}

/// Both inputs, read and seeded.
struct Inputs<'a> {
    original: &'a Package,
    revised: &'a Package,
    original_body: &'a Body,
    revised_body: &'a Body,
    options: &'a CompareOptions,
}

/// An aligned paragraph pair whose text differs, with its differing stretches.
struct PairDiff {
    original: usize,
    revised: usize,
    hunks: Vec<text::Hunk>,
}

/// Everything inspection established: the inputs and the supported text differences.
pub(crate) struct ComparisonInspection {
    original: Package,
    revised: Package,
    original_body: Body,
    revised_body: Body,
    pairs: Vec<PairDiff>,
}

/// Checks an aligned paragraph pair, returning its text differences when they are supported.
fn inspect_pair(
    inputs: &Inputs<'_>,
    (o, r): (&BodyParagraph, &BodyParagraph),
    budget: &mut Budget,
    diagnostics: &mut Diagnostics,
) -> Result<Option<Vec<text::Hunk>>, Stop> {
    let (original, revised) = (inputs.original, inputs.revised);
    let both = || vec![original.at(&o.path), revised.at(&r.path)];
    if o.text == r.text {
        if !o.source.same_content(&r.source) {
            let (code, message) = o.source.difference(&r.source);
            diagnostics.block(code, message, both())?;
        }
        return Ok(None);
    }
    for (package, paragraph) in [(original, o), (revised, r)] {
        if let Some(blocker) = &paragraph.source.blocker {
            diagnostics.block(
                blocker.code,
                format!("a changed paragraph cannot be edited: {}", blocker.message),
                vec![package.at(&paragraph.path)],
            )?;
            return Ok(None);
        }
    }
    if o.source.text != o.text
        || r.source.text != r.text
        || o.source.units.len() != o.text.encode_utf16().count()
        || r.source.units.len() != r.text.encode_utf16().count()
        || o.formats.is_none()
        || r.formats.is_none()
    {
        diagnostics.block(
            CompareDiagnosticCode::ProvenanceUnavailable,
            "a changed paragraph's source text does not match its session text",
            both(),
        )?;
        return Ok(None);
    }
    if o.source.ppr != r.source.ppr {
        diagnostics.block(
            CompareDiagnosticCode::FormattingChange,
            "paragraph properties changed on a paragraph whose text changed",
            both(),
        )?;
        return Ok(None);
    }
    let limits = &inputs.options.limits;
    let diff = text::diff_paragraphs(
        (&o.text, &o.atoms),
        (&r.text, &r.atoms),
        inputs.options.granularity,
        budget,
    )
    .map_err(|exhausted| diagnostics.budget(exhausted, limits))?;
    let mut retained = diff.equal;
    for hunk in &diff.hunks {
        let original_text = text::slice(&o.text, hunk.original.clone());
        let revised_text = text::slice(&r.text, hunk.revised.clone());
        if original_text.contains('\u{FFFC}') || revised_text.contains('\u{FFFC}') {
            diagnostics.block(
                CompareDiagnosticCode::UnsupportedContent,
                "a line break or other inline object is inserted or deleted",
                vec![
                    original.span(&o.path, hunk.original.clone()),
                    revised.span(&r.path, hunk.revised.clone()),
                ],
            )?;
            return Ok(None);
        }
        if !original_text.is_empty() && !revised_text.is_empty() {
            let inner = text::retained_within(original_text, revised_text, budget)
                .map_err(|exhausted| diagnostics.budget(exhausted, limits))?;
            retained.extend(inner.into_iter().map(|span| text::Hunk {
                original: span.original.start + hunk.original.start
                    ..span.original.end + hunk.original.start,
                revised: span.revised.start + hunk.revised.start
                    ..span.revised.end + hunk.revised.start,
            }));
        }
    }
    for span in &retained {
        let changed = span
            .original
            .clone()
            .zip(span.revised.clone())
            .find(|(i, j)| o.source.units[*i as usize].rpr != r.source.units[*j as usize].rpr);
        if let Some((i, j)) = changed {
            diagnostics.block(
                CompareDiagnosticCode::FormattingChange,
                "formatting changed on text both documents share",
                vec![
                    original.span(&o.path, i..i + 1),
                    revised.span(&r.path, j..j + 1),
                ],
            )?;
            return Ok(None);
        }
    }
    Ok(Some(diff.hunks))
}

fn check_structure(inputs: &Inputs<'_>, diagnostics: &mut Diagnostics) -> Result<bool, Stop> {
    let (a, b) = (inputs.original_body, inputs.revised_body);
    if a.blocks == b.blocks && a.rest != b.rest {
        diagnostics.block(
            CompareDiagnosticCode::StructureChange,
            "the main document part differs around its body blocks, for example in wrappers, body-level markers or final section properties",
            vec![
                location(CompareInput::Original, Some(&inputs.original.document_part), None),
                location(CompareInput::Revised, Some(&inputs.revised.document_part), None),
            ],
        )?;
    }
    let same_shape = a.others.len() == b.others.len()
        && a.others
            .iter()
            .zip(&b.others)
            .all(|(x, y)| x.local == y.local);
    if !same_shape {
        let describe = |others: &[OtherBlock]| {
            others
                .iter()
                .map(OtherBlock::label)
                .collect::<Vec<_>>()
                .join(", ")
        };
        let code = if a
            .others
            .iter()
            .chain(&b.others)
            .any(|other| other.local == "tbl")
        {
            CompareDiagnosticCode::TableChange
        } else {
            CompareDiagnosticCode::StructureChange
        };
        diagnostics.block(
            code,
            format!(
                "the body's tables and other blocks differ: [{}] against [{}]",
                describe(&a.others),
                describe(&b.others)
            ),
            a.others
                .iter()
                .map(|other| inputs.original.at(&other.path))
                .chain(b.others.iter().map(|other| inputs.revised.at(&other.path)))
                .collect(),
        )?;
        return Ok(false);
    }
    for (x, y) in a.others.iter().zip(&b.others) {
        if x.canonical != y.canonical {
            diagnostics.block(
                x.code(),
                format!(
                    "a {} changed; only body paragraph text is compared",
                    x.label()
                ),
                vec![inputs.original.at(&x.path), inputs.revised.at(&y.path)],
            )?;
        }
    }
    Ok(true)
}

fn report_issues(
    inputs: &Inputs<'_>,
    (segment_a, segment_b): (&[usize], &[usize]),
    issues: Vec<Issue>,
    diagnostics: &mut Diagnostics,
) -> Result<(), Stop> {
    let at_a = |index: usize| {
        inputs
            .original
            .at(&inputs.original_body.paragraphs[segment_a[index]].path)
    };
    let at_b = |index: usize| {
        inputs
            .revised
            .at(&inputs.revised_body.paragraphs[segment_b[index]].path)
    };
    for issue in issues {
        match issue {
            Issue::Moved { original, revised } => diagnostics.block(
                CompareDiagnosticCode::ParagraphMove,
                "a paragraph moved; moves are not recorded as tracked changes",
                vec![at_a(original), at_b(revised)],
            )?,
            Issue::Count { original, revised } => {
                let (code, message) = if revised.len() > original.len() {
                    (
                        CompareDiagnosticCode::ParagraphInsertion,
                        "whole paragraphs were added; only text within existing paragraphs is compared",
                    )
                } else {
                    (
                        CompareDiagnosticCode::ParagraphDeletion,
                        "whole paragraphs were removed; only text within existing paragraphs is compared",
                    )
                };
                diagnostics.block(
                    code,
                    message,
                    original.map(at_a).chain(revised.map(at_b)).collect(),
                )?;
            }
            Issue::Ambiguous { original, revised } => diagnostics.block(
                CompareDiagnosticCode::AmbiguousAlignment,
                "several paragraphs changed with no evidence of which replaced which",
                original.map(at_a).chain(revised.map(at_b)).collect(),
            )?,
        }
    }
    Ok(())
}

fn inspect_bodies(
    inputs: &Inputs<'_>,
    diagnostics: &mut Diagnostics,
) -> Result<Vec<PairDiff>, Stop> {
    let limits = &inputs.options.limits;
    let (a, b) = (inputs.original_body, inputs.revised_body);
    let mut seen = HashMap::new();
    let repeated: Vec<CompareLocation> = a
        .paragraphs
        .iter()
        .filter_map(|paragraph| {
            let id = paragraph.source.para_id.as_deref()?;
            (*seen.entry(id).and_modify(|count| *count += 1).or_insert(1) == 2)
                .then(|| inputs.original.at(&paragraph.path))
        })
        .collect();
    if !repeated.is_empty() {
        diagnostics.push(
            CompareDiagnosticCode::AmbiguousIdentity,
            Severity::Warning,
            "paragraph ids repeat in the original; changes are addressed by position instead"
                .to_owned(),
            repeated,
        )?;
    }
    let mut pairs = Vec::new();
    if !check_structure(inputs, diagnostics)? {
        return Ok(pairs);
    }
    let mut budget = Budget {
        alignment_cells: limits.max_alignment_cells,
        diff_cells: limits.max_diff_cells,
    };
    for (segment_a, segment_b) in a.segments.iter().zip(&b.segments) {
        let texts_a: Vec<&str> = segment_a
            .iter()
            .map(|i| a.paragraphs[*i].text.as_str())
            .collect();
        let texts_b: Vec<&str> = segment_b
            .iter()
            .map(|j| b.paragraphs[*j].text.as_str())
            .collect();
        let alignment = align::align(&texts_a, &texts_b, &mut budget)
            .map_err(|exhausted| diagnostics.budget(exhausted, limits))?;
        report_issues(
            inputs,
            (segment_a, segment_b),
            alignment.issues,
            diagnostics,
        )?;
        for (i, j) in alignment.pairs {
            let (original, revised) = (segment_a[i], segment_b[j]);
            if let Some(hunks) = inspect_pair(
                inputs,
                (&a.paragraphs[original], &b.paragraphs[revised]),
                &mut budget,
                diagnostics,
            )? {
                pairs.push(PairDiff {
                    original,
                    revised,
                    hunks,
                });
            }
        }
    }
    Ok(pairs)
}

/// Reads and seeds both inputs and inspects every difference between them. `doc` receives the
/// original and `revised_doc` the revised package.
fn inspect_comparison(
    (doc, original_bytes): (&EditingDoc, &[u8]),
    (revised_doc, revised_bytes): (&EditingDoc, &[u8]),
    options: &CompareOptions,
    diagnostics: &mut Diagnostics,
) -> Result<ComparisonInspection, Stop> {
    let limits = &options.limits;
    let original = Package::read(CompareInput::Original, original_bytes, limits, diagnostics)?;
    let revised = Package::read(CompareInput::Revised, revised_bytes, limits, diagnostics)?;
    let mut budget = InspectionBudget {
        paragraphs: 0,
        max_paragraphs: limits.max_paragraphs,
        text_units: limits.max_text_units,
    };
    package::scan(&original, &mut budget, diagnostics)?;
    package::scan(&revised, &mut budget, diagnostics)?;
    if diagnostics.blocked {
        return Err(Stop);
    }
    package::compare_parts(&original, &revised, diagnostics)?;
    let styles = match RunStyles::read(&original.parts) {
        Ok(styles) => styles,
        Err(message) => {
            diagnostics.block(
                CompareDiagnosticCode::InvalidDocx,
                format!("the original document's styles cannot be read: {message}"),
                vec![location(CompareInput::Original, None, None)],
            )?;
            return Err(Stop);
        }
    };
    seed(doc, &original, original_bytes, diagnostics)?;
    seed(revised_doc, &revised, revised_bytes, diagnostics)?;
    let original_body = read_body(&original, doc, &styles, diagnostics)?;
    let revised_body = read_body(&revised, revised_doc, &styles, diagnostics)?;
    let pairs = inspect_bodies(
        &Inputs {
            original: &original,
            revised: &revised,
            original_body: &original_body,
            revised_body: &revised_body,
            options,
        },
        diagnostics,
    )?;
    Ok(ComparisonInspection {
        original,
        revised,
        original_body,
        revised_body,
        pairs,
    })
}

/// The changes, the batch steps that author them, and what the saved result must show.
pub(crate) struct ComparePlan {
    changes: Vec<ComparedChange>,
    steps: Vec<RichReplacement>,
    paragraphs: Vec<verify::ChangedParagraph>,
}

/// Plans one rich tracked replacement per differing stretch, in document order.
fn plan_comparison(
    inspection: &ComparisonInspection,
    revised_doc: &EditingDoc,
    options: &CompareOptions,
) -> ComparePlan {
    let (original, revised) = (&inspection.original, &inspection.revised);
    let suggest = EditSuggestion {
        author: options.author.clone(),
        date: options.date.clone(),
    };
    let mut plan = ComparePlan {
        changes: Vec::new(),
        steps: Vec::new(),
        paragraphs: Vec::new(),
    };
    for pair in &inspection.pairs {
        let o = &inspection.original_body.paragraphs[pair.original];
        let r = &inspection.revised_body.paragraphs[pair.revised];
        let revised_attrs = text::unit_attrs(
            revised_doc,
            BODY,
            &inspection.revised_body.view.paragraphs[r.session],
        );
        let first = plan.changes.len();
        for hunk in &pair.hunks {
            let span =
                |package: &Package, paragraph: &BodyParagraph, range: std::ops::Range<u32>| {
                    CompareTextSpan {
                        part: package.document_part.clone(),
                        part_sha256: package.document_sha256.clone(),
                        path: paragraph.path.clone(),
                        start: range.start,
                        end: range.end,
                        text: text::slice(&paragraph.text, range).to_owned(),
                    }
                };
            let kind = match (hunk.original.is_empty(), hunk.revised.is_empty()) {
                (true, _) => ChangeKind::Insertion,
                (_, true) => ChangeKind::Deletion,
                _ => ChangeKind::Replacement,
            };
            let original_span = span(original, o, hunk.original.clone());
            plan.steps.push(RichReplacement {
                story: BODY.to_owned(),
                paragraph: o.session,
                start: hunk.original.start,
                end: hunk.original.end,
                expect: EditGuard {
                    text: original_span.text.clone(),
                },
                runs: text::rich_runs(&r.text, hunk.revised.clone(), &revised_attrs),
                suggest: suggest.clone(),
            });
            plan.changes.push(ComparedChange {
                id: format!("change-{}", plan.changes.len()),
                kind,
                original: original_span,
                revised: span(revised, r, hunk.revised.clone()),
            });
        }
        plan.paragraphs.push(verify::ChangedParagraph {
            session: o.session,
            block: o.block,
            path: o.path.clone(),
            revised_text: r.text.clone(),
            revised_atoms: r.atoms.clone(),
            revised_attrs: revised_attrs.iter().map(verify::comparable).collect(),
            original_formats: o.formats.clone().unwrap_or_default(),
            revised_formats: r.formats.clone().unwrap_or_default(),
            changes: first..plan.changes.len(),
            hunks: pair.hunks.clone(),
        });
    }
    plan
}

/// The largest numeric `w:id` the main document part already uses.
fn highest_annotation_id(package: &Package) -> u32 {
    package
        .part(&package.document_part)
        .and_then(|bytes| package::parse(bytes, &package.document_part))
        .and_then(|document| {
            let root = document.root()?;
            Some(xml::highest_id(root, &mut Namespaces::default()))
        })
        .unwrap_or(0)
}

/// The save clock: the revision date at millisecond precision.
fn save_clock(date: &str) -> String {
    let (seconds, fraction) = date
        .trim_end_matches('Z')
        .split_once('.')
        .unwrap_or((date.trim_end_matches('Z'), ""));
    format!("{seconds}.{:0<3.3}Z", fraction)
}

fn seed(
    doc: &EditingDoc,
    package: &Package,
    bytes: &[u8],
    diagnostics: &mut Diagnostics,
) -> Result<(), Stop> {
    let seeded = crate::seed::parse_docx_with_parts(bytes).and_then(|(envelope, parts)| {
        crate::seed::seed_parsed_docx_with(doc, envelope, Some(&parts))
    });
    if let Err(message) = seeded {
        diagnostics.block(
            CompareDiagnosticCode::InvalidDocx,
            format!("the document could not be read: {message}"),
            vec![location(package.input, None, None)],
        )?;
        return Err(Stop);
    }
    Ok(())
}

/// Applied changes awaiting their saved bytes: what the save needs and the saved result must
/// satisfy.
pub(crate) struct CompareApplied {
    changes: Vec<ComparedChange>,
    diagnostics: Diagnostics,
    manifest: SaveManifest,
    postconditions: ComparePostconditions,
}

impl CompareApplied {
    /// The final result for `bytes`, the saved comparison: verified against both inputs.
    pub(crate) fn finish(
        mut self,
        bytes: &[u8],
        limits: &CompareLimits,
    ) -> Result<String, serde_json::Error> {
        if let Err(found) = verify::verify_saved_comparison(bytes, &self.postconditions, limits) {
            for diagnostic in found {
                if self
                    .diagnostics
                    .push(
                        diagnostic.code,
                        diagnostic.severity,
                        diagnostic.message,
                        diagnostic.locations,
                    )
                    .is_err()
                {
                    break;
                }
            }
        }
        result_json(
            !self.diagnostics.blocked,
            false,
            &self.changes,
            &self.diagnostics.items,
            limits.max_result_bytes,
        )
    }

    /// The final result when the save projection failed with `message`.
    pub(crate) fn fail(
        mut self,
        message: &str,
        limits: &CompareLimits,
    ) -> Result<String, serde_json::Error> {
        let _ = self.diagnostics.block(
            CompareDiagnosticCode::SerializationFailed,
            format!("the compared document could not be saved: {message}"),
            Vec::new(),
        );
        self.diagnostics.blocked = true;
        result_json(
            false,
            false,
            &[],
            &self.diagnostics.items,
            limits.max_result_bytes,
        )
    }
}

/// Applies `plan` to `doc` as one batch outside undo history and maps each revision the batch
/// recorded to the serialized ids of its deletion and insertion. `Ok(Err)` is a refusal already
/// recorded in `diagnostics`; `Err` is an internal failure.
fn apply_comparison(
    doc: &EditingDoc,
    undo: &UndoSession,
    plan: ComparePlan,
    original: (&Package, &[u8]),
    options: &CompareOptions,
    diagnostics: &mut Diagnostics,
) -> EditResult<Result<(Vec<ComparedChange>, SaveManifest, ComparePostconditions), Stop>> {
    let (package, original_bytes) = original;
    let needed = plan
        .changes
        .iter()
        .map(|change| {
            u64::from(!change.original.text.is_empty()) + u64::from(!change.revised.text.is_empty())
        })
        .sum::<u64>();
    let first_id = highest_annotation_id(package);
    if u64::from(first_id) + needed > i32::MAX as u64 {
        return Ok(Err(diagnostics.exhausted(
            "the original's annotation ids leave no room for the revisions".to_owned(),
        )));
    }
    let steps: Vec<BatchStep<'_>> = plan.steps.iter().map(BatchStep::Rich).collect();
    let application = match doc.apply_steps(
        &doc.version(),
        EditSource::Host,
        EditHistory::None,
        &steps,
        undo,
        options.limits.max_staged_bytes,
    )? {
        Ok(application) => application,
        Err(refusal) => {
            let code = serde_json::to_value(refusal.failure.code)
                .ok()
                .and_then(|code| code.as_str().map(str::to_owned))
                .unwrap_or_default();
            let locations = refusal
                .failure
                .step_index
                .and_then(|index| plan.changes.get(index as usize))
                .map(|change| {
                    vec![package.span(
                        &change.original.path,
                        change.original.start..change.original.end,
                    )]
                })
                .unwrap_or_default();
            let _ = diagnostics.block(
                CompareDiagnosticCode::BatchRefused,
                format!(
                    "the planned changes could not be applied ({code}): {}",
                    refusal.failure.message
                ),
                locations,
            );
            return Ok(Err(Stop));
        }
    };
    let mut next = first_id;
    let mut allocate = || {
        next += 1;
        next
    };
    let mut revision_ids: BTreeMap<String, RevisionNumbers> = BTreeMap::new();
    let mut change_ids = Vec::with_capacity(plan.changes.len());
    for (change, receipt) in plan.changes.iter().zip(&application.receipts) {
        let [session_id] = receipt.revision_ids.as_slice() else {
            return Err(EditError::InvalidUpdate(format!(
                "{} recorded {} revisions instead of one",
                change.id,
                receipt.revision_ids.len()
            )));
        };
        if revision_ids.contains_key(session_id) {
            return Err(EditError::InvalidUpdate(format!(
                "{} shares a revision with another change",
                change.id
            )));
        }
        let mut numbers = RevisionNumbers::default();
        if !change.original.text.is_empty() {
            numbers.deletion = Some(allocate());
        }
        if !change.revised.text.is_empty() {
            numbers.insertion = Some(allocate());
        }
        revision_ids.insert(session_id.clone(), numbers);
        change_ids.push(numbers);
    }
    let manifest = SaveManifest {
        part_sha256: package.document_sha256.clone(),
        paragraphs: plan
            .paragraphs
            .iter()
            .map(|paragraph| SourceParagraphTarget {
                path: paragraph.path.clone(),
                block: paragraph.block,
            })
            .collect(),
        revision_ids,
        seed: sha256(original_bytes),
        now: save_clock(&options.date),
    };
    Ok(Ok((
        plan.changes,
        manifest,
        ComparePostconditions {
            original: original_bytes.to_vec(),
            document_part: package.document_part.clone(),
            author: options.author.clone(),
            date: options.date.clone(),
            paragraphs: plan.paragraphs,
            change_ids,
        },
    )))
}

/// Compares `original` with `revised` and applies the differences to `doc`, which must not hold
/// a document yet and is seeded from `original`. `undo` is `doc`'s history; the comparison stays
/// out of it. `Err` is an internal failure.
pub(crate) fn compare_into(
    doc: &EditingDoc,
    undo: &UndoSession,
    original_bytes: &[u8],
    revised_bytes: &[u8],
    options: &CompareOptions,
) -> EditResult<CompareOutcome> {
    let empty = {
        let txn = doc.yrs_doc().transact();
        txn.get_map(crate::STORIES)
            .is_none_or(|stories| yrs::Map::len(&stories, &txn) == 0)
    };
    if !empty {
        return Err(EditError::InvalidUpdate(
            "a comparison needs a session that holds no document".to_owned(),
        ));
    }
    let limits = &options.limits;
    let mut diagnostics = Diagnostics::new(limits.max_diagnostics, options.unsupported);
    let revised_doc = EditingDoc::new(REVISED_CLIENT_ID);
    let inspection = inspect_comparison(
        (doc, original_bytes),
        (&revised_doc, revised_bytes),
        options,
        &mut diagnostics,
    );
    let inspection = match inspection {
        Ok(inspection) if !diagnostics.blocked => inspection,
        _ => return Ok(CompareOutcome::Refused(diagnostics)),
    };
    for pair in &inspection.pairs {
        for (doc, body, index) in [
            (doc, &inspection.original_body, pair.original),
            (&revised_doc, &inspection.revised_body, pair.revised),
        ] {
            let paragraph = &body.paragraphs[index];
            let view = &body.view.paragraphs[paragraph.session];
            text::state_complex_script(doc, BODY, view, &paragraph.complex_script);
        }
    }
    let plan = plan_comparison(&inspection, &revised_doc, options);
    if plan.changes.len() > limits.max_changes {
        let _ = diagnostics.exhausted(format!(
            "the comparison has {} changes, more than maxChanges ({})",
            plan.changes.len(),
            limits.max_changes
        ));
        return Ok(CompareOutcome::Refused(diagnostics));
    }
    if plan.steps.is_empty() {
        if original_bytes.len() > limits.max_output_bytes {
            let _ = diagnostics.exhausted(format!(
                "the unchanged document has {} bytes, more than maxOutputBytes ({})",
                original_bytes.len(),
                limits.max_output_bytes
            ));
            return Ok(CompareOutcome::Refused(diagnostics));
        }
        return Ok(CompareOutcome::Unchanged(diagnostics));
    }
    match apply_comparison(
        doc,
        undo,
        plan,
        (&inspection.original, original_bytes),
        options,
        &mut diagnostics,
    )? {
        Ok((changes, manifest, postconditions)) => {
            Ok(CompareOutcome::Applied(Box::new(CompareApplied {
                changes,
                diagnostics,
                manifest,
                postconditions,
            })))
        }
        Err(Stop) => Ok(CompareOutcome::Refused(diagnostics)),
    }
}

/// Parses bridge options: `Ok(Err(diagnostic))` when they are well-formed but unusable.
pub(crate) fn parse_options(
    json: &str,
) -> Result<Result<CompareOptions, CompareDiagnostic>, serde_json::Error> {
    let wire: CompareOptionsWire = serde_json::from_str(json)?;
    Ok(
        CompareOptions::resolve(wire).map_err(|message| CompareDiagnostic {
            code: CompareDiagnosticCode::InvalidOptions,
            severity: Severity::Error,
            message,
            locations: Vec::new(),
        }),
    )
}

/// The refusal JSON for unusable options.
pub(crate) fn refused_json(diagnostic: CompareDiagnostic) -> Result<String, serde_json::Error> {
    result_json(
        false,
        false,
        &[],
        &[diagnostic],
        options::CEILINGS.max_result_bytes,
    )
}

#[cfg(test)]
mod tests;
