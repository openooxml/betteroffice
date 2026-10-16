//! Markdown rendering of structured content: slide sections, title placeholders as headings,
//! paragraphs, literal list markers, tables, placeholders for unrepresented objects, and notes
//! and comments when exported. Tables Markdown cannot express are rendered as HTML, whose text is
//! entity-escaped rather than Markdown-escaped. Nothing is evaluated or fetched.

use std::collections::{BTreeSet, HashSet};
use std::fmt::Write as _;

use super::*;

/// Renders `content` as Markdown within `options.max_bytes`, stopping at a block boundary.
pub fn render_pptx_markdown(
    content: &PptxStructuredContent,
    options: &PptxMarkdownOptions,
) -> Result<PptxMarkdownContent, ExportFailure> {
    render(content, options).map(|(rendered, _)| rendered)
}

/// The rendering and the bytes of block text it built on the way.
fn render(
    content: &PptxStructuredContent,
    options: &PptxMarkdownOptions,
) -> Result<(PptxMarkdownContent, usize), ExportFailure> {
    let max_bytes = byte_limit(options.max_bytes)?;
    validate(content)?;
    let mut renderer = Renderer {
        output: String::new(),
        anchors: Vec::new(),
        lossy: BTreeSet::new(),
        omitted: Vec::new(),
        max_bytes,
        truncated: false,
        in_list: false,
        cut: false,
        built: 0,
    };
    'slides: for slide in &content.slides {
        if !renderer.slide(slide) {
            break;
        }
        for shape in &slide.shapes {
            if !renderer.shape(shape) {
                break 'slides;
            }
        }
        if let Some(notes) = &slide.notes
            && !renderer.notes(notes)
        {
            break;
        }
        for (index, comment) in slide.comments.iter().enumerate() {
            if !renderer.comment(comment, index == 0) {
                break 'slides;
            }
        }
        renderer.end_list();
    }
    let mut diagnostics = content.diagnostics.clone();
    diagnostics.extend(renderer.omitted);
    for message in &renderer.lossy {
        diagnostics.push(ExportDiagnostic {
            code: ExportDiagnosticCode::MarkdownLossy,
            severity: ExportSeverity::Info,
            anchor: None,
            message: (*message).to_owned(),
        });
    }
    if renderer.truncated {
        diagnostics.push(ExportDiagnostic {
            code: ExportDiagnosticCode::Truncated,
            severity: ExportSeverity::Warning,
            anchor: None,
            message:
                "The Markdown stopped at its byte limit; content after this point is not included."
                    .to_owned(),
        });
    }
    let markdown = if renderer.output.is_empty() {
        String::new()
    } else {
        renderer.output.trim_end().to_owned() + "\n"
    };
    Ok((
        PptxMarkdownContent {
            markdown,
            anchors: renderer.anchors,
            diagnostics,
            truncated: content.truncated || renderer.truncated,
        },
        renderer.built,
    ))
}

fn invalid(message: impl Into<String>) -> ExportFailure {
    ExportFailure {
        code: ExportFailureCode::InvalidContent,
        target: None,
        message: message.into(),
    }
}

/// Shape nesting the renderer accepts.
const MAX_RENDER_DEPTH: usize = 128;
/// Records and metadata entries one rendering validates before it refuses the content.
const MAX_RENDER_RECORDS: usize = 16 * MAX_BLOCKS_LIMIT as usize;
/// Checks the content contract rendering relies on: the schema version, anchor kinds and owners,
/// UTF-16 ranges that match the text they carry, paragraphs that tile their story (a truncated
/// export may stop short), at most one formatting mark of each kind per run, unique ids, cells in
/// column order, resolvable merge origins and a bounded count of records and metadata entries.
fn validate(content: &PptxStructuredContent) -> Result<(), ExportFailure> {
    if content.schema_version != SCHEMA_VERSION {
        return Err(invalid(format!(
            "structured content schema version {} is not supported; expected {SCHEMA_VERSION}",
            content.schema_version
        )));
    }
    let mut validator = Validator {
        truncated: content.truncated,
        ..Validator::default()
    };
    for slide in &content.slides {
        validator.slide(slide)?;
    }
    for diagnostic in &content.diagnostics {
        validator.charge(1)?;
        if let Some(anchor) = &diagnostic.anchor {
            if let PptxAnchor::SourcePart { path, .. } = anchor {
                validator.charge(path.len())?;
            }
            ordered(anchor)?;
        }
    }
    Ok(())
}

fn range_of(anchor: &PptxAnchor) -> Option<TextSpan> {
    match anchor {
        PptxAnchor::Range(range) => Some(TextSpan {
            start: range.start,
            end: range.end,
        }),
        PptxAnchor::Notes { range, .. } | PptxAnchor::Comment { range, .. } => Some(*range),
        _ => None,
    }
}

fn ordered(anchor: &PptxAnchor) -> Result<(), ExportFailure> {
    match range_of(anchor) {
        Some(range) if range.end < range.start => Err(invalid(format!(
            "anchor range {}..{} ends before it starts",
            range.start, range.end
        ))),
        _ => Ok(()),
    }
}

fn units(text: &str) -> u32 {
    text.encode_utf16().count() as u32
}

#[derive(Default)]
struct Validator {
    ids: HashSet<String>,
    records: usize,
    /// Whether the content says it stopped at its limits, so a story may end early.
    truncated: bool,
}

impl Validator {
    fn charge(&mut self, entries: usize) -> Result<(), ExportFailure> {
        self.records = self.records.saturating_add(entries);
        if self.records > MAX_RENDER_RECORDS {
            return Err(ExportFailure {
                code: ExportFailureCode::LimitExceeded,
                target: None,
                message: format!(
                    "content may hold at most {MAX_RENDER_RECORDS} records and metadata entries"
                ),
            });
        }
        Ok(())
    }

    fn record(&mut self, id: &str) -> Result<(), ExportFailure> {
        self.charge(1)?;
        if !self.ids.insert(id.to_owned()) {
            return Err(invalid(format!("id {id:?} occurs more than once")));
        }
        Ok(())
    }

    fn slide(&mut self, slide: &ExportSlide) -> Result<(), ExportFailure> {
        self.record(&slide.id)?;
        let PptxAnchor::Slide { slide_id } = &slide.anchor else {
            return Err(invalid(format!(
                "slide {} is not anchored to a slide",
                slide.id
            )));
        };
        for shape in &slide.shapes {
            self.shape(shape, slide_id, 0)?;
        }
        if let Some(notes) = &slide.notes {
            self.record(&notes.id)?;
            let expected = PptxAnchor::Notes {
                slide_id: slide_id.clone(),
                range: TextSpan {
                    start: 0,
                    end: units(&notes.text),
                },
            };
            if notes.anchor != expected {
                return Err(invalid(format!(
                    "notes {} are not anchored to their text",
                    notes.id
                )));
            }
        }
        for comment in &slide.comments {
            self.record(&comment.id)?;
            let expected = PptxAnchor::Comment {
                slide_id: slide_id.clone(),
                comment_id: comment.comment_id.clone(),
                range: TextSpan {
                    start: 0,
                    end: units(&comment.text),
                },
            };
            if comment.anchor != expected {
                return Err(invalid(format!(
                    "comment {} is not anchored to its text",
                    comment.id
                )));
            }
        }
        Ok(())
    }

    fn shape(
        &mut self,
        shape: &ExportShape,
        slide_id: &str,
        depth: usize,
    ) -> Result<(), ExportFailure> {
        self.record(&shape.id)?;
        if depth >= MAX_RENDER_DEPTH {
            return Err(invalid(format!(
                "shapes nest deeper than {MAX_RENDER_DEPTH} levels"
            )));
        }
        let shape_id = match (&shape.anchor, shape.kind) {
            (PptxAnchor::SourcePart { path, .. }, ExportShapeKind::Unknown) => {
                self.charge(path.len())?;
                None
            }
            (
                PptxAnchor::Shape {
                    slide_id: owner,
                    shape_id,
                },
                kind,
            ) if owner == slide_id && kind != ExportShapeKind::Unknown => Some(shape_id.as_str()),
            _ => {
                return Err(invalid(format!(
                    "shape {} is not anchored to itself on its slide",
                    shape.id
                )));
            }
        };
        if let Some(object) = &shape.object {
            self.record(&object.id)?;
            self.charge(
                object.relationship_ids.len() + object.parts.len() + object.external_targets.len(),
            )?;
        }
        for story in &shape.stories {
            self.story(story, slide_id, shape_id)?;
        }
        if let Some(table) = &shape.table {
            self.table(table, slide_id, shape_id, &shape.anchor)?;
        }
        for child in &shape.children {
            self.shape(child, slide_id, depth + 1)?;
        }
        Ok(())
    }

    fn table(
        &mut self,
        table: &ExportTable,
        slide_id: &str,
        shape_id: Option<&str>,
        shape_anchor: &PptxAnchor,
    ) -> Result<(), ExportFailure> {
        let rows = table.rows.len() as u64;
        let cell_at = |row: u32, column: u32| {
            let cells = &table.rows.get(row as usize)?.cells;
            let index = cells
                .binary_search_by_key(&column, |cell| cell.column)
                .ok()?;
            cells.get(index)
        };
        for (row_index, row) in table.rows.iter().enumerate() {
            for (index, cell) in row.cells.iter().enumerate() {
                self.record(&cell.id)?;
                let fits = cell.row as usize == row_index
                    && index
                        .checked_sub(1)
                        .is_none_or(|previous| row.cells[previous].column < cell.column)
                    && cell.column < table.columns
                    && cell.grid_span > 0
                    && cell.row_span > 0
                    && u64::from(cell.column) + u64::from(cell.grid_span)
                        <= u64::from(table.columns)
                    && u64::from(cell.row) + u64::from(cell.row_span) <= rows;
                let origin = match cell.merge_origin {
                    None => true,
                    Some(origin) => {
                        cell.merged
                            && cell_at(origin.row, origin.column).is_some_and(|owner| {
                                !owner.merged
                                    && (owner.row, owner.column) != (cell.row, cell.column)
                                    && owner.row <= cell.row
                                    && owner.column <= cell.column
                                    && owner.row + owner.row_span > cell.row
                                    && owner.column + owner.grid_span > cell.column
                            })
                    }
                };
                if !fits || !origin {
                    return Err(invalid(format!(
                        "table cell {} does not fit its table's grid",
                        cell.id
                    )));
                }
                match &cell.story {
                    Some(story) => {
                        if cell.anchor != story.anchor {
                            return Err(invalid(format!(
                                "table cell {} is not anchored to its story",
                                cell.id
                            )));
                        }
                        self.story(story, slide_id, shape_id)?;
                    }
                    None if cell.anchor != *shape_anchor => {
                        return Err(invalid(format!(
                            "table cell {} is not anchored to its table",
                            cell.id
                        )));
                    }
                    None => {}
                }
            }
        }
        Ok(())
    }

    fn story(
        &mut self,
        story: &ExportStory,
        slide_id: &str,
        shape_id: Option<&str>,
    ) -> Result<(), ExportFailure> {
        self.record(&story.id)?;
        let text_of = |anchor: &PptxAnchor| match anchor {
            PptxAnchor::Range(range)
                if range.slide_id == slide_id && Some(range.shape_id.as_str()) == shape_id =>
            {
                Some((
                    range.story_id.clone(),
                    TextSpan {
                        start: range.start,
                        end: range.end,
                    },
                ))
            }
            _ => None,
        };
        let Some((story_id, span)) = text_of(&story.anchor).filter(|(_, span)| span.start == 0)
        else {
            return Err(invalid(format!(
                "story {} is not anchored to its shape's text",
                story.id
            )));
        };
        let mut previous_end: Option<u32> = None;
        for paragraph in &story.paragraphs {
            self.record(&paragraph.id)?;
            let range = match text_of(&paragraph.anchor) {
                Some((id, range)) if id == story_id => range,
                _ => {
                    return Err(invalid(format!(
                        "paragraph {} is not anchored in its story",
                        paragraph.id
                    )));
                }
            };
            let expected_start = previous_end.map_or(0, |end| u64::from(end) + 1);
            if range.start > range.end
                || range.end > span.end
                || u64::from(range.start) != expected_start
            {
                return Err(invalid(format!(
                    "paragraph {} is out of order or outside its story",
                    paragraph.id
                )));
            }
            previous_end = Some(range.end);
            let mut at = range.start;
            self.charge(paragraph.runs.len())?;
            for run in &paragraph.runs {
                if run.marks.as_deref().is_some_and(repeats_a_mark) {
                    return Err(invalid(format!(
                        "a run of paragraph {} carries a formatting mark more than once",
                        paragraph.id
                    )));
                }
                let run_range = match text_of(&run.anchor) {
                    Some((id, run_range)) if id == story_id => run_range,
                    _ => {
                        return Err(invalid(format!(
                            "a run of paragraph {} is not anchored in its story",
                            paragraph.id
                        )));
                    }
                };
                let length = match &run.content {
                    ExportRunKind::Text { text } | ExportRunKind::Field { text, .. } => units(text),
                    ExportRunKind::LineBreak => 1,
                    ExportRunKind::Unsupported { .. } => 0,
                };
                if run_range.start != at
                    || run_range.end < run_range.start
                    || run_range.end - run_range.start != length
                {
                    return Err(invalid(format!(
                        "a run of paragraph {} does not match its range",
                        paragraph.id
                    )));
                }
                at = run_range.end;
            }
            if at != range.end {
                return Err(invalid(format!(
                    "the runs of paragraph {} do not cover its range",
                    paragraph.id
                )));
            }
        }
        if !self.truncated && previous_end.unwrap_or(0) != span.end {
            return Err(invalid(format!(
                "the paragraphs of story {} do not cover its text",
                story.id
            )));
        }
        Ok(())
    }
}

/// Whether `marks` holds a kind of mark more than once.
fn repeats_a_mark(marks: &[ExportMark]) -> bool {
    let mut seen = 0_u8;
    marks.iter().any(|mark| {
        let bit = 1_u8
            << match mark {
                ExportMark::Bold => 0,
                ExportMark::Italic => 1,
                ExportMark::Underline { .. } => 2,
                ExportMark::Superscript => 3,
                ExportMark::Subscript => 4,
                ExportMark::SmallCaps => 5,
                ExportMark::AllCaps => 6,
            };
        let repeated = seen & bit != 0;
        seen |= bit;
        repeated
    })
}

struct Renderer {
    output: String,
    anchors: Vec<MarkdownAnchor>,
    lossy: BTreeSet<&'static str>,
    omitted: Vec<ExportDiagnostic>,
    max_bytes: usize,
    truncated: bool,
    in_list: bool,
    /// Whether the block being built stopped early because it cannot fit.
    cut: bool,
    /// Bytes of block text built, kept or not.
    built: usize,
}

/// Whether `ch` ends a line in Markdown or in a browser's text layout.
fn is_line_break(ch: char) -> bool {
    matches!(ch, '\r' | '\n' | '\u{2028}' | '\u{2029}')
}

/// Text safe inside a one-line HTML comment.
fn comment_text(text: &str) -> String {
    let mut output = String::with_capacity(text.len());
    for ch in text.replace("--", "- -").chars() {
        match ch {
            '<' => output.push_str("&lt;"),
            '>' => output.push_str("&gt;"),
            ch if is_line_break(ch) => output.push(' '),
            ch => output.push(ch),
        }
    }
    output
}

/// Escapes Markdown and HTML syntax in document text and puts it on one line, including a
/// trailing `!` that a following link would turn into an image.
fn escape(text: &str) -> String {
    let mut output = String::with_capacity(text.len());
    let mut chars = text.chars().peekable();
    while let Some(ch) = chars.next() {
        if is_line_break(ch) {
            output.push(' ');
            continue;
        }
        let entity = ch == '&'
            && chars
                .peek()
                .is_some_and(|next| next.is_ascii_alphanumeric() || *next == '#');
        if matches!(
            ch,
            '\\' | '`' | '*' | '_' | '[' | ']' | '<' | '>' | '|' | '~'
        ) || entity
        {
            output.push('\\');
        }
        output.push(ch);
    }
    if output.ends_with('!') {
        output.insert(output.len() - 1, '\\');
    }
    output
}

/// Escapes document text for HTML element content.
fn html_text(text: &str) -> String {
    let mut output = String::with_capacity(text.len());
    for ch in text.chars() {
        match ch {
            '&' => output.push_str("&amp;"),
            '<' => output.push_str("&lt;"),
            '>' => output.push_str("&gt;"),
            ch if is_line_break(ch) => output.push(' '),
            ch => output.push(ch),
        }
    }
    output
}

/// Escapes text for an HTML attribute value.
fn attribute(text: &str) -> String {
    let mut output = String::with_capacity(text.len());
    for ch in text.chars() {
        match ch {
            '&' => output.push_str("&amp;"),
            '"' => output.push_str("&quot;"),
            '\'' => output.push_str("&#39;"),
            '<' => output.push_str("&lt;"),
            '>' => output.push_str("&gt;"),
            ch => output.push(ch),
        }
    }
    output
}

/// Link schemes rendered as links, the same as the DOCX and XLSX renderers.
const LINK_SCHEMES: [&str; 3] = ["http", "https", "mailto"];

/// `href` when the scheme a browser would read from it is allowed.
fn safe_href(href: &str) -> Option<&str> {
    let effective = effective_href(href)?;
    let (scheme, _) = effective.split_once(':')?;
    LINK_SCHEMES.contains(&scheme).then_some(href)
}

/// Rounds of decoding after which a still-changing link target is refused.
const MAX_DECODE_ROUNDS: usize = 8;

/// All of `href` as a browser reads its scheme: character references and percent escapes
/// decoded until stable, whitespace and controls (which browsers strip or skip) removed,
/// lowercased. `None` when decoding does not settle within [`MAX_DECODE_ROUNDS`].
fn effective_href(href: &str) -> Option<String> {
    let mut current = href.to_owned();
    for _ in 0..MAX_DECODE_ROUNDS {
        let decoded = percent_decoded(&references_decoded(&current));
        if decoded == current {
            return Some(
                current
                    .chars()
                    .filter(|ch| !ch.is_whitespace() && !ch.is_control())
                    .collect::<String>()
                    .to_ascii_lowercase(),
            );
        }
        current = decoded;
    }
    None
}

/// Named character references that spell URL syntax, whitespace or markup.
const NAMED_REFERENCES: [(&str, char); 28] = [
    ("AMP", '&'),
    ("GT", '>'),
    ("LT", '<'),
    ("NewLine", '\n'),
    ("QUOT", '"'),
    ("Tab", '\t'),
    ("amp", '&'),
    ("apos", '\''),
    ("bsol", '\\'),
    ("colon", ':'),
    ("comma", ','),
    ("commat", '@'),
    ("equals", '='),
    ("excl", '!'),
    ("gt", '>'),
    ("lowbar", '_'),
    ("lpar", '('),
    ("lt", '<'),
    ("nbsp", '\u{a0}'),
    ("num", '#'),
    ("percnt", '%'),
    ("period", '.'),
    ("plus", '+'),
    ("quest", '?'),
    ("quot", '"'),
    ("rpar", ')'),
    ("semi", ';'),
    ("sol", '/'),
];

/// `text` with character references decoded as a browser would: numeric and named, with or
/// without the closing `;`.
fn references_decoded(text: &str) -> String {
    let mut output = String::with_capacity(text.len());
    let mut rest = text;
    while let Some(start) = rest.find('&') {
        output.push_str(&rest[..start]);
        let tail = &rest[start + 1..];
        match reference(tail) {
            Some((ch, length)) => {
                output.push(ch);
                rest = &tail[length..];
                rest = rest.strip_prefix(';').unwrap_or(rest);
            }
            None => {
                output.push('&');
                rest = tail;
            }
        }
    }
    output.push_str(rest);
    output
}

/// The character the reference after an `&` spells, with the bytes its number or name takes.
fn reference(text: &str) -> Option<(char, usize)> {
    if let Some(number) = text.strip_prefix('#') {
        let (digits, radix, prefix) = match number.strip_prefix(['x', 'X']) {
            Some(hex) => (hex, 16, 2),
            None => (number, 10, 1),
        };
        let length = digits
            .find(|ch: char| !ch.is_digit(radix))
            .unwrap_or(digits.len());
        if length == 0 {
            return None;
        }
        let value = u32::from_str_radix(&digits[..length], radix)
            .ok()
            .filter(|value| *value != 0)
            .and_then(char::from_u32)
            .unwrap_or('\u{fffd}');
        return Some((value, prefix + length));
    }
    NAMED_REFERENCES
        .iter()
        .filter(|(name, _)| text.starts_with(name))
        .max_by_key(|(name, _)| name.len())
        .map(|(name, ch)| (*ch, name.len()))
}

fn percent_decoded(text: &str) -> String {
    let bytes = text.as_bytes();
    let mut output = Vec::with_capacity(bytes.len());
    let mut index = 0;
    while index < bytes.len() {
        let hex = bytes
            .get(index + 1..index + 3)
            .and_then(|hex| std::str::from_utf8(hex).ok())
            .and_then(|hex| u8::from_str_radix(hex, 16).ok());
        match (bytes[index], hex) {
            (b'%', Some(byte)) => {
                output.push(byte);
                index += 3;
            }
            (byte, _) => {
                output.push(byte);
                index += 1;
            }
        }
    }
    String::from_utf8_lossy(&output).into_owned()
}

/// Percent-encodes whitespace and controls in a link target. A Markdown destination also
/// percent-encodes what could end it early and writes `&` as `&amp;`, which Markdown decodes
/// back to `&`.
fn destination(href: &str, markdown: bool) -> String {
    let mut output = String::with_capacity(href.len());
    for ch in href.chars() {
        match ch {
            ch if ch.is_whitespace() || ch.is_control() => {
                let mut buffer = [0u8; 4];
                for byte in ch.encode_utf8(&mut buffer).bytes() {
                    let _ = write!(output, "%{byte:02X}");
                }
            }
            ch if !markdown => output.push(ch),
            '&' => output.push_str("&amp;"),
            '(' => output.push_str("%28"),
            ')' => output.push_str("%29"),
            '<' => output.push_str("%3C"),
            '>' => output.push_str("%3E"),
            '\\' => output.push_str("%5C"),
            '|' => output.push_str("%7C"),
            ch => output.push(ch),
        }
    }
    output
}

/// Escapes the pipes a Markdown table cell would otherwise split at.
fn escape_pipes(text: &str) -> String {
    let mut output = String::with_capacity(text.len());
    let mut backslashes = 0usize;
    for ch in text.chars() {
        if ch == '|' && backslashes.is_multiple_of(2) {
            output.push('\\');
        }
        backslashes = if ch == '\\' { backslashes + 1 } else { 0 };
        output.push(ch);
    }
    output
}

/// Escapes what would open a block construct at the start of a trimmed line.
fn line_start(text: String) -> String {
    let digits = text.chars().take_while(char::is_ascii_digit).count();
    if digits > 0
        && digits <= 9
        && text[digits..].starts_with(['.', ')'])
        && (text.len() == digits + 1 || text[digits + 1..].starts_with([' ', '\t']))
    {
        return format!("{}\\{}", &text[..digits], &text[digits..]);
    }
    if text.starts_with(['-', '+', '=', '#']) {
        return format!("\\{text}");
    }
    text
}

/// Escaped `text` at the start of a line, its leading spaces and tabs kept as no-break spaces,
/// which cannot indent a code block, and what follows escaped against opening a block.
fn indented_line_start(text: &str) -> String {
    let rest = text.trim_start_matches([' ', '\t']);
    let mut output = String::with_capacity(text.len() + 8);
    for ch in text[..text.len() - rest.len()].chars() {
        output.push_str(if ch == '\t' {
            "\u{a0}\u{a0}\u{a0}\u{a0}"
        } else {
            "\u{a0}"
        });
    }
    output + &line_start(rest.to_owned())
}

fn is_ordered_marker(marker: &str) -> bool {
    let digits = marker.chars().take_while(char::is_ascii_digit).count();
    (1..=9).contains(&digits) && digits + 1 == marker.len() && marker.ends_with(['.', ')'])
}

fn is_title(shape: &ExportShape) -> bool {
    shape.placeholder.as_ref().is_some_and(|placeholder| {
        matches!(
            placeholder.placeholder_type.as_deref(),
            Some("title" | "ctrTitle")
        )
    })
}

fn has_text(paragraph: &ExportParagraph) -> bool {
    paragraph.runs.iter().any(|run| match &run.content {
        ExportRunKind::Text { text } | ExportRunKind::Field { text, .. } => !text.is_empty(),
        ExportRunKind::LineBreak | ExportRunKind::Unsupported { .. } => true,
    })
}

fn story_text(story: &ExportStory) -> bool {
    story.paragraphs.iter().any(has_text)
}

/// The fewest bytes `text` escapes to: a line separator, three bytes, becomes one space and
/// nothing else shrinks.
fn least(text: &str) -> usize {
    text.len() / 3
}

/// The fewest bytes runs render to with whitespace trimmed at line starts: their text, a unit per
/// line break and the comment an unsupported inline leaves.
fn least_runs(runs: &[ExportRun]) -> usize {
    let mut line_start = true;
    let mut total = 0usize;
    for run in runs {
        let bytes = match &run.content {
            ExportRunKind::Text { text } | ExportRunKind::Field { text, .. } if line_start => {
                least(text.trim_start())
            }
            ExportRunKind::Text { text } | ExportRunKind::Field { text, .. } => least(text),
            ExportRunKind::LineBreak => 1,
            ExportRunKind::Unsupported { element } => least(element),
        };
        line_start = run.content == ExportRunKind::LineBreak || (line_start && bytes == 0);
        total = total.saturating_add(bytes);
    }
    total
}

/// The fewest bytes a table renders to: its shown cells' runs and a delimiter per cell.
fn least_table(table: &ExportTable) -> usize {
    table
        .rows
        .iter()
        .flat_map(|row| &row.cells)
        .filter(|cell| !cell.merged)
        .map(|cell| {
            cell.story
                .iter()
                .flat_map(|story| &story.paragraphs)
                .fold(3usize, |least, paragraph| {
                    least.saturating_add(least_runs(&paragraph.runs))
                })
        })
        .fold(0, usize::saturating_add)
}

fn object_label(kind: ExportObjectKind) -> &'static str {
    match kind {
        ExportObjectKind::Picture => "Picture",
        ExportObjectKind::Video => "Video",
        ExportObjectKind::Audio => "Audio",
        ExportObjectKind::Chart => "Chart",
        ExportObjectKind::SmartArt => "SmartArt",
        ExportObjectKind::EmbeddedObject => "Embedded object",
        ExportObjectKind::Unknown => "Unsupported content",
    }
}

fn alternative_text(shape: &ExportShape) -> Option<&str> {
    [shape.description.as_deref(), shape.title.as_deref()]
        .into_iter()
        .flatten()
        .find(|text| !text.trim().is_empty())
}

impl Renderer {
    /// Appends `text` if it fits; otherwise marks the rendering truncated.
    fn push(&mut self, text: &str) -> bool {
        if self.truncated || self.output.len() + text.len() > self.max_bytes {
            self.truncated = true;
            return false;
        }
        self.output.push_str(text);
        true
    }

    /// Whether `pending` more bytes would run past the byte limit.
    fn over(&self, pending: usize) -> bool {
        self.output.len().saturating_add(pending) > self.max_bytes
    }

    /// Appends one rendered block, dropping the markers it recorded when it does not fit. A
    /// block whose text alone, `least` bytes, cannot fit is not rendered at all.
    fn block(&mut self, least: usize, render: impl FnOnce(&mut Self) -> String) -> bool {
        if self.truncated || self.over(least) {
            self.truncated = true;
            return false;
        }
        let anchors = self.anchors.len();
        let omitted = self.omitted.len();
        let in_list = self.in_list;
        let rendered = render(self);
        self.built += rendered.len();
        if !std::mem::take(&mut self.cut) && self.push(&rendered) {
            return true;
        }
        self.truncated = true;
        self.anchors.truncate(anchors);
        self.omitted.truncate(omitted);
        self.in_list = in_list;
        false
    }

    fn end_list(&mut self) {
        if self.in_list {
            self.in_list = false;
            self.push("\n");
        }
    }

    fn lossy(&mut self, message: &'static str) {
        self.lossy.insert(message);
    }

    /// The marker comment for a record, recorded with its anchor.
    fn marker(&mut self, anchor: &PptxAnchor) -> String {
        let marker = format!("pptx-export:{}", self.anchors.len());
        self.anchors.push(MarkdownAnchor {
            marker: marker.clone(),
            anchor: anchor.clone(),
        });
        format!("<!-- {marker} -->")
    }

    /// Output that leaves a list first separates itself from it.
    fn leave_list(&mut self) -> &'static str {
        if std::mem::take(&mut self.in_list) {
            "\n"
        } else {
            ""
        }
    }

    fn slide(&mut self, slide: &ExportSlide) -> bool {
        self.end_list();
        let name = slide.name.as_deref().filter(|name| !name.trim().is_empty());
        self.block(name.map_or(0, least), |renderer| {
            let mut heading = format!("## Slide {}", u64::from(slide.index) + 1);
            if let Some(name) = name {
                let _ = write!(heading, ": {}", escape(name));
            }
            if slide.hidden == Some(true) {
                heading.push_str(" (hidden)");
            }
            let marker = renderer.marker(&slide.anchor);
            format!("{marker}\n{heading}\n\n")
        })
    }

    fn shape(&mut self, shape: &ExportShape) -> bool {
        let title = is_title(shape);
        for story in &shape.stories {
            for paragraph in story
                .paragraphs
                .iter()
                .filter(|paragraph| has_text(paragraph))
            {
                if !self.block(least_runs(&paragraph.runs), |renderer| {
                    renderer.paragraph(paragraph, title)
                }) {
                    return false;
                }
            }
        }
        if let Some(table) = &shape.table
            && !self.block(least_table(table), |renderer| {
                let leave = renderer.leave_list();
                let marker = renderer.marker(&shape.anchor);
                let table = renderer.table(table);
                format!("{leave}{marker}\n{table}\n")
            })
        {
            return false;
        }
        if let Some(object) = &shape.object
            && !self.block(alternative_text(shape).map_or(0, least), |renderer| {
                renderer.object(shape, object)
            })
        {
            return false;
        }
        shape.children.iter().all(|child| self.shape(child))
    }

    fn object(&mut self, shape: &ExportShape, object: &ExportObject) -> String {
        let leave = self.leave_list();
        let marker = self.marker(&shape.anchor);
        let body = match (object.kind, alternative_text(shape)) {
            (ExportObjectKind::Picture, alt) => {
                format!("![{}]()", escape(alt.unwrap_or("picture")))
            }
            (ExportObjectKind::Unknown, alt) => {
                let mut body = format!(
                    "<!-- pptx-unsupported: {} -->",
                    comment_text(&object.element)
                );
                if let Some(alt) = alt {
                    let _ = write!(body, "\n\\[{}\\]", escape(alt));
                }
                body
            }
            (kind, Some(alt)) => format!("\\[{}: {}\\]", object_label(kind), escape(alt)),
            (kind, None) => format!("\\[{}\\]", object_label(kind)),
        };
        format!("{leave}{marker}\n{body}\n\n")
    }

    fn paragraph(&mut self, paragraph: &ExportParagraph, title: bool) -> String {
        let text = self.runs(&paragraph.runs, title);
        let mut output = String::new();
        match &paragraph.list {
            Some(list) if !title => {
                let indent = "    ".repeat(paragraph.level.min(8) as usize);
                let marker = self.list_marker(list);
                let anchor = self.marker(&paragraph.anchor);
                let _ = write!(
                    output,
                    "{indent}{anchor}\n{indent}{marker} {}\n",
                    line_start(text.trim_start().to_owned())
                );
                self.in_list = true;
            }
            _ => {
                output.push_str(self.leave_list());
                let anchor = self.marker(&paragraph.anchor);
                if title {
                    if paragraph.list.is_some() {
                        self.lossy("List markers of title paragraphs are not rendered.");
                    }
                    let _ = write!(output, "{anchor}\n### {}\n\n", text.trim_start());
                } else {
                    let _ = write!(
                        output,
                        "{anchor}\n{}\n\n",
                        line_start(text.trim_start().to_owned())
                    );
                }
            }
        }
        output
    }

    fn list_marker(&mut self, list: &ExportList) -> String {
        match list {
            ExportList::Bullet { .. } => "-".to_owned(),
            ExportList::Number { marker: None, .. } => {
                self.lossy("Unresolved list numbers are rendered as bullets.");
                "-".to_owned()
            }
            ExportList::Number {
                marker: Some(marker),
                ..
            } if is_ordered_marker(marker) => marker.clone(),
            ExportList::Number {
                marker: Some(marker),
                ..
            } => {
                self.lossy("List markers Markdown cannot number are kept as literal text.");
                format!("- {}", escape(marker))
            }
        }
    }

    /// The runs' Markdown, cut short once even without its leading whitespace it cannot fit. On a
    /// `single_line` line breaks become spaces; otherwise each line after a break is trimmed and
    /// escaped against opening a block.
    fn runs(&mut self, runs: &[ExportRun], single_line: bool) -> String {
        let mut output = String::new();
        let mut lead = None;
        let mut at_line_start = false;
        for run in runs {
            let broken = run.content == ExportRunKind::LineBreak;
            let mut rendered = if broken && single_line {
                " ".to_owned()
            } else {
                self.run(run, false)
            };
            if std::mem::take(&mut at_line_start) {
                let trimmed = rendered.trim_start();
                at_line_start = trimmed.is_empty();
                rendered = if at_line_start {
                    String::new()
                } else {
                    line_start(trimmed.to_owned())
                };
            }
            at_line_start |= broken && !single_line;
            if lead.is_none() && !rendered.trim_start().is_empty() {
                lead = Some(output.len() + rendered.len() - rendered.trim_start().len());
            }
            output.push_str(&rendered);
            if self.over(output.len() - lead.unwrap_or(output.len())) {
                self.cut = true;
                break;
            }
        }
        output
    }

    fn run(&mut self, run: &ExportRun, html: bool) -> String {
        let body = match &run.content {
            ExportRunKind::Text { text } | ExportRunKind::Field { text, .. } => {
                if html {
                    html_text(text)
                } else {
                    escape(text)
                }
            }
            ExportRunKind::LineBreak if html => "<br>".to_owned(),
            ExportRunKind::LineBreak => "\\\n".to_owned(),
            ExportRunKind::Unsupported { element } => {
                format!("<!-- pptx-unsupported: {} -->", comment_text(element))
            }
        };
        if body.is_empty()
            || !matches!(
                run.content,
                ExportRunKind::Text { .. } | ExportRunKind::Field { .. }
            )
        {
            return body;
        }
        let text = self.marked(body, run.marks.as_deref().unwrap_or_default(), html);
        match &run.link {
            None => text,
            Some(link) if !link.external => {
                self.lossy("Links within the presentation are rendered as plain text.");
                text
            }
            Some(link) => {
                match safe_href(&link.href) {
                    None => {
                        self.lossy("Link targets whose scheme is not http, https or mailto are not linked.");
                        text
                    }
                    Some(href) if html => {
                        format!(
                            "<a href=\"{}\">{text}</a>",
                            attribute(&destination(href, false))
                        )
                    }
                    Some(href) => format!("[{text}]({})", destination(href, true)),
                }
            }
        }
    }

    /// Wraps `text` in its marks, keeping surrounding whitespace outside Markdown delimiters.
    fn marked(&mut self, text: String, marks: &[ExportMark], html: bool) -> String {
        if marks.is_empty() || text.trim().is_empty() {
            return text;
        }
        let leading = text.len() - text.trim_start().len();
        let trailing = text.len() - text.trim_end().len();
        let mut core = text[leading..text.len() - trailing].to_owned();
        for mark in marks {
            core = match (mark, html) {
                (ExportMark::Bold, false) => format!("**{core}**"),
                (ExportMark::Bold, true) => format!("<strong>{core}</strong>"),
                (ExportMark::Italic, false) => format!("*{core}*"),
                (ExportMark::Italic, true) => format!("<em>{core}</em>"),
                (ExportMark::Underline { .. }, _) => format!("<u>{core}</u>"),
                (ExportMark::Superscript, _) => format!("<sup>{core}</sup>"),
                (ExportMark::Subscript, _) => format!("<sub>{core}</sub>"),
                (ExportMark::SmallCaps | ExportMark::AllCaps, _) => {
                    self.lossy("Capitalization is rendered as authored.");
                    core
                }
            };
        }
        format!(
            "{}{core}{}",
            &text[..leading],
            &text[text.len() - trailing..]
        )
    }

    /// A pipe table when every shown cell is at most one line of text on an unmerged grid, and
    /// HTML otherwise.
    fn table(&mut self, table: &ExportTable) -> String {
        let cells = || table.rows.iter().flat_map(|row| &row.cells);
        let simple = !table.rows.is_empty()
            && table
                .rows
                .iter()
                .all(|row| row.cells.len() == table.columns as usize)
            && cells().all(|cell| {
                !cell.merged
                    && cell.grid_span == 1
                    && cell.row_span == 1
                    && cell.story.as_ref().is_none_or(|story| {
                        story
                            .paragraphs
                            .iter()
                            .filter(|paragraph| has_text(paragraph))
                            .count()
                            <= 1
                            && story.paragraphs.iter().all(|paragraph| {
                                paragraph.list.is_none()
                                    && paragraph
                                        .runs
                                        .iter()
                                        .all(|run| run.content != ExportRunKind::LineBreak)
                            })
                    })
            });
        for cell in cells().filter(|cell| cell.merged) {
            if cell.story.as_ref().is_some_and(story_text) {
                self.omitted.push(ExportDiagnostic {
                    code: ExportDiagnosticCode::MergeContinuationContentOmitted,
                    severity: ExportSeverity::Warning,
                    anchor: Some(cell.anchor.clone()),
                    message: "A merged-away table cell holds text PowerPoint does not show; the \
                              Markdown leaves it out."
                        .to_owned(),
                });
            }
        }
        if !simple {
            return self.html_table(table);
        }
        let mut output = String::new();
        for (index, row) in table.rows.iter().enumerate() {
            output.push('|');
            for cell in &row.cells {
                let mut text = String::new();
                for paragraph in cell
                    .story
                    .iter()
                    .flat_map(|story| &story.paragraphs)
                    .filter(|paragraph| has_text(paragraph))
                {
                    text.push_str(&self.marker(&paragraph.anchor));
                    text.push_str(&self.runs(&paragraph.runs, true));
                }
                let _ = write!(output, " {} |", escape_pipes(&text));
                if self.over(output.len()) {
                    self.cut = true;
                    return output;
                }
            }
            output.push('\n');
            if index == 0 {
                output.push('|');
                for _ in &row.cells {
                    output.push_str(" --- |");
                }
                output.push('\n');
            }
        }
        output
    }

    fn html_table(&mut self, table: &ExportTable) -> String {
        let mut output = String::from("<table>\n");
        for row in &table.rows {
            output.push_str("<tr>");
            for cell in row.cells.iter().filter(|cell| !cell.merged) {
                let mut attributes = String::new();
                if cell.grid_span > 1 {
                    let _ = write!(attributes, " colspan=\"{}\"", cell.grid_span);
                }
                if cell.row_span > 1 {
                    let _ = write!(attributes, " rowspan=\"{}\"", cell.row_span);
                }
                let mut content = String::new();
                for paragraph in cell
                    .story
                    .iter()
                    .flat_map(|story| &story.paragraphs)
                    .filter(|paragraph| has_text(paragraph))
                {
                    content.push_str(&self.marker(&paragraph.anchor));
                    let marker = match &paragraph.list {
                        Some(ExportList::Bullet { character, .. }) => {
                            format!("{} ", html_text(character))
                        }
                        Some(ExportList::Number {
                            marker: Some(marker),
                            ..
                        }) => format!("{} ", html_text(marker)),
                        Some(ExportList::Number { marker: None, .. }) => {
                            self.lossy("Unresolved list numbers are left out of HTML tables.");
                            String::new()
                        }
                        None => String::new(),
                    };
                    let mut runs = String::new();
                    for run in &paragraph.runs {
                        runs.push_str(&self.run(run, true));
                        if self.over(output.len() + runs.len()) {
                            self.cut = true;
                            return output;
                        }
                    }
                    let _ = write!(content, "<p>{marker}{runs}</p>");
                }
                let _ = write!(output, "<td{attributes}>{content}</td>");
                if self.over(output.len()) {
                    self.cut = true;
                    return output;
                }
            }
            output.push_str("</tr>\n");
        }
        output.push_str("</table>\n");
        output
    }

    fn notes(&mut self, notes: &ExportNotes) -> bool {
        self.end_list();
        let least = notes
            .text
            .chars()
            .filter(|ch| !ch.is_whitespace())
            .map(char::len_utf8)
            .sum();
        self.block(least, |renderer| {
            let marker = renderer.marker(&notes.anchor);
            let quoted = notes
                .text
                .split('\n')
                .map(|line| {
                    if line.trim().is_empty() {
                        ">".to_owned()
                    } else {
                        format!("> {}", indented_line_start(&escape(line)))
                    }
                })
                .collect::<Vec<_>>()
                .join("\n");
            format!("**Notes**\n\n{marker}\n{quoted}\n\n")
        })
    }

    fn comment(&mut self, comment: &ExportComment, first: bool) -> bool {
        if first {
            self.end_list();
        }
        self.block(least(comment.text.trim_start()), |renderer| {
            let heading = if first { "**Comments**\n\n" } else { "" };
            let marker = renderer.marker(&comment.anchor);
            let indent = if comment.parent_id.is_some() {
                "    "
            } else {
                ""
            };
            let mut attribution = Vec::new();
            if let Some(author) = &comment.author {
                attribution.push(format!("**{}**", escape(author)));
            }
            let details: Vec<String> = comment
                .date
                .iter()
                .map(|date| escape(date))
                .chain(comment.resolved.then(|| "resolved".to_owned()))
                .collect();
            if !details.is_empty() {
                attribution.push(format!("({})", details.join(", ")));
            }
            let text = comment
                .text
                .split('\n')
                .map(escape)
                .collect::<Vec<_>>()
                .join("<br>");
            let line = if attribution.is_empty() {
                indented_line_start(&text)
            } else {
                format!("{}: {text}", attribution.join(" "))
            };
            renderer.in_list = true;
            format!("{heading}{indent}{marker}\n{indent}- {line}\n")
        })
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::structured::walk::tests::{deck_with_shapes, deck_with_table};

    fn exported(bytes: &[u8]) -> PptxStructuredContent {
        export_pptx_structured(bytes, &PptxExportOptions::default()).unwrap()
    }

    /// Slide 1 holding one text box whose single paragraph is `runs`.
    fn paragraph(runs: &str) -> PptxStructuredContent {
        exported(&deck_with_shapes(&format!(
            r#"<p:sp><p:nvSpPr><p:cNvPr id="2" name="Body"/><p:cNvSpPr txBox="1"/><p:nvPr/></p:nvSpPr><p:spPr/><p:txBody><a:bodyPr/><a:lstStyle/><a:p>{runs}</a:p></p:txBody></p:sp>"#
        )))
    }

    fn rendered(content: &PptxStructuredContent, max_bytes: u32) -> (PptxMarkdownContent, usize) {
        render(
            content,
            &PptxMarkdownOptions {
                max_bytes: Some(max_bytes),
            },
        )
        .unwrap()
    }

    #[test]
    fn text_before_a_link_cannot_make_it_an_image() {
        let mut content = paragraph("<a:r><a:t>see!Docs</a:t></a:r>");
        let runs = &mut content.slides[0].shapes[0].stories[0].paragraphs[0].runs;
        let mut linked = runs[0].clone();
        for (run, text, first) in [(&mut runs[0], "see!", true), (&mut linked, "Docs", false)] {
            if let PptxAnchor::Range(range) = &mut run.anchor {
                if first {
                    range.end = range.start + 4;
                } else {
                    range.start += 4;
                }
            }
            run.content = ExportRunKind::Text {
                text: text.to_owned(),
            };
        }
        linked.link = Some(ExportLink {
            href: "https://example.test/pixel".to_owned(),
            external: true,
        });
        runs.push(linked);
        let (markdown, _) = rendered(&content, 1 << 20);
        assert!(
            markdown
                .markdown
                .contains(r"see\![Docs](https://example.test/pixel)"),
            "{}",
            markdown.markdown
        );
    }

    #[test]
    fn a_block_too_long_for_the_limit_is_not_built() {
        let content = paragraph(&format!("<a:r><a:t>{}</a:t></a:r>", "a".repeat(1 << 20)));
        let (markdown, built) = rendered(&content, 1_024);
        assert!(markdown.truncated && !markdown.markdown.contains("aaaa"));
        assert!(built < 1_024, "{built} bytes built");
    }

    #[test]
    fn blocks_stop_building_once_they_pass_the_limit() {
        let content = paragraph(&"<a:r><a:t>**********</a:t></a:r><a:br/>".repeat(1_000));
        let (markdown, built) = rendered(&content, 16_384);
        assert!(markdown.truncated && !markdown.markdown.contains(r"\*"));
        assert!(built < 16_384 + 1_024, "{built} bytes built");
        let content = exported(&deck_with_table(2_000));
        let (markdown, built) = rendered(&content, 65_536);
        assert!(markdown.truncated && !markdown.markdown.contains("left 0"));
        assert!(built < 65_536 + 1_024, "{built} bytes built");
    }
}
