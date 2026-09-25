//! Markdown rendering of structured content. The rendering keeps story boundaries, heading,
//! list and table semantics and markup-view revisions; it never evaluates fields, fetches
//! relationships or claims to preserve Word layout. Tables Markdown cannot express are rendered
//! as HTML, whose text is entity-escaped rather than Markdown-escaped. Every document string is
//! inert: no line break survives, so none can open a block, and only http, https, mailto and
//! internal-anchor links are linked, judged after the decoding a consumer applies.

use std::collections::BTreeSet;
use std::fmt::Write as _;

use super::{
    Block, BlockKind, BreakType, CachedResult, Diagnostic, DiagnosticCode, DocxStructuredContent,
    ExportFailure, FormattingMark, HeadingInfo, Inline, InlineKind, Link, ListInfo, MarkdownAnchor,
    MarkdownContent, MarkdownOptions, NoteKind, Revision, RevisionKind, Severity, StoryKind,
    TableData, VerticalMerge, byte_limit,
};

/// Renders `content` as Markdown within `options.max_bytes`, stopping at a block boundary.
pub fn render_docx_markdown(
    content: &DocxStructuredContent,
    options: &MarkdownOptions,
) -> Result<MarkdownContent, ExportFailure> {
    let max_bytes = byte_limit(options.max_bytes)?;
    let mut renderer = Renderer {
        output: String::new(),
        anchors: Vec::new(),
        lossy: BTreeSet::new(),
        max_bytes,
        truncated: false,
        in_list: false,
    };
    'stories: for story in &content.stories {
        let kind = match story.kind {
            StoryKind::Body => "body",
            StoryKind::Header => "header",
            StoryKind::Footer => "footer",
            StoryKind::Footnote => "footnote",
            StoryKind::Endnote => "endnote",
            StoryKind::Comment => "comment",
        };
        let header = format!(
            "<!-- docx-story: {} {kind} -->\n\n",
            comment_text(&story.story)
        );
        if !renderer.push(&header) {
            break;
        }
        let label = match (story.kind, &story.note_id) {
            (StoryKind::Footnote, Some(id)) => Some(format!("[^fn-{}]: ", label_text(id))),
            (StoryKind::Endnote, Some(id)) => Some(format!("[^en-{}]: ", label_text(id))),
            _ => None,
        };
        for (index, block) in story.blocks.iter().enumerate() {
            let prefix = if index == 0 { label.as_deref() } else { None };
            if !renderer.block(block, prefix) {
                break 'stories;
            }
        }
        renderer.end_list();
    }
    let mut diagnostics = content.diagnostics.clone();
    for (code, message) in &renderer.lossy {
        diagnostics.push(Diagnostic {
            code: *code,
            severity: Severity::Info,
            anchor: None,
            message: (*message).to_owned(),
        });
    }
    if renderer.truncated {
        diagnostics.push(Diagnostic {
            code: DiagnosticCode::Truncated,
            severity: Severity::Warning,
            anchor: None,
            message:
                "The Markdown stopped at its byte limit; content after this point is not included."
                    .to_owned(),
        });
    }
    let markdown = renderer.output.trim_end().to_owned() + "\n";
    Ok(MarkdownContent {
        markdown: if renderer.output.is_empty() {
            String::new()
        } else {
            markdown
        },
        anchors: renderer.anchors,
        diagnostics,
        truncated: content.truncated || renderer.truncated,
    })
}

const UNLINKED: &str =
    "Link targets other than http, https, mailto and internal anchors are not linked.";

struct Renderer {
    output: String,
    anchors: Vec<MarkdownAnchor>,
    lossy: BTreeSet<(DiagnosticCode, &'static str)>,
    max_bytes: usize,
    truncated: bool,
    in_list: bool,
}

/// Whether a consumer can read `ch` as ending a line.
fn line_break(ch: char) -> bool {
    matches!(ch, '\r' | '\n' | '\u{2028}' | '\u{2029}')
}

/// Text safe inside an HTML comment.
fn comment_text(text: &str) -> String {
    let mut output = html_text(text);
    while output.contains("--") {
        output = output.replace("--", "- -");
    }
    output
}

/// Text safe inside a footnote label.
fn label_text(text: &str) -> String {
    text.chars()
        .map(|ch| if ch.is_alphanumeric() { ch } else { '-' })
        .collect()
}

/// Escapes Markdown and HTML syntax in document text, including a trailing `!` that a
/// following link would turn into an image.
fn escape(text: &str) -> String {
    let mut output = String::with_capacity(text.len());
    let mut chars = text.chars().peekable();
    while let Some(ch) = chars.next() {
        let entity = ch == '&'
            && chars
                .peek()
                .is_some_and(|next| next.is_ascii_alphanumeric() || *next == '#');
        if line_break(ch) {
            output.push(' ');
            continue;
        }
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

/// Escapes text for a double-quoted Markdown link title.
fn title_text(text: &str) -> String {
    let mut output = String::with_capacity(text.len());
    for ch in text.chars() {
        match ch {
            '\\' | '"' | '<' | '>' | '&' => {
                output.push('\\');
                output.push(ch);
            }
            ch if line_break(ch) => output.push(' '),
            ch => output.push(ch),
        }
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
            ch if line_break(ch) => output.push(' '),
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
            ch if line_break(ch) => output.push(' '),
            ch => output.push(ch),
        }
    }
    output
}

/// `href` as a link destination, `&` still unescaped, when what a consumer resolves it to is an
/// http, https or mailto URL or an internal anchor.
fn link_target(href: &str) -> Option<String> {
    let effective = effective_destination(href)?.to_ascii_lowercase();
    let allowed = effective.starts_with('#')
        || ["http:", "https:", "mailto:"]
            .iter()
            .any(|scheme| effective.starts_with(scheme));
    allowed.then(|| destination(href))
}

/// What a consumer reads the scheme of `href` from: ASCII whitespace and control characters
/// removed, HTML entities decoded (named, decimal and hex, with or without the semicolon) and
/// percent escapes decoded, until nothing changes. `None` when that does not settle.
fn effective_destination(href: &str) -> Option<String> {
    let mut current = href.to_owned();
    for _ in 0..8 {
        let stripped: String = current
            .chars()
            .filter(|ch| !ch.is_ascii_whitespace() && !ch.is_control())
            .collect();
        let next = percent_decoded(&entities_decoded(&stripped));
        if next == current {
            return Some(next);
        }
        current = next;
    }
    None
}

fn entities_decoded(text: &str) -> String {
    let mut output = String::with_capacity(text.len());
    let mut rest = text;
    while let Some(at) = rest.find('&') {
        output.push_str(&rest[..at]);
        rest = &rest[at..];
        match entity(rest) {
            Some((ch, len)) => {
                output.push(ch);
                rest = &rest[len..];
            }
            None => {
                output.push('&');
                rest = &rest[1..];
            }
        }
    }
    output.push_str(rest);
    output
}

/// The character the entity reference opening `text` stands for, and the reference's length.
fn entity(text: &str) -> Option<(char, usize)> {
    let semicolon = |len: usize| len + usize::from(text[len..].starts_with(';'));
    if let Some(numeric) = text.strip_prefix("&#") {
        let (radix, start) = match numeric.strip_prefix(['x', 'X']) {
            Some(_) => (16, 3),
            None => (10, 2),
        };
        let digits = text[start..]
            .chars()
            .take_while(|ch| ch.is_digit(radix))
            .count();
        if digits == 0 {
            return None;
        }
        let ch = u32::from_str_radix(&text[start..start + digits], radix)
            .ok()
            .and_then(char::from_u32)
            .unwrap_or('\u{FFFD}');
        return Some((ch, semicolon(start + digits)));
    }
    let name = text[1..]
        .chars()
        .take_while(char::is_ascii_alphanumeric)
        .take(LONGEST_ENTITY_NAME)
        .count();
    (1..=name).rev().find_map(|end| {
        named_entity(&text[1..=end].to_ascii_lowercase()).map(|ch| (ch, semicolon(1 + end)))
    })
}

/// The length of the longest name `named_entity` knows, which bounds the prefixes tried.
const LONGEST_ENTITY_NAME: usize = 7;

fn named_entity(name: &str) -> Option<char> {
    Some(match name {
        "colon" => ':',
        "tab" => '\t',
        "newline" => '\n',
        "nbsp" => '\u{A0}',
        "amp" => '&',
        "lt" => '<',
        "gt" => '>',
        "quot" => '"',
        "apos" => '\'',
        "lpar" => '(',
        "rpar" => ')',
        "sol" => '/',
        "bsol" => '\\',
        "num" => '#',
        "percnt" => '%',
        "period" => '.',
        "plus" => '+',
        "hyphen" | "dash" => '-',
        "excl" => '!',
        "commat" => '@',
        "semi" => ';',
        "equals" => '=',
        "quest" => '?',
        "comma" => ',',
        "lowbar" => '_',
        "ast" => '*',
        "dollar" => '$',
        "grave" => '`',
        "verbar" | "vert" => '|',
        "lsqb" | "lbrack" => '[',
        "rsqb" | "rbrack" => ']',
        "lcub" | "lbrace" => '{',
        "rcub" | "rbrace" => '}',
        "hat" => '^',
        _ => return None,
    })
}

fn percent_decoded(text: &str) -> String {
    let bytes = text.as_bytes();
    let hex = |byte: u8| (byte as char).to_digit(16);
    let mut output = Vec::with_capacity(bytes.len());
    let mut index = 0;
    while index < bytes.len() {
        let escape = (bytes[index] == b'%')
            .then(|| Some(hex(*bytes.get(index + 1)?)? * 16 + hex(*bytes.get(index + 2)?)?))
            .flatten();
        match escape {
            Some(value) => {
                output.push(value as u8);
                index += 3;
            }
            None => {
                output.push(bytes[index]);
                index += 1;
            }
        }
    }
    String::from_utf8_lossy(&output).into_owned()
}

/// Escapes the pipes a Markdown table cell would otherwise split at.
fn escape_pipes(text: &str) -> String {
    let mut output = String::with_capacity(text.len());
    let mut backslashes = 0usize;
    for ch in text.chars() {
        if ch == '|' && backslashes % 2 == 0 {
            output.push('\\');
        }
        backslashes = if ch == '\\' { backslashes + 1 } else { 0 };
        output.push(ch);
    }
    output
}

/// Percent-encodes what would end a link destination early. `&` is left for the caller to
/// write as `&amp;`, so no entity decoding can change the destination.
fn destination(href: &str) -> String {
    let mut output = String::with_capacity(href.len());
    for ch in href.chars() {
        match ch {
            '(' => output.push_str("%28"),
            ')' => output.push_str("%29"),
            '<' => output.push_str("%3C"),
            '>' => output.push_str("%3E"),
            '\\' => output.push_str("%5C"),
            '|' => output.push_str("%7C"),
            ch if ch.is_control() || ch.is_whitespace() => {
                let mut buffer = [0u8; 4];
                for byte in ch.encode_utf8(&mut buffer).bytes() {
                    let _ = write!(output, "%{byte:02X}");
                }
            }
            ch => output.push(ch),
        }
    }
    output
}

/// Escapes what would open a block construct at the start of a line.
fn line_start(text: String) -> String {
    let digits = text.chars().take_while(char::is_ascii_digit).count();
    if digits > 0
        && digits <= 9
        && text[digits..].starts_with(['.', ')'])
        && text[digits + 1..].starts_with([' ', '\t'])
    {
        return format!("{}\\{}", &text[..digits], &text[digits..]);
    }
    if text.starts_with(['-', '+', '=', '#']) || text.starts_with("    ") {
        return format!("\\{text}");
    }
    text
}

/// `text` with every line after a hard line break escaped as [`line_start`] escapes the first.
fn lines(text: &str) -> String {
    text.split("\\\n")
        .map(|line| line_start(line.to_owned()))
        .collect::<Vec<_>>()
        .join("\\\n")
}

fn is_ordered_marker(marker: &str) -> bool {
    let digits = marker.chars().take_while(char::is_ascii_digit).count();
    (1..=9).contains(&digits) && digits + 1 == marker.len() && marker.ends_with(['.', ')'])
}

fn revision_attributes(revision: &Revision) -> String {
    let mut attributes = String::new();
    match revision.kind {
        RevisionKind::MoveFrom => attributes.push_str(" data-move=\"from\""),
        RevisionKind::MoveTo => attributes.push_str(" data-move=\"to\""),
        RevisionKind::Insertion | RevisionKind::Deletion => {}
    }
    if let Some(author) = &revision.author {
        let _ = write!(attributes, " data-author=\"{}\"", attribute(author));
    }
    if let Some(date) = &revision.date {
        let _ = write!(attributes, " data-date=\"{}\"", attribute(date));
    }
    attributes
}

fn revision_tag(revision: &Revision) -> &'static str {
    match revision.kind {
        RevisionKind::Insertion | RevisionKind::MoveTo => "ins",
        RevisionKind::Deletion | RevisionKind::MoveFrom => "del",
    }
}

/// A page or column break inside a paragraph, which Markdown cannot show.
fn break_comment(break_type: BreakType) -> String {
    match break_type {
        BreakType::Column => "<!-- docx-break: column -->".to_owned(),
        _ => "<!-- docx-break: page -->".to_owned(),
    }
}

fn heading_level(heading: &HeadingInfo) -> usize {
    usize::from(heading.outline_level.min(5)) + 1
}

/// The blocks a paragraph's fields return as block results, which render after it.
fn result_blocks(inlines: &[Inline]) -> Vec<&Block> {
    let mut blocks = Vec::new();
    for inline in inlines {
        match &inline.content {
            InlineKind::Field {
                cached_result: CachedResult::Blocks { blocks: result },
                ..
            } => blocks.extend(result),
            InlineKind::Field {
                cached_result: CachedResult::Inline { inlines },
                ..
            }
            | InlineKind::ContentControl { inlines, .. } => blocks.extend(result_blocks(inlines)),
            _ => {}
        }
    }
    blocks
}

fn paragraph_of(block: &Block) -> Option<&[Inline]> {
    match &block.content {
        BlockKind::Paragraph { paragraph }
        | BlockKind::Heading { paragraph, .. }
        | BlockKind::ListItem { paragraph, .. } => Some(&paragraph.inlines),
        _ => None,
    }
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

    fn end_list(&mut self) {
        if self.in_list {
            self.in_list = false;
            self.push("\n");
        }
    }

    fn lossy(&mut self, message: &'static str) {
        self.lossy.insert((DiagnosticCode::MarkdownLossy, message));
    }

    /// The marker comment for `block`, recorded with its anchor.
    fn marker(&mut self, block: &Block) -> String {
        let marker = format!("docx-export:{}", self.anchors.len());
        self.anchors.push(MarkdownAnchor {
            marker: marker.clone(),
            anchor: block.anchor.clone(),
        });
        format!("<!-- {marker} -->")
    }

    /// Renders one root block with its marker, all or nothing.
    fn block(&mut self, block: &Block, prefix: Option<&str>) -> bool {
        let anchors = self.anchors.len();
        let in_list = self.in_list;
        let rendered = self.render_block(block, prefix);
        if self.push(&rendered) {
            return true;
        }
        self.anchors.truncate(anchors);
        self.in_list = in_list;
        false
    }

    fn render_block(&mut self, block: &Block, prefix: Option<&str>) -> String {
        let mut output = String::new();
        let list_item = matches!(block.content, BlockKind::ListItem { .. });
        if self.in_list && !list_item {
            self.in_list = false;
            output.push('\n');
        }
        if let BlockKind::ListItem { list, .. } = &block.content {
            output.push_str(&"    ".repeat(usize::from(list.level)));
        }
        output.push_str(&self.marker(block));
        output.push('\n');
        let prefix = prefix.unwrap_or_default();
        match &block.content {
            BlockKind::Paragraph { paragraph } => {
                let text = self.inlines(&paragraph.inlines);
                output.push_str(prefix);
                output.push_str(&lines(&text));
                output.push_str("\n\n");
            }
            BlockKind::Heading { paragraph, heading } => {
                if heading.outline_level > 5 {
                    self.lossy("Heading levels deeper than 6 are rendered as level 6.");
                }
                let text = self.inlines(&paragraph.inlines).replace("\\\n", "<br>");
                output.push_str(prefix);
                let _ = write!(
                    output,
                    "{} {}\n\n",
                    "#".repeat(heading_level(heading)),
                    text.trim_start()
                );
            }
            BlockKind::ListItem {
                paragraph, list, ..
            } => {
                let text = lines(&self.inlines(&paragraph.inlines));
                let marker = self.list_marker(list);
                output.push_str(prefix);
                let _ = writeln!(
                    output,
                    "{}{marker} {}",
                    "    ".repeat(usize::from(list.level)),
                    text.trim_start()
                );
                self.in_list = true;
            }
            BlockKind::Table { table } => {
                output.push_str(&self.table(table));
                output.push('\n');
            }
            BlockKind::ContentControl { blocks, .. } => {
                for block in blocks {
                    let rendered = self.render_block(block, None);
                    output.push_str(&rendered);
                }
            }
            BlockKind::SectionBreak { break_type, .. } => {
                let kind = serde_json::to_value(break_type)
                    .ok()
                    .and_then(|value| value.as_str().map(str::to_owned))
                    .unwrap_or_default();
                let _ = write!(output, "<!-- docx-section-break: {kind} -->\n\n");
            }
            BlockKind::Break { break_type } => {
                let kind = serde_json::to_value(break_type)
                    .ok()
                    .and_then(|value| value.as_str().map(str::to_owned))
                    .unwrap_or_default();
                let _ = write!(output, "<!-- docx-break: {kind} -->\n\n");
            }
            BlockKind::Unsupported { element } => {
                let _ = write!(
                    output,
                    "<!-- docx-unsupported: {} -->\n\n",
                    comment_text(element)
                );
            }
        }
        if let Some(inlines) = paragraph_of(block) {
            let results = result_blocks(inlines);
            if !results.is_empty() {
                self.lossy("Field results that span blocks are rendered after their paragraph.");
                for result in results {
                    let rendered = self.render_block(result, None);
                    output.push_str(&rendered);
                }
            }
        }
        output
    }

    fn list_marker(&mut self, list: &ListInfo) -> String {
        match list.marker.as_deref() {
            None => {
                self.lossy("Unresolved list markers are rendered as bullets.");
                "-".to_owned()
            }
            Some(_) if list.marker_hidden => {
                self.lossy("Hidden list markers are rendered as bullets.");
                "-".to_owned()
            }
            Some(marker) if list.format == "bullet" || marker.is_empty() => "-".to_owned(),
            Some(marker) if is_ordered_marker(marker) => marker.to_owned(),
            Some(marker) => {
                self.lossy("List markers Markdown cannot number are kept as literal text.");
                format!("- {}", escape(marker))
            }
        }
    }

    /// A table as a Markdown pipe table when every cell is at most one plain paragraph on an
    /// unmerged grid, and as HTML otherwise.
    fn table(&mut self, table: &TableData) -> String {
        let simple = !table.rows.is_empty()
            && table.rows.iter().all(|row| {
                row.grid_before == 0
                    && row.grid_after == 0
                    && row.cells.len() == table.rows[0].cells.len()
                    && row.cells.iter().all(|cell| {
                        cell.grid_span == 1
                            && cell.row_span == 1
                            && cell.vertical_merge == VerticalMerge::None
                            && cell.blocks.len() <= 1
                            && cell.blocks.iter().all(|block| {
                                matches!(block.content, BlockKind::Paragraph { .. })
                                    && paragraph_of(block)
                                        .is_some_and(|inlines| result_blocks(inlines).is_empty())
                            })
                    })
            });
        if !simple {
            return self.html_table(table);
        }
        let mut output = String::new();
        for (index, row) in table.rows.iter().enumerate() {
            output.push('|');
            for cell in &row.cells {
                let mut text = String::new();
                for block in &cell.blocks {
                    text.push_str(&self.marker(block));
                    if let BlockKind::Paragraph { paragraph } = &block.content {
                        let inlines = self.inlines(&paragraph.inlines);
                        text.push_str(&inlines.replace("\\\n", "<br>").replace('\n', " "));
                    }
                }
                let _ = write!(output, " {} |", escape_pipes(&text));
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

    fn html_table(&mut self, table: &TableData) -> String {
        let mut output = String::from("<table>\n");
        for row in &table.rows {
            output.push_str("<tr>");
            let tag = if row.header { "th" } else { "td" };
            for cell in row.cells.iter().filter(|cell| cell.row_span > 0) {
                let mut attributes = String::new();
                if cell.grid_span > 1 {
                    let _ = write!(attributes, " colspan=\"{}\"", cell.grid_span);
                }
                if cell.row_span > 1 {
                    let _ = write!(attributes, " rowspan=\"{}\"", cell.row_span);
                }
                let content = self.html_blocks(&cell.blocks);
                let _ = write!(output, "<{tag}{attributes}>{content}</{tag}>");
            }
            output.push_str("</tr>\n");
        }
        output.push_str("</table>\n");
        if table
            .rows
            .iter()
            .any(|row| row.grid_before > 0 || row.grid_after > 0)
        {
            self.lossy("Skipped grid columns are not rendered in HTML tables.");
        }
        output
    }

    /// Blocks inside an HTML table cell, each preceded by its marker.
    fn html_blocks<'b>(&mut self, blocks: impl IntoIterator<Item = &'b Block>) -> String {
        let mut output = String::new();
        for block in blocks {
            output.push_str(&self.marker(block));
            match &block.content {
                BlockKind::Paragraph { paragraph } => {
                    let _ = write!(output, "<p>{}</p>", self.html_inlines(&paragraph.inlines));
                }
                BlockKind::Heading { paragraph, heading } => {
                    let level = heading_level(heading);
                    let _ = write!(
                        output,
                        "<h{level}>{}</h{level}>",
                        self.html_inlines(&paragraph.inlines)
                    );
                }
                BlockKind::ListItem {
                    paragraph, list, ..
                } => {
                    let marker = match list.marker.as_deref() {
                        Some(marker) if !list.marker_hidden && !marker.is_empty() => {
                            format!("{} ", html_text(marker))
                        }
                        _ => {
                            self.lossy("List markers without a rendered value are left out of HTML tables.");
                            String::new()
                        }
                    };
                    let _ = write!(
                        output,
                        "<p>{marker}{}</p>",
                        self.html_inlines(&paragraph.inlines)
                    );
                }
                BlockKind::Table { table } => output.push_str(&self.html_table(table)),
                BlockKind::ContentControl { blocks, .. } => {
                    let inner = self.html_blocks(blocks);
                    output.push_str(&inner);
                }
                BlockKind::SectionBreak { .. } => {
                    output.push_str("<!-- docx-section-break -->");
                }
                BlockKind::Break { .. } => output.push_str("<!-- docx-break -->"),
                BlockKind::Unsupported { element } => {
                    let _ = write!(
                        output,
                        "<!-- docx-unsupported: {} -->",
                        comment_text(element)
                    );
                }
            }
            if let Some(inlines) = paragraph_of(block) {
                let results = result_blocks(inlines);
                if !results.is_empty() {
                    self.lossy(
                        "Field results that span blocks are rendered after their paragraph.",
                    );
                    let inner = self.html_blocks(results);
                    output.push_str(&inner);
                }
            }
        }
        output
    }

    fn html_inlines(&mut self, inlines: &[Inline]) -> String {
        let mut output = String::new();
        for inline in inlines {
            let rendered = self.html_inline(inline);
            output.push_str(&rendered);
        }
        output
    }

    fn html_inline(&mut self, inline: &Inline) -> String {
        let body = match &inline.content {
            InlineKind::Text { text } => html_text(text),
            InlineKind::Tab => "\t".to_owned(),
            InlineKind::Break {
                break_type: BreakType::Line,
            } => "<br>".to_owned(),
            InlineKind::Break { break_type } => break_comment(*break_type),
            InlineKind::NoteReference {
                note_kind, note_id, ..
            } => {
                let prefix = match note_kind {
                    NoteKind::Footnote => "fn",
                    NoteKind::Endnote => "en",
                };
                format!("<sup>[{prefix}-{}]</sup>", html_text(&label_text(note_id)))
            }
            InlineKind::CommentReference { .. } => String::new(),
            InlineKind::Field { cached_result, .. } => match cached_result {
                CachedResult::Inline { inlines } => self.html_inlines(inlines),
                CachedResult::Blocks { .. } | CachedResult::Missing => String::new(),
            },
            InlineKind::Image { alt_text, .. } => format!(
                "<img alt=\"{}\">",
                attribute(alt_text.as_deref().unwrap_or("image"))
            ),
            InlineKind::ContentControl { inlines, .. } => self.html_inlines(inlines),
            InlineKind::Unsupported { element, alt_text } => format!(
                "<!-- docx-unsupported: {} -->{}",
                comment_text(element),
                alt_text.as_deref().map(html_text).unwrap_or_default()
            ),
        };
        if body.is_empty() {
            return body;
        }
        let mut text = match &inline.content {
            InlineKind::Text { .. } => {
                self.html_marked(body, inline.marks.as_deref().unwrap_or_default())
            }
            _ => body,
        };
        if let Some(link) = &inline.link {
            text = self.html_link(text, link);
        }
        for revision in &inline.revisions {
            let tag = revision_tag(revision);
            text = format!("<{tag}{}>{text}</{tag}>", revision_attributes(revision));
        }
        text
    }

    fn html_marked(&mut self, text: String, marks: &[FormattingMark]) -> String {
        let mut text = text;
        for mark in marks {
            text = match mark {
                FormattingMark::Bold => format!("<strong>{text}</strong>"),
                FormattingMark::Italic => format!("<em>{text}</em>"),
                FormattingMark::Strike => format!("<s>{text}</s>"),
                FormattingMark::Underline { .. } => format!("<u>{text}</u>"),
                FormattingMark::Subscript => format!("<sub>{text}</sub>"),
                FormattingMark::Superscript => format!("<sup>{text}</sup>"),
                FormattingMark::Hidden => {
                    self.lossy("Hidden text is rendered as visible text.");
                    text
                }
            };
        }
        text
    }

    fn html_link(&mut self, text: String, link: &Link) -> String {
        let Some(href) = link_target(&link.href) else {
            self.lossy(UNLINKED);
            return text;
        };
        let title = link
            .title
            .as_deref()
            .map(|title| format!(" title=\"{}\"", attribute(title)))
            .unwrap_or_default();
        format!("<a href=\"{}\"{title}>{text}</a>", attribute(&href))
    }

    fn inlines(&mut self, inlines: &[Inline]) -> String {
        let mut output = String::new();
        for inline in inlines {
            let rendered = self.inline(inline);
            output.push_str(&rendered);
        }
        output
    }

    fn inline(&mut self, inline: &Inline) -> String {
        let body = match &inline.content {
            InlineKind::Text { text } => escape(text),
            InlineKind::Tab => "\t".to_owned(),
            InlineKind::Break {
                break_type: BreakType::Line,
            } => "\\\n".to_owned(),
            InlineKind::Break { break_type } => break_comment(*break_type),
            InlineKind::NoteReference {
                note_kind, note_id, ..
            } => {
                let prefix = match note_kind {
                    NoteKind::Footnote => "fn",
                    NoteKind::Endnote => "en",
                };
                format!("[^{prefix}-{}]", label_text(note_id))
            }
            InlineKind::CommentReference { .. } => String::new(),
            InlineKind::Field { cached_result, .. } => match cached_result {
                CachedResult::Inline { inlines } => self.inlines(inlines),
                CachedResult::Blocks { .. } | CachedResult::Missing => String::new(),
            },
            InlineKind::Image { alt_text, .. } => {
                format!("![{}]()", escape(alt_text.as_deref().unwrap_or("image")))
            }
            InlineKind::ContentControl { inlines, .. } => self.inlines(inlines),
            InlineKind::Unsupported { element, alt_text } => match alt_text {
                Some(alt) => format!(
                    "<!-- docx-unsupported: {} -->{}",
                    comment_text(element),
                    escape(alt)
                ),
                None => format!("<!-- docx-unsupported: {} -->", comment_text(element)),
            },
        };
        if body.is_empty() {
            return body;
        }
        let mut text = match &inline.content {
            InlineKind::Text { .. } => {
                self.marked(body, inline.marks.as_deref().unwrap_or_default())
            }
            _ => body,
        };
        if let Some(link) = &inline.link {
            text = match link_target(&link.href) {
                None => {
                    self.lossy(UNLINKED);
                    text
                }
                Some(href) => {
                    let href = href.replace('&', "&amp;");
                    match &link.title {
                        Some(title) => format!("[{text}]({href} \"{}\")", title_text(title)),
                        None => format!("[{text}]({href})"),
                    }
                }
            };
        }
        for revision in &inline.revisions {
            let tag = revision_tag(revision);
            text = format!("<{tag}{}>{text}</{tag}>", revision_attributes(revision));
        }
        text
    }

    /// Wraps `text` in its marks, keeping surrounding whitespace outside the delimiters.
    fn marked(&mut self, text: String, marks: &[FormattingMark]) -> String {
        if marks.is_empty() || text.trim().is_empty() {
            return text;
        }
        let leading = text.len() - text.trim_start().len();
        let trailing = text.len() - text.trim_end().len();
        let mut core = text[leading..text.len() - trailing].to_owned();
        for mark in marks {
            core = match mark {
                FormattingMark::Bold => format!("**{core}**"),
                FormattingMark::Italic => format!("*{core}*"),
                FormattingMark::Strike => format!("~~{core}~~"),
                FormattingMark::Underline { .. } => format!("<u>{core}</u>"),
                FormattingMark::Subscript => format!("<sub>{core}</sub>"),
                FormattingMark::Superscript => format!("<sup>{core}</sup>"),
                FormattingMark::Hidden => {
                    self.lossy("Hidden text is rendered as visible text.");
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
}

#[cfg(test)]
mod tests {
    use super::super::{
        AnchorScope, ExportStory, ParagraphData, RevisionView, TableCell, TableRow,
    };
    use super::*;
    use crate::read_types::Anchor;

    fn anchor() -> Anchor {
        Anchor::Paragraph {
            story: "body".to_owned(),
            para_id: String::new(),
        }
    }

    fn text(text: &str, link: Option<Link>) -> Inline {
        Inline {
            id: String::new(),
            anchor: anchor(),
            marks: Some(Vec::new()),
            link,
            revisions: Vec::new(),
            content: InlineKind::Text {
                text: text.to_owned(),
            },
        }
    }

    fn paragraph(inlines: Vec<Inline>) -> Block {
        Block {
            id: String::new(),
            anchor: anchor(),
            content: BlockKind::Paragraph {
                paragraph: ParagraphData {
                    style_id: None,
                    inlines,
                },
            },
        }
    }

    /// `inlines` as a paragraph on its own and inside a table only HTML can render.
    fn render(inlines: Vec<Inline>) -> (String, String) {
        let cell = |blocks: Vec<Block>| TableCell {
            anchor: anchor(),
            story: None,
            column: 0,
            grid_span: 2,
            row_span: 1,
            vertical_merge: VerticalMerge::None,
            merge_origin: None,
            blocks,
        };
        let table = Block {
            id: String::new(),
            anchor: anchor(),
            content: BlockKind::Table {
                table: TableData {
                    grid_columns: 2,
                    rows: vec![TableRow {
                        header: false,
                        grid_before: 0,
                        grid_after: 0,
                        cells: vec![cell(vec![paragraph(inlines.clone())])],
                    }],
                },
            },
        };
        let markdown = |block: Block| {
            let content = DocxStructuredContent {
                schema_version: super::super::SCHEMA_VERSION,
                revision_view: RevisionView::Accepted,
                anchor_scope: AnchorScope::Snapshot,
                included_stories: Vec::new(),
                include_formatting: true,
                stories: vec![ExportStory {
                    story: "body".to_owned(),
                    kind: StoryKind::Body,
                    part: None,
                    note_id: None,
                    comment: None,
                    uses: Vec::new(),
                    blocks: vec![block],
                }],
                diagnostics: Vec::new(),
                truncated: false,
            };
            let mut markdown = render_docx_markdown(&content, &MarkdownOptions::default())
                .unwrap()
                .markdown;
            for marker in [
                "<!-- docx-story: body body -->",
                "<!-- docx-export:0 -->",
                "<!-- docx-export:1 -->",
            ] {
                markdown = markdown.replace(marker, "");
            }
            markdown.trim().to_owned()
        };
        (markdown(paragraph(inlines)), markdown(table))
    }

    fn link(href: &str, title: Option<&str>) -> Option<Link> {
        Some(Link {
            href: href.to_owned(),
            title: title.map(str::to_owned),
        })
    }

    #[test]
    fn script_destinations_are_never_linked_whatever_their_encoding() {
        for href in [
            "javascript&#58;alert(1)",
            "javascript:alert(1)",
            "JaVaScRiPt:alert(1)",
            "java&#x09;script:alert(1)",
            "java&Tab;script:alert(1)",
            "&#106;avascript:alert(1)",
            "&#x6A;avascript:alert(1)",
            "&#0000106avascript:alert(1)",
            "javascript&colon;alert(1)",
            "javascript&amp;#58;alert(1)",
            "javascript%3Aalert(1)",
            "%6Aavascript:alert(1)",
            " javascript:alert(1)",
            "\tjavascript:alert(1)",
            "\u{1}javascript:alert(1)",
            "java\nscript:alert(1)",
            "java\u{0}script:alert(1)",
            "vbscript:msgbox(1)",
            "data:text/html,<script>alert(1)</script>",
            "file:///etc/passwd",
        ] {
            let (markdown, html) = render(vec![text("click", link(href, None))]);
            assert_eq!(markdown, "click", "{href:?}");
            assert_eq!(
                html, "<table>\n<tr><td colspan=\"2\"><p>click</p></td></tr>\n</table>",
                "{href:?}"
            );
        }
    }

    #[test]
    fn allowed_destinations_are_encoded_against_entity_decoding() {
        let (markdown, html) = render(vec![text(
            "go",
            link("https://example.com/a b?x=(1)&y=<2>", None),
        )]);
        assert_eq!(
            markdown,
            "[go](https://example.com/a%20b?x=%281%29&amp;y=%3C2%3E)"
        );
        assert!(
            html.contains(r#"<a href="https://example.com/a%20b?x=%281%29&amp;y=%3C2%3E">go</a>"#),
            "{html}"
        );
        assert_eq!(
            render(vec![text("q", link("https://a.test/?a=1&b=2", None))]).0,
            "[q](https://a.test/?a=1&amp;b=2)"
        );
        for (href, expected) in [
            ("#_Toc1", "[go](#_Toc1)"),
            ("MAILTO:a@example.com", "[go](MAILTO:a@example.com)"),
            (
                "https://a.test/?q=&#58;",
                "[go](https://a.test/?q=&amp;#58;)",
            ),
            ("http://example.com", "[go](http://example.com)"),
        ] {
            assert_eq!(render(vec![text("go", link(href, None))]).0, expected);
        }
    }

    #[test]
    fn titles_and_alt_text_stay_on_one_line_and_inert() {
        let tooltip =
            "tip\n\n<img src=x onerror=alert(1)>\n\n</a><script>alert(1)</script>\u{2028}&amp;";
        let (markdown, html) = render(vec![text(
            "Link",
            link("https://example.com", Some(tooltip)),
        )]);
        assert_eq!(
            markdown,
            r#"[Link](https://example.com "tip  \<img src=x onerror=alert(1)\>  \</a\>\<script\>alert(1)\</script\> \&amp;")"#
        );
        assert!(
            html.contains(r#"title="tip  &lt;img src=x onerror=alert(1)&gt;  &lt;/a&gt;&lt;script&gt;alert(1)&lt;/script&gt; &amp;amp;""#),
            "{html}"
        );
        let image = Inline {
            content: InlineKind::Image {
                alt_text: Some(
                    "</a><script>alert(1)</script>\n\n<img src=x onerror=alert(1)>".to_owned(),
                ),
                relationship_id: None,
                part: None,
                external_target: None,
            },
            ..text("", None)
        };
        let (markdown, html) = render(vec![image]);
        assert_eq!(
            markdown,
            r"![\</a\>\<script\>alert(1)\</script\>  \<img src=x onerror=alert(1)\>]()"
        );
        assert!(
            html.contains(r#"<img alt="&lt;/a&gt;&lt;script&gt;alert(1)&lt;/script&gt;  &lt;img src=x onerror=alert(1)&gt;">"#),
            "{html}"
        );
    }

    #[test]
    fn document_text_cannot_open_a_block() {
        let (markdown, html) = render(vec![text(
            "a\n\n<img src=x onerror=alert(1)>\r\n# b\u{2029}[x](javascript:alert(1))",
            None,
        )]);
        assert_eq!(
            markdown,
            r"a  \<img src=x onerror=alert(1)\>  # b \[x\](javascript:alert(1))"
        );
        assert!(!markdown.contains('\n'));
        assert_eq!(
            html,
            "<table>\n<tr><td colspan=\"2\"><p>a  &lt;img src=x onerror=alert(1)&gt;  # b [x](javascript:alert(1))</p></td></tr>\n</table>"
        );
    }

    #[test]
    fn long_names_after_an_ampersand_decode_in_linear_time() {
        let href = format!("https://example.test/?x=1&{}=1", "A".repeat(1 << 20));
        assert_eq!(effective_destination(&href).as_deref(), Some(href.as_str()));
        assert_eq!(entities_decoded("&newlinex&verbarx&amp"), "\nx|x&");
    }

    #[test]
    fn text_before_a_link_cannot_make_it_an_image() {
        let (markdown, _) = render(vec![
            text("see!", None),
            text("Docs", link("https://example.test/pixel", None)),
        ]);
        assert_eq!(markdown, r"see\![Docs](https://example.test/pixel)");
    }
}
