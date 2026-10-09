//! Markdown projection of a structured export: per sheet one bounded grid labelled with its
//! A1 columns and row numbers, then its tables, links and objects, then the defined names.
//! Every block starts with a `<!-- xlsx-export:N -->` marker whose anchor is returned.
//! Document text is escaped; nothing is evaluated or fetched.

use std::collections::BTreeMap;

use xlsx_model::addr::col_to_letters;
use xlsx_model::{CellRange, CellRef, MAX_COLS, MAX_ROWS};

use super::*;

/// Room kept for closing an open table and the truncation note.
const RESERVE: usize = 160;

pub(super) struct MarkdownLimits {
    rows: usize,
    columns: usize,
    cells: usize,
    bytes: usize,
}

fn refused(code: XlsxExportFailureCode, message: String) -> XlsxExportFailure {
    XlsxExportFailure {
        code,
        target: None,
        message,
    }
}

pub(super) fn limits(
    options: &XlsxMarkdownOptions,
) -> std::result::Result<MarkdownLimits, XlsxExportFailure> {
    let bounded = |value: Option<u32>, default: u32, max: u32, name: &str| {
        let value = value.unwrap_or(default);
        if value == 0 {
            return Err(refused(
                XlsxExportFailureCode::InvalidOptions,
                format!("{name} must be positive"),
            ));
        }
        if value > max {
            return Err(refused(
                XlsxExportFailureCode::LimitExceeded,
                format!("{name} is at most {max}"),
            ));
        }
        Ok(value as usize)
    };
    let bytes = options.max_bytes.unwrap_or(DEFAULT_EXPORT_MAX_BYTES);
    if bytes < MIN_MARKDOWN_BYTES {
        return Err(refused(
            XlsxExportFailureCode::LimitExceeded,
            format!("maxBytes must be at least {MIN_MARKDOWN_BYTES}"),
        ));
    }
    Ok(MarkdownLimits {
        rows: bounded(
            options.max_rows,
            DEFAULT_MARKDOWN_MAX_ROWS,
            MAX_MARKDOWN_ROWS,
            "maxRows",
        )?,
        columns: bounded(
            options.max_columns,
            DEFAULT_MARKDOWN_MAX_COLUMNS,
            MAX_COLS,
            "maxColumns",
        )?,
        cells: bounded(
            options.max_cells,
            DEFAULT_MARKDOWN_MAX_CELLS,
            MAX_MARKDOWN_CELLS,
            "maxCells",
        )?,
        bytes: bounded(Some(bytes), bytes, MAX_EXPORT_BYTES, "maxBytes")?,
    })
}

/// Renders `content` as Markdown. Content that does not validate, and options outside
/// their bounds, are [`Error::InvalidRequest`].
pub fn render_xlsx_markdown(
    content: &XlsxStructuredContent,
    options: &XlsxMarkdownOptions,
) -> Result<XlsxMarkdownContent> {
    let limits = limits(options).map_err(|failure| Error::InvalidRequest(failure.message))?;
    validate(content).map_err(Error::InvalidRequest)?;
    let mut writer = Writer {
        out: String::new(),
        max: limits.bytes - RESERVE,
        anchors: Vec::new(),
        diagnostics: content.diagnostics.clone(),
        stopped: false,
        truncated: content.truncated,
    };
    if content.sheets.iter().any(|sheet| {
        sheet.cells.iter().any(|cell| {
            cell.formula.is_some() || !matches!(cell.value, XlsxExportValue::Text { .. })
        })
    }) {
        writer.lossy(
            None,
            "Markdown shows each cell's display text; formulas, typed values and number formats stay in the structured export.",
        );
    }
    if content.sheets.iter().any(|sheet| {
        sheet
            .cells
            .iter()
            .any(|cell| cell.display_text.chars().any(breaks_line))
    }) {
        writer.lossy(
            None,
            "Line breaks and control characters inside cell text are shown as spaces.",
        );
    }
    let mut cells = limits.cells;
    for sheet in &content.sheets {
        if writer.stopped {
            break;
        }
        render_sheet(&mut writer, content, sheet, &limits, &mut cells)
            .map_err(Error::InvalidRequest)?;
    }
    if !content.defined_names.is_empty() && writer.push("## Defined names\n\n") {
        for defined in &content.defined_names {
            let scope = defined
                .local_sheet
                .as_ref()
                .map(|sheet| format!(" ({})", inline(&sheet.name)))
                .unwrap_or_default();
            let hidden = if defined.hidden { " (hidden)" } else { "" };
            let line = format!(
                "- {} {}{scope}{hidden}: {}\n",
                writer.marker(),
                inline(&defined.name),
                inline(&defined.formula)
            );
            if !writer.item(defined.anchor.clone(), &line) {
                break;
            }
        }
    }
    Ok(writer.finish())
}

struct Writer {
    out: String,
    max: usize,
    anchors: Vec<XlsxMarkdownAnchor>,
    diagnostics: Vec<XlsxExportDiagnostic>,
    stopped: bool,
    truncated: bool,
}

impl Writer {
    /// Appends `text` whole, or stops the rendering if it does not fit.
    fn push(&mut self, text: &str) -> bool {
        if self.stopped || self.out.len() + text.len() > self.max {
            self.stopped = true;
            return false;
        }
        self.out.push_str(text);
        true
    }

    /// The marker the next recorded anchor takes.
    fn marker(&self) -> String {
        format!("<!-- xlsx-export:{} -->", self.anchors.len())
    }

    /// `body` on the lines after a marker line.
    fn marked(&mut self, anchor: XlsxAnchor, body: &str) -> bool {
        let text = format!("{}\n{body}", self.marker());
        self.item(anchor, &text)
    }

    /// A line that already carries [`Writer::marker`].
    fn item(&mut self, anchor: XlsxAnchor, text: &str) -> bool {
        if !self.push(text) {
            return false;
        }
        let marker = self.marker();
        self.anchors.push(XlsxMarkdownAnchor { marker, anchor });
        true
    }

    fn lossy(&mut self, anchor: Option<XlsxAnchor>, message: &str) {
        self.diagnostics.push(XlsxExportDiagnostic {
            code: XlsxExportDiagnosticCode::MarkdownLossy,
            severity: XlsxExportSeverity::Info,
            anchor,
            message: message.to_owned(),
        });
    }

    fn finish(mut self) -> XlsxMarkdownContent {
        if self.stopped {
            self.truncated = true;
            self.out
                .push_str("\n_Markdown stopped at its maxBytes limit._\n");
            self.diagnostics.push(XlsxExportDiagnostic {
                code: XlsxExportDiagnosticCode::Truncated,
                severity: XlsxExportSeverity::Warning,
                anchor: None,
                message: "Markdown stopped at its maxBytes limit; later content is not rendered."
                    .to_owned(),
            });
        }
        XlsxMarkdownContent {
            markdown: self.out,
            anchors: self.anchors,
            diagnostics: self.diagnostics,
            truncated: self.truncated,
        }
    }
}

fn render_sheet(
    writer: &mut Writer,
    content: &XlsxStructuredContent,
    sheet: &XlsxExportSheet,
    limits: &MarkdownLimits,
    budget: &mut usize,
) -> std::result::Result<(), String> {
    let XlsxAnchor::Sheet { sheet: identity } = &sheet.anchor else {
        return Ok(());
    };
    let kind = match sheet.kind {
        XlsxSheetKind::Worksheet => "",
        XlsxSheetKind::Chartsheet => " (chartsheet)",
        XlsxSheetKind::Dialogsheet => " (dialogsheet)",
        XlsxSheetKind::Macrosheet => " (macrosheet)",
        XlsxSheetKind::Other => " (unsupported sheet)",
    };
    let hidden = match sheet.visibility {
        XlsxSheetVisibility::Visible => "",
        XlsxSheetVisibility::Hidden => " (hidden)",
        XlsxSheetVisibility::VeryHidden => " (very hidden)",
        XlsxSheetVisibility::Unknown => " (visibility unknown)",
    };
    if !writer.marked(
        sheet.anchor.clone(),
        &format!("## {}{kind}{hidden}\n\n", inline(&identity.name)),
    ) {
        return Ok(());
    }
    let range = sheet
        .selected_range
        .as_deref()
        .and_then(|a1| CellRange::parse_a1(a1).ok());
    match range {
        Some(range) if sheet.kind == XlsxSheetKind::Worksheet => {
            render_grid(writer, content, sheet, identity, range, limits, budget)?;
        }
        None if sheet.kind == XlsxSheetKind::Worksheet => {
            writer.push("_No cells._\n\n");
        }
        _ => {}
    }
    for table in &sheet.tables {
        let a1 = range_a1(&table.anchor);
        let columns = table
            .columns
            .iter()
            .map(|column| inline(column))
            .collect::<Vec<_>>()
            .join(", ");
        let clipped = if table.clipped { ", clipped" } else { "" };
        let line = format!(
            "- {} Table {} at {a1}{clipped}: {} header rows, {} totals rows; columns {columns}\n",
            writer.marker(),
            inline(&table.name),
            table.header_rows,
            table.totals_rows,
        );
        if !writer.item(table.anchor.clone(), &line) {
            return Ok(());
        }
    }
    for link in &sheet.hyperlinks {
        let external = link.external_target.as_deref().map(|target| {
            let label = inline(link.display.as_deref().unwrap_or(target));
            match link_destination(target) {
                Some(destination) => format!("[{label}]({destination})"),
                None if label == inline(target) => format!("{label} (not linked)"),
                None => format!("{label}: {} (not linked)", inline(target)),
            }
        });
        let location = link
            .location
            .as_deref()
            .map(|location| format!("in-workbook location {}", inline(location)));
        let target = [external, location]
            .into_iter()
            .flatten()
            .collect::<Vec<_>>()
            .join("; ");
        let line = format!(
            "- {} Link at {}: {target}\n",
            writer.marker(),
            range_a1(&link.anchor)
        );
        if !writer.item(link.anchor.clone(), &line) {
            return Ok(());
        }
    }
    for object in &sheet.objects {
        let kind = match object.kind {
            XlsxObjectKind::Chart => "Chart",
            XlsxObjectKind::Picture => "Picture",
            XlsxObjectKind::Shape => "Shape",
            XlsxObjectKind::Group => "Group",
            XlsxObjectKind::Connector => "Connector",
            XlsxObjectKind::Diagram => "Diagram",
            XlsxObjectKind::GraphicFrame => "Graphic frame",
            XlsxObjectKind::ContentPart => "Ink",
            XlsxObjectKind::Unknown => "Drawing object",
        };
        let name = object
            .name
            .as_deref()
            .map(|name| format!(" {}", inline(name)))
            .unwrap_or_default();
        let position = object
            .anchor
            .as_ref()
            .map(|anchor| format!(" at {}", range_a1(anchor)))
            .unwrap_or_else(|| " (position unknown)".to_owned());
        let alt = object
            .alt_text
            .as_deref()
            .map(|alt| format!("; alt text: {}", inline(alt)))
            .unwrap_or_default();
        let hidden = if object.hidden { " (hidden)" } else { "" };
        let line = format!(
            "- {} {kind}{name}{position}{hidden}{alt}\n",
            writer.marker()
        );
        let anchor = match (&object.anchor, &object.source) {
            (Some(anchor), _) => anchor.clone(),
            (None, Some(source)) => XlsxAnchor::SourcePart {
                part: source.part.clone(),
                part_sha256: source.part_sha256.clone(),
                path: source.path.clone(),
            },
            (None, None) => sheet.anchor.clone(),
        };
        if !writer.item(anchor, &line) {
            return Ok(());
        }
    }
    if !(sheet.tables.is_empty() && sheet.hyperlinks.is_empty() && sheet.objects.is_empty()) {
        writer.push("\n");
    }
    Ok(())
}

#[allow(clippy::too_many_arguments)]
fn render_grid(
    writer: &mut Writer,
    content: &XlsxStructuredContent,
    sheet: &XlsxExportSheet,
    identity: &XlsxSheetIdentity,
    range: CellRange,
    limits: &MarkdownLimits,
    budget: &mut usize,
) -> std::result::Result<(), String> {
    let hidden_rows = spans(&sheet.hidden_rows, parse_row_span);
    let hidden_columns = spans(&sheet.hidden_columns, parse_column_span);
    let covered_until = if sheet.truncated {
        match sheet.cells.last().and_then(|cell| cell_ref(&cell.anchor)) {
            Some(last) => Some(last),
            None => {
                writer.push("_Export stopped before this sheet's cells._\n\n");
                return Ok(());
            }
        }
    } else {
        None
    };
    let (columns, columns_clipped) = axis(
        range.start.col,
        range.end.col,
        (!content.included.hidden_columns).then_some(hidden_columns.as_slice()),
        limits.columns,
    );
    let (Some(&first_col), Some(&last_col)) = (columns.first(), columns.last()) else {
        writer.push("_Every column of the selected range is hidden and excluded._\n\n");
        return Ok(());
    };
    let row_limit = limits.rows.min(*budget / columns.len());
    if row_limit == 0 {
        writer.truncated = true;
        writer.lossy(
            Some(sheet.anchor.clone()),
            "The Markdown cell limit was reached before this sheet; its grid is omitted.",
        );
        writer.push("_Grid omitted: the Markdown cell limit was reached._\n\n");
        return Ok(());
    }
    let (mut rows, mut rows_clipped) = axis(
        range.start.row,
        range.end.row,
        (!content.included.hidden_rows).then_some(hidden_rows.as_slice()),
        row_limit,
    );
    if let Some(last) = covered_until {
        let full = last.col >= last_col;
        let before = rows.len();
        rows.retain(|&row| row < last.row || (row == last.row && full));
        rows_clipped |= rows.len() < before;
    }
    let (Some(&first_row), Some(&last_row)) = (rows.first(), rows.last()) else {
        writer.push("_No rows of this sheet are covered._\n\n");
        return Ok(());
    };
    *budget -= rows.len() * columns.len();
    if rows_clipped || columns_clipped {
        writer.truncated = true;
        writer.lossy(
            Some(sheet.anchor.clone()),
            "Markdown shows only part of this sheet's selected range; the rest stays in the structured export.",
        );
    }
    let window = CellRange::new(
        CellRef::new(first_row, first_col),
        CellRef::new(last_row, last_col),
    );
    let mut values = BTreeMap::new();
    for cell in &sheet.cells {
        let Some(at) = cell_ref(&cell.anchor) else {
            continue;
        };
        if at.row > last_row {
            break;
        }
        if rows.binary_search(&at.row).is_ok() && columns.binary_search(&at.col).is_ok() {
            values.insert((at.row, at.col), cell);
        }
    }
    let spans = sheet
        .merges
        .iter()
        .filter_map(|merge| span(canonical_range(&range_a1(&merge.anchor))?, &rows, &columns))
        .collect::<Vec<_>>();
    let anchor = XlsxAnchor::Range {
        sheet: identity.clone(),
        a1: window_a1(window),
    };
    if spans.is_empty() {
        pipe_table(writer, anchor, &rows, &columns, &values);
        return Ok(());
    }
    let mut covering = BTreeMap::new();
    for (index, span) in spans.iter().enumerate() {
        for row in span.rows.clone() {
            for col in span.columns.clone() {
                if covering.insert((row, col), index).is_some() {
                    return Err(format!("merges on sheet {} overlap", identity.index));
                }
            }
        }
    }
    html_table(
        writer, sheet, anchor, &rows, &columns, &values, &spans, &covering,
    );
    Ok(())
}

/// A merge as the grid shows it: the range, the rendered row and column index spans it
/// covers, and whether the grid cuts part of it off.
struct MergeSpan {
    range: CellRange,
    rows: std::ops::Range<usize>,
    columns: std::ops::Range<usize>,
}

impl MergeSpan {
    fn clipped(&self, rows: &[u32], columns: &[u32]) -> bool {
        let top_left = (rows[self.rows.start], columns[self.columns.start]);
        top_left != (self.range.start.row, self.range.start.col)
            || self.rows.len() != (self.range.end.row - self.range.start.row + 1) as usize
            || self.columns.len() != (self.range.end.col - self.range.start.col + 1) as usize
    }
}

fn span(range: CellRange, rows: &[u32], columns: &[u32]) -> Option<MergeSpan> {
    let within = |axis: &[u32], start: u32, end: u32| {
        axis.partition_point(|&index| index < start)..axis.partition_point(|&index| index <= end)
    };
    let rows = within(rows, range.start.row, range.end.row);
    let columns = within(columns, range.start.col, range.end.col);
    (!rows.is_empty() && !columns.is_empty()).then_some(MergeSpan {
        range,
        rows,
        columns,
    })
}

fn pipe_table(
    writer: &mut Writer,
    anchor: XlsxAnchor,
    rows: &[u32],
    columns: &[u32],
    values: &BTreeMap<(u32, u32), &XlsxExportCell>,
) {
    let mut header = String::from("|  |");
    let mut rule = String::from("| ---: |");
    for &col in columns {
        header.push_str(&format!(" {} |", col_to_letters(col)));
        rule.push_str(" --- |");
    }
    if !writer.marked(anchor, &format!("{header}\n{rule}\n")) {
        return;
    }
    for &row in rows {
        let mut line = format!("| {} |", row + 1);
        for &col in columns {
            let text = values
                .get(&(row, col))
                .map(|cell| inline(&cell.display_text))
                .unwrap_or_default();
            line.push_str(&format!(" {text} |"));
        }
        line.push('\n');
        if !writer.push(&line) {
            break;
        }
    }
    writer.out.push('\n');
}

#[allow(clippy::too_many_arguments)]
fn html_table(
    writer: &mut Writer,
    sheet: &XlsxExportSheet,
    anchor: XlsxAnchor,
    rows: &[u32],
    columns: &[u32],
    values: &BTreeMap<(u32, u32), &XlsxExportCell>,
    spans: &[MergeSpan],
    covering: &BTreeMap<(usize, usize), usize>,
) {
    let mut header = String::from("<table>\n<thead><tr><th></th>");
    for &col in columns {
        header.push_str(&format!("<th>{}</th>", col_to_letters(col)));
    }
    header.push_str("</tr></thead>\n<tbody>\n");
    if !writer.marked(anchor, &header) {
        return;
    }
    let (mut suppressed, mut clipped) = (false, false);
    for (row_index, &row) in rows.iter().enumerate() {
        let mut line = format!("<tr><th>{}</th>", row + 1);
        for (col_index, &col) in columns.iter().enumerate() {
            let value = values.get(&(row, col));
            let Some(&merge) = covering.get(&(row_index, col_index)) else {
                let text = value
                    .map(|cell| html(&cell.display_text))
                    .unwrap_or_default();
                line.push_str(&format!("<td>{text}</td>"));
                continue;
            };
            let span = &spans[merge];
            if (row_index, col_index) != (span.rows.start, span.columns.start) {
                suppressed |= value.is_some_and(|cell| !cell.display_text.is_empty());
                continue;
            }
            let origin = CellRef::new(span.range.start.row, span.range.start.col);
            let text = if (row, col) == (origin.row, origin.col) {
                value
                    .map(|cell| html(&cell.display_text))
                    .unwrap_or_default()
            } else {
                suppressed |= value.is_some_and(|cell| !cell.display_text.is_empty());
                format!("[merged from {}]", origin.to_a1())
            };
            clipped |= span.clipped(rows, columns);
            line.push_str(&format!(
                "<td rowspan=\"{}\" colspan=\"{}\" data-merge=\"{}\">{text}</td>",
                span.rows.len(),
                span.columns.len(),
                html(&window_a1(span.range))
            ));
        }
        line.push_str("</tr>\n");
        if !writer.push(&line) {
            break;
        }
    }
    writer.out.push_str("</tbody>\n</table>\n\n");
    if suppressed {
        writer.lossy(
            Some(sheet.anchor.clone()),
            "Values stored in cells a merge covers are not shown in Markdown; the structured export keeps them.",
        );
    }
    if clipped {
        writer.lossy(
            Some(sheet.anchor.clone()),
            "Some merged ranges extend past the rendered grid; each shows only its rendered part, and data-merge names the full range.",
        );
    }
}

/// Up to `limit` indices from `start` through `end`, skipping `hidden` spans, and whether
/// any were left out by the limit.
fn axis(start: u32, end: u32, hidden: Option<&[(u32, u32)]>, limit: usize) -> (Vec<u32>, bool) {
    let mut out = Vec::new();
    let mut next = start;
    let mut spans = hidden.unwrap_or_default().iter().peekable();
    while next <= end {
        while spans.next_if(|span| span.1 < next).is_some() {}
        if let Some(&&(first, last)) = spans.peek()
            && first <= next
        {
            match last.checked_add(1) {
                Some(after) => next = after,
                None => break,
            }
            continue;
        }
        if out.len() == limit {
            return (out, true);
        }
        out.push(next);
        match next.checked_add(1) {
            Some(after) => next = after,
            None => break,
        }
    }
    (out, false)
}

fn spans(text: &[String], parse: fn(&str) -> Option<(u32, u32)>) -> Vec<(u32, u32)> {
    let mut spans = text
        .iter()
        .filter_map(|span| parse(span))
        .collect::<Vec<_>>();
    spans.sort_unstable();
    spans
}

fn parse_row_span(text: &str) -> Option<(u32, u32)> {
    let (start, end) = text.split_once(':')?;
    let parse = |value: &str| {
        let row = value.parse::<u32>().ok()?;
        (value == row.to_string() && (1..=MAX_ROWS).contains(&row)).then(|| row - 1)
    };
    let (start, end) = (parse(start)?, parse(end)?);
    (start <= end).then_some((start, end))
}

fn parse_column_span(text: &str) -> Option<(u32, u32)> {
    let (start, end) = text.split_once(':')?;
    let parse = |value: &str| {
        let col = CellRef::parse_a1(&format!("{value}1")).ok()?.col;
        (col_to_letters(col) == value).then_some(col)
    };
    let (start, end) = (parse(start)?, parse(end)?);
    (start <= end).then_some((start, end))
}

fn cell_ref(anchor: &XlsxAnchor) -> Option<CellRef> {
    match anchor {
        XlsxAnchor::Cell { a1, .. } => CellRef::parse_a1(a1).ok(),
        _ => None,
    }
}

fn range_a1(anchor: &XlsxAnchor) -> String {
    match anchor {
        XlsxAnchor::Cell { a1, .. } | XlsxAnchor::Range { a1, .. } => a1.clone(),
        _ => String::new(),
    }
}

fn window_a1(range: CellRange) -> String {
    CellRange::new(
        CellRef::new(range.start.row, range.start.col),
        CellRef::new(range.end.row, range.end.col),
    )
    .to_a1()
}

/// Whether `c` could end a line for some renderer: controls, NEL and the Unicode line and
/// paragraph separators.
fn breaks_line(c: char) -> bool {
    c.is_control() || matches!(c, '\u{2028}' | '\u{2029}')
}

/// Escapes text for single-line Markdown inline content, pipe-table cells included:
/// entities for what HTML would read, backslashes for Markdown punctuation, and a space
/// for anything that could end the line.
fn inline(text: &str) -> String {
    let mut out = String::with_capacity(text.len());
    for c in text.chars() {
        match c {
            '&' => out.push_str("&amp;"),
            '<' => out.push_str("&lt;"),
            '>' => out.push_str("&gt;"),
            '\\' | '`' | '*' | '_' | '[' | ']' | '|' | '~' | '!' | '$' | '#' | '{' | '}' | '('
            | ')' => {
                out.push('\\');
                out.push(c);
            }
            c if breaks_line(c) => out.push(' '),
            c => out.push(c),
        }
    }
    out
}

/// Escapes text for an HTML element or attribute value on one line.
fn html(text: &str) -> String {
    let mut out = String::with_capacity(text.len());
    for c in text.chars() {
        match c {
            '&' => out.push_str("&amp;"),
            '<' => out.push_str("&lt;"),
            '>' => out.push_str("&gt;"),
            '"' => out.push_str("&quot;"),
            '\'' => out.push_str("&#39;"),
            c if breaks_line(c) => out.push(' '),
            c => out.push(c),
        }
    }
    out
}

/// A Markdown link destination for `target`, or `None` unless its effective scheme is
/// `http`, `https` or `mailto` once entities, percent-escapes and whitespace or control
/// characters a browser would drop are taken into account. Characters that could end the
/// destination are percent-encoded and `&` is written as `&amp;`, so entity decoding cannot
/// rebuild a scheme while query strings keep their meaning.
fn link_destination(target: &str) -> Option<String> {
    let effective = decode_entities(target)
        .chars()
        .filter(|c| !c.is_ascii_whitespace() && !c.is_control())
        .collect::<String>();
    let (scheme, _) = effective.split_once(':')?;
    let scheme = percent_decode(scheme)?.to_ascii_lowercase();
    if !matches!(scheme.as_str(), "http" | "https" | "mailto") {
        return None;
    }
    let mut out = String::with_capacity(target.len());
    for c in target.trim().chars() {
        match c {
            '&' => out.push_str("&amp;"),
            ' ' | '(' | ')' | '<' | '>' | '\\' | '"' | '\'' | '`' | '[' | ']' => {
                out.push_str(&format!("%{:02X}", c as u32));
            }
            c if c.is_control() || matches!(c, '\u{2028}' | '\u{2029}') => {
                let mut bytes = [0; 4];
                for byte in c.encode_utf8(&mut bytes).bytes() {
                    out.push_str(&format!("%{byte:02X}"));
                }
            }
            c => out.push(c),
        }
    }
    Some(out)
}

/// `text` with the HTML character references that can spell a scheme decoded, terminated
/// or not: numeric ones and the named ones for ASCII punctuation and whitespace. Other
/// names stay literal, so their `&` keeps any scheme they sit in from matching.
fn decode_entities(text: &str) -> String {
    let mut out = String::with_capacity(text.len());
    let mut rest = text;
    while let Some(at) = rest.find('&') {
        out.push_str(&rest[..at]);
        rest = &rest[at + 1..];
        if let Some(number) = rest.strip_prefix('#') {
            let (digits, radix, skip) = match number.strip_prefix(['x', 'X']) {
                Some(hex) => (hex, 16, 2),
                None => (number, 10, 1),
            };
            let len = digits
                .find(|c: char| !c.is_digit(radix))
                .unwrap_or(digits.len());
            if len == 0 {
                out.push('&');
                continue;
            }
            let code = u32::from_str_radix(&digits[..len], radix).ok();
            out.push(code.and_then(char::from_u32).unwrap_or('\u{FFFD}'));
            rest = &rest[skip + len..];
            rest = rest.strip_prefix(';').unwrap_or(rest);
            continue;
        }
        let len = rest
            .find(|c: char| !c.is_ascii_alphanumeric())
            .unwrap_or(rest.len());
        let decoded = match &rest[..len] {
            "amp" => '&',
            "lt" => '<',
            "gt" => '>',
            "quot" => '"',
            "apos" => '\'',
            "colon" => ':',
            "Tab" => '\t',
            "NewLine" => '\n',
            "sol" => '/',
            "period" => '.',
            "plus" => '+',
            "percnt" => '%',
            "hyphen" | "dash" => '-',
            "nbsp" => '\u{A0}',
            _ => {
                out.push('&');
                continue;
            }
        };
        out.push(decoded);
        rest = &rest[len..];
        rest = rest.strip_prefix(';').unwrap_or(rest);
    }
    out.push_str(rest);
    out
}

/// `text` with `%XX` escapes decoded; `None` when they do not form UTF-8.
fn percent_decode(text: &str) -> Option<String> {
    let bytes = text.as_bytes();
    let mut out = Vec::with_capacity(bytes.len());
    let mut index = 0;
    while index < bytes.len() {
        let hex = bytes
            .get(index + 1..index + 3)
            .and_then(|pair| std::str::from_utf8(pair).ok())
            .and_then(|pair| u8::from_str_radix(pair, 16).ok());
        match (bytes[index], hex) {
            (b'%', Some(byte)) => {
                out.push(byte);
                index += 3;
            }
            (byte, _) => {
                out.push(byte);
                index += 1;
            }
        }
    }
    String::from_utf8(out).ok()
}

/// A1 in the canonical form exports write.
fn canonical_range(a1: &str) -> Option<CellRange> {
    let range = CellRange::parse_a1(a1).ok()?;
    (window_a1(range) == a1).then_some(range)
}

/// A cell anchor on `sheet` in canonical A1.
fn cell_on(anchor: &XlsxAnchor, sheet: &XlsxSheetIdentity) -> Option<CellRef> {
    let XlsxAnchor::Cell { sheet: on, a1 } = anchor else {
        return None;
    };
    let at = CellRef::parse_a1(a1).ok()?;
    let at = CellRef::new(at.row, at.col);
    (on == sheet && at.to_a1() == *a1).then_some(at)
}

/// A range anchor on `sheet` in canonical A1.
fn range_on(anchor: &XlsxAnchor, sheet: &XlsxSheetIdentity) -> Option<CellRange> {
    let XlsxAnchor::Range { sheet: on, a1 } = anchor else {
        return None;
    };
    (on == sheet).then(|| canonical_range(a1)).flatten()
}

/// Checks the contract the renderer relies on, so content from elsewhere is refused
/// rather than misrendered.
fn validate(content: &XlsxStructuredContent) -> std::result::Result<(), String> {
    let mut previous: Option<u32> = None;
    for sheet in &content.sheets {
        let XlsxAnchor::Sheet { sheet: identity } = &sheet.anchor else {
            return Err(format!("sheet {} has no sheet anchor", sheet.id));
        };
        let at = identity.index;
        if previous.is_some_and(|index| index >= at) {
            return Err("sheets must be in increasing index order".to_owned());
        }
        previous = Some(at);
        if identity.name.is_empty() {
            return Err(format!("sheet {at} has no name"));
        }
        if sheet.truncated && !content.truncated {
            return Err(format!("sheet {at} is truncated but the content is not"));
        }
        let selected = match &sheet.selected_range {
            Some(a1) => Some(
                canonical_range(a1)
                    .ok_or_else(|| format!("sheet {at} has an invalid selectedRange"))?,
            ),
            None => None,
        };
        if sheet
            .used_range
            .as_deref()
            .is_some_and(|a1| canonical_range(a1).is_none())
        {
            return Err(format!("sheet {at} has an invalid usedRange"));
        }
        if sheet
            .hidden_rows
            .iter()
            .any(|span| parse_row_span(span).is_none())
            || sheet
                .hidden_columns
                .iter()
                .any(|span| parse_column_span(span).is_none())
        {
            return Err(format!("sheet {at} has an invalid hidden span"));
        }
        let ranges = sheet
            .merges
            .iter()
            .map(|merge| &merge.anchor)
            .chain(sheet.tables.iter().map(|table| &table.anchor))
            .chain(sheet.hyperlinks.iter().map(|link| &link.anchor));
        for anchor in ranges {
            if range_on(anchor, identity).is_none() {
                return Err(format!("sheet {at} has a range record anchored elsewhere"));
            }
        }
        for object in &sheet.objects {
            let placed = object.anchor.as_ref().is_none_or(|anchor| {
                range_on(anchor, identity).is_some() || cell_on(anchor, identity).is_some()
            });
            if !placed {
                return Err(format!("object {} is not anchored on its sheet", object.id));
            }
        }
        let mut last: Option<(u32, u32)> = None;
        for cell in &sheet.cells {
            let Some(position) = cell_on(&cell.anchor, identity) else {
                return Err(format!("cell {} has no cell anchor on its sheet", cell.id));
            };
            if last.is_some_and(|last| last >= (position.row, position.col)) {
                return Err(format!("cells of sheet {at} are not in row-major order"));
            }
            if !selected.is_some_and(|range| range.contains(position)) {
                return Err(format!("cell {} lies outside the selected range", cell.id));
            }
            if cell
                .merge
                .as_ref()
                .is_some_and(|merge| canonical_range(&merge.range).is_none())
            {
                return Err(format!("cell {} has an invalid merge range", cell.id));
            }
            last = Some((position.row, position.col));
        }
    }
    for defined in &content.defined_names {
        if !matches!(defined.anchor, XlsxAnchor::DefinedName { .. }) {
            return Err(format!(
                "defined name {} has no defined-name anchor",
                defined.id
            ));
        }
    }
    Ok(())
}
