//! Sheet qualifiers in formula text: `Sheet1!`, `'My Sheet'!`, `Jan:Mar!` and
//! `'Jan:Mar'!`, found outside strings and external-book brackets, and the
//! token Excel writes for a sheet or a span of sheets. Shared by the sheet
//! rename in `xlsx-ops` and the writer's rename of preserved fragments.

use crate::CellRef;

/// One sheet qualifier found in a formula. A 3-D qualifier names two
/// endpoints and covers every sheet between them in workbook order, whether it
/// is written `Jan:Mar!` or, as Excel writes it, `'Jan:Mar'!`.
pub struct Qualifier {
    pub span: core::ops::Range<usize>,
    pub first: String,
    pub last: String,
    /// preceded by `]`, so it names a sheet in another workbook.
    pub external: bool,
    /// followed by `!`, so it really qualifies a reference.
    pub bound: bool,
    /// followed by `(`, so it is a function call rather than a sheet name.
    pub call: bool,
}

impl Qualifier {
    pub fn is_span(&self) -> bool {
        self.first != self.last
    }
}

/// Every sheet qualifier in `source`, in order, skipping strings and the
/// `[book]` prefixes of external references.
pub fn sheet_qualifiers(source: &str) -> Vec<Qualifier> {
    let bytes = source.as_bytes();
    let mut out = Vec::new();
    let mut index = 0;
    let mut bracket_depth = 0_u32;
    while index < bytes.len() {
        if bytes[index] == b'"' {
            index = skip_string(source, index);
            continue;
        }
        if bytes[index] == b'[' {
            bracket_depth = bracket_depth.saturating_add(1);
            index += 1;
            continue;
        }
        if bytes[index] == b']' {
            bracket_depth = bracket_depth.saturating_sub(1);
            index += 1;
            continue;
        }
        if bracket_depth != 0 {
            index += next_char_len(source, index);
            continue;
        }
        let Some(first) = parse_sheet_token(source, index) else {
            index += next_char_len(source, index);
            continue;
        };
        let external = first.start > 0 && bytes[first.start - 1] == b']';
        if bytes.get(first.end) == Some(&b'!') {
            let (start, end) = quoted_endpoints(&first.name);
            out.push(Qualifier {
                span: first.start..first.end,
                first: start,
                last: end,
                external,
                bound: true,
                call: false,
            });
            index = first.end + 1;
            continue;
        }
        if bytes.get(first.end) == Some(&b':')
            && let Some(second) = parse_sheet_token(source, first.end + 1)
            && bytes.get(second.end) == Some(&b'!')
        {
            out.push(Qualifier {
                span: first.start..second.end,
                first: first.name,
                last: second.name,
                external,
                bound: true,
                call: false,
            });
            index = second.end + 1;
            continue;
        }
        let call = is_function_call(source, &first);
        index = first.end;
        out.push(Qualifier {
            span: first.start..first.end,
            first: first.name.clone(),
            last: first.name,
            external,
            bound: false,
            call,
        });
    }
    out
}

/// `'Jan:Mar'` names the span from `Jan` to `Mar`; a sheet name cannot itself
/// contain a colon, so the split is unambiguous.
pub fn quoted_endpoints(name: &str) -> (String, String) {
    match name.split_once(':') {
        Some((first, last)) => (first.to_owned(), last.to_owned()),
        None => (name.to_owned(), name.to_owned()),
    }
}

/// An unquoted token followed by `(` is a function call, never a sheet name.
pub fn is_function_call(source: &str, token: &ParsedSheetToken) -> bool {
    !source[token.start..].starts_with('\'') && source[token.end..].trim_start().starts_with('(')
}

pub fn next_char_len(source: &str, index: usize) -> usize {
    source[index..]
        .chars()
        .next()
        .map(char::len_utf8)
        .unwrap_or(1)
}

pub struct ParsedSheetToken {
    pub start: usize,
    pub end: usize,
    pub name: String,
}

pub fn parse_sheet_token(source: &str, start: usize) -> Option<ParsedSheetToken> {
    let bytes = source.as_bytes();
    if bytes.get(start) == Some(&b'\'') {
        let mut index = start + 1;
        let mut name = String::new();
        while index < bytes.len() {
            if bytes[index] == b'\'' {
                if bytes.get(index + 1) == Some(&b'\'') {
                    name.push('\'');
                    index += 2;
                } else {
                    return Some(ParsedSheetToken {
                        start,
                        end: index + 1,
                        name,
                    });
                }
            } else {
                let character = source[index..].chars().next()?;
                name.push(character);
                index += character.len_utf8();
            }
        }
        return None;
    }
    let first = source[start..].chars().next()?;
    if !is_unquoted_sheet_char(first) {
        return None;
    }
    let mut end = start;
    while end < bytes.len() {
        let character = source[end..].chars().next()?;
        if !is_unquoted_sheet_char(character) {
            break;
        }
        end += character.len_utf8();
    }
    Some(ParsedSheetToken {
        start,
        end,
        name: source[start..end].to_string(),
    })
}

pub fn skip_string(source: &str, start: usize) -> usize {
    let bytes = source.as_bytes();
    let mut index = start + 1;
    while index < bytes.len() {
        if bytes[index] == b'"' {
            index += 1;
            if bytes.get(index) == Some(&b'"') {
                index += 1;
            } else {
                break;
            }
        } else {
            index += source[index..]
                .chars()
                .next()
                .map(char::len_utf8)
                .unwrap_or(1);
        }
    }
    index
}

pub fn sheet_names_equal(left: &str, right: &str) -> bool {
    left.to_lowercase() == right.to_lowercase()
}

pub fn is_unquoted_sheet_char(character: char) -> bool {
    !character.is_whitespace()
        && !matches!(
            character,
            '"' | '\''
                | '!'
                | ':'
                | '+'
                | '-'
                | '*'
                | '/'
                | '^'
                | '&'
                | '='
                | '<'
                | '>'
                | '('
                | ')'
                | ','
                | '%'
                | '['
                | ']'
        )
}

pub fn sheet_token(name: &str) -> String {
    if is_simple_sheet_name(name) {
        name.to_string()
    } else {
        format!("'{}'", name.replace('\'', "''"))
    }
}

/// A qualifier naming one sheet, or the span between two. Excel writes a span
/// unquoted when both endpoints are simple and `'first:last'` otherwise.
pub fn qualifier_token(first: &str, last: &str) -> String {
    if sheet_names_equal(first, last) {
        return sheet_token(first);
    }
    if is_simple_sheet_name(first) && is_simple_sheet_name(last) {
        return format!("{first}:{last}");
    }
    format!(
        "'{}:{}'",
        first.replace('\'', "''"),
        last.replace('\'', "''")
    )
}

pub fn is_simple_sheet_name(name: &str) -> bool {
    !name.is_empty()
        && name
            .chars()
            .all(|character| character.is_ascii_alphanumeric() || character == '_')
        && !name
            .chars()
            .next()
            .is_some_and(|character| character.is_ascii_digit())
        && CellRef::parse_a1(name).is_err()
        && !is_r1c1_reference(name)
}

pub fn is_r1c1_reference(name: &str) -> bool {
    let upper = name.to_ascii_uppercase();
    if matches!(upper.as_str(), "R" | "C") {
        return true;
    }
    let Some(rest) = upper.strip_prefix('R') else {
        return false;
    };
    let digits = rest.bytes().take_while(u8::is_ascii_digit).count();
    rest.as_bytes().get(digits) == Some(&b'C')
}

/// `source` with the sheets `renamed` maps (each sheet's old name to its new
/// one) renamed in every qualifier that names them, all in one pass, so a
/// swap or a chain lands on each sheet's own new name rather than replaying
/// the renames in turn. External references, strings and tokens that only
/// look like sheet names are left alone. Also returns which entries of
/// `renamed` were written, so a caller can check what it wrote.
pub fn rename_sheets(source: &str, renamed: &[(String, String)]) -> (String, Vec<usize>) {
    let lookup = |name: &str| {
        renamed
            .iter()
            .position(|(old, _)| sheet_names_equal(old, name))
    };
    let mut used = Vec::new();
    let mut replacements = Vec::new();
    for qualifier in sheet_qualifiers(source) {
        if qualifier.external || !qualifier.bound {
            continue;
        }
        let first = lookup(&qualifier.first);
        let last = lookup(&qualifier.last);
        if first.is_none() && last.is_none() {
            continue;
        }
        used.extend(first.into_iter().chain(last));
        let first = first.map_or(qualifier.first.as_str(), |at| renamed[at].1.as_str());
        let last = last.map_or(qualifier.last.as_str(), |at| renamed[at].1.as_str());
        replacements.push((qualifier.span, qualifier_token(first, last)));
    }
    if replacements.is_empty() {
        return (source.to_owned(), used);
    }
    let mut output = String::with_capacity(source.len());
    let mut copied_until = 0;
    for (span, replacement) in replacements {
        output.push_str(&source[copied_until..span.start]);
        output.push_str(&replacement);
        copied_until = span.end;
    }
    output.push_str(&source[copied_until..]);
    used.sort_unstable();
    used.dedup();
    (output, used)
}

#[cfg(test)]
mod tests {
    use super::rename_sheets;

    fn renamed(pairs: &[(&str, &str)]) -> Vec<(String, String)> {
        pairs
            .iter()
            .map(|(old, new)| ((*old).to_owned(), (*new).to_owned()))
            .collect()
    }

    #[test]
    fn a_swap_renames_each_qualifier_once() {
        let (text, used) = rename_sheets("A!$A$1>B!$A$1", &renamed(&[("A", "B"), ("B", "A")]));
        assert_eq!(text, "B!$A$1>A!$A$1");
        assert_eq!(used, [0, 1]);
    }

    #[test]
    fn strings_external_references_and_bare_names_are_left_alone() {
        let pairs = renamed(&[("Data", "New Data")]);
        for source in ["\"Data!A1\"&Data", "[1]Data!A1", "Data", "SUM(Data)"] {
            assert_eq!(
                rename_sheets(source, &pairs),
                (source.to_owned(), vec![]),
                "{source}"
            );
        }
        assert_eq!(rename_sheets("'data'!A1", &pairs).0, "'New Data'!A1");
        assert_eq!(
            rename_sheets("Data:Other!A1", &pairs).0,
            "'New Data:Other'!A1"
        );
    }
}

/// `Some(Some(text))` narrows the span onto the sheets that remain,
/// `Some(None)` means nothing it covered survives, `None` means the removal
/// does not touch it.
pub fn narrowed_span(
    qualifier: &Qualifier,
    removed_name: &str,
    order_before: &[String],
) -> Option<Option<String>> {
    let position = |name: &str| {
        order_before
            .iter()
            .position(|sheet| sheet_names_equal(sheet, name))
    };
    let (Some(start), Some(end)) = (position(&qualifier.first), position(&qualifier.last)) else {
        return (sheet_names_equal(&qualifier.first, removed_name)
            || sheet_names_equal(&qualifier.last, removed_name))
        .then_some(None);
    };
    let covered = &order_before[start.min(end)..=start.max(end)];
    if !covered
        .iter()
        .any(|sheet| sheet_names_equal(sheet, removed_name))
    {
        return None;
    }
    let remaining = covered
        .iter()
        .filter(|sheet| !sheet_names_equal(sheet, removed_name))
        .collect::<Vec<_>>();
    let (Some(first), Some(last)) = (remaining.first(), remaining.last()) else {
        return Some(None);
    };
    Some(Some(qualifier_token(first, last)))
}

/// The end of the address that follows a sheet qualifier's `!` at `start`: a
/// cell, a range, whole rows or columns, a sheet-scoped name, or an error
/// already standing in for one. A `:` joins only an address shaped like a
/// cell, columns or rows; a name, a function call or another qualifier after
/// it is a separate reference.
fn reference_end(source: &str, start: usize) -> usize {
    if source[start..].starts_with("#REF!") {
        return start + "#REF!".len();
    }
    let run = |from: usize| {
        from + source[from..]
            .chars()
            .take_while(|ch| ch.is_alphanumeric() || matches!(ch, '$' | '_' | '.' | '\\'))
            .map(char::len_utf8)
            .sum::<usize>()
    };
    let end = run(start);
    if source[end..].starts_with(':') {
        let further = run(end + 1);
        let next = source[further..].trim_start();
        if further > end + 1
            && address_shape(&source[start..end]).is_some()
            && address_shape(&source[start..end]) == address_shape(&source[end + 1..further])
            && !next.starts_with([':', '!', '('])
        {
            return further;
        }
    }
    end
}

/// What a range endpoint is: a cell, a column (up to XFD) or a row. A `:`
/// joins only two endpoints of the same shape, as Excel reads a range, so
/// `A1:Tax` is a cell and a name, while `A:Tax` is a range of columns.
#[derive(Clone, Copy, PartialEq, Eq)]
enum AddressShape {
    Cell,
    Column,
    Row,
}

fn address_shape(text: &str) -> Option<AddressShape> {
    let bare = text.replace('$', "");
    if bare.is_empty() {
        return None;
    }
    if bare.bytes().all(|byte| byte.is_ascii_digit()) {
        return Some(AddressShape::Row);
    }
    if bare.len() <= 3 && bare.bytes().all(|byte| byte.is_ascii_alphabetic()) {
        let column = bare.bytes().fold(0_u32, |acc, byte| {
            acc * 26 + u32::from(byte.to_ascii_uppercase() - b'A' + 1)
        });
        return (column <= 16_384).then_some(AddressShape::Column);
    }
    CellRef::parse_a1(&bare.to_ascii_uppercase())
        .is_ok()
        .then_some(AddressShape::Cell)
}

/// `source` with every reference to the sheet `removed_name` rewritten as
/// Excel writes it when that sheet is deleted: a reference that named only
/// that sheet becomes `#REF!`, its address included; a 3-D span whose
/// endpoint is that sheet narrows onto the sheets that remain between its
/// endpoints in `order_before`, or becomes `#REF!` when none remain, and a
/// span the sheet only lies inside keeps its text. `None` when nothing
/// changes. External references and strings are left alone.
pub fn remove_sheet_references(
    source: &str,
    removed_name: &str,
    order_before: &[String],
) -> Option<String> {
    remove_sheet_references_noting_spans(source, removed_name, order_before).map(|(text, _)| text)
}

/// [`remove_sheet_references`], also saying whether a 3-D span narrowed,
/// which depends on the sheet order the removal is read against.
pub fn remove_sheet_references_noting_spans(
    source: &str,
    removed_name: &str,
    order_before: &[String],
) -> Option<(String, bool)> {
    let mut replacements: Vec<(core::ops::Range<usize>, String)> = Vec::new();
    let mut narrowed = false;
    for qualifier in sheet_qualifiers(source) {
        if !qualifier.bound || qualifier.external {
            continue;
        }
        let whole = || qualifier.span.start..reference_end(source, qualifier.span.end + 1);
        if !qualifier.is_span() {
            if sheet_names_equal(&qualifier.first, removed_name) {
                replacements.push((whole(), "#REF!".to_owned()));
            }
            continue;
        }
        if !sheet_names_equal(&qualifier.first, removed_name)
            && !sheet_names_equal(&qualifier.last, removed_name)
        {
            continue;
        }
        match narrowed_span(&qualifier, removed_name, order_before) {
            Some(Some(token)) => {
                narrowed = true;
                replacements.push((qualifier.span.clone(), token));
            }
            Some(None) => replacements.push((whole(), "#REF!".to_owned())),
            None => {}
        }
    }
    if replacements.is_empty() {
        return None;
    }
    let mut merged: Vec<(core::ops::Range<usize>, String)> = Vec::with_capacity(replacements.len());
    for (span, replacement) in replacements {
        match merged.last_mut() {
            Some((last, text)) if span.start < last.end => {
                last.end = last.end.max(span.end);
                *text = "#REF!".to_owned();
            }
            _ => merged.push((span, replacement)),
        }
    }
    let mut output = String::with_capacity(source.len());
    let mut copied = 0;
    for (span, replacement) in merged {
        output.push_str(&source[copied..span.start]);
        output.push_str(&replacement);
        copied = span.end;
    }
    output.push_str(&source[copied..]);
    Some((output, narrowed))
}

#[cfg(test)]
mod removal_tests {
    use super::remove_sheet_references;

    fn sheets(names: &[&str]) -> Vec<String> {
        names.iter().map(|name| (*name).to_owned()).collect()
    }

    #[test]
    fn a_reference_to_the_deleted_sheet_becomes_ref_with_its_address() {
        let order = sheets(&["Data", "Other"]);
        for (source, removed) in [
            ("Data!A1+1", "#REF!+1"),
            ("SUM('Data'!$A$1:$B$5)*2", "SUM(#REF!)*2"),
            ("Data!A:C", "#REF!"),
            ("data!1:3", "#REF!"),
            ("Data!Total", "#REF!"),
            ("Data!#REF!", "#REF!"),
        ] {
            assert_eq!(
                remove_sheet_references(source, "Data", &order).as_deref(),
                Some(removed),
                "{source}"
            );
        }
        for (source, removed) in [
            ("SUM(Data!A1:Data!B2)", "SUM(#REF!)"),
            (
                "SUM(Data!A1:INDEX(Data!A:A,2))",
                "SUM(#REF!:INDEX(#REF!,2))",
            ),
            ("Data!Café*2", "#REF!*2"),
            ("SUM(Data!A1:Total)", "SUM(#REF!:Total)"),
            ("SUM(Data!A1:Tax)", "SUM(#REF!:Tax)"),
            ("SUM(Data!A:Tax)", "SUM(#REF!)"),
            ("SUM(Data!$A$1:$B$2)", "SUM(#REF!)"),
            ("SUM(Data!1:3)", "SUM(#REF!)"),
        ] {
            assert_eq!(
                remove_sheet_references(source, "Data", &order).as_deref(),
                Some(removed),
                "{source}"
            );
        }
        for source in ["Other!A1", "\"Data!A1\"", "[1]Data!A1", "Data", "DataX!A1"] {
            assert_eq!(
                remove_sheet_references(source, "Data", &order),
                None,
                "{source}"
            );
        }
    }

    #[test]
    fn a_span_narrows_inward_or_becomes_ref_when_nothing_remains() {
        let order = sheets(&["Jan", "Feb", "Mar", "Apr"]);
        assert_eq!(
            remove_sheet_references("SUM(Jan:Mar!A1)", "Jan", &order).as_deref(),
            Some("SUM(Feb:Mar!A1)")
        );
        assert_eq!(
            remove_sheet_references("SUM(Jan:Mar!A1)", "Mar", &order).as_deref(),
            Some("SUM(Jan:Feb!A1)")
        );
        assert_eq!(
            remove_sheet_references("SUM(Jan:Mar!A1)", "Feb", &order),
            None
        );
        let months = sheets(&["Data", "Feb", "Mar"]);
        assert_eq!(
            remove_sheet_references("SUM(Data!A1:Feb:Mar!B2)", "Data", &months).as_deref(),
            Some("SUM(#REF!:Feb:Mar!B2)")
        );
        assert_eq!(
            remove_sheet_references("SUM(Data!A1:Data:Mar!B2)", "Data", &months).as_deref(),
            Some("SUM(#REF!:Feb:Mar!B2)")
        );
        let alone = sheets(&["Jan", "Other"]);
        assert_eq!(
            remove_sheet_references("SUM(Jan:Jan!A1:B2)", "Jan", &alone).as_deref(),
            Some("SUM(#REF!)")
        );
    }
}
