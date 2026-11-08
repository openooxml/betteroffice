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
