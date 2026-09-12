use std::collections::BTreeMap;

use crate::VsdxError;

pub const MAX_PATCH_EDITS: usize = 16_384;
pub const MAX_PATCH_BYTES: usize = 16 * 1024 * 1024;

#[derive(
    Clone, Copy, Debug, PartialEq, Eq, PartialOrd, Ord, serde::Serialize, serde::Deserialize,
)]
#[serde(rename_all = "camelCase")]
pub struct SourceSpan {
    pub offset: usize,
    pub length: usize,
}

impl SourceSpan {
    pub fn end(self) -> Option<usize> {
        self.offset.checked_add(self.length)
    }
}

#[derive(Clone, Debug, PartialEq, Eq)]
pub struct AttributeSpan {
    pub value: SourceSpan,
    pub quote: u8,
}

#[derive(Clone, Debug, PartialEq, Eq)]
pub struct ElementSpan {
    pub name: String,
    pub span: SourceSpan,
    pub start_tag: SourceSpan,
    pub attributes: BTreeMap<String, AttributeSpan>,
}

#[derive(Clone, Debug, PartialEq, Eq)]
pub struct SpanEdit {
    pub span: SourceSpan,
    pub replacement: Vec<u8>,
}

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub enum CellAttribute {
    Formula,
    Value,
}

impl CellAttribute {
    pub(crate) fn name(self) -> &'static str {
        match self {
            Self::Formula => "F",
            Self::Value => "V",
        }
    }
}

#[derive(Clone, Debug, PartialEq, Eq)]
pub struct CellEdit {
    pub part_path: String,
    pub cell_span: SourceSpan,
    pub attribute: CellAttribute,
    pub value: String,
}

pub fn apply_span_edits(source: &[u8], edits: &[SpanEdit]) -> Result<Vec<u8>, VsdxError> {
    if edits.len() > MAX_PATCH_EDITS {
        return Err(VsdxError::PatchLimit { kind: "editCount" });
    }
    let replacement_bytes = edits.iter().try_fold(0_usize, |total, edit| {
        total
            .checked_add(edit.replacement.len())
            .ok_or(VsdxError::PatchLimit { kind: "editBytes" })
    })?;
    if replacement_bytes > MAX_PATCH_BYTES {
        return Err(VsdxError::PatchLimit { kind: "editBytes" });
    }
    let mut ordered = edits.to_vec();
    ordered.sort_by_key(|edit| edit.span.offset);
    let mut cursor = 0;
    let mut previous_end = None;
    let mut output = Vec::with_capacity(source.len().saturating_add(replacement_bytes));
    for edit in ordered {
        let Some(end) = edit.span.end() else {
            return Err(VsdxError::InvalidSpan);
        };
        if end > source.len()
            || edit.span.offset < cursor
            || (previous_end == Some(edit.span.offset) && edit.span.length == 0)
        {
            return Err(VsdxError::InvalidSpan);
        }
        if !utf8_boundary(source, edit.span.offset) || !utf8_boundary(source, end) {
            return Err(VsdxError::InvalidSpan);
        }
        output.extend_from_slice(&source[cursor..edit.span.offset]);
        output.extend_from_slice(&edit.replacement);
        cursor = end;
        previous_end = Some(end);
    }
    output.extend_from_slice(&source[cursor..]);
    Ok(output)
}

pub fn escape_attribute_value(value: &str, _quote: u8) -> Result<Vec<u8>, VsdxError> {
    if !value.is_char_boundary(value.len()) {
        return Err(VsdxError::InvalidSpan);
    }
    let mut output = String::with_capacity(value.len());
    for character in value.chars() {
        if !xml_10_character(character) {
            return Err(VsdxError::InvalidXmlCharacter);
        }
        match character {
            '&' => output.push_str("&amp;"),
            '<' => output.push_str("&lt;"),
            '"' => output.push_str("&quot;"),
            '\'' => output.push_str("&apos;"),
            '\t' => output.push_str("&#x9;"),
            '\n' => output.push_str("&#xA;"),
            '\r' => output.push_str("&#xD;"),
            _ => output.push(character),
        }
    }
    Ok(output.into_bytes())
}

pub(crate) fn scan_element_spans(source: &[u8]) -> Result<Vec<ElementSpan>, VsdxError> {
    let mut spans = Vec::new();
    let mut stack: Vec<(usize, String, BTreeMap<String, AttributeSpan>)> = Vec::new();
    let mut cursor = 0;
    while let Some(relative) = source[cursor..].iter().position(|byte| *byte == b'<') {
        let start = cursor + relative;
        let token = token(source, start).ok_or(VsdxError::InvalidSpan)?;
        match token.kind {
            TokenKind::Ignored => {}
            TokenKind::Close(name) => {
                let Some((offset, open_name, attributes)) = stack.pop() else {
                    return Err(VsdxError::InvalidSpan);
                };
                if open_name != name {
                    return Err(VsdxError::InvalidSpan);
                }
                spans.push(ElementSpan {
                    name: open_name,
                    span: SourceSpan {
                        offset,
                        length: token.end + 1 - offset,
                    },
                    start_tag: SourceSpan {
                        offset,
                        length: token.end + 1 - offset,
                    },
                    attributes,
                });
            }
            TokenKind::Open => {
                if let Some((name, attributes, empty)) = opening_tag(source, start, token.end) {
                    if empty {
                        spans.push(ElementSpan {
                            name,
                            span: SourceSpan {
                                offset: start,
                                length: token.end + 1 - start,
                            },
                            start_tag: SourceSpan {
                                offset: start,
                                length: token.end + 1 - start,
                            },
                            attributes,
                        });
                    } else {
                        stack.push((start, name, attributes));
                    }
                }
            }
        }
        cursor = token.end + 1;
    }
    if stack.is_empty() {
        Ok(spans)
    } else {
        Err(VsdxError::InvalidSpan)
    }
}

struct Token {
    end: usize,
    kind: TokenKind,
}

enum TokenKind {
    Ignored,
    Open,
    Close(String),
}

fn token(source: &[u8], start: usize) -> Option<Token> {
    if source[start..].starts_with(b"<!--") {
        return terminated_token(source, start + 4, b"-->");
    }
    if source[start..].starts_with(b"<![CDATA[") {
        return terminated_token(source, start + 9, b"]]>");
    }
    if source[start..].starts_with(b"<?") {
        return terminated_token(source, start + 2, b"?>");
    }
    if source[start..].starts_with(b"<!DOCTYPE") {
        return doctype_token(source, start + 9);
    }
    if source[start..].starts_with(b"<!") {
        return terminated_token(source, start + 2, b">");
    }
    if source[start..].starts_with(b"</") {
        let end = tag_end(source, start)?;
        let mut index = start + 2;
        skip_space(source, &mut index, end);
        let name_start = index;
        while index < end && !source[index].is_ascii_whitespace() {
            index += 1;
        }
        if name_start == index {
            return None;
        }
        let name = String::from_utf8(source[name_start..index].to_vec()).ok()?;
        skip_space(source, &mut index, end);
        if index != end {
            return None;
        }
        return Some(Token {
            end,
            kind: TokenKind::Close(name),
        });
    }
    Some(Token {
        end: tag_end(source, start)?,
        kind: TokenKind::Open,
    })
}

fn terminated_token(source: &[u8], index: usize, terminator: &[u8]) -> Option<Token> {
    let end = source[index..]
        .windows(terminator.len())
        .position(|window| window == terminator)?
        + index
        + terminator.len()
        - 1;
    Some(Token {
        end,
        kind: TokenKind::Ignored,
    })
}

fn doctype_token(source: &[u8], mut index: usize) -> Option<Token> {
    let mut quote = None;
    let mut comment = false;
    let mut subset_depth = 0;
    while index < source.len() {
        if comment {
            if source[index..].starts_with(b"-->") {
                comment = false;
                index += 3;
                continue;
            }
            index += 1;
            continue;
        }
        if quote.is_none() && source[index..].starts_with(b"<!--") {
            comment = true;
            index += 4;
            continue;
        }
        match (quote, source[index]) {
            (Some(active), byte) if byte == active => quote = None,
            (Some(_), _) => {}
            (None, b'\'' | b'\"') => quote = Some(source[index]),
            (None, b'[') => subset_depth += 1,
            (None, b']') if subset_depth > 0 => subset_depth -= 1,
            (None, b'>') if subset_depth == 0 => {
                return Some(Token {
                    end: index,
                    kind: TokenKind::Ignored,
                });
            }
            _ => {}
        }
        index += 1;
    }
    None
}

fn tag_end(source: &[u8], start: usize) -> Option<usize> {
    let mut index = start + 1;
    let mut quote = None;
    while index < source.len() {
        let byte = source[index];
        if let Some(active) = quote {
            if byte == active {
                quote = None;
            }
        } else if byte == b'\'' || byte == b'"' {
            quote = Some(byte);
        } else if byte == b'>' {
            return Some(index);
        }
        index += 1;
    }
    None
}

fn opening_tag(
    source: &[u8],
    start: usize,
    end: usize,
) -> Option<(String, BTreeMap<String, AttributeSpan>, bool)> {
    let mut index = start + 1;
    skip_space(source, &mut index, end);
    let name_start = index;
    while index < end && !source[index].is_ascii_whitespace() && source[index] != b'/' {
        index += 1;
    }
    if name_start == index {
        return None;
    }
    let name = String::from_utf8(source[name_start..index].to_vec()).ok()?;
    let mut attributes = BTreeMap::new();
    while index < end {
        skip_space(source, &mut index, end);
        if index >= end || source[index] == b'/' {
            break;
        }
        let key_start = index;
        while index < end && !source[index].is_ascii_whitespace() && source[index] != b'=' {
            index += 1;
        }
        let key = String::from_utf8(source[key_start..index].to_vec()).ok()?;
        skip_space(source, &mut index, end);
        if index >= end || source[index] != b'=' {
            return None;
        }
        index += 1;
        skip_space(source, &mut index, end);
        if index >= end || !matches!(source[index], b'\'' | b'"') {
            return None;
        }
        let quote = source[index];
        let value_start = index + 1;
        index = value_start;
        while index < end && source[index] != quote {
            index += 1;
        }
        if index >= end {
            return None;
        }
        attributes.insert(
            key,
            AttributeSpan {
                value: SourceSpan {
                    offset: value_start,
                    length: index - value_start,
                },
                quote,
            },
        );
        index += 1;
    }
    Some((
        name,
        attributes,
        source[..end]
            .iter()
            .rposition(|byte| !byte.is_ascii_whitespace())
            .and_then(|index| source.get(index))
            == Some(&b'/'),
    ))
}

fn xml_10_character(character: char) -> bool {
    matches!(character as u32, 0x9 | 0xA | 0xD | 0x20..=0xD7FF | 0xE000..=0xFFFD | 0x10000..=0x10FFFF)
}

fn skip_space(source: &[u8], index: &mut usize, end: usize) {
    while *index < end && source[*index].is_ascii_whitespace() {
        *index += 1;
    }
}

fn utf8_boundary(source: &[u8], index: usize) -> bool {
    index == 0 || index == source.len() || source[index] & 0b1100_0000 != 0b1000_0000
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn patcher_preserves_unedited_bytes_and_rejects_overlap() {
        let source = b"a\xC3\xA4bc";
        assert_eq!(
            apply_span_edits(
                source,
                &[SpanEdit {
                    span: SourceSpan {
                        offset: 3,
                        length: 1
                    },
                    replacement: b"X".to_vec()
                }]
            )
            .unwrap(),
            b"a\xC3\xA4Xc"
        );
        assert!(matches!(
            apply_span_edits(
                b"abcd",
                &[
                    SpanEdit {
                        span: SourceSpan {
                            offset: 1,
                            length: 2
                        },
                        replacement: vec![]
                    },
                    SpanEdit {
                        span: SourceSpan {
                            offset: 2,
                            length: 1
                        },
                        replacement: vec![]
                    }
                ]
            ),
            Err(VsdxError::InvalidSpan)
        ));
        assert!(matches!(
            apply_span_edits(
                b"abcd",
                &[SpanEdit {
                    span: SourceSpan {
                        offset: 4,
                        length: 1
                    },
                    replacement: vec![]
                }]
            ),
            Err(VsdxError::InvalidSpan)
        ));
    }

    #[test]
    fn scanner_ignores_non_markup_and_tracks_nested_elements() {
        let source = b"<?pi value='>'?><Root><!-- <Cell V='x'/> --><![CDATA[<Cell> >]]><Cell V='x' /><Inner><Deep><Cell V='y'/></Deep></Inner></Root>";
        let spans = scan_element_spans(source).unwrap();
        let cells: Vec<_> = spans.iter().filter(|span| span.name == "Cell").collect();
        assert_eq!(cells.len(), 2);
        assert_eq!(cells[0].span.length, b"<Cell V='x' />".len());
        assert!(spans.iter().any(|span| span.name == "Inner"));
        assert!(spans.iter().any(|span| span.name == "Deep"));
        assert!(spans.iter().any(|span| span.name == "Root"));
    }

    #[test]
    fn scanner_ignores_doctype_comments() {
        let source = b"<!DOCTYPE Root [<!-- ]> <Cell V='fake'/> -->]><Root><Cell V='real'/></Root>";
        let spans = scan_element_spans(source).unwrap();
        let cells: Vec<_> = spans.iter().filter(|span| span.name == "Cell").collect();
        assert_eq!(cells.len(), 1);
        assert_eq!(
            &source[cells[0].span.offset..cells[0].span.end().unwrap()],
            b"<Cell V='real'/>"
        );
    }

    #[test]
    fn scanner_rejects_mismatched_close_tags() {
        assert!(matches!(
            scan_element_spans(b"<Root><Cell/></Wrong>"),
            Err(VsdxError::InvalidSpan)
        ));
    }

    #[test]
    fn attribute_escaping_rejects_invalid_xml_and_preserves_whitespace() {
        for value in ["bad\0", "bad\u{8}"] {
            assert!(matches!(
                escape_attribute_value(value, b'\''),
                Err(VsdxError::InvalidXmlCharacter)
            ));
        }
        assert_eq!(
            escape_attribute_value("a\tb\nc\rd", b'\'').unwrap(),
            b"a&#x9;b&#xA;c&#xD;d"
        );
    }

    #[test]
    fn patcher_rejects_duplicate_zero_length_spans() {
        assert!(matches!(
            apply_span_edits(
                b"abcd",
                &[
                    SpanEdit {
                        span: SourceSpan {
                            offset: 2,
                            length: 0
                        },
                        replacement: b"X".to_vec()
                    },
                    SpanEdit {
                        span: SourceSpan {
                            offset: 2,
                            length: 0
                        },
                        replacement: b"Y".to_vec()
                    },
                ],
            ),
            Err(VsdxError::InvalidSpan)
        ));
    }
}
