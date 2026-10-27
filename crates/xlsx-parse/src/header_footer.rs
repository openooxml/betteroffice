use quick_xml::events::Event;
use serde::{Deserialize, Serialize};

use crate::package::{XmlAttribute, attributes_from_fragment};
use crate::read::unescape_name;
use crate::xml::{collect_text, local_name, next_event, reader};

#[derive(Clone, Debug, Default, Deserialize, Eq, PartialEq, Serialize)]
#[serde(rename_all = "camelCase")]
pub struct HeaderFooterText {
    pub raw: String,
    pub left: String,
    pub center: String,
    pub right: String,
}

#[derive(Clone, Debug, Default, Deserialize, Eq, PartialEq, Serialize)]
#[serde(rename_all = "camelCase")]
pub struct HeaderFooterPair {
    pub header: Option<HeaderFooterText>,
    pub footer: Option<HeaderFooterText>,
}

#[derive(Clone, Debug, Deserialize, Eq, PartialEq, Serialize)]
#[serde(rename_all = "camelCase")]
pub struct SheetHeaderFooter {
    pub different_first: bool,
    pub different_odd_even: bool,
    pub scale_with_doc: bool,
    pub align_with_margins: bool,
    pub odd: HeaderFooterPair,
    pub even: Option<HeaderFooterPair>,
    pub first: Option<HeaderFooterPair>,
}

/// Decodes stored header/footer text into printable sections and dialog fields.
pub fn parse_header_footer_text(source: &str) -> HeaderFooterText {
    let raw = unescape_name(source);
    let chars: Vec<char> = raw.chars().collect();
    let mut sections: [String; 3] = [String::new(), String::new(), String::new()];
    let mut current = 1usize;
    let mut index = 0usize;
    while index < chars.len() {
        let c = chars[index];
        if c == '\r' {
            sections[current].push('\n');
            index += 1;
            if chars.get(index) == Some(&'\n') {
                index += 1;
            }
            continue;
        }
        if c == '\n' {
            sections[current].push('\n');
            index += 1;
            continue;
        }
        if c != '&' {
            sections[current].push(c);
            index += 1;
            continue;
        }
        let Some(&next) = chars.get(index + 1) else {
            sections[current].push('&');
            break;
        };
        match next {
            '&' => {
                sections[current].push('&');
                index += 2;
            }
            'L' => {
                current = 0;
                index += 2;
            }
            'C' => {
                current = 1;
                index += 2;
            }
            'R' => {
                current = 2;
                index += 2;
            }
            'P' => {
                sections[current].push_str("&[Page]");
                index += 2;
            }
            'N' => {
                sections[current].push_str("&[Pages]");
                index += 2;
            }
            'D' => {
                sections[current].push_str("&[Date]");
                index += 2;
            }
            'T' => {
                sections[current].push_str("&[Time]");
                index += 2;
            }
            'F' => {
                sections[current].push_str("&[File]");
                index += 2;
            }
            'Z' => {
                sections[current].push_str("&[Path]");
                index += 2;
            }
            'A' => {
                sections[current].push_str("&[Tab]");
                index += 2;
            }
            'G' => {
                sections[current].push_str("&[Picture]");
                index += 2;
            }
            '"' => {
                index += 2;
                while index < chars.len() && chars[index] != '"' {
                    index += 1;
                }
                if index < chars.len() {
                    index += 1;
                }
            }
            '0'..='9' => {
                index += 1;
                while index < chars.len() && chars[index].is_ascii_digit() {
                    index += 1;
                }
            }
            'K' => {
                let rest: &[char] = if index + 2 <= chars.len() {
                    &chars[index + 2..]
                } else {
                    &[]
                };
                let theme = rest.len() >= 6
                    && rest[0].is_ascii_digit()
                    && rest[1].is_ascii_digit()
                    && matches!(rest[2], '+' | '-')
                    && rest[3..6].iter().all(char::is_ascii_digit);
                let hex = rest.len() >= 6 && rest[..6].iter().all(char::is_ascii_hexdigit);
                index += if theme || hex { 8 } else { 2 };
            }
            'B' | 'I' | 'U' | 'E' | 'S' | 'X' | 'Y' | 'O' | 'H' => {
                index += 2;
            }
            '\r' | '\n' => {
                sections[current].push('&');
                sections[current].push('\n');
                index += 2;
                if next == '\r' && chars.get(index) == Some(&'\n') {
                    index += 1;
                }
            }
            _ => {
                sections[current].push('&');
                sections[current].push(next);
                index += 2;
            }
        }
    }
    let [left, center, right] = sections;
    HeaderFooterText {
        raw,
        left,
        center,
        right,
    }
}

fn bool_attribute(attributes: &[XmlAttribute], name: &str, default: bool) -> bool {
    match attributes
        .iter()
        .find(|attribute| attribute.local_name() == name)
        .map(|attribute| attribute.value.as_str())
    {
        Some("1" | "true") => true,
        Some("0" | "false") => false,
        _ => default,
    }
}

fn part(stored: Option<String>) -> Option<HeaderFooterText> {
    match stored {
        Some(text) if !text.is_empty() => Some(parse_header_footer_text(&text)),
        _ => None,
    }
}

/// Reads printable header/footer variants from a worksheet element.
pub fn parse_header_footer(source: &[u8]) -> Option<SheetHeaderFooter> {
    let attributes = attributes_from_fragment(source).ok()?;
    let different_first = bool_attribute(&attributes, "differentFirst", false);
    let different_odd_even = bool_attribute(&attributes, "differentOddEven", false);
    let scale_with_doc = bool_attribute(&attributes, "scaleWithDoc", true);
    let align_with_margins = bool_attribute(&attributes, "alignWithMargins", true);
    let mut odd_header: Option<String> = None;
    let mut odd_footer: Option<String> = None;
    let mut even_header: Option<String> = None;
    let mut even_footer: Option<String> = None;
    let mut first_header: Option<String> = None;
    let mut first_footer: Option<String> = None;
    let mut reader = reader(source);
    let mut buf = Vec::new();
    let mut depth = 0usize;
    loop {
        match next_event(&mut reader, &mut buf, &mut depth) {
            Ok(Event::Start(element)) => {
                let name = local_name(&element);
                let target = match name.as_slice() {
                    b"oddHeader" => &mut odd_header,
                    b"oddFooter" => &mut odd_footer,
                    b"evenHeader" => &mut even_header,
                    b"evenFooter" => &mut even_footer,
                    b"firstHeader" => &mut first_header,
                    b"firstFooter" => &mut first_footer,
                    _ => continue,
                };
                *target = Some(collect_text(&mut reader, &mut buf, &mut depth).ok()?);
            }
            Ok(Event::Eof) => break,
            Ok(_) => {}
            Err(_) => return None,
        }
    }
    let any_part = [
        &odd_header,
        &odd_footer,
        &even_header,
        &even_footer,
        &first_header,
        &first_footer,
    ]
    .iter()
    .any(|stored| stored.as_deref().is_some_and(|text| !text.is_empty()));
    if !any_part {
        return None;
    }
    let odd = HeaderFooterPair {
        header: part(odd_header),
        footer: part(odd_footer),
    };
    let even = different_odd_even.then(|| HeaderFooterPair {
        header: part(even_header),
        footer: part(even_footer),
    });
    let first = different_first.then(|| HeaderFooterPair {
        header: part(first_header),
        footer: part(first_footer),
    });
    Some(SheetHeaderFooter {
        different_first,
        different_odd_even,
        scale_with_doc,
        align_with_margins,
        odd,
        even,
        first,
    })
}
