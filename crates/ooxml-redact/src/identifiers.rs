use std::collections::HashMap;

use crate::RedactError;
use crate::mask::TextMasker;

/// Neutral names for user-authored bookmark and VML shape identifiers, shared across a package.
#[derive(Default)]
pub(crate) struct Identifiers {
    bookmarks: HashMap<String, String>,
    shapes: HashMap<String, String>,
}

impl Identifiers {
    pub(crate) fn bookmark(&mut self, name: &str) -> String {
        if name.is_empty() || generated_bookmark(name) {
            return name.to_owned();
        }
        let next = self.bookmarks.len() + 1;
        self.bookmarks
            .entry(name.to_lowercase())
            .or_insert_with(|| {
                let hidden = if name.starts_with('_') { "_" } else { "" };
                format!("{hidden}Bm{next}")
            })
            .clone()
    }

    pub(crate) fn shape(&mut self, id: &str) -> String {
        if id.is_empty() || generated_shape(id) {
            return id.to_owned();
        }
        let next = self.shapes.len() + 1;
        self.shapes
            .entry(id.to_owned())
            .or_insert_with(|| format!("Shape{next}"))
            .clone()
    }

    pub(crate) fn shape_reference(&mut self, reference: &str) -> String {
        match reference.strip_prefix('#') {
            Some(id) => format!("#{}", self.shape(id)),
            None => self.shape(reference),
        }
    }

    pub(crate) fn ole_object(&mut self, id: &str) -> String {
        if id
            .strip_prefix('_')
            .is_some_and(|digits| !digits.is_empty() && digits.bytes().all(|b| b.is_ascii_digit()))
        {
            id.to_owned()
        } else {
            self.shape(id)
        }
    }

    pub(crate) fn vml_style(&mut self, style: &str) -> String {
        style
            .split(';')
            .map(|declaration| match declaration.split_once(':') {
                Some((property, value))
                    if property.trim().eq_ignore_ascii_case("mso-next-textbox") =>
                {
                    match value.trim().strip_prefix('#') {
                        Some(id) => format!("{property}:#{}", self.shape(id)),
                        None => declaration.to_owned(),
                    }
                }
                _ => declaration.to_owned(),
            })
            .collect::<Vec<_>>()
            .join(";")
    }
}

fn digits_after(value: &str, prefix: &str) -> bool {
    value
        .strip_prefix(prefix)
        .is_some_and(|rest| !rest.is_empty() && rest.bytes().all(|b| b.is_ascii_digit()))
}

fn generated_bookmark(name: &str) -> bool {
    let lower = name.to_ascii_lowercase();
    lower == "_goback"
        || ["_toc", "_ref", "_hlk"]
            .iter()
            .any(|prefix| digits_after(&lower, prefix))
}

fn generated_shape(id: &str) -> bool {
    id.strip_prefix("_x0000_").is_some_and(|rest| {
        rest.chars()
            .next()
            .is_some_and(|kind| kind.is_ascii_lowercase() && digits_after(rest, &kind.to_string()))
    })
}

/// Rewrites a field instruction that names a bookmark, or `None` for other fields.
pub(crate) fn rewrite_instruction(
    instruction: &str,
    masker: &mut TextMasker,
) -> Result<Option<String>, RedactError> {
    let Some(edits) = instruction_edits(instruction, &[], masker)? else {
        return Ok(None);
    };
    let mut output = String::with_capacity(instruction.len());
    let mut copied = 0;
    for edit in edits {
        output.push_str(&instruction[copied..edit.offset]);
        output.push_str(&edit.replacement);
        copied = edit.offset + edit.length;
    }
    output.push_str(&instruction[copied..]);
    Ok(Some(output))
}

pub(crate) struct InstructionEdit {
    pub(crate) offset: usize,
    pub(crate) length: usize,
    pub(crate) replacement: String,
}

pub(crate) fn instruction_edits(
    instruction: &str,
    breaks: &[usize],
    masker: &mut TextMasker,
) -> Result<Option<Vec<InstructionEdit>>, RedactError> {
    let tokens = tokenize(instruction);
    let Some(keyword) = tokens.iter().find(|token| !token.space) else {
        return Ok(None);
    };
    let field = keyword.text.to_ascii_uppercase();
    if !matches!(
        field.as_str(),
        "REF" | "PAGEREF" | "NOTEREF" | "HYPERLINK" | "SET" | "ASK" | "GOTOBUTTON"
    ) {
        return Ok(None);
    }
    let mut edits = Vec::new();
    let mut seen_keyword = false;
    let mut positional = 0;
    let mut pending = None;
    for token in &tokens {
        if token.space || !seen_keyword {
            seen_keyword |= !token.space;
            continue;
        }
        if let Some(switch) = switch(token.text) {
            pending = takes_argument(&field, switch).then_some(switch);
            continue;
        }
        let bookmark = match pending.take() {
            Some('l') if field == "HYPERLINK" => true,
            Some('*')
                if FORMAT_SWITCHES
                    .iter()
                    .any(|format| format.eq_ignore_ascii_case(token.text)) =>
            {
                continue;
            }
            Some(_) => false,
            None => {
                positional += 1;
                positional == 1 && field != "HYPERLINK"
            }
        };
        let (offset, text) = match token.text.strip_prefix('"') {
            Some(inner) => (token.offset + 1, inner.strip_suffix('"').unwrap_or(inner)),
            None => (token.offset, token.text),
        };
        let token_end = offset + text.len();
        if bookmark {
            if breaks
                .iter()
                .any(|&boundary| offset < boundary && boundary < token_end)
            {
                return Err(RedactError::Xml {
                    part: "field instruction".to_owned(),
                    message: "bookmark argument crosses a nested field".to_owned(),
                });
            }
            edits.push(InstructionEdit {
                offset,
                length: text.len(),
                replacement: masker.ids.bookmark(text),
            });
        } else {
            let mut start = offset;
            for end in breaks
                .iter()
                .copied()
                .filter(|&boundary| offset < boundary && boundary < token_end)
                .chain(std::iter::once(token_end))
            {
                let text = &instruction[start..end];
                let replacement = masker.replace(text)?;
                if replacement != text {
                    edits.push(InstructionEdit {
                        offset: start,
                        length: end - start,
                        replacement,
                    });
                }
                start = end;
            }
        }
    }
    Ok(Some(edits))
}

const FORMAT_SWITCHES: &[&str] = &[
    "MERGEFORMAT",
    "CHARFORMAT",
    "Caps",
    "FirstCap",
    "Lower",
    "Upper",
    "Arabic",
    "ArabicDash",
    "Alphabetic",
    "Roman",
    "Hex",
    "Ordinal",
    "CardText",
    "DollarText",
    "OrdText",
];

struct Token<'a> {
    offset: usize,
    text: &'a str,
    space: bool,
}

fn tokenize(instruction: &str) -> Vec<Token<'_>> {
    let mut tokens = Vec::new();
    let mut rest = instruction;
    while let Some(first) = rest.chars().next() {
        let length = if first.is_whitespace() {
            rest.find(|c: char| !c.is_whitespace())
                .unwrap_or(rest.len())
        } else if switch_prefix(rest).is_some() {
            2
        } else if first == '"' {
            let mut escaped = false;
            rest.char_indices()
                .skip(1)
                .find(|&(_, c)| {
                    let close = c == '"' && !escaped;
                    escaped = c == '\\' && !escaped;
                    close
                })
                .map_or(rest.len(), |(index, _)| index + 1)
        } else {
            rest.char_indices()
                .skip(1)
                .find(|&(index, c)| {
                    c.is_whitespace() || c == '"' || switch_prefix(&rest[index..]).is_some()
                })
                .map_or(rest.len(), |(index, _)| index)
        };
        tokens.push(Token {
            offset: instruction.len() - rest.len(),
            text: &rest[..length],
            space: first.is_whitespace(),
        });
        rest = &rest[length..];
    }
    tokens
}

fn switch_prefix(token: &str) -> Option<char> {
    let mut characters = token.chars();
    (characters.next() == Some('\\'))
        .then(|| characters.next())
        .flatten()
        .filter(|c| c.is_ascii_alphabetic() || matches!(*c, '*' | '@' | '#' | '!'))
}

fn switch(token: &str) -> Option<char> {
    switch_prefix(token)
        .filter(|_| token.len() == 2)
        .map(|switch| switch.to_ascii_lowercase())
}

fn takes_argument(field: &str, switch: char) -> bool {
    matches!(switch, '*' | '#' | '@')
        || matches!(
            (field, switch),
            ("REF" | "ASK", 'd') | ("HYPERLINK", 'l' | 'o' | 't')
        )
}

/// Neutral replacement for a custom theme, color, font, or effect scheme name.
pub(crate) fn theme_name(element: &str, value: &str) -> Option<&'static str> {
    let neutral = match element {
        "theme" | "themeFamily" => "Theme",
        "clrScheme" => "Colors",
        "fontScheme" => "Fonts",
        "fmtScheme" => "Effects",
        _ => return None,
    };
    (!BUILT_IN_THEME_NAMES.contains(&value)).then_some(neutral)
}

const BUILT_IN_THEME_NAMES: &[&str] = &[
    "Office Theme",
    "Office",
    "Office 2007 - 2010",
    "Office 2007 - 2010 Theme",
    "Office 2013 - 2022",
    "Office 2013 - 2022 Theme",
    "Office Classic",
    "Office Classic 2",
    "Office-Design",
    "Thème Office",
    "Tema de Office",
    "Tema di Office",
    "Adjacency",
    "Angles",
    "Apex",
    "Apothecary",
    "Aspect",
    "Atlas",
    "Austin",
    "Badge",
    "Banded",
    "Basis",
    "Berlin",
    "Black Tie",
    "Blue",
    "Blue II",
    "Blue Green",
    "Blue Warm",
    "Celestial",
    "Circuit",
    "Civic",
    "Clarity",
    "Composite",
    "Concourse",
    "Couture",
    "Crop",
    "Damask",
    "Depth",
    "Dividend",
    "Droplet",
    "Elemental",
    "Equity",
    "Essential",
    "Executive",
    "Facet",
    "Feathered",
    "Flow",
    "Foundry",
    "Frame",
    "Gallery",
    "Gift",
    "Grayscale",
    "Green",
    "Green Yellow",
    "Grid",
    "Hardcover",
    "Headlines",
    "Horizon",
    "Integral",
    "Ion",
    "Ion Boardroom",
    "Madison",
    "Main Event",
    "Marquee",
    "Median",
    "Mesh",
    "Metro",
    "Metropolitan",
    "Module",
    "Newsprint",
    "Opulent",
    "Orange",
    "Orange Red",
    "Organic",
    "Oriel",
    "Origin",
    "Paper",
    "Parallax",
    "Parcel",
    "Perspective",
    "Pushpin",
    "Quotable",
    "Red",
    "Red Orange",
    "Red Violet",
    "Retrospect",
    "Savon",
    "Slice",
    "Slipstream",
    "Solstice",
    "Technic",
    "Thatch",
    "Trek",
    "Urban",
    "Vapor Trail",
    "Verve",
    "View",
    "Violet",
    "Violet II",
    "Waveform",
    "Wisp",
    "Wood Type",
    "Yellow",
    "Yellow Orange",
    "Arial",
    "Arial Black-Arial",
    "Calibri",
    "Calibri Light-Constantia",
    "Calibri-Cambria",
    "Cambria",
    "Candara",
    "Century Gothic",
    "Consolas-Verdana",
    "Constantia-Franklin Gothic Book",
    "Corbel",
    "Franklin Gothic",
    "Garamond",
    "Garamond-TrebuchetMs",
    "Georgia",
    "Gill Sans MT",
    "Times New Roman",
    "Trebuchet MS",
    "Tw Cen MT",
    "Tw Cen MT Condensed",
    "Verdana",
    "Banded Edge",
    "Extreme Shadow",
    "Frosted Glass",
    "Glossy",
    "Glow Edge",
    "Grunge Texture",
    "Inset",
    "Milk Glass",
    "Reflection",
    "Riblet",
    "Subtle Solids",
    "Top Shadow",
];
