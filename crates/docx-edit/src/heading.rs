//! Heading classification shared by the structured export and the session's heading reads.

use std::collections::BTreeMap;

use yrs::Any;

use crate::read_types::{HeadingInfo, OutlineSource};
use crate::seed::SourceMetadata;
use crate::{EditResult, EditingDoc, ParagraphId};

/// What a paragraph's outline level is read from.
pub(crate) struct OutlineInputs<'a> {
    /// The outline level the paragraph resolves to in the view being read, when known: the
    /// stream's value, which seeding and style changes resolve through the style chain.
    pub effective: Option<f64>,
    /// The paragraph's own outline level.
    pub direct: Option<f64>,
    /// The paragraph style as written.
    pub style_id: Option<&'a str>,
}

fn builtin_level(style_id: &str) -> Option<u8> {
    let digit = style_id
        .strip_prefix("Heading")
        .or_else(|| style_id.strip_prefix("heading"))?;
    match digit.as_bytes() {
        [digit @ b'1'..=b'9'] => Some(digit - b'1'),
        _ => None,
    }
}

/// The paragraph's heading: direct formatting wins, then its style with inheritance, then the
/// document defaults; only without any outline level does a `Heading1`..`Heading9` style id
/// count. Outline level 9 is body text.
pub(crate) fn resolve_heading(
    inputs: &OutlineInputs<'_>,
    source: Option<&SourceMetadata>,
) -> Option<HeadingInfo> {
    let effective_style = inputs
        .style_id
        .filter(|style| source.is_none_or(|source| source.has_style(style)))
        .or_else(|| source.and_then(SourceMetadata::default_paragraph_style));
    let style_level = effective_style
        .and_then(|style| source.and_then(|source| source.style_outline_level(style)));
    let default_level = source.and_then(SourceMetadata::default_outline_level);
    let level = inputs
        .effective
        .or(inputs.direct)
        .or(style_level)
        .or(default_level);
    let Some(level) = level else {
        let style_id = inputs.style_id?;
        return builtin_level(style_id).map(|outline_level| HeadingInfo {
            outline_level,
            source: OutlineSource::BuiltinStyleId {
                style_id: style_id.to_owned(),
            },
        });
    };
    let outline_level =
        (level.fract() == 0.0 && (0.0..=8.0).contains(&level)).then_some(level as u8)?;
    let source = if inputs.direct == Some(level) {
        OutlineSource::Direct
    } else if let Some(style) = effective_style.filter(|_| style_level == Some(level)) {
        OutlineSource::Style {
            style_id: style.to_owned(),
        }
    } else if default_level == Some(level) {
        OutlineSource::DocumentDefault
    } else {
        OutlineSource::Direct
    };
    Some(HeadingInfo {
        outline_level,
        source,
    })
}

fn number(value: Option<&Any>) -> Option<f64> {
    match value? {
        Any::Number(value) => Some(*value),
        Any::BigInt(value) => Some(*value as f64),
        _ => None,
    }
}

/// The outline inputs of a paragraph mark's current properties.
pub(crate) fn current_inputs(properties: &BTreeMap<String, Any>) -> OutlineInputs<'_> {
    OutlineInputs {
        effective: number(properties.get("outlineLevel")),
        direct: match properties.get("_originalFormatting") {
            Some(Any::Map(formatting)) => number(formatting.get("outlineLevel")),
            _ => None,
        },
        style_id: match properties.get("pStyle") {
            Some(Any::String(style)) => Some(style),
            _ => None,
        },
    }
}

impl EditingDoc {
    /// The headings of `story` in document order, classified the way the structured export
    /// classifies them.
    pub fn paragraph_headings(&self, story: &str) -> EditResult<Vec<(ParagraphId, HeadingInfo)>> {
        let source = self.source_metadata();
        Ok(self
            .paragraphs(story)?
            .into_iter()
            .filter_map(|paragraph| {
                let heading =
                    resolve_heading(&current_inputs(&paragraph.properties), source.as_deref())?;
                Some((paragraph.para_id, heading))
            })
            .collect())
    }
}
