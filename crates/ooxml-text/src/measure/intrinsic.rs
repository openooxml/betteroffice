use crate::font_store::FontStore;
use crate::line_break::break_opportunities;

use super::input;
use super::line_filler::span_width;
use super::prepare::{PreparedRun, PreparedText, prepare_runs};
use super::{MAX_RUNS, MeasureError, MeasureRequest, list_marker, tabs};

#[derive(Default)]
struct Widths {
    minimum: f32,
    maximum: f32,
    line: f32,
    visible_line: f32,
    word: f32,
    visible_word: f32,
    indent: f32,
    first_line: f32,
    first_word: f32,
}

impl Widths {
    fn add(&mut self, width: f32, is_visible: bool) {
        if is_visible {
            self.visible_line = self.line + width;
            self.visible_word = self.word + width;
        }
        self.line += width;
        self.word += width;
    }

    fn end_word(&mut self) {
        if self.word == 0.0 {
            return;
        }
        self.minimum = self
            .minimum
            .max(self.visible_word + self.indent + self.first_word);
        self.word = 0.0;
        self.visible_word = 0.0;
        self.first_word = 0.0;
    }

    fn end_line(&mut self) {
        self.end_word();
        self.maximum = self
            .maximum
            .max(self.visible_line + self.indent + self.first_line);
        self.line = 0.0;
        self.visible_line = 0.0;
        self.first_line = 0.0;
        self.first_word = 0.0;
    }
}

fn text_widths(texts: &[(&str, &PreparedText)], widths: &mut Widths) {
    let text: String = texts.iter().map(|(text, _)| *text).collect();
    let mut previous = 0;
    let mut offset = 0;
    let boundaries: Vec<usize> = break_opportunities(&text)
        .into_iter()
        .map(|boundary| {
            offset += text[previous..boundary.byte_index].encode_utf16().count();
            previous = boundary.byte_index;
            offset
        })
        .collect();
    let mut boundary = 0;
    let mut run_start = 0;
    for (_, text) in texts {
        for cluster in &text.chars {
            while boundaries
                .get(boundary)
                .is_some_and(|limit| *limit <= run_start + cluster.utf16_offset as usize)
            {
                widths.end_word();
                boundary += 1;
            }
            let advance = cluster.advance_with_spacing(text.letter_spacing);
            widths.add(
                advance,
                !cluster.is_fit_space && (advance > 0.0 || text.letter_spacing < 0.0),
            );
        }
        run_start += text.utf16_len as usize;
    }
    widths.end_word();
}

fn following_width(runs: &[PreparedRun]) -> f32 {
    runs.iter()
        .take_while(|run| !matches!(run, PreparedRun::Tab(_) | PreparedRun::LineBreak))
        .map(|run| match run {
            PreparedRun::Text(text) => span_width(&text.chars, text.letter_spacing),
            PreparedRun::Field(field) => field.width,
            PreparedRun::InlineImage(image) | PreparedRun::OwnLineImage(image) => image.width,
            PreparedRun::SkippedImage { width, .. } => *width,
            _ => 0.0,
        })
        .sum()
}

/// Minimum unbreakable and maximum unwrapped paragraph widths, in pixels.
pub fn measure_intrinsic_widths(
    store: &FontStore,
    request: &MeasureRequest<'_>,
) -> Result<(f32, f32), MeasureError> {
    if request.block.kind != "paragraph" || request.block.runs.len() > MAX_RUNS {
        return Err(MeasureError::Unsupported(
            "intrinsic paragraph input".to_owned(),
        ));
    }
    input::validate_pt_size(request.defaults.font_size, "defaults.fontSize")?;
    let attrs = request.block.attrs.as_ref();
    let indent = attrs.and_then(|attrs| attrs.indent.as_ref());
    if let Some(indent) = indent {
        indent.validate()?;
    }
    let stops = attrs.and_then(|attrs| attrs.tabs.as_deref()).unwrap_or(&[]);
    input::validate_tabs(stops)?;
    let left = indent.and_then(|indent| indent.left).unwrap_or(0.0);
    let right = indent.and_then(|indent| indent.right).unwrap_or(0.0);
    let hanging = indent.and_then(|indent| indent.hanging).unwrap_or(0.0);
    let marker = match attrs {
        Some(attrs) if hanging == 0.0 => {
            list_marker::list_marker_inline_width(store, request, attrs)?
        }
        _ => 0.0,
    };
    let visible_marker = attrs.is_some_and(|attrs| {
        !attrs.list_marker_hidden
            && attrs
                .list_marker
                .as_ref()
                .is_some_and(|marker| !marker.is_empty())
    });
    let first_line = marker
        + if visible_marker && hanging > 0.0 {
            0.0
        } else {
            indent.and_then(|indent| indent.first_line).unwrap_or(0.0) - hanging
        };
    let prepared = prepare_runs(store, request)?;
    let mut widths = Widths {
        indent: left + right,
        first_line,
        first_word: first_line,
        ..Widths::default()
    };
    let mut index = 0;
    while index < prepared.len() {
        match &prepared[index] {
            PreparedRun::Text(_) => {
                let mut texts = Vec::new();
                while index < prepared.len() {
                    match &prepared[index] {
                        PreparedRun::Text(text) => texts.push((
                            request.block.runs[index].text.as_deref().unwrap_or(""),
                            text,
                        )),
                        PreparedRun::Hidden { .. } | PreparedRun::SkippedImage { .. } => {}
                        _ => break,
                    }
                    index += 1;
                }
                text_widths(&texts, &mut widths);
                continue;
            }
            PreparedRun::LineBreak => widths.end_line(),
            PreparedRun::Tab(_) => {
                widths.end_word();
                let advance = tabs::calculate_tab_width(
                    left + widths.line + widths.first_line,
                    stops,
                    tabs::px_to_twips(left),
                    following_width(&prepared[index + 1..]),
                    f32::MAX,
                )
                .width;
                widths.add(advance, advance > 0.0);
                widths.end_word();
            }
            PreparedRun::Field(field) => {
                widths.add(field.width, field.width > 0.0);
                widths.end_word();
            }
            PreparedRun::InlineImage(image) => {
                widths.end_word();
                widths.add(image.width, image.width > 0.0);
                widths.end_word();
            }
            PreparedRun::OwnLineImage(image) => {
                if widths.line > 0.0 {
                    widths.end_line();
                }
                widths.add(image.width, image.width > 0.0);
                widths.end_line();
            }
            PreparedRun::Hidden { .. } | PreparedRun::SkippedImage { .. } => {}
        }
        index += 1;
    }
    widths.end_line();
    Ok((
        widths.minimum.max(0.0),
        widths.maximum.max(widths.minimum).max(0.0),
    ))
}

#[cfg(test)]
mod tests {
    use super::super::line_filler::visible_span_width;
    use super::*;
    use crate::measure::MeasureInput;

    #[test]
    fn condensed_intrinsic_widths_include_nonpositive_visible_clusters() {
        let mut store = FontStore::new();
        let font = store
            .register(include_bytes!("../../tests/fonts/LiberationSans-Regular.ttf").to_vec())
            .unwrap();
        for source in ["WW\u{200b}", "WWi", "WWi\u{301}"] {
            for spacing in [-2.0, -3.0] {
                for trailing in ["", "  "] {
                    let input: MeasureInput = serde_json::from_value(serde_json::json!({
                        "block": {
                            "kind": "paragraph",
                            "runs": [{
                                "kind": "text", "text": format!("{source}{trailing}"),
                                "letterSpacing": spacing
                            }]
                        },
                        "maxWidth": 1000,
                        "fontChains": {"arial|0|0": [font.to_u32()]},
                        "defaults": {"fontFamily": "Arial", "fontSize": 6},
                        "authoritativeShaping": true
                    }))
                    .unwrap();
                    let request = MeasureRequest {
                        block: &input.block,
                        max_width: input.max_width,
                        font_chains: input::FontChains::Hash(&input.font_chains),
                        defaults: &input.defaults,
                        compat: input.compat,
                        floating_zones: None,
                        paragraph_y_offset: None,
                        authoritative_shaping: input.authoritative_shaping,
                    };
                    let prepared = prepare_runs(&store, &request).unwrap();
                    let PreparedRun::Text(text) = &prepared[0] else {
                        unreachable!()
                    };
                    let last = text
                        .chars
                        .iter()
                        .rfind(|cluster| !cluster.is_fit_space)
                        .unwrap();
                    assert!(last.advance_with_spacing(text.letter_spacing) <= 0.0);
                    let expected = visible_span_width(&text.chars, text.letter_spacing);
                    assert!(expected > 0.0);
                    let (minimum, maximum) = measure_intrinsic_widths(&store, &request).unwrap();
                    assert!((minimum - expected).abs() < 0.001, "{source:?}, {spacing}");
                    assert!((maximum - expected).abs() < 0.001, "{source:?}, {spacing}");
                }
            }
        }
    }
}
