//! Typed bridge between layout blocks and `ooxml-text` measurement.
//!
//! Field-for-field equivalent of the JSON envelope this used to serialize:
//! every mapping here mirrors what serde did at that boundary, including its
//! strictness — a value the JSON parse would have rejected makes
//! [`measure_paragraph`] return `None` so the caller falls back to the
//! synthetic extent, exactly as an `"invalid: "` error used to.

use ooxml_text::measure::{
    AttrsIn, BlockIn, CompatIn, DefaultsIn, FloatZoneIn, FontChains, IndentIn, MeasureRequest,
    RotationBoundsIn, RunFontSlotsIn, RunIn, RunLanguageSlotsIn, SpacingIn, TabStopIn,
};
use serde_json::Value;

use crate::measure_blocks::{FloatingZone, MeasurementConfig};
use crate::types::{
    ImageRun, ParagraphAttrs, ParagraphBlock, Run, RunFormatting, TabStop, TypesetBidiSlice,
    TypesetClusterAdvance, TypesetRow, TypesetRowSegment, TypesetRunAdvance,
};

/// Measure one paragraph through the typed path; `None` mirrors the old
/// fallbacks (unparseable envelope, `UNSUPPORTED`, or a measurement error).
pub(crate) fn measure_paragraph(
    paragraph: &ParagraphBlock,
    content_width: f64,
    config: &MeasurementConfig,
    floating_zones: Option<&[FloatingZone]>,
    cumulative_y: f64,
) -> Option<crate::types::ParagraphExtent> {
    let block = block_in(paragraph)?;
    let defaults = defaults_in(&config.defaults)?;
    let compat = compat_in(&config.compat)?;
    let zones = floating_zones.map(|zones| zones.iter().map(zone_in).collect::<Vec<FloatZoneIn>>());
    let request = MeasureRequest {
        block: &block,
        max_width: content_width as f32,
        font_chains: FontChains::BTree(&config.font_chains),
        defaults: &defaults,
        compat,
        floating_zones: zones.as_deref(),
        paragraph_y_offset: floating_zones.is_some().then(|| cumulative_y as f32),
        authoritative_shaping: config.authoritative_shaping,
    };
    let extent = crate::measure_paragraph_typed_resident(&request).ok()?;
    Some(extent_from_out(extent))
}

fn block_in(paragraph: &ParagraphBlock) -> Option<BlockIn> {
    Some(BlockIn {
        kind: "paragraph".to_owned(),
        runs: paragraph
            .runs
            .iter()
            .map(run_in)
            .collect::<Option<Vec<_>>>()?,
        attrs: paragraph.attrs.as_ref().map(attrs_in),
    })
}

fn defaults_in(defaults: &Value) -> Option<DefaultsIn> {
    let fields = defaults.as_object()?;
    Some(DefaultsIn {
        font_size: fields.get("fontSize")?.as_f64()? as f32,
        font_family: fields.get("fontFamily")?.as_str()?.to_owned(),
    })
}

fn compat_in(compat: &Value) -> Option<CompatIn> {
    match compat {
        Value::Null => Some(CompatIn::default()),
        Value::Object(fields) => {
            let mut parsed = CompatIn::default();
            parsed.no_leading = bool_flag(fields.get("noLeading"))?;
            parsed.do_not_expand_shift_return = bool_flag(fields.get("doNotExpandShiftReturn"))?;
            Some(parsed)
        }
        _ => None,
    }
}

fn bool_flag(value: Option<&Value>) -> Option<bool> {
    match value {
        None => Some(false),
        Some(Value::Bool(flag)) => Some(*flag),
        Some(_) => None,
    }
}

fn number_f32(value: Option<&Value>) -> Option<Option<f32>> {
    match value {
        None | Some(Value::Null) => Some(None),
        Some(Value::Number(number)) => Some(Some(number.as_f64()? as f32)),
        Some(_) => None,
    }
}

fn zone_in(zone: &FloatingZone) -> FloatZoneIn {
    FloatZoneIn {
        left_margin: zone.left_margin as f32,
        right_margin: zone.right_margin as f32,
        top_y: zone.top_y as f32,
        bottom_y: zone.bottom_y as f32,
        segments: None,
        full_width_block: zone.full_width_block,
    }
}

fn attrs_in(attrs: &ParagraphAttrs) -> AttrsIn {
    AttrsIn {
        alignment: attrs.alignment.clone(),
        spacing: attrs.spacing.as_ref().map(|spacing| SpacingIn {
            before: spacing.before.map(|px| px as f32),
            after: spacing.after.map(|px| px as f32),
            line: spacing.line.map(|px| px as f32),
            line_unit: spacing.line_unit.clone(),
            line_rule: spacing.line_rule.clone(),
        }),
        indent: attrs.indent.as_ref().map(|indent| IndentIn {
            left: indent.left.map(|px| px as f32),
            right: indent.right.map(|px| px as f32),
            first_line: indent.first_line.map(|px| px as f32),
            hanging: indent.hanging.map(|px| px as f32),
        }),
        tabs: attrs
            .tabs
            .as_ref()
            .map(|stops| stops.iter().map(tab_stop_in).collect()),
        bidi: attrs.bidi.unwrap_or(false),
        default_font_size: attrs.default_font_size.map(|pt| pt as f32),
        default_font_family: attrs.default_font_family.clone(),
        suppress_empty_paragraph_height: attrs.suppress_empty_paragraph_height.unwrap_or(false),
        list_marker: attrs.list_marker.clone(),
        list_marker_hidden: attrs.list_marker_hidden.unwrap_or(false),
        list_marker_font_family: attrs.list_marker_font_family.clone(),
        list_marker_font_size: attrs.list_marker_font_size.map(|pt| pt as f32),
        list_marker_suffix: attrs.list_marker_suffix.clone(),
        default_tab_stop_twips: attrs.default_tab_stop_twips.map(|twips| twips as f32),
    }
}

fn tab_stop_in(stop: &TabStop) -> TabStopIn {
    TabStopIn {
        val: stop.val.clone(),
        pos: stop.pos as f32,
    }
}

fn run_in(run: &Run) -> Option<RunIn> {
    match run {
        Run::Text(text) => {
            let mut out = formatted_run("text", &text.fmt);
            out.text = Some(text.text.clone());
            Some(out)
        }
        Run::Tab(tab) => Some(formatted_run("tab", &tab.fmt)),
        Run::LineBreak(_) => Some(bare_run("lineBreak")),
        Run::Field(field) => {
            let mut out = formatted_run("field", &field.fmt);
            out.fallback = field.fallback.clone();
            Some(out)
        }
        Run::Image(image) => image_run(image),
        Run::Unsupported => Some(bare_run("unsupported")),
    }
}

fn formatted_run(kind: &str, fmt: &RunFormatting) -> RunIn {
    let mut out = bare_run(kind);
    out.bold = fmt.bold.unwrap_or(false);
    out.italic = fmt.italic.unwrap_or(false);
    out.bold_cs = fmt.bold_cs;
    out.italic_cs = fmt.italic_cs;
    out.font_size = fmt.font_size.map(|pt| pt as f32);
    out.font_size_cs = fmt.font_size_cs.map(|pt| pt as f32);
    out.font_family = fmt.font_family.clone();
    out.font_slots = fmt.font_slots.as_ref().map(|slots| RunFontSlotsIn {
        ascii: slots.ascii.clone(),
        h_ansi: slots.h_ansi.clone(),
        east_asia: slots.east_asia.clone(),
        cs: slots.cs.clone(),
        hint: slots.hint.clone(),
    });
    out.complex_script = fmt.complex_script.unwrap_or(false);
    out.language = fmt.language.as_ref().map(|slots| RunLanguageSlotsIn {
        latin: slots.latin.clone(),
        east_asia: slots.east_asia.clone(),
        bidi: slots.bidi.clone(),
    });
    out.letter_spacing = fmt.letter_spacing.map(|spacing| spacing as f32);
    out.all_caps = fmt.all_caps.unwrap_or(false);
    out.small_caps = fmt.small_caps.unwrap_or(false);
    out.horizontal_scale = fmt.horizontal_scale.map(|scale| scale as f32);
    out.kerning_min_pt = fmt.kerning_min_pt.map(|kerning| kerning as f32);
    out.superscript = fmt.superscript.unwrap_or(false);
    out.subscript = fmt.subscript.unwrap_or(false);
    out.hidden = fmt.hidden.unwrap_or(false);
    out.rtl = fmt.rtl.unwrap_or(false);
    out
}

fn bare_run(kind: &str) -> RunIn {
    RunIn {
        kind: kind.to_owned(),
        text: None,
        bold: false,
        italic: false,
        bold_cs: None,
        italic_cs: None,
        font_size: None,
        font_size_cs: None,
        font_family: None,
        font_slots: None,
        complex_script: false,
        language: None,
        letter_spacing: None,
        all_caps: false,
        small_caps: false,
        horizontal_scale: None,
        kerning_min_pt: None,
        superscript: false,
        subscript: false,
        hidden: false,
        rtl: false,
        fallback: None,
        width: None,
        height: None,
        rotation_bounds: None,
        dist_top: None,
        dist_bottom: None,
        wrap_type: None,
        display_mode: None,
        position: None,
    }
}

fn image_run(image: &ImageRun) -> Option<RunIn> {
    let mut out = bare_run("image");
    out.width = Some(image.width as f32);
    out.height = Some(image.height as f32);
    out.rotation_bounds = rotation_bounds_in(image.rotation_bounds.as_ref())?;
    out.dist_top = image.dist_top.map(|distance| distance as f32);
    out.dist_bottom = image.dist_bottom.map(|distance| distance as f32);
    out.wrap_type = image.wrap_type.clone();
    out.display_mode = image.display_mode.clone();
    out.position = image.position.as_ref().map(|_| serde::de::IgnoredAny);
    Some(out)
}

/// Mirrors serde's strictness at the old JSON boundary: absent or `null` is
/// no bounds; anything else must be an object or a positional sequence of
/// the two optional dimensions, exactly as serde's derived `Deserialize`
/// accepts — any deviation fails the whole request (`None`) where envelope
/// parsing used to.
fn rotation_bounds_in(bounds: Option<&Value>) -> Option<Option<RotationBoundsIn>> {
    match bounds {
        None | Some(Value::Null) => Some(None),
        Some(Value::Object(fields)) => Some(Some(RotationBoundsIn {
            width: number_f32(fields.get("width"))?,
            height: number_f32(fields.get("height"))?,
        })),
        Some(Value::Array(items)) => {
            if items.len() > 2 {
                return None;
            }
            Some(Some(RotationBoundsIn {
                width: number_f32(items.first())?,
                height: number_f32(items.get(1))?,
            }))
        }
        Some(_) => None,
    }
}

fn extent_from_out(extent: ooxml_text::ParagraphExtentOut) -> crate::types::ParagraphExtent {
    crate::types::ParagraphExtent {
        lines: extent.lines.into_iter().map(row_from_out).collect(),
        total_height: extent.total_height as f64,
    }
}

fn row_from_out(row: ooxml_text::TypesetRowOut) -> TypesetRow {
    TypesetRow {
        head_run: row.head_run as usize,
        head_char: row.head_char as usize,
        tail_run: row.tail_run as usize,
        tail_char: row.tail_char as usize,
        width: row.width as f64,
        ascent: row.ascent as f64,
        descent: row.descent as f64,
        line_height: row.line_height as f64,
        synthetic_fallback: None,
        left_offset: row.left_offset.map(f64::from),
        right_offset: row.right_offset.map(f64::from),
        segments: row.segments.map(|segments| {
            segments
                .into_iter()
                .map(|segment| TypesetRowSegment {
                    head_run: segment.head_run as usize,
                    head_char: segment.head_char as usize,
                    tail_run: segment.tail_run as usize,
                    tail_char: segment.tail_char as usize,
                    left_offset: segment.left_offset as f64,
                    available_width: segment.available_width as f64,
                    width: segment.width as f64,
                })
                .collect()
        }),
        float_skip_before: row.float_skip_before.map(f64::from),
        run_advances: row.run_advances.map(|advances| {
            advances
                .into_iter()
                .map(|advance| TypesetRunAdvance {
                    run_index: Some(u64::from(advance.run_index)),
                    start_char: Some(u64::from(advance.start_char)),
                    end_char: Some(u64::from(advance.end_char)),
                    advance: Some(advance.advance as f64),
                    logical_order: Some(u64::from(advance.logical_order)),
                })
                .collect()
        }),
        cluster_advances: row.cluster_advances.map(|clusters| {
            clusters
                .into_iter()
                .map(|cluster| TypesetClusterAdvance {
                    run_index: Some(u64::from(cluster.run_index)),
                    start_char: Some(u64::from(cluster.start_char)),
                    end_char: Some(u64::from(cluster.end_char)),
                    advance: Some(cluster.advance as f64),
                    x_offset: Some(cluster.x_offset as f64),
                    bidi_level: Some(cluster.bidi_level),
                    logical_order: Some(u64::from(cluster.logical_order)),
                })
                .collect()
        }),
        bidi_slices: row.bidi_slices.map(|slices| {
            slices
                .into_iter()
                .map(|slice| TypesetBidiSlice {
                    run_index: Some(u64::from(slice.run_index)),
                    start_char: Some(u64::from(slice.start_char)),
                    end_char: Some(u64::from(slice.end_char)),
                    advance: Some(slice.advance as f64),
                    bidi_level: Some(slice.bidi_level),
                    visual_order: Some(u64::from(slice.visual_order)),
                    logical_order: Some(u64::from(slice.logical_order)),
                })
                .collect()
        }),
    }
}

#[cfg(test)]
mod parity_tests {
    use super::*;
    use crate::measure_blocks::measure_paragraph as measure_paragraph_entry;
    use crate::types::{BlockId, FieldRun, LayoutBlock, ParagraphExtent, TabRun, TextRun};
    use serde_json::json;
    use std::collections::BTreeMap;

    const FIXTURE: &[u8] =
        include_bytes!("../../ooxml-text/tests/fonts/LiberationSans-Regular.ttf");

    struct Fixture {
        config: MeasurementConfig,
    }

    fn fixture() -> Fixture {
        crate::clear_measure_fonts();
        let id = crate::register_measure_font(FIXTURE).expect("register fixture font");
        let mut font_chains = BTreeMap::new();
        for bold in 0..2 {
            for italic in 0..2 {
                font_chains.insert(format!("liberation sans|{bold}|{italic}"), vec![id]);
            }
        }
        Fixture {
            config: MeasurementConfig {
                font_chains,
                defaults: json!({ "fontSize": 12.0, "fontFamily": "Liberation Sans" }),
                compat: json!({}),
                authoritative_shaping: true,
            },
        }
    }

    fn text_run(text: &str) -> Run {
        Run::Text(TextRun {
            fmt: RunFormatting::default(),
            text: text.to_owned(),
            pm_start: None,
            pm_end: None,
            inline_sdt_widget: None,
        })
    }

    fn paragraph(runs: Vec<Run>, attrs: Option<crate::types::ParagraphAttrs>) -> ParagraphBlock {
        ParagraphBlock {
            sdt_groups: None,
            id: BlockId::Num(0.0),
            para_id: None,
            runs,
            attrs,
            pm_start: None,
            pm_end: None,
        }
    }

    fn legacy_envelope(
        paragraph: &ParagraphBlock,
        content_width: f64,
        config: &MeasurementConfig,
        floating_zones: Option<&[FloatingZone]>,
        cumulative_y: f64,
    ) -> String {
        let mut envelope = json!({
            "block": LayoutBlock::Paragraph(paragraph.clone()),
            "maxWidth": content_width,
            "fontChains": config.font_chains,
            "authoritativeShaping": config.authoritative_shaping,
        });
        let fields = envelope.as_object_mut().unwrap();
        if !config.defaults.is_null() {
            fields.insert("defaults".to_owned(), config.defaults.clone());
        }
        if !config.compat.is_null() {
            fields.insert("compat".to_owned(), config.compat.clone());
        }
        if let Some(zones) = floating_zones {
            fields.insert(
                "floatingZones".to_owned(),
                json!(
                    zones
                        .iter()
                        .map(|zone| {
                            json!({
                                "leftMargin": zone.left_margin,
                                "rightMargin": zone.right_margin,
                                "topY": zone.top_y,
                                "bottomY": zone.bottom_y,
                                "fullWidthBlock": zone.full_width_block,
                            })
                        })
                        .collect::<Vec<_>>()
                ),
            );
            fields.insert("paragraphYOffset".to_owned(), json!(cumulative_y));
        }
        envelope.to_string()
    }

    fn assert_close(name: &str, left: &Value, right: &Value, path: &str) {
        match (left, right) {
            (Value::Number(left), Value::Number(right)) => {
                let left = left.as_f64().unwrap();
                let right = right.as_f64().unwrap();
                let tolerance = 1e-4_f64.max(left.abs() * 1e-5);
                assert!(
                    (left - right).abs() <= tolerance,
                    "{name}: {path} {left} != {right}"
                );
            }
            (Value::Array(left), Value::Array(right)) => {
                assert_eq!(left.len(), right.len(), "{name}: {path} length");
                for (index, (a, b)) in left.iter().zip(right).enumerate() {
                    assert_close(name, a, b, &format!("{path}[{index}]"));
                }
            }
            (Value::Object(left), Value::Object(right)) => {
                assert_eq!(left.len(), right.len(), "{name}: {path} keys");
                for (key, value) in left {
                    let other = right
                        .get(key)
                        .unwrap_or_else(|| panic!("{name}: {path}.{key} missing"));
                    assert_close(name, value, other, &format!("{path}.{key}"));
                }
            }
            (left, right) => assert_eq!(left, right, "{name}: {path}"),
        }
    }

    fn assert_parity(
        name: &str,
        block: &ParagraphBlock,
        content_width: f64,
        config: &MeasurementConfig,
        floating_zones: Option<&[FloatingZone]>,
        cumulative_y: f64,
    ) {
        let legacy = crate::measure_paragraph_json_resident(&legacy_envelope(
            block,
            content_width,
            config,
            floating_zones,
            cumulative_y,
        ))
        .ok()
        .and_then(|out| serde_json::from_str::<ParagraphExtent>(&out).ok());
        let typed = measure_paragraph(block, content_width, config, floating_zones, cumulative_y);
        match (legacy, typed) {
            (Some(legacy), Some(typed)) => {
                // The JSON path quantized every f32 through its shortest
                // decimal form before widening to f64; the typed path widens
                // exactly, so compare numerically with a tight tolerance.
                assert_close(
                    name,
                    &serde_json::to_value(&legacy).unwrap(),
                    &serde_json::to_value(&typed).unwrap(),
                    "$",
                );
            }
            (None, None) => {}
            (legacy, typed) => {
                panic!("{name}: fallback divergence — legacy {legacy:?}, typed {typed:?}")
            }
        }
    }

    #[test]
    fn plain_wrapping_text_matches_the_json_path() {
        let fixture = fixture();
        let block = paragraph(
            vec![text_run("The quick brown fox jumps over the lazy dog")],
            None,
        );
        assert_parity("plain", &block, 200.0, &fixture.config, None, 0.0);
        let via_entry = measure_paragraph_entry(&block, 200.0, &fixture.config).unwrap();
        let via_typed = measure_paragraph(&block, 200.0, &fixture.config, None, 0.0).unwrap();
        assert_eq!(
            serde_json::to_value(&via_entry).unwrap(),
            serde_json::to_value(&via_typed).unwrap(),
            "the production entry point must agree bit-for-bit with the typed path"
        );
    }

    #[test]
    fn cjk_and_bidi_runs_match_the_json_path() {
        let fixture = fixture();
        let attrs = crate::types::ParagraphAttrs {
            bidi: Some(true),
            ..Default::default()
        };
        let mut rtl = RunFormatting::default();
        rtl.rtl = Some(true);
        let runs = vec![
            text_run("שלום"),
            Run::Text(TextRun {
                fmt: rtl,
                text: " mixed 世界".to_owned(),
                pm_start: None,
                pm_end: None,
                inline_sdt_widget: None,
            }),
        ];
        assert_parity(
            "cjk-bidi",
            &paragraph(runs, Some(attrs)),
            240.0,
            &fixture.config,
            None,
            0.0,
        );
    }

    #[test]
    fn tab_stops_match_the_json_path() {
        let fixture = fixture();
        let attrs = crate::types::ParagraphAttrs {
            tabs: Some(vec![crate::types::TabStop {
                val: "end".to_owned(),
                pos: 4000.0,
                leader: None,
            }]),
            ..Default::default()
        };
        let runs = vec![
            text_run("before"),
            Run::Tab(TabRun {
                fmt: RunFormatting::default(),
                pm_start: None,
                pm_end: None,
                width: None,
                leader_glyphs: None,
            }),
            text_run("after the stop"),
        ];
        assert_parity(
            "tabs",
            &paragraph(runs, Some(attrs)),
            400.0,
            &fixture.config,
            None,
            0.0,
        );
    }

    #[test]
    fn list_marker_matches_the_json_path() {
        let fixture = fixture();
        let attrs = crate::types::ParagraphAttrs {
            list_marker: Some("7.".to_owned()),
            list_marker_font_family: Some("Liberation Sans".to_owned()),
            list_marker_font_size: Some(12.0),
            indent: Some(crate::types::ParagraphIndent {
                left: Some(20.0),
                right: None,
                first_line: None,
                hanging: None,
            }),
            default_tab_stop_twips: Some(720.0),
            ..Default::default()
        };
        assert_parity(
            "list-marker",
            &paragraph(vec![text_run("numbered item text")], Some(attrs)),
            220.0,
            &fixture.config,
            None,
            0.0,
        );
    }

    #[test]
    fn field_run_matches_the_json_path() {
        let fixture = fixture();
        let runs = vec![Run::Field(FieldRun {
            fmt: RunFormatting::default(),
            field_type: "PAGE".to_owned(),
            raw_type: None,
            instruction: None,
            fallback: Some("42".to_owned()),
            pm_start: None,
            pm_end: None,
        })];
        assert_parity(
            "field",
            &paragraph(runs, None),
            150.0,
            &fixture.config,
            None,
            0.0,
        );
    }

    fn img_run(rotation_bounds: Value) -> Run {
        Run::Image(
            serde_json::from_value(json!({
                "src": "a.png",
                "width": 24.0,
                "height": 18.0,
                "displayMode": "inline",
                "rotationBounds": rotation_bounds,
            }))
            .unwrap(),
        )
    }

    #[test]
    fn inline_image_matches_the_json_path() {
        let fixture = fixture();
        let block = paragraph(
            vec![img_run(json!({ "width": 30.0, "height": 22.0 }))],
            None,
        );
        assert_parity("image", &block, 200.0, &fixture.config, None, 0.0);
    }

    /// serde's derived `Deserialize` also takes a struct from a positional
    /// sequence, so `[width, height]` bounds must measure identically.
    #[test]
    fn positional_sequence_rotation_bounds_match_the_json_path() {
        let fixture = fixture();
        for bounds in [json!([]), json!([30.0]), json!([30.0, 22.0])] {
            assert_parity(
                "image-seq-bounds",
                &paragraph(vec![img_run(bounds)], None),
                200.0,
                &fixture.config,
                None,
                0.0,
            );
        }
    }

    /// Absent and explicit-`null` bounds are not malformed: serde parses
    /// `null` into `None`, so both paths must keep measuring.
    #[test]
    fn absent_and_null_rotation_bounds_measure_on_both_paths() {
        let fixture = fixture();
        let bare = Run::Image(
            serde_json::from_value(json!({
                "src": "a.png",
                "width": 24.0,
                "height": 18.0,
                "displayMode": "inline",
            }))
            .unwrap(),
        );
        assert_parity(
            "image-no-bounds",
            &paragraph(vec![bare], None),
            200.0,
            &fixture.config,
            None,
            0.0,
        );
        let null = paragraph(vec![img_run(Value::Null)], None);
        assert_parity(
            "image-null-bounds",
            &null,
            200.0,
            &fixture.config,
            None,
            0.0,
        );
        assert!(measure_paragraph(&null, 200.0, &fixture.config, None, 0.0).is_some());
    }

    /// Values the old envelope's serde pass rejected — a non-object bounds
    /// value, or an object with a non-number dimension — must fall back on
    /// both paths; the typed conversion may not quietly drop them and measure.
    #[test]
    fn malformed_rotation_bounds_fall_back_on_both_paths() {
        let fixture = fixture();
        for bounds in [
            json!(1),
            json!("bad"),
            json!(true),
            json!({ "width": "bad" }),
            json!({ "width": true }),
            json!({ "height": [] }),
            json!(["x"]),
            json!([1, "x"]),
            json!([1, 2, 3]),
        ] {
            let block = paragraph(vec![img_run(bounds.clone())], None);
            let legacy = crate::measure_paragraph_json_resident(&legacy_envelope(
                &block,
                200.0,
                &fixture.config,
                None,
                0.0,
            ))
            .ok()
            .and_then(|out| serde_json::from_str::<ParagraphExtent>(&out).ok());
            let typed = measure_paragraph(&block, 200.0, &fixture.config, None, 0.0);
            assert!(legacy.is_none(), "{bounds}: legacy must fall back");
            assert!(typed.is_none(), "{bounds}: typed must fall back");
        }
    }

    #[test]
    fn floating_zones_match_the_json_path() {
        let fixture = fixture();
        let zones = [FloatingZone {
            left_margin: 60.0,
            right_margin: 10.0,
            top_y: -5.0,
            bottom_y: 40.0,
            full_width_block: false,
        }];
        assert_parity(
            "zones",
            &paragraph(
                vec![text_run(
                    "flowing around a floating image with several words",
                )],
                None,
            ),
            260.0,
            &fixture.config,
            Some(&zones),
            12.0,
        );
    }

    #[test]
    fn unsupported_run_falls_back_on_both_paths() {
        let fixture = fixture();
        let block = paragraph(vec![Run::Unsupported, text_run("after")], None);
        assert_parity("unsupported", &block, 200.0, &fixture.config, None, 0.0);
    }
}
