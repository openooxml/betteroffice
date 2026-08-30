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
use serde::Deserialize;
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
    let zones = match floating_zones {
        None => None,
        Some(zones) => Some(
            zones
                .iter()
                .map(zone_in)
                .collect::<Option<Vec<FloatZoneIn>>>()?,
        ),
    };
    let request = MeasureRequest {
        block: &block,
        max_width: content_width as f32,
        font_chains: FontChains::BTree(&config.font_chains),
        defaults: &defaults,
        compat,
        floating_zones: zones.as_deref(),
        paragraph_y_offset: floating_zones
            .is_some()
            .then_some(cumulative_y)
            .and_then(finite),
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
        attrs: match paragraph.attrs.as_ref() {
            None => None,
            Some(attrs) => Some(attrs_in(attrs)?),
        },
    })
}

/// A non-finite float serialized as `null`, which the JSON boundary read back
/// as absent; the typed path drops it the same way.
fn finite(value: f64) -> Option<f32> {
    value.is_finite().then_some(value as f32)
}

fn defaults_in(defaults: &Value) -> Option<DefaultsIn> {
    DefaultsIn::deserialize(defaults).ok()
}

fn compat_in(compat: &Value) -> Option<CompatIn> {
    if compat.is_null() {
        return Some(CompatIn::default());
    }
    CompatIn::deserialize(compat).ok()
}

fn zone_in(zone: &FloatingZone) -> Option<FloatZoneIn> {
    Some(FloatZoneIn {
        left_margin: finite(zone.left_margin)?,
        right_margin: finite(zone.right_margin)?,
        top_y: finite(zone.top_y)?,
        bottom_y: finite(zone.bottom_y)?,
        segments: None,
        full_width_block: zone.full_width_block,
    })
}

fn attrs_in(attrs: &ParagraphAttrs) -> Option<AttrsIn> {
    Some(AttrsIn {
        alignment: attrs.alignment.clone(),
        spacing: attrs.spacing.as_ref().map(|spacing| SpacingIn {
            before: spacing.before.and_then(finite),
            after: spacing.after.and_then(finite),
            line: spacing.line.and_then(finite),
            line_unit: spacing.line_unit.clone(),
            line_rule: spacing.line_rule.clone(),
        }),
        indent: attrs.indent.as_ref().map(|indent| IndentIn {
            left: indent.left.and_then(finite),
            right: indent.right.and_then(finite),
            first_line: indent.first_line.and_then(finite),
            hanging: indent.hanging.and_then(finite),
        }),
        tabs: match attrs.tabs.as_ref() {
            None => None,
            Some(stops) => Some(stops.iter().map(tab_stop_in).collect::<Option<Vec<_>>>()?),
        },
        bidi: attrs.bidi.unwrap_or(false),
        default_font_size: attrs.default_font_size.and_then(finite),
        default_font_family: attrs.default_font_family.clone(),
        suppress_empty_paragraph_height: attrs.suppress_empty_paragraph_height.unwrap_or(false),
        list_marker: attrs.list_marker.clone(),
        list_marker_hidden: attrs.list_marker_hidden.unwrap_or(false),
        list_marker_font_family: attrs.list_marker_font_family.clone(),
        list_marker_font_size: attrs.list_marker_font_size.and_then(finite),
        list_marker_suffix: attrs.list_marker_suffix.clone(),
        default_tab_stop_twips: attrs.default_tab_stop_twips.and_then(finite),
    })
}

fn tab_stop_in(stop: &TabStop) -> Option<TabStopIn> {
    Some(TabStopIn {
        val: stop.val.clone(),
        pos: finite(stop.pos)?,
    })
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
    out.font_size = fmt.font_size.and_then(finite);
    out.font_size_cs = fmt.font_size_cs.and_then(finite);
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
    out.letter_spacing = fmt.letter_spacing.and_then(finite);
    out.all_caps = fmt.all_caps.unwrap_or(false);
    out.small_caps = fmt.small_caps.unwrap_or(false);
    out.horizontal_scale = fmt.horizontal_scale.and_then(finite);
    out.kerning_min_pt = fmt.kerning_min_pt.and_then(finite);
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
    out.width = finite(image.width);
    out.height = finite(image.height);
    out.rotation_bounds = rotation_bounds_in(image.rotation_bounds.as_ref())?;
    out.dist_top = image.dist_top.and_then(finite);
    out.dist_bottom = image.dist_bottom.and_then(finite);
    out.wrap_type = image.wrap_type.clone();
    out.display_mode = image.display_mode.clone();
    out.position = image.position.as_ref().map(|_| serde::de::IgnoredAny);
    Some(out)
}

/// Defers to the same derived `Deserialize` the envelope parse used, so
/// object, positional-sequence and rejected forms all behave as before.
fn rotation_bounds_in(bounds: Option<&Value>) -> Option<Option<RotationBoundsIn>> {
    match bounds {
        None => Some(None),
        Some(value) => Option::<RotationBoundsIn>::deserialize(value).ok(),
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
            "the production entry point must delegate to the typed core"
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

    struct Rng(u64);

    impl Rng {
        fn next(&mut self) -> u64 {
            let mut x = self.0;
            x ^= x << 13;
            x ^= x >> 7;
            x ^= x << 17;
            self.0 = x;
            x
        }
        fn below(&mut self, n: u64) -> u64 {
            self.next() % n
        }
        fn chance(&mut self, n: u64) -> bool {
            self.below(n) == 0
        }
    }

    const FAMILIES: [&str; 4] = ["Liberation Sans", "Arial", "MS Mincho", ""];
    const TEXTS: [&str; 8] = [
        "The quick brown fox jumps over the lazy dog",
        "short",
        "",
        " ",
        "\t\t",
        "\u{5e9}\u{5dc}\u{5d5}\u{5dd} mixed \u{627}\u{644}\u{639}",
        "\u{65e5}\u{672c}\u{8a9e}\u{306e}\u{30c6}\u{30ad}\u{30b9}\u{30c8}",
        "ligature ffi fi fl and combining e\u{301}",
    ];

    fn num(rng: &mut Rng) -> f64 {
        match rng.below(7) {
            0 => 0.0,
            1 => 1.0,
            2 => 12.0,
            3 => 720.0,
            4 => (rng.below(4000) as f64) / 16.0,
            5 => -((rng.below(400) as f64) / 8.0),
            _ => (rng.below(100_000) as f64) / 1000.0,
        }
    }

    fn maybe_num(rng: &mut Rng) -> Option<Value> {
        (!rng.chance(3)).then(|| json!(num(rng)))
    }

    fn maybe_bool(rng: &mut Rng) -> Option<Value> {
        (!rng.chance(3)).then(|| json!(rng.chance(2)))
    }

    fn maybe_str(rng: &mut Rng, pool: &[&str]) -> Option<Value> {
        (!rng.chance(3)).then(|| json!(pool[rng.below(pool.len() as u64) as usize]))
    }

    fn put(map: &mut serde_json::Map<String, Value>, key: &str, value: Option<Value>) {
        if let Some(value) = value {
            map.insert(key.to_owned(), value);
        }
    }

    /// `RunFormatting` is `#[serde(flatten)]`, so its fields sit at run level;
    /// nesting them under `fmt` would silently drop every one.
    fn formatting(rng: &mut Rng, target: &mut serde_json::Map<String, Value>) {
        put(target, "bold", maybe_bool(rng));
        put(target, "italic", maybe_bool(rng));
        put(target, "boldCs", maybe_bool(rng));
        put(target, "italicCs", maybe_bool(rng));
        put(target, "fontSize", maybe_num(rng));
        put(target, "fontSizeCs", maybe_num(rng));
        put(target, "fontFamily", maybe_str(rng, &FAMILIES));
        put(target, "complexScript", maybe_bool(rng));
        put(target, "letterSpacing", maybe_num(rng));
        put(target, "allCaps", maybe_bool(rng));
        put(target, "smallCaps", maybe_bool(rng));
        put(target, "horizontalScale", maybe_num(rng));
        put(target, "kerningMinPt", maybe_num(rng));
        put(target, "superscript", maybe_bool(rng));
        put(target, "subscript", maybe_bool(rng));
        put(target, "hidden", maybe_bool(rng));
        put(target, "rtl", maybe_bool(rng));
        if !rng.chance(3) {
            let mut slots = serde_json::Map::new();
            put(&mut slots, "ascii", maybe_str(rng, &FAMILIES));
            put(&mut slots, "hAnsi", maybe_str(rng, &FAMILIES));
            put(&mut slots, "eastAsia", maybe_str(rng, &FAMILIES));
            put(&mut slots, "cs", maybe_str(rng, &FAMILIES));
            put(&mut slots, "hint", maybe_str(rng, &["eastAsia", "default"]));
            target.insert("fontSlots".to_owned(), Value::Object(slots));
        }
        if !rng.chance(4) {
            let mut lang = serde_json::Map::new();
            put(&mut lang, "latin", maybe_str(rng, &["en-US"]));
            put(&mut lang, "eastAsia", maybe_str(rng, &["ja-JP"]));
            put(&mut lang, "bidi", maybe_str(rng, &["he-IL"]));
            target.insert("language".to_owned(), Value::Object(lang));
        }
    }

    fn random_run(rng: &mut Rng) -> Value {
        let mut m = serde_json::Map::new();
        match rng.below(7) {
            0..=2 => {
                m.insert("kind".to_owned(), json!("text"));
                let text = TEXTS[rng.below(TEXTS.len() as u64) as usize];
                m.insert("text".to_owned(), json!(text));
                formatting(rng, &mut m);
            }
            3 => {
                m.insert("kind".to_owned(), json!("tab"));
                formatting(rng, &mut m);
            }
            4 => {
                m.insert("kind".to_owned(), json!("lineBreak"));
            }
            5 => {
                m.insert("kind".to_owned(), json!("field"));
                m.insert("fieldType".to_owned(), json!("PAGE"));
                put(&mut m, "fallback", maybe_str(rng, &["42", "", "iv"]));
                formatting(rng, &mut m);
            }
            _ => {
                m.insert("kind".to_owned(), json!("image"));
                m.insert("src".to_owned(), json!("a.png"));
                m.insert("width".to_owned(), json!(num(rng)));
                m.insert("height".to_owned(), json!(num(rng)));
                put(&mut m, "distTop", maybe_num(rng));
                put(&mut m, "distBottom", maybe_num(rng));
                put(&mut m, "wrapType", maybe_str(rng, &["inline", "square"]));
                put(&mut m, "displayMode", maybe_str(rng, &["inline", "block"]));
                if !rng.chance(3) {
                    m.insert("position".to_owned(), json!({ "behindDoc": false }));
                }
                if !rng.chance(3) {
                    let bounds = match rng.below(5) {
                        0 => json!({ "width": num(rng), "height": num(rng) }),
                        1 => json!([num(rng), num(rng)]),
                        2 => json!([]),
                        3 => Value::Null,
                        _ => json!({ "width": num(rng) }),
                    };
                    m.insert("rotationBounds".to_owned(), bounds);
                }
            }
        }
        Value::Object(m)
    }

    fn random_attrs(rng: &mut Rng) -> Value {
        let mut m = serde_json::Map::new();
        put(
            &mut m,
            "alignment",
            maybe_str(rng, &["left", "center", "both"]),
        );
        if !rng.chance(3) {
            let mut sp = serde_json::Map::new();
            put(&mut sp, "before", maybe_num(rng));
            put(&mut sp, "after", maybe_num(rng));
            put(&mut sp, "line", maybe_num(rng));
            put(&mut sp, "lineUnit", maybe_str(rng, &["px", "multiple"]));
            put(
                &mut sp,
                "lineRule",
                maybe_str(rng, &["auto", "exact", "atLeast"]),
            );
            m.insert("spacing".to_owned(), Value::Object(sp));
        }
        if !rng.chance(3) {
            let mut ind = serde_json::Map::new();
            put(&mut ind, "left", maybe_num(rng));
            put(&mut ind, "right", maybe_num(rng));
            put(&mut ind, "firstLine", maybe_num(rng));
            put(&mut ind, "hanging", maybe_num(rng));
            m.insert("indent".to_owned(), Value::Object(ind));
        }
        if !rng.chance(3) {
            let stops: Vec<Value> = (0..rng.below(4))
                .map(|_| {
                    let val = ["start", "end", "center", "decimal"][rng.below(4) as usize];
                    let pos = num(rng);
                    json!({ "val": val, "pos": pos })
                })
                .collect();
            m.insert("tabs".to_owned(), json!(stops));
        }
        put(&mut m, "bidi", maybe_bool(rng));
        put(&mut m, "defaultFontSize", maybe_num(rng));
        put(&mut m, "defaultFontFamily", maybe_str(rng, &FAMILIES));
        put(&mut m, "suppressEmptyParagraphHeight", maybe_bool(rng));
        put(
            &mut m,
            "listMarker",
            maybe_str(rng, &["1.", "\u{2022}", ""]),
        );
        put(&mut m, "listMarkerHidden", maybe_bool(rng));
        put(&mut m, "listMarkerFontFamily", maybe_str(rng, &FAMILIES));
        put(&mut m, "listMarkerFontSize", maybe_num(rng));
        put(
            &mut m,
            "listMarkerSuffix",
            maybe_str(rng, &["tab", "space"]),
        );
        put(&mut m, "defaultTabStopTwips", maybe_num(rng));
        Value::Object(m)
    }

    fn random_paragraph(rng: &mut Rng) -> Value {
        let runs: Vec<Value> = (0..1 + rng.below(4)).map(|_| random_run(rng)).collect();
        let mut m = serde_json::Map::new();
        m.insert("id".to_owned(), json!(0.0));
        m.insert("runs".to_owned(), json!(runs));
        if !rng.chance(4) {
            m.insert("attrs".to_owned(), random_attrs(rng));
        }
        Value::Object(m)
    }

    /// The `BlockIn` the legacy envelope actually delivered to the engine.
    fn legacy_block_in(paragraph: &ParagraphBlock) -> Option<ooxml_text::measure::BlockIn> {
        let envelope = json!({ "block": LayoutBlock::Paragraph(paragraph.clone()) }).to_string();
        let parsed: Value = serde_json::from_str(&envelope).ok()?;
        serde_json::from_value(parsed.get("block")?.clone()).ok()
    }

    /// Re-narrows to f32 so the JSON path's f32 -> shortest-decimal -> f64
    /// widening cannot mask a real difference.
    fn narrow(value: &Value) -> Value {
        match value {
            Value::Number(n) => json!(format!("{:?}", n.as_f64().unwrap_or(f64::NAN) as f32)),
            Value::Array(items) => Value::Array(items.iter().map(narrow).collect()),
            Value::Object(map) => {
                Value::Object(map.iter().map(|(k, v)| (k.clone(), narrow(v))).collect())
            }
            other => other.clone(),
        }
    }

    fn legacy_measure(
        paragraph: &ParagraphBlock,
        content_width: f64,
        config: &MeasurementConfig,
        zones: Option<&[FloatingZone]>,
        cumulative_y: f64,
    ) -> Option<ParagraphExtent> {
        let envelope = legacy_envelope(paragraph, content_width, config, zones, cumulative_y);
        let extent = crate::measure_paragraph_json_resident(&envelope).ok()?;
        serde_json::from_str::<ParagraphExtent>(&extent).ok()
    }

    /// Randomised whole-surface differential. Every mapped field varies, so a
    /// dropped or mistyped mapping shows up as a `BlockIn` difference before it
    /// can reach measurement.
    #[test]
    fn randomised_paragraphs_match_the_json_path() {
        let fixture = fixture();
        let mut rng = Rng(0x2f6e_2b13);
        let mut measured = 0usize;
        let mut fell_back = 0usize;
        for case in 0..3000 {
            let spec = random_paragraph(&mut rng);
            let paragraph: ParagraphBlock =
                serde_json::from_value(spec.clone()).expect("generated paragraph parses");
            let width = [200.0_f64, 60.0, 1000.0, 13.5][rng.below(4) as usize];
            let size = [8.0_f64, 11.0, 12.0, 24.0][rng.below(4) as usize];
            let family = FAMILIES[rng.below(FAMILIES.len() as u64) as usize];
            let no_leading = rng.chance(2);
            let no_expand = rng.chance(2);
            let config = MeasurementConfig {
                font_chains: fixture.config.font_chains.clone(),
                defaults: json!({ "fontSize": size, "fontFamily": family }),
                compat: json!({
                    "noLeading": no_leading,
                    "doNotExpandShiftReturn": no_expand,
                }),
                authoritative_shaping: rng.chance(2),
            };
            let full_width = rng.chance(2);
            let zones = rng.chance(4).then(|| {
                vec![FloatingZone {
                    left_margin: 40.0,
                    right_margin: 8.0,
                    top_y: -4.0,
                    bottom_y: 60.0,
                    full_width_block: full_width,
                }]
            });

            assert_eq!(
                legacy_block_in(&paragraph).map(|b| format!("{b:?}")),
                block_in(&paragraph).map(|b| format!("{b:?}")),
                "case {case}: BlockIn differs for {spec}"
            );

            let legacy = legacy_measure(&paragraph, width, &config, zones.as_deref(), 12.0);
            let typed = measure_paragraph(&paragraph, width, &config, zones.as_deref(), 12.0);
            match (&legacy, &typed) {
                (None, None) => fell_back += 1,
                (Some(a), Some(b)) => {
                    measured += 1;
                    assert_eq!(
                        narrow(&serde_json::to_value(a).unwrap()),
                        narrow(&serde_json::to_value(b).unwrap()),
                        "case {case}: extent differs for {spec}"
                    );
                }
                _ => panic!(
                    "case {case}: fallback differs (legacy {}, typed {}) for {spec}",
                    legacy.is_some(),
                    typed.is_some()
                ),
            }
        }
        assert!(measured > 500, "corpus measured only {measured} paragraphs");
        assert!(
            fell_back > 100,
            "corpus exercised only {fell_back} fallbacks"
        );
    }

    /// `serde_json` writes a non-finite float as `null`, which the envelope read
    /// back as absent. Every site that crosses the boundary must still do that.
    #[test]
    fn non_finite_numbers_match_the_json_path() {
        use crate::types::{ParagraphIndent, ParagraphSpacing};
        let fixture = fixture();
        for site in 0..24 {
            for value in [f64::NAN, f64::INFINITY, f64::NEG_INFINITY] {
                let mut fmt = RunFormatting::default();
                let mut attrs = crate::types::ParagraphAttrs::default();
                let mut spacing = ParagraphSpacing {
                    before: Some(2.0),
                    after: Some(3.0),
                    line: Some(1.0),
                    line_unit: Some("multiple".to_owned()),
                    line_rule: Some("auto".to_owned()),
                };
                let mut indent = ParagraphIndent {
                    left: Some(4.0),
                    right: Some(2.0),
                    first_line: Some(6.0),
                    hanging: Some(0.0),
                };
                let mut tab = crate::types::TabStop {
                    val: "end".to_owned(),
                    pos: 2000.0,
                    leader: None,
                };
                let mut image: crate::types::ImageRun = serde_json::from_value(json!({
                    "src": "a.png", "width": 20.0, "height": 14.0,
                    "displayMode": "inline", "distTop": 1.0, "distBottom": 1.0,
                }))
                .unwrap();
                let mut cumulative_y = 10.0_f64;
                let mut zone = [40.0_f64, 8.0, -4.0, 60.0];
                let mut width = 200.0_f64;
                let mut with_image = false;
                let mut with_zones = false;

                let name = match site {
                    0 => {
                        fmt.font_size = Some(value);
                        "run.fontSize"
                    }
                    1 => {
                        fmt.font_size_cs = Some(value);
                        "run.fontSizeCs"
                    }
                    2 => {
                        fmt.letter_spacing = Some(value);
                        "run.letterSpacing"
                    }
                    3 => {
                        fmt.horizontal_scale = Some(value);
                        "run.horizontalScale"
                    }
                    4 => {
                        fmt.kerning_min_pt = Some(value);
                        "run.kerningMinPt"
                    }
                    5 => {
                        attrs.default_font_size = Some(value);
                        "attrs.defaultFontSize"
                    }
                    6 => {
                        attrs.list_marker = Some("1.".to_owned());
                        attrs.list_marker_font_size = Some(value);
                        "attrs.listMarkerFontSize"
                    }
                    7 => {
                        attrs.default_tab_stop_twips = Some(value);
                        "attrs.defaultTabStopTwips"
                    }
                    8 => {
                        spacing.before = Some(value);
                        attrs.spacing = Some(spacing);
                        "spacing.before"
                    }
                    9 => {
                        spacing.after = Some(value);
                        attrs.spacing = Some(spacing);
                        "spacing.after"
                    }
                    10 => {
                        spacing.line = Some(value);
                        attrs.spacing = Some(spacing);
                        "spacing.line"
                    }
                    11 => {
                        indent.left = Some(value);
                        attrs.indent = Some(indent);
                        "indent.left"
                    }
                    12 => {
                        indent.right = Some(value);
                        attrs.indent = Some(indent);
                        "indent.right"
                    }
                    13 => {
                        indent.first_line = Some(value);
                        attrs.indent = Some(indent);
                        "indent.firstLine"
                    }
                    14 => {
                        indent.hanging = Some(value);
                        attrs.indent = Some(indent);
                        "indent.hanging"
                    }
                    15 => {
                        tab.pos = value;
                        attrs.tabs = Some(vec![tab]);
                        "tab.pos"
                    }
                    16 => {
                        image.width = value;
                        with_image = true;
                        "image.width"
                    }
                    17 => {
                        image.height = value;
                        with_image = true;
                        "image.height"
                    }
                    18 => {
                        image.dist_top = Some(value);
                        with_image = true;
                        "image.distTop"
                    }
                    19 => {
                        image.dist_bottom = Some(value);
                        with_image = true;
                        "image.distBottom"
                    }
                    20 => {
                        cumulative_y = value;
                        with_zones = true;
                        "paragraphYOffset"
                    }
                    21 => {
                        zone[0] = value;
                        with_zones = true;
                        "zone.leftMargin"
                    }
                    22 => {
                        zone[2] = value;
                        with_zones = true;
                        "zone.topY"
                    }
                    _ => {
                        width = value;
                        "maxWidth"
                    }
                };

                let mut runs = vec![Run::Text(TextRun {
                    fmt,
                    text: "hello world wrapping across a few lines".to_owned(),
                    pm_start: None,
                    pm_end: None,
                    inline_sdt_widget: None,
                })];
                if with_image {
                    runs.push(Run::Image(image));
                }
                let block = paragraph(runs, Some(attrs));
                let zones = with_zones.then(|| {
                    vec![FloatingZone {
                        left_margin: zone[0],
                        right_margin: zone[1],
                        top_y: zone[2],
                        bottom_y: zone[3],
                        full_width_block: false,
                    }]
                });
                assert_eq!(
                    legacy_block_in(&block).map(|b| format!("{b:?}")),
                    block_in(&block).map(|b| format!("{b:?}")),
                    "{name} = {value}: BlockIn differs"
                );
                let legacy = legacy_measure(
                    &block,
                    width,
                    &fixture.config,
                    zones.as_deref(),
                    cumulative_y,
                );
                let typed = measure_paragraph(
                    &block,
                    width,
                    &fixture.config,
                    zones.as_deref(),
                    cumulative_y,
                );
                match (&legacy, &typed) {
                    (None, None) => {}
                    (Some(a), Some(b)) => assert_eq!(
                        narrow(&serde_json::to_value(a).unwrap()),
                        narrow(&serde_json::to_value(b).unwrap()),
                        "{name} = {value}: extent differs"
                    ),
                    _ => panic!(
                        "{name} = {value}: fallback differs (legacy {}, typed {})",
                        legacy.is_some(),
                        typed.is_some()
                    ),
                }
            }
        }
    }

    /// serde's derived `Deserialize` also takes a struct from a positional
    /// sequence, so the envelope accepted these and the typed path must too.
    #[test]
    fn sequence_form_defaults_and_compat_match_the_json_path() {
        let base = fixture();
        let block = paragraph(vec![text_run("hello world")], None);
        for (name, defaults, compat) in [
            ("defaults-seq", json!([12.0, "Liberation Sans"]), json!({})),
            (
                "compat-seq",
                json!({ "fontSize": 12.0, "fontFamily": "Liberation Sans" }),
                json!([true, false]),
            ),
            (
                "compat-seq-one",
                json!({ "fontSize": 12.0, "fontFamily": "Liberation Sans" }),
                json!([true]),
            ),
            (
                "compat-seq-empty",
                json!({ "fontSize": 12.0, "fontFamily": "Liberation Sans" }),
                json!([]),
            ),
            ("defaults-null", Value::Null, json!({})),
            (
                "compat-null",
                json!({ "fontSize": 12.0, "fontFamily": "Liberation Sans" }),
                Value::Null,
            ),
            ("defaults-scalar", json!(3), json!({})),
            (
                "compat-scalar",
                json!({ "fontSize": 12.0, "fontFamily": "Liberation Sans" }),
                json!(3),
            ),
            (
                "defaults-missing-key",
                json!({ "fontSize": 12.0 }),
                json!({}),
            ),
            (
                "defaults-wrong-type",
                json!({ "fontSize": "big", "fontFamily": "Liberation Sans" }),
                json!({}),
            ),
            (
                "compat-wrong-type",
                json!({ "fontSize": 12.0, "fontFamily": "Liberation Sans" }),
                json!({ "noLeading": "yes" }),
            ),
        ] {
            let config = MeasurementConfig {
                font_chains: base.config.font_chains.clone(),
                defaults,
                compat,
                authoritative_shaping: true,
            };
            let legacy = legacy_measure(&block, 200.0, &config, None, 0.0);
            let typed = measure_paragraph(&block, 200.0, &config, None, 0.0);
            assert_eq!(
                legacy.is_some(),
                typed.is_some(),
                "{name}: fallback differs (legacy {}, typed {})",
                legacy.is_some(),
                typed.is_some()
            );
        }
    }
}
