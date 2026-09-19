//! Integration tests against Liberation Sans Regular 2.1.5 (SIL OFL 1.1).

use ooxml_text::{
    BaseDirection, BreakOpportunity, CompatFlags, FontMetrics, FontStore, LineBox, LineSpacingRule,
    ShapeDirection, ShapeFeature, apply_spacing_rule, bidi_paragraphs, break_opportunities,
    kern_enabled, kern_features, line_is_justified, shape, shape_with_direction, single_line_box,
    snap_line_box, snap_line_height, stretch_spaces,
};

const LIBERATION_SANS: &[u8] = include_bytes!("fonts/LiberationSans-Regular.ttf");
const NOTO_NASKH_ARABIC: &[u8] =
    include_bytes!("../../../packages/fonts/assets/NotoNaskhArabic-Regular.ttf");

fn store_with_font() -> (FontStore, ooxml_text::FontId) {
    let mut store = FontStore::new();
    let id = store
        .register(LIBERATION_SANS.to_vec())
        .expect("fixture font registers");
    (store, id)
}

fn store_with_arabic_font() -> (FontStore, ooxml_text::FontId) {
    let mut store = FontStore::new();
    let id = store
        .register(NOTO_NASKH_ARABIC.to_vec())
        .expect("Arabic fixture font registers");
    (store, id)
}

// hand-computed from LiberationSans-Regular.ttf head/hhea/OS/2 tables
#[test]
fn metrics_match_hand_computed_table_values() {
    let (store, id) = store_with_font();
    let m = store.metrics(id).unwrap();
    assert_eq!(m.units_per_em, 2048);
    assert_eq!(m.hhea_ascender, 1854);
    assert_eq!(m.hhea_descender, -434);
    assert_eq!(m.hhea_line_gap, 67);
    assert_eq!(m.os2_typo_ascender, 1491);
    assert_eq!(m.os2_typo_descender, -431);
    assert_eq!(m.os2_typo_line_gap, 307);
    assert_eq!(m.os2_win_ascent, 1854);
    assert_eq!(m.os2_win_descent, 434);
    assert_eq!(m.os2_fs_selection, 0x40);
    assert_eq!(m.os2_version, 3);
}

// hand-computed from cmap format-4 + hmtx: (char, glyph id, advance in font units)
#[test]
fn advance_widths_match_hand_computed_hmtx_values() {
    let (store, id) = store_with_font();
    let expected = [
        ('A', 36u16, 1366.0f32),
        ('V', 57, 1366.0),
        ('W', 58, 1933.0),
        (' ', 3, 569.0),
        ('i', 76, 455.0),
        ('x', 91, 1024.0),
        ('0', 19, 1139.0),
    ];
    for (ch, gid, advance) in expected {
        assert_eq!(store.glyph_id(id, ch).unwrap(), Some(gid), "gid of {ch:?}");
        assert_eq!(
            store.advance_width(id, ch).unwrap(),
            Some(advance),
            "advance of {ch:?}"
        );
    }
}

#[test]
fn uncovered_char_reports_no_glyph() {
    let (store, id) = store_with_font();
    // Liberation Sans has no CJK coverage
    assert_eq!(store.glyph_id(id, '\u{4E2D}').unwrap(), None);
    assert_eq!(store.advance_width(id, '\u{4E2D}').unwrap(), None);
    assert!(!store.covers(id, '\u{4E2D}').unwrap());
}

#[test]
fn rejects_garbage_bytes() {
    let mut store = FontStore::new();
    assert!(store.register(b"definitely not a font".to_vec()).is_err());
}

// shaping at size == upem makes shaped advances directly comparable to
// font-unit hmtx advances
#[test]
fn shaping_kerned_pair_differs_from_sum_of_advances() {
    let (store, id) = store_with_font();
    let upem = store.metrics(id).unwrap().units_per_em as f32;

    let glyphs = shape(&store, id, "AV", upem, &[]).unwrap();
    assert_eq!(glyphs.len(), 2);
    assert_eq!(glyphs[0].glyph_id, 36); // A
    assert_eq!(glyphs[1].glyph_id, 57); // V
    assert_eq!(glyphs[0].cluster, 0);
    assert_eq!(glyphs[1].cluster, 1);

    let shaped_total: f32 = glyphs.iter().map(|g| g.x_advance).sum();
    let sum_of_advances = 1366.0 + 1366.0;
    assert!(
        shaped_total < sum_of_advances,
        "GPOS kerning must tighten AV: shaped {shaped_total} vs plain {sum_of_advances}"
    );
}

#[test]
fn shaping_scales_advances_to_size() {
    let (store, id) = store_with_font();
    // 'A' at 16px: 1366 * 16 / 2048 = 10.671875
    let glyphs = shape(&store, id, "A", 16.0, &[]).unwrap();
    assert_eq!(glyphs.len(), 1);
    assert!((glyphs[0].x_advance - 1366.0 * 16.0 / 2048.0).abs() < 1e-4);
}

#[test]
fn disabling_kern_feature_restores_plain_advances() {
    let (store, id) = store_with_font();
    let upem = store.metrics(id).unwrap().units_per_em as f32;
    let no_kern = [ShapeFeature {
        tag: *b"kern",
        value: 0,
    }];
    let glyphs = shape(&store, id, "AV", upem, &no_kern).unwrap();
    let total: f32 = glyphs.iter().map(|g| g.x_advance).sum();
    assert_eq!(total, 1366.0 + 1366.0);
}

#[test]
fn explicit_rtl_shape_direction_outputs_visual_clusters() {
    let (store, id) = store_with_font();
    let upem = store.metrics(id).unwrap().units_per_em as f32;

    let glyphs = shape_with_direction(&store, id, "אבג", upem, &[], ShapeDirection::Rtl).unwrap();
    assert_eq!(glyphs.len(), 3);
    assert!(
        glyphs.first().unwrap().cluster > glyphs.last().unwrap().cluster,
        "RTL shaping should return visual-order glyphs with descending source clusters: {glyphs:?}"
    );
}

#[test]
fn arabic_rtl_shaping_applies_joining_substitutions() {
    let (store, id) = store_with_arabic_font();
    let size = 48.0;
    let text = "سلام";

    let joined = shape_with_direction(&store, id, text, size, &[], ShapeDirection::Rtl).unwrap();
    assert!(!joined.is_empty(), "Arabic word shaped to glyphs");

    let isolated_visual_ids: Vec<u32> = text
        .chars()
        .rev()
        .map(|ch| {
            let s = ch.to_string();
            shape_with_direction(&store, id, &s, size, &[], ShapeDirection::Rtl)
                .unwrap()
                .first()
                .unwrap()
                .glyph_id
        })
        .collect();
    let joined_ids: Vec<u32> = joined.iter().map(|g| g.glyph_id).collect();

    assert_ne!(
        joined_ids, isolated_visual_ids,
        "Arabic word should not shape as isolated glyph forms"
    );
    assert!(
        joined.first().unwrap().cluster > joined.last().unwrap().cluster,
        "RTL Arabic glyphs should be in visual order: {joined:?}"
    );
}

#[test]
fn break_opportunities_for_mixed_latin_cjk_space_text() {
    // "foo bar " (8 bytes ascii) + 漢字 (3 bytes each)
    let text = "foo bar 漢字";
    let breaks = break_opportunities(text);

    // after each space the next word may start (UAX-14: break index = start
    // of the would-be next line)
    let allowed: Vec<usize> = breaks
        .iter()
        .filter(|b| !b.mandatory)
        .map(|b| b.byte_index)
        .collect();
    assert!(allowed.contains(&4), "break before 'bar': {allowed:?}");
    assert!(allowed.contains(&8), "break before 漢: {allowed:?}");
    // Ideographs break between every pair.
    assert!(
        allowed.contains(&11),
        "break between 漢 and 字: {allowed:?}"
    );

    // end of text is the only mandatory break, at a char-safe byte index
    let mandatory: Vec<usize> = breaks
        .iter()
        .filter(|b| b.mandatory)
        .map(|b| b.byte_index)
        .collect();
    assert_eq!(mandatory, vec![text.len()]);
    assert!(text.is_char_boundary(text.len()));
}

#[test]
fn newline_is_a_mandatory_break() {
    let breaks = break_opportunities("a\nb");
    assert_eq!(
        breaks,
        vec![
            BreakOpportunity {
                byte_index: 2,
                mandatory: true
            },
            BreakOpportunity {
                byte_index: 3,
                mandatory: true
            },
        ]
    );
}

// Every opportunity is a character boundary through supplementary-plane text.
#[test]
fn break_opportunities_are_char_boundaries_in_supplementary_plane_text() {
    let text = "a\u{1F600}\u{20BB7}b \u{1F3B4}c";
    for b in break_opportunities(text) {
        assert!(
            text.is_char_boundary(b.byte_index),
            "not a char boundary: {}",
            b.byte_index
        );
    }
}

#[test]
fn fallback_resolution_picks_first_covering_font_in_chain_order() {
    let mut store = FontStore::new();
    let first = store.register(LIBERATION_SANS.to_vec()).unwrap();
    let second = store.register(LIBERATION_SANS.to_vec()).unwrap();
    assert_ne!(first, second);

    // first covering font wins, in chain order
    assert_eq!(store.resolve(&[first, second], 'A'), Some(first));
    assert_eq!(store.resolve(&[second, first], 'A'), Some(second));
    // no font in the chain covers CJK -> None (host degrades that run)
    assert_eq!(store.resolve(&[first, second], '\u{4E2D}'), None);
    // empty chain resolves nothing
    assert_eq!(store.resolve(&[], 'A'), None);
}

#[test]
fn bidi_splits_mixed_ltr_rtl_into_level_runs() {
    // "abc " (4 bytes) + אבג (2 bytes each = 6)
    let text = "abc אבג";
    let paras = bidi_paragraphs(text, BaseDirection::Auto);
    assert_eq!(paras.len(), 1);
    let para = &paras[0];
    assert_eq!(para.base_level % 2, 0, "first strong char is LTR");

    assert_eq!(para.runs.len(), 2);
    assert_eq!((para.runs[0].start, para.runs[0].end), (0, 4));
    assert!(!para.runs[0].is_rtl());
    assert_eq!((para.runs[1].start, para.runs[1].end), (4, 10));
    assert!(para.runs[1].is_rtl());
}

#[test]
fn bidi_forced_rtl_base_direction() {
    let paras = bidi_paragraphs("abc", BaseDirection::Rtl);
    assert_eq!(paras.len(), 1);
    assert_eq!(paras[0].base_level, 1);
    // latin text stays an LTR run inside the RTL paragraph
    assert_eq!(paras[0].runs.len(), 1);
    assert!(!paras[0].runs[0].is_rtl());
}

// ---- word_metrics -------------------------------------------------------
//
// All expected values below are hand-computed from the fixture's raw tables:
// upem 2048, usWinAscent 1854, usWinDescent 434, hhea ascender 1854,
// descender -434, lineGap 67. At 16px the scale is 16/2048 = 0.0078125, an
// exact binary fraction, so every expected value is exact in f32 and
// assert_eq! is legitimate.

/// Fixture single-spacing box at 16px, computed independently of the crate.
fn liberation_single_16px() -> LineBox {
    LineBox {
        ascent: 15.0078125, // (1854 + 67) * 16 / 2048
        descent: 3.390625,  // 434 * 16 / 2048
        leading: 0.0,
    }
}

#[test]
fn single_line_box_uses_hhea_metrics_with_the_gap_above_the_ascender() {
    let (store, id) = store_with_font();
    let m = store.metrics(id).unwrap();

    let line = single_line_box(m, 16.0, &CompatFlags::default());
    assert_eq!(line, liberation_single_16px());
    // full pitch = 18.3984375, exact in f32
    assert_eq!(line.height(), (1854.0 + 434.0 + 67.0) * 16.0 / 2048.0);
}

#[test]
fn the_win_box_does_not_move_the_single_spacing_line() {
    let (store, id) = store_with_font();
    let inflated = FontMetrics {
        os2_win_ascent: 4000,
        os2_win_descent: 1200,
        ..*store.metrics(id).unwrap()
    };
    let shrunk = FontMetrics {
        os2_win_ascent: 1000,
        os2_win_descent: 200,
        ..*store.metrics(id).unwrap()
    };

    assert_eq!(
        single_line_box(&inflated, 16.0, &CompatFlags::default()),
        liberation_single_16px()
    );
    assert_eq!(
        single_line_box(&shrunk, 16.0, &CompatFlags::default()),
        liberation_single_16px()
    );
}

#[test]
fn aptos_measures_its_hhea_span_not_its_taller_win_box() {
    // Word 16.113's own Aptos: upem 2048, usWin 2068/563, hhea 1923/-577/0.
    let aptos = FontMetrics {
        units_per_em: 2048,
        hhea_ascender: 1923,
        hhea_descender: -577,
        hhea_line_gap: 0,
        os2_win_ascent: 2068,
        os2_win_descent: 563,
        os2_typo_ascender: 1923,
        os2_typo_descender: -577,
        os2_typo_line_gap: 0,
        os2_fs_selection: USE_TYPO_METRICS,
        os2_version: 4,
        os2_code_page_range1: 0,
    };

    let line = single_line_box(&aptos, 2048.0, &CompatFlags::default());
    assert_eq!(line.height(), 2500.0);
    assert_eq!(line.ascent, 1923.0);
}

#[test]
fn no_leading_compat_flag_drops_the_line_gap_only() {
    let (store, id) = store_with_font();
    let m = store.metrics(id).unwrap();

    let compat = CompatFlags {
        no_leading: true,
        ..CompatFlags::default()
    };
    let line = single_line_box(m, 16.0, &compat);
    assert_eq!(
        line,
        LineBox {
            ascent: 14.484375, // 1854 * 16 / 2048
            descent: 3.390625,
            leading: 0.0,
        }
    );
}

const USE_TYPO_METRICS: u16 = 0x0080;

fn synthetic_metrics() -> FontMetrics {
    FontMetrics {
        units_per_em: 1000,
        hhea_ascender: 620,
        hhea_descender: -170,
        hhea_line_gap: 70,
        os2_typo_ascender: 555,
        os2_typo_descender: -155,
        os2_typo_line_gap: 65,
        os2_win_ascent: 620,
        os2_win_descent: 170,
        os2_fs_selection: 0,
        os2_version: 4,
        os2_code_page_range1: 0,
    }
}

fn typo_metrics() -> FontMetrics {
    FontMetrics {
        os2_fs_selection: USE_TYPO_METRICS,
        ..synthetic_metrics()
    }
}

fn gdi_flags() -> CompatFlags {
    CompatFlags {
        gdi_line_metrics: true,
        ..CompatFlags::default()
    }
}

fn typo_flags() -> CompatFlags {
    CompatFlags {
        typo_line_spacing: true,
        ..CompatFlags::default()
    }
}

fn gdi_typo_flags() -> CompatFlags {
    CompatFlags {
        gdi_line_metrics: true,
        typo_line_spacing: true,
        ..CompatFlags::default()
    }
}

#[test]
fn default_path_does_not_clamp_spec_valid_vertical_metrics() {
    let metrics = FontMetrics {
        units_per_em: 1024,
        hhea_ascender: 20_480,
        hhea_descender: 0,
        hhea_line_gap: 0,
        os2_win_ascent: 20_480,
        os2_win_descent: 0,
        ..synthetic_metrics()
    };

    assert_eq!(
        single_line_box(&metrics, 16.0, &CompatFlags::default()),
        LineBox {
            ascent: 320.0,
            descent: 0.0,
            leading: 0.0,
        }
    );
    assert_eq!(single_line_box(&metrics, 16.0, &gdi_flags()).ascent, 256.0);
}

#[test]
fn the_two_line_metric_flags_are_independent() {
    let m = typo_metrics();
    let float_win = single_line_box(&m, 20.0, &CompatFlags::default());
    let float_typo = single_line_box(&m, 20.0, &typo_flags());
    let gdi_win = single_line_box(&m, 20.0, &gdi_flags());
    let gdi_typo = single_line_box(&m, 20.0, &gdi_typo_flags());

    assert!((float_win.height() - 17.2).abs() < 1e-4, "{float_win:?}");
    assert!((float_typo.height() - 15.5).abs() < 1e-4, "{float_typo:?}");
    assert_eq!(gdi_win.height(), 16.0);
    assert_eq!(gdi_typo.height(), 15.0);
}

#[test]
fn gdi_metrics_round_each_component_before_summing() {
    let m = synthetic_metrics();
    assert!(
        !CompatFlags::default().gdi_line_metrics,
        "quantization is opt-in; the float path stays the default"
    );

    let quantized = single_line_box(&m, 20.0, &gdi_flags());
    assert_eq!(
        quantized,
        LineBox {
            ascent: 12.0,
            descent: 3.0,
            leading: 1.0,
        }
    );
    assert_eq!(quantized.height(), 16.0);

    let float = single_line_box(&m, 20.0, &CompatFlags::default());
    assert!((float.height() - 17.2).abs() < 1e-4, "{}", float.height());
    assert_eq!(float.height().round(), 17.0);
}

#[test]
fn gdi_metrics_snap_the_em_size_to_an_integer_ppem_first() {
    let m = synthetic_metrics();

    let eleven_pt = single_line_box(&m, 11.0 * 96.0 / 72.0, &gdi_flags());
    assert_eq!(
        eleven_pt,
        LineBox {
            ascent: 9.0,
            descent: 3.0,
            leading: 1.0,
        }
    );

    for size_px in [14.5, 14.9, 15.0, 15.49] {
        assert_eq!(
            single_line_box(&m, size_px, &gdi_flags()),
            eleven_pt,
            "{size_px}"
        );
    }
    assert_ne!(single_line_box(&m, 15.5, &gdi_flags()), eleven_pt);
}

#[test]
fn use_typo_metrics_bit_selects_the_typographic_family() {
    let win = synthetic_metrics();
    let typo = typo_metrics();
    assert!(!win.use_typo_metrics());
    assert!(typo.use_typo_metrics());

    assert_eq!(
        single_line_box(&typo, 20.0, &gdi_typo_flags()),
        LineBox {
            ascent: 11.0,
            descent: 3.0,
            leading: 1.0,
        }
    );
    assert_eq!(
        single_line_box(&win, 20.0, &gdi_typo_flags()).height(),
        16.0
    );

    for compat in [CompatFlags::default(), gdi_flags()] {
        assert_eq!(
            single_line_box(&typo, 20.0, &compat),
            single_line_box(&win, 20.0, &compat)
        );
    }
}

#[test]
fn typo_line_gap_stays_signed() {
    // A real macOS font has a negative typo gap; clamping it would incorrectly
    // loosen its lines.
    let tamil = FontMetrics {
        units_per_em: 2048,
        os2_typo_ascender: 1550,
        os2_typo_descender: -717,
        os2_typo_line_gap: -210,
        os2_fs_selection: 0xC0,
        os2_version: 4,
        os2_code_page_range1: 0,
        ..synthetic_metrics()
    };
    assert_eq!(
        single_line_box(&tamil, 16.0, &gdi_typo_flags()),
        LineBox {
            ascent: 12.0,
            descent: 6.0,
            leading: -2.0,
        }
    );
    let single = single_line_box(&tamil, 16.0, &gdi_typo_flags());
    assert_eq!(single.height(), 16.0);
}

#[test]
fn use_typo_metrics_is_ignored_before_os2_version_4() {
    for version in [0u16, 1, 2, 3] {
        let old = FontMetrics {
            os2_version: version,
            os2_code_page_range1: 0,
            ..typo_metrics()
        };
        assert!(!old.use_typo_metrics(), "version {version}");
        assert_eq!(
            single_line_box(&old, 20.0, &gdi_typo_flags()).height(),
            16.0,
            "version {version} measures on win metrics"
        );
    }
}

#[test]
fn use_typo_metrics_falls_back_when_the_typo_box_is_unusable() {
    let win_box = single_line_box(&synthetic_metrics(), 20.0, &gdi_typo_flags());
    let fallback = |ascender: i16, descender: i16, line_gap: i16, why: &str| {
        let m = FontMetrics {
            os2_typo_ascender: ascender,
            os2_typo_descender: descender,
            os2_typo_line_gap: line_gap,
            ..typo_metrics()
        };
        assert_eq!(
            single_line_box(&m, 20.0, &gdi_typo_flags()),
            win_box,
            "{why}"
        );
    };

    fallback(0, 0, 0, "absent");
    fallback(555, 155, 0, "positive descender");
    fallback(555, 555, 0, "positive descender summing to zero");
    fallback(555, i16::MIN, 65, "descent past the design ceiling");
    fallback(-100, 155, 900, "non-positive ascent");
}

#[test]
fn gdi_no_leading_drops_the_leading_and_nothing_else() {
    let compat = CompatFlags {
        no_leading: true,
        ..gdi_flags()
    };
    let line = single_line_box(&synthetic_metrics(), 20.0, &compat);
    assert_eq!(
        line,
        LineBox {
            ascent: 12.0,
            descent: 3.0,
            leading: 0.0,
        }
    );

    // Drop leading before quantization so no rounded remainder survives.
    let both = CompatFlags {
        no_leading: true,
        ..gdi_typo_flags()
    };
    assert_eq!(single_line_box(&typo_metrics(), 20.0, &both).leading, 0.0);
}

#[test]
fn degenerate_sizes_yield_a_zero_box_on_both_paths() {
    let zero = LineBox {
        ascent: 0.0,
        descent: 0.0,
        leading: 0.0,
    };
    let broken_upem = FontMetrics {
        units_per_em: 0,
        ..synthetic_metrics()
    };
    let sizes = [
        f32::NAN,
        0.0,
        -20.0,
        f32::NEG_INFINITY,
        f32::INFINITY,
        -f32::MIN_POSITIVE,
    ];
    for compat in [CompatFlags::default(), gdi_flags(), gdi_typo_flags()] {
        assert_eq!(single_line_box(&broken_upem, 20.0, &compat), zero);
        for size_px in sizes {
            let line = single_line_box(&synthetic_metrics(), size_px, &compat);
            assert_eq!(line, zero, "{size_px} / gdi={}", compat.gdi_line_metrics);
        }
    }

    let tiny = single_line_box(&synthetic_metrics(), 0.2, &gdi_flags());
    assert_eq!(tiny.height(), 1.0);
    assert!(tiny.ascent >= 0.0 && tiny.descent >= 0.0 && tiny.leading >= 0.0);
}

#[test]
fn experimental_metrics_stay_bounded_and_non_negative() {
    let tiny_em = FontMetrics {
        units_per_em: 16,
        os2_win_ascent: u16::MAX,
        os2_win_descent: u16::MAX,
        hhea_ascender: i16::MAX,
        hhea_descender: i16::MIN,
        hhea_line_gap: i16::MAX,
        ..synthetic_metrics()
    };
    for compat in [gdi_flags(), typo_flags(), gdi_typo_flags()] {
        let line = single_line_box(&tiny_em, 16.0, &compat);
        for part in [line.ascent, line.descent, line.leading] {
            assert!(
                (0.0..=256.0).contains(&part),
                "{part} / gdi={}",
                compat.gdi_line_metrics
            );
        }
    }

    for compat in [gdi_flags(), typo_flags(), gdi_typo_flags()] {
        for size_px in [1.0e30, f32::MAX] {
            let line = single_line_box(&tiny_em, size_px, &compat);
            for part in [line.ascent, line.descent, line.leading] {
                assert!(
                    part.is_finite() && (0.0..=34_944.0).contains(&part),
                    "{part} at {size_px} / gdi={}",
                    compat.gdi_line_metrics
                );
            }
        }
    }

    let capped = single_line_box(&synthetic_metrics(), 2184.0, &gdi_flags());
    assert_eq!(
        capped,
        LineBox {
            ascent: 1354.0,
            descent: 371.0,
            leading: 153.0,
        }
    );
    for size_px in [1.0e9, f32::MAX] {
        assert_eq!(
            single_line_box(&synthetic_metrics(), size_px, &gdi_flags()),
            capped,
            "{size_px}"
        );
    }
}

#[test]
fn gdi_metrics_quantize_the_fixture_font_read_at_runtime() {
    let (store, id) = store_with_font();
    let m = store.metrics(id).unwrap();
    assert!(!m.use_typo_metrics(), "Liberation Sans leaves bit 7 clear");

    let line = single_line_box(m, 16.0, &gdi_flags());
    assert_eq!(
        line,
        LineBox {
            ascent: 14.0,
            descent: 3.0,
            leading: 1.0,
        }
    );
    assert_eq!(line.height(), 18.0);
    assert_eq!(liberation_single_16px().height() - line.height(), 0.3984375);
}

#[test]
fn auto_240_is_identity_and_480_doubles_height_into_leading() {
    let single = liberation_single_16px();

    let same = apply_spacing_rule(single, &LineSpacingRule::Auto { line_240ths: 240 });
    assert_eq!(same, single);

    let double = apply_spacing_rule(single, &LineSpacingRule::Auto { line_240ths: 480 });
    assert_eq!(double.height(), 2.0 * single.height());
    // ascent/descent stay put — all the extra pitch goes below the descent,
    // so selection rects hug the text at the top of the line box like Word
    assert_eq!(double.ascent, single.ascent);
    assert_eq!(double.descent, single.descent);
    assert_eq!(double.leading, 18.398438); // 36.796875 - 15.0078125 - 3.390625
}

/// Word splits an `exact` box 80/20 about the baseline whatever the content.
#[test]
fn auto_240_preserves_negative_leading_exactly() {
    let content = LineBox {
        ascent: 12.0,
        descent: 6.0,
        leading: -2.0,
    };

    assert_eq!(
        apply_spacing_rule(content, &LineSpacingRule::Auto { line_240ths: 240 }),
        content
    );

    let double = apply_spacing_rule(content, &LineSpacingRule::Auto { line_240ths: 480 });
    assert_eq!(
        double,
        LineBox {
            ascent: 12.0,
            descent: 6.0,
            leading: 14.0,
        }
    );

    let half = apply_spacing_rule(content, &LineSpacingRule::Auto { line_240ths: 120 });
    assert_eq!(half.height(), 8.0);
    assert_eq!(half.leading, 0.0);
}

#[test]
fn exact_rule_splits_the_fixed_box_by_a_font_independent_constant() {
    let single = liberation_single_16px();

    // smaller than content: clips (measurement just fixes the box)
    let clipped = apply_spacing_rule(single, &LineSpacingRule::Exact { px: 10.0 });
    assert_eq!(clipped.height(), 10.0);
    assert_eq!(clipped.ascent, 8.0);
    assert_eq!(clipped.descent, 2.0);
    assert_eq!(clipped.leading, 0.0);

    // taller than content: same split, still exactly the fixed height
    let padded = apply_spacing_rule(single, &LineSpacingRule::Exact { px: 40.0 });
    assert_eq!(padded.height(), 40.0);
    assert_eq!(padded.ascent, 32.0);
    assert_eq!(padded.descent, 8.0);

    // the split ignores the content box entirely
    let other = LineBox {
        ascent: 1.0,
        descent: 9.0,
        leading: 4.0,
    };
    assert_eq!(
        apply_spacing_rule(other, &LineSpacingRule::Exact { px: 40.0 }),
        padded
    );
}

#[test]
fn at_least_rule_floors_but_never_shrinks() {
    let single = liberation_single_16px();

    // floor below the content height: content wins untouched
    let unchanged = apply_spacing_rule(single, &LineSpacingRule::AtLeast { px: 10.0 });
    assert_eq!(unchanged, single);

    // floor above: height is exactly the floor and the slack lands above the
    // ascent, so the content descent is preserved from the bottom
    let floored = apply_spacing_rule(single, &LineSpacingRule::AtLeast { px: 30.0 });
    assert_eq!(floored.height(), 30.0);
    assert_eq!(floored.descent, single.descent);
    assert_eq!(floored.ascent, 30.0 - single.descent);
    assert_eq!(floored.leading, 0.0);
}

/// `Exact` and a floor-active `AtLeast` leave no leading, so `ascent +
/// descent` is the whole box and every baseline model agrees on the result.
/// A content-winning `AtLeast` keeps the content's natural leading.
#[test]
fn fixed_rules_leave_no_leading() {
    let single = liberation_single_16px();
    for px in [1.0, 10.0, 18.0, 40.0, 500.0] {
        for rule in [
            LineSpacingRule::Exact { px },
            LineSpacingRule::AtLeast { px },
        ] {
            let box_ = apply_spacing_rule(single, &rule);
            assert!(
                box_.ascent + box_.descent <= box_.height() + 1e-4,
                "{rule:?} at {px}: {box_:?}"
            );
            if box_.height() == px {
                assert_eq!(box_.leading, 0.0, "{rule:?} at {px}");
            }
        }
    }
}

#[test]
fn stretch_spaces_gives_equal_share_to_space_clusters_only() {
    let mut advances = [10.0, 5.0, 10.0, 5.0, 10.0];
    let is_space = [false, true, false, true, false];

    stretch_spaces(&mut advances, &is_space, 4.0);
    // 4px slack over 2 space clusters = +2 each; letters untouched
    assert_eq!(advances, [10.0, 7.0, 10.0, 7.0, 10.0]);
}

#[test]
fn stretch_spaces_is_a_no_op_without_slack_or_spaces() {
    let original = [10.0, 5.0, 10.0];

    // negative / zero slack
    let mut advances = original;
    stretch_spaces(&mut advances, &[false, true, false], 0.0);
    assert_eq!(advances, original);
    stretch_spaces(&mut advances, &[false, true, false], -3.0);
    assert_eq!(advances, original);

    // no expandable spaces
    let mut advances = original;
    stretch_spaces(&mut advances, &[false, false, false], 4.0);
    assert_eq!(advances, original);
}

#[test]
fn justification_gate_matches_word_soft_return_semantics() {
    let default = CompatFlags::default();
    let compat = CompatFlags {
        do_not_expand_shift_return: true,
        ..CompatFlags::default()
    };

    // mid-paragraph wrapped line: justified
    assert!(line_is_justified(false, false, &default));
    // final line ended by the paragraph mark: never justified
    assert!(!line_is_justified(true, false, &default));
    // soft-return line (even the paragraph's last): justified by default...
    assert!(line_is_justified(false, true, &default));
    assert!(line_is_justified(true, true, &default));
    // ...until w:doNotExpandShiftReturn turns it off
    assert!(!line_is_justified(false, true, &compat));
    assert!(!line_is_justified(true, true, &compat));
}

#[test]
fn kern_enabled_truth_table() {
    // w:kern of 0 disables kerning outright
    assert!(!kern_enabled(24, 0));
    assert!(!kern_enabled(0, 0));
    // font size below the threshold: no kerning
    assert!(!kern_enabled(19, 20));
    // at or above the threshold: kerning on
    assert!(kern_enabled(20, 20));
    assert!(kern_enabled(40, 20));
}

// end-to-end proof that the kern_features contract holds through rustybuzz:
// kern-on shaping tightens "AV" below the plain hmtx sum, kern-off equals it
#[test]
fn kern_features_gate_pair_kerning_in_shaping() {
    let (store, id) = store_with_font();
    let upem = store.metrics(id).unwrap().units_per_em as f32;

    let on = kern_features(true);
    assert!(on.is_empty(), "enabled kerning must not override defaults");
    let kerned: f32 = shape(&store, id, "AV", upem, &on)
        .unwrap()
        .iter()
        .map(|g| g.x_advance)
        .sum();

    let off = kern_features(false);
    let plain: f32 = shape(&store, id, "AV", upem, &off)
        .unwrap()
        .iter()
        .map(|g| g.x_advance)
        .sum();

    // hmtx advances of A and V are 1366 each
    assert_eq!(plain, 1366.0 + 1366.0);
    assert!(
        kerned < plain,
        "kern_features(true) must keep GPOS pair kerning: {kerned} vs {plain}"
    );
}

/// Word 16.113 (macOS) measures a face claiming an East Asian code page at
/// 1.3 x the hhea ascent-to-descent span, half-leading split, ignoring the
/// win and sTypo families and hhea.lineGap.
#[test]
fn east_asian_code_pages_select_the_cjk_line_pitch() {
    let m = FontMetrics {
        os2_code_page_range1: 0x0002_0000,
        ..synthetic_metrics()
    };
    let line = single_line_box(&m, 1000.0, &CompatFlags::default());
    assert_eq!(line.height(), 1.3 * 790.0);
    assert_eq!(line.ascent, 620.0 + 0.15 * 790.0);
    assert_eq!(line.descent, 170.0 + 0.15 * 790.0);
    assert_eq!(line.leading, 0.0);

    for bits in [0x0004_0000, 0x0008_0000, 0x0010_0000] {
        let gated = FontMetrics {
            os2_code_page_range1: bits,
            ..synthetic_metrics()
        };
        assert_eq!(
            single_line_box(&gated, 1000.0, &CompatFlags::default()),
            line
        );
    }
}

/// Johab (bit 21) and Thai (bit 16) are not East Asian code pages for this
/// rule, and neither is a face that only covers CJK.
#[test]
fn non_gating_code_pages_keep_the_latin_line_pitch() {
    let latin = single_line_box(&synthetic_metrics(), 1000.0, &CompatFlags::default());
    for bits in [0x0001_0000, 0x0020_0000, 0x0040_0000, 0x4000_01ff] {
        let m = FontMetrics {
            os2_code_page_range1: bits,
            ..synthetic_metrics()
        };
        assert_eq!(
            single_line_box(&m, 1000.0, &CompatFlags::default()),
            latin,
            "code page bits {bits:#x} must not gate"
        );
    }
}

/// The East Asian pitch reads hhea only: the win family, hhea.lineGap and
/// USE_TYPO_METRICS all leave it untouched.
#[test]
fn east_asian_line_pitch_ignores_win_gap_and_typo_metrics() {
    let base = FontMetrics {
        os2_code_page_range1: 0x0002_0000,
        ..synthetic_metrics()
    };
    let expected = single_line_box(&base, 1000.0, &CompatFlags::default());
    let variants = [
        FontMetrics {
            os2_win_ascent: 1200,
            os2_win_descent: 400,
            ..base
        },
        FontMetrics {
            hhea_line_gap: 600,
            ..base
        },
        FontMetrics {
            os2_fs_selection: USE_TYPO_METRICS,
            ..base
        },
    ];
    for m in variants {
        assert_eq!(
            single_line_box(&m, 1000.0, &CompatFlags::default()),
            expected
        );
    }
}

/// A face the host substituted measures the way Word measures the face the
/// document named, while every glyph-side answer stays the substitute's.
#[test]
fn a_substitute_measures_as_the_face_it_stands_in_for() {
    let (mut store, base) = store_with_font();
    let mincho = ooxml_text::word_fonts::requested_line_metrics("ＭＳ 明朝")
        .expect("MS Mincho is a known East Asian face");
    let view = store
        .register_substitute(base, mincho)
        .expect("a measurement view registers");
    assert_ne!(view, base, "the view is its own font id");

    let size_px = 32.0;
    let line = single_line_box(
        store.metrics(view).unwrap(),
        size_px,
        &CompatFlags::default(),
    );
    assert!(
        (line.height() - size_px * 1.3).abs() < 0.01,
        "MS Mincho spans one em, so Word's East Asian pitch is 1.3 em: {}",
        line.height()
    );
    let base_line = single_line_box(
        store.metrics(base).unwrap(),
        size_px,
        &CompatFlags::default(),
    );
    assert!(
        base_line.height() < line.height(),
        "Liberation Sans measures shorter on its own: {} vs {}",
        base_line.height(),
        line.height()
    );

    assert_eq!(
        store.metrics(view).unwrap().units_per_em,
        store.metrics(base).unwrap().units_per_em,
        "units per em describes the bytes, which shaping and outlines scale by"
    );
    assert_eq!(
        store.glyph_id(view, 'A').unwrap(),
        store.glyph_id(base, 'A').unwrap()
    );
    assert_eq!(
        store.advance_width(view, 'A').unwrap(),
        store.advance_width(base, 'A').unwrap()
    );
    assert_eq!(
        store.outline_glyph_json(view, 36).unwrap(),
        store.outline_glyph_json(base, 36).unwrap()
    );
    assert_eq!(
        shape(&store, view, "Ag fi", size_px, &[]).unwrap(),
        shape(&store, base, "Ag fi", size_px, &[]).unwrap()
    );
}

#[test]
fn a_substitute_view_rescales_the_requested_span_into_its_own_units() {
    let (mut store, base) = store_with_font();
    let malgun = ooxml_text::word_fonts::requested_line_metrics("Malgun Gothic").unwrap();
    let view = store.register_substitute(base, malgun).unwrap();
    let metrics = store.metrics(view).unwrap();
    let span = f32::from(metrics.hhea_ascender) - f32::from(metrics.hhea_descender);
    let requested_span = f32::from(malgun.hhea_ascender) - f32::from(malgun.hhea_descender);
    assert!(
        (span / f32::from(metrics.units_per_em) - requested_span / f32::from(malgun.units_per_em))
            .abs()
            < 1e-3
    );
}

/// A Latin substitute is measured the Latin way — the win box plus external
/// leading — off the requested face's span, not the East Asian pitch and not
/// the substitute's own box.
#[test]
fn a_latin_substitute_measures_at_the_requested_span() {
    let (mut store, base) = store_with_font();
    for (family, span_em) in [("Open Sans", 2789.0 / 2048.0), ("Lato", 2400.0 / 2000.0)] {
        let requested = ooxml_text::word_fonts::requested_line_metrics(family).expect(family);
        let view = store.register_substitute(base, requested).expect(family);
        let metrics = store.metrics(view).unwrap();
        assert!(
            !metrics.east_asian_line_metrics(),
            "{family} must not take the East Asian pitch"
        );
        let size_px = 32.0;
        let line = single_line_box(metrics, size_px, &CompatFlags::default());
        assert!(
            (line.height() - size_px * span_em).abs() < 0.01,
            "{family}: {} vs {}",
            line.height(),
            size_px * span_em
        );
        assert_eq!(
            store.advance_width(view, 'A').unwrap(),
            store.advance_width(base, 'A').unwrap(),
            "{family} keeps the substitute's advances"
        );
    }
}

/// The win box of a Latin view spans exactly the requested ascender to
/// descender, so the requested span survives whatever box the substitute has.
#[test]
fn a_latin_view_pins_the_win_box_to_the_requested_span() {
    let (mut store, base) = store_with_font();
    let requested = ooxml_text::word_fonts::requested_line_metrics("Lucida Sans Unicode").unwrap();
    let view = store.register_substitute(base, requested).unwrap();
    let metrics = store.metrics(view).unwrap();
    assert_eq!(
        i32::from(metrics.hhea_ascender),
        i32::from(metrics.os2_win_ascent)
    );
    assert_eq!(
        i32::from(metrics.hhea_descender),
        -i32::from(metrics.os2_win_descent)
    );
}

#[test]
fn a_family_with_no_known_metrics_leaves_its_substitute_alone() {
    for family in ["Arial", "Times New Roman", "Helvetica", "폴라리스바탕"] {
        assert!(
            ooxml_text::word_fonts::requested_line_metrics(family).is_none(),
            "{family}"
        );
    }
}

/// The Word faces the visual-fidelity corpus names and no bundled face stands
/// in for: each measured at the last-resort Liberation span before, so every
/// line under one sat high by the difference.
#[test]
fn a_substituted_corpus_face_measures_at_the_span_word_embeds() {
    let (mut store, base) = store_with_font();
    let size_px = 32.0;
    for (family, span_em) in [
        ("Century Gothic", 2440.0 / 2048.0),
        ("Lucida Sans", 2332.0 / 2048.0),
        ("Lucida Calligraphy", 2566.0 / 2048.0),
        ("Arial Narrow", 2319.0 / 2048.0),
        ("Cambria Math", 2403.0 / 2048.0),
    ] {
        let requested = ooxml_text::word_fonts::requested_line_metrics(family).expect(family);
        let view = store.register_substitute(base, requested).expect(family);
        let line = single_line_box(
            store.metrics(view).unwrap(),
            size_px,
            &CompatFlags::default(),
        );
        assert!(
            (line.height() - size_px * span_em).abs() < 0.01,
            "{family}: {} vs {}",
            line.height(),
            size_px * span_em
        );
    }
}

/// Yu Mincho is a sixth taller than Yu Gothic; measuring it with Yu Gothic's
/// span shortened every Mincho line in a Japanese document.
#[test]
fn yu_mincho_does_not_measure_at_yu_gothics_span() {
    let gothic = ooxml_text::word_fonts::requested_line_metrics("Yu Gothic").unwrap();
    let mincho = ooxml_text::word_fonts::requested_line_metrics("Yu Mincho").unwrap();
    assert_eq!(
        mincho,
        ooxml_text::word_fonts::requested_line_metrics("游明朝").unwrap()
    );
    assert_ne!(mincho, gothic);
    assert_eq!(mincho.hhea_ascender, 2038);
    assert_eq!(mincho.hhea_descender, -598);
    assert!(mincho.east_asian);
}

#[test]
fn a_grid_rounds_a_line_up_to_a_whole_number_of_rows() {
    let pitch = 24.0;
    assert_eq!(
        snap_line_height(10.0, pitch),
        pitch,
        "a short line fills its row"
    );
    assert_eq!(
        snap_line_height(pitch, pitch),
        pitch,
        "an exact row does not grow"
    );
    assert_eq!(
        snap_line_height(pitch + 0.5, pitch),
        2.0 * pitch,
        "past a row takes two"
    );
    assert_eq!(snap_line_height(2.0 * pitch, pitch), 2.0 * pitch);
    assert_eq!(snap_line_height(2.0 * pitch + 0.1, pitch), 3.0 * pitch);
}

#[test]
fn a_row_boundary_is_not_pushed_over_by_float_error() {
    let pitch = 17.7;
    for rows in 1..=40 {
        let height = rows as f32 * pitch;
        assert_eq!(
            snap_line_height(height, pitch),
            height,
            "{rows} whole rows stay {rows} rows"
        );
    }
}

#[test]
fn a_grid_grown_box_keeps_ascent_and_descent() {
    let content = LineBox {
        ascent: 20.0,
        descent: 5.0,
        leading: 0.0,
    };
    let snapped = snap_line_box(content, 40.0);
    assert_eq!(snapped.ascent, content.ascent);
    assert_eq!(snapped.descent, content.descent);
    assert_eq!(snapped.height(), 40.0);
    let two_rows = snap_line_box(
        LineBox {
            ascent: 36.0,
            descent: 9.0,
            leading: 0.0,
        },
        40.0,
    );
    assert_eq!(two_rows.ascent, 36.0);
    assert_eq!(two_rows.descent, 9.0);
    assert_eq!(two_rows.height(), 80.0, "content past one row takes two");
}
