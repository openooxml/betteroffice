use ooxml_text::{
    BaseDirection, BreakOpportunity, CompatFlags, FontStore, LineBox, LineSpacingRule,
    ShapeDirection, ShapeFeature, apply_spacing_rule, bidi_paragraphs, break_opportunities,
    kern_enabled, kern_features, line_is_justified, shape, shape_with_direction, single_line_box,
    stretch_spaces,
};

const LIBERATION_SANS: &[u8] = include_bytes!("fonts/LiberationSans-Regular.ttf");
const NOTO_NASKH_ARABIC: &[u8] = include_bytes!("fonts/NotoNaskhArabic-Regular.ttf");

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
    assert_eq!(glyphs[0].glyph_id, 36);
    assert_eq!(glyphs[1].glyph_id, 57);
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

// ---- line metrics -------------------------------------------------------
//
// All expected values below are hand-computed from the fixture's raw tables:
// upem 2048, usWinAscent 1854, usWinDescent 434, hhea ascender 1854,
// descender -434, lineGap 67. At 16px the scale is 16/2048 = 0.0078125, an
// exact binary fraction, so every expected value is exact in f32 and
// assert_eq! is legitimate.

/// Fixture single-spacing box at 16px, computed independently of the crate.
fn liberation_single_16px() -> LineBox {
    LineBox {
        ascent: 14.484375,
        descent: 3.390625,
        leading: 0.5234375,
    }
}

#[test]
fn single_line_box_uses_win_metrics_and_gdi_external_leading() {
    let (store, id) = store_with_font();
    let m = store.metrics(id).unwrap();

    let line = single_line_box(m, 16.0, &CompatFlags::default());
    assert_eq!(line, liberation_single_16px());
    // full pitch = 18.3984375, exact in f32
    assert_eq!(line.height(), (1854.0 + 434.0 + 67.0) * 16.0 / 2048.0);
}

#[test]
fn no_leading_compat_flag_drops_external_leading_only() {
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
            leading: 0.0,
            ..liberation_single_16px()
        }
    );
}

#[test]
fn auto_240_is_identity_and_480_doubles_height_into_leading() {
    let single = liberation_single_16px();

    let same = apply_spacing_rule(single, &LineSpacingRule::Auto { line_240ths: 240 });
    assert_eq!(same, single);

    let double = apply_spacing_rule(single, &LineSpacingRule::Auto { line_240ths: 480 });
    assert_eq!(double.height(), 2.0 * single.height());
    // ascent/descent stay put — all the extra pitch goes below the descent,
    // Selection geometry stays at the top of the line box.
    assert_eq!(double.ascent, single.ascent);
    assert_eq!(double.descent, single.descent);
    assert_eq!(double.leading, 18.921875);
}

#[test]
fn exact_rule_fixes_height_and_preserves_descent_bottom_up() {
    let single = liberation_single_16px();

    // smaller than content: clips (measurement just fixes the box), the
    // baseline keeps the content descent from the bottom, ascent absorbs it
    let clipped = apply_spacing_rule(single, &LineSpacingRule::Exact { px: 10.0 });
    assert_eq!(clipped.height(), 10.0);
    assert_eq!(clipped.descent, single.descent);
    assert_eq!(clipped.ascent, 10.0 - single.descent);
    assert_eq!(clipped.leading, 0.0);

    // taller than content: still exactly the fixed height
    let padded = apply_spacing_rule(single, &LineSpacingRule::Exact { px: 40.0 });
    assert_eq!(padded.height(), 40.0);
    assert_eq!(padded.descent, single.descent);
}

#[test]
fn at_least_rule_floors_but_never_shrinks() {
    let single = liberation_single_16px();

    // floor below the content height: content wins untouched
    let unchanged = apply_spacing_rule(single, &LineSpacingRule::AtLeast { px: 10.0 });
    assert_eq!(unchanged, single);

    // floor above: height is exactly the floor, extra goes to leading
    let floored = apply_spacing_rule(single, &LineSpacingRule::AtLeast { px: 30.0 });
    assert_eq!(floored.height(), 30.0);
    assert_eq!(floored.ascent, single.ascent);
    assert_eq!(floored.descent, single.descent);
    assert_eq!(floored.leading, 30.0 - single.ascent - single.descent);
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
fn justification_gate_honors_soft_return_compatibility() {
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
    // Compatibility settings can disable soft-return justification.
    assert!(!line_is_justified(false, true, &compat));
    assert!(!line_is_justified(true, true, &compat));
}

#[test]
fn kern_enabled_truth_table() {
    // A zero threshold disables kerning.
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
