use ooxml_text::family_metrics::{family_advance, family_metrics, runs_at_own_widths};
use ooxml_text::{
    CompatFlags, FontId, FontStore, PathCmd, ShapeDirection, echo_glyph_set, kern_features, shape,
    shape_with_direction, single_line_box,
};

const LIBERATION_SANS: &[u8] = include_bytes!("fonts/LiberationSans-Regular.ttf");

fn store_with_font() -> (FontStore, FontId) {
    let mut store = FontStore::new();
    let base = store.register(LIBERATION_SANS.to_vec()).unwrap();
    (store, base)
}

fn table_width(family: &str, text: &str, size: f32) -> f32 {
    let metrics = family_metrics(family, false, false).unwrap();
    text.chars()
        .map(|ch| family_advance(metrics, ch).unwrap())
        .sum::<f32>()
        * size
}

fn shaped_width(store: &FontStore, font: FontId, text: &str, size: f32) -> f32 {
    shape(store, font, text, size, &kern_features(true))
        .unwrap()
        .iter()
        .map(|glyph| glyph.x_advance)
        .sum()
}

fn table_offset(bytes: &[u8], tag: &[u8; 4]) -> usize {
    let count = u16::from_be_bytes(bytes[4..6].try_into().unwrap()) as usize;
    let entry = (0..count)
        .map(|index| 12 + index * 16)
        .find(|&offset| &bytes[offset..offset + 4] == tag)
        .unwrap();
    u32::from_be_bytes(bytes[entry + 8..entry + 12].try_into().unwrap()) as usize
}

#[test]
fn arial_narrow_uses_its_widths_for_shaping_and_character_advances() {
    let (mut store, base) = store_with_font();
    let view = store
        .register_named_substitute(base, "  ARIAL Narrow ")
        .unwrap();
    assert_ne!(view, base);
    assert!((store.advance_scale(view).unwrap() - 0.82).abs() < 0.01);
    let text = "AVATAR, office: minimum width 0123.";
    let size = 32.0;
    let expected = table_width("arial narrow", text, size);
    assert!((shaped_width(&store, view, text, size) - expected).abs() < 1e-3);
    assert!(shaped_width(&store, base, text, size) > expected);
    assert_eq!(store.advance_scale(base).unwrap(), 1.0);
    assert_eq!(
        store.font_bytes(view).unwrap(),
        store.font_bytes(base).unwrap()
    );
    let upem = f32::from(store.metrics(view).unwrap().units_per_em);
    for ch in text.chars() {
        let advance = store.advance_width(view, ch).unwrap().unwrap();
        assert!(
            (advance / upem * size - table_width("arial narrow", &ch.to_string(), size)).abs()
                < 1e-3
        );
    }
}

#[test]
fn arial_narrow_scales_outlines_and_untabulated_characters_by_the_base_ratio() {
    let (mut store, base) = store_with_font();
    let view = store
        .register_named_substitute(base, "Arial Narrow")
        .unwrap();
    let ratio = store.advance_scale(view).unwrap();
    let extent = |id| {
        let glyph = store.glyph_id(id, 'm').unwrap().unwrap();
        store
            .outline_glyph(id, glyph)
            .unwrap()
            .cmds
            .iter()
            .filter_map(|cmd| match *cmd {
                PathCmd::MoveTo { x, .. }
                | PathCmd::LineTo { x, .. }
                | PathCmd::QuadTo { x, .. }
                | PathCmd::CubicTo { x, .. } => Some(x),
                PathCmd::Close => None,
            })
            .fold(f32::MIN, f32::max)
    };
    assert!((extent(view) / extent(base) - ratio).abs() < 1e-4);
    let ch = '\u{03A9}';
    assert!(family_advance(family_metrics("arial narrow", false, false).unwrap(), ch).is_none());
    assert!(store.covers(base, ch).unwrap());
    assert_eq!(
        store.advance_width(view, ch).unwrap().unwrap(),
        store.advance_width(base, ch).unwrap().unwrap() * ratio
    );
    assert!(
        (shaped_width(&store, view, &ch.to_string(), 32.0)
            - shaped_width(&store, base, &ch.to_string(), 32.0) * ratio)
            .abs()
            < 1e-4
    );
    assert_eq!(store.advance_width(view, '\u{4E2D}').unwrap(), None);
}

#[test]
fn matching_and_unknown_families_preserve_base_kerning() {
    let (mut store, base) = store_with_font();
    let metrics = family_metrics("helvetica", false, false).unwrap();
    assert!(runs_at_own_widths(&store, base, metrics));
    let kerned = shape(&store, base, "AVATAR To", 32.0, &kern_features(true)).unwrap();
    assert_ne!(
        kerned,
        shape(&store, base, "AVATAR To", 32.0, &kern_features(false)).unwrap()
    );
    for family in ["Arial", "Helvetica", "Unmeasured Family"] {
        let view = store.register_named_substitute(base, family).unwrap();
        assert_eq!(view, base, "{family}");
        assert_eq!(
            shape(&store, view, "AVATAR To", 32.0, &kern_features(true)).unwrap(),
            kerned,
            "{family}"
        );
    }
}

#[test]
fn matching_widths_with_requested_line_metrics_keep_the_base_advance_scale() {
    let (mut store, original) = store_with_font();
    let mut bytes = LIBERATION_SANS.to_vec();
    let hmtx = table_offset(&bytes, b"hmtx");
    let metrics = family_metrics("arial narrow", false, false).unwrap();
    let upem = f32::from(store.metrics(original).unwrap().units_per_em);
    for ch in ' '..='~' {
        let glyph = store.glyph_id(original, ch).unwrap().unwrap();
        let offset = hmtx + usize::from(glyph) * 4;
        let advance = (family_advance(metrics, ch).unwrap() * upem).round() as u16;
        bytes[offset..offset + 2].copy_from_slice(&advance.to_be_bytes());
    }
    let base = store.register(bytes).unwrap();
    assert!(runs_at_own_widths(&store, base, metrics));
    let view = store
        .register_named_substitute(base, "Arial Narrow")
        .unwrap();
    assert_ne!(view, base);
    assert!(store.family_widths(view).unwrap().is_none());
    assert_eq!(store.advance_scale(view).unwrap(), 1.0);
    assert_eq!(
        shape(&store, view, "AVATAR To", 32.0, &kern_features(true)).unwrap(),
        shape(&store, base, "AVATAR To", 32.0, &kern_features(true)).unwrap()
    );
}

#[test]
fn outline_scale_is_computed_from_the_supplied_face_advances() {
    let (mut store, original) = store_with_font();
    let mut bytes = LIBERATION_SANS.to_vec();
    let hmtx = table_offset(&bytes, b"hmtx");
    let hhea = table_offset(&bytes, b"hhea");
    let count = u16::from_be_bytes(bytes[hhea + 34..hhea + 36].try_into().unwrap());
    for glyph in 0..count {
        let offset = hmtx + usize::from(glyph) * 4;
        let advance = u16::from_be_bytes(bytes[offset..offset + 2].try_into().unwrap());
        let advance = (f32::from(advance) * 0.9).round() as u16;
        bytes[offset..offset + 2].copy_from_slice(&advance.to_be_bytes());
    }
    let narrower = store.register(bytes).unwrap();
    let original_view = store
        .register_named_substitute(original, "Arial Narrow")
        .unwrap();
    let narrower_view = store
        .register_named_substitute(narrower, "Arial Narrow")
        .unwrap();
    assert!(store.family_widths(narrower_view).unwrap().is_some());
    assert!(
        (store.advance_scale(narrower_view).unwrap() / store.advance_scale(original_view).unwrap()
            - 1.0 / 0.9)
            .abs()
            < 1e-3
    );
    assert_eq!(
        shaped_width(&store, narrower_view, "minimum widths", 32.0),
        shaped_width(&store, original_view, "minimum widths", 32.0)
    );
}

#[test]
fn named_substitute_chooses_widths_from_base_selection_flags() {
    let mut store = FontStore::new();
    for (bold, italic) in [(false, false), (true, false), (false, true), (true, true)] {
        let mut bytes = LIBERATION_SANS.to_vec();
        let os2 = table_offset(&bytes, b"OS/2");
        let selection = (bold as u16) << 5 | italic as u16;
        bytes[os2 + 62..os2 + 64].copy_from_slice(&selection.to_be_bytes());
        let base = store.register(bytes).unwrap();
        let view = store
            .register_named_substitute(base, "Arial Narrow")
            .unwrap();
        let metrics = family_metrics("arial narrow", bold, italic).unwrap();
        let text = "AVATAR minimum";
        let expected = text
            .chars()
            .map(|ch| family_advance(metrics, ch).unwrap())
            .sum::<f32>()
            * 32.0;
        assert!((shaped_width(&store, view, text, 32.0) - expected).abs() < 1e-3);
    }
}

#[test]
fn lucida_bright_preserves_its_static_advance_ratio() {
    let (mut store, base) = store_with_font();
    let view = store
        .register_named_substitute(base, "Lucida Bright")
        .unwrap();
    assert_eq!(store.advance_scale(view).unwrap(), 1.113);
    assert!(
        (shaped_width(&store, view, "AVATAR To", 32.0)
            / shaped_width(&store, base, "AVATAR To", 32.0)
            - 1.113)
            .abs()
            < 1e-4
    );
}

#[test]
fn width_tables_without_vertical_entries_keep_all_base_vertical_metrics() {
    let (mut store, base) = store_with_font();
    for family in ["Book Antiqua", "Garamond", "Tw Cen MT Condensed"] {
        assert!(ooxml_text::word_fonts::requested_line_metrics(family).is_none());
        let view = store.register_named_substitute(base, family).unwrap();
        assert_ne!(view, base, "{family}");
        assert_eq!(store.metrics(view).unwrap(), store.metrics(base).unwrap());
        assert!(
            (shaped_width(&store, view, "AVATAR minimum", 32.0)
                - table_width(&family.to_lowercase(), "AVATAR minimum", 32.0))
            .abs()
                < 1e-3
        );
    }
}

#[test]
fn cambria_substitute_uses_requested_vertical_metrics() {
    for upem in [2048u16, 1000] {
        let mut bytes = LIBERATION_SANS.to_vec();
        let head = table_offset(&bytes, b"head");
        bytes[head + 18..head + 20].copy_from_slice(&upem.to_be_bytes());
        let mut store = FontStore::new();
        let base = store.register(bytes).unwrap();
        let view = store.register_named_substitute(base, "Cambria").unwrap();
        let metrics = store.metrics(view).unwrap();
        let scale = f32::from(metrics.units_per_em) / 2048.0;
        assert_eq!(metrics.hhea_ascender, (1946.0 * scale).round() as i16);
        assert_eq!(metrics.hhea_descender, (-455.0 * scale).round() as i16);
        assert_eq!(metrics.hhea_line_gap, 0);
        assert_eq!(metrics.os2_win_ascent, metrics.hhea_ascender as u16);
        assert_eq!(metrics.os2_win_descent, -metrics.hhea_descender as u16);
        let line = single_line_box(metrics, 32.0, &CompatFlags::default());
        assert!((line.height() - 32.0 * 2401.0 / 2048.0).abs() < 0.02);
        assert!(
            (shaped_width(&store, view, "AVATAR minimum", 32.0)
                - table_width("cambria", "AVATAR minimum", 32.0))
            .abs()
                < 1e-3
        );
    }
}

#[test]
fn substituted_widths_follow_logical_cluster_boundaries_in_both_directions() {
    let (mut store, base) = store_with_font();
    let view = store
        .register_named_substitute(base, "Arial Narrow")
        .unwrap();
    for direction in [ShapeDirection::Ltr, ShapeDirection::Rtl] {
        let text = "AV\u{00E9} i";
        let glyphs = shape_with_direction(&store, view, text, 32.0, &[], direction).unwrap();
        for glyph in &glyphs {
            let start = glyph.cluster as usize;
            let end = glyphs
                .iter()
                .map(|next| next.cluster as usize)
                .filter(|&offset| offset > start)
                .min()
                .unwrap_or(text.len());
            let expected = table_width("arial narrow", &text[start..end], 32.0);
            let actual: f32 = glyphs
                .iter()
                .filter(|next| next.cluster == glyph.cluster)
                .map(|next| next.x_advance)
                .sum();
            assert!((actual - expected).abs() < 1e-3, "{direction:?} {start}");
        }
        let text = "x\u{0301}";
        let plain = shape_with_direction(&store, base, text, 32.0, &[], direction).unwrap();
        let substituted = shape_with_direction(&store, view, text, 32.0, &[], direction).unwrap();
        assert!(
            family_advance(
                family_metrics("arial narrow", false, false).unwrap(),
                '\u{0301}'
            )
            .is_none()
        );
        assert_eq!(plain.len(), substituted.len());
        let ratio = store.advance_scale(view).unwrap();
        for (plain, substituted) in plain.iter().zip(&substituted) {
            assert!((substituted.x_advance - plain.x_advance * ratio).abs() < 1e-4);
            assert!((substituted.x_offset - plain.x_offset * ratio).abs() < 1e-4);
            assert_eq!(substituted.y_offset, plain.y_offset);
        }
    }
}

#[test]
fn substituted_echo_advances_equal_table_shaping() {
    let (mut store, base) = store_with_font();
    let view = store
        .register_named_substitute(base, "Arial Narrow")
        .unwrap();
    let alphabet: Vec<char> = "AVTo.x i".chars().collect();
    for size in [16.0, 13.333_333] {
        for kern in [false, true] {
            let set = echo_glyph_set(&store, view, size, kern, &alphabet).unwrap();
            assert!(set.glyph_ids.iter().all(|&id| id != 0));
            assert!(set.kerning.is_empty());
            for (index, ch) in alphabet.iter().enumerate() {
                let expected = shaped_width(&store, view, &ch.to_string(), size);
                assert_eq!(set.advances[index] as f32 * set.horizontal, expected);
            }
        }
    }
}

#[test]
fn soft_hyphens_preserve_zero_advance_in_substituted_words() {
    let (mut store, base) = store_with_font();
    let view = store
        .register_named_substitute(base, "Arial Narrow")
        .unwrap();
    for direction in [ShapeDirection::Ltr, ShapeDirection::Rtl] {
        let width = |text| {
            shape_with_direction(&store, view, text, 32.0, &kern_features(false), direction)
                .unwrap()
                .iter()
                .map(|glyph| glyph.x_advance)
                .sum::<f32>()
        };
        assert!((width("co\u{00AD}operate") - width("cooperate")).abs() < 1e-4);
        assert_eq!(width("\u{00AD}"), 0.0);
        assert!((width("-") - table_width("arial narrow", "-", 32.0)).abs() < 1e-4);
    }
    let measure = |text: &str, max_width| {
        let input = serde_json::from_value(serde_json::json!({
            "block": {
                "kind": "paragraph",
                "runs": [{"kind": "text", "text": text, "fontFamily": "Arial Narrow"}]
            },
            "maxWidth": max_width,
            "fontChains": {"arial narrow|0|0": [view.to_u32()]},
            "defaults": {"fontSize": 24.0, "fontFamily": "Arial Narrow"}
        }))
        .unwrap();
        ooxml_text::measure_paragraph(&store, &input).unwrap()
    };
    let plain = measure("cooperate", 1000.0);
    let discretionary = measure("co\u{00AD}operate", 1000.0);
    assert_eq!(plain.lines.len(), 1);
    assert_eq!(discretionary.lines.len(), 1);
    assert!((plain.lines[0].width - discretionary.lines[0].width).abs() < 1e-4);
    assert_eq!(
        measure("co\u{00AD}operate", plain.lines[0].width)
            .lines
            .len(),
        1
    );
    let wrapped = measure("co\u{00AD}operate", plain.lines[0].width * 0.9);
    assert!(wrapped.lines.len() > 1);
    assert_eq!(wrapped.lines[0].tail_char, 3);
    let echo = echo_glyph_set(&store, view, 32.0, false, &['c', '\u{00AD}', '-']).unwrap();
    assert_eq!(echo.advances[1], 0);
    assert_ne!(echo.glyph_ids[2], 0);
    assert!(
        (echo.advances[2] as f32 * echo.horizontal - table_width("arial narrow", "-", 32.0)).abs()
            < 1e-4
    );
}
