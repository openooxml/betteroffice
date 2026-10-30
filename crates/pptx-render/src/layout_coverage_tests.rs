mod coverage_regressions {
    use super::*;

    fn latin_renderer() -> SlideRenderer {
        let mut renderer = SlideRenderer::new();
        renderer
            .register_font(
                "Arial",
                false,
                false,
                include_bytes!("../../fontpack/tests/fixtures/Synthetic-Regular.ttf"),
            )
            .unwrap();
        renderer
    }

    fn split(renderer: &SlideRenderer, text: &str) -> Vec<ResolvedRun> {
        let mut run = paragraph(renderer, "l", text).runs.remove(0);
        run.start = 7;
        let mut result = Vec::new();
        renderer.split_by_coverage(run, &mut result);
        result
    }

    #[test]
    fn browser_coverage_uses_primary_faces_for_cyrillic() {
        let mut renderer = latin_renderer();
        let own = renderer.fallback_font().unwrap();
        let fallback = renderer
            .register_font("Trebuchet MS", false, false, FONT)
            .unwrap();
        assert!(
            renderer
                .resolve_face(
                    "Trebuchet MS",
                    false,
                    false,
                    &mut SubstitutionLog::default()
                )
                .unwrap()
                .widths
                .is_some()
        );
        let cyrillic = "България";
        for character in cyrillic.chars() {
            assert_eq!(renderer.fonts.covers(own, character), Ok(false));
            assert_eq!(
                renderer.fonts.covers(FontId::from_u32(fallback), character),
                Ok(true)
            );
        }
        let text = format!("A{cyrillic}B");
        for enabled in [false, true, false] {
            renderer.set_browser_coverage_fallback(enabled);
            let pieces = split(&renderer, &text);
            let actual: Vec<_> = pieces
                .iter()
                .map(|piece| (piece.text.as_str(), piece.style.face.id.to_u32()))
                .collect();
            if enabled {
                assert_eq!(
                    actual,
                    [
                        ("A", own.to_u32()),
                        (cyrillic, fallback),
                        ("B", own.to_u32())
                    ]
                );
            } else {
                assert_eq!(actual, [(text.as_str(), own.to_u32())]);
            }
        }

        renderer.set_browser_coverage_fallback(true);
        let mut paragraph = paragraph(&renderer, "l", &text);
        let original_line = style_line_box(&renderer.fonts, &paragraph.runs[0].style, 1.0).unwrap();
        paragraph.runs = split(&renderer, &text);
        let selected = &paragraph.runs[1].style;
        assert_eq!(selected.face.line_id, own);
        assert!(selected.face.widths.is_none());
        let selected_line = style_line_box(&renderer.fonts, selected, 1.0).unwrap();
        assert_eq!(selected_line.ascent, original_line.ascent);
        assert_eq!(selected_line.descent, original_line.descent);
        assert_eq!(selected_line.leading, original_line.leading);

        let expected = shape(
            &renderer.fonts,
            FontId::from_u32(fallback),
            cyrillic,
            points_to_px(selected.font_size_pt),
            &[],
        )
        .unwrap();
        let expected_width: f32 = expected.iter().map(|glyph| glyph.x_advance).sum();
        let lines = layout_paragraph(
            &renderer.fonts,
            &paragraph,
            0.0,
            0.0,
            1_000.0,
            1.0,
            false,
            false,
        )
        .unwrap();
        let drawn = lines[0]
            .runs
            .iter()
            .find(|run| run.text == cyrillic)
            .unwrap();
        assert_eq!(drawn.font_id, fallback);
        assert_eq!(drawn.font_family, "Trebuchet MS");
        assert!((drawn.width - expected_width).abs() < 0.01);
        let drawn_advance: f32 = drawn.glyphs.iter().map(|glyph| glyph.advance).sum();
        assert!((drawn_advance - expected_width).abs() < 0.01);
        assert!(drawn.glyphs.iter().all(|glyph| glyph.glyph_id != 0));
    }

    #[test]
    fn browser_coverage_still_uses_browser_when_no_primary_face_covers() {
        for registered_primary in [false, true] {
            let mut renderer = latin_renderer();
            if registered_primary {
                renderer
                    .register_font(
                        "Other Latin",
                        false,
                        false,
                        include_bytes!("../../fontpack/tests/fixtures/Synthetic-Regular.ttf"),
                    )
                    .unwrap();
            }
            renderer.set_browser_coverage_fallback(true);
            let own = renderer.fallback_font().unwrap().to_u32();
            let pieces = split(&renderer, "AБB");
            let actual: Vec<_> = pieces
                .iter()
                .map(|piece| (piece.text.as_str(), piece.style.face.id.to_u32()))
                .collect();
            assert_eq!(actual, [("A", own), ("Б", BROWSER_FONT_ID), ("B", own)]);
        }
    }

    #[test]
    fn primary_coverage_requires_the_whole_grapheme() {
        let mut renderer = latin_renderer();
        renderer.set_browser_coverage_fallback(true);
        let partial = renderer
            .register_font("Partial primary", false, false, FONT)
            .unwrap();
        renderer
            .register_fallback_font("Partial coverage", false, false, FONT)
            .unwrap();
        assert_eq!(
            renderer.fonts.covers(FontId::from_u32(partial), 'Б'),
            Ok(true)
        );
        assert_eq!(
            renderer.fonts.covers(FontId::from_u32(partial), '\u{20DD}'),
            Ok(false)
        );
        let pieces = split(&renderer, "AБ\u{20DD}B");
        assert_eq!(pieces.len(), 3);
        assert_eq!(pieces[1].text, "Б\u{20DD}");
        assert_eq!(pieces[1].style.face.id.to_u32(), BROWSER_FONT_ID);

        let covering = renderer
            .register_font(
                "Carlito",
                false,
                false,
                include_bytes!("../../../packages/fonts/assets/Carlito-Regular.ttf"),
            )
            .unwrap();
        for character in "Б\u{20DD}".chars() {
            assert_eq!(
                renderer.fonts.covers(FontId::from_u32(covering), character),
                Ok(true)
            );
        }
        let pieces = split(&renderer, "AБ\u{20DD}B");
        assert_eq!(pieces.len(), 3);
        assert_eq!(pieces[1].text, "Б\u{20DD}");
        assert_eq!(pieces[1].style.face.id.to_u32(), covering);
    }

    #[test]
    fn primary_coverage_prefers_style_then_registration_order() {
        for (bold, italic) in [(false, false), (false, true), (true, false), (true, true)] {
            let mut renderer = latin_renderer();
            let face = renderer
                .resolve_face("Arial", false, false, &mut SubstitutionLog::default())
                .unwrap();
            let selected = |renderer: &SlideRenderer| {
                renderer
                    .browser_coverage_face(&face, "Б", bold, italic)
                    .unwrap()
                    .id
                    .to_u32()
            };
            let first_any = renderer
                .register_font("First any", !bold, !italic, FONT)
                .unwrap();
            renderer
                .register_font("Second any", !bold, italic, FONT)
                .unwrap();
            assert_eq!(selected(&renderer), first_any);

            let first_bold = renderer
                .register_font("First bold", bold, !italic, FONT)
                .unwrap();
            renderer
                .register_font("Second bold", bold, !italic, FONT)
                .unwrap();
            assert_eq!(selected(&renderer), first_bold);

            let first_exact = renderer
                .register_font("First exact", bold, italic, FONT)
                .unwrap();
            renderer
                .register_font("Second exact", bold, italic, FONT)
                .unwrap();
            assert_eq!(selected(&renderer), first_exact);

            let coverage = renderer
                .register_fallback_font("Coverage", !bold, !italic, FONT)
                .unwrap();
            assert_eq!(selected(&renderer), coverage);
        }
    }

    #[test]
    fn browser_coverage_is_off_by_default() {
        for registered_fallback in [false, true] {
            let mut renderer = renderer();
            if registered_fallback {
                renderer
                    .register_fallback_font("Liberation Sans", false, false, FONT)
                    .unwrap();
            }
            let own = renderer.fallback_font().unwrap();
            for text in ["AहB", "A👩‍💻B", "A1️⃣B", "A\u{E000}B"] {
                let pieces = split(&renderer, text);
                assert_eq!(pieces.len(), 1, "{text:?}");
                assert_eq!(pieces[0].text, text);
                assert_eq!(pieces[0].start, 7);
                assert_eq!(pieces[0].style.face.id, own);
            }
        }
    }

    #[test]
    fn disabled_browser_coverage_keeps_registered_fallbacks_scalar() {
        let mut renderer = renderer();
        let own = renderer.fallback_font().unwrap();
        let fallback = renderer
            .register_fallback_font(
                "Carlito",
                false,
                false,
                include_bytes!("../../../packages/fonts/assets/Carlito-Regular.ttf"),
            )
            .unwrap();
        assert_eq!(renderer.fonts.covers(own, '\u{20DD}'), Ok(false));
        assert_eq!(
            renderer.fonts.covers(FontId::from_u32(fallback), '\u{20DD}'),
            Ok(true)
        );
        for enabled in [false, true, false] {
            renderer.set_browser_coverage_fallback(enabled);
            let pieces = split(&renderer, "a\u{20DD}b");
            let actual: Vec<_> = pieces
                .iter()
                .map(|piece| {
                    (
                        piece.text.as_str(),
                        piece.start,
                        piece.style.face.id.to_u32(),
                    )
                })
                .collect();
            if enabled {
                assert_eq!(actual, [("a\u{20DD}", 7, fallback), ("b", 9, own.to_u32())]);
            } else {
                assert_eq!(
                    actual,
                    [
                        ("a", 7, own.to_u32()),
                        ("\u{20DD}", 8, fallback),
                        ("b", 9, own.to_u32()),
                    ]
                );
            }
        }
    }

    #[test]
    fn browser_coverage_keeps_grapheme_clusters_whole() {
        for registered_fallback in [false, true] {
            let mut renderer = renderer();
            renderer.set_browser_coverage_fallback(true);
            if registered_fallback {
                renderer
                    .register_fallback_font("Liberation Sans", false, false, FONT)
                    .unwrap();
            }
            let own = renderer.fallback_font().unwrap().to_u32();
            for cluster in [
                "🏞️",
                "👩‍💻",
                "👍🏽",
                "1️⃣",
                "🇩🇪",
                "🏞\u{FE0E}",
                "🏞️\u{0301}",
                "🏞️\u{20DD}",
                "🏴\u{E0067}\u{E0062}\u{E0065}\u{E006E}\u{E0067}\u{E007F}",
            ] {
                for (text, expected) in [
                    (
                        format!("{cluster}x"),
                        vec![
                            (cluster, 7, BROWSER_FONT_ID),
                            ("x", 7 + utf16_len(cluster), own),
                        ],
                    ),
                    (
                        format!("A{cluster}x"),
                        vec![
                            ("A", 7, own),
                            (cluster, 8, BROWSER_FONT_ID),
                            ("x", 8 + utf16_len(cluster), own),
                        ],
                    ),
                ] {
                    let pieces = split(&renderer, &text);
                    let actual: Vec<_> = pieces
                        .iter()
                        .map(|piece| {
                            (
                                piece.text.as_str(),
                                piece.start,
                                piece.style.face.id.to_u32(),
                            )
                        })
                        .collect();
                    assert_eq!(actual, expected, "{text:?}");
                }
            }
        }
    }

    #[test]
    fn browser_coverage_keeps_invisible_only_runs_unchanged() {
        for registered_fallback in [false, true] {
            let mut renderer = renderer();
            renderer.set_browser_coverage_fallback(true);
            if registered_fallback {
                renderer
                    .register_fallback_font("Liberation Sans", false, false, FONT)
                    .unwrap();
            }
            let own = renderer.fallback_font().unwrap();
            for text in [
                "",
                "\n\t\u{200B}\u{200D}\u{FE0E}\u{FE0F}\u{E0067}\u{E007F}\u{E0100}",
            ] {
                let pieces = split(&renderer, text);
                assert_eq!(pieces.len(), 1);
                assert_eq!(pieces[0].text, text);
                assert_eq!(pieces[0].start, 7);
                assert_eq!(pieces[0].style.face.id, own);
            }
        }
    }

    #[test]
    fn alternating_browser_coverage_splitting_scales_linearly() {
        let mut renderer = renderer();
        renderer.set_browser_coverage_fallback(true);
        let measure = |count| {
            let text = "Aह".repeat(count);
            (0..3)
                .map(|_| {
                    let start = std::time::Instant::now();
                    let pieces = split(&renderer, &text);
                    let elapsed = start.elapsed();
                    assert_eq!(pieces.len(), count * 2);
                    for (index, piece) in pieces.iter().enumerate() {
                        assert_eq!(piece.start, 7 + index as u32);
                        assert_eq!(piece.text, if index % 2 == 0 { "A" } else { "ह" });
                    }
                    elapsed
                })
                .min()
                .unwrap()
        };
        let small = measure(2_000);
        let large = measure(20_000);
        assert!(large <= small * 30 + std::time::Duration::from_millis(10));
    }

    #[test]
    fn coverage_offsets_keep_supplementary_characters_and_run_origins() {
        let mut renderer = renderer();
        renderer.set_browser_coverage_fallback(true);
        let pieces = split(&renderer, "A😀हA");
        let pieces: Vec<_> = pieces
            .iter()
            .map(|piece| (piece.text.as_str(), piece.start))
            .collect();
        assert_eq!(pieces, [("A", 7), ("😀ह", 8), ("A", 11)]);
    }

    #[test]
    fn chart_coverage_offsets_keep_supplementary_characters() {
        for enabled in [false, true] {
            chart_coverage_offsets(enabled);
        }
    }

    fn chart_coverage_offsets(enabled: bool) {
        let mut renderer = renderer();
        renderer.set_browser_coverage_fallback(enabled);
        renderer
            .register_fallback_font(
                "Noto Sans Arabic",
                false,
                false,
                include_bytes!("../../../packages/fonts/assets/NotoSansArabic-Regular.ttf"),
            )
            .unwrap();
        let primitive = chart_text_primitive(
            &renderer,
            &Theme::default(),
            "chart:1",
            ChartText {
                object_id: 1,
                text: "A😀مA",
                x: 0.0,
                baseline_y: 40.0,
                width: 200.0,
                font: PlotFont {
                    weight: 400,
                    size_px: 14.0,
                    family: "Arial".to_owned(),
                    italic: false,
                    letter_spacing_px: 0.0,
                },
                color: "#000000",
                align: PlotTextAlign::Start,
                rotation_deg: 0.0,
            },
            &mut SubstitutionLog::default(),
        )
        .unwrap();
        let Primitive::TextBox { lines, .. } = primitive else {
            panic!("a text box");
        };
        let runs: Vec<_> = lines[0]
            .runs
            .iter()
            .map(|run| (run.text.as_str(), run.start, run.end))
            .collect();
        if enabled {
            assert_eq!(runs, [("A", 0, 1), ("😀", 1, 3), ("م", 3, 4), ("A", 4, 5)]);
            assert_eq!(lines[0].runs[1].font_id, BROWSER_FONT_ID);
        } else {
            assert_eq!(runs, [("A😀", 0, 3), ("م", 3, 4), ("A", 4, 5)]);
            assert!(lines[0].runs.iter().all(|run| run.font_id != BROWSER_FONT_ID));
        }
        let arabic = lines[0].runs.iter().find(|run| run.text == "م").unwrap();
        assert_eq!(arabic.glyphs[0].cluster, 3);
        assert_eq!(lines[0].runs.last().unwrap().glyphs[0].cluster, 4);
    }

    fn chart_run_widths(enabled: bool, text: &str) -> Vec<(u32, f32)> {
        let mut renderer = renderer();
        renderer.set_browser_coverage_fallback(enabled);
        let primitive = chart_text_primitive(
            &renderer,
            &Theme::default(),
            "chart:1",
            ChartText {
                object_id: 1,
                text,
                x: 0.0,
                baseline_y: 40.0,
                width: 2000.0,
                font: PlotFont {
                    weight: 400,
                    size_px: 20.0,
                    family: "Arial".to_owned(),
                    italic: false,
                    letter_spacing_px: 0.0,
                },
                color: "#000000",
                align: PlotTextAlign::Start,
                rotation_deg: 0.0,
            },
            &mut SubstitutionLog::default(),
        )
        .unwrap();
        let Primitive::TextBox { lines, .. } = primitive else {
            panic!("a text box");
        };
        lines[0]
            .runs
            .iter()
            .map(|run| (run.font_id, run.width))
            .collect()
    }

    #[test]
    fn browser_coverage_gives_full_width_east_asian_text_its_em_box() {
        for text in ["本町商店街周辺", "データ提供元", "한국어", "Ａ区間"] {
            let count = text.chars().count() as f32;
            let browser = chart_run_widths(true, text);
            assert_eq!(browser.len(), 1, "{text}");
            assert_eq!(browser[0].0, BROWSER_FONT_ID, "{text}");
            assert!(
                (browser[0].1 - count * 20.0).abs() <= count * 20.0 * 0.02,
                "{text}: {} for {count} characters",
                browser[0].1
            );
            let scalar: f32 = chart_run_widths(false, text).iter().map(|run| run.1).sum();
            assert!(browser[0].1 >= scalar, "{text}: {} < {scalar}", browser[0].1);
        }
        for (decomposed, composed) in [("か\u{3099}か\u{3099}", "がが"), ("ハ\u{309a}", "パ")] {
            let decomposed = chart_run_widths(true, decomposed);
            let composed = chart_run_widths(true, composed);
            assert_eq!(decomposed.len(), 1);
            assert!((decomposed[0].1 - composed[0].1).abs() < 0.01, "{decomposed:?} {composed:?}");
        }
        let halfwidth = chart_run_widths(true, "ｱｲｳ");
        assert_eq!(halfwidth[0].0, BROWSER_FONT_ID);
        assert!((halfwidth[0].1 - 30.0).abs() <= 0.6, "{}", halfwidth[0].1);
    }
}
