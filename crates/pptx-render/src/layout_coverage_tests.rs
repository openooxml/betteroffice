mod coverage_regressions {
    use super::*;

    fn split(renderer: &SlideRenderer, text: &str) -> Vec<ResolvedRun> {
        let mut run = paragraph(renderer, "l", text).runs.remove(0);
        run.start = 7;
        let pieces = renderer.coverage_pieces(&run.style.face, &run.text, false, false, true);
        let mut result = Vec::new();
        split_coverage_pieces(run, pieces, &mut result);
        result
    }

    #[test]
    fn browser_coverage_keeps_grapheme_clusters_whole() {
        for registered_fallback in [false, true] {
            let mut renderer = renderer();
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
        let renderer = renderer();
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
        let renderer = renderer();
        let pieces = split(&renderer, "A😀हA");
        let pieces: Vec<_> = pieces
            .iter()
            .map(|piece| (piece.text.as_str(), piece.start))
            .collect();
        assert_eq!(pieces, [("A", 7), ("😀ह", 8), ("A", 11)]);
    }

    #[test]
    fn chart_coverage_offsets_keep_supplementary_characters() {
        let mut renderer = renderer();
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
        if cfg!(target_arch = "wasm32") {
            assert_eq!(runs, [("A", 0, 1), ("😀", 1, 3), ("م", 3, 4), ("A", 4, 5)]);
        } else {
            assert_eq!(runs, [("A😀", 0, 3), ("م", 3, 4), ("A", 4, 5)]);
        }
        let arabic = lines[0].runs.iter().find(|run| run.text == "م").unwrap();
        assert_eq!(arabic.glyphs[0].cluster, 3);
        assert_eq!(lines[0].runs.last().unwrap().glyphs[0].cluster, 4);
    }
}
