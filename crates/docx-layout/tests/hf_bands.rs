use docx_layout::display_list::{DisplayList, HfKind, build_display_list_json};
use docx_layout::hit::{HitRegion, hit_test, hit_test_regions, range_rects, range_rects_in_region};

const SCENARIOS: &[&str] = &["hf-default-both", "hf-title-page", "hf-even-odd"];

fn fixture_path(name: &str, suffix: &str) -> std::path::PathBuf {
    std::path::Path::new(env!("CARGO_MANIFEST_DIR"))
        .join("tests/fixtures/hf")
        .join(format!("{name}.{suffix}.json"))
}

fn read(name: &str, suffix: &str) -> String {
    std::fs::read_to_string(fixture_path(name, suffix))
        .unwrap_or_else(|_| panic!("missing {name}.{suffix}.json"))
}

fn build(name: &str) -> DisplayList {
    let json = build_display_list_json(&read(name, "input")).expect("builds");
    serde_json::from_str(&json).expect("parses")
}

fn approx(a: f64, b: f64) {
    assert!((a - b).abs() < 2e-3, "expected {b}, got {a}");
}

#[test]
fn hf_regions_match_expected_mapping_and_geometry() {
    for name in SCENARIOS {
        let dl = build(name);
        let expect: serde_json::Value = serde_json::from_str(&read(name, "expect")).unwrap();
        let pages = expect["pages"].as_array().unwrap();
        assert_eq!(dl.pages.len(), pages.len(), "{name}: page count");

        for (i, (page, exp)) in dl.pages.iter().zip(pages).enumerate() {
            for (kind, region) in [("header", &page.header), ("footer", &page.footer)] {
                let exp_rid = exp[format!("{kind}RId")].as_str();
                let region = match exp_rid {
                    None => {
                        assert!(region.is_none(), "{name} p{i}: unexpected {kind} region");
                        continue;
                    }
                    Some(rid) => {
                        let r = region
                            .as_ref()
                            .unwrap_or_else(|| panic!("{name} p{i}: missing {kind} region"));
                        assert_eq!(r.r_id, rid, "{name} p{i}: {kind} rId");
                        r
                    }
                };
                assert_eq!(
                    region.kind,
                    if kind == "header" {
                        HfKind::Header
                    } else {
                        HfKind::Footer
                    }
                );
                let g = &exp[kind];
                approx(
                    region.y.as_f64().unwrap(),
                    g["y"]
                        .as_f64()
                        .unwrap_or_else(|| panic!("{name} p{i} {kind} y")),
                );
                approx(
                    region.height.as_f64().unwrap(),
                    g["height"].as_f64().unwrap(),
                );
                assert!(
                    !region.primitives.is_empty(),
                    "{name} p{i}: {kind} region painted nothing"
                );
            }
        }
    }
}

#[test]
fn hf_scenarios_snapshot_and_determinism() {
    let update = std::env::var("DL_SNAPSHOT_UPDATE").as_deref() == Ok("1");

    for name in SCENARIOS {
        let input = read(name, "input");
        let a = build_display_list_json(&input).expect("builds");
        let b = build_display_list_json(&input).expect("builds");
        assert_eq!(a, b, "{name}: display list build is not deterministic");

        let snapshot_file = fixture_path(name, "displaylist");
        let pretty: serde_json::Value = serde_json::from_str(&a).unwrap();
        let actual = format!("{}\n", serde_json::to_string_pretty(&pretty).unwrap());
        if update {
            std::fs::write(&snapshot_file, actual).expect("write snapshot");
            continue;
        }
        let expected = std::fs::read_to_string(&snapshot_file).unwrap_or_else(|_| {
            panic!(
                "missing snapshot for {name}; run DL_SNAPSHOT_UPDATE=1 cargo test -p docx-layout"
            )
        });
        assert_eq!(
            actual, expected,
            "{name}: display list drifted from snapshot"
        );
    }
}

#[test]
fn hf_payload_absent_emits_no_regions() {
    // strip headersFooters from a fixture and confirm the pages carry no
    // header/footer keys (byte-level backward compatibility)
    let mut v: serde_json::Value = serde_json::from_str(&read("hf-default-both", "input")).unwrap();
    v.as_object_mut().unwrap().remove("headersFooters");
    let out = build_display_list_json(&v.to_string()).expect("builds");
    assert!(
        !out.contains("\"header\"") && !out.contains("\"footer\""),
        "regions leaked without an HF payload"
    );
}

#[test]
fn hit_test_regions_resolves_bands_and_body() {
    // hf-default-both: header band y 48..72, footer band y 984..1008,
    // margins.left 96; header paragraph pm span [1,15]
    let dl = build("hf-default-both");

    let hit = hit_test_regions(&dl, 0, 120.0, 60.0).unwrap();
    assert_eq!(hit.region, HitRegion::Header);
    assert_eq!(hit.r_id.as_deref(), Some("rIdHdrDefault"));
    let pos = hit
        .pos
        .expect("header band with content resolves a position");
    assert!((1..=15).contains(&pos), "header pos out of HF span: {pos}");

    let hit = hit_test_regions(&dl, 0, 120.0, 990.0).unwrap();
    assert_eq!(hit.region, HitRegion::Footer);
    assert_eq!(hit.r_id.as_deref(), Some("rIdFtrDefault"));
    let pos = hit
        .pos
        .expect("footer band with content resolves a position");
    assert!((1..=15).contains(&pos), "footer pos out of HF span: {pos}");

    let hit = hit_test_regions(&dl, 0, 100.0, 150.0).unwrap();
    assert_eq!(hit.region, HitRegion::Body);
    assert_eq!(hit.r_id, None);
    assert_eq!(hit.pos, hit_test(&dl, 0, 100.0, 150.0));
    assert!(hit.pos.is_some());

    // even-odd: page 2 resolves the even header's rId
    let dl = build("hf-even-odd");
    let hit = hit_test_regions(&dl, 1, 120.0, 60.0).unwrap();
    assert_eq!(hit.region, HitRegion::Header);
    assert_eq!(hit.r_id.as_deref(), Some("rIdHdrEven"));

    // out-of-range page
    assert!(hit_test_regions(&dl, 9, 0.0, 0.0).is_none());
}

#[test]
fn range_rects_resolve_inside_hf_bands_scoped_by_rid() {
    // hf-default-both: header band y 48..72, footer band y 984..1008, header +
    // footer paragraphs both span pm [1,15]. A range in the HEADER doc must
    // yield rects inside the header band only (not the footer, not the body).
    let dl = build("hf-default-both");

    let header = range_rects_in_region(&dl, HitRegion::Header, Some("rIdHdrDefault"), 1, 15);
    assert!(!header.is_empty(), "header range yielded no rects");
    for r in &header {
        assert!(
            r.y >= 48.0 - 1.0 && r.y <= 72.0 + 1.0,
            "header rect y {} outside the header band [48,72]",
            r.y
        );
        assert!(r.width > 0.0 && r.height > 0.0);
    }

    let footer = range_rects_in_region(&dl, HitRegion::Footer, Some("rIdFtrDefault"), 1, 15);
    assert!(!footer.is_empty(), "footer range yielded no rects");
    for r in &footer {
        assert!(
            r.y >= 984.0 - 1.0 && r.y <= 1008.0 + 1.0,
            "footer rect y {} outside the footer band [984,1008]",
            r.y
        );
    }

    // an rId that no band carries yields nothing (variant disambiguation)
    assert!(
        range_rects_in_region(&dl, HitRegion::Header, Some("rIdNope"), 1, 15).is_empty(),
        "non-matching rId must not contribute rects"
    );

    // the body wrapper equals the region call with HitRegion::Body — the
    // r_id argument is ignored for the body doc
    let body_wrapper = range_rects(&dl, 1, 8);
    let body_region = range_rects_in_region(&dl, HitRegion::Body, None, 1, 8);
    assert_eq!(
        body_wrapper, body_region,
        "body wrapper drifted from region"
    );

    // a collapsed range is empty in every region
    assert!(range_rects_in_region(&dl, HitRegion::Header, Some("rIdHdrDefault"), 5, 5).is_empty());
}

#[test]
fn range_rects_region_matches_json_and_handle_paths() {
    use docx_layout::hit::range_rects_region_json;
    use docx_layout::session::{
        close_display_list, open_display_list, range_rects_region_by_handle,
    };

    let input = read("hf-default-both", "input");
    let dl_json = build_display_list_json(&input).expect("builds");

    // native fn, JSON-arg export, and by-handle export must all agree byte-for-byte
    let native = range_rects_in_region(
        &build("hf-default-both"),
        HitRegion::Header,
        Some("rIdHdrDefault"),
        1,
        15,
    );
    let via_json =
        range_rects_region_json(&dl_json, "header", "rIdHdrDefault", 1, 15).expect("json ok");
    assert_eq!(via_json, serde_json::to_string(&native).unwrap());

    let handle = open_display_list(&dl_json).expect("opens");
    let via_handle =
        range_rects_region_by_handle(handle, "header", "rIdHdrDefault", 1, 15).expect("handle ok");
    assert_eq!(via_handle, via_json, "by-handle drifted from JSON-arg");
    close_display_list(handle);

    // an unparseable region is an error, not a panic
    assert!(range_rects_region_json(&dl_json, "sidebar", "", 1, 15).is_err());
}
