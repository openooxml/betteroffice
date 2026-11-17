#[allow(dead_code)]
#[path = "../../pptx-parse/tests/support/header_footer.rs"]
mod support;

use pptx_edit::{DeckSession, EditCtx, TextStyle};
use pptx_render::{Primitive, SlideRenderer, TextLink, TextLinkTarget, TextParagraph};

const FONT: &[u8] = include_bytes!("../../ooxml-text/tests/fonts/LiberationSans-Regular.ttf");
const REL: &str = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

fn run(text: &str, attributes: &str) -> String {
    format!(r#"<a:r><a:rPr><a:hlinkClick {attributes}/></a:rPr><a:t>{text}</a:t></a:r>"#)
}

fn shape(paragraphs: &str) -> String {
    support::placeholder(2, "body", paragraphs, false).replace(
        "<p:spPr/>",
        r#"<p:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="8000000" cy="4000000"/></a:xfrm></p:spPr>"#,
    )
}

fn relationship(id: &str, kind: &str, target: &str, external: bool) -> String {
    let mode = if external {
        r#" TargetMode="External""#
    } else {
        ""
    };
    format!(r#"<Relationship Id="{id}" Type="{REL}/{kind}" Target="{target}"{mode}/>"#)
}

fn bytes(paragraphs: &str, relationships: &str) -> Vec<u8> {
    let shape = shape(paragraphs);
    let bytes = support::deck(
        &[(&shape, false), ("", false), ("", false)],
        "",
        "",
        "",
        Some(12),
    );
    let mut parts = ooxml_opc::unzip_parts(&bytes).unwrap();
    let (_, xml) = parts
        .iter_mut()
        .find(|(path, _)| path == "ppt/slides/_rels/slide1.xml.rels")
        .unwrap();
    *xml = String::from_utf8(xml.clone())
        .unwrap()
        .replace(
            "</Relationships>",
            &format!("{relationships}</Relationships>"),
        )
        .into_bytes();
    ooxml_opc::rezip_parts(&parts).unwrap()
}

fn open(paragraphs: &str, relationships: &str) -> DeckSession {
    DeckSession::open(&bytes(paragraphs, relationships), 846).unwrap()
}

fn rendered_at(session: &DeckSession, index: usize) -> Vec<TextParagraph> {
    fn paragraphs(primitive: Primitive) -> Vec<TextParagraph> {
        match primitive {
            Primitive::TextBox { paragraphs, .. } => paragraphs,
            Primitive::Table { primitives, .. } => {
                primitives.into_iter().flat_map(paragraphs).collect()
            }
            _ => Vec::new(),
        }
    }
    let mut renderer = SlideRenderer::new();
    renderer.register_font("Arial", false, false, FONT).unwrap();
    let full = renderer
        .layout_slide(session.package(), &session.snapshot().unwrap(), index)
        .unwrap();
    let scoped = renderer
        .layout_scoped_slide(session.package(), &session.slide_scope(index).unwrap())
        .unwrap();
    assert_eq!(full.display_list, scoped.display_list);
    full.display_list
        .primitives
        .into_iter()
        .flat_map(paragraphs)
        .collect()
}

fn rendered(session: &DeckSession) -> TextParagraph {
    rendered_at(session, 0).remove(0)
}

fn url(href: &str) -> TextLinkTarget {
    TextLinkTarget::Url {
        href: href.to_owned(),
    }
}

fn link(start: u32, end: u32, target: TextLinkTarget) -> TextLink {
    TextLink { start, end, target }
}

#[test]
fn external_relationship_targets_are_emitted_raw() {
    for href in ["https://example.test/docs", "javascript:alert(1)"] {
        let session = open(
            &format!("<a:p>{}</a:p>", run("Read more", r#"r:id="link""#)),
            &relationship("link", "hyperlink", href, true),
        );
        assert!(session.slide_scope(0).unwrap().slide_part_paths.is_empty());
        assert_eq!(rendered(&session).links, [link(0, 9, url(href))]);
    }
}

#[test]
fn internal_slide_relationships_and_slide_jump_actions_resolve_to_deck_indices() {
    for (kind, action) in [
        ("slide", ""),
        ("slide", r#"action="ppaction://hlinksldjump""#),
        ("hyperlink", r#"action="ppaction://hlinksldjump""#),
    ] {
        let session = open(
            &format!(
                "<a:p>{}</a:p>",
                run("Third slide", &format!(r#"r:id="link" {action}"#))
            ),
            &relationship("link", kind, "slide3.xml", false),
        );
        assert_eq!(
            rendered(&session).links,
            [link(0, 11, TextLinkTarget::Slide { index: 2 })]
        );
    }
}

#[test]
fn show_jumps_have_named_navigation_targets() {
    for (action, jump) in [
        ("nextslide", "next"),
        ("previousslide", "previous"),
        ("firstslide", "first"),
        ("lastslide", "last"),
    ] {
        let session = open(
            &format!(
                "<a:p>{}</a:p>",
                run(
                    "Go",
                    &format!(r#"action="ppaction://hlinkshowjump?jump={action}""#)
                )
            ),
            "",
        );
        assert!(session.slide_scope(0).unwrap().slide_part_paths.is_empty());
        assert_eq!(
            rendered(&session).links,
            [link(
                0,
                2,
                TextLinkTarget::Jump {
                    jump: jump.to_owned()
                }
            )]
        );
    }
}

#[test]
fn unresolved_unsupported_and_empty_targets_do_not_become_links() {
    let relationships = [
        relationship("notes", "hyperlink", "../media/image1.png", false),
        relationship("missing", "slide", "slide99.xml", false),
        relationship("empty", "hyperlink", "", true),
        relationship("url", "hyperlink", "https://example.test", true),
    ]
    .concat();
    for attributes in [
        r#"r:id="absent""#,
        r#"r:id="""#,
        r#"r:id="notes""#,
        r#"r:id="missing""#,
        r#"r:id="empty""#,
        r#"r:id="url" action="ppaction://hlinkfile""#,
        r#"action="ppaction://hlinkshowjump?jump=endshow""#,
        "",
    ] {
        let paragraph = rendered(&open(
            &format!("<a:p>{}</a:p>", run("Plain", attributes)),
            &relationships,
        ));
        assert!(paragraph.links.is_empty());
        let json = serde_json::to_value(&paragraph).unwrap();
        assert!(json.get("links").is_none());
        assert!(
            serde_json::from_value::<TextParagraph>(json)
                .unwrap()
                .links
                .is_empty()
        );
    }
}

#[test]
fn utf16_offsets_and_adjacent_equal_targets_survive_coalesced_runs() {
    let session = open(
        &format!(
            r#"<a:p><a:r><a:t>A😀 </a:t></a:r>{}{}<a:r><a:t> end</a:t></a:r></a:p>"#,
            run("Read ", r#"r:id="link""#),
            run("more😀", r#"r:id="alias""#)
        ),
        &[
            relationship("link", "hyperlink", "https://example.test", true),
            relationship("alias", "hyperlink", "https://example.test", true),
        ]
        .concat(),
    );
    assert_eq!(
        rendered(&session).links,
        [link(4, 15, url("https://example.test"))]
    );
}

#[test]
fn changed_paragraphs_do_not_borrow_source_links() {
    let session = open(
        &format!("<a:p>{}</a:p>", run("Read more", r#"r:id="link""#)),
        &relationship("link", "hyperlink", "https://example.test", true),
    );
    let story = session.snapshot().unwrap().slides[0].shapes[0].text_stories[0]
        .id
        .clone();
    session
        .insert_text(
            &EditCtx::local("test"),
            &story,
            0,
            "New ",
            &TextStyle::default(),
        )
        .unwrap();
    assert!(session.slide_scope(0).unwrap().slide_part_paths.is_empty());
    assert!(rendered(&session).links.is_empty());
}

#[test]
fn slide_number_fields_use_the_rendered_number_length() {
    for cached in ["9", ""] {
        let paragraph = format!(
            r#"<a:p><a:fld id="{{00000000-0000-0000-0000-000000000001}}" type="slidenum"><a:rPr><a:hlinkClick r:id="link"/></a:rPr><a:t>{cached}</a:t></a:fld>{}</a:p>"#,
            run(" more", r#"r:id="link""#)
        );
        assert_eq!(
            rendered(&open(
                &paragraph,
                &relationship("link", "hyperlink", "https://example.test", true)
            ))
            .links,
            [link(0, 7, url("https://example.test"))]
        );
    }
}

#[test]
fn slide_targets_follow_live_deck_order() {
    let session = open(
        &format!("<a:p>{}</a:p>", run("Third", r#"r:id="link""#)),
        &relationship("link", "slide", "slide3.xml", false),
    );
    let id = session.snapshot().unwrap().slides[2].id.clone();
    session.move_slide(&EditCtx::local("test"), &id, 0).unwrap();
    assert_eq!(
        rendered_at(&session, 1)[0].links,
        [link(0, 5, TextLinkTarget::Slide { index: 0 })]
    );
}

#[test]
fn inherited_links_use_the_layout_or_master_relationships() {
    let parent = shape(&format!(
        "<a:p>{}</a:p>",
        run("Inherited", r#"r:id="link""#)
    ))
    .replace(r#"<p:ph type="body"/>"#, "");
    let source = support::deck(&[("", false)], "", "", &parent, None);
    let mut parts = ooxml_opc::unzip_parts(&source).unwrap();
    for (path, xml) in &mut parts {
        let href = if path == "ppt/slideLayouts/_rels/slideLayout1.xml.rels" {
            "https://example.test/layout"
        } else if path == "ppt/slideMasters/_rels/slideMaster1.xml.rels" {
            "https://example.test/master"
        } else {
            continue;
        };
        *xml = String::from_utf8(xml.clone())
            .unwrap()
            .replace(
                "</Relationships>",
                &format!(
                    "{}</Relationships>",
                    relationship("link", "hyperlink", href, true)
                ),
            )
            .into_bytes();
    }
    let session = DeckSession::open(&ooxml_opc::rezip_parts(&parts).unwrap(), 847).unwrap();
    let paragraphs = rendered_at(&session, 0);
    assert_eq!(
        paragraphs
            .iter()
            .flat_map(|paragraph| paragraph.links.clone())
            .collect::<Vec<_>>(),
        [
            link(0, 9, url("https://example.test/master")),
            link(0, 9, url("https://example.test/layout"))
        ]
    );
}

#[test]
fn unchanged_hyperlink_actions_round_trip_without_part_changes() {
    let source = bytes(
        &format!(
            "<a:p>{}</a:p>",
            run("Go", r#"r:id="link" action="ppaction://hlinksldjump""#)
        ),
        &relationship("link", "slide", "slide3.xml", false),
    );
    let session = DeckSession::open(&source, 848).unwrap();
    let original: std::collections::BTreeMap<_, _> = ooxml_opc::unzip_parts(&source)
        .unwrap()
        .into_iter()
        .collect();
    let saved: std::collections::BTreeMap<_, _> = ooxml_opc::unzip_parts(&session.save().unwrap())
        .unwrap()
        .into_iter()
        .collect();
    assert_eq!(original, saved);
}

fn paragraph_text(paragraphs: &[TextParagraph]) -> String {
    paragraphs
        .iter()
        .flat_map(|paragraph| &paragraph.runs)
        .map(|run| run.text.as_str())
        .collect()
}

fn href_bytes(paragraphs: &[TextParagraph]) -> usize {
    paragraphs
        .iter()
        .flat_map(|paragraph| &paragraph.links)
        .map(|link| match &link.target {
            TextLinkTarget::Url { href } => href.len(),
            _ => 0,
        })
        .sum()
}

#[test]
fn long_https_paths_and_query_strings_remain_links() {
    let href = format!(
        "https://example.test/shared/{}report.pdf?trackingToken={}",
        "project-reports/".repeat(80),
        "AbCdEf0123456789".repeat(150),
    );
    assert!((3_000..4_000).contains(&href.encode_utf16().count()));
    let session = open(
        &format!("<a:p>{}</a:p>", run("Open report", r#"r:id="link""#)),
        &relationship("link", "hyperlink", &href, true),
    );
    let paragraph = rendered(&session);
    assert_eq!(paragraph.links, [link(0, 11, url(&href))]);
    assert_eq!(paragraph_text(&[paragraph]), "Open report");
}

#[test]
fn oversized_external_hrefs_keep_their_text_without_link_metadata() {
    let prefix = "https://example.test/";
    let over_limit = format!("{prefix}{}", "x".repeat(8_193 - prefix.len()));
    assert_eq!(over_limit.encode_utf16().count(), 8_193);
    for href in [over_limit, format!("{prefix}{}", "😀".repeat(4_096))] {
        assert!(href.encode_utf16().count() > 8_192);
        let session = open(
            &format!("<a:p>{}</a:p>", run("Read more", r#"r:id="link""#)),
            &relationship("link", "hyperlink", &href, true),
        );
        let paragraph = rendered(&session);
        assert!(paragraph.links.is_empty());
        assert_eq!(paragraph_text(&[paragraph]), "Read more");
    }
}

#[test]
fn distinct_links_are_kept_across_paragraphs() {
    let relationships = (0..6)
        .map(|index| {
            relationship(
                &format!("link{index}"),
                "hyperlink",
                &format!("https://example.test/{index}"),
                true,
            )
        })
        .collect::<String>();
    let runs = (0..6)
        .map(|index| run("x", &format!(r#"r:id="link{index}""#)))
        .collect::<Vec<_>>();
    let paragraphs = format!(
        "<a:p>{}</a:p><a:p>{}</a:p>",
        runs[..3].concat(),
        runs[3..].concat(),
    );
    let session = open(&paragraphs, &relationships);
    let paragraphs = rendered_at(&session, 0);
    assert_eq!(paragraph_text(&paragraphs), "x".repeat(6));
    assert_eq!(paragraphs[0].links.len(), 3);
    assert_eq!(paragraphs[1].links.len(), 3);
    assert_eq!(paragraphs[1].links.last().unwrap().end, 3);
    assert_eq!(
        paragraphs[1].links.last().unwrap().target,
        url("https://example.test/5")
    );
}

#[test]
fn href_bytes_are_budgeted_per_slide_and_later_spans_are_omitted() {
    let prefix = "https://example.test/";
    let href_a = format!("{prefix}{}a", "é".repeat(8_192 - prefix.len() - 1));
    let href_b = format!("{prefix}{}b", "é".repeat(8_192 - prefix.len() - 1));
    assert_eq!(href_a.encode_utf16().count(), 8_192);
    assert_eq!(href_a.len(), href_b.len());
    let runs = (0..256)
        .map(|index| {
            run(
                "x",
                if index % 2 == 0 {
                    r#"r:id="a""#
                } else {
                    r#"r:id="b""#
                },
            )
        })
        .collect::<String>();
    let session = open(
        &format!(
            "<a:p>{runs}</a:p><a:p>{}</a:p>",
            run("Later", r#"r:id="short""#)
        ),
        &[
            relationship("a", "hyperlink", &href_a, true),
            relationship("b", "hyperlink", &href_b, true),
            relationship("short", "hyperlink", "https://example.test", true),
        ]
        .concat(),
    );
    let paragraphs = rendered_at(&session, 0);
    let retained = (256 * 1024) / href_a.len();
    assert_eq!(paragraphs[0].links.len(), retained);
    assert!(paragraphs[1].links.is_empty());
    assert_eq!(href_bytes(&paragraphs), retained * href_a.len());
    assert!(href_bytes(&paragraphs) <= 256 * 1024);
    assert_eq!(
        paragraph_text(&paragraphs),
        format!("{}Later", "x".repeat(256))
    );
}

#[test]
fn alternating_maximum_length_hrefs_produce_bounded_metadata() {
    let prefix = "https://example.test/";
    let href_a = format!("{prefix}{}a", "x".repeat(8_192 - prefix.len() - 1));
    let href_b = format!("{prefix}{}b", "x".repeat(8_192 - prefix.len() - 1));
    let runs = (0..40)
        .map(|index| {
            run(
                "x",
                if index % 2 == 0 {
                    r#"r:id="a""#
                } else {
                    r#"r:id="b""#
                },
            )
        })
        .collect::<String>();
    let session = open(
        &format!("<a:p>{runs}</a:p>"),
        &[
            relationship("a", "hyperlink", &href_a, true),
            relationship("b", "hyperlink", &href_b, true),
        ]
        .concat(),
    );
    let paragraphs = rendered_at(&session, 0);
    assert_eq!(paragraphs[0].links.len(), 32);
    assert!(href_bytes(&paragraphs) <= 256 * 1024);
    assert_eq!(href_bytes(&paragraphs), 256 * 1024);
    assert_eq!(paragraph_text(&paragraphs), "x".repeat(40));
}

#[test]
fn adjacent_runs_charge_one_span_and_one_href_after_merging() {
    let prefix = "https://example.test/";
    let href = format!("{prefix}{}", "x".repeat(8_192 - prefix.len()));
    let runs = run("x", r#"r:id="link""#).repeat(8);
    let session = open(
        &format!("<a:p>{runs}</a:p>"),
        &relationship("link", "hyperlink", &href, true),
    );
    let paragraphs = rendered_at(&session, 0);
    assert_eq!(paragraphs[0].links, [link(0, 8, url(&href))]);
    assert_eq!(href_bytes(&paragraphs), 8_192);
    assert_eq!(paragraph_text(&paragraphs), "x".repeat(8));
}

#[test]
fn repeated_links_to_a_megabyte_padded_target_keep_text_without_links() {
    let href = format!("https://example.test/{}", " ".repeat(1024 * 1024));
    let runs = run("x", r#"r:id="link""#).repeat(8);
    let session = open(
        &format!("<a:p>{runs}</a:p>"),
        &relationship("link", "hyperlink", &href, true),
    );
    let paragraph = rendered(&session);
    assert!(paragraph.links.is_empty());
    assert_eq!(paragraph_text(&[paragraph]), "x".repeat(8));
}

#[test]
fn distinct_missing_relationship_ids_keep_text_without_links() {
    let runs = (0..16)
        .map(|index| run("x", &format!(r#"r:id="missing{index}""#)))
        .collect::<String>();
    let session = open(
        &format!("<a:p>{runs}</a:p>"),
        &relationship("present", "hyperlink", "https://example.test", true),
    );
    let paragraph = rendered(&session);
    assert!(paragraph.links.is_empty());
    assert_eq!(paragraph_text(&[paragraph]), "x".repeat(16));
}

#[test]
fn relationship_ids_over_256_bytes_are_omitted() {
    for length in [256, 257] {
        let id = "r".repeat(length);
        let session = open(
            &format!("<a:p>{}</a:p>", run("Read", &format!(r#"r:id="{id}""#))),
            &relationship(&id, "hyperlink", "https://example.test", true),
        );
        let paragraph = rendered(&session);
        if length == 256 {
            assert_eq!(paragraph.links, [link(0, 4, url("https://example.test"))]);
        } else {
            assert!(paragraph.links.is_empty());
        }
        assert_eq!(paragraph_text(&[paragraph]), "Read");
    }
}

#[test]
fn relationship_index_keeps_the_first_occurrence_of_duplicate_ids() {
    let session = open(
        &format!("<a:p>{}</a:p>", run("Read", r#"r:id="link""#)),
        &[
            relationship("link", "hyperlink", "https://example.test/first", true),
            relationship("link", "hyperlink", "https://example.test/second", true),
        ]
        .concat(),
    );
    assert_eq!(
        rendered(&session).links,
        [link(0, 4, url("https://example.test/first"))]
    );
}

#[test]
fn cached_internal_relationships_distinguish_slide_jump_actions() {
    let session = open(
        &format!(
            "<a:p>{}{}</a:p>",
            run("Plain", r#"r:id="link""#),
            run("Jump", r#"r:id="link" action="ppaction://hlinksldjump""#),
        ),
        &relationship("link", "hyperlink", "slide3.xml", false),
    );
    assert_eq!(
        rendered(&session).links,
        [link(5, 9, TextLinkTarget::Slide { index: 2 })]
    );
}

#[test]
fn link_work_exhaustion_preserves_earlier_links_and_all_later_text() {
    let unused = (0..65_600)
        .map(|index| format!(r#"<Relationship Id="u{index}" Type="h" Target="x"/>"#))
        .collect::<String>();
    let relationships = format!(
        "{}{}",
        relationship("link", "hyperlink", "https://example.test", true),
        unused
    );
    let session = open(
        &format!(
            "<a:p>{}</a:p><a:p>{}</a:p><a:p>{}</a:p>",
            run(
                "Early",
                r#"action="ppaction://hlinkshowjump?jump=nextslide""#
            ),
            run("Trigger", r#"r:id="link""#),
            run(
                "Later",
                r#"action="ppaction://hlinkshowjump?jump=nextslide""#
            ),
        ),
        &relationships,
    );
    let paragraphs = rendered_at(&session, 0);
    assert_eq!(
        paragraphs[0].links,
        [link(
            0,
            5,
            TextLinkTarget::Jump {
                jump: "next".to_owned()
            }
        )]
    );
    assert!(paragraphs[1].links.is_empty());
    assert!(paragraphs[2].links.is_empty());
    assert_eq!(paragraph_text(&paragraphs), "EarlyTriggerLater");
}
