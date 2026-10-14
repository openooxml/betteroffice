use docx_parse::{S9ParseOptions, S9WireEnvelope, parse_docx_s9_wire};

const W: &str = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const WP: &str = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing";
const A: &str = "http://schemas.openxmlformats.org/drawingml/2006/main";
const R: &str = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const W14: &str = "http://schemas.microsoft.com/office/word/2010/wordml";
const WPS: &str = "http://schemas.microsoft.com/office/word/2010/wordprocessingShape";
const MC: &str = "http://schemas.openxmlformats.org/markup-compatibility/2006";

fn story(root: &str, content: &str) -> String {
    format!(
        r#"<w:{root} xmlns:w="{W}" xmlns:wp="{WP}" xmlns:a="{A}" xmlns:r="{R}" xmlns:w14="{W14}" xmlns:wps="{WPS}" xmlns:mc="{MC}">{content}</w:{root}>"#
    )
}

fn package(body: &str, extra: &[(&str, String)]) -> S9WireEnvelope {
    let mut parts = vec![
        ("[Content_Types].xml".to_owned(), br#"<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="xml" ContentType="application/xml"/><Default Extension="PNG" ContentType="image/png"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/custom/covered.payload" ContentType="application/octet-stream"/></Types>"#.to_vec()),
        ("_rels/.rels".to_owned(), format!(r#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="doc" Type="{R}/officeDocument" Target="word/document.xml"/></Relationships>"#).into_bytes()),
        ("word/_rels/document.xml.rels".to_owned(), format!(r#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="header" Type="{R}/header" Target="header1.xml"/><Relationship Id="footer" Type="{R}/footer" Target="footer1.xml"/><Relationship Id="chart" Type="{R}/chart" Target="charts/missing.xml"/></Relationships>"#).into_bytes()),
        ("word/document.xml".to_owned(), story("document", &format!("<w:body>{body}</w:body>")).into_bytes()),
    ];
    parts.extend(
        extra
            .iter()
            .map(|(name, xml)| ((*name).to_owned(), xml.as_bytes().to_vec())),
    );
    parse_docx_s9_wire(
        &ooxml_opc::rezip_parts(&parts).unwrap(),
        S9ParseOptions::default(),
    )
    .unwrap()
}

fn image(id: &str) -> String {
    format!(
        r#"<w:r><w:drawing><wp:inline><wp:extent cx="914400" cy="914400"/><wp:docPr id="{id}" name="Image"/><a:graphic><a:graphicData/></a:graphic></wp:inline></w:drawing></w:r>"#
    )
}

fn shape(id: &str, text: Option<&str>) -> String {
    let text = text
        .map(|text| format!("<wps:txbx><w:txbxContent>{text}</w:txbxContent></wps:txbx>"))
        .unwrap_or_default();
    format!(
        r#"<w:r><w:drawing><wp:anchor><wp:extent cx="914400" cy="914400"/><wp:docPr id="{id}" name="Shape"/><a:graphic><a:graphicData><wps:wsp><wps:cNvPr id="1"/><wps:spPr><a:prstGeom prst="rect"/></wps:spPr>{text}<wps:bodyPr/></wps:wsp></a:graphicData></a:graphic></wp:anchor></w:drawing></w:r>"#
    )
}

fn bookmark(id: &str, name: &str) -> String {
    format!(r#"<w:bookmarkStart w:id="{id}" w:name="{name}"/><w:bookmarkEnd w:id="{id}"/>"#)
}

#[test]
fn clean_package_has_no_integrity_warnings() {
    let alternate = format!(
        r#"<w:r><mc:AlternateContent><mc:Choice Requires="wps">{}</mc:Choice><mc:Fallback>{}</mc:Fallback></mc:AlternateContent></w:r>"#,
        image("4").replace("<w:r>", "").replace("</w:r>", ""),
        image("4").replace("<w:r>", "").replace("</w:r>", "")
    );
    let body = format!(
        r#"<w:p w14:paraId="0000000A">{}{}{}{}{alternate}</w:p><w:p w14:paraId="0000000B">{}</w:p>"#,
        image("1"),
        shape("2", None),
        shape(
            "3",
            Some(&format!("<w:p>{}</w:p>", bookmark("2", "inside")))
        ),
        bookmark("1", "body"),
        bookmark("3", "next")
    );
    let parsed = package(
        &body,
        &[
            (
                "word/header1.xml",
                story(
                    "hdr",
                    &format!(
                        r#"<w:p w14:paraId="0000000C">{}{}</w:p>"#,
                        image("5"),
                        bookmark("4", "header")
                    ),
                ),
            ),
            ("custom/covered.payload", "covered".to_owned()),
            ("word/media/covered.PNG", "covered by default".to_owned()),
        ],
    );
    assert_eq!(parsed.document.warnings, None);
}

#[test]
fn uncovered_parts_are_counted_without_rejecting_the_package() {
    let parsed = package(
        "<w:p/>",
        &[
            ("custom/orphan.bin", "uncovered".to_owned()),
            ("custom/no-extension", "uncovered".to_owned()),
            ("custom/covered.payload", "covered by override".to_owned()),
            ("custom/covered.xml", "<root/>".to_owned()),
            ("word/media/covered.png", "covered by default".to_owned()),
        ],
    );
    assert_eq!(parsed.document.package.document.content.len(), 1);
    assert_eq!(
        parsed.document.warnings.unwrap(),
        ["DOCX contains 2 orphan OPC parts with no declared content type."]
    );
}

#[test]
fn drawing_duplicates_cover_images_shapes_text_boxes_and_opaque_drawings() {
    let body = format!(
        "<w:p>{}{}{}{}</w:p>",
        image("7"),
        shape("07", None),
        shape("7", Some("<w:p/>")),
        r#"<w:r><w:drawing><wp:inline><wp:docPr id="7"/><a:graphic><a:graphicData><c:chart xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" r:id="chart"/></a:graphicData></a:graphic></wp:inline></w:drawing></w:r>"#
    );
    let parsed = package(
        &body,
        &[
            (
                "word/header1.xml",
                story("hdr", &format!("<w:p>{}</w:p>", image("7"))),
            ),
            (
                "word/comments.xml",
                story(
                    "comments",
                    &format!(
                        r#"<w:comment w:id="0"><w:p>{}</w:p></w:comment>"#,
                        image("7")
                    ),
                ),
            ),
        ],
    );
    assert!(
        serde_json::to_string(&parsed.document.package.document.content)
            .unwrap()
            .contains("opaqueDrawing")
    );
    assert_eq!(
        parsed.document.warnings.unwrap(),
        ["DOCX contains 5 drawing doc properties with duplicate `wp:docPr id` values."]
    );
}

#[test]
fn bookmark_duplicates_cover_nested_content_and_other_stories() {
    let body = format!(
        r#"<w:p>{}<w:hyperlink w:anchor="first">{}</w:hyperlink><w:sdt><w:sdtContent>{}</w:sdtContent></w:sdt>{}</w:p>"#,
        bookmark("4", "first"),
        bookmark("04", "link"),
        bookmark("4", "control"),
        shape(
            "1",
            Some(&format!("<w:p>{}</w:p>", bookmark("4", "textBox")))
        )
    );
    let parsed = package(
        &body,
        &[
            (
                "word/header1.xml",
                story("hdr", &format!("<w:p>{}</w:p>", bookmark("4", "header"))),
            ),
            (
                "word/comments.xml",
                story(
                    "comments",
                    &format!(
                        r#"<w:comment w:id="0"><w:p>{}</w:p></w:comment>"#,
                        bookmark("4", "comment")
                    ),
                ),
            ),
        ],
    );
    assert_eq!(
        parsed.document.warnings.unwrap(),
        ["DOCX contains 5 bookmark starts with duplicate `w:bookmarkStart w:id` values."]
    );
}

#[test]
fn paragraph_duplicate_count_matches_flags_across_stories_and_separators() {
    let paragraph = r#"<w:p w14:paraId="0000000a"/>"#;
    let body = format!(
        r#"<w:p w14:paraId="0000000A"/><w:tbl><w:tr><w:tc>{paragraph}</w:tc></w:tr></w:tbl><w:sdt><w:sdtContent>{paragraph}</w:sdtContent></w:sdt>"#
    );
    let parsed = package(
        &body,
        &[
            ("word/header1.xml", story("hdr", paragraph)),
            ("word/footer1.xml", story("ftr", paragraph)),
            (
                "word/footnotes.xml",
                story(
                    "footnotes",
                    &format!(
                        r#"<w:footnote w:id="1">{paragraph}</w:footnote><w:footnote w:id="-1" w:type="separator">{paragraph}</w:footnote>"#
                    ),
                ),
            ),
            (
                "word/endnotes.xml",
                story(
                    "endnotes",
                    &format!(
                        r#"<w:endnote w:id="1">{paragraph}</w:endnote><w:endnote w:id="-1" w:type="separator">{paragraph}</w:endnote>"#
                    ),
                ),
            ),
        ],
    );
    let package = serde_json::to_value(&parsed.document.package).unwrap();
    fn flagged(value: &serde_json::Value) -> usize {
        match value {
            serde_json::Value::Array(values) => values.iter().map(flagged).sum(),
            serde_json::Value::Object(values) => {
                usize::from(values.get("repeatedParaId") == Some(&serde_json::Value::Bool(true)))
                    + values.values().map(flagged).sum::<usize>()
            }
            _ => 0,
        }
    }
    assert_eq!(flagged(&package), 8);
    assert_eq!(package["document"]["content"][0]["paraId"], "0000000A");
    assert_eq!(
        package["headerEntries"][0][1]["content"][0]["paraId"],
        "0000000a"
    );
    assert_eq!(
        parsed.document.warnings.unwrap(),
        [
            "BetterOffice found 8 paragraphs sharing a `w14:paraId` on open. A host addressing paragraphs by that raw Word ID should not assume it is unique; prefer a session anchor with `resolveParagraphAnchor()` or call `persistParagraphIds()` to repair the duplicates."
        ]
    );
}

#[test]
fn identifier_scan_resolves_namespaces_and_ignores_foreign_lookalikes() {
    let body = format!(
        r#"<w:p xmlns:d="{WP}" xmlns:b="{W}" xmlns:f="urn:foreign"><w:r><w:drawing><wp:inline><d:docPr id="2"/><a:graphic><a:graphicData/></a:graphic></wp:inline></w:drawing></w:r>{}<b:bookmarkStart b:id="2" b:name="alias"/><b:bookmarkEnd b:id="2"/><f:docPr id="2"/><f:bookmarkStart f:id="2"/></w:p>"#,
        bookmark("2", "original")
    );
    let parsed = package(
        &body,
        &[(
            "word/header1.xml",
            story("hdr", &format!("<w:p>{}</w:p>", image("2"))),
        )],
    );
    assert_eq!(
        parsed.document.warnings.unwrap(),
        [
            "DOCX contains 1 drawing doc properties with duplicate `wp:docPr id` values.",
            "DOCX contains 1 bookmark starts with duplicate `w:bookmarkStart w:id` values."
        ]
    );
}

#[test]
fn integrity_categories_are_additive_to_existing_warnings() {
    let paragraph = format!(
        r#"<w:p w14:paraId="00000001">{}{}</w:p>"#,
        image("1"),
        bookmark("1", "duplicate")
    );
    let smart_art = r#"<w:p><w:r><w:drawing><wp:inline><wp:docPr id="2"/><a:graphic><a:graphicData><dgm:relIds xmlns:dgm="http://schemas.openxmlformats.org/drawingml/2006/diagram" r:dm="missing"/></a:graphicData></a:graphic></wp:inline></w:drawing></w:r></w:p>"#;
    let parsed = package(
        &format!("{paragraph}{paragraph}{smart_art}"),
        &[("custom/orphan.bin", "uncovered".to_owned())],
    );
    let warnings = parsed.document.warnings.unwrap();
    assert_eq!(warnings.len(), 5, "{warnings:?}");
    assert!(warnings[0].contains("1 orphan OPC parts"));
    assert!(warnings[1].contains("1 drawing doc properties"));
    assert!(warnings[2].contains("1 bookmark starts"));
    assert!(warnings[3].starts_with("BetterOffice found 1 paragraphs sharing"));
    assert_eq!(
        warnings[4],
        "SmartArt diagram was skipped: no diagram data part was found."
    );
}
