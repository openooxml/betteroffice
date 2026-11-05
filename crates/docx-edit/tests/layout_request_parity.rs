use std::path::Path;

use docx_edit::layout_request::region_layout_request;
use serde_json::Value;

const WORD_NAMESPACE: &str =
    r#"xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main""#;
const RELATIONSHIPS: &str = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

fn theme_accent_docx() -> Vec<u8> {
    let parts = [
        (
            "[Content_Types].xml",
            r#"<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/theme/theme1.xml" ContentType="application/vnd.openxmlformats-officedocument.theme+xml"/><Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/></Types>"#,
        ),
        (
            "_rels/.rels",
            r#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>"#,
        ),
        (
            "word/document.xml",
            r#"<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml"><w:body><w:p w14:paraId="11111111"><w:r><w:rPr><w:rFonts w:ascii="Calibri" w:hAnsi="Calibri"/><w:b/><w:sz w:val="96"/><w:color w:val="000000" w:themeColor="accent1"/></w:rPr><w:t>Red theme accent</w:t></w:r></w:p><w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"/></w:sectPr></w:body></w:document>"#,
        ),
        (
            "word/_rels/document.xml.rels",
            r#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdTheme" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="theme/theme1.xml"/><Relationship Id="rIdStyles" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/></Relationships>"#,
        ),
        (
            "word/theme/theme1.xml",
            r#"<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="Red accent"><a:themeElements><a:clrScheme name="Red accent"><a:dk1><a:srgbClr val="000000"/></a:dk1><a:lt1><a:srgbClr val="FFFFFF"/></a:lt1><a:dk2><a:srgbClr val="44546A"/></a:dk2><a:lt2><a:srgbClr val="E7E6E6"/></a:lt2><a:accent1><a:srgbClr val="FF0000"/></a:accent1><a:accent2><a:srgbClr val="ED7D31"/></a:accent2><a:accent3><a:srgbClr val="A5A5A5"/></a:accent3><a:accent4><a:srgbClr val="FFC000"/></a:accent4><a:accent5><a:srgbClr val="5B9BD5"/></a:accent5><a:accent6><a:srgbClr val="70AD47"/></a:accent6><a:hlink><a:srgbClr val="0563C1"/></a:hlink><a:folHlink><a:srgbClr val="954F72"/></a:folHlink></a:clrScheme><a:fontScheme name="Office"><a:majorFont><a:latin typeface="Calibri"/></a:majorFont><a:minorFont><a:latin typeface="Calibri"/></a:minorFont></a:fontScheme><a:fmtScheme name="Office"><a:fillStyleLst/><a:lnStyleLst/><a:effectStyleLst/><a:bgFillStyleLst/></a:fmtScheme></a:themeElements></a:theme>"#,
        ),
        (
            "word/styles.xml",
            r#"<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:docDefaults><w:rPrDefault><w:rPr><w:sz w:val="24"/><w:lang w:val="en-US" w:eastAsia="ja-JP"/></w:rPr></w:rPrDefault></w:docDefaults><w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style><w:style w:type="paragraph" w:styleId="TOC1"><w:name w:val="TOC 1"/><w:basedOn w:val="Normal"/></w:style><w:style w:type="character" w:styleId="Hyperlink"><w:name w:val="Hyperlink"/><w:rPr><w:sz w:val="28"/></w:rPr></w:style></w:styles>"#,
        ),
    ]
    .into_iter()
    .map(|(name, xml)| (name.to_owned(), xml.as_bytes().to_vec()))
    .collect::<Vec<_>>();
    ooxml_opc::rezip_parts(&parts).unwrap()
}

fn request_edges_docx() -> Vec<u8> {
    let parts = [
        (
            "[Content_Types].xml",
            r#"<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/><Override PartName="/word/settings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/><Override PartName="/word/footnotes.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml"/><Override PartName="/word/endnotes.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.endnotes+xml"/><Override PartName="/word/header1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"/></Types>"#,
        ),
        (
            "_rels/.rels",
            r#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="REL/officeDocument" Target="word/document.xml"/></Relationships>"#,
        ),
        (
            "word/_rels/document.xml.rels",
            r#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdStyles" Type="REL/styles" Target="styles.xml"/><Relationship Id="rIdSettings" Type="REL/settings" Target="settings.xml"/><Relationship Id="rIdFoot" Type="REL/footnotes" Target="footnotes.xml"/><Relationship Id="rIdEnd" Type="REL/endnotes" Target="endnotes.xml"/><Relationship Id="rIdHeader" Type="REL/header" Target="header1.xml"/></Relationships>"#,
        ),
        (
            "word/document.xml",
            r#"<w:document WNS xmlns:r="REL"><w:body><w:p><w:pPr><w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr></w:pPr><w:r><w:t>First section</w:t></w:r></w:p><w:p><w:r><w:t>Noted</w:t></w:r><w:r><w:footnoteReference w:id="2"/></w:r><w:r><w:endnoteReference w:id="3"/></w:r></w:p><w:sectPr><w:headerReference w:type="default" r:id="rIdHeader"/><w:pgSz w:w="15840" w:h="12240" w:orient="landscape"/></w:sectPr></w:body></w:document>"#,
        ),
        (
            "word/header1.xml",
            r##"<w:hdr WNS xmlns:v="urn:schemas-microsoft-com:vml"><w:p><w:r><w:pict><v:shape id="PowerPlusWaterMarkObject1" type="#_x0000_t136" style="rotation:315" fillcolor="C0C0C0"><v:fill opacity=".5"/><v:textpath string="DRAFT"/></v:shape></w:pict></w:r></w:p></w:hdr>"##,
        ),
        (
            "word/footnotes.xml",
            r#"<w:footnotes WNS><w:footnote w:type="separator" w:id="0"><w:p/></w:footnote><w:footnote w:type="continuationSeparator" w:id="1"><w:p/></w:footnote><w:footnote w:id="2"><w:p><w:r><w:t>Foot</w:t></w:r></w:p></w:footnote></w:footnotes>"#,
        ),
        (
            "word/endnotes.xml",
            r#"<w:endnotes WNS><w:endnote w:type="separator" w:id="0"><w:p/></w:endnote><w:endnote w:id="3"><w:p><w:r><w:t>End</w:t></w:r></w:p></w:endnote></w:endnotes>"#,
        ),
        (
            "word/settings.xml",
            r#"<w:settings WNS><w:defaultTabStop w:val="360"/><w:compat><w:noLeading/><w:doNotExpandShiftReturn/><w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="14"/></w:compat></w:settings>"#,
        ),
        (
            "word/styles.xml",
            r#"<w:styles WNS><w:docDefaults><w:rPrDefault><w:rPr><w:lang w:val="de-DE" w:bidi=""/></w:rPr></w:rPrDefault></w:docDefaults><w:style w:type="paragraph" w:styleId="Normal"><w:name w:val="Normal"/></w:style><w:style w:type="paragraph" w:styleId="Loop1"><w:name w:val="Loop 1"/><w:basedOn w:val="Loop2"/></w:style><w:style w:type="paragraph" w:styleId="Loop2"><w:name w:val="Loop 2"/><w:basedOn w:val="Loop1"/></w:style><w:style w:type="paragraph" w:styleId="Child"><w:name w:val="Child"/><w:basedOn w:val="Sized"/></w:style><w:style w:type="paragraph" w:styleId="Sized"><w:name w:val="Sized"/><w:basedOn w:val="Missing"/><w:rPr><w:sz w:val="31"/></w:rPr></w:style><w:style w:type="paragraph" w:styleId="Orphan"><w:name w:val="Orphan"/><w:basedOn w:val="Missing"/></w:style><w:style w:type="paragraph" w:styleId="Contents2"><w:name w:val="toc 2"/></w:style><w:style w:type="paragraph" w:styleId="TOC3"><w:name w:val="Third level"/></w:style><w:style w:type="character" w:styleId="Link"><w:name w:val="FollowedHyperlink"/><w:basedOn w:val="LinkBase"/></w:style><w:style w:type="character" w:styleId="LinkBase"><w:name w:val="Link base"/><w:rPr><w:sz w:val="18"/></w:rPr></w:style><w:style w:type="character" w:styleId="Hyperlink"><w:name w:val="Hyperlink"/></w:style></w:styles>"#,
        ),
    ]
    .into_iter()
    .map(|(name, xml)| {
        let xml = xml.replace("WNS", WORD_NAMESPACE).replace("REL", RELATIONSHIPS);
        (name.to_owned(), xml.into_bytes())
    })
    .collect::<Vec<_>>();
    ooxml_opc::rezip_parts(&parts).unwrap()
}

#[test]
fn layout_requests_match_goldens() {
    let root = Path::new(env!("CARGO_MANIFEST_DIR")).join("../..");
    let fixtures = root.join("crates/docx-edit/tests/fixtures/layout-request");
    let update = std::env::var("UPDATE_LAYOUT_REQUEST_GOLDENS").as_deref() == Ok("1");
    if update {
        std::fs::write(fixtures.join("theme-accent.docx"), theme_accent_docx()).unwrap();
        std::fs::write(fixtures.join("request-edges.docx"), request_edges_docx()).unwrap();
    }
    let cases: std::collections::BTreeMap<String, String> =
        serde_json::from_slice(&std::fs::read(fixtures.join("cases.json")).unwrap()).unwrap();
    for (name, source) in &cases {
        let input = root.join(source);
        let bytes = std::fs::read(&input).unwrap_or_else(|error| {
            panic!(
                "{}: {error}; generate with UPDATE_LAYOUT_REQUEST_GOLDENS=1",
                input.display()
            )
        });
        let parsed = docx_parse::parse_docx_s9_wire(
            &bytes,
            docx_parse::S9ParseOptions {
                parse_headers_footers: true,
                parse_notes: true,
                detect_variables: false,
                include_canonical: false,
                ..Default::default()
            },
        )
        .unwrap();
        let package = serde_json::to_value(parsed.document.package).unwrap();
        let request = region_layout_request(&package, 24.0);
        serde_json::from_value::<docx_layout::regions::RegionLayoutInput>(request.clone())
            .unwrap_or_else(|error| panic!("{name}: layout cannot read the request: {error}"));
        let golden = fixtures.join(format!("{name}.request.json"));
        if update {
            std::fs::write(
                &golden,
                format!("{}\n", serde_json::to_string_pretty(&request).unwrap()),
            )
            .unwrap();
        } else {
            let expected: Value = serde_json::from_slice(&std::fs::read(&golden).unwrap()).unwrap();
            assert_eq!(request, expected, "{name}");
        }
    }
}
