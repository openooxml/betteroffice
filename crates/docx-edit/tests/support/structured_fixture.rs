//! The synthetic package the structured-export tests read: headings from every source, nested
//! and overridden numbering, a merged table, controls, revisions, notes, comments, an image, fields
//! and opaque content.

use std::path::PathBuf;

pub const NS: &str = concat!(
    r#"xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" "#,
    r#"xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" "#,
    r#"xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml" "#,
    r#"xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" "#,
    r#"xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" "#,
    r#"xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture" "#,
    r#"xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape" "#,
    r#"xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" "#,
    r#"xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math" "#,
    r#"xmlns:bofx="urn:fidelity""#
);

const REL: &str = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const OFFICE: &str = "application/vnd.openxmlformats-officedocument.wordprocessingml";

pub fn fixture_dir() -> PathBuf {
    PathBuf::from(env!("CARGO_MANIFEST_DIR")).join("tests/fixtures/structured-export")
}

fn p(id: &str, content: &str) -> String {
    format!(r#"<w:p w14:paraId="{id}">{content}</w:p>"#)
}

fn styled(id: &str, style: &str, content: &str) -> String {
    format!(r#"<w:p w14:paraId="{id}"><w:pPr><w:pStyle w:val="{style}"/></w:pPr>{content}</w:p>"#)
}

fn numbered(id: &str, style: Option<&str>, num: u32, level: u32, content: &str) -> String {
    let style = style.map_or_else(String::new, |style| {
        format!(r#"<w:pStyle w:val="{style}"/>"#)
    });
    format!(
        r#"<w:p w14:paraId="{id}"><w:pPr>{style}<w:numPr><w:ilvl w:val="{level}"/><w:numId w:val="{num}"/></w:numPr></w:pPr>{content}</w:p>"#
    )
}

fn r(text: &str) -> String {
    format!(r#"<w:r><w:t xml:space="preserve">{text}</w:t></w:r>"#)
}

fn cell(properties: &str, content: &str) -> String {
    format!(r#"<w:tc><w:tcPr>{properties}</w:tcPr>{content}</w:tc>"#)
}

fn styles() -> String {
    let style = |kind: &str, id: &str, extra: &str| {
        format!(
            r#"<w:style w:type="{kind}" w:styleId="{id}"><w:name w:val="{id}"/>{extra}</w:style>"#
        )
    };
    format!(
        r#"<w:styles {NS}>{}{}{}{}{}{}</w:styles>"#,
        r#"<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style>"#,
        style(
            "paragraph",
            "Heading1",
            r#"<w:pPr><w:keepNext/><w:outlineLvl w:val="0"/></w:pPr><w:rPr><w:b/></w:rPr>"#
        ),
        style("paragraph", "Heading2", ""),
        style(
            "paragraph",
            "Heading3",
            r#"<w:pPr><w:outlineLvl w:val="9"/></w:pPr>"#
        ),
        style(
            "paragraph",
            "Title",
            r#"<w:basedOn w:val="Heading1"/><w:pPr><w:jc w:val="center"/></w:pPr>"#
        ),
        style(
            "character",
            "Hyperlink",
            r#"<w:rPr><w:u w:val="single"/></w:rPr>"#
        ),
    )
}

fn numbering() -> String {
    let level = |ilvl: u32, format: &str, text: &str| {
        format!(
            r#"<w:lvl w:ilvl="{ilvl}"><w:start w:val="1"/><w:numFmt w:val="{format}"/><w:lvlText w:val="{text}"/></w:lvl>"#
        )
    };
    format!(
        r#"<w:numbering {NS}><w:abstractNum w:abstractNumId="0">{}{}{}</w:abstractNum><w:abstractNum w:abstractNumId="1">{}</w:abstractNum><w:abstractNum w:abstractNumId="2">{}</w:abstractNum><w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num><w:num w:numId="2"><w:abstractNumId w:val="0"/><w:lvlOverride w:ilvl="0"><w:startOverride w:val="7"/></w:lvlOverride></w:num><w:num w:numId="3"><w:abstractNumId w:val="1"/></w:num><w:num w:numId="4"><w:abstractNumId w:val="2"/></w:num></w:numbering>"#,
        level(0, "decimal", "%1."),
        level(1, "lowerLetter", "%2)"),
        level(2, "bullet", "\u{2022}"),
        level(0, "upperRoman", "%1."),
        level(0, "cardinalText", "%1"),
    )
}

fn body() -> String {
    let image = r#"<w:r><w:drawing><wp:inline><wp:extent cx="914400" cy="457200"/><wp:docPr id="1" name="logo" descr="Company logo"/><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:pic><pic:blipFill><a:blip r:embed="rIdImage"/></pic:blipFill></pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing></w:r>"#;
    let inline_control = |tag: &str, id: u32, text: &str| {
        format!(
            r#"<w:sdt><w:sdtPr><w:alias w:val="Name"/><w:tag w:val="{tag}"/><w:id w:val="{id}"/><w:text/></w:sdtPr><w:sdtContent>{}</w:sdtContent></w:sdt>"#,
            r(text)
        )
    };
    let table = format!(
        r#"<w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/></w:tblPr><w:tblGrid><w:gridCol w:w="2000"/><w:gridCol w:w="2000"/><w:gridCol w:w="2000"/></w:tblGrid><w:tr><w:trPr><w:tblHeader/></w:trPr>{}{}</w:tr><w:tr>{}{}{}</w:tr><w:tr>{}{}{}</w:tr><w:tr><w:trPr><w:gridBefore w:val="1"/></w:trPr>{}</w:tr></w:tbl>"#,
        cell(
            r#"<w:gridSpan w:val="2"/>"#,
            &p("10000001", &r("Wide header"))
        ),
        cell("", &p("10000002", &r("Right header"))),
        cell(
            r#"<w:vMerge w:val="restart"/>"#,
            &p("10000003", &r("Merged"))
        ),
        cell("", &p("10000004", &r("Middle"))),
        cell(
            "",
            &format!(
                "{}{}",
                p("10000005", &r("Nested:")),
                r#"<w:tbl><w:tblGrid><w:gridCol w:w="900"/></w:tblGrid><w:tr><w:tc><w:p w14:paraId="10000006"><w:r><w:t>Inner</w:t></w:r></w:p></w:tc></w:tr></w:tbl><w:p w14:paraId="10000007"/>"#
            )
        ),
        cell(r#"<w:vMerge/>"#, &p("10000008", &r("Continuation text"))),
        cell("", &p("10000009", &r("Lower middle"))),
        cell("", &p("1000000A", &r("Lower right"))),
        cell(
            r#"<w:gridSpan w:val="2"/>"#,
            &p("1000000B", &r("After a skipped column"))
        ),
    );
    [
        styled("00000001", "Heading1", &r("Structured export")),
        format!(
            r#"<w:p w14:paraId="00000002"><w:pPr><w:outlineLvl w:val="1"/></w:pPr>{}</w:p>"#,
            r("Direct outline")
        ),
        styled("00000003", "Title", &r("Inherited from Heading1")),
        styled("00000004", "Heading2", &r("Built-in style fallback")),
        styled("00000005", "Heading3", &r("Explicit body level")),
        numbered("00000006", None, 1, 0, &r("First")),
        numbered("00000007", None, 1, 1, &r("Nested")),
        numbered("00000008", None, 1, 2, &r("Bullet")),
        numbered("00000009", None, 1, 0, &r("Second")),
        numbered("0000000A", None, 2, 0, &r("Overridden start")),
        numbered("0000000B", Some("Heading1"), 3, 0, &r("Numbered heading")),
        numbered("0000000C", None, 4, 0, &r("Unrendered format")),
        table,
        format!(
            r#"<w:sdt><w:sdtPr><w:alias w:val="Clause"/><w:tag w:val="clause"/><w:id w:val="-501"/><w:showingPlcHdr/></w:sdtPr><w:sdtContent>{}</w:sdtContent></w:sdt>"#,
            p("0000000D", &r("Inside a block control"))
        ),
        p(
            "0000000E",
            &format!(
                "{}{}{}{}",
                r("Dear "),
                inline_control("name", 11, "Ada"),
                r(" and "),
                inline_control("name", 12, "Grace")
            ),
        ),
        p(
            "0000000F",
            &format!(
                r#"{}<w:ins w:id="21" w:author="Ann" w:date="2026-01-02T03:04:05Z">{}</w:ins><w:del w:id="22" w:author="Bob" w:date="2026-01-03T00:00:00Z"><w:r><w:delText xml:space="preserve">removed </w:delText></w:r></w:del>{}"#,
                r("Keep "),
                r("added "),
                r("end")
            ),
        ),
        p("00000010", &format!("{}{image}", r("Logo: "))),
        p(
            "00000011",
            &format!(
                r#"{}<w:fldSimple w:instr=" PAGE "><w:r><w:t>4</w:t></w:r></w:fldSimple>{}<w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText xml:space="preserve"> DATE \@ "yyyy" </w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r><w:r><w:t>2026</w:t></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r>{}<w:fldSimple w:instr=" AUTHOR "/>"#,
                r("Page "),
                r(", year "),
                r(", author ")
            ),
        ),
        p(
            "00000012",
            &format!(
                r#"<w:commentRangeStart w:id="1"/>{}<w:commentRangeEnd w:id="1"/><w:r><w:commentReference w:id="1"/></w:r>{}<w:r><w:rPr><w:vertAlign w:val="superscript"/></w:rPr><w:footnoteReference w:id="1"/></w:r>"#,
                r("Commented words"),
                r(" with a note")
            ),
        ),
        p(
            "00000013",
            &format!(
                r#"{}<w:hyperlink r:id="rIdLink" w:tooltip="Visit"><w:r><w:rPr><w:rStyle w:val="Hyperlink"/></w:rPr><w:t>a link</w:t></w:r></w:hyperlink>"#,
                r("See ")
            ),
        ),
        p(
            "00000014",
            &format!(
                "{}<w:r><w:rPr><w:i/></w:rPr><w:tab/><w:t>tabbed</w:t><w:br/><w:t xml:space=\"preserve\">e\u{0301} \u{1F600} \u{FFFC} </w:t></w:r>{}",
                r("Atoms:"),
                r("done")
            ),
        ),
        r#"<bofx:block bofx:value="opaque"><bofx:child>kept</bofx:child></bofx:block>"#.to_owned(),
        p(
            "00000015",
            &format!(
                r#"{}<w:r><w:drawing><wp:inline><wp:extent cx="914400" cy="457200"/><wp:docPr id="2" name="Arrow" descr="Process arrow"/><a:graphic><a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape"><wps:wsp><wps:spPr><a:prstGeom prst="rightArrow"/></wps:spPr><wps:bodyPr/></wps:wsp></a:graphicData></a:graphic></wp:inline></w:drawing></w:r>{}<m:oMath><m:r><m:t>x=1</m:t></m:r></m:oMath>{}<w:r><w:drawing><wp:inline><wp:extent cx="914400" cy="457200"/><wp:docPr id="3" name="Chart"/><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart"><c:chart r:id="rIdChart"/></a:graphicData></a:graphic></wp:inline></w:drawing></w:r>{}"#,
                r("Shape "),
                r(", equation "),
                r(", unread chart "),
                r("end")
            ),
        ),
        p(
            "00000018",
            &format!(
                r#"{}<w:moveFrom w:id="71" w:author="Cy" w:date="2026-01-06T00:00:00Z"><w:r><w:delText>from here</w:delText></w:r></w:moveFrom><w:moveTo w:id="72" w:author="Cy" w:date="2026-01-06T00:00:00Z"><w:r><w:t>to here</w:t></w:r></w:moveTo>"#,
                r("Moved ")
            ),
        ),
        p(
            "00000019",
            &format!(
                r#"{}<w:r><w:br w:type="page"/></w:r>{}<w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText xml:space="preserve"> REF _Ref1 \h </w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r><w:r><w:rPr><w:b/></w:rPr><w:t>Bold</w:t></w:r><w:r><w:tab/><w:t>result</w:t></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r>"#,
                r("Page"),
                r("break ")
            ),
        ),
        format!(
            r#"<w:p w14:paraId="00000016"><w:pPr><w:sectPr><w:headerReference w:type="default" r:id="rIdHeader1"/><w:footerReference w:type="default" r:id="rIdFooter1"/><w:type w:val="continuous"/></w:sectPr></w:pPr>{}</w:p>"#,
            r("End of section one")
        ),
        p("00000017", &r("Section two")),
    ]
    .concat()
}

/// The fixture's parts, in archive order.
pub fn principal_parts() -> Vec<(String, Vec<u8>)> {
    let rel = |id: &str, kind: &str, target: &str| {
        format!(r#"<Relationship Id="{id}" Type="{REL}/{kind}" Target="{target}"/>"#)
    };
    let content_types = format!(
        r#"<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Default Extension="png" ContentType="image/png"/><Override PartName="/word/document.xml" ContentType="{OFFICE}.document.main+xml"/><Override PartName="/word/styles.xml" ContentType="{OFFICE}.styles+xml"/><Override PartName="/word/numbering.xml" ContentType="{OFFICE}.numbering+xml"/><Override PartName="/word/header1.xml" ContentType="{OFFICE}.header+xml"/><Override PartName="/word/header2.xml" ContentType="{OFFICE}.header+xml"/><Override PartName="/word/footer1.xml" ContentType="{OFFICE}.footer+xml"/><Override PartName="/word/footnotes.xml" ContentType="{OFFICE}.footnotes+xml"/><Override PartName="/word/endnotes.xml" ContentType="{OFFICE}.endnotes+xml"/><Override PartName="/word/comments.xml" ContentType="{OFFICE}.comments+xml"/></Types>"#
    );
    let root_rels = format!(
        r#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">{}</Relationships>"#,
        rel("rId1", "officeDocument", "word/document.xml")
    );
    let document_rels = format!(
        r#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">{}{}{}{}{}{}{}{}{}{}<Relationship Id="rIdLink" Type="{REL}/hyperlink" Target="https://example.com/docs" TargetMode="External"/></Relationships>"#,
        rel("rIdStyles", "styles", "styles.xml"),
        rel("rIdNumbering", "numbering", "numbering.xml"),
        rel("rIdHeader1", "header", "header1.xml"),
        rel("rIdHeader2", "header", "header2.xml"),
        rel("rIdFooter1", "footer", "footer1.xml"),
        rel("rIdFootnotes", "footnotes", "footnotes.xml"),
        rel("rIdEndnotes", "endnotes", "endnotes.xml"),
        rel("rIdComments", "comments", "comments.xml"),
        rel("rIdImage", "image", "media/logo.png"),
        rel("rIdChart", "chart", "charts/chart1.xml"),
    );
    let document = format!(
        r#"<w:document {NS}><w:body>{}<w:sectPr><w:headerReference w:type="default" r:id="rIdHeader1"/><w:headerReference w:type="first" r:id="rIdHeader2"/><w:titlePg/></w:sectPr></w:body></w:document>"#,
        body()
    );
    let story = |root: &str, id: &str, text: &str| {
        format!(r#"<w:{root} {NS}>{}</w:{root}>"#, p(id, &r(text)))
    };
    let notes = |root: &str, note: &str, id: &str, text: &str| {
        format!(
            r#"<w:{root} {NS}><w:{note} w:type="separator" w:id="-1"><w:p><w:r><w:separator/></w:r></w:p></w:{note}><w:{note} w:type="continuationSeparator" w:id="0"><w:p><w:r><w:continuationSeparator/></w:r></w:p></w:{note}><w:{note} w:id="1">{}</w:{note}></w:{root}>"#,
            p(id, &r(text))
        )
    };
    let comments = format!(
        r#"<w:comments {NS}><w:comment w:id="1" w:author="Ann" w:date="2026-01-04T00:00:00Z" w:initials="A">{}</w:comment><w:comment w:id="2" w:author="Bob" w:date="2026-01-05T00:00:00Z">{}</w:comment></w:comments>"#,
        p("30000001", &r("Please review")),
        p("30000002", &r("Unanchored remark"))
    );
    [
        ("[Content_Types].xml", content_types.into_bytes()),
        ("_rels/.rels", root_rels.into_bytes()),
        ("word/_rels/document.xml.rels", document_rels.into_bytes()),
        ("word/document.xml", document.into_bytes()),
        ("word/styles.xml", styles().into_bytes()),
        ("word/numbering.xml", numbering().into_bytes()),
        (
            "word/header1.xml",
            story("hdr", "20000001", "Default header").into_bytes(),
        ),
        (
            "word/header2.xml",
            story("hdr", "20000002", "First page header").into_bytes(),
        ),
        (
            "word/footer1.xml",
            story("ftr", "20000003", "Footer text").into_bytes(),
        ),
        (
            "word/footnotes.xml",
            notes("footnotes", "footnote", "20000004", "Footnote text").into_bytes(),
        ),
        (
            "word/endnotes.xml",
            notes("endnotes", "endnote", "20000005", "Unreferenced endnote").into_bytes(),
        ),
        ("word/comments.xml", comments.into_bytes()),
        (
            "word/media/logo.png",
            vec![137, 80, 78, 71, 13, 10, 26, 10, 0, 0, 0, 0],
        ),
    ]
    .into_iter()
    .map(|(name, bytes)| (name.to_owned(), bytes))
    .collect()
}

/// The committed fixture, rewritten from [`principal_parts`] when `STRUCTURED_EXPORT_UPDATE=1`.
pub fn principal_docx() -> Vec<u8> {
    let path = fixture_dir().join("principal.docx");
    if std::env::var("STRUCTURED_EXPORT_UPDATE").as_deref() == Ok("1") {
        let bytes = ooxml_opc::rezip_parts(&principal_parts()).unwrap();
        std::fs::create_dir_all(fixture_dir()).unwrap();
        static STAGED: std::sync::atomic::AtomicUsize = std::sync::atomic::AtomicUsize::new(0);
        let staged = path.with_extension(format!(
            "{}.{}.tmp",
            std::process::id(),
            STAGED.fetch_add(1, std::sync::atomic::Ordering::Relaxed)
        ));
        std::fs::write(&staged, &bytes).unwrap();
        std::fs::rename(&staged, &path).unwrap();
        return bytes;
    }
    std::fs::read(&path).unwrap()
}

/// Compares `actual` with the committed golden file `name`, rewriting it when
/// `STRUCTURED_EXPORT_UPDATE=1`.
pub fn golden(name: &str, actual: &str) {
    let path = fixture_dir().join(name);
    if std::env::var("STRUCTURED_EXPORT_UPDATE").as_deref() == Ok("1") {
        std::fs::create_dir_all(fixture_dir()).unwrap();
        std::fs::write(&path, actual).unwrap();
    }
    let expected = std::fs::read_to_string(&path)
        .unwrap_or_else(|error| panic!("golden file {name} is missing: {error}"));
    assert_eq!(actual, expected, "golden file {name} differs");
}

/// A small package around a body, for tests that need one feature at a time.
pub struct Package {
    body: String,
    parts: Vec<(String, String, Option<&'static str>)>,
    rels: Vec<String>,
    part_rels: Vec<(String, String)>,
}

impl Package {
    pub fn new(body: &str) -> Self {
        Self {
            body: body.to_owned(),
            parts: Vec::new(),
            rels: Vec::new(),
            part_rels: Vec::new(),
        }
    }

    /// Adds `word/{name}`, related from the main document as `id` of relationship `kind` when
    /// `kind` is not empty.
    pub fn part(
        mut self,
        name: &str,
        id: &str,
        kind: &str,
        content_type: &'static str,
        xml: &str,
    ) -> Self {
        self.parts
            .push((format!("word/{name}"), xml.to_owned(), Some(content_type)));
        if !kind.is_empty() {
            self.rels.push(format!(
                r#"<Relationship Id="{id}" Type="{REL}/{kind}" Target="{name}"/>"#
            ));
        }
        self
    }

    pub fn styles(self, xml: &str) -> Self {
        self.part("styles.xml", "rIdStyles", "styles", "styles", xml)
    }

    pub fn numbering(self, xml: &str) -> Self {
        self.part(
            "numbering.xml",
            "rIdNumbering",
            "numbering",
            "numbering",
            xml,
        )
    }

    /// Adds a relationship of the main document.
    pub fn rel(mut self, id: &str, kind: &str, target: &str) -> Self {
        self.rels.push(format!(
            r#"<Relationship Id="{id}" Type="{REL}/{kind}" Target="{target}"/>"#
        ));
        self
    }

    /// Adds a relationship of `word/{part}`.
    pub fn part_rel(mut self, part: &str, id: &str, kind: &str, target: &str) -> Self {
        self.part_rels.push((
            part.to_owned(),
            format!(r#"<Relationship Id="{id}" Type="{REL}/{kind}" Target="{target}"/>"#),
        ));
        self
    }

    pub fn parts(&self) -> Vec<(String, Vec<u8>)> {
        let overrides: String = self
            .parts
            .iter()
            .filter_map(|(name, _, kind)| {
                kind.map(|kind| {
                    format!(r#"<Override PartName="/{name}" ContentType="{OFFICE}.{kind}+xml"/>"#)
                })
            })
            .collect();
        let content_types = format!(
            r#"<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Default Extension="png" ContentType="image/png"/><Override PartName="/word/document.xml" ContentType="{OFFICE}.document.main+xml"/>{overrides}</Types>"#
        );
        let relationships = |entries: &str| {
            format!(
                r#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">{entries}</Relationships>"#
            )
        };
        let mut parts = vec![
            ("[Content_Types].xml".to_owned(), content_types),
            (
                "_rels/.rels".to_owned(),
                relationships(&format!(
                    r#"<Relationship Id="rId1" Type="{REL}/officeDocument" Target="word/document.xml"/>"#
                )),
            ),
            (
                "word/_rels/document.xml.rels".to_owned(),
                relationships(&self.rels.concat()),
            ),
            (
                "word/document.xml".to_owned(),
                format!(
                    r#"<w:document {NS}><w:body>{}</w:body></w:document>"#,
                    self.body
                ),
            ),
        ];
        for (name, xml, _) in &self.parts {
            parts.push((name.clone(), xml.clone()));
        }
        let mut owners: Vec<&String> = self.part_rels.iter().map(|(part, _)| part).collect();
        owners.dedup();
        for owner in owners {
            let entries: String = self
                .part_rels
                .iter()
                .filter(|(part, _)| part == owner)
                .map(|(_, entry)| entry.as_str())
                .collect();
            parts.push((format!("word/_rels/{owner}.rels"), relationships(&entries)));
        }
        let mut parts: Vec<(String, Vec<u8>)> = parts
            .into_iter()
            .map(|(name, xml)| (name, xml.into_bytes()))
            .collect();
        parts.push((
            "word/media/image1.png".to_owned(),
            vec![137, 80, 78, 71, 13, 10, 26, 10, 0, 0, 0, 0],
        ));
        parts
    }

    pub fn bytes(&self) -> Vec<u8> {
        ooxml_opc::rezip_parts(&self.parts()).unwrap()
    }
}

pub fn para(id: &str, content: &str) -> String {
    p(id, content)
}

pub fn run(text: &str) -> String {
    r(text)
}

/// An inline image run relating `id`.
pub fn image(id: &str, description: &str) -> String {
    format!(
        r#"<w:r><w:drawing><wp:inline><wp:extent cx="914400" cy="457200"/><wp:docPr id="9" name="image" descr="{description}"/><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:pic><pic:blipFill><a:blip r:embed="{id}"/></pic:blipFill></pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing></w:r>"#
    )
}

pub fn namespaces() -> &'static str {
    NS
}
