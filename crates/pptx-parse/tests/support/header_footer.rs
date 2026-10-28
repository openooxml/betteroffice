pub const MASTER: &str = "ppt/slideMasters/slideMaster1.xml";
pub const LAYOUT: &str = "ppt/slideLayouts/slideLayout1.xml";
const NS: &str = r#"xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships""#;
const REL: &str = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const GROUP: &str =
    r#"<p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/>"#;

pub fn placeholder(id: u32, kind: &str, paragraphs: &str, hidden: bool) -> String {
    let hidden = if hidden { r#" hidden="1""# } else { "" };
    format!(
        r#"<p:sp><p:nvSpPr><p:cNvPr id="{id}" name="{kind}"{hidden}/><p:cNvSpPr/><p:nvPr><p:ph type="{kind}"/></p:nvPr></p:nvSpPr><p:spPr/><p:txBody><a:bodyPr/><a:lstStyle/>{paragraphs}</p:txBody></p:sp>"#
    )
}

pub fn text(value: &str) -> String {
    format!("<a:p><a:r><a:t>{value}</a:t></a:r></a:p>")
}

pub fn field(kind: &str, cached: &str) -> String {
    format!(
        r#"<a:fld id="{{00000000-0000-0000-0000-000000000001}}" type="{kind}"><a:t>{cached}</a:t></a:fld>"#
    )
}

pub fn group(shapes: &str) -> String {
    format!(
        "<p:grpSp>{}{shapes}</p:grpSp>",
        GROUP.replace(r#"id="1""#, r#"id="10""#)
    )
}

fn relationships(entries: &[(&str, &str, &str)]) -> String {
    let entries = entries
        .iter()
        .map(|(id, kind, target)| {
            format!(r#"<Relationship Id="{id}" Type="{REL}/{kind}" Target="{target}"/>"#)
        })
        .collect::<String>();
    format!(
        r#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">{entries}</Relationships>"#
    )
}

pub fn deck(
    slides: &[(&str, bool)],
    layout_hf: &str,
    master_hf: &str,
    parent_shapes: &str,
    first_slide_num: Option<u32>,
) -> Vec<u8> {
    let mut parts = Vec::new();
    let mut add = |path: &str, xml: String| parts.push((path.to_owned(), xml.into_bytes()));
    add(
        "[Content_Types].xml",
        r#"<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/ppt/presentation.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/></Types>"#.to_owned(),
    );
    add(
        "_rels/.rels",
        relationships(&[("rId1", "officeDocument", "ppt/presentation.xml")]),
    );
    let first = first_slide_num.map_or(String::new(), |value| {
        format!(r#" firstSlideNum="{value}""#)
    });
    let ids = (0..slides.len())
        .map(|index| format!(r#"<p:sldId id="{}" r:id="rId{}"/>"#, 256 + index, index + 2))
        .collect::<String>();
    add(
        "ppt/presentation.xml",
        format!(
            r#"<p:presentation {NS}{first}><p:sldMasterIdLst><p:sldMasterId id="2147483648" r:id="rId1"/></p:sldMasterIdLst><p:sldIdLst>{ids}</p:sldIdLst><p:sldSz cx="9144000" cy="5143500"/></p:presentation>"#
        ),
    );
    let mut rels = vec![(
        "rId1".to_owned(),
        "slideMaster",
        "slideMasters/slideMaster1.xml".to_owned(),
    )];
    for (index, (shapes, hidden)) in slides.iter().enumerate() {
        let number = index + 1;
        rels.push((
            format!("rId{}", index + 2),
            "slide",
            format!("slides/slide{number}.xml"),
        ));
        let show = if *hidden { r#" show="0""# } else { "" };
        add(
            &format!("ppt/slides/slide{number}.xml"),
            format!(
                r#"<p:sld {NS}{show}><p:cSld><p:spTree>{GROUP}{shapes}</p:spTree></p:cSld></p:sld>"#
            ),
        );
        add(
            &format!("ppt/slides/_rels/slide{number}.xml.rels"),
            relationships(&[("rId1", "slideLayout", "../slideLayouts/slideLayout1.xml")]),
        );
    }
    let rels = rels
        .iter()
        .map(|(id, kind, target)| (id.as_str(), *kind, target.as_str()))
        .collect::<Vec<_>>();
    add("ppt/_rels/presentation.xml.rels", relationships(&rels));
    add(
        MASTER,
        format!(
            r#"<p:sldMaster {NS}><p:cSld><p:spTree>{GROUP}{parent_shapes}</p:spTree></p:cSld><p:clrMap accent1="accent1" bg1="lt1" tx1="dk1"/><p:sldLayoutIdLst><p:sldLayoutId id="2147483649" r:id="rId1"/></p:sldLayoutIdLst>{master_hf}</p:sldMaster>"#
        ),
    );
    add(
        "ppt/slideMasters/_rels/slideMaster1.xml.rels",
        relationships(&[("rId1", "slideLayout", "../slideLayouts/slideLayout1.xml")]),
    );
    add(
        LAYOUT,
        format!(
            r#"<p:sldLayout {NS} type="title"><p:cSld name="Header layout"><p:spTree>{GROUP}{parent_shapes}</p:spTree></p:cSld>{layout_hf}</p:sldLayout>"#
        ),
    );
    add(
        "ppt/slideLayouts/_rels/slideLayout1.xml.rels",
        relationships(&[("rId1", "slideMaster", "../slideMasters/slideMaster1.xml")]),
    );
    ooxml_opc::rezip_parts(&parts).unwrap()
}
