const NS: &str = r#"xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships""#;
const REL: &str = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

pub fn deck() -> Vec<u8> {
    let shape = |id, placeholder: &str, style: &str, paragraphs: &str| {
        format!(
            r#"<p:sp><p:nvSpPr><p:cNvPr id="{id}" name="Strike {id}"/><p:cNvSpPr txBox="1"/><p:nvPr>{placeholder}</p:nvPr></p:nvSpPr><p:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="6000000" cy="1000000"/></a:xfrm></p:spPr><p:txBody><a:bodyPr/><a:lstStyle><a:defPPr><a:defRPr sz="1800"><a:latin typeface="Arial"/></a:defRPr></a:defPPr>{style}</a:lstStyle>{paragraphs}</p:txBody></p:sp>"#
        )
    };
    let shapes = [
        shape(
            2,
            "",
            r#"<a:lvl1pPr><a:defRPr strike="dblStrike"/></a:lvl1pPr><a:lvl2pPr><a:defRPr strike="noStrike"/></a:lvl2pPr>"#,
            r#"<a:p><a:r><a:rPr lang="en-US" strike="sngStrike" dirty="0"/><a:t>A</a:t></a:r><a:r><a:rPr strike="dblStrike"/><a:t>B</a:t></a:r><a:r><a:rPr strike="noStrike"/><a:t>C</a:t></a:r><a:r><a:t>D</a:t></a:r></a:p><a:p><a:pPr lvl="1"/><a:r><a:t>E</a:t></a:r></a:p>"#,
        ),
        shape(
            3,
            r#"<p:ph type="body" idx="1"/>"#,
            "",
            r#"<a:p><a:r><a:t>M</a:t></a:r><a:r><a:rPr strike="noStrike"/><a:t>N</a:t></a:r></a:p>"#,
        ),
        shape(
            4,
            r#"<p:ph type="title"/>"#,
            "",
            r#"<a:p><a:r><a:t>L</a:t></a:r><a:r><a:rPr strike="noStrike"/><a:t>O</a:t></a:r></a:p>"#,
        ),
        shape(5, "", "", r#"<a:p><a:r><a:t>P</a:t></a:r><a:r><a:rPr strike="sngStrike"/><a:t>Q</a:t></a:r></a:p>"#),
    ]
    .join("");
    let slide = format!(r#"<p:sld {NS}><p:cSld><p:spTree>{shapes}</p:spTree></p:cSld></p:sld>"#);
    let layout_shape = shape(
        2,
        r#"<p:ph type="title"/>"#,
        r#"<a:lvl1pPr><a:defRPr strike="dblStrike"/></a:lvl1pPr>"#,
        "<a:p/>",
    );
    let layout = format!(
        r#"<p:sldLayout {NS}><p:cSld><p:spTree>{layout_shape}</p:spTree></p:cSld></p:sldLayout>"#
    );
    let master = format!(
        r#"<p:sldMaster {NS}><p:cSld><p:spTree/></p:cSld><p:txStyles><p:titleStyle><a:lvl1pPr><a:defRPr strike="sngStrike"/></a:lvl1pPr></p:titleStyle><p:bodyStyle><a:lvl1pPr><a:defRPr strike="sngStrike"/></a:lvl1pPr></p:bodyStyle></p:txStyles></p:sldMaster>"#
    );
    let parts = [
        ("[Content_Types].xml", r#"<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="xml" ContentType="application/xml"/><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Override PartName="/ppt/presentation.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/><Override PartName="/ppt/slides/slide1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/><Override PartName="/ppt/slideLayouts/slideLayout1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml"/><Override PartName="/ppt/slideMasters/slideMaster1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideMaster+xml"/></Types>"#.to_owned()),
        ("_rels/.rels", format!(r#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="{REL}/officeDocument" Target="ppt/presentation.xml"/></Relationships>"#)),
        ("ppt/presentation.xml", format!(r#"<p:presentation {NS}><p:sldMasterIdLst><p:sldMasterId id="2147483648" r:id="rId1"/></p:sldMasterIdLst><p:sldIdLst><p:sldId id="256" r:id="rId2"/></p:sldIdLst><p:sldSz cx="9144000" cy="6858000"/></p:presentation>"#)),
        ("ppt/_rels/presentation.xml.rels", format!(r#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="{REL}/slideMaster" Target="slideMasters/slideMaster1.xml"/><Relationship Id="rId2" Type="{REL}/slide" Target="slides/slide1.xml"/></Relationships>"#)),
        ("ppt/slides/slide1.xml", slide),
        ("ppt/slides/_rels/slide1.xml.rels", format!(r#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="{REL}/slideLayout" Target="../slideLayouts/slideLayout1.xml"/></Relationships>"#)),
        ("ppt/slideLayouts/slideLayout1.xml", layout),
        ("ppt/slideLayouts/_rels/slideLayout1.xml.rels", format!(r#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="{REL}/slideMaster" Target="../slideMasters/slideMaster1.xml"/></Relationships>"#)),
        ("ppt/slideMasters/slideMaster1.xml", master),
    ]
    .into_iter()
    .map(|(path, xml)| (path.to_owned(), xml.into_bytes()))
    .collect::<Vec<_>>();
    ooxml_opc::rezip_parts(&parts).unwrap()
}
