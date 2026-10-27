pub const MEDIA: &str = "ppt/media/preview.png";
pub const OBJECT: &str = r#"<p:oleObj spid="_x0000_s3075" name="Equation" r:id="rId5" imgW="710891" imgH="444307" progId="Equation.3"><p:embed/></p:oleObj>"#;
pub const VML: &str = r##"<xml xmlns:v="urn:schemas-microsoft-com:vml" xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><v:shapetype id="_x0000_t75"/><v:shape id="_x0000_s3075" type="#_x0000_t75" style="position:absolute;left:462pt;top:343.75pt;width:90.75pt;height:58.75pt"><v:imagedata o:relid="rId1" o:title=""/></v:shape></xml>"##;

pub fn parts(vml: &str, object: &str) -> Vec<(String, Vec<u8>)> {
    let office = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
    let relationships = "http://schemas.openxmlformats.org/package/2006/relationships";
    let xml = [
        (
            "[Content_Types].xml",
            r#"<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="xml" ContentType="application/xml"/><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="vml" ContentType="application/vnd.openxmlformats-officedocument.vmlDrawing"/><Default Extension="png" ContentType="image/png"/><Override PartName="/ppt/presentation.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/><Override PartName="/ppt/slides/slide1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/></Types>"#.to_owned(),
        ),
        (
            "_rels/.rels",
            format!(r#"<Relationships xmlns="{relationships}"><Relationship Id="rId1" Type="{office}/officeDocument" Target="ppt/presentation.xml"/></Relationships>"#),
        ),
        (
            "ppt/presentation.xml",
            format!(r#"<p:presentation xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:r="{office}"><p:sldIdLst><p:sldId id="256" r:id="rId1"/></p:sldIdLst><p:sldSz cx="12192000" cy="6858000"/></p:presentation>"#),
        ),
        (
            "ppt/_rels/presentation.xml.rels",
            format!(r#"<Relationships xmlns="{relationships}"><Relationship Id="rId1" Type="{office}/slide" Target="slides/slide1.xml"/></Relationships>"#),
        ),
        (
            "ppt/slides/slide1.xml",
            format!(r#"<p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="{office}" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"><p:cSld><p:spTree><p:graphicFrame><p:nvGraphicFramePr><p:cNvPr id="4" name="Object 3"/><p:cNvGraphicFramePr/><p:nvPr/></p:nvGraphicFramePr><p:xfrm><a:off x="5868144" y="4365104"/><a:ext cx="1152128" cy="747326"/></p:xfrm><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/presentationml/2006/ole">{object}</a:graphicData></a:graphic></p:graphicFrame></p:spTree></p:cSld></p:sld>"#),
        ),
        (
            "ppt/slides/_rels/slide1.xml.rels",
            format!(r#"<Relationships xmlns="{relationships}"><Relationship Id="rId1" Type="{office}/image" Target="../media/wrong.png"/><Relationship Id="rId3" Type="{office}/vmlDrawing" Target="../drawings/vmlDrawing1.vml"/><Relationship Id="rId5" Type="{office}/oleObject" Target="../embeddings/oleObject1.bin"/></Relationships>"#),
        ),
        ("ppt/drawings/vmlDrawing1.vml", vml.to_owned()),
        (
            "ppt/drawings/_rels/vmlDrawing1.vml.rels",
            format!(r#"<Relationships xmlns="{relationships}"><Relationship Id="rId1" Type="{office}/image" Target="../media/preview.png"/></Relationships>"#),
        ),
    ];
    let mut parts: Vec<_> = xml
        .into_iter()
        .map(|(path, xml)| (path.to_owned(), xml.into_bytes()))
        .collect();
    parts.push((
        MEDIA.to_owned(),
        vec![
            0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a, 0, 0, 0, 0x0d, 0x49, 0x48, 0x44, 0x52,
            0, 0, 0, 1, 0, 0, 0, 1, 8, 6, 0, 0, 0, 0x1f, 0x15, 0xc4, 0x89, 0, 0, 0, 0x0d, 0x49,
            0x44, 0x41, 0x54, 0x78, 0x9c, 0x63, 0xf8, 0xcf, 0xc0, 0xf0, 0x1f, 0, 0x05, 0, 0x01,
            0xff, 0x7d, 0x31, 0x96, 0xd5, 0, 0, 0, 0, 0x49, 0x45, 0x4e, 0x44, 0xae, 0x42, 0x60,
            0x82,
        ],
    ));
    parts.push(("ppt/embeddings/oleObject1.bin".to_owned(), vec![0]));
    parts
}

pub fn deck(vml: &str, object: &str) -> Vec<u8> {
    ooxml_opc::rezip_parts(&parts(vml, object)).unwrap()
}
