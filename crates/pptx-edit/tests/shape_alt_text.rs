use std::collections::BTreeMap;

use pptx_edit::structured::{ExportShape, PptxExportOptions, export_pptx_structured};
use pptx_edit::{DeckSession, ReadRequest, ShapeKind, ShapeSnapshot};

const NS: &str = r#"xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main""#;
const REL: &str = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const DECORATIVE_URI: &str = "{C183D7F6-B498-43B3-948B-1728B52AA6E4}";

fn decorative(val: &str) -> String {
    format!(
        r#"<a:extLst><a:ext uri="{DECORATIVE_URI}"><adec:decorative xmlns:adec="http://schemas.microsoft.com/office/drawing/2017/decorative" val="{val}"/></a:ext></a:extLst>"#
    )
}

fn slide() -> String {
    let shape = format!(
        r#"<p:sp><p:nvSpPr><p:cNvPr id="2" name="Alt shape" title="Shape title" descr="Shape description">{}</p:cNvPr><p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:spPr><a:xfrm><a:off x="100000" y="100000"/><a:ext cx="1000000" cy="500000"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr><p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:t>Shape text</a:t></a:r></a:p></p:txBody></p:sp>"#,
        decorative("1")
    );
    let picture = format!(
        r#"<p:pic><p:nvPicPr><p:cNvPr id="3" name="Alt picture" title="Picture title" descr="">{}</p:cNvPr><p:cNvPicPr/><p:nvPr/></p:nvPicPr><p:blipFill><a:blip r:embed="rId2"/><a:stretch><a:fillRect/></a:stretch></p:blipFill><p:spPr><a:xfrm><a:off x="1200000" y="100000"/><a:ext cx="800000" cy="500000"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr></p:pic>"#,
        decorative("0")
    );
    let table = r#"<p:graphicFrame><p:nvGraphicFramePr><p:cNvPr id="4" name="Alt table" title="Table title" descr="Table description"/><p:cNvGraphicFramePr/><p:nvPr/></p:nvGraphicFramePr><p:xfrm><a:off x="100000" y="700000"/><a:ext cx="2000000" cy="500000"/></p:xfrm><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/table"><a:tbl><a:tblPr/><a:tblGrid><a:gridCol w="2000000"/></a:tblGrid><a:tr h="500000"><a:tc><a:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:t>Cell text</a:t></a:r></a:p></a:txBody><a:tcPr/></a:tc></a:tr></a:tbl></a:graphicData></a:graphic></p:graphicFrame>"#;
    let child = format!(
        r#"<p:pic><p:nvPicPr><p:cNvPr id="6" name="Alt child picture" title="Child title" descr="Child description">{}</p:cNvPr><p:cNvPicPr/><p:nvPr/></p:nvPicPr><p:blipFill><a:blip r:embed="rId2"/><a:stretch><a:fillRect/></a:stretch></p:blipFill><p:spPr><a:xfrm><a:off x="200000" y="1400000"/><a:ext cx="800000" cy="500000"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr></p:pic>"#,
        decorative("true")
    );
    let group = format!(
        r#"<p:grpSp><p:nvGrpSpPr><p:cNvPr id="5" name="Alt group" title="Group title" descr="Group description"/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr><a:xfrm><a:off x="100000" y="1300000"/><a:ext cx="2000000" cy="1000000"/><a:chOff x="100000" y="1300000"/><a:chExt cx="2000000" cy="1000000"/></a:xfrm></p:grpSpPr>{child}</p:grpSp>"#
    );
    let connector = r#"<p:cxnSp><p:nvCxnSpPr><p:cNvPr id="7" name="Alt connector" title="Connector title" descr="Connector description"/><p:cNvCxnSpPr/><p:nvPr/></p:nvCxnSpPr><p:spPr><a:xfrm><a:off x="100000" y="2500000"/><a:ext cx="1000000" cy="0"/></a:xfrm><a:prstGeom prst="line"><a:avLst/></a:prstGeom></p:spPr></p:cxnSp>"#;
    format!(
        r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?><p:sld {NS}><p:cSld><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/>{shape}{picture}{table}{group}{connector}</p:spTree></p:cSld></p:sld>"#
    )
}

fn relationships(entries: &[(&str, &str, &str, Option<&str>)]) -> String {
    let entries: String = entries
        .iter()
        .map(|(id, kind, target, mode)| {
            let mode = mode.map_or(String::new(), |mode| format!(r#" TargetMode="{mode}""#));
            format!(r#"<Relationship Id="{id}" Type="{REL}/{kind}" Target="{target}"{mode}/>"#)
        })
        .collect();
    format!(
        r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">{entries}</Relationships>"#
    )
}

fn fixture() -> Vec<u8> {
    let parts: Vec<(String, Vec<u8>)> = [
        (
            "[Content_Types].xml",
            r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Default Extension="png" ContentType="image/png"/><Override PartName="/ppt/presentation.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/><Override PartName="/ppt/slides/slide1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/><Override PartName="/ppt/slideLayouts/slideLayout1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml"/></Types>"#
                .to_owned(),
        ),
        (
            "_rels/.rels",
            relationships(&[("rId1", "officeDocument", "ppt/presentation.xml", None)]),
        ),
        (
            "ppt/presentation.xml",
            format!(
                r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?><p:presentation {NS}><p:sldIdLst><p:sldId id="256" r:id="rId1"/></p:sldIdLst><p:sldSz cx="9144000" cy="5143500"/></p:presentation>"#
            ),
        ),
        (
            "ppt/_rels/presentation.xml.rels",
            relationships(&[("rId1", "slide", "slides/slide1.xml", None)]),
        ),
        ("ppt/slides/slide1.xml", slide()),
        (
            "ppt/slides/_rels/slide1.xml.rels",
            relationships(&[
                ("rId1", "slideLayout", "../slideLayouts/slideLayout1.xml", None),
                ("rId2", "image", "../media/image1.png", None),
            ]),
        ),
        (
            "ppt/slideLayouts/slideLayout1.xml",
            format!(
                r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?><p:sldLayout {NS} type="blank"><p:cSld name="Blank"><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/></p:spTree></p:cSld></p:sldLayout>"#
            ),
        ),
    ]
    .into_iter()
    .map(|(path, body)| (path.to_owned(), body.into_bytes()))
    .chain([(
        "ppt/media/image1.png".to_owned(),
        vec![0x89, b'P', b'N', b'G', 0x0d, 0x0a, 0x1a, 0x0a],
    )])
    .collect();
    ooxml_opc::rezip_parts(&parts).unwrap()
}

fn parts(bytes: &[u8]) -> BTreeMap<String, Vec<u8>> {
    ooxml_opc::unzip_parts(bytes).unwrap().into_iter().collect()
}

fn find_opt<'a>(shapes: &'a [ShapeSnapshot], name: &str) -> Option<&'a ShapeSnapshot> {
    shapes.iter().find_map(|shape| {
        if shape.name == name {
            Some(shape)
        } else {
            find_opt(&shape.children, name)
        }
    })
}

fn find<'a>(shapes: &'a [ShapeSnapshot], name: &str) -> &'a ShapeSnapshot {
    find_opt(shapes, name).unwrap_or_else(|| panic!("no shape named {name}"))
}

fn find_export<'a>(shapes: &'a [ExportShape], name: &str) -> Option<&'a ExportShape> {
    shapes.iter().find_map(|shape| {
        if shape.name == name {
            Some(shape)
        } else {
            find_export(&shape.children, name)
        }
    })
}

#[test]
fn read_snapshot_carries_shape_alt_text_at_every_depth() {
    let bytes = fixture();
    let session = DeckSession::open(&bytes, 41).unwrap();
    let read = session
        .read_content(&ReadRequest::default())
        .unwrap()
        .unwrap();
    let shapes = &read.slides[0].shapes;
    assert_eq!(shapes.len(), 5);

    let shape = find(shapes, "Alt shape");
    assert_eq!(shape.kind, ShapeKind::Shape);
    assert_eq!(shape.title.as_deref(), Some("Shape title"));
    assert_eq!(shape.description.as_deref(), Some("Shape description"));
    assert!(shape.decorative);

    let picture = find(shapes, "Alt picture");
    assert_eq!(picture.kind, ShapeKind::Picture);
    assert_eq!(picture.title.as_deref(), Some("Picture title"));
    assert_eq!(picture.description, None);
    assert!(!picture.decorative);
    let picture_json = serde_json::to_value(picture).unwrap();
    assert!(picture_json.get("description").is_none());
    assert!(picture_json.get("decorative").is_none());

    let table = find(shapes, "Alt table");
    assert_eq!(table.kind, ShapeKind::GraphicFrame);
    assert_eq!(table.title.as_deref(), Some("Table title"));
    assert_eq!(table.description.as_deref(), Some("Table description"));
    assert!(!table.decorative);

    let group = find(shapes, "Alt group");
    assert_eq!(group.kind, ShapeKind::Group);
    assert_eq!(group.title.as_deref(), Some("Group title"));
    assert_eq!(group.description.as_deref(), Some("Group description"));

    let child = find(shapes, "Alt child picture");
    assert_eq!(child.kind, ShapeKind::Picture);
    assert_eq!(child.title.as_deref(), Some("Child title"));
    assert_eq!(child.description.as_deref(), Some("Child description"));
    assert!(child.decorative);

    let connector = find(shapes, "Alt connector");
    assert_eq!(connector.title.as_deref(), Some("Connector title"));
    assert_eq!(
        connector.description.as_deref(),
        Some("Connector description")
    );

    let exported = export_pptx_structured(&bytes, &PptxExportOptions::default()).unwrap();
    for name in [
        "Alt shape",
        "Alt picture",
        "Alt table",
        "Alt group",
        "Alt child picture",
        "Alt connector",
    ] {
        let snapshot = find(shapes, name);
        let record = find_export(&exported.slides[0].shapes, name)
            .unwrap_or_else(|| panic!("no exported shape named {name}"));
        let present = |value: &Option<String>| value.clone().filter(|text| !text.is_empty());
        assert_eq!(present(&record.title), snapshot.title, "{name} title");
        assert_eq!(
            present(&record.description),
            snapshot.description,
            "{name} description"
        );
    }

    assert_eq!(parts(&session.save().unwrap()), parts(&bytes));
}
