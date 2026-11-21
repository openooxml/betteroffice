use docx_edit::{EngineSession, seed_from_docx};
use serde_json::{Value, json};

const FONT: &[u8] = include_bytes!("../../ooxml-text/tests/fonts/LiberationSans-Regular.ttf");
const W: &str = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const R: &str = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const A: &str = "http://schemas.openxmlformats.org/drawingml/2006/main";
const WP: &str = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing";
const DGM: &str = "http://schemas.openxmlformats.org/drawingml/2006/diagram";
const DSP: &str = "http://schemas.microsoft.com/office/drawing/2008/diagram";
const RELS: &str = "http://schemas.openxmlformats.org/package/2006/relationships";

fn smart_art_docx() -> Vec<u8> {
    let drawing = format!(
        r#"<w:drawing><wp:inline distT="0" distB="0" distL="0" distR="0"><wp:extent cx="2743200" cy="1371600"/><wp:docPr id="71" name="Diagram 1"/><wp:cNvGraphicFramePr/><a:graphic><a:graphicData uri="{DGM}"><dgm:relIds r:dm="rId10" r:lo="rId11" r:qs="rId12" r:cs="rId13"/></a:graphicData></a:graphic></wp:inline></w:drawing>"#
    );
    let document = format!(
        r#"<w:document xmlns:w="{W}" xmlns:r="{R}" xmlns:a="{A}" xmlns:wp="{WP}" xmlns:dgm="{DGM}" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml"><w:body><w:p w14:paraId="10000000"><w:r><w:t>Opening paragraph above.</w:t></w:r></w:p><w:p w14:paraId="10000001"><w:r><w:t xml:space="preserve">BEFORE </w:t></w:r><w:r>{drawing}</w:r><w:r><w:t xml:space="preserve"> AFTER</w:t></w:r></w:p><w:p w14:paraId="10000002"><w:r><w:t>Next paragraph.</w:t></w:r></w:p><w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"/></w:sectPr></w:body></w:document>"#
    );
    let parts_xml: Vec<(&str, &str, String)> = vec![
        (
            "data1.xml",
            "diagramData",
            format!(
                r#"<dgm:dataModel xmlns:dgm="{DGM}" xmlns:a="{A}"><dgm:ptLst><dgm:pt modelId="root" type="doc"><dgm:prSet loTypeId="urn:microsoft.com/office/officeart/2005/8/layout/process1"/></dgm:pt><dgm:pt modelId="item"><dgm:prSet/><dgm:t><a:bodyPr/><a:p><a:r><a:t>Plan</a:t></a:r></a:p></dgm:t></dgm:pt></dgm:ptLst><dgm:cxnLst><dgm:cxn modelId="edge" srcId="root" destId="item" srcOrd="0" destOrd="0"/></dgm:cxnLst><dgm:extLst><a:ext uri="{DSP}"><dsp:dataModelExt xmlns:dsp="{DSP}" relId="rId14"/></a:ext></dgm:extLst></dgm:dataModel>"#
            ),
        ),
        (
            "layout1.xml",
            "diagramLayout",
            format!(
                r#"<dgm:layoutDef xmlns:dgm="{DGM}" uniqueId="urn:microsoft.com/office/officeart/2005/8/layout/process1"><dgm:layoutNode name="process"/></dgm:layoutDef>"#
            ),
        ),
        (
            "quickStyle1.xml",
            "diagramQuickStyle",
            format!(r#"<dgm:styleDef xmlns:dgm="{DGM}" uniqueId="s"/>"#),
        ),
        (
            "colors1.xml",
            "diagramColors",
            format!(r#"<dgm:colorsDef xmlns:dgm="{DGM}" uniqueId="c"/>"#),
        ),
        (
            "drawing1.xml",
            "diagramDrawing",
            format!(
                r#"<dsp:drawing xmlns:dsp="{DSP}" xmlns:a="{A}"><dsp:spTree><dsp:nvGrpSpPr><dsp:cNvPr id="0" name=""/><dsp:cNvGrpSpPr/></dsp:nvGrpSpPr><dsp:grpSpPr/><dsp:sp modelId="item"><dsp:nvSpPr><dsp:cNvPr id="2" name="Plan"/><dsp:cNvSpPr/></dsp:nvSpPr><dsp:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="1600000" cy="960000"/></a:xfrm><a:prstGeom prst="roundRect"><a:avLst/></a:prstGeom></dsp:spPr><dsp:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:t>Plan</a:t></a:r></a:p></dsp:txBody></dsp:sp></dsp:spTree></dsp:drawing>"#
            ),
        ),
    ];
    let rel_type = |kind: &str| {
        if kind == "diagramDrawing" {
            "http://schemas.microsoft.com/office/2007/relationships/diagramDrawing".to_owned()
        } else {
            format!("{R}/{kind}")
        }
    };
    let ids = ["rId10", "rId11", "rId12", "rId13", "rId14"];
    let rels: String = parts_xml
        .iter()
        .zip(ids)
        .map(|((name, kind, _), id)| {
            format!(
                r#"<Relationship Id="{id}" Type="{}" Target="diagrams/{name}"/>"#,
                rel_type(kind)
            )
        })
        .collect();
    let overrides: String = parts_xml
        .iter()
        .map(|(name, kind, _)| {
            let content = if *kind == "diagramDrawing" {
                "application/vnd.ms-office.drawingml.diagramDrawing+xml".to_owned()
            } else {
                format!(
                    "application/vnd.openxmlformats-officedocument.drawingml.{}+xml",
                    if *kind == "diagramQuickStyle" {
                        "diagramStyle"
                    } else {
                        kind
                    }
                )
            };
            format!(r#"<Override PartName="/word/diagrams/{name}" ContentType="{content}"/>"#)
        })
        .collect();
    let mut parts = vec![
        ("[Content_Types].xml".to_owned(), format!(r#"<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>{overrides}</Types>"#).into_bytes()),
        ("_rels/.rels".to_owned(), format!(r#"<Relationships xmlns="{RELS}"><Relationship Id="main" Type="{R}/officeDocument" Target="word/document.xml"/></Relationships>"#).into_bytes()),
        ("word/document.xml".to_owned(), document.into_bytes()),
        ("word/_rels/document.xml.rels".to_owned(), format!(r#"<Relationships xmlns="{RELS}">{rels}</Relationships>"#).into_bytes()),
        ("word/diagrams/_rels/data1.xml.rels".to_owned(), format!(r#"<Relationships xmlns="{RELS}"><Relationship Id="rId14" Type="http://schemas.microsoft.com/office/2007/relationships/diagramDrawing" Target="drawing1.xml"/></Relationships>"#).into_bytes()),
    ];
    for (name, _, xml) in parts_xml {
        parts.push((format!("word/diagrams/{name}"), xml.into_bytes()));
    }
    ooxml_opc::rezip_parts(&parts).unwrap()
}

fn layout_request(engine: &EngineSession, bytes: &[u8]) -> String {
    let package = docx_parse::parse_docx_s9_wire(bytes, Default::default())
        .unwrap()
        .document
        .package;
    let sections: Vec<_> = package
        .document
        .sections
        .unwrap()
        .into_iter()
        .map(|section| json!({"properties": section.properties}))
        .collect();
    let font = docx_layout::register_measure_font(FONT).unwrap();
    let mut request = json!({
        "bodyStory": "body", "renderEnv": {},
        "regions": {"sections": sections, "settings": package.settings},
        "measurement": {"fontChains": {"arial|0|0": [font]}, "defaults": {"fontFamily": "Arial", "fontSize": 12}, "authoritativeShaping": true}
    });
    let requirements: Vec<Value> = serde_json::from_str(
        &engine
            .layout_font_requirements_json(&request.to_string())
            .unwrap(),
    )
    .unwrap();
    for requirement in requirements {
        request["measurement"]["fontChains"][requirement["key"].as_str().unwrap()] = json!([font]);
    }
    request.to_string()
}

#[test]
fn an_inline_smart_art_stays_between_the_text_around_it() {
    let bytes = smart_art_docx();
    let engine = EngineSession::new(41);
    seed_from_docx(engine.doc(), &bytes).unwrap();
    let request = layout_request(&engine, &bytes);
    let output: Value =
        serde_json::from_str(&engine.layout_document_with_regions_json(&request).unwrap()).unwrap();
    let fragments = output["layout"]["pages"][0]["fragments"]
        .as_array()
        .unwrap();
    let listed = serde_json::to_string(fragments).unwrap();
    let own = |fragment: &Value| {
        fragment["blockId"]
            .as_str()
            .is_some_and(|id| id.starts_with("10000001"))
    };
    let y = |fragment: &Value| fragment["y"].as_f64().unwrap();
    let diagram = fragments
        .iter()
        .find(|fragment| fragment["kind"] == "shape")
        .unwrap_or_else(|| panic!("the inline diagram is laid out: {listed}"));
    let first = fragments.iter().find(|fragment| own(fragment)).unwrap();
    let last = fragments
        .iter()
        .rev()
        .find(|fragment| own(fragment))
        .unwrap();
    assert!(
        y(first) <= y(diagram) && y(last) >= y(diagram),
        "the inline diagram is laid out outside the text around it: {listed}"
    );
}
