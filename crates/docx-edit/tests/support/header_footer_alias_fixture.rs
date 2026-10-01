use std::collections::BTreeSet;

const NS: &str = r#"xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml""#;
const REL: &str = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const OFFICE: &str = "application/vnd.openxmlformats-officedocument.wordprocessingml";

pub fn package(headers: &[(&str, &str)], footers: &[(&str, &str)]) -> Vec<u8> {
    let mut parts = Vec::new();
    let mut relationships = String::new();
    let mut overrides = String::new();
    let mut paths = BTreeSet::new();
    for (kind, root, entries) in [("header", "hdr", headers), ("footer", "ftr", footers)] {
        for (id, target) in entries {
            relationships.push_str(&format!(
                r#"<Relationship Id="{id}" Type="{REL}/{kind}" Target="{target}"/>"#
            ));
            let path = format!("word/{}", target.trim_start_matches("./"));
            if paths.insert(path.clone()) {
                overrides.push_str(&format!(
                    r#"<Override PartName="/{path}" ContentType="{OFFICE}.{kind}+xml"/>"#
                ));
                parts.push((
                    path,
                    format!(
                        r#"<w:{root} {NS}><w:p w14:paraId="20000001"><w:r><w:t>Shared band</w:t></w:r></w:p></w:{root}>"#
                    ),
                ));
            }
        }
    }
    let mut body = String::new();
    let sections = headers.len().max(footers.len()).max(1);
    for index in 0..sections {
        let mut references = String::new();
        for (kind, entries) in [("header", headers), ("footer", footers)] {
            if let Some((id, _)) = entries.get(index).or_else(|| entries.last()) {
                references.push_str(&format!(
                    r#"<w:{kind}Reference w:type="default" r:id="{id}"/>"#
                ));
            }
        }
        let properties = format!(
            r#"<w:sectPr>{references}<w:type w:val="nextPage"/><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720"/></w:sectPr>"#
        );
        let id = format!("{:08X}", index + 1);
        if index + 1 < sections {
            body.push_str(&format!(
                r#"<w:p w14:paraId="{id}"><w:pPr>{properties}</w:pPr><w:r><w:t>Section {index}</w:t></w:r></w:p>"#
            ));
        } else {
            body.push_str(&format!(
                r#"<w:p w14:paraId="{id}"><w:r><w:t>Section {index}</w:t></w:r></w:p>{properties}"#
            ));
        }
    }
    parts.extend([
        (
            "[Content_Types].xml".to_owned(),
            format!(
                r#"<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="{OFFICE}.document.main+xml"/>{overrides}</Types>"#
            ),
        ),
        (
            "_rels/.rels".to_owned(),
            format!(
                r#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="doc" Type="{REL}/officeDocument" Target="word/document.xml"/></Relationships>"#
            ),
        ),
        (
            "word/_rels/document.xml.rels".to_owned(),
            format!(
                r#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">{relationships}</Relationships>"#
            ),
        ),
        (
            "word/document.xml".to_owned(),
            format!(r#"<w:document {NS}><w:body>{body}</w:body></w:document>"#),
        ),
    ]);
    ooxml_opc::rezip_parts(
        &parts
            .into_iter()
            .map(|(path, xml)| (path, xml.into_bytes()))
            .collect::<Vec<_>>(),
    )
    .unwrap()
}
