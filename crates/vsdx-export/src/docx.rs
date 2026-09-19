use std::collections::BTreeMap;

use ooxml_drawingml::{cust_geom, emu, escape_xml, place_rect, srgb_hex, xfrm_xml};
use vsdx_parse::VsdxPackage;
use vsdx_render::Primitive;

use crate::ExportError;
use crate::ExportReport;
use crate::Page;
use crate::metadata::shape_data;
use crate::shared::{
    Media, collect_media, emit_matrix, fallback_rect, fill_string, flat, has_text, line_string,
    scale_list,
};

const CONTENT_WIDTH_IN: f64 = 6.5;

pub fn build(
    pages: &[Page],
    package: &VsdxPackage,
    report: &mut ExportReport,
) -> Result<Vec<u8>, ExportError> {
    let data = shape_data(package)?;
    let mut media: Vec<Media> = Vec::new();
    let mut index_by_asset: BTreeMap<String, usize> = BTreeMap::new();
    for page in pages {
        collect_media(
            &page.list.primitives,
            package,
            &mut media,
            &mut index_by_asset,
            "word/media",
        );
    }
    let mut parts: Vec<(String, Vec<u8>)> = vec![
        ("[Content_Types].xml".to_owned(), content_types(&media)),
        ("_rels/.rels".to_owned(), package_rels()),
        (
            "word/_rels/document.xml.rels".to_owned(),
            document_rels(&media),
        ),
        (
            "word/document.xml".to_owned(),
            document_xml(pages, &data, &index_by_asset, report),
        ),
        ("word/styles.xml".to_owned(), styles_xml()),
        ("docProps/core.xml".to_owned(), core_xml()),
        ("docProps/app.xml".to_owned(), app_xml()),
    ];
    for item in &media {
        parts.push((item.part.clone(), item.bytes.clone()));
    }
    ooxml_opc::rezip_parts(&parts).map_err(ExportError::Package)
}

fn fit_scale(page: &Page) -> f64 {
    (CONTENT_WIDTH_IN / page.width_in).min(1.0)
}

fn document_xml(
    pages: &[Page],
    data: &[crate::ShapeDatum],
    index_by_asset: &BTreeMap<String, usize>,
    report: &mut ExportReport,
) -> Vec<u8> {
    let mut body = String::new();
    for (index, page) in pages.iter().enumerate() {
        let doc_id = index as u32 + 1;
        let name = escape_xml(&page.name);
        body.push_str(&format!(
            "<w:p><w:pPr><w:pStyle w:val=\"Heading1\"/></w:pPr><w:r><w:t xml:space=\"preserve\">Page {name}</w:t></w:r></w:p>"
        ));
        let scale = fit_scale(page);
        let list = scale_list(&page.list, scale);
        let page_height = f64::from(list.height) / 96.0;
        let width = emu(f64::from(list.width) / 96.0).max(1);
        let height = emu(page_height).max(1);
        let mut members = String::new();
        for primitive in flat(&list.primitives) {
            members.push_str(&member_xml(&primitive, page_height, index_by_asset, report));
        }
        body.push_str(&format!(
            "<w:p><w:r><w:drawing><wp:inline distT=\"0\" distB=\"0\" distL=\"0\" distR=\"0\"><wp:extent cx=\"{width}\" cy=\"{height}\"/><wp:docPr id=\"{doc_id}\" name=\"{name}\"/><a:graphic><a:graphicData uri=\"http://schemas.microsoft.com/office/word/2010/wordprocessingGroup\"><wpg:wgp><wpg:cNvGrpSpPr/><wpg:grpSpPr><a:xfrm><a:off x=\"0\" y=\"0\"/><a:ext cx=\"{width}\" cy=\"{height}\"/><a:chOff x=\"0\" y=\"0\"/><a:chExt cx=\"{width}\" cy=\"{height}\"/></a:xfrm></wpg:grpSpPr>{members}</wpg:wgp></a:graphicData></a:graphic></wp:inline></w:drawing></w:r></w:p>"
        ));
        let rows: Vec<&crate::ShapeDatum> = data
            .iter()
            .filter(|datum| datum.page == page.name)
            .collect();
        if !rows.is_empty() {
            let title = escape_xml(&page.name);
            body.push_str(&format!(
                "<w:p><w:pPr><w:pStyle w:val=\"Heading2\"/></w:pPr><w:r><w:t xml:space=\"preserve\">Shape data for {title}</w:t></w:r></w:p>"
            ));
            body.push_str(&table_xml(&rows));
        }
    }
    format!("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?><w:document xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\" xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\" xmlns:wp=\"http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing\" xmlns:wps=\"http://schemas.microsoft.com/office/word/2010/wordprocessingShape\" xmlns:wpg=\"http://schemas.microsoft.com/office/word/2010/wordprocessingGroup\" xmlns:pic=\"http://schemas.openxmlformats.org/drawingml/2006/picture\"><w:body>{body}<w:sectPr><w:pgSz w:w=\"12240\" w:h=\"15840\"/><w:pgMar w:top=\"1440\" w:right=\"1440\" w:bottom=\"1440\" w:left=\"1440\" w:header=\"720\" w:footer=\"720\" w:gutter=\"0\"/></w:sectPr></w:body></w:document>").into_bytes()
}

fn member_xml(
    primitive: &Primitive,
    page_height: f64,
    index_by_asset: &BTreeMap<String, usize>,
    report: &mut ExportReport,
) -> String {
    match primitive {
        Primitive::Shape {
            id,
            path,
            fill,
            stroke,
            ..
        } => {
            let fill = fill_string(fill);
            let line = line_string(stroke);
            if let Some(geom) = cust_geom(path, page_height) {
                return format!(
                    "<wps:wsp><wps:cNvSpPr/><wps:spPr><a:xfrm><a:off x=\"{}\" y=\"{}\"/><a:ext cx=\"{}\" cy=\"{}\"/></a:xfrm><a:custGeom><a:avLst/><a:gdLst/><a:ahLst/><a:cxnLst/><a:pathLst><a:path w=\"{}\" h=\"{}\">{}</a:path></a:pathLst></a:custGeom>{fill}{line}</wps:spPr><wps:txbx><w:txbxContent><w:p/></w:txbxContent></wps:txbx><wps:bodyPr/></wps:wsp>",
                    geom.x, geom.y, geom.w, geom.h, geom.w, geom.h, geom.body
                );
            }
            report.empty_geometry.push(id.clone());
            match fallback_rect(path, page_height) {
                Some((x, y, w, h)) => format!(
                    "<wps:wsp><wps:cNvSpPr/><wps:spPr><a:xfrm><a:off x=\"{x}\" y=\"{y}\"/><a:ext cx=\"{w}\" cy=\"{h}\"/></a:xfrm><a:prstGeom prst=\"rect\"><a:avLst/></a:prstGeom>{fill}{line}</wps:spPr><wps:txbx><w:txbxContent><w:p/></w:txbxContent></wps:txbx><wps:bodyPr/></wps:wsp>"
                ),
                None => {
                    report.skipped.push(id.clone());
                    String::new()
                }
            }
        }
        Primitive::TextBox {
            id,
            x,
            y,
            width,
            height,
            paragraphs,
            transform,
            ..
        } => {
            let Some(placed) = place_rect(
                *x,
                *y,
                *width,
                *height,
                emit_matrix(*transform),
                page_height,
            ) else {
                report.skipped.push(id.clone());
                return String::new();
            };
            if placed.degraded {
                report.bbox_fallbacks.push(id.clone());
            }
            let mut text = wml_runs(paragraphs);
            if !has_text(paragraphs) {
                text = "<w:p/>".to_owned();
            }
            format!(
                "<wps:wsp><wps:cNvSpPr txBox=\"1\"/><wps:spPr>{}<a:prstGeom prst=\"rect\"><a:avLst/></a:prstGeom><a:noFill/></wps:spPr><wps:txbx><w:txbxContent>{text}</w:txbxContent></wps:txbx><wps:bodyPr/></wps:wsp>",
                xfrm_xml(&placed)
            )
        }
        Primitive::Placeholder {
            x,
            y,
            width,
            height,
            reason,
            ..
        } => {
            let label = escape_xml(reason);
            format!(
                "<wps:wsp><wps:cNvSpPr txBox=\"1\"/><wps:spPr><a:xfrm><a:off x=\"{}\" y=\"{}\"/><a:ext cx=\"{}\" cy=\"{}\"/></a:xfrm><a:prstGeom prst=\"rect\"><a:avLst/></a:prstGeom><a:noFill/><a:ln w=\"12700\"><a:solidFill><a:srgbClr val=\"8A94A6\"/></a:solidFill><a:prstDash val=\"dash\"/></a:ln></wps:spPr><wps:txbx><w:txbxContent><w:p><w:r><w:rPr><w:rFonts w:ascii=\"Calibri\" w:hAnsi=\"Calibri\" w:cs=\"Calibri\"/><w:color w:val=\"5D6675\"/><w:sz w:val=\"24\"/><w:szCs w:val=\"24\"/></w:rPr><w:t xml:space=\"preserve\">{label}</w:t></w:r></w:p></w:txbxContent></wps:txbx><wps:bodyPr/></wps:wsp>",
                emu(f64::from(*x)),
                emu(page_height - f64::from(*y) - f64::from(*height)),
                emu(f64::from(*width)).max(1),
                emu(f64::from(*height)).max(1),
            )
        }
        Primitive::Image {
            id,
            x,
            y,
            width,
            height,
            asset_id,
            transform,
            ..
        } => {
            let Some(media_index) = index_by_asset.get(asset_id) else {
                report.unsupported_images.push(id.clone());
                return image_placeholder(id, *x, *y, *width, *height, page_height);
            };
            let Some(placed) = place_rect(
                *x,
                *y,
                *width,
                *height,
                emit_matrix(*transform),
                page_height,
            ) else {
                report.skipped.push(id.clone());
                return String::new();
            };
            if placed.degraded {
                report.bbox_fallbacks.push(id.clone());
            }
            let name = escape_xml(id);
            let embed = format!("rId{}", media_index + 2);
            format!(
                "<pic:pic><pic:nvPicPr><pic:cNvPr id=\"0\" name=\"{name}\"/><pic:cNvPicPr/></pic:nvPicPr><pic:blipFill><a:blip r:embed=\"{embed}\"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill><pic:spPr>{}<a:prstGeom prst=\"rect\"><a:avLst/></a:prstGeom></pic:spPr></pic:pic>",
                xfrm_xml(&placed)
            )
        }
        Primitive::Group { .. } => String::new(),
    }
}

fn image_placeholder(
    id: &str,
    x: f32,
    y: f32,
    width: f32,
    height: f32,
    page_height: f64,
) -> String {
    let name = escape_xml(id);
    format!(
        "<wps:wsp><wps:cNvSpPr txBox=\"1\"/><wps:spPr><a:xfrm><a:off x=\"{}\" y=\"{}\"/><a:ext cx=\"{}\" cy=\"{}\"/></a:xfrm><a:prstGeom prst=\"rect\"><a:avLst/></a:prstGeom><a:noFill/><a:ln w=\"12700\"><a:prstDash val=\"dash\"/></a:ln></wps:spPr><wps:txbx><w:txbxContent><w:p><w:r><w:t>Unsupported image: {name}</w:t></w:r></w:p></w:txbxContent></wps:txbx><wps:bodyPr/></wps:wsp>",
        emu(f64::from(x)),
        emu(page_height - f64::from(y) - f64::from(height)),
        emu(f64::from(width)).max(1),
        emu(f64::from(height)).max(1)
    )
}

fn wml_runs(paragraphs: &[vsdx_render::TextParagraph]) -> String {
    let mut out = String::new();
    for paragraph in paragraphs {
        out.push_str("<w:p>");
        for run in &paragraph.runs {
            out.push_str("<w:r><w:rPr>");
            let family = escape_xml(&run.family);
            out.push_str(&format!(
                "<w:rFonts w:ascii=\"{family}\" w:hAnsi=\"{family}\" w:cs=\"{family}\"/>"
            ));
            if run.bold {
                out.push_str("<w:b/><w:bCs/>");
            }
            if run.italic {
                out.push_str("<w:i/><w:iCs/>");
            }
            if run.small_caps {
                out.push_str("<w:smallCaps/>");
            }
            if let Some(hex) = srgb_hex(&run.color) {
                out.push_str(&format!("<w:color w:val=\"{hex}\"/>"));
            }
            let half = (f64::from(run.size_in) * 144.0).round() as i64;
            out.push_str(&format!(
                "<w:sz w:val=\"{half}\"/><w:szCs w:val=\"{half}\"/>",
                half = half.max(2)
            ));
            if run.underline {
                out.push_str("<w:u w:val=\"single\"/>");
            }
            if run.superscript {
                out.push_str("<w:vertAlign w:val=\"superscript\"/>");
            }
            if run.subscript {
                out.push_str("<w:vertAlign w:val=\"subscript\"/>");
            }
            out.push_str("</w:rPr>");
            let text = escape_xml(&run.text);
            out.push_str(&format!("<w:t xml:space=\"preserve\">{text}</w:t></w:r>"));
        }
        out.push_str("</w:p>");
    }
    out
}

fn table_xml(rows: &[&crate::ShapeDatum]) -> String {
    let mut body = String::from(
        "<w:tbl><w:tblPr><w:tblW w:w=\"0\" w:type=\"auto\"/><w:tblBorders><w:top w:val=\"single\" w:sz=\"4\" w:color=\"8A94A6\"/><w:left w:val=\"single\" w:sz=\"4\" w:color=\"8A94A6\"/><w:bottom w:val=\"single\" w:sz=\"4\" w:color=\"8A94A6\"/><w:right w:val=\"single\" w:sz=\"4\" w:color=\"8A94A6\"/><w:insideH w:val=\"single\" w:sz=\"4\" w:color=\"8A94A6\"/><w:insideV w:val=\"single\" w:sz=\"4\" w:color=\"8A94A6\"/></w:tblBorders></w:tblPr><w:tblGrid><w:gridCol w:w=\"2000\"/><w:gridCol w:w=\"3000\"/><w:gridCol w:w=\"4360\"/></w:tblGrid>",
    );
    body.push_str(&table_row(&["Shape", "Label", "Value"], true));
    for datum in rows {
        body.push_str(&table_row(
            &[&datum.shape, &datum.label, &datum.value],
            false,
        ));
    }
    body.push_str("</w:tbl>");
    body
}

fn table_row(cells: &[&str], header: bool) -> String {
    let mut out = String::from("<w:tr>");
    for cell in cells {
        let text = escape_xml(cell);
        let run = if header {
            format!(
                "<w:r><w:rPr><w:b/><w:bCs/></w:rPr><w:t xml:space=\"preserve\">{text}</w:t></w:r>"
            )
        } else {
            format!("<w:r><w:t xml:space=\"preserve\">{text}</w:t></w:r>")
        };
        out.push_str(&format!("<w:tc><w:p>{run}</w:p></w:tc>"));
    }
    out.push_str("</w:tr>");
    out
}

fn content_types(media: &[Media]) -> Vec<u8> {
    let mut out = String::from(
        "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?><Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\"><Default Extension=\"rels\" ContentType=\"application/vnd.openxmlformats-package.relationships+xml\"/><Default Extension=\"xml\" ContentType=\"application/xml\"/>",
    );
    let mut extensions: Vec<String> = Vec::new();
    for item in media {
        let ext = item.part.rsplit('.').next().unwrap_or("bin").to_owned();
        if !extensions.contains(&ext) {
            extensions.push(ext);
        }
    }
    for ext in &extensions {
        let content_type = media
            .iter()
            .find(|item| item.part.ends_with(&format!(".{ext}")))
            .map(|item| item.content_type.clone())
            .unwrap_or_else(|| "application/octet-stream".to_owned());
        out.push_str(&format!(
            "<Default Extension=\"{ext}\" ContentType=\"{content_type}\"/>"
        ));
    }
    out.push_str("<Override PartName=\"/word/document.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml\"/><Override PartName=\"/word/styles.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml\"/><Override PartName=\"/docProps/core.xml\" ContentType=\"application/vnd.openxmlformats-package.core-properties+xml\"/><Override PartName=\"/docProps/app.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.extended-properties+xml\"/>");
    for item in media {
        out.push_str(&format!(
            "<Override PartName=\"/{}\" ContentType=\"{}\"/>",
            item.part, item.content_type
        ));
    }
    out.push_str("</Types>");
    out.into_bytes()
}

fn package_rels() -> Vec<u8> {
    "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?><Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\"><Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument\" Target=\"word/document.xml\"/><Relationship Id=\"rId2\" Type=\"http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties\" Target=\"docProps/core.xml\"/><Relationship Id=\"rId3\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties\" Target=\"docProps/app.xml\"/></Relationships>".to_owned().into_bytes()
}

fn document_rels(media: &[Media]) -> Vec<u8> {
    let mut out = String::from(
        "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?><Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\"><Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles\" Target=\"styles.xml\"/>",
    );
    for (index, item) in media.iter().enumerate() {
        let id = index + 2;
        let target = item.part.strip_prefix("word/").unwrap_or(&item.part);
        out.push_str(&format!("<Relationship Id=\"rId{id}\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/image\" Target=\"{target}\"/>"));
    }
    out.push_str("</Relationships>");
    out.into_bytes()
}

fn styles_xml() -> Vec<u8> {
    "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?><w:styles xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\"><w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii=\"Calibri\" w:hAnsi=\"Calibri\" w:cs=\"Calibri\"/><w:sz w:val=\"22\"/><w:szCs w:val=\"22\"/></w:rPr></w:rPrDefault><w:pPrDefault><w:pPr><w:spacing w:after=\"160\" w:line=\"259\" w:lineRule=\"auto\"/></w:pPr></w:pPrDefault></w:docDefaults><w:style w:type=\"paragraph\" w:default=\"1\" w:styleId=\"Normal\"><w:name w:val=\"Normal\"/><w:rPr><w:rFonts w:ascii=\"Calibri\" w:hAnsi=\"Calibri\" w:cs=\"Calibri\"/><w:sz w:val=\"22\"/><w:szCs w:val=\"22\"/></w:rPr></w:style><w:style w:type=\"paragraph\" w:styleId=\"Heading1\"><w:name w:val=\"heading 1\"/><w:basedOn w:val=\"Normal\"/><w:next w:val=\"Normal\"/><w:rPr><w:b/><w:bCs/><w:color w:val=\"1F2937\"/><w:sz w:val=\"32\"/><w:szCs w:val=\"32\"/></w:rPr></w:style><w:style w:type=\"paragraph\" w:styleId=\"Heading2\"><w:name w:val=\"heading 2\"/><w:basedOn w:val=\"Normal\"/><w:next w:val=\"Normal\"/><w:rPr><w:b/><w:bCs/><w:color w:val=\"374151\"/><w:sz w:val=\"26\"/><w:szCs w:val=\"26\"/></w:rPr></w:style></w:styles>".to_owned().into_bytes()
}

fn core_xml() -> Vec<u8> {
    "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?><cp:coreProperties xmlns:cp=\"http://schemas.openxmlformats.org/package/2006/metadata/core-properties\" xmlns:dc=\"http://purl.org/dc/elements/1.1/\"><dc:title>Diagram export</dc:title><dc:creator>BetterOffice</dc:creator><cp:revision>1</cp:revision></cp:coreProperties>".to_owned().into_bytes()
}

fn app_xml() -> Vec<u8> {
    "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?><Properties xmlns=\"http://schemas.openxmlformats.org/officeDocument/2006/extended-properties\"><Application>BetterOffice</Application></Properties>".to_owned().into_bytes()
}

#[cfg(test)]
mod tests {
    use super::*;
    use vsdx_render::Affine;

    #[test]
    fn image_member_keeps_its_placement() {
        let image = Primitive::Image {
            id: "image".into(),
            z_order: 1,
            asset_id: "asset".into(),
            x: 1.0,
            y: 2.0,
            width: 3.0,
            height: 4.0,
            transform: Affine {
                e: 5.0,
                f: 0.0,
                ..Affine::identity()
            },
        };
        let mut media = BTreeMap::new();
        media.insert("asset".into(), 0);
        let mut report = crate::ExportReport::default();
        let xml = member_xml(&image, 8.0, &media, &mut report);
        assert!(xml.contains("<pic:pic>"));
        assert!(xml.contains("x=\"5486400\""));
        assert!(report.summary().is_empty());
    }
}
