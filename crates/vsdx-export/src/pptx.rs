use std::collections::BTreeMap;

use ooxml_drawingml::{EMU_PER_INCH, cust_geom, emu, escape_xml, place_rect, xfrm_xml};
use vsdx_parse::VsdxPackage;
use vsdx_render::Primitive;

use crate::ExportError;
use crate::ExportReport;
use crate::Page;
use crate::shared::{
    Media, collect_media, emit_matrix, fallback_rect, fill_string, flat, line_string,
    paragraph_string, scale_list, translate_primitives, used_media,
};

pub fn build(
    pages: &[Page],
    package: &VsdxPackage,
    report: &mut ExportReport,
) -> Result<Vec<u8>, ExportError> {
    let mut media: Vec<Media> = Vec::new();
    let mut index_by_asset: BTreeMap<String, usize> = BTreeMap::new();
    for page in pages {
        collect_media(
            &page.list.primitives,
            package,
            &mut media,
            &mut index_by_asset,
            "ppt/media",
        );
    }
    let (max_w, max_h) = pages
        .first()
        .map(|page| presentation_size(page.width_in, page.height_in))
        .unwrap_or((11_433_600, 8_575_200));
    let mut parts: Vec<(String, Vec<u8>)> = vec![
        (
            "[Content_Types].xml".to_owned(),
            content_types(pages, &media),
        ),
        ("_rels/.rels".to_owned(), package_rels()),
        (
            "ppt/presentation.xml".to_owned(),
            presentation_xml(pages, max_w, max_h),
        ),
        (
            "ppt/_rels/presentation.xml.rels".to_owned(),
            presentation_rels(pages),
        ),
        ("ppt/slideMasters/slideMaster1.xml".to_owned(), master_xml()),
        (
            "ppt/slideMasters/_rels/slideMaster1.xml.rels".to_owned(),
            master_rels(),
        ),
        ("ppt/slideLayouts/slideLayout1.xml".to_owned(), layout_xml()),
        (
            "ppt/slideLayouts/_rels/slideLayout1.xml.rels".to_owned(),
            layout_rels(),
        ),
        ("ppt/theme/theme1.xml".to_owned(), theme_xml()),
        ("docProps/core.xml".to_owned(), core_xml()),
        ("docProps/app.xml".to_owned(), app_xml(pages)),
    ];
    for (slide_index, page) in pages.iter().enumerate() {
        let number = slide_index + 1;
        let used = used_media(&page.list.primitives, &index_by_asset);
        let (scale, dx, dy) = fit_placement(page.width_in, page.height_in, max_w, max_h);
        let mut list = scale_list(&page.list, scale);
        translate_primitives(&mut list.primitives, dx as f32, dy as f32);
        parts.push((
            format!("ppt/slides/slide{number}.xml"),
            slide_xml(
                &page.name,
                &list,
                &index_by_asset,
                max_h as f64 / 914400.0,
                report,
            ),
        ));
        parts.push((
            format!("ppt/slides/_rels/slide{number}.xml.rels"),
            slide_rels(&media, &used),
        ));
    }
    for item in &media {
        parts.push((item.part.clone(), item.bytes.clone()));
    }
    ooxml_opc::rezip_parts(&parts).map_err(ExportError::Package)
}

fn slide_xml(
    name: &str,
    list: &vsdx_render::VsdxDisplayList,
    index_by_asset: &BTreeMap<String, usize>,
    page_height: f64,
    report: &mut ExportReport,
) -> Vec<u8> {
    let mut shapes = String::new();
    let mut next_id: u32 = 2;
    for primitive in &flat(&list.primitives) {
        match primitive {
            Primitive::Shape {
                id,
                path,
                fill,
                stroke,
                ..
            } => {
                let name = escape_xml(id);
                let fill = fill_string(fill);
                let line = line_string(stroke);
                let id_value = next_id;
                next_id += 1;
                if let Some(geom) = cust_geom(path, page_height) {
                    shapes.push_str(&format!(
                        "<p:sp><p:nvSpPr><p:cNvPr id=\"{id_value}\" name=\"{name}\"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:spPr><a:xfrm><a:off x=\"{}\" y=\"{}\"/><a:ext cx=\"{}\" cy=\"{}\"/></a:xfrm><a:custGeom><a:avLst/><a:gdLst/><a:ahLst/><a:cxnLst/><a:pathLst><a:path w=\"{}\" h=\"{}\">{}</a:path></a:pathLst></a:custGeom>{fill}{line}</p:spPr></p:sp>",
                        geom.x, geom.y, geom.w, geom.h, geom.w, geom.h, geom.body
                    ));
                    continue;
                }
                report.empty_geometry.push(id.clone());
                let Some((x, y, w, h)) = fallback_rect(path, page_height) else {
                    report.skipped.push(id.clone());
                    continue;
                };
                shapes.push_str(&format!(
                    "<p:sp><p:nvSpPr><p:cNvPr id=\"{id_value}\" name=\"{name}\"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:spPr><a:xfrm><a:off x=\"{x}\" y=\"{y}\"/><a:ext cx=\"{w}\" cy=\"{h}\"/></a:xfrm><a:prstGeom prst=\"rect\"><a:avLst/></a:prstGeom>{fill}{line}</p:spPr></p:sp>"
                ));
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
                    continue;
                };
                if placed.degraded {
                    report.bbox_fallbacks.push(id.clone());
                }
                let name = escape_xml(id);
                let mut body = paragraph_string(paragraphs);
                if body.is_empty() {
                    body.push_str("<a:p/>");
                }
                let id_value = next_id;
                next_id += 1;
                shapes.push_str(&format!(
                    "<p:sp><p:nvSpPr><p:cNvPr id=\"{id_value}\" name=\"{name}\"/><p:cNvSpPr><a:spLocks noChangeArrowheads=\"1\"/></p:cNvSpPr><p:nvPr/></p:nvSpPr><p:spPr>{}<a:prstGeom prst=\"rect\"><a:avLst/></a:prstGeom><a:noFill/></p:spPr><p:txBody><a:bodyPr/><a:lstStyle/>{body}</p:txBody></p:sp>",
                    xfrm_xml(&placed)
                ));
            }
            Primitive::Placeholder {
                id,
                x,
                y,
                width,
                height,
                reason,
                ..
            } => {
                let placed = ooxml_drawingml::Placed {
                    x: emu(f64::from(*x)),
                    y: emu(page_height - f64::from(*y) - f64::from(*height)),
                    w: emu(f64::from(*width)).max(1),
                    h: emu(f64::from(*height)).max(1),
                    rot: 0,
                    flip_h: false,
                    flip_v: false,
                    degraded: false,
                };
                let id_value = next_id;
                next_id += 1;
                shapes.push_str(&placeholder_xml(id, reason, &placed, id_value));
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
                    shapes.push_str(&image_placeholder(
                        id,
                        (*x, *y, *width, *height),
                        *transform,
                        page_height,
                        next_id,
                        report,
                    ));
                    next_id += 1;
                    continue;
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
                    continue;
                };
                if placed.degraded {
                    report.bbox_fallbacks.push(id.clone());
                }
                let name = escape_xml(id);
                let embed = format!("rId{}", media_index + 2);
                let id_value = next_id;
                next_id += 1;
                shapes.push_str(&format!(
                    "<p:pic><p:nvPicPr><p:cNvPr id=\"{id_value}\" name=\"{name}\"/><p:cNvPicPr><a:picLocks noChangeAspect=\"1\"/></p:cNvPicPr><p:nvPr/></p:nvPicPr><p:blipFill><a:blip r:embed=\"{embed}\"/><a:stretch><a:fillRect/></a:stretch></p:blipFill><p:spPr>{}<a:prstGeom prst=\"rect\"><a:avLst/></a:prstGeom></p:spPr></p:pic>",
                    xfrm_xml(&placed)
                ));
            }
            Primitive::Group { .. } => {}
        }
    }
    let name = escape_xml(name);
    format!(
        "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?><p:sld xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\" xmlns:p=\"http://schemas.openxmlformats.org/presentationml/2006/main\"><p:cSld name=\"{name}\"><p:spTree><p:nvGrpSpPr><p:cNvPr id=\"1\" name=\"\"/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr><a:xfrm><a:off x=\"0\" y=\"0\"/><a:ext cx=\"0\" cy=\"0\"/><a:chOff x=\"0\" y=\"0\"/><a:chExt cx=\"0\" cy=\"0\"/></a:xfrm></p:grpSpPr>{shapes}</p:spTree></p:cSld><p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:sld>"
    )
    .into_bytes()
}

fn image_placeholder(
    id: &str,
    rect: (f32, f32, f32, f32),
    transform: vsdx_render::Affine,
    page_height: f64,
    id_value: u32,
    report: &mut ExportReport,
) -> String {
    let (x, y, width, height) = rect;
    let Some(placed) = place_rect(x, y, width, height, emit_matrix(transform), page_height) else {
        report.skipped.push(id.to_owned());
        return String::new();
    };
    if placed.degraded {
        report.bbox_fallbacks.push(id.to_owned());
    }
    placeholder_xml(id, &format!("Unsupported image: {id}"), &placed, id_value)
}

fn placeholder_xml(
    id: &str,
    label: &str,
    placed: &ooxml_drawingml::Placed,
    id_value: u32,
) -> String {
    let name = escape_xml(id);
    let label = escape_xml(label);
    format!(
        "<p:sp><p:nvSpPr><p:cNvPr id=\"{id_value}\" name=\"{name}\"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:spPr>{}<a:prstGeom prst=\"rect\"><a:avLst/></a:prstGeom><a:noFill/><a:ln w=\"12700\"><a:solidFill><a:srgbClr val=\"8A94A6\"/></a:solidFill><a:prstDash val=\"dash\"/></a:ln></p:spPr><p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:rPr sz=\"1200\"><a:solidFill><a:srgbClr val=\"5D6675\"/></a:solidFill><a:latin typeface=\"Calibri\"/></a:rPr><a:t>{label}</a:t></a:r></a:p></p:txBody></p:sp>",
        xfrm_xml(placed)
    )
}

/// Slide canvas in EMUs, taken from the first page; later pages fit inside.
fn presentation_size(width_in: f64, height_in: f64) -> (i64, i64) {
    (emu(width_in), emu(height_in))
}

/// Uniform fit scale plus centring offsets in inches for a page on canvas.
fn fit_placement(
    page_w_in: f64,
    page_h_in: f64,
    canvas_w_emu: i64,
    canvas_h_emu: i64,
) -> (f64, f64, f64) {
    let scale = (canvas_w_emu as f64 / emu(page_w_in) as f64)
        .min(canvas_h_emu as f64 / emu(page_h_in) as f64);
    let dx = (canvas_w_emu as f64 - emu(page_w_in * scale) as f64) / EMU_PER_INCH / 2.0;
    let dy = (canvas_h_emu as f64 - emu(page_h_in * scale) as f64) / EMU_PER_INCH / 2.0;
    (scale, dx, dy)
}

fn content_types(pages: &[Page], media: &[Media]) -> Vec<u8> {
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
    out.push_str("<Override PartName=\"/ppt/presentation.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml\"/><Override PartName=\"/ppt/slideMasters/slideMaster1.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.presentationml.slideMaster+xml\"/><Override PartName=\"/ppt/slideLayouts/slideLayout1.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml\"/><Override PartName=\"/ppt/theme/theme1.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.theme+xml\"/><Override PartName=\"/docProps/core.xml\" ContentType=\"application/vnd.openxmlformats-package.core-properties+xml\"/><Override PartName=\"/docProps/app.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.extended-properties+xml\"/>");
    for (index, _) in pages.iter().enumerate() {
        let number = index + 1;
        out.push_str(&format!("<Override PartName=\"/ppt/slides/slide{number}.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.presentationml.slide+xml\"/>"));
    }
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
    "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?><Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\"><Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument\" Target=\"ppt/presentation.xml\"/><Relationship Id=\"rId2\" Type=\"http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties\" Target=\"docProps/core.xml\"/><Relationship Id=\"rId3\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties\" Target=\"docProps/app.xml\"/></Relationships>".to_owned().into_bytes()
}

fn presentation_xml(pages: &[Page], max_w: i64, max_h: i64) -> Vec<u8> {
    let mut ids = String::new();
    for (index, _) in pages.iter().enumerate() {
        let number = index + 1;
        let id = 256 + index as u32;
        ids.push_str(&format!("<p:sldId id=\"{id}\" r:id=\"rId{number}\"/>"));
    }
    format!("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?><p:presentation xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\" xmlns:p=\"http://schemas.openxmlformats.org/presentationml/2006/main\"><p:sldMasterIdLst><p:sldMasterId id=\"2147483648\" r:id=\"rId{}\"/></p:sldMasterIdLst><p:sldIdLst>{ids}</p:sldIdLst><p:sldSz cx=\"{max_w}\" cy=\"{max_h}\"/><p:notesSz cx=\"6858000\" cy=\"9144000\"/></p:presentation>", pages.len() + 1).into_bytes()
}

fn presentation_rels(pages: &[Page]) -> Vec<u8> {
    let mut out = String::from(
        "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?><Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">",
    );
    for (index, _) in pages.iter().enumerate() {
        let number = index + 1;
        out.push_str(&format!("<Relationship Id=\"rId{number}\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide\" Target=\"slides/slide{number}.xml\"/>"));
    }
    let master = pages.len() + 1;
    out.push_str(&format!("<Relationship Id=\"rId{master}\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster\" Target=\"slideMasters/slideMaster1.xml\"/></Relationships>"));
    out.into_bytes()
}

fn slide_rels(media: &[Media], used: &[usize]) -> Vec<u8> {
    let mut out = String::from(
        "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?><Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\"><Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout\" Target=\"../slideLayouts/slideLayout1.xml\"/>",
    );
    for index in used {
        let item = &media[*index];
        let id = index + 2;
        let relative = relative_media(&item.part);
        out.push_str(&format!("<Relationship Id=\"rId{id}\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/image\" Target=\"{relative}\"/>"));
    }
    out.push_str("</Relationships>");
    out.into_bytes()
}

fn relative_media(part: &str) -> String {
    part.strip_prefix("ppt/")
        .map(|rest| format!("../{rest}"))
        .unwrap_or_else(|| part.to_owned())
}

fn master_xml() -> Vec<u8> {
    "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?><p:sldMaster xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\" xmlns:p=\"http://schemas.openxmlformats.org/presentationml/2006/main\"><p:cSld><p:bg><p:bgPr><a:solidFill><a:srgbClr val=\"FFFFFF\"/></a:solidFill><a:effectLst/></p:bgPr></p:bg><p:spTree><p:nvGrpSpPr><p:cNvPr id=\"1\" name=\"\"/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr><a:xfrm><a:off x=\"0\" y=\"0\"/><a:ext cx=\"0\" cy=\"0\"/><a:chOff x=\"0\" y=\"0\"/><a:chExt cx=\"0\" cy=\"0\"/></a:xfrm></p:grpSpPr></p:spTree></p:cSld><p:clrMap bg1=\"lt1\" tx1=\"dk1\" bg2=\"lt2\" tx2=\"dk2\" accent1=\"accent1\" accent2=\"accent2\" accent3=\"accent3\" accent4=\"accent4\" accent5=\"accent5\" accent6=\"accent6\" hlink=\"hlink\" folHlink=\"folHlink\"/><p:sldLayoutIdLst><p:sldLayoutId id=\"2147483649\" r:id=\"rId1\"/></p:sldLayoutIdLst><p:txStyles><p:titleStyle><a:lvl1pPr><a:defRPr sz=\"4400\" b=\"1\"><a:solidFill><a:schemeClr val=\"tx1\"/></a:solidFill><a:latin typeface=\"Calibri\"/></a:defRPr></a:lvl1pPr></p:titleStyle><p:bodyStyle><a:lvl1pPr><a:defRPr sz=\"3200\"><a:solidFill><a:schemeClr val=\"tx1\"/></a:solidFill><a:latin typeface=\"Calibri\"/></a:defRPr></a:lvl1pPr></p:bodyStyle><p:otherStyle><a:lvl1pPr><a:defRPr sz=\"1800\"><a:solidFill><a:schemeClr val=\"tx1\"/></a:solidFill><a:latin typeface=\"Calibri\"/></a:defRPr></a:lvl1pPr></p:otherStyle></p:txStyles></p:sldMaster>".to_owned().into_bytes()
}

fn master_rels() -> Vec<u8> {
    "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?><Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\"><Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout\" Target=\"../slideLayouts/slideLayout1.xml\"/><Relationship Id=\"rId2\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme\" Target=\"../theme/theme1.xml\"/></Relationships>".to_owned().into_bytes()
}

fn layout_xml() -> Vec<u8> {
    "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?><p:sldLayout xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\" xmlns:p=\"http://schemas.openxmlformats.org/presentationml/2006/main\" type=\"blank\" preserve=\"1\"><p:cSld name=\"Blank\"><p:spTree><p:nvGrpSpPr><p:cNvPr id=\"1\" name=\"\"/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr><a:xfrm><a:off x=\"0\" y=\"0\"/><a:ext cx=\"0\" cy=\"0\"/><a:chOff x=\"0\" y=\"0\"/><a:chExt cx=\"0\" cy=\"0\"/></a:xfrm></p:grpSpPr></p:spTree></p:cSld><p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:sldLayout>".to_owned().into_bytes()
}

fn layout_rels() -> Vec<u8> {
    "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?><Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\"><Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster\" Target=\"../slideMasters/slideMaster1.xml\"/></Relationships>".to_owned().into_bytes()
}

fn theme_xml() -> Vec<u8> {
    "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?><a:theme xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\" name=\"Export\"><a:themeElements><a:clrScheme name=\"Export\"><a:dk1><a:sysClr val=\"windowText\" lastClr=\"000000\"/></a:dk1><a:lt1><a:sysClr val=\"window\" lastClr=\"FFFFFF\"/></a:lt1><a:dk2><a:srgbClr val=\"1F2937\"/></a:dk2><a:lt2><a:srgbClr val=\"E5E7EB\"/></a:lt2><a:accent1><a:srgbClr val=\"2563EB\"/></a:accent1><a:accent2><a:srgbClr val=\"059669\"/></a:accent2><a:accent3><a:srgbClr val=\"D97706\"/></a:accent3><a:accent4><a:srgbClr val=\"DC2626\"/></a:accent4><a:accent5><a:srgbClr val=\"7C3AED\"/></a:accent5><a:accent6><a:srgbClr val=\"0891B2\"/></a:accent6><a:hlink><a:srgbClr val=\"2563EB\"/></a:hlink><a:folHlink><a:srgbClr val=\"7C3AED\"/></a:folHlink></a:clrScheme><a:fontScheme name=\"Export\"><a:majorFont><a:latin typeface=\"Calibri\"/><a:ea typeface=\"\"/><a:cs typeface=\"\"/></a:majorFont><a:minorFont><a:latin typeface=\"Calibri\"/><a:ea typeface=\"\"/><a:cs typeface=\"\"/></a:minorFont></a:fontScheme><a:fmtScheme name=\"Export\"><a:fillStyleLst><a:solidFill><a:schemeClr val=\"phClr\"/></a:solidFill><a:gradFill rotWithShape=\"1\"><a:gsLst><a:gs pos=\"0\"><a:schemeClr val=\"phClr\"/></a:gs><a:gs pos=\"50000\"><a:schemeClr val=\"phClr\"/></a:gs><a:gs pos=\"100000\"><a:schemeClr val=\"phClr\"/></a:gs></a:gsLst><a:lin ang=\"5400000\" scaled=\"0\"/></a:gradFill><a:gradFill rotWithShape=\"1\"><a:gsLst><a:gs pos=\"0\"><a:schemeClr val=\"phClr\"/></a:gs><a:gs pos=\"50000\"><a:schemeClr val=\"phClr\"/></a:gs><a:gs pos=\"100000\"><a:schemeClr val=\"phClr\"/></a:gs></a:gsLst><a:lin ang=\"5400000\" scaled=\"0\"/></a:gradFill></a:fillStyleLst><a:lnStyleLst><a:ln w=\"12700\" cap=\"flat\" cmpd=\"sng\" algn=\"ctr\"><a:solidFill><a:schemeClr val=\"phClr\"/></a:solidFill><a:prstDash val=\"solid\"/><a:miter lim=\"800000\"/></a:ln><a:ln w=\"19050\" cap=\"flat\" cmpd=\"sng\" algn=\"ctr\"><a:solidFill><a:schemeClr val=\"phClr\"/></a:solidFill><a:prstDash val=\"solid\"/><a:miter lim=\"800000\"/></a:ln><a:ln w=\"25400\" cap=\"flat\" cmpd=\"sng\" algn=\"ctr\"><a:solidFill><a:schemeClr val=\"phClr\"/></a:solidFill><a:prstDash val=\"solid\"/><a:miter lim=\"800000\"/></a:ln></a:lnStyleLst><a:effectStyleLst><a:effectStyle><a:effectLst/></a:effectStyle><a:effectStyle><a:effectLst/></a:effectStyle><a:effectStyle><a:effectLst><a:outerShdw blurRad=\"57150\" dist=\"19050\" dir=\"5400000\" algn=\"ctr\" rotWithShape=\"0\"><a:srgbClr val=\"000000\"><a:alpha val=\"63000\"/></a:srgbClr></a:outerShdw></a:effectLst></a:effectStyle></a:effectStyleLst><a:bgFillStyleLst><a:solidFill><a:schemeClr val=\"phClr\"/></a:solidFill><a:solidFill><a:schemeClr val=\"phClr\"/></a:solidFill><a:gradFill rotWithShape=\"1\"><a:gsLst><a:gs pos=\"0\"><a:schemeClr val=\"phClr\"/></a:gs><a:gs pos=\"50000\"><a:schemeClr val=\"phClr\"/></a:gs><a:gs pos=\"100000\"><a:schemeClr val=\"phClr\"/></a:gs></a:gsLst><a:lin ang=\"5400000\" scaled=\"0\"/></a:gradFill></a:bgFillStyleLst></a:fmtScheme></a:themeElements></a:theme>".to_owned().into_bytes()
}

fn core_xml() -> Vec<u8> {
    "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?><cp:coreProperties xmlns:cp=\"http://schemas.openxmlformats.org/package/2006/metadata/core-properties\" xmlns:dc=\"http://purl.org/dc/elements/1.1/\"><dc:title>Diagram export</dc:title><dc:creator>BetterOffice</dc:creator><cp:revision>1</cp:revision></cp:coreProperties>".to_owned().into_bytes()
}

fn app_xml(pages: &[Page]) -> Vec<u8> {
    format!("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?><Properties xmlns=\"http://schemas.openxmlformats.org/officeDocument/2006/extended-properties\"><Application>BetterOffice</Application><PresentationFormat>On-screen Show (4:3)</PresentationFormat><Slides>{}</Slides></Properties>", pages.len()).into_bytes()
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn canvas_size_comes_from_the_first_page() {
        assert_eq!(presentation_size(10.0, 8.0), (9_144_000, 7_315_200));
    }

    #[test]
    fn fitted_page_centres_on_the_canvas() {
        let (scale, dx, dy) = fit_placement(4.0, 12.0, 9_144_000, 7_315_200);
        assert!((scale - 2.0 / 3.0).abs() < 1e-9);
        assert!(dx > 3.6 && dx < 3.7);
        assert_eq!(dy, 0.0);
    }

    #[test]
    fn matching_page_needs_no_offset() {
        let (scale, dx, dy) = fit_placement(10.0, 8.0, 9_144_000, 7_315_200);
        assert_eq!((scale, dx, dy), (1.0, 0.0, 0.0));
    }
}
