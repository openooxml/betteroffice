//! The raster pictures a sheet's drawings anchor, as a frame draws them.

use xlsx_model::{PictureCrop, SheetPicture};

use crate::MAX_CHART_ANCHORS;
use crate::chart::{
    NS_RELATIONSHIPS, NS_SPREADSHEET_DRAWING, anchor_geometry, directory_of, is_anchor,
    parse_relationships, relationship_part_path, sheet_drawings, type_is,
};
use crate::package::PreservedPackage;
use crate::tree::{Element, parse_tree};
use crate::xml::{find_part, resolve_part_path};

/// Every picture the drawings of the sheet stored at `sheet_path` anchor directly and do not hide,
/// in drawing then anchor order. `parts` must carry the sheet's relationship part, its drawing
/// parts and their relationship parts; one that is missing or does not read contributes nothing.
pub fn sheet_pictures(parts: &[(String, Vec<u8>)], sheet_path: &str) -> Vec<SheetPicture> {
    let Ok(drawings) = sheet_drawings(parts, sheet_path) else {
        return Vec::new();
    };
    let mut pictures = Vec::new();
    for (drawing, bytes) in drawings {
        let Ok(root) = parse_tree(bytes) else {
            continue;
        };
        if !root.is(NS_SPREADSHEET_DRAWING, "wsDr") {
            continue;
        }
        let relationships = find_part(parts, &relationship_part_path(&drawing))
            .and_then(|rels| parse_relationships(rels).ok())
            .unwrap_or_default();
        let base = directory_of(&drawing);
        let anchors = root.child_elements().filter(is_anchor);
        for (anchor_index, element) in anchors.enumerate().take(MAX_CHART_ANCHORS) {
            let Some(pic) = element
                .child_elements()
                .find(|child| child.is(NS_SPREADSHEET_DRAWING, "pic"))
            else {
                continue;
            };
            let hidden = pic
                .child("nvPicPr")
                .and_then(|properties| properties.child("cNvPr"))
                .and_then(|properties| properties.attribute_local("hidden"))
                .is_some_and(|hidden| matches!(hidden, "1" | "true"));
            let Some(fill) = pic.child("blipFill") else {
                continue;
            };
            let media = fill
                .child("blip")
                .and_then(|blip| blip.attribute_ns(NS_RELATIONSHIPS, "embed"))
                .and_then(|id| {
                    relationships
                        .iter()
                        .find(|(rel_id, kind, _)| rel_id == id && type_is(kind, "image"))
                })
                .map(|(_, _, target)| resolve_part_path(base, target));
            let (false, Some(media), Ok(anchor)) = (hidden, media, anchor_geometry(element)) else {
                continue;
            };
            pictures.push(SheetPicture {
                media,
                drawing: drawing.clone(),
                anchor_index,
                anchor,
                crop: fill.child("srcRect").map(source_rect).unwrap_or_default(),
            });
        }
    }
    pictures
}

fn source_rect(rect: &Element) -> PictureCrop {
    let inset = |name: &str| {
        rect.attribute_local(name)
            .and_then(|value| value.trim().parse().ok())
            .unwrap_or(0)
    };
    PictureCrop {
        left: inset("l"),
        top: inset("t"),
        right: inset("r"),
        bottom: inset("b"),
    }
}

impl PreservedPackage {
    /// [`sheet_pictures`] of source sheet `index`.
    pub fn source_sheet_pictures(&self, index: usize) -> Vec<SheetPicture> {
        self.source_sheet_part(index)
            .map_or_else(Vec::new, |path| sheet_pictures(&self.parts, path))
    }
}

#[cfg(test)]
mod tests {
    use xlsx_model::{AnchorCell, AnchorEditAs, AnchorExtent, ChartAnchor};

    use super::*;

    const RELS: &str = "http://schemas.openxmlformats.org/package/2006/relationships";
    const DOC_RELS: &str = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

    fn rels(entries: &[(&str, &str, &str)]) -> Vec<u8> {
        let entries: String = entries
            .iter()
            .map(|(id, kind, target)| {
                format!(r#"<Relationship Id="{id}" Type="{DOC_RELS}/{kind}" Target="{target}"/>"#)
            })
            .collect();
        format!(r#"<Relationships xmlns="{RELS}">{entries}</Relationships>"#).into_bytes()
    }

    fn pic(embed: &str, extra: &str) -> String {
        format!(
            r#"<xdr:pic><xdr:nvPicPr><xdr:cNvPr id="2" name="Logo"{extra}/><xdr:cNvPicPr/></xdr:nvPicPr><xdr:blipFill><a:blip r:embed="{embed}"/><a:srcRect l="10000" r="-5000"/><a:stretch/></xdr:blipFill><xdr:spPr/></xdr:pic><xdr:clientData/>"#
        )
    }

    fn parts(anchors: &str) -> Vec<(String, Vec<u8>)> {
        vec![
            (
                "xl/worksheets/_rels/sheet1.xml.rels".into(),
                rels(&[("d", "drawing", "../drawings/drawing1.xml")]),
            ),
            (
                "xl/drawings/drawing1.xml".into(),
                format!(
                    r#"<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="{DOC_RELS}">{anchors}</xdr:wsDr>"#
                )
                .into_bytes(),
            ),
            (
                "xl/drawings/_rels/drawing1.xml.rels".into(),
                rels(&[
                    ("img", "image", "../media/image1.png"),
                    ("chart", "chart", "../charts/chart1.xml"),
                ]),
            ),
        ]
    }

    fn marker(tag: &str, col: u32, row: u32) -> String {
        format!(
            "<xdr:{tag}><xdr:col>{col}</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>{row}</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:{tag}>"
        )
    }

    #[test]
    fn reads_visible_pictures_with_their_anchor_order_media_and_crop() {
        let anchors = [
            format!(
                "<xdr:twoCellAnchor editAs=\"oneCell\">{}{}{}</xdr:twoCellAnchor>",
                marker("from", 1, 2),
                marker("to", 3, 5),
                pic("img", "")
            ),
            format!(
                "<xdr:oneCellAnchor>{}<xdr:ext cx=\"914400\" cy=\"457200\"/><xdr:graphicFrame><a:graphic><a:graphicData><c:chart xmlns:c=\"http://schemas.openxmlformats.org/drawingml/2006/chart\" r:id=\"chart\"/></a:graphicData></a:graphic></xdr:graphicFrame><xdr:clientData/></xdr:oneCellAnchor>",
                marker("from", 0, 0)
            ),
            format!(
                "<xdr:oneCellAnchor>{}<xdr:ext cx=\"914400\" cy=\"457200\"/>{}</xdr:oneCellAnchor>",
                marker("from", 4, 4),
                pic("img", " hidden=\"1\"")
            ),
            format!(
                "<xdr:oneCellAnchor>{}<xdr:ext cx=\"914400\" cy=\"457200\"/>{}</xdr:oneCellAnchor>",
                marker("from", 6, 7),
                pic("img", "")
            ),
            format!(
                "<xdr:oneCellAnchor>{}<xdr:ext cx=\"1\" cy=\"1\"/>{}</xdr:oneCellAnchor>",
                marker("from", 0, 0),
                pic("missing", "")
            ),
        ]
        .concat();
        let pictures = sheet_pictures(&parts(&anchors), "xl/worksheets/sheet1.xml");
        let cell = |col, row| AnchorCell {
            col,
            col_off: 0,
            row,
            row_off: 0,
        };
        let crop = PictureCrop {
            left: 10_000,
            right: -5_000,
            ..PictureCrop::default()
        };
        assert_eq!(
            pictures,
            vec![
                SheetPicture {
                    media: "xl/media/image1.png".into(),
                    drawing: "xl/drawings/drawing1.xml".into(),
                    anchor_index: 0,
                    anchor: ChartAnchor::TwoCell {
                        from: cell(1, 2),
                        to: cell(3, 5),
                        edit_as: AnchorEditAs::OneCell,
                    },
                    crop,
                },
                SheetPicture {
                    media: "xl/media/image1.png".into(),
                    drawing: "xl/drawings/drawing1.xml".into(),
                    anchor_index: 3,
                    anchor: ChartAnchor::OneCell {
                        from: cell(6, 7),
                        extent: AnchorExtent {
                            cx: 914_400,
                            cy: 457_200,
                        },
                    },
                    crop,
                },
            ]
        );
    }

    #[test]
    fn a_sheet_without_drawings_has_no_pictures() {
        assert!(sheet_pictures(&[], "xl/worksheets/sheet1.xml").is_empty());
    }
}
