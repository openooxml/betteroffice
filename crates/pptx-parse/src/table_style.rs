//! Parses `ppt/tableStyles.xml`.

use ooxml_drawingml::{
    TableCellBorder, TableCellBorders, TableCellStyle, TableStyle, TableStyleList, TableStylePart,
    TableTextStyle,
};

use crate::drawing::{parse_color_container, parse_fill_element, parse_outline_element};
use crate::xml::XmlElement;

pub(crate) fn parse_table_styles(root: &XmlElement) -> TableStyleList {
    TableStyleList {
        default_style_id: root.attribute("def").map(str::to_owned),
        styles: root
            .children_named("tblStyle")
            .map(parse_table_style)
            .collect(),
    }
}

fn parse_table_style(element: &XmlElement) -> TableStyle {
    TableStyle {
        style_id: element.attribute("styleId").unwrap_or_default().to_owned(),
        style_name: element.attribute("styleName").map(str::to_owned),
        whole_table: element.child("wholeTbl").map(parse_part),
        band1_row: element.child("band1H").map(parse_part),
        band2_row: element.child("band2H").map(parse_part),
        first_row: element.child("firstRow").map(parse_part),
        last_row: element.child("lastRow").map(parse_part),
        first_column: element.child("firstCol").map(parse_part),
        last_column: element.child("lastCol").map(parse_part),
    }
}

fn parse_part(element: &XmlElement) -> TableStylePart {
    TableStylePart {
        text: element
            .child("tcTxStyle")
            .map(parse_text_style)
            .unwrap_or_default(),
        cell: element
            .child("tcStyle")
            .map(parse_cell_style)
            .unwrap_or_default(),
    }
}

fn parse_text_style(element: &XmlElement) -> TableTextStyle {
    TableTextStyle {
        bold: on_off(element.attribute("b")),
        italic: on_off(element.attribute("i")),
        color: parse_color_container(element),
    }
}

fn parse_cell_style(element: &XmlElement) -> TableCellStyle {
    TableCellStyle {
        fill: element
            .child("fill")
            .and_then(|fill| fill.child_elements().find_map(parse_fill_element)),
        borders: element
            .child("tcBdr")
            .map(parse_borders)
            .unwrap_or_default(),
    }
}

fn parse_borders(element: &XmlElement) -> TableCellBorders {
    TableCellBorders {
        left: element.child("left").and_then(parse_border),
        right: element.child("right").and_then(parse_border),
        top: element.child("top").and_then(parse_border),
        bottom: element.child("bottom").and_then(parse_border),
        inside_horizontal: element.child("insideH").and_then(parse_border),
        inside_vertical: element.child("insideV").and_then(parse_border),
    }
}

fn parse_border(element: &XmlElement) -> Option<TableCellBorder> {
    let line = element.child("ln")?;
    if line.child("noFill").is_some() {
        return Some(TableCellBorder::None);
    }
    parse_outline_element(line).map(|outline| TableCellBorder::Line(Box::new(outline)))
}

fn on_off(value: Option<&str>) -> Option<bool> {
    match value? {
        "on" | "1" | "true" => Some(true),
        "off" | "0" | "false" => Some(false),
        _ => None,
    }
}

#[cfg(test)]
mod tests {
    use ooxml_drawingml::TableStyle;

    use super::*;
    use crate::xml::{ParseBudget, ParseLimits, parse_xml};

    const NO_GRID: &[u8] = br#"<a:tblStyleLst xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" def="{DEF}"><a:tblStyle styleId="{NOGRID}" styleName="No Style, No Grid"><a:wholeTbl><a:tcTxStyle b="off" i="def"/><a:tcStyle><a:tcBdr><a:left><a:ln><a:noFill/></a:ln></a:left><a:right><a:ln><a:noFill/></a:ln></a:right><a:top><a:ln><a:noFill/></a:ln></a:top><a:bottom><a:ln><a:noFill/></a:ln></a:bottom><a:insideH><a:ln><a:noFill/></a:ln></a:insideH><a:insideV><a:ln><a:noFill/></a:ln></a:insideV></a:tcBdr><a:fill><a:noFill/></a:fill></a:tcStyle></a:wholeTbl></a:tblStyle></a:tblStyleLst>"#;

    const REGIONS: &[u8] = br#"<a:tblStyleLst xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:tblStyle styleId="{R}"><a:band2H><a:tcStyle><a:fill><a:solidFill><a:srgbClr val="B2B2B2"/></a:solidFill></a:fill></a:tcStyle></a:band2H><a:lastRow><a:tcStyle><a:fill><a:solidFill><a:srgbClr val="111111"/></a:solidFill></a:fill></a:tcStyle></a:lastRow><a:firstCol><a:tcStyle><a:fill><a:solidFill><a:srgbClr val="222222"/></a:solidFill></a:fill></a:tcStyle></a:firstCol><a:lastCol><a:tcStyle><a:fill><a:solidFill><a:srgbClr val="333333"/></a:solidFill></a:fill></a:tcStyle></a:lastCol></a:tblStyle></a:tblStyleLst>"#;

    fn styles(xml: &[u8]) -> TableStyleList {
        let limits = ParseLimits::default();
        let mut budget = ParseBudget::new(&limits);
        let root = parse_xml(xml, "ppt/tableStyles.xml", &mut budget).unwrap();
        parse_table_styles(&root)
    }

    fn only(xml: &[u8]) -> TableStyle {
        styles(xml).styles.into_iter().next().unwrap()
    }

    fn fill_of(part: Option<&TableStylePart>) -> Option<&str> {
        part?.cell.fill.as_ref()?.color.as_ref()?.rgb.as_deref()
    }

    #[test]
    fn every_no_fill_edge_parses_as_an_explicit_clear() {
        let style = only(NO_GRID);
        let borders = &style.whole_table.as_ref().unwrap().cell.borders;

        for edge in [
            &borders.left,
            &borders.right,
            &borders.top,
            &borders.bottom,
            &borders.inside_horizontal,
            &borders.inside_vertical,
        ] {
            assert_eq!(edge.as_ref(), Some(&TableCellBorder::None));
        }
    }

    #[test]
    fn a_no_fill_cell_fill_survives_as_an_explicit_none() {
        let style = only(NO_GRID);
        let fill = style.whole_table.as_ref().unwrap().cell.fill.as_ref();

        assert_eq!(fill.map(|fill| fill.fill_type.as_str()), Some("none"));
    }

    #[test]
    fn an_off_or_def_text_flag_reads_as_off_or_unset() {
        let style = only(NO_GRID);
        let text = &style.whole_table.as_ref().unwrap().text;

        assert_eq!(text.bold, Some(false));
        assert_eq!(text.italic, None);
    }

    #[test]
    fn each_region_element_lands_on_its_own_slot() {
        let style = only(REGIONS);

        assert_eq!(fill_of(style.band2_row.as_ref()), Some("B2B2B2"));
        assert_eq!(fill_of(style.last_row.as_ref()), Some("111111"));
        assert_eq!(fill_of(style.first_column.as_ref()), Some("222222"));
        assert_eq!(fill_of(style.last_column.as_ref()), Some("333333"));
        assert!(style.whole_table.is_none());
        assert!(style.band1_row.is_none());
        assert!(style.first_row.is_none());
    }

    #[test]
    fn a_style_list_round_trips_through_the_stored_package_encoding() {
        let list = styles(NO_GRID);
        let json = serde_json::to_vec(&list).unwrap();

        assert_eq!(
            serde_json::from_slice::<TableStyleList>(&json).unwrap(),
            list
        );
    }
}
