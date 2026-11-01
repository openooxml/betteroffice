use std::collections::BTreeMap;
use std::fmt::Debug;

use serde::de::DeserializeOwned;
use serde_json::Value;
use xlsx_model::{
    AnchorCell, AnchorExtent, AnchorPos, AutoFilter, ChartRef, ChartRefKind, ConditionalFormatRule,
};

use super::*;
use crate::formatting::{
    BorderLineStyle, BorderPatch, BorderPreset, HorizontalAlignment, StyleProperty, TextWrapping,
    VerticalAlignment,
};

struct Golden {
    op: Op,
    json: &'static str,
    minimal: &'static str,
    canonical_minimal: &'static str,
}

fn goldens() -> Vec<Golden> {
    let sheet = SheetId(2);
    let at = CellRef {
        row: 1,
        col: 2,
        abs_row: true,
        abs_col: true,
    };
    let range = CellRange::new(CellRef::new(0, 0), CellRef::new(1, 1));
    let cell = CellState {
        value: CellValue::Number { value: 3.5 },
        formula: Some("1+2.5".into()),
        style: Some(7),
    };
    let hyperlink = Hyperlink {
        range,
        external_target: Some("https://example.com".into()),
        location: Some("Sheet1!A1".into()),
        tooltip: Some("Tip".into()),
        display: Some("Link".into()),
    };
    let filter = AutoFilter {
        range,
        raw_xml: Some("<autoFilter/>".into()),
        source_count: 1,
        filter_columns: vec![0],
        has_sort_state: true,
    };
    let format = DifferentialFormat {
        font_color: Some("#112233".into()),
        bold: Some(true),
        italic: Some(true),
        fill: Some("#445566".into()),
    };
    let conditional = ConditionalFormatting {
        sqref: "A1:B2".into(),
        rules: vec![ConditionalFormatRule {
            id: "cf1".into(),
            rule_type: "cellIs".into(),
            operator: Some("greaterThan".into()),
            formulas: vec!["1".into()],
            priority: 1,
            dxf_id: Some(2),
            format: format.clone(),
            raw_xml: Some("<cfRule/>".into()),
        }],
    };
    let anchor = ChartAnchor::Absolute {
        pos: AnchorPos { x: 1, y: 2 },
        extent: AnchorExtent { cx: 3, cy: 4 },
    };
    let to_anchor = ChartAnchor::OneCell {
        from: AnchorCell::default(),
        extent: AnchorExtent { cx: 3, cy: 4 },
    };
    let chart = SheetChart {
        part: "xl/charts/chart1.xml".into(),
        drawing: "xl/drawings/drawing1.xml".into(),
        anchor_index: 0,
        anchor,
        refs: vec![ChartRef {
            kind: ChartRefKind::Values,
            formula: "Sheet1!$A$1".into(),
        }],
        raw_xml: Some("<chart/>".into()),
    };
    let patch = StylePatch {
        bold: Some(true),
        italic: Some(true),
        strikethrough: Some(true),
        font_family: Some("Arial".into()),
        font_size: Some(12.5),
        text_color: Some("#112233".into()),
        fill_color: Some("#445566".into()),
        border: Some(BorderPatch {
            preset: Some(BorderPreset::All),
            style: Some(BorderLineStyle::Solid),
            color: Some("#778899".into()),
        }),
        horizontal_alignment: Some(HorizontalAlignment::Center),
        vertical_alignment: Some(VerticalAlignment::Middle),
        text_wrapping: Some(TextWrapping::Wrap),
        clear: vec![StyleProperty::Bold],
    };
    let defined_name = DefinedName {
        name: "Total".into(),
        formula: "Sheet1!$A$1".into(),
        local_sheet: Some(sheet),
        hidden: true,
    };

    vec![
        Golden {
            op: Op::SetCell {
                sheet,
                at,
                cell: cell.clone(),
            },
            json: concat!(
                r##"{"type":"setCell","sheet":2,"at":{"row":1,"col":2,"abs_row":true,"##,
                r##""abs_col":true},"cell":{"value":{"kind":"number","value":3.5},"##,
                r##""formula":"1+2.5","style":7}}"##,
            ),
            minimal: r##"{"type":"setCell","sheet":2,"at":{"row":1,"col":2},"cell":{}}"##,
            canonical_minimal: concat!(
                r##"{"type":"setCell","sheet":2,"at":{"row":1,"col":2},"##,
                r##""cell":{"value":{"kind":"empty"}}}"##,
            ),
        },
        Golden {
            op: Op::InsertRows {
                sheet,
                at: 3,
                count: 2,
            },
            json: r##"{"type":"insertRows","sheet":2,"at":3,"count":2}"##,
            minimal: r##"{"type":"insertRows","sheet":2,"at":3,"count":2}"##,
            canonical_minimal: r##"{"type":"insertRows","sheet":2,"at":3,"count":2}"##,
        },
        Golden {
            op: Op::DeleteRows {
                sheet,
                at: 3,
                count: 2,
            },
            json: r##"{"type":"deleteRows","sheet":2,"at":3,"count":2}"##,
            minimal: r##"{"type":"deleteRows","sheet":2,"at":3,"count":2}"##,
            canonical_minimal: r##"{"type":"deleteRows","sheet":2,"at":3,"count":2}"##,
        },
        Golden {
            op: Op::InsertCols {
                sheet,
                at: 3,
                count: 2,
            },
            json: r##"{"type":"insertCols","sheet":2,"at":3,"count":2}"##,
            minimal: r##"{"type":"insertCols","sheet":2,"at":3,"count":2}"##,
            canonical_minimal: r##"{"type":"insertCols","sheet":2,"at":3,"count":2}"##,
        },
        Golden {
            op: Op::DeleteCols {
                sheet,
                at: 3,
                count: 2,
            },
            json: r##"{"type":"deleteCols","sheet":2,"at":3,"count":2}"##,
            minimal: r##"{"type":"deleteCols","sheet":2,"at":3,"count":2}"##,
            canonical_minimal: r##"{"type":"deleteCols","sheet":2,"at":3,"count":2}"##,
        },
        Golden {
            op: Op::SetColWidth {
                sheet,
                col: 3,
                width: Some(12.5),
            },
            json: r##"{"type":"setColWidth","sheet":2,"col":3,"width":12.5}"##,
            minimal: r##"{"type":"setColWidth","sheet":2,"col":3}"##,
            canonical_minimal: r##"{"type":"setColWidth","sheet":2,"col":3,"width":null}"##,
        },
        Golden {
            op: Op::SetColVisibility {
                sheet,
                col: 3,
                hidden: true,
            },
            json: r##"{"type":"setColVisibility","sheet":2,"col":3,"hidden":true}"##,
            minimal: r##"{"type":"setColVisibility","sheet":2,"col":3,"hidden":true}"##,
            canonical_minimal: r##"{"type":"setColVisibility","sheet":2,"col":3,"hidden":true}"##,
        },
        Golden {
            op: Op::RestoreColDimension {
                sheet,
                col: 3,
                width: Some(12.5),
                hidden_width: Some(8.5),
            },
            json: concat!(
                r##"{"type":"restoreColDimension","sheet":2,"col":3,"width":12.5,"##,
                r##""hidden_width":8.5}"##,
            ),
            minimal: r##"{"type":"restoreColDimension","sheet":2,"col":3}"##,
            canonical_minimal: concat!(
                r##"{"type":"restoreColDimension","sheet":2,"col":3,"width":null,"##,
                r##""hidden_width":null}"##,
            ),
        },
        Golden {
            op: Op::SetRowHeight {
                sheet,
                row: 3,
                height: Some(12.5),
            },
            json: r##"{"type":"setRowHeight","sheet":2,"row":3,"height":12.5}"##,
            minimal: r##"{"type":"setRowHeight","sheet":2,"row":3}"##,
            canonical_minimal: r##"{"type":"setRowHeight","sheet":2,"row":3,"height":null}"##,
        },
        Golden {
            op: Op::SetRowVisibility {
                sheet,
                row: 3,
                hidden: true,
            },
            json: r##"{"type":"setRowVisibility","sheet":2,"row":3,"hidden":true}"##,
            minimal: r##"{"type":"setRowVisibility","sheet":2,"row":3,"hidden":true}"##,
            canonical_minimal: r##"{"type":"setRowVisibility","sheet":2,"row":3,"hidden":true}"##,
        },
        Golden {
            op: Op::RestoreRowDimension {
                sheet,
                row: 3,
                height: Some(12.5),
                hidden_height: Some(8.5),
            },
            json: concat!(
                r##"{"type":"restoreRowDimension","sheet":2,"row":3,"height":12.5,"##,
                r##""hidden_height":8.5}"##,
            ),
            minimal: r##"{"type":"restoreRowDimension","sheet":2,"row":3}"##,
            canonical_minimal: concat!(
                r##"{"type":"restoreRowDimension","sheet":2,"row":3,"height":null,"##,
                r##""hidden_height":null}"##,
            ),
        },
        Golden {
            op: Op::SetFreezePane {
                sheet,
                pane: Some(FreezePane::new(1, 2, at)),
            },
            json: concat!(
                r##"{"type":"setFreezePane","sheet":2,"pane":{"rows":1,"cols":2,"##,
                r##""top_left":{"row":1,"col":2,"abs_row":true,"abs_col":true}}}"##,
            ),
            minimal: r##"{"type":"setFreezePane","sheet":2}"##,
            canonical_minimal: r##"{"type":"setFreezePane","sheet":2,"pane":null}"##,
        },
        Golden {
            op: Op::SetHyperlinks {
                sheet,
                hyperlinks: vec![hyperlink.clone()],
            },
            json: concat!(
                r##"{"type":"setHyperlinks","sheet":2,"##,
                r##""hyperlinks":[{"range":{"start":{"row":0,"col":0},"end":{"row":1,"##,
                r##""col":1}},"external_target":"https://example.com","##,
                r##""location":"Sheet1!A1","tooltip":"Tip","display":"Link"}]}"##,
            ),
            minimal: r##"{"type":"setHyperlinks","sheet":2,"hyperlinks":[]}"##,
            canonical_minimal: r##"{"type":"setHyperlinks","sheet":2,"hyperlinks":[]}"##,
        },
        Golden {
            op: Op::SetHyperlink { sheet, hyperlink },
            json: concat!(
                r##"{"type":"setHyperlink","sheet":2,"##,
                r##""hyperlink":{"range":{"start":{"row":0,"col":0},"end":{"row":1,"##,
                r##""col":1}},"external_target":"https://example.com","##,
                r##""location":"Sheet1!A1","tooltip":"Tip","display":"Link"}}"##,
            ),
            minimal: concat!(
                r##"{"type":"setHyperlink","sheet":2,"##,
                r##""hyperlink":{"range":{"start":{"row":0,"col":0},"end":{"row":1,"##,
                r##""col":1}}}}"##,
            ),
            canonical_minimal: concat!(
                r##"{"type":"setHyperlink","sheet":2,"##,
                r##""hyperlink":{"range":{"start":{"row":0,"col":0},"end":{"row":1,"##,
                r##""col":1}},"external_target":null,"location":null,"tooltip":null,"##,
                r##""display":null}}"##,
            ),
        },
        Golden {
            op: Op::RemoveHyperlink { sheet, range },
            json: concat!(
                r##"{"type":"removeHyperlink","sheet":2,"range":{"start":{"row":0,"##,
                r##""col":0},"end":{"row":1,"col":1}}}"##,
            ),
            minimal: concat!(
                r##"{"type":"removeHyperlink","sheet":2,"range":{"start":{"row":0,"##,
                r##""col":0},"end":{"row":1,"col":1}}}"##,
            ),
            canonical_minimal: concat!(
                r##"{"type":"removeHyperlink","sheet":2,"range":{"start":{"row":0,"##,
                r##""col":0},"end":{"row":1,"col":1}}}"##,
            ),
        },
        Golden {
            op: Op::SetAutoFilter {
                sheet,
                range: Some(range),
            },
            json: r##"{"type":"setAutoFilter","sheet":2,"range":"A1:B2"}"##,
            minimal: r##"{"type":"setAutoFilter","sheet":2,"range":null}"##,
            canonical_minimal: r##"{"type":"setAutoFilter","sheet":2,"range":null}"##,
        },
        Golden {
            op: Op::RestoreAutoFilter {
                sheet,
                filter: Some(filter),
            },
            json: concat!(
                r##"{"type":"restoreAutoFilter","sheet":2,"##,
                r##""filter":{"range":{"start":{"row":0,"col":0},"end":{"row":1,"##,
                r##""col":1}},"raw_xml":"<autoFilter/>","source_count":1,"##,
                r##""filter_columns":[0],"has_sort_state":true}}"##,
            ),
            minimal: r##"{"type":"restoreAutoFilter","sheet":2}"##,
            canonical_minimal: r##"{"type":"restoreAutoFilter","sheet":2,"filter":null}"##,
        },
        Golden {
            op: Op::AddConditionalFormat {
                sheet,
                range,
                kind: "cellIs".into(),
                operator: Some("greaterThan".into()),
                formulas: vec!["1".into()],
                format: format.clone(),
            },
            json: concat!(
                r##"{"type":"addConditionalFormat","sheet":2,"##,
                r##""range":{"start":{"row":0,"col":0},"end":{"row":1,"col":1}},"##,
                r##""kind":"cellIs","operator":"greaterThan","formulas":["1"],"##,
                r##""format":{"fontColor":"#112233","bold":true,"italic":true,"##,
                r##""fill":"#445566"}}"##,
            ),
            minimal: concat!(
                r##"{"type":"addConditionalFormat","sheet":2,"##,
                r##""range":{"start":{"row":0,"col":0},"end":{"row":1,"col":1}},"##,
                r##""kind":"cellIs","formulas":[],"format":{}}"##,
            ),
            canonical_minimal: concat!(
                r##"{"type":"addConditionalFormat","sheet":2,"##,
                r##""range":{"start":{"row":0,"col":0},"end":{"row":1,"col":1}},"##,
                r##""kind":"cellIs","formulas":[],"format":{}}"##,
            ),
        },
        Golden {
            op: Op::RemoveConditionalFormat {
                sheet,
                id: "cf1".into(),
            },
            json: r##"{"type":"removeConditionalFormat","sheet":2,"id":"cf1"}"##,
            minimal: r##"{"type":"removeConditionalFormat","sheet":2,"id":"cf1"}"##,
            canonical_minimal: r##"{"type":"removeConditionalFormat","sheet":2,"id":"cf1"}"##,
        },
        Golden {
            op: Op::SetConditionalFormats {
                sheet,
                formats: vec![conditional],
            },
            json: concat!(
                r##"{"type":"setConditionalFormats","sheet":2,"##,
                r##""formats":[{"sqref":"A1:B2","rules":[{"id":"cf1","type":"cellIs","##,
                r##""operator":"greaterThan","formulas":["1"],"priority":1,"dxfId":2,"##,
                r##""format":{"fontColor":"#112233","bold":true,"italic":true,"##,
                r##""fill":"#445566"},"rawXml":"<cfRule/>"}]}]}"##,
            ),
            minimal: r##"{"type":"setConditionalFormats","sheet":2,"formats":[]}"##,
            canonical_minimal: r##"{"type":"setConditionalFormats","sheet":2,"formats":[]}"##,
        },
        Golden {
            op: Op::RestoreColStyles {
                sheet,
                styles: vec![ColStyle {
                    first: 0,
                    last: 2,
                    xf: 1,
                }],
            },
            json: concat!(
                r##"{"type":"restoreColStyles","sheet":2,"styles":[{"first":0,"##,
                r##""last":2,"xf":1}]}"##,
            ),
            minimal: r##"{"type":"restoreColStyles","sheet":2,"styles":[]}"##,
            canonical_minimal: r##"{"type":"restoreColStyles","sheet":2,"styles":[]}"##,
        },
        Golden {
            op: Op::RestoreRowStyles {
                sheet,
                styles: BTreeMap::from([(3, 1), (70_000, 2)]),
            },
            json: r##"{"type":"restoreRowStyles","sheet":2,"styles":{"3":1,"70000":2}}"##,
            minimal: r##"{"type":"restoreRowStyles","sheet":2,"styles":{}}"##,
            canonical_minimal: r##"{"type":"restoreRowStyles","sheet":2,"styles":{}}"##,
        },
        Golden {
            op: Op::SetCharts {
                sheet,
                charts: vec![chart],
            },
            json: concat!(
                r##"{"type":"setCharts","sheet":2,"##,
                r##""charts":[{"part":"xl/charts/chart1.xml","##,
                r##""drawing":"xl/drawings/drawing1.xml","anchorIndex":0,"##,
                r##""anchor":{"kind":"absolute","pos":{"x":1,"y":2},"extent":{"cx":3,"##,
                r##""cy":4}},"refs":[{"kind":"values","formula":"Sheet1!$A$1"}],"##,
                r##""rawXml":"<chart/>"}]}"##,
            ),
            minimal: r##"{"type":"setCharts","sheet":2,"charts":[]}"##,
            canonical_minimal: r##"{"type":"setCharts","sheet":2,"charts":[]}"##,
        },
        Golden {
            op: Op::SetChartAnchor {
                sheet,
                frame: "frame".into(),
                part: "chart.xml".into(),
                from: anchor,
                to: to_anchor,
            },
            json: concat!(
                r##"{"type":"setChartAnchor","sheet":2,"frame":"frame","##,
                r##""part":"chart.xml","from":{"kind":"absolute","pos":{"x":1,"y":2},"##,
                r##""extent":{"cx":3,"cy":4}},"to":{"kind":"oneCell","from":{"col":0,"##,
                r##""colOff":0,"row":0,"rowOff":0},"extent":{"cx":3,"cy":4}}}"##,
            ),
            minimal: concat!(
                r##"{"type":"setChartAnchor","sheet":2,"frame":"frame","##,
                r##""part":"chart.xml","from":{"kind":"absolute","pos":{"x":1,"y":2},"##,
                r##""extent":{"cx":3,"cy":4}},"to":{"kind":"oneCell","from":{"col":0,"##,
                r##""colOff":0,"row":0,"rowOff":0},"extent":{"cx":3,"cy":4}}}"##,
            ),
            canonical_minimal: concat!(
                r##"{"type":"setChartAnchor","sheet":2,"frame":"frame","##,
                r##""part":"chart.xml","from":{"kind":"absolute","pos":{"x":1,"y":2},"##,
                r##""extent":{"cx":3,"cy":4}},"to":{"kind":"oneCell","from":{"col":0,"##,
                r##""colOff":0,"row":0,"rowOff":0},"extent":{"cx":3,"cy":4}}}"##,
            ),
        },
        Golden {
            op: Op::MergeCells { sheet, range },
            json: concat!(
                r##"{"type":"mergeCells","sheet":2,"range":{"start":{"row":0,"##,
                r##""col":0},"end":{"row":1,"col":1}}}"##,
            ),
            minimal: concat!(
                r##"{"type":"mergeCells","sheet":2,"range":{"start":{"row":0,"##,
                r##""col":0},"end":{"row":1,"col":1}}}"##,
            ),
            canonical_minimal: concat!(
                r##"{"type":"mergeCells","sheet":2,"range":{"start":{"row":0,"##,
                r##""col":0},"end":{"row":1,"col":1}}}"##,
            ),
        },
        Golden {
            op: Op::UnmergeCells { sheet, range },
            json: concat!(
                r##"{"type":"unmergeCells","sheet":2,"range":{"start":{"row":0,"##,
                r##""col":0},"end":{"row":1,"col":1}}}"##,
            ),
            minimal: concat!(
                r##"{"type":"unmergeCells","sheet":2,"range":{"start":{"row":0,"##,
                r##""col":0},"end":{"row":1,"col":1}}}"##,
            ),
            canonical_minimal: concat!(
                r##"{"type":"unmergeCells","sheet":2,"range":{"start":{"row":0,"##,
                r##""col":0},"end":{"row":1,"col":1}}}"##,
            ),
        },
        Golden {
            op: Op::PatchRangeStyle {
                sheet,
                range,
                patch,
            },
            json: concat!(
                r##"{"type":"patchRangeStyle","sheet":2,"range":{"start":{"row":0,"##,
                r##""col":0},"end":{"row":1,"col":1}},"patch":{"bold":true,"##,
                r##""italic":true,"strikethrough":true,"fontFamily":"Arial","##,
                r##""fontSize":12.5,"textColor":"#112233","fillColor":"#445566","##,
                r##""border":{"preset":"all","style":"solid","color":"#778899"},"##,
                r##""horizontalAlignment":"center","verticalAlignment":"middle","##,
                r##""textWrapping":"wrap","clear":["bold"]}}"##,
            ),
            minimal: concat!(
                r##"{"type":"patchRangeStyle","sheet":2,"range":{"start":{"row":0,"##,
                r##""col":0},"end":{"row":1,"col":1}},"patch":{}}"##,
            ),
            canonical_minimal: concat!(
                r##"{"type":"patchRangeStyle","sheet":2,"range":{"start":{"row":0,"##,
                r##""col":0},"end":{"row":1,"col":1}},"patch":{}}"##,
            ),
        },
        Golden {
            op: Op::SetRangeNumberFormat {
                sheet,
                range,
                format: NumberFormatMutation::Custom {
                    pattern: "0.00".into(),
                },
            },
            json: concat!(
                r##"{"type":"setRangeNumberFormat","sheet":2,"##,
                r##""range":{"start":{"row":0,"col":0},"end":{"row":1,"col":1}},"##,
                r##""format":{"type":"custom","pattern":"0.00"}}"##,
            ),
            minimal: concat!(
                r##"{"type":"setRangeNumberFormat","sheet":2,"##,
                r##""range":{"start":{"row":0,"col":0},"end":{"row":1,"col":1}},"##,
                r##""format":{"type":"custom","pattern":"0.00"}}"##,
            ),
            canonical_minimal: concat!(
                r##"{"type":"setRangeNumberFormat","sheet":2,"##,
                r##""range":{"start":{"row":0,"col":0},"end":{"row":1,"col":1}},"##,
                r##""format":{"type":"custom","pattern":"0.00"}}"##,
            ),
        },
        Golden {
            op: Op::ApplyRangeFormat {
                sheet,
                range,
                format: CapturedFormat {
                    rows: 1,
                    columns: 1,
                    formats: vec![],
                    kept: vec![],
                    source_styles: vec![Some(1)],
                    source_formats: vec![],
                },
            },
            json: concat!(
                r##"{"type":"applyRangeFormat","sheet":2,"range":{"start":{"row":0,"##,
                r##""col":0},"end":{"row":1,"col":1}},"format":{"rows":1,"columns":1,"##,
                r##""formats":[],"sourceStyles":[1]}}"##,
            ),
            minimal: concat!(
                r##"{"type":"applyRangeFormat","sheet":2,"range":{"start":{"row":0,"##,
                r##""col":0},"end":{"row":1,"col":1}},"format":{"rows":1,"columns":1,"##,
                r##""formats":[]}}"##,
            ),
            canonical_minimal: concat!(
                r##"{"type":"applyRangeFormat","sheet":2,"range":{"start":{"row":0,"##,
                r##""col":0},"end":{"row":1,"col":1}},"format":{"rows":1,"columns":1,"##,
                r##""formats":[]}}"##,
            ),
        },
        Golden {
            op: Op::AddSheet {
                index: 1,
                name: "Data".into(),
            },
            json: r##"{"type":"addSheet","index":1,"name":"Data"}"##,
            minimal: r##"{"type":"addSheet","index":1,"name":"Data"}"##,
            canonical_minimal: r##"{"type":"addSheet","index":1,"name":"Data"}"##,
        },
        Golden {
            op: Op::RemoveSheet { index: 1 },
            json: r##"{"type":"removeSheet","index":1}"##,
            minimal: r##"{"type":"removeSheet","index":1}"##,
            canonical_minimal: r##"{"type":"removeSheet","index":1}"##,
        },
        Golden {
            op: Op::MoveSheet { from: 1, to: 2 },
            json: r##"{"type":"moveSheet","from":1,"to":2}"##,
            minimal: r##"{"type":"moveSheet","from":1,"to":2}"##,
            canonical_minimal: r##"{"type":"moveSheet","from":1,"to":2}"##,
        },
        Golden {
            op: Op::RenameSheet {
                sheet,
                name: "Data".into(),
            },
            json: r##"{"type":"renameSheet","sheet":2,"name":"Data"}"##,
            minimal: r##"{"type":"renameSheet","sheet":2,"name":"Data"}"##,
            canonical_minimal: r##"{"type":"renameSheet","sheet":2,"name":"Data"}"##,
        },
        Golden {
            op: Op::SetDefinedName {
                name: "Total".into(),
                refers_to: "Sheet1!$A$1".into(),
                scope: Some(sheet),
                hidden: Some(true),
            },
            json: concat!(
                r##"{"type":"setDefinedName","name":"Total","##,
                r##""refers_to":"Sheet1!$A$1","scope":2,"hidden":true}"##,
            ),
            minimal: r##"{"type":"setDefinedName","name":"Total","refers_to":"Sheet1!$A$1"}"##,
            canonical_minimal: concat!(
                r##"{"type":"setDefinedName","name":"Total","##,
                r##""refers_to":"Sheet1!$A$1","scope":null,"hidden":null}"##,
            ),
        },
        Golden {
            op: Op::DeleteDefinedName {
                name: "Total".into(),
                scope: Some(sheet),
            },
            json: r##"{"type":"deleteDefinedName","name":"Total","scope":2}"##,
            minimal: r##"{"type":"deleteDefinedName","name":"Total"}"##,
            canonical_minimal: r##"{"type":"deleteDefinedName","name":"Total","scope":null}"##,
        },
        Golden {
            op: Op::SetDefinedNames {
                defined_names: vec![defined_name],
            },
            json: concat!(
                r##"{"type":"setDefinedNames","defined_names":[{"name":"Total","##,
                r##""formula":"Sheet1!$A$1","local_sheet":2,"hidden":true}]}"##,
            ),
            minimal: r##"{"type":"setDefinedNames","defined_names":[]}"##,
            canonical_minimal: r##"{"type":"setDefinedNames","defined_names":[]}"##,
        },
        Golden {
            op: Op::RestoreSheet {
                sheet,
                name: "Data".into(),
                formulas: vec![(sheet, at, cell)],
            },
            json: concat!(
                r##"{"type":"restoreSheet","sheet":2,"name":"Data","##,
                r##""formulas":[[2,{"row":1,"col":2,"abs_row":true,"##,
                r##""abs_col":true},{"value":{"kind":"number","value":3.5},"##,
                r##""formula":"1+2.5","style":7}]]}"##,
            ),
            minimal: r##"{"type":"restoreSheet","sheet":2,"name":"Data","formulas":[]}"##,
            canonical_minimal: r##"{"type":"restoreSheet","sheet":2,"name":"Data","formulas":[]}"##,
        },
    ]
}

#[derive(Debug, PartialEq, Serialize, Deserialize)]
struct Ops {
    ops: Vec<Op>,
}

#[derive(Debug, PartialEq, Serialize, Deserialize)]
struct Outer<T> {
    op: T,
    marker: String,
}

#[derive(Debug, PartialEq, Serialize, Deserialize)]
#[serde(untagged)]
enum Untagged<T> {
    Operation(T),
}

#[derive(Debug, PartialEq, Serialize, Deserialize)]
#[serde(tag = "wrapper", rename_all = "camelCase")]
enum Tagged<T> {
    Operation { op: T },
}

#[derive(Debug, PartialEq, Serialize, Deserialize)]
struct Flattened<T> {
    #[serde(flatten)]
    payload: Outer<T>,
}

#[derive(Debug, PartialEq, Serialize, Deserialize)]
struct FlatOp<T> {
    marker: String,
    #[serde(flatten)]
    op: T,
}

mod oracle_context {
    use super::*;

    #[derive(Serialize, Deserialize)]
    pub(super) struct Ops {
        pub(super) ops: Vec<OracleOp>,
    }

    #[derive(Serialize, Deserialize)]
    #[serde(rename_all = "camelCase")]
    pub(super) struct Transaction {
        pub(super) ops: Vec<OracleOp>,
        pub(super) author: Provenance,
        #[serde(default)]
        pub(super) proposed: bool,
    }
}

use oracle_context::{Ops as OracleOps, Transaction as OracleTransaction};

impl From<OracleOps> for Ops {
    fn from(oracle: OracleOps) -> Self {
        Self {
            ops: oracle.ops.into_iter().map(Op::from).collect(),
        }
    }
}

impl From<OracleTransaction> for Transaction {
    fn from(oracle: OracleTransaction) -> Self {
        Self {
            ops: oracle.ops.into_iter().map(Op::from).collect(),
            author: oracle.author,
            proposed: oracle.proposed,
        }
    }
}

fn entry_points<T: DeserializeOwned>(input: &str) -> Vec<(&'static str, serde_json::Result<T>)> {
    let mut decoded = vec![
        ("from_str", serde_json::from_str(input)),
        ("from_slice", serde_json::from_slice(input.as_bytes())),
    ];
    if let Ok(value) = serde_json::from_str::<Value>(input) {
        decoded.push(("from_value", serde_json::from_value(value)));
    }
    decoded
}

fn assert_canonical<T, O>(op: &T, convert: &impl Fn(O) -> T)
where
    T: Debug + PartialEq + Serialize + DeserializeOwned,
    O: DeserializeOwned,
{
    let canonical = serde_json::to_string(op).unwrap();
    for (entry, decoded) in entry_points::<T>(&canonical) {
        assert_eq!(&decoded.unwrap(), op, "new/{entry}: {canonical}");
    }
    for (entry, decoded) in entry_points::<O>(&canonical) {
        assert_eq!(
            &convert(decoded.unwrap()),
            op,
            "oracle/{entry}: {canonical}"
        );
    }
}

fn assert_success<T, O>(input: &str, convert: impl Fn(O) -> T)
where
    T: Debug + PartialEq + Serialize + DeserializeOwned,
    O: Serialize + DeserializeOwned,
{
    for ((entry, direct), (_, oracle)) in entry_points::<T>(input)
        .into_iter()
        .zip(entry_points::<O>(input))
    {
        let direct = direct.unwrap_or_else(|error| panic!("new/{entry}: {input}: {error}"));
        let oracle = oracle.unwrap_or_else(|error| panic!("oracle/{entry}: {input}: {error}"));
        assert_eq!(
            serde_json::to_string(&direct).unwrap(),
            serde_json::to_string(&oracle).unwrap(),
            "serialization/{entry}: {input}"
        );
        assert_eq!(direct, convert(oracle), "value/{entry}: {input}");
        assert_canonical::<T, O>(&direct, &convert);
    }
}

fn assert_value_result<T, O>(input: &str, convert: impl Fn(O) -> T)
where
    T: Debug + PartialEq + Serialize + DeserializeOwned,
    O: Serialize + DeserializeOwned,
{
    let Ok(value) = serde_json::from_str::<Value>(input) else {
        return;
    };
    let direct = serde_json::from_value::<T>(value.clone());
    match serde_json::from_value::<O>(value) {
        Ok(oracle) => {
            let direct = direct.unwrap();
            assert_eq!(
                serde_json::to_string(&direct).unwrap(),
                serde_json::to_string(&oracle).unwrap(),
                "from_value serialization: {input}"
            );
            assert_eq!(direct, convert(oracle), "from_value: {input}");
            assert_canonical::<T, O>(&direct, &convert);
        }
        Err(_) => assert!(direct.is_err(), "new/from_value accepted {input}"),
    }
}

fn assert_both_err<T, O>(input: &str, convert: impl Fn(O) -> T)
where
    T: Debug + PartialEq + Serialize + DeserializeOwned,
    O: Serialize + DeserializeOwned,
{
    assert!(
        serde_json::from_str::<T>(input).is_err(),
        "new/from_str: {input}"
    );
    assert!(
        serde_json::from_str::<O>(input).is_err(),
        "oracle/from_str: {input}"
    );
    assert!(
        serde_json::from_slice::<T>(input.as_bytes()).is_err(),
        "new/from_slice: {input}"
    );
    assert!(
        serde_json::from_slice::<O>(input.as_bytes()).is_err(),
        "oracle/from_slice: {input}"
    );
    assert_value_result::<T, O>(input, convert);
}

fn assert_well_formed(input: &str) {
    assert_success::<Op, OracleOp>(input, Op::from);
    assert_success::<Vec<Op>, Vec<OracleOp>>(&format!("[{input}]"), |ops| {
        ops.into_iter().map(Op::from).collect()
    });
    assert_success::<Ops, OracleOps>(&format!("{{\"ops\":[{input}]}}"), Ops::from);
    assert_success::<Outer<Op>, Outer<OracleOp>>(
        &format!("{{\"marker\":\"outer\",\"op\":{input}}}"),
        |outer| Outer {
            op: outer.op.into(),
            marker: outer.marker,
        },
    );
    for author in [
        r##"{"kind":"user"}"##,
        r##"{"kind":"agent","id":"bot"}"##,
        r##"{"kind":"remote","actor":"peer"}"##,
    ] {
        assert_success::<Transaction, OracleTransaction>(
            &format!("{{\"ops\":[{input}],\"author\":{author}}}"),
            Transaction::from,
        );
    }
}

fn assert_replayed(input: &str) {
    assert_success::<Untagged<Op>, Untagged<OracleOp>>(input, |parent| match parent {
        Untagged::Operation(op) => Untagged::Operation(op.into()),
    });
    assert_success::<Tagged<Op>, Tagged<OracleOp>>(
        &format!("{{\"wrapper\":\"operation\",\"op\":{input}}}"),
        |parent| match parent {
            Tagged::Operation { op } => Tagged::Operation { op: op.into() },
        },
    );
    assert_success::<Flattened<Op>, Flattened<OracleOp>>(
        &format!("{{\"marker\":\"outer\",\"op\":{input}}}"),
        |parent| Flattened {
            payload: Outer {
                op: parent.payload.op.into(),
                marker: parent.payload.marker,
            },
        },
    );
    if input.starts_with('{') {
        assert_success::<FlatOp<Op>, FlatOp<OracleOp>>(
            &format!("{{\"marker\":\"outer\",{}}}", &input[1..input.len() - 1]),
            |parent| FlatOp {
                op: parent.op.into(),
                marker: parent.marker,
            },
        );
    }
}

fn assert_malformed(input: &str) {
    assert_both_err::<Op, OracleOp>(input, Op::from);
    assert_both_err::<Vec<Op>, Vec<OracleOp>>(&format!("[{input}]"), |ops| {
        ops.into_iter().map(Op::from).collect()
    });
    assert_both_err::<Ops, OracleOps>(&format!("{{\"ops\":[{input}]}}"), Ops::from);
    assert_both_err::<Outer<Op>, Outer<OracleOp>>(
        &format!("{{\"op\":{input},\"marker\":\"outer\"}}"),
        |outer| Outer {
            op: outer.op.into(),
            marker: outer.marker,
        },
    );
    assert_both_err::<Transaction, OracleTransaction>(
        &format!("{{\"ops\":[{input}],\"author\":{{\"kind\":\"user\"}}}}"),
        Transaction::from,
    );
}

fn tag_orders(input: &str) -> Vec<String> {
    let value = serde_json::from_str::<Value>(input).unwrap();
    let mut fields = value.as_object().unwrap().clone();
    let tag = fields.remove("type").unwrap();
    let fields: Vec<_> = fields
        .into_iter()
        .map(|(key, value)| format!("{}:{value}", serde_json::to_string(&key).unwrap()))
        .collect();
    let mut inputs = Vec::new();
    for position in [0, fields.len() / 2, fields.len()] {
        for key in [
            r##""type""##,
            r##""t\u0079pe""##,
            r##""\u0074\u0079\u0070\u0065""##,
        ] {
            let mut ordered = fields.clone();
            ordered.insert(position, format!("{key}:{tag}"));
            inputs.push(format!("{{{}}}", ordered.join(",")));
        }
    }
    inputs
}

fn field_mutations(input: &str) -> Vec<String> {
    let value = serde_json::from_str::<Value>(input).unwrap();
    let fields = value.as_object().unwrap();
    let mut inputs = Vec::new();
    for (key, value) in fields {
        if key == "type" {
            continue;
        }
        let mut missing = fields.clone();
        missing.remove(key);
        inputs.extend(tag_orders(&serde_json::to_string(&missing).unwrap()));

        let invalid = if value.is_boolean() {
            Value::String("wrong".into())
        } else {
            Value::Bool(false)
        };
        for replacement in [invalid, Value::Null] {
            let mut wrong = fields.clone();
            wrong.insert(key.clone(), replacement);
            inputs.extend(tag_orders(&serde_json::to_string(&wrong).unwrap()));
        }
        inputs.push(format!(
            "{},{}:{value}}}",
            &input[..input.len() - 1],
            serde_json::to_string(key).unwrap()
        ));
    }
    inputs
}

mod well_formed {
    use super::*;

    #[test]
    fn every_variant_has_full_and_minimal_serialization_goldens() {
        let cases = goldens();
        assert_eq!(cases.len(), 37);
        for Golden {
            op,
            json,
            minimal,
            canonical_minimal,
        } in cases
        {
            assert_eq!(serde_json::to_string(&op).unwrap(), json);
            let oracle = serde_json::from_str::<OracleOp>(json).unwrap();
            assert_eq!(serde_json::to_string(&oracle).unwrap(), json);
            assert_eq!(op, Op::from(oracle));
            let minimal_op = serde_json::from_str::<OracleOp>(minimal).unwrap();
            assert_eq!(
                serde_json::to_string(&minimal_op).unwrap(),
                canonical_minimal
            );
            assert_eq!(
                serde_json::to_string(&Op::from(minimal_op)).unwrap(),
                canonical_minimal
            );
            for input in [json, minimal, canonical_minimal] {
                assert_well_formed(input);
                for ordered in tag_orders(input) {
                    assert_well_formed(&ordered);
                    assert_replayed(&ordered);
                }
            }
        }
    }

    #[test]
    fn every_variant_accepts_positional_fields() {
        for golden in goldens() {
            for input in [golden.json, golden.minimal, golden.canonical_minimal] {
                let value = serde_json::from_str::<Value>(input).unwrap();
                let tag = serde_json::from_value::<Tag>(value["type"].clone()).unwrap();
                let mut fields = vec![value["type"].clone()];
                fields.extend(
                    tag.fields()
                        .iter()
                        .map(|name| value.get(*name).cloned().unwrap_or(Value::Null)),
                );
                let input = serde_json::to_string(&fields).unwrap();
                assert_well_formed(&input);
                assert_replayed(&input);
            }
        }
    }

    #[test]
    fn optional_fields_and_ignored_fields_preserve_values() {
        for golden in goldens() {
            for input in field_mutations(golden.json) {
                if serde_json::from_str::<OracleOp>(&input).is_ok() {
                    assert_well_formed(&input);
                }
            }
            let unknown = format!(
                "{},\"unknown\":{{\"nested\":[null,1,{{\"type\":\"ignored\"}}]}}}}",
                &golden.json[..golden.json.len() - 1]
            );
            for input in tag_orders(&unknown) {
                assert_well_formed(&input);
            }
        }
    }

    #[test]
    fn crafted_inputs_match_in_all_entry_points_and_contexts() {
        for input in INPUTS {
            assert_well_formed(input);
        }
        let maximum_index = format!("{{\"type\":\"removeSheet\",\"index\":{}}}", usize::MAX);
        assert_well_formed(&maximum_index);
    }

    #[test]
    fn transaction_fields_preserve_values() {
        for input in [
            r##"{"ops":[],"author":{"kind":"user"}}"##,
            r##"{"ops":[],"author":{"kind":"agent","id":"bot"},"proposed":true}"##,
            r##"{"ops":[],"author":{"kind":"remote","actor":"peer"},"proposed":false}"##,
            r##"[[["removeSheet",0]],{"kind":"user"},true]"##,
        ] {
            assert_success::<Transaction, OracleTransaction>(input, Transaction::from);
        }
    }

    #[test]
    fn from_value_preserves_float_payloads() {
        for bits in [
            0,
            1,
            0x000f_ffff_ffff_ffff,
            0x0010_0000_0000_0000,
            0x3fb9_9999_9999_9999,
            0x3fb9_9999_9999_999a,
            0x3fd3_3333_3333_3333,
            0x3fef_ffff_ffff_ffff,
            0x3ff0_0000_0000_0000,
            0x3ff0_0000_0000_0001,
            0x4340_0000_0000_0001,
            0x7fef_ffff_ffff_ffff,
            0x8000_0000_0000_0000,
            0xbfb9_9999_9999_999a,
        ] {
            let value = serde_json::json!({
                "type": "setColWidth",
                "sheet": 0,
                "col": 0,
                "width": f64::from_bits(bits),
            });
            let direct = serde_json::from_value::<Op>(value.clone()).unwrap();
            let oracle = serde_json::from_value::<OracleOp>(value.clone()).unwrap();
            assert_eq!(direct, Op::from(oracle));
            let Op::SetColWidth { width, .. } = direct else {
                panic!("expected setColWidth");
            };
            assert_eq!(width.unwrap().to_bits(), bits);
            assert_well_formed(&value.to_string());
        }
    }

    const INPUTS: &[&str] = &[
        r##"["removeSheet",0]"##,
        r##"["addSheet",0,"Data"]"##,
        r##"["insertRows",0,1,2]"##,
        r##"["setColWidth",0,1,null]"##,
        r##"["setColWidth",0,1,12.5]"##,
        r##"["setDefinedName","N","A1",null,null]"##,
        r##"["setAutoFilter",0,null]"##,
        r##"["setAutoFilter",0,"A1:B2"]"##,
        r##"["restoreRowStyles",0,{"3":1,"70000":2}]"##,
        r##"{"type":"removeSheet","index":0}"##,
        r##"{"index":0,"type":"removeSheet"}"##,
        r##"{"index":0,"type":"removeSheet","x":1}"##,
        r##"{"x":1,"type":"removeSheet","index":0}"##,
        r##"{"t\u0079pe":"removeSheet","index":0}"##,
        r##"{"\u0074\u0079\u0070\u0065":"removeSheet","index":0}"##,
        r##"{"type":"removeSheet","in\u0064ex":0}"##,
        r##"{"type":"removeSheet","index":0,"x":1,"x":2}"##,
        r##"{"x":1,"x":2,"type":"removeSheet","index":0}"##,
        r##"{"type":"setAutoFilter","sheet":0,"range":null}"##,
        r##"{"type":"setAutoFilter","sheet":0,"range":"A1:B2"}"##,
        r##"{"type":"setDefinedName","name":"N","refers_to":"A1"}"##,
        r##"{"type":"setDefinedName","name":"N","refers_to":"A1","scope":null,"hidden":null}"##,
        r##"{"type":"setDefinedName","name":"N","refers_to":"A1","hidden":false}"##,
        r##"{"type":"deleteDefinedName","name":"N"}"##,
        r##"{"type":"restoreRowStyles","sheet":0,"styles":{}}"##,
        r##"{"type":"restoreRowStyles","sheet":0,"styles":{"03":1,"3":2}}"##,
        r##"{"type":"insertRows","sheet":0,"at":0,"count":4294967295}"##,
        r##"{"type":"setColWidth","sheet":0,"col":0,"width":-0}"##,
        r##"{"type":"setColWidth","sheet":0,"col":0,"width":0}"##,
        r##"{"type":"setColWidth","sheet":0,"col":0,"width":1.7976931348623157e308}"##,
        r##"{"type":"setColWidth","sheet":0,"col":0,"width":5e-324}"##,
        r##"{"type":"setColWidth","sheet":0,"col":0,"width":1e-400}"##,
        r##"{"type":"setColWidth","sheet":0,"col":0,"width":9007199254740993}"##,
        r##"{"type":"setColWidth","sheet":0,"col":0,"width":18446744073709551616}"##,
        concat!(
            r##"{"type":"setCell","sheet":0,"at":{"row":0,"col":0},"##,
            r##""cell":{"value":{"kind":"empty"}}}"##,
        ),
        concat!(
            r##"{"type":"setCell","sheet":0,"at":{"row":0,"col":0},"##,
            r##""cell":{"value":{"kind":"text","value":"a\nb"}}}"##,
        ),
        concat!(
            r##"{"type":"setCell","sheet":0,"at":{"row":0,"col":0},"##,
            r##""cell":{"value":{"kind":"number","value":3.5}}}"##,
        ),
        concat!(
            r##"{"type":"setCell","sheet":0,"at":{"row":0,"col":0},"##,
            r##""cell":{"value":{"kind":"bool","value":true}}}"##,
        ),
        concat!(
            r##"{"type":"setCell","sheet":0,"at":{"row":0,"col":0},"##,
            r##""cell":{"value":{"kind":"error","value":"#DIV/0!"}}}"##,
        ),
        r##"{"type":"setCell","sheet":0,"at":{"row":0,"col":0},"cell":{"value":["number",3.5]}}"##,
        concat!(
            r##"{"type":"patchRangeStyle","sheet":0,"range":{"start":{"row":0,"col":0},"##,
            r##""end":{"row":1,"col":1}},"patch":{}}"##,
        ),
        concat!(
            r##"{"type":"patchRangeStyle","sheet":0,"range":{"start":{"row":0,"col":0},"##,
            r##""end":{"row":1,"col":1}},"patch":{"bold":null}}"##,
        ),
        concat!(
            r##"{"type":"patchRangeStyle","sheet":0,"range":{"start":{"row":0,"col":0},"##,
            r##""end":{"row":1,"col":1}},"patch":{"textWrapping":"wrap","##,
            r##""fillColor":"#112233"}}"##,
        ),
        concat!(
            r##"{"type":"setRangeNumberFormat","sheet":0,"range":{"start":{"row":0,"col":0},"##,
            r##""end":{"row":1,"col":1}},"format":{"type":"automatic"}}"##,
        ),
        concat!(
            r##"{"type":"setRangeNumberFormat","sheet":0,"range":{"start":{"row":0,"col":0},"##,
            r##""end":{"row":1,"col":1}},"format":{"type":"custom","pattern":"0.00"}}"##,
        ),
    ];
}

fn assert_rejected_as<T, O>(input: &str, expected: &T, convert: impl Fn(O) -> T)
where
    T: Debug + PartialEq + Serialize + DeserializeOwned,
    O: DeserializeOwned,
{
    for ((entry, direct), (_, oracle)) in entry_points::<T>(input)
        .into_iter()
        .zip(entry_points::<O>(input))
    {
        assert!(direct.is_err(), "new/{entry}: {input}");
        assert_eq!(
            &convert(oracle.unwrap()),
            expected,
            "oracle/{entry}: {input}"
        );
    }
    assert_canonical::<T, O>(expected, &convert);
}

fn assert_pinned_rejection(input: &str, expected: &Op) {
    assert_rejected_as::<Op, OracleOp>(input, expected, Op::from);
    assert_rejected_as::<Vec<Op>, Vec<OracleOp>>(
        &format!("[{input}]"),
        &vec![expected.clone()],
        |ops| ops.into_iter().map(Op::from).collect(),
    );
    assert_rejected_as::<Ops, OracleOps>(
        &format!("{{\"ops\":[{input}]}}"),
        &Ops {
            ops: vec![expected.clone()],
        },
        Ops::from,
    );
    assert_rejected_as::<Outer<Op>, Outer<OracleOp>>(
        &format!("{{\"op\":{input},\"marker\":\"outer\"}}"),
        &Outer {
            op: expected.clone(),
            marker: "outer".into(),
        },
        |outer| Outer {
            op: outer.op.into(),
            marker: outer.marker,
        },
    );
    assert_rejected_as::<Transaction, OracleTransaction>(
        &format!("{{\"ops\":[{input}],\"author\":{{\"kind\":\"user\"}}}}"),
        &Transaction::new(vec![expected.clone()], Provenance::User),
        Transaction::from,
    );
}

fn assert_pinned_acceptance<T, O>(input: &str, expected: &T, convert: impl Fn(O) -> T)
where
    T: Debug + PartialEq + Serialize + DeserializeOwned,
    O: DeserializeOwned,
{
    for (entry, result) in entry_points::<T>(input) {
        assert_eq!(&result.unwrap(), expected, "new/{entry}: {input}");
    }
    assert!(
        serde_json::from_str::<O>(input).is_err(),
        "oracle/from_str: {input}"
    );
    assert!(
        serde_json::from_slice::<O>(input.as_bytes()).is_err(),
        "oracle/from_slice: {input}"
    );
    assert_canonical::<T, O>(expected, &convert);
}

mod pinned {
    use super::*;

    fn cell(value: CellValue) -> Op {
        Op::SetCell {
            sheet: SheetId(0),
            at: CellRef::new(0, 0),
            cell: CellState {
                value,
                ..CellState::default()
            },
        }
    }

    #[test]
    fn numeric_nested_cell_value_tags_are_rejected() {
        for (value, expected) in [
            (r##"{"kind":0}"##, CellValue::Empty),
            (
                r##"{"kind":1,"value":3.5}"##,
                CellValue::Number { value: 3.5 },
            ),
        ] {
            let input = format!(
                "{{\"type\":\"setCell\",\"sheet\":0,\"at\":{{\"row\":0,\"col\":0}},\
                 \"cell\":{{\"value\":{value}}}}}"
            );
            for input in tag_orders(&input) {
                assert_pinned_rejection(&input, &cell(expected.clone()));
            }
        }
    }

    #[test]
    fn numeric_positional_nested_cell_values_are_rejected() {
        let input = concat!(
            r##"{"type":"setCell","sheet":0,"at":{"row":0,"col":0},"##,
            r##""cell":{"value":[1,3.5]}}"##,
        );
        for input in tag_orders(input) {
            assert_pinned_rejection(&input, &cell(CellValue::Number { value: 3.5 }));
        }
        assert_pinned_rejection(
            r##"["setCell",0,{"row":0,"col":0},{"value":[1,3.5]}]"##,
            &cell(CellValue::Number { value: 3.5 }),
        );
    }

    #[test]
    fn numeric_nested_number_format_tags_are_rejected() {
        for (format, expected) in [
            (r##"{"type":0}"##, NumberFormatMutation::Automatic),
            (
                r##"{"type":8,"pattern":"0.00"}"##,
                NumberFormatMutation::Custom {
                    pattern: "0.00".into(),
                },
            ),
        ] {
            let input = format!(
                "{{\"type\":\"setRangeNumberFormat\",\"sheet\":0,\
                 \"range\":{{\"start\":{{\"row\":0,\"col\":0}},\
                 \"end\":{{\"row\":1,\"col\":1}}}},\"format\":{format}}}"
            );
            let expected = Op::SetRangeNumberFormat {
                sheet: SheetId(0),
                range: CellRange::new(CellRef::new(0, 0), CellRef::new(1, 1)),
                format: expected,
            };
            for input in tag_orders(&input) {
                assert_pinned_rejection(&input, &expected);
            }
        }
    }

    #[test]
    fn numeric_nested_chart_anchor_tags_are_rejected() {
        let input = concat!(
            r##"{"type":"setChartAnchor","sheet":0,"frame":"frame","part":"chart.xml","##,
            r##""from":{"kind":2,"pos":{"x":1,"y":2},"extent":{"cx":3,"cy":4}},"##,
            r##""to":{"kind":"absolute","pos":{"x":1,"y":2},"extent":{"cx":3,"cy":4}}}"##,
        );
        let anchor = ChartAnchor::Absolute {
            pos: AnchorPos { x: 1, y: 2 },
            extent: AnchorExtent { cx: 3, cy: 4 },
        };
        let expected = Op::SetChartAnchor {
            sheet: SheetId(0),
            frame: "frame".into(),
            part: "chart.xml".into(),
            from: anchor,
            to: anchor,
        };
        for input in tag_orders(input) {
            assert_pinned_rejection(&input, &expected);
        }
    }

    fn assert_ignored_op(input: &str, expected: &Op) {
        assert_pinned_acceptance::<Op, OracleOp>(input, expected, Op::from);
        assert_pinned_acceptance::<Vec<Op>, Vec<OracleOp>>(
            &format!("[{input}]"),
            &vec![expected.clone()],
            |ops| ops.into_iter().map(Op::from).collect(),
        );
        assert_pinned_acceptance::<Ops, OracleOps>(
            &format!("{{\"ops\":[{input}]}}"),
            &Ops {
                ops: vec![expected.clone()],
            },
            Ops::from,
        );
        assert_pinned_acceptance::<Outer<Op>, Outer<OracleOp>>(
            &format!("{{\"op\":{input},\"marker\":\"outer\"}}"),
            &Outer {
                op: expected.clone(),
                marker: "outer".into(),
            },
            |outer| Outer {
                op: outer.op.into(),
                marker: outer.marker,
            },
        );
        assert_pinned_acceptance::<Transaction, OracleTransaction>(
            &format!("{{\"ops\":[{input}],\"author\":{{\"kind\":\"user\"}}}}"),
            &Transaction::new(vec![expected.clone()], Provenance::User),
            Transaction::from,
        );
    }

    fn assert_ignored(junk: &str) {
        let direct = format!("{{\"type\":\"removeSheet\",\"index\":0,\"unknown\":{junk}}}");
        assert_ignored_op(&direct, &Op::RemoveSheet { index: 0 });
        for fallback in [
            format!("{{\"index\":0,\"type\":\"removeSheet\",\"unknown\":{junk}}}"),
            format!("{{\"unknown\":{junk},\"index\":0,\"type\":\"removeSheet\"}}"),
        ] {
            assert_malformed(&fallback);
        }
        for input in [
            format!(
                "{{\"type\":\"setCell\",\"sheet\":0,\"at\":{{\"row\":0,\"col\":0}},\
                 \"cell\":{{\"unknown\":{junk}}}}}"
            ),
            format!(
                "{{\"type\":\"setCell\",\"sheet\":0,\
                 \"at\":{{\"row\":0,\"col\":0,\"unknown\":{junk}}},\"cell\":{{}}}}"
            ),
        ] {
            assert_ignored_op(&input, &cell(CellValue::Empty));
            let fields = &input[r##"{"type":"setCell","##.len()..input.len() - 1];
            assert_malformed(&format!("{{{fields},\"type\":\"setCell\"}}"));
        }
    }

    #[test]
    fn overflowing_ignored_numbers_are_accepted_directly_and_rejected_in_fallback() {
        assert_ignored("1e309");
    }

    #[test]
    fn ignored_surrogates_are_accepted_directly_and_rejected_in_fallback() {
        assert_ignored(r##""\ud800""##);
        assert_ignored(r##""\udc00""##);
    }

    #[test]
    fn invalid_utf8_in_a_directly_ignored_field_is_accepted() {
        let input = b"{\"type\":\"removeSheet\",\"index\":0,\"unknown\":\"\xff\"}";
        let expected = Op::RemoveSheet { index: 0 };
        assert_eq!(serde_json::from_slice::<Op>(input).unwrap(), expected);
        assert!(serde_json::from_slice::<OracleOp>(input).is_err());
        assert_canonical::<Op, OracleOp>(&expected, &Op::from);
        let fallback = b"{\"index\":0,\"type\":\"removeSheet\",\"unknown\":\"\xff\"}";
        assert!(serde_json::from_slice::<Op>(fallback).is_err());
        assert!(serde_json::from_slice::<OracleOp>(fallback).is_err());
    }

    #[test]
    fn ignored_excessive_nesting_is_accepted_directly_and_rejected_in_fallback() {
        for depth in [128, 256, 512] {
            let nested = format!("{}0{}", "[".repeat(depth), "]".repeat(depth));
            assert_ignored(&nested);
        }
    }
}

mod malformed {
    use super::*;

    #[test]
    fn malformed_inputs_fail_in_every_context() {
        for input in INPUTS {
            if input.trim().is_empty() {
                assert_both_err::<Op, OracleOp>(input, Op::from);
            } else {
                assert_malformed(input);
            }
        }
        if usize::BITS < 64 {
            assert_malformed(r##"{"type":"removeSheet","index":18446744073709551615}"##);
        }
    }

    #[test]
    fn required_wrong_and_duplicate_fields_fail() {
        for golden in goldens() {
            for input in field_mutations(golden.json) {
                if serde_json::from_str::<OracleOp>(&input).is_err() {
                    assert_malformed(&input);
                }
            }
        }
    }

    #[test]
    fn fallback_preserves_nested_duplicate_fields() {
        for input in [
            concat!(
                r##"{"at":{"row":0,"row":1,"col":0},"type":"setCell","##,
                r##""sheet":0,"cell":{}}"##,
            ),
            concat!(
                r##"{"sheet":0,"type":"setCell","at":{"row":0,"col":0},"##,
                r##""cell":{"value":{"kind":"empty","kind":"number","value":3.5}}}"##,
            ),
        ] {
            assert_malformed(input);
        }
    }

    #[test]
    fn positional_fields_require_every_position_and_reject_extras() {
        for golden in goldens() {
            let value = serde_json::from_str::<Value>(golden.json).unwrap();
            let tag = serde_json::from_value::<Tag>(value["type"].clone()).unwrap();
            let mut fields = vec![value["type"].clone()];
            fields.extend(
                tag.fields()
                    .iter()
                    .map(|name| value.get(*name).cloned().unwrap_or(Value::Null)),
            );
            for length in 1..fields.len() {
                assert_malformed(&serde_json::to_string(&fields[..length]).unwrap());
            }
            fields.push(Value::Null);
            assert_malformed(&serde_json::to_string(&fields).unwrap());
        }
    }

    #[test]
    fn malformed_transaction_fields_fail() {
        for input in [
            r##"{"ops":[],"author":{"kind":"remote","actor":"peer"},"proposed":null}"##,
            r##"{"ops":[],"author":{"kind":"user"},"proposed":"yes"}"##,
            r##"{"ops":[],"author":{"kind":"user"},"ops":[]}"##,
            r##"{"ops":[],"author":{"kind":"unknown"}}"##,
            r##"{"ops":[{"type":"removeSheet","index":0}]}"##,
            r##"{"author":{"kind":"user"}}"##,
            r##"{"ops":null,"author":{"kind":"user"}}"##,
        ] {
            assert_both_err::<Transaction, OracleTransaction>(input, Transaction::from);
        }
    }

    #[test]
    fn invalid_bytes_fail() {
        for input in [
            b"{\"type\":\"removeSheet\",\"index\":0,\"\xff\":null}".as_slice(),
            b"{\"type\":\"renameSheet\",\"sheet\":0,\"name\":\"\xff\"}".as_slice(),
            b"{\"type\":\"\xff\",\"index\":0}".as_slice(),
            b"{\"type\":\"removeSheet\",\"index\":0}\xff".as_slice(),
        ] {
            assert!(
                serde_json::from_slice::<Op>(input).is_err(),
                "new: {input:?}"
            );
            assert!(
                serde_json::from_slice::<OracleOp>(input).is_err(),
                "oracle: {input:?}"
            );
        }
    }

    fn assert_boundary<T: DeserializeOwned>(make_input: impl Fn(usize) -> String, maximum: usize) {
        for depth in maximum - 1..=maximum + 1 {
            let input = make_input(depth);
            assert_eq!(
                serde_json::from_str::<T>(&input).is_ok(),
                depth <= maximum,
                "from_str/max={maximum}/arrays={depth}: {input}"
            );
            assert_eq!(
                serde_json::from_slice::<T>(input.as_bytes()).is_ok(),
                depth <= maximum,
                "from_slice/max={maximum}/arrays={depth}: {input}"
            );
        }
    }

    #[test]
    fn current_recursion_maxima_are_pinned() {
        let ignored = |depth| {
            let nested = format!("{}0{}", "[".repeat(depth), "]".repeat(depth));
            format!("{{\"type\":\"removeSheet\",\"index\":0,\"unknown\":{nested}}}")
        };
        assert_boundary::<OracleOp>(ignored, 126);
        assert_boundary::<Vec<OracleOp>>(|depth| format!("[{}]", ignored(depth)), 125);
        assert_boundary::<OracleOps>(|depth| format!("{{\"ops\":[{}]}}", ignored(depth)), 124);
        assert_boundary::<OracleTransaction>(
            |depth| {
                format!(
                    "{{\"ops\":[{}],\"author\":{{\"kind\":\"user\"}}}}",
                    ignored(depth)
                )
            },
            124,
        );
        assert_boundary::<Outer<OracleOp>>(
            |depth| format!("{{\"op\":{},\"marker\":\"outer\"}}", ignored(depth)),
            125,
        );

        let fallback = |depth| {
            let nested = format!("{}0{}", "[".repeat(depth), "]".repeat(depth));
            format!("{{\"index\":0,\"type\":\"removeSheet\",\"unknown\":{nested}}}")
        };
        assert_boundary::<Op>(fallback, 126);
        assert_boundary::<Vec<Op>>(|depth| format!("[{}]", fallback(depth)), 125);
        assert_boundary::<Ops>(|depth| format!("{{\"ops\":[{}]}}", fallback(depth)), 124);
        assert_boundary::<Transaction>(
            |depth| {
                format!(
                    "{{\"ops\":[{}],\"author\":{{\"kind\":\"user\"}}}}",
                    fallback(depth)
                )
            },
            124,
        );
        assert_boundary::<Outer<Op>>(
            |depth| format!("{{\"op\":{},\"marker\":\"outer\"}}", fallback(depth)),
            125,
        );

        let cell = |depth, tag_first| {
            let nested = format!("{}0{}", "[".repeat(depth), "]".repeat(depth));
            let fields = format!(
                "\"sheet\":0,\"at\":{{\"row\":0,\"col\":0}},\
                 \"cell\":{{\"value\":{{\"kind\":\"empty\",\"unknown\":{nested}}}}}"
            );
            if tag_first {
                format!("{{\"type\":\"setCell\",{fields}}}")
            } else {
                format!("{{{fields},\"type\":\"setCell\"}}")
            }
        };
        for tag_first in [true, false] {
            assert_boundary::<OracleOp>(|depth| cell(depth, tag_first), 124);
            assert_boundary::<Vec<OracleOp>>(|depth| format!("[{}]", cell(depth, tag_first)), 123);
            assert_boundary::<OracleOps>(
                |depth| format!("{{\"ops\":[{}]}}", cell(depth, tag_first)),
                122,
            );
            assert_boundary::<OracleTransaction>(
                |depth| {
                    format!(
                        "{{\"ops\":[{}],\"author\":{{\"kind\":\"user\"}}}}",
                        cell(depth, tag_first)
                    )
                },
                122,
            );
            assert_boundary::<Outer<OracleOp>>(
                |depth| format!("{{\"op\":{},\"marker\":\"outer\"}}", cell(depth, tag_first)),
                123,
            );
            assert_boundary::<Op>(|depth| cell(depth, tag_first), 124);
            assert_boundary::<Vec<Op>>(|depth| format!("[{}]", cell(depth, tag_first)), 123);
            assert_boundary::<Ops>(
                |depth| format!("{{\"ops\":[{}]}}", cell(depth, tag_first)),
                122,
            );
            assert_boundary::<Transaction>(
                |depth| {
                    format!(
                        "{{\"ops\":[{}],\"author\":{{\"kind\":\"user\"}}}}",
                        cell(depth, tag_first)
                    )
                },
                122,
            );
            assert_boundary::<Outer<Op>>(
                |depth| format!("{{\"op\":{},\"marker\":\"outer\"}}", cell(depth, tag_first)),
                123,
            );
        }
        let positional = |depth| {
            let nested = format!("{}0{}", "[".repeat(depth), "]".repeat(depth));
            format!(
                "[\"setCell\",0,{{\"row\":0,\"col\":0}},\
                 {{\"value\":{{\"kind\":\"empty\",\"unknown\":{nested}}}}}]"
            )
        };
        assert_boundary::<Op>(positional, 124);
        assert_boundary::<Vec<Op>>(|depth| format!("[{}]", positional(depth)), 123);
        assert_boundary::<Ops>(|depth| format!("{{\"ops\":[{}]}}", positional(depth)), 122);
        assert_boundary::<Transaction>(
            |depth| {
                format!(
                    "{{\"ops\":[{}],\"author\":{{\"kind\":\"user\"}}}}",
                    positional(depth)
                )
            },
            122,
        );
        assert_boundary::<Outer<Op>>(
            |depth| format!("{{\"op\":{},\"marker\":\"outer\"}}", positional(depth)),
            123,
        );
    }

    const INPUTS: &[&str] = &[
        r##""##,
        r##" "##,
        r##"{"##,
        r##"["##,
        r##"null"##,
        r##"true"##,
        r##"false"##,
        r##"0"##,
        r##"-1"##,
        r##"1.5"##,
        r##""removeSheet""##,
        r##"{}"##,
        r##"[]"##,
        r##"[null]"##,
        r##"[true]"##,
        r##"[0]"##,
        r##"[30,0]"##,
        r##"[37,0]"##,
        r##"[-1,0]"##,
        r##"[1.0,0]"##,
        r##"["removeSheet"]"##,
        r##"["removeSheet",0,1]"##,
        r##"["removeSheet",null]"##,
        r##"["removeSheet","0"]"##,
        r##"["removeSheet",0,{},[]]"##,
        r##"["addSheet",0]"##,
        r##"["addSheet",0,"Data",true]"##,
        r##"["insertRows",0,1]"##,
        r##"["insertRows",0,1,2,3]"##,
        r##"["setColWidth",0,1]"##,
        r##"["setDefinedName","N","A1"]"##,
        r##"{"index":0}"##,
        r##"{"type":"removeSheet","type":"addSheet","index":0}"##,
        r##"{"type":"removeSheet","t\u0079pe":"removeSheet","index":0}"##,
        r##"{"type":"removeSheet","type":null}"##,
        r##"{"type":"nope","type":"removeSheet"}"##,
        r##"{"type":"removeSheet","index":0,"index":1}"##,
        r##"{"index":0,"type":"removeSheet","index":1}"##,
        r##"{"index":0,"in\u0064ex":1,"type":"removeSheet"}"##,
        r##"{"type":"removeSheet","index":0,"in\u0064ex":1}"##,
        r##"{"type":"removeSheet","index":"bad","index":0}"##,
        r##"{"TYPE":"removeSheet","index":0}"##,
        r##"{"type":"RemoveSheet","index":0}"##,
        r##"{"type":"remove_sheet","index":0}"##,
        r##"{"type":"","index":0}"##,
        r##"{"type":null}"##,
        r##"{"type":true}"##,
        r##"{"type":[]}"##,
        r##"{"type":{}}"##,
        r##"{"type":-1}"##,
        r##"{"type":1.0}"##,
        r##"{"type":30,"index":0}"##,
        r##"{"type":37}"##,
        r##"{"type":18446744073709551615}"##,
        r##"{"type":"removeSheet"}"##,
        r##"{"type":"removeSheet","index":null}"##,
        r##"{"index":"bad","type":"unknown"}"##,
        r##"{"type":"removeSheet","index":"bad","type":"unknown"}"##,
        r##"{"index":"bad","type":"removeSheet","type":"removeSheet"}"##,
        r##"{"index":"bad"}"##,
        r##"{"type":"removeSheet","index":0,}"##,
        r##"{"type":"removeSheet","index":}"##,
        r##"{"type":"removeSheet","index":[}"##,
        r##"{"type":"removeSheet","index":0} trailing"##,
        "{\n  \"index\": \"bad\",\n  \"type\": \"removeSheet\"\n}",
        "{\n  \"index\": 0,\n  \"type\": \"unknown\"\n}",
        r##"{"type":"setAutoFilter","sheet":0}"##,
        r##"{"type":"setAutoFilter","sheet":0,"range":"not a range"}"##,
        r##"{"type":"setAutoFilter","sheet":0,"range":{}}"##,
        r##"{"type":"setDefinedName","name":"N","refers_to":"A1","hidden":0}"##,
        r##"{"type":"restoreRowStyles","sheet":0,"styles":{"-1":1}}"##,
        r##"{"type":"restoreRowStyles","sheet":0,"styles":{"4294967296":1}}"##,
        r##"{"type":"restoreRowStyles","sheet":0,"styles":{"3":"bad"}}"##,
        r##"{"type":"restoreRowStyles","sheet":0,"styles":[]}"##,
        r##"{"type":"insertRows","sheet":0,"at":0,"count":-1}"##,
        r##"{"type":"insertRows","sheet":0,"at":0,"count":-0}"##,
        r##"{"type":"insertRows","sheet":0,"at":0,"count":0.0}"##,
        r##"{"type":"insertRows","sheet":0,"at":0,"count":1e0}"##,
        r##"{"type":"insertRows","sheet":0,"at":0,"count":4294967296}"##,
        r##"{"type":"insertRows","sheet":0,"at":0,"count":18446744073709551615}"##,
        r##"{"type":"insertRows","sheet":0,"at":0,"count":18446744073709551616}"##,
        r##"{"type":"insertRows","sheet":0,"at":0,"count":-9223372036854775809}"##,
        r##"{"type":"insertRows","sheet":0,"at":0,"count":1e309}"##,
        r##"{"type":"insertRows","sheet":0,"at":0,"count":1e-400}"##,
        r##"{"type":"removeSheet","index":-1}"##,
        r##"{"type":"removeSheet","index":-0}"##,
        r##"{"type":"removeSheet","index":0.0}"##,
        r##"{"type":"removeSheet","index":18446744073709551616}"##,
        r##"{"type":"setColWidth","sheet":0,"col":0,"width":1.7976931348623159e308}"##,
        r##"{"type":"setCell","sheet":0,"at":{"row":0,"col":0},"cell":{"value":{}}}"##,
        r##"{"type":"setCell","sheet":0,"at":{"row":0,"col":0},"cell":{"value":null}}"##,
        concat!(
            r##"{"type":"setCell","sheet":0,"at":{"row":0,"col":0},"##,
            r##""cell":{"value":{"kind":"number","value":"bad"}}}"##,
        ),
        concat!(
            r##"{"type":"setCell","sheet":0,"at":{"row":0,"col":0},"##,
            r##""cell":{"value":{"kind":"error","value":"unknown"}}}"##,
        ),
        concat!(
            r##"{"type":"setCell","sheet":0,"at":{"row":0,"col":0},"##,
            r##""cell":{"value":{"kind":"unknown"}}}"##,
        ),
        concat!(
            r##"{"type":"setCell","sheet":0,"at":{"row":0,"col":0},"##,
            r##""cell":{"value":{"kind":"empty","kind":"number"}}}"##,
        ),
        concat!(
            r##"{"type":"patchRangeStyle","sheet":0,"range":{"start":{"row":0,"col":0},"##,
            r##""end":{"row":1,"col":1}},"patch":{"bold":true,"bold":false}}"##,
        ),
        concat!(
            r##"{"type":"patchRangeStyle","sheet":0,"range":{"start":{"row":0,"col":0},"##,
            r##""end":{"row":1,"col":1}},"patch":{"fontSize":1e309}}"##,
        ),
        concat!(
            r##"{"type":"patchRangeStyle","sheet":0,"range":{"start":{"row":0,"col":0},"##,
            r##""end":{"row":1,"col":1}},"patch":{"border":{"style":"unknown"}}}"##,
        ),
        concat!(
            r##"{"type":"patchRangeStyle","sheet":0,"range":{"start":{"row":0,"col":0},"##,
            r##""end":{"row":1,"col":1}},"patch":{"horizontalAlignment":0}}"##,
        ),
        concat!(
            r##"{"type":"patchRangeStyle","sheet":0,"range":{"start":{"row":0,"col":0},"##,
            r##""end":{"row":1,"col":1}},"patch":{"horizontalAlignment":{}}}"##,
        ),
        concat!(
            r##"{"type":"patchRangeStyle","sheet":0,"range":{"start":{"row":0,"col":0},"##,
            r##""end":{"row":1,"col":1}},"patch":{"clear":["bold","unknown"]}}"##,
        ),
        concat!(
            r##"{"type":"setRangeNumberFormat","sheet":0,"range":{"start":{"row":0,"col":0},"##,
            r##""end":{"row":1,"col":1}},"format":{"type":"custom"}}"##,
        ),
        concat!(
            r##"{"type":"setRangeNumberFormat","sheet":0,"range":{"start":{"row":0,"col":0},"##,
            r##""end":{"row":1,"col":1}},"format":{"type":"unknown"}}"##,
        ),
        concat!(
            r##"{"type":"patchRangeStyle","sheet":0,"range":{"start":{"row":0,"col":0},"##,
            r##""end":{"row":1,"col":1}},"patch":{"horizontalAlignment":{"##,
            r##""center":null,"left":null}}}"##,
        ),
        concat!(
            r##"{"type":"patchRangeStyle","sheet":0,"range":{"start":{"row":0,"col":0},"##,
            r##""end":{"row":1,"col":1}},"patch":{"horizontalAlignment":{"##,
            r##""unknown":null,"left":null}}}"##,
        ),
        concat!(
            r##"{"type":"setCell","sheet":0,"at":{"row":0,"col":0},"##,
            r##""cell":{"value":{"kind":5}}}"##,
        ),
        r##"{"type":"renameSheet","sheet":0,"name":"\ud800"}"##,
        r##"{"name":"\ud800","type":"unknown"}"##,
        r##"{"unknown":1e309,"type":"unknown"}"##,
        r##"{"type":"setCell","sheet":0,"at":[0,0,false,false,1],"cell":{}}"##,
        r##"["restoreSheet",0,"N",[[0,{"row":0,"col":0},{},1]]]"##,
    ];
}
