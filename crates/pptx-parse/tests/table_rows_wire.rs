use ooxml_drawingml::{ColorValue, ShapeFill, ShapeOutline};
use pptx_parse::{
    GraphicFrameData, RunProperties, Table, TableCell, TableCellBorders, TableRow, TextBody,
    TextParagraph, TextRun,
};
use proptest::collection::vec;
use proptest::option;
use proptest::prelude::*;
use proptest::sample::Index;
use serde_json::Value;

const EMU_PER_POINT: f64 = 12_700.0;

fn colour() -> impl Strategy<Value = ColorValue> {
    prop_oneof![
        prop_oneof![Just("C00000"), Just("DDEBF7")].prop_map(|rgb| ColorValue {
            rgb: Some(rgb.to_owned()),
            ..ColorValue::default()
        }),
        Just(ColorValue {
            theme_color: Some("accent1".to_owned()),
            ..ColorValue::default()
        }),
    ]
}

fn run() -> impl Strategy<Value = TextRun> {
    ("[ab ]{0,3}", option::of(any::<bool>())).prop_map(|(text, bold)| TextRun {
        text,
        properties: RunProperties {
            bold,
            ..RunProperties::default()
        },
        field_id: None,
        field_type: None,
        line_break: false,
    })
}

fn text_body() -> impl Strategy<Value = TextBody> {
    (
        option::of(prop_oneof![Just("t"), Just("ctr"), Just("b")]),
        option::of(0..200_000_i64),
        vec(
            vec(run(), 0..3).prop_map(|runs| TextParagraph {
                runs,
                ..TextParagraph::default()
            }),
            0..3,
        ),
    )
        .prop_map(|(anchor, inset_left, paragraphs)| TextBody {
            anchor: anchor.map(str::to_owned),
            inset_left,
            paragraphs,
            ..TextBody::default()
        })
}

fn fill() -> impl Strategy<Value = ShapeFill> {
    prop_oneof![
        Just(ShapeFill::named("none")),
        Just(ShapeFill::named("picture")),
        option::of(colour()).prop_map(|color| ShapeFill {
            color,
            ..ShapeFill::named("solid")
        }),
    ]
}

fn border() -> impl Strategy<Value = ShapeOutline> {
    prop_oneof![
        Just(ShapeOutline::default()),
        (1..=4_u32, option::of(colour()), option::of(Just("dash"))).prop_map(
            |(points, color, style)| ShapeOutline {
                width: Some(f64::from(points) * EMU_PER_POINT),
                color,
                style: style.map(str::to_owned),
                ..ShapeOutline::default()
            }
        ),
    ]
}

fn borders() -> impl Strategy<Value = TableCellBorders> {
    (
        option::of(border()),
        option::of(border()),
        option::of(border()),
        option::of(border()),
    )
        .prop_map(|(left, top, right, bottom)| TableCellBorders {
            left,
            top,
            right,
            bottom,
        })
}

fn one_border() -> impl Strategy<Value = TableCellBorders> {
    (0..4_usize, border()).prop_map(|(side, line)| {
        let mut borders = TableCellBorders::default();
        *[
            &mut borders.left,
            &mut borders.top,
            &mut borders.right,
            &mut borders.bottom,
        ]
        .into_iter()
        .nth(side)
        .unwrap() = Some(line);
        borders
    })
}

fn plain_cell() -> impl Strategy<Value = TableCell> {
    text_body().prop_map(TableCell::from_text)
}

fn any_cell() -> impl Strategy<Value = TableCell> {
    (
        text_body(),
        1..=3_u32,
        1..=3_u32,
        any::<bool>(),
        option::of(fill()),
        borders(),
    )
        .prop_map(
            |(text, grid_span, row_span, merged, fill, borders)| TableCell {
                text,
                grid_span,
                row_span,
                merged,
                fill,
                borders,
            },
        )
}

fn row() -> impl Strategy<Value = TableRow> {
    (
        prop_oneof![3 => Just(0_i64), 1 => 1..1_000_000_i64],
        vec(prop_oneof![2 => plain_cell(), 1 => any_cell()], 0..4),
    )
        .prop_map(|(height, cells)| TableRow { height, cells })
}

fn plain_row() -> impl Strategy<Value = TableRow> {
    vec(plain_cell(), 0..4).prop_map(|cells| TableRow { height: 0, cells })
}

/// One thing a row can carry beyond cell text, each enough to leave the released encoding.
#[derive(Clone, Debug)]
enum Beyond {
    Height(i64),
    GridSpan(u32),
    RowSpan(u32),
    Merged,
    Fill(Box<ShapeFill>),
    Borders(Box<TableCellBorders>),
}

fn beyond() -> impl Strategy<Value = Beyond> {
    prop_oneof![
        (1..1_000_000_i64).prop_map(Beyond::Height),
        (2..=3_u32).prop_map(Beyond::GridSpan),
        (2..=3_u32).prop_map(Beyond::RowSpan),
        Just(Beyond::Merged),
        fill().prop_map(|fill| Beyond::Fill(Box::new(fill))),
        one_border().prop_map(|borders| Beyond::Borders(Box::new(borders))),
    ]
}

fn decorated_row() -> impl Strategy<Value = TableRow> {
    (
        vec(plain_cell(), 0..3),
        plain_cell(),
        beyond(),
        any::<Index>(),
    )
        .prop_map(|(mut cells, mut cell, beyond, at)| {
            let mut height = 0;
            match beyond {
                Beyond::Height(value) => height = value,
                Beyond::GridSpan(span) => cell.grid_span = span,
                Beyond::RowSpan(span) => cell.row_span = span,
                Beyond::Merged => cell.merged = true,
                Beyond::Fill(fill) => cell.fill = Some(*fill),
                Beyond::Borders(borders) => cell.borders = *borders,
            }
            cells.insert(at.index(cells.len() + 1), cell);
            TableRow { height, cells }
        })
}

fn frame(rows: Vec<TableRow>) -> GraphicFrameData {
    GraphicFrameData::Table(Table {
        rows,
        ..Table::default()
    })
}

fn wire_rows(json: &str) -> Vec<Value> {
    let mut frame: Value = serde_json::from_str(json).unwrap();
    let Value::Array(rows) = frame["rows"].take() else {
        panic!("rows must serialize as an array: {json}");
    };
    rows
}

proptest! {
    #[test]
    fn table_rows_survive_the_stored_json(rows in vec(row(), 0..4)) {
        let data = frame(rows);
        let json = serde_json::to_string(&data).unwrap();
        prop_assert_eq!(serde_json::from_str::<GraphicFrameData>(&json).unwrap(), data);
    }

    #[test]
    fn text_only_rows_keep_the_released_array_shape(texts in vec(vec(text_body(), 0..4), 0..4)) {
        let rows = texts
            .iter()
            .map(|row| TableRow {
                height: 0,
                cells: row.iter().cloned().map(TableCell::from_text).collect(),
            })
            .collect();
        let data = frame(rows);
        let json = serde_json::to_string(&data).unwrap();
        let wire = wire_rows(&json);
        prop_assert_eq!(wire.len(), texts.len());
        for (encoded, row) in wire.iter().zip(&texts) {
            prop_assert_eq!(encoded, &serde_json::to_value(row).unwrap());
        }
        prop_assert_eq!(serde_json::from_str::<GraphicFrameData>(&json).unwrap(), data);
    }

    #[test]
    fn a_row_carrying_anything_beyond_text_never_takes_the_array_shape(
        before in vec(plain_row(), 0..3),
        decorated in decorated_row(),
        after in vec(plain_row(), 0..3),
    ) {
        let at = before.len();
        let rows: Vec<_> = before.into_iter().chain([decorated]).chain(after).collect();
        let data = frame(rows);
        let json = serde_json::to_string(&data).unwrap();
        let wire = wire_rows(&json);
        prop_assert!(wire[at].get("cells").is_some(), "decorated row lost its cells object: {}", json);
        prop_assert!(wire.iter().all(Value::is_object), "the encoding is chosen per table: {}", json);
        prop_assert_eq!(serde_json::from_str::<GraphicFrameData>(&json).unwrap(), data);
    }
}
