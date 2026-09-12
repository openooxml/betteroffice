use ooxml_drawingml::{
    ColorValue, ResolvedCellStyle, ShapeFill, ShapeOutline, TableCellBorder, TableCellBorders,
    TableCellPosition, TableCellStyle, TableStyle, TableStyleFlags, TableStyleList, TableStylePart,
    TableTextStyle,
};
use proptest::collection::vec;
use proptest::option;
use proptest::prelude::*;
use proptest::test_runner::TestCaseError;

const IDS: [&str; 4] = ["{A}", "{B}", "{C}", "{D}"];
const MAX_SIDE: usize = 4;
const MAX_BANDED_ROWS: usize = 6;

/// A table style part.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
enum Part {
    WholeTable,
    Band1Row,
    Band2Row,
    FirstColumn,
    LastColumn,
    FirstRow,
    LastRow,
}

const PRECEDENCE: [Part; 7] = [
    Part::WholeTable,
    Part::Band1Row,
    Part::Band2Row,
    Part::FirstColumn,
    Part::LastColumn,
    Part::FirstRow,
    Part::LastRow,
];

/// How a generated style defines one part: absent, text only, or text with fill and borders.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
enum Defined {
    Absent,
    TextOnly,
    Full(Borders),
}

/// One generated `a:tcBdr` edge: unset, an `a:noFill` clear, or a sentinel line.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
enum Stroke {
    Unset,
    Clear,
    Line,
}

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
struct Borders {
    left: Stroke,
    right: Stroke,
    top: Stroke,
    bottom: Stroke,
    inside_horizontal: Stroke,
    inside_vertical: Stroke,
}

const UNSET_BORDERS: Borders = Borders {
    left: Stroke::Unset,
    right: Stroke::Unset,
    top: Stroke::Unset,
    bottom: Stroke::Unset,
    inside_horizontal: Stroke::Unset,
    inside_vertical: Stroke::Unset,
};

impl Borders {
    fn get(self, edge: Edge) -> Stroke {
        match edge {
            Edge::Left => self.left,
            Edge::Right => self.right,
            Edge::Top => self.top,
            Edge::Bottom => self.bottom,
            Edge::InsideHorizontal => self.inside_horizontal,
            Edge::InsideVertical => self.inside_vertical,
        }
    }
}

/// An edge of a style part's `a:tcBdr`.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
enum Edge {
    Left,
    Right,
    Top,
    Bottom,
    InsideHorizontal,
    InsideVertical,
}

const EDGES: [Edge; 6] = [
    Edge::Left,
    Edge::Right,
    Edge::Top,
    Edge::Bottom,
    Edge::InsideHorizontal,
    Edge::InsideVertical,
];

/// A side of a resolved cell.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
enum Side {
    Left,
    Right,
    Top,
    Bottom,
}

const SIDES: [Side; 4] = [Side::Left, Side::Right, Side::Top, Side::Bottom];

#[derive(Clone, Copy, Debug)]
struct Parts {
    whole_table: Defined,
    band1_row: Defined,
    band2_row: Defined,
    first_column: Defined,
    last_column: Defined,
    first_row: Defined,
    last_row: Defined,
}

impl Parts {
    fn get(self, part: Part) -> Defined {
        match part {
            Part::WholeTable => self.whole_table,
            Part::Band1Row => self.band1_row,
            Part::Band2Row => self.band2_row,
            Part::FirstColumn => self.first_column,
            Part::LastColumn => self.last_column,
            Part::FirstRow => self.first_row,
            Part::LastRow => self.last_row,
        }
    }
}

fn rgb(value: String) -> ColorValue {
    ColorValue {
        rgb: Some(value),
        ..ColorValue::default()
    }
}

fn solid(value: &str) -> ShapeFill {
    ShapeFill {
        fill_type: "solid".to_owned(),
        color: Some(rgb(value.to_owned())),
        gradient: None,
    }
}

fn line(width: f64, value: String) -> TableCellBorder {
    TableCellBorder::Line(Box::new(ShapeOutline {
        width: Some(width),
        color: Some(rgb(value)),
        ..ShapeOutline::default()
    }))
}

fn fill_sentinel(style: usize, part: Part) -> ShapeFill {
    solid(&format!("F{style}{}FFF", part as usize))
}

fn text_sentinel(style: usize, part: Part) -> ColorValue {
    rgb(format!("C{style}{}CCC", part as usize))
}

fn border_sentinel(style: usize, part: Part, edge: Edge) -> ShapeOutline {
    ShapeOutline {
        width: Some(12700.0),
        color: Some(rgb(format!("B{style}{}{}BB", part as usize, edge as usize))),
        ..ShapeOutline::default()
    }
}

fn build_edge(style: usize, part: Part, edge: Edge, stroke: Stroke) -> Option<TableCellBorder> {
    match stroke {
        Stroke::Unset => None,
        Stroke::Clear => Some(TableCellBorder::None),
        Stroke::Line => Some(TableCellBorder::Line(Box::new(border_sentinel(
            style, part, edge,
        )))),
    }
}

fn build_part(style: usize, part: Part, defined: Defined) -> Option<TableStylePart> {
    let text = TableTextStyle {
        color: Some(text_sentinel(style, part)),
        ..TableTextStyle::default()
    };
    match defined {
        Defined::Absent => None,
        Defined::TextOnly => Some(TableStylePart {
            text,
            cell: TableCellStyle::default(),
        }),
        Defined::Full(borders) => {
            let edge = |edge| build_edge(style, part, edge, borders.get(edge));
            Some(TableStylePart {
                text,
                cell: TableCellStyle {
                    fill: Some(fill_sentinel(style, part)),
                    borders: TableCellBorders {
                        left: edge(Edge::Left),
                        right: edge(Edge::Right),
                        top: edge(Edge::Top),
                        bottom: edge(Edge::Bottom),
                        inside_horizontal: edge(Edge::InsideHorizontal),
                        inside_vertical: edge(Edge::InsideVertical),
                    },
                },
            })
        }
    }
}

fn build_style(style: usize, parts: Parts) -> TableStyle {
    let part = |part| build_part(style, part, parts.get(part));
    TableStyle {
        style_id: IDS[style].to_owned(),
        style_name: None,
        whole_table: part(Part::WholeTable),
        band1_row: part(Part::Band1Row),
        band2_row: part(Part::Band2Row),
        first_row: part(Part::FirstRow),
        last_row: part(Part::LastRow),
        first_column: part(Part::FirstColumn),
        last_column: part(Part::LastColumn),
    }
}

fn style_list(styles: &[Parts], default: Option<usize>) -> TableStyleList {
    TableStyleList {
        default_style_id: default.map(|index| IDS[index].to_owned()),
        styles: styles
            .iter()
            .enumerate()
            .map(|(index, &parts)| build_style(index, parts))
            .collect(),
    }
}

fn is_header(flags: TableStyleFlags, cell: TableCellPosition) -> bool {
    flags.first_row && cell.row == 0
}

fn is_total(flags: TableStyleFlags, cell: TableCellPosition) -> bool {
    flags.last_row && cell.row + 1 == cell.row_count
}

/// The cell's index among banded rows, counted from the first row after any header.
fn band_index(flags: TableStyleFlags, cell: TableCellPosition) -> Option<usize> {
    (flags.band_row && !is_header(flags, cell) && !is_total(flags, cell))
        .then(|| cell.row - usize::from(flags.first_row))
}

/// Whether `part` covers `cell` under `flags`, restated independently of the resolver.
fn applies(part: Part, flags: TableStyleFlags, cell: TableCellPosition) -> bool {
    match part {
        Part::WholeTable => true,
        Part::Band1Row => band_index(flags, cell).is_some_and(|index| index.is_multiple_of(2)),
        Part::Band2Row => band_index(flags, cell).is_some_and(|index| !index.is_multiple_of(2)),
        Part::FirstColumn => flags.first_column && cell.column == 0,
        Part::LastColumn => flags.last_column && cell.column + 1 == cell.column_count,
        Part::FirstRow => is_header(flags, cell),
        Part::LastRow => is_total(flags, cell),
    }
}

fn winner(
    parts: Parts,
    flags: TableStyleFlags,
    cell: TableCellPosition,
    carries: impl Fn(Defined) -> bool,
) -> Option<Part> {
    PRECEDENCE
        .into_iter()
        .rev()
        .find(|&part| applies(part, flags, cell) && carries(parts.get(part)))
}

/// The part's own `side` edge when `cell` is on that side of the part's region, else its
/// inside edge. `wholeTbl` covers the table, a column part its column, any other part its row.
fn selected_edge(part: Part, side: Side, cell: TableCellPosition) -> Edge {
    let (spans_rows, spans_columns) = match part {
        Part::WholeTable => (true, true),
        Part::FirstColumn | Part::LastColumn => (true, false),
        _ => (false, true),
    };
    let on_boundary = match side {
        Side::Left => !spans_columns || cell.column == 0,
        Side::Right => !spans_columns || cell.column + 1 == cell.column_count,
        Side::Top => !spans_rows || cell.row == 0,
        Side::Bottom => !spans_rows || cell.row + 1 == cell.row_count,
    };
    match (side, on_boundary) {
        (Side::Left, true) => Edge::Left,
        (Side::Right, true) => Edge::Right,
        (Side::Top, true) => Edge::Top,
        (Side::Bottom, true) => Edge::Bottom,
        (Side::Left | Side::Right, false) => Edge::InsideVertical,
        (Side::Top | Side::Bottom, false) => Edge::InsideHorizontal,
    }
}

/// The part and edge that should draw `side`, or `None` when no line survives the cascade.
fn expected_edge(
    parts: Parts,
    flags: TableStyleFlags,
    cell: TableCellPosition,
    side: Side,
) -> Option<(Part, Edge)> {
    PRECEDENCE
        .into_iter()
        .filter(|&part| applies(part, flags, cell))
        .fold(None, |current, part| {
            let Defined::Full(borders) = parts.get(part) else {
                return current;
            };
            let edge = selected_edge(part, side, cell);
            match borders.get(edge) {
                Stroke::Unset => current,
                Stroke::Clear => None,
                Stroke::Line => Some((part, edge)),
            }
        })
}

fn resolved_side(resolved: &ResolvedCellStyle, side: Side) -> Option<&ShapeOutline> {
    match side {
        Side::Left => resolved.left.as_ref(),
        Side::Right => resolved.right.as_ref(),
        Side::Top => resolved.top.as_ref(),
        Side::Bottom => resolved.bottom.as_ref(),
    }
}

fn drawn_by(outline: Option<&ShapeOutline>) -> Option<(Part, Edge)> {
    let outline = outline?;
    PRECEDENCE
        .into_iter()
        .flat_map(|part| EDGES.map(|edge| (part, edge)))
        .find(|&(part, edge)| *outline == border_sentinel(0, part, edge))
}

fn check_precedence(
    parts: Parts,
    flags: TableStyleFlags,
    cell: TableCellPosition,
) -> Result<(), TestCaseError> {
    let resolved = build_style(0, parts).resolve_cell(flags, cell);
    for side in SIDES {
        let expected = expected_edge(parts, flags, cell, side);
        let actual = resolved_side(&resolved, side);
        prop_assert_eq!(
            actual.cloned(),
            expected.map(|(part, edge)| border_sentinel(0, part, edge)),
            "{:?} should come from {:?}, not {:?}",
            side,
            expected,
            drawn_by(actual)
        );
    }
    let fill_owner = winner(parts, flags, cell, |defined| {
        matches!(defined, Defined::Full(_))
    });
    let text_owner = winner(parts, flags, cell, |defined| defined != Defined::Absent);

    prop_assert_eq!(
        resolved.fill,
        fill_owner.map(|part| fill_sentinel(0, part)),
        "{:?} should own the fill",
        fill_owner
    );
    prop_assert_eq!(
        resolved.color,
        text_owner.map(|part| text_sentinel(0, part)),
        "{:?} should own the text colour",
        text_owner
    );
    Ok(())
}

fn inherit_unless_set(
    own: &Option<TableCellBorder>,
    inherited: &Option<ShapeOutline>,
) -> Option<ShapeOutline> {
    match own {
        None => inherited.clone(),
        Some(TableCellBorder::None) => None,
        Some(TableCellBorder::Line(outline)) => Some(outline.as_ref().clone()),
    }
}

fn stroke() -> impl Strategy<Value = Stroke> {
    prop_oneof![Just(Stroke::Unset), Just(Stroke::Clear), Just(Stroke::Line)]
}

prop_compose! {
    fn borders()(
        left in stroke(),
        right in stroke(),
        top in stroke(),
        bottom in stroke(),
        inside_horizontal in stroke(),
        inside_vertical in stroke(),
    ) -> Borders {
        Borders { left, right, top, bottom, inside_horizontal, inside_vertical }
    }
}

fn defined() -> impl Strategy<Value = Defined> {
    prop_oneof![
        Just(Defined::Absent),
        Just(Defined::TextOnly),
        borders().prop_map(Defined::Full)
    ]
}

prop_compose! {
    fn parts()(
        whole_table in defined(),
        band1_row in defined(),
        band2_row in defined(),
        first_column in defined(),
        last_column in defined(),
        first_row in defined(),
        last_row in defined(),
    ) -> Parts {
        Parts { whole_table, band1_row, band2_row, first_column, last_column, first_row, last_row }
    }
}

prop_compose! {
    fn flags()(
        first_row in any::<bool>(),
        last_row in any::<bool>(),
        first_column in any::<bool>(),
        last_column in any::<bool>(),
        band_row in any::<bool>(),
    ) -> TableStyleFlags {
        TableStyleFlags { first_row, last_row, first_column, last_column, band_row }
    }
}

fn position(max_side: usize) -> impl Strategy<Value = TableCellPosition> {
    (1..=max_side, 1..=max_side).prop_flat_map(|(row_count, column_count)| {
        (0..row_count, 0..column_count).prop_map(move |(row, column)| TableCellPosition {
            row,
            column,
            row_count,
            column_count,
        })
    })
}

fn own_edge() -> impl Strategy<Value = Option<TableCellBorder>> {
    prop_oneof![
        Just(None),
        Just(Some(TableCellBorder::None)),
        Just(Some(line(19050.0, "D1D1D1".to_owned()))),
    ]
}

prop_compose! {
    fn own_cell_style()(
        fill in prop_oneof![Just(solid("D0D0D0")), Just(ShapeFill::named("none"))],
        left in own_edge(),
        right in own_edge(),
        top in own_edge(),
        bottom in own_edge(),
    ) -> TableCellStyle {
        TableCellStyle {
            fill: Some(fill),
            borders: TableCellBorders {
                left,
                right,
                top,
                bottom,
                inside_horizontal: None,
                inside_vertical: None,
            },
        }
    }
}

fn styles_and_an_undefined_id() -> impl Strategy<Value = (Vec<Parts>, usize)> {
    vec(parts(), 0..IDS.len()).prop_flat_map(|styles| {
        let defined = styles.len();
        (Just(styles), defined..IDS.len())
    })
}

proptest! {
    #[test]
    fn a_cells_own_formatting_wins_over_any_style(
        parts in parts(),
        flags in flags(),
        cell in position(MAX_SIDE),
        own in own_cell_style(),
    ) {
        let inherited = build_style(0, parts).resolve_cell(flags, cell);
        let mut resolved = inherited.clone();
        resolved.apply_cell_style(&own);

        prop_assert_eq!(&resolved.fill, &own.fill);
        prop_assert_eq!(&resolved.left, &inherit_unless_set(&own.borders.left, &inherited.left));
        prop_assert_eq!(&resolved.right, &inherit_unless_set(&own.borders.right, &inherited.right));
        prop_assert_eq!(&resolved.top, &inherit_unless_set(&own.borders.top, &inherited.top));
        prop_assert_eq!(
            &resolved.bottom,
            &inherit_unless_set(&own.borders.bottom, &inherited.bottom)
        );
        prop_assert_eq!(&resolved.color, &inherited.color);
    }

    #[test]
    fn the_highest_applicable_part_wins(
        parts in parts(),
        flags in flags(),
        cell in position(MAX_SIDE),
    ) {
        check_precedence(parts, flags, cell)?;
    }

    #[test]
    fn the_highest_applicable_part_wins_in_a_one_by_one_table(
        parts in parts(),
        flags in flags(),
        cell in position(1),
    ) {
        check_precedence(parts, flags, cell)?;
    }

    #[test]
    fn banded_rows_alternate_from_band1_after_the_header(
        whole_table in defined(),
        flags in flags(),
        cell in position(MAX_BANDED_ROWS),
    ) {
        let flags = TableStyleFlags { band_row: true, ..flags };
        let style = build_style(0, Parts {
            whole_table,
            band1_row: Defined::Full(UNSET_BORDERS),
            band2_row: Defined::Full(UNSET_BORDERS),
            first_column: Defined::Absent,
            last_column: Defined::Absent,
            first_row: Defined::Absent,
            last_row: Defined::Absent,
        });
        let bands: Vec<Option<Part>> = (0..cell.row_count)
            .map(|row| {
                let fill = style.resolve_cell(flags, TableCellPosition { row, ..cell }).fill;
                [Part::Band1Row, Part::Band2Row]
                    .into_iter()
                    .find(|&part| fill == Some(fill_sentinel(0, part)))
            })
            .collect();
        let banded: Vec<usize> = (0..cell.row_count)
            .filter(|&row| {
                let cell = TableCellPosition { row, ..cell };
                !is_header(flags, cell) && !is_total(flags, cell)
            })
            .collect();

        for (row, band) in bands.iter().enumerate() {
            if !banded.contains(&row) {
                prop_assert_eq!(*band, None, "row {} is a header or total row", row);
            }
        }
        if let Some(&first) = banded.first() {
            prop_assert_eq!(bands[first], Some(Part::Band1Row), "first banded row {}", first);
        }
        for pair in banded.windows(2) {
            prop_assert!(
                bands[pair[1]].is_some() && bands[pair[0]] != bands[pair[1]],
                "rows {} and {} took {:?} and {:?}",
                pair[0],
                pair[1],
                bands[pair[0]],
                bands[pair[1]]
            );
        }
    }

    #[test]
    fn an_id_the_package_does_not_define_resolves_to_no_style(
        (styles, undefined) in styles_and_an_undefined_id(),
        default in option::of(0..IDS.len()),
        flags in flags(),
        cell in position(MAX_SIDE),
    ) {
        let list = style_list(&styles, default);

        prop_assert_eq!(
            list.resolve_cell(Some(IDS[undefined]), flags, cell),
            ResolvedCellStyle::default()
        );
    }

    #[test]
    fn a_table_naming_no_style_resolves_through_def(
        styles in vec(parts(), 0..IDS.len()),
        default in option::of(0..IDS.len()),
        flags in flags(),
        cell in position(MAX_SIDE),
    ) {
        let list = style_list(&styles, default);
        let expected = default
            .and_then(|index| list.styles.get(index))
            .map(|style| style.resolve_cell(flags, cell))
            .unwrap_or_default();

        prop_assert_eq!(list.resolve_cell(None, flags, cell), expected);
    }
}
