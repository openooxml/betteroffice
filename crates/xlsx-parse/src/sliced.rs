use ooxml_opc::WorkBudget;
use xlsx_model::{Sheet, SheetChart, Stylesheet, Table, Workbook};

pub async fn clone_charts(source: &[SheetChart], work: &WorkBudget) -> Vec<SheetChart> {
    let mut charts = Vec::with_capacity(source.len());
    for chart in source {
        work.step().await;
        charts.push(SheetChart {
            part: chart.part.clone(), drawing: chart.drawing.clone(),
            anchor_index: chart.anchor_index, anchor: chart.anchor,
            refs: work.clone_slice(&chart.refs).await,
        });
    }
    charts
}

pub async fn clone_tables(source: &[Table], work: &WorkBudget) -> Vec<Table> {
    let mut tables = Vec::with_capacity(source.len());
    for table in source {
        work.step().await;
        tables.push(Table {
            name: table.name.clone(), sheet: table.sheet, range: table.range,
            header_rows: table.header_rows, totals_rows: table.totals_rows,
            columns: work.clone_slice(&table.columns).await,
        });
    }
    tables
}

pub async fn retire_workbook(mut model: Workbook, work: &WorkBudget) {
    for sheet in &mut model.sheets {
        loop {
            let at = sheet.iter_cells().next().map(|(at, _)| at);
            let Some(at) = at else { break; };
            work.step().await;
            sheet.set_cell(at, xlsx_model::Cell::default());
        }
        while !sheet.col_widths.is_empty() { work.step().await; sheet.col_widths.pop_first(); }
        while !sheet.row_heights.is_empty() { work.step().await; sheet.row_heights.pop_first(); }
        loop {
            let at = sheet.array_formulas().next().map(|(at, _)| at);
            let Some(at) = at else { break; };
            work.step().await;
            sheet.clear_array_formula(at);
        }
        while !sheet.hyperlinks.is_empty() { work.step().await; sheet.hyperlinks.pop(); }
        while let Some(mut chart) = sheet.charts.pop() {
            work.step().await;
            while !chart.refs.is_empty() { work.step().await; chart.refs.pop(); }
        }
        while !sheet.merges.is_empty() { work.step().await; sheet.merges.pop(); }
        while !sheet.col_styles.is_empty() { work.step().await; sheet.col_styles.pop(); }
    }
    while !model.sheets.is_empty() { work.step().await; model.sheets.pop(); }
    while !model.defined_names.is_empty() { work.step().await; model.defined_names.pop(); }
    while !model.shared_strings.is_empty() { work.step().await; model.shared_strings.pop(); }
    while let Some(mut table) = model.tables.pop() {
        work.step().await;
        while !table.columns.is_empty() { work.step().await; table.columns.pop(); }
    }
    while !model.styles.fonts.is_empty() { work.step().await; model.styles.fonts.pop(); }
    while !model.styles.fills.is_empty() { work.step().await; model.styles.fills.pop(); }
    while !model.styles.borders.is_empty() { work.step().await; model.styles.borders.pop(); }
    while !model.styles.cell_xfs.is_empty() { work.step().await; model.styles.cell_xfs.pop(); }
    while !model.styles.num_fmts.is_empty() { work.step().await; model.styles.num_fmts.pop(); }
    while !model.styles.indexed_colors.is_empty() { work.step().await; model.styles.indexed_colors.pop(); }
}

pub async fn clone_stylesheet(source: &Stylesheet, work: &WorkBudget) -> Stylesheet {
    let mut styles = Stylesheet::default();
    styles.fonts = work.clone_slice(&source.fonts).await;
    styles.fills = work.clone_slice(&source.fills).await;
    styles.borders = work.clone_slice(&source.borders).await;
    styles.cell_xfs = work.clone_slice(&source.cell_xfs).await;
    styles.num_fmts = work.clone_slice(&source.num_fmts).await;
    styles.indexed_colors = work.clone_slice(&source.indexed_colors).await;
    styles.theme = source.theme.clone();
    styles
}

pub async fn clone_workbook(source: &Workbook, work: &WorkBudget) -> Workbook {
    let mut model = Workbook {
        date_system: source.date_system,
        defined_names: work.clone_slice(&source.defined_names).await,
        shared_strings: work.clone_slice(&source.shared_strings).await,
        styles: clone_stylesheet(&source.styles, work).await,
        tables: clone_tables(&source.tables, work).await,
        sheets: Vec::with_capacity(source.sheets.len()),
    };
    for source in &source.sheets {
        work.step().await;
        let mut sheet = Sheet::new(&source.name);
        sheet.freeze_pane = source.freeze_pane;
        sheet.format = source.format;
        sheet.hyperlinks = work.clone_slice(&source.hyperlinks).await;
        sheet.charts = clone_charts(&source.charts, work).await;
        sheet.merges = work.clone_slice(&source.merges).await;
        sheet.col_styles = work.clone_slice(&source.col_styles).await;
        for (&at, &width) in &source.col_widths {
            work.step().await;
            sheet.col_widths.insert(at, width);
        }
        for (&at, &height) in &source.row_heights {
            work.step().await;
            sheet.row_heights.insert(at, height);
        }
        for (at, cell) in source.iter_cells() {
            work.step().await;
            sheet.set_cell(at, cell.clone());
        }
        for (at, range) in source.array_formulas() {
            work.step().await;
            sheet.set_array_formula(at, range);
        }
        model.sheets.push(sheet);
    }
    model
}
