use std::collections::{BTreeMap, BTreeSet};
use std::hint::black_box;
use std::sync::OnceLock;
use std::time::Instant;

use serde_json::{Value, json};
use xlsx_engine::{Document, Host, Session};
use xlsx_model::open_timing::{self, Sample};

fn timer_ms() -> f64 {
    static START: OnceLock<Instant> = OnceLock::new();
    START.get_or_init(Instant::now).elapsed().as_secs_f64() * 1000.0
}

fn sample(bytes: &[u8]) -> Sample {
    open_timing::begin();
    let total = Instant::now();
    let opened = Instant::now();
    let document = Document::open(
        bytes,
        Host {
            now_serial: || Ok(Some(45000.25)),
            timer_ms,
            progress: || {},
            version: Session::version(),
        },
    )
    .unwrap();
    let session_open = opened.elapsed();
    let prepare = open_timing::span("worker_prepare");
    let info: Value = serde_json::from_str(&document.sheet_info_json().unwrap()).unwrap();
    let sheet = info["activeSheet"].as_u64().unwrap() as u32;
    black_box(document.document_version());
    black_box(document.sheet_info_json().unwrap());
    let catalog: Value =
        serde_json::from_str(&document.read_cells_json(r#"{"ranges":[]}"#).unwrap()).unwrap();
    assert_eq!(catalog["ok"], true);
    black_box(document.history_state_json().unwrap());
    let range = json!({ "sheet": sheet, "range": "A1" }).to_string();
    if catalog["sheets"][sheet as usize]["editable"] == true {
        black_box(document.selection_formatting_json(&range).unwrap());
    }
    black_box(document.merged_ranges_json(&range).unwrap());
    black_box(document.calculation_status_json().unwrap());
    let view: Value = serde_json::from_str(&document.sheet_info_for_json(sheet).unwrap()).unwrap();
    let edge = json!({
        "sheet": sheet,
        "row": view["frozenRows"].as_u64().unwrap().saturating_sub(1),
        "col": view["frozenCols"].as_u64().unwrap().saturating_sub(1),
    })
    .to_string();
    black_box(document.cell_rect_json(&edge).unwrap());
    let viewport = json!({
        "x": view["initialScrollX"], "y": view["initialScrollY"],
        "width": 800.0, "height": 600.0,
    })
    .to_string();
    drop(prepare);
    let frame = open_timing::span("first_viewport_frame");
    black_box(document.display_list_for_json(sheet, &viewport).unwrap());
    black_box(document.document_version());
    black_box(
        document
            .visible_merged_ranges_json(sheet, &viewport)
            .unwrap(),
    );
    drop(frame);
    let elapsed = total.elapsed();
    let mut sample = open_timing::finish();
    sample.phases.insert("session_open", session_open);
    sample.phases.insert("open_to_first_frame", elapsed);
    let rest = session_open
        .saturating_sub(sample.phases["package_parse"])
        .saturating_sub(sample.phases["workbook_base"])
        .saturating_sub(sample.phases["authority_rest"]);
    sample.phases.insert("session_rest", rest);
    let base_rest = sample.phases["workbook_base"]
        .saturating_sub(sample.phases["fingerprint_plan"])
        .saturating_sub(sample.phases["fingerprint_compute"])
        .saturating_sub(sample.phases["fingerprint_build"]);
    sample.phases.insert("workbook_base_rest", base_rest);
    for phase in [
        "parse_columns",
        "style_column_runs",
        "style_materialize",
        "style_inherited",
        "style_effective",
        "style_autofit",
        "style_fills",
        "style_borders",
    ] {
        sample.phases.entry(phase).or_default();
    }
    for name in [
        "style_column_runs_input",
        "style_column_runs_visited",
        "style_column_scan_runs",
        "style_materialize_rows",
        "style_materialize_col_runs",
        "style_inherited_rows",
        "style_inherited_cols",
        "style_effective_cells",
        "style_autofit_rows",
        "style_autofit_cells",
        "style_missing_row_cells",
        "style_fill_rows",
        "style_fill_cols",
        "style_fill_cells",
        "style_border_rows",
        "style_border_cols",
        "style_border_cells",
        "parse_dimension_cols",
        "parse_preserved_empty_cells",
    ] {
        sample.counts.entry(name).or_default();
    }
    sample
}

fn percentiles(mut values: Vec<f64>) -> (f64, f64) {
    values.sort_by(f64::total_cmp);
    let middle = values.len() / 2;
    let median = if values.len().is_multiple_of(2) {
        (values[middle - 1] + values[middle]) / 2.0
    } else {
        values[middle]
    };
    (median, values[(values.len() * 9).div_ceil(10) - 1])
}

#[test]
#[ignore]
fn worker_open_timing() {
    let path = std::env::var_os("BO_TIMING_XLSX").expect("set BO_TIMING_XLSX to an .xlsx path");
    let runs = match std::env::var("BO_TIMING_RUNS") {
        Ok(value) => value
            .parse::<usize>()
            .expect("BO_TIMING_RUNS must be a positive integer"),
        Err(std::env::VarError::NotPresent) => 20,
        Err(error) => panic!("invalid BO_TIMING_RUNS: {error}"),
    };
    assert!(runs > 0, "BO_TIMING_RUNS must be positive");
    let bytes = std::fs::read(path).unwrap();
    for _ in 0..3 {
        black_box(sample(&bytes));
    }
    let samples: Vec<_> = (0..runs).map(|_| sample(&bytes)).collect();
    let mut phases = BTreeMap::<_, Vec<f64>>::new();
    let mut counts = BTreeMap::<_, BTreeSet<u64>>::new();
    for sample in samples {
        for (phase, elapsed) in sample.phases {
            phases
                .entry(phase)
                .or_default()
                .push(elapsed.as_secs_f64() * 1000.0);
        }
        for (name, count) in sample.counts {
            counts.entry(name).or_default().insert(count);
        }
    }
    for (phase, values) in phases {
        assert_eq!(values.len(), runs, "phase missing from a run: {phase}");
        let (median, p90) = percentiles(values);
        println!("TIMING {phase} {median:.3} {p90:.3}");
    }
    for (name, values) in counts {
        println!("COUNT {name} {values:?}");
    }
}
