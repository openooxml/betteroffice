use std::alloc::{GlobalAlloc, Layout, System};
use std::sync::atomic::{AtomicUsize, Ordering};
use std::thread;
use std::time::{Duration, Instant};

use betteroffice_xlsx::{
    CalculationOptions, Cell, CellRef, CellValue, DefinedName, DefinedNamesRequest, Op, Sheet,
    SheetId, Workbook, WorkbookModel,
};

/// Tracks live heap bytes and their high-water mark.
struct Peak;

static LIVE: AtomicUsize = AtomicUsize::new(0);
static HIGH: AtomicUsize = AtomicUsize::new(0);

unsafe impl GlobalAlloc for Peak {
    unsafe fn alloc(&self, layout: Layout) -> *mut u8 {
        let ptr = unsafe { System.alloc(layout) };
        if !ptr.is_null() {
            let live = LIVE.fetch_add(layout.size(), Ordering::Relaxed) + layout.size();
            HIGH.fetch_max(live, Ordering::Relaxed);
        }
        ptr
    }

    unsafe fn dealloc(&self, ptr: *mut u8, layout: Layout) {
        unsafe { System.dealloc(ptr, layout) };
        LIVE.fetch_sub(layout.size(), Ordering::Relaxed);
    }
}

#[global_allocator]
static ALLOCATOR: Peak = Peak;

const NAMES: usize = 20_000;
const HEAP_BOUND: usize = 256 << 20;
const TIME_BOUND: Duration = Duration::from_secs(120);

/// Three sheets and `count` names: workbook and sheet scopes, one name in two
/// scopes, names that reference other names, and a hidden filter range.
fn package(count: usize) -> Vec<u8> {
    let mut model = WorkbookModel::default();
    for name in ["Data", "Second", "Third"] {
        let mut sheet = Sheet::new(name);
        for row in 0..100 {
            sheet.set_cell(
                CellRef::new(row, 0),
                Cell {
                    value: CellValue::Number {
                        value: f64::from(row),
                    },
                    ..Cell::default()
                },
            );
        }
        model.sheets.push(sheet);
    }
    model.defined_names.push(DefinedName {
        name: "_xlnm._FilterDatabase".into(),
        formula: "Data!$A$1:$A$10".into(),
        local_sheet: Some(SheetId(0)),
        hidden: true,
    });
    model
        .defined_names
        .extend((0..count).map(|index| match index % 10 {
            0..=5 => DefinedName {
                name: format!("Name{index}"),
                formula: format!("Data!$A${}", index % 100 + 1),
                local_sheet: None,
                hidden: false,
            },
            6 | 7 => DefinedName {
                name: format!("Local{index}"),
                formula: format!("Second!$A$1:$A${}", index % 100 + 1),
                local_sheet: Some(SheetId((index % 3) as u32)),
                hidden: false,
            },
            8 => DefinedName {
                name: format!("Name{}", index - 8),
                formula: "Second!$A$2".into(),
                local_sheet: Some(SheetId(1)),
                hidden: false,
            },
            _ => DefinedName {
                name: format!("Ref{index}"),
                formula: format!("Name{}*2+SUM(Name{},1)", index - 9, index - 8),
                local_sheet: None,
                hidden: false,
            },
        }));
    for (row, formula) in ["Name0*2", "SUM(Ref9)", "Name1+Name2"]
        .into_iter()
        .enumerate()
    {
        model.sheets[0].set_cell(
            CellRef::new(row as u32, 1),
            Cell {
                value: CellValue::Number { value: 0.0 },
                formula: Some(formula.into()),
                style: None,
            },
        );
    }
    let parts = xlsx_parse::serialize_workbook(&model).unwrap();
    ooxml_opc::rezip_parts(&parts).unwrap()
}

fn recommit_ops(workbook: &Workbook) -> Vec<Op> {
    let mut ops = Vec::new();
    let mut offset = 0;
    loop {
        let page = workbook
            .list_defined_names(&DefinedNamesRequest {
                offset: Some(offset),
                limit: Some(100),
                ..Default::default()
            })
            .unwrap();
        ops.extend(page.items.into_iter().map(|info| Op::SetDefinedName {
            name: info.name,
            refers_to: info.refers_to,
            scope: info.scope.map(SheetId),
            hidden: None,
        }));
        match page.next_offset {
            Some(next) if next > offset => offset = next,
            _ => return ops,
        }
    }
}

#[test]
fn recommitting_every_name_of_a_large_table_in_one_batch_takes_bounded_memory_and_time() {
    let bytes = package(NAMES);
    let mut workbook = Workbook::open(&bytes).unwrap();
    let version = workbook.version();
    let ops = recommit_ops(&workbook);
    assert_eq!(ops.len(), NAMES + 1);

    let start = LIVE.load(Ordering::Relaxed);
    HIGH.store(start, Ordering::Relaxed);
    let started = Instant::now();
    let batch = thread::spawn(move || {
        let applied = workbook
            .apply_ops(ops, CalculationOptions::default())
            .map(|result| result.applied);
        (workbook, applied)
    });
    while !batch.is_finished() {
        let held = HIGH.load(Ordering::Relaxed).saturating_sub(start);
        assert!(
            held <= HEAP_BOUND,
            "the batch held {} MiB more than before it started",
            held >> 20
        );
        assert!(
            started.elapsed() <= TIME_BOUND,
            "the batch ran past {TIME_BOUND:?}"
        );
        thread::sleep(Duration::from_millis(5));
    }
    let (workbook, applied) = batch.join().unwrap();
    let held = HIGH.load(Ordering::Relaxed).saturating_sub(start);

    assert!(!applied.unwrap());
    assert!(
        held <= HEAP_BOUND,
        "the batch held {} MiB more than before it started",
        held >> 20
    );
    assert_eq!(workbook.version(), version);
    assert_eq!(workbook.save().unwrap(), bytes);
}
