use std::alloc::{GlobalAlloc, Layout, System};
use std::hint::black_box;
use std::sync::atomic::{AtomicUsize, Ordering};

use xlsx_model::{Cell, CellRange, CellRef, CellValue, Sheet, SheetId, Workbook};
use xlsx_ops::{Op, StylePatch, apply};

struct TrackingAllocator;

static LIVE_BYTES: AtomicUsize = AtomicUsize::new(0);
static PEAK_BYTES: AtomicUsize = AtomicUsize::new(0);

fn record_allocation(size: usize) {
    let live = LIVE_BYTES.fetch_add(size, Ordering::SeqCst) + size;
    PEAK_BYTES.fetch_max(live, Ordering::SeqCst);
}

unsafe impl GlobalAlloc for TrackingAllocator {
    unsafe fn alloc(&self, layout: Layout) -> *mut u8 {
        let ptr = unsafe { System.alloc(layout) };
        if !ptr.is_null() {
            record_allocation(layout.size());
        }
        ptr
    }

    unsafe fn dealloc(&self, ptr: *mut u8, layout: Layout) {
        unsafe { System.dealloc(ptr, layout) };
        LIVE_BYTES.fetch_sub(layout.size(), Ordering::SeqCst);
    }

    unsafe fn realloc(&self, ptr: *mut u8, layout: Layout, new_size: usize) -> *mut u8 {
        let next = unsafe { System.realloc(ptr, layout, new_size) };
        if !next.is_null() {
            if new_size >= layout.size() {
                record_allocation(new_size - layout.size());
            } else {
                LIVE_BYTES.fetch_sub(layout.size() - new_size, Ordering::SeqCst);
            }
        }
        next
    }
}

#[global_allocator]
static ALLOCATOR: TrackingAllocator = TrackingAllocator;

#[test]
fn range_format_staging_holds_only_style_indices() {
    const CELLS: usize = 1_000;
    const TEXT_BYTES: usize = 8 * 1_024;

    let mut sheet = Sheet::new("Data");
    for row in 0..CELLS {
        sheet.set_cell(
            CellRef::new(row as u32, 0),
            Cell {
                value: CellValue::Text {
                    value: "x".repeat(TEXT_BYTES),
                },
                ..Cell::default()
            },
        );
    }
    let mut workbook = Workbook::default();
    workbook.sheets.push(sheet);
    let op = Op::PatchRangeStyle {
        sheet: SheetId(0),
        range: CellRange::new(CellRef::new(0, 0), CellRef::new((CELLS - 1) as u32, 0)),
        patch: StylePatch {
            italic: Some(true),
            ..StylePatch::default()
        },
    };

    let baseline = LIVE_BYTES.load(Ordering::SeqCst);
    PEAK_BYTES.store(baseline, Ordering::SeqCst);
    let inverse = apply(&mut workbook, &op).unwrap();
    let added_peak = PEAK_BYTES.load(Ordering::SeqCst) - baseline;
    black_box(inverse);

    assert!(
        added_peak < CELLS * TEXT_BYTES * 3 / 2,
        "range staging retained {added_peak} bytes above baseline"
    );
}
