use std::alloc::{GlobalAlloc, Layout, System};
use std::cell::Cell;

struct TrackedAllocator;

thread_local! {
    static ALLOCATED: Cell<Option<usize>> = const { Cell::new(None) };
}

unsafe impl GlobalAlloc for TrackedAllocator {
    unsafe fn alloc(&self, layout: Layout) -> *mut u8 {
        let _ = ALLOCATED.try_with(|count| {
            if let Some(bytes) = count.get() {
                count.set(Some(bytes + layout.size()));
            }
        });
        unsafe { System.alloc(layout) }
    }

    unsafe fn dealloc(&self, ptr: *mut u8, layout: Layout) {
        unsafe { System.dealloc(ptr, layout) }
    }
}

#[global_allocator]
static ALLOCATOR: TrackedAllocator = TrackedAllocator;

fn woff1(tables: &[(u32, u32, u32)]) -> Vec<u8> {
    let mut bytes = vec![0; 44 + 20 * tables.len() + tables.len()];
    bytes[..4].copy_from_slice(b"wOFF");
    bytes[4..8].copy_from_slice(&0x10000_u32.to_be_bytes());
    let length = bytes.len() as u32;
    bytes[8..12].copy_from_slice(&length.to_be_bytes());
    bytes[12..14].copy_from_slice(&(tables.len() as u16).to_be_bytes());
    bytes[16..20].copy_from_slice(&28_u32.to_be_bytes());
    for (index, (offset, compressed, original)) in tables.iter().enumerate() {
        let start = 44 + 20 * index;
        bytes[start..start + 4].copy_from_slice(&(index as u32).to_be_bytes());
        bytes[start + 4..start + 8].copy_from_slice(&offset.to_be_bytes());
        bytes[start + 8..start + 12].copy_from_slice(&compressed.to_be_bytes());
        bytes[start + 12..start + 16].copy_from_slice(&original.to_be_bytes());
    }
    bytes
}

fn reject_without_allocating(bytes: &[u8], expected: docx_edit::ResidentFontError) {
    ALLOCATED.set(Some(0));
    let result = docx_edit::decode_resident_font(bytes);
    let allocated = ALLOCATED.replace(None).unwrap();
    assert!(allocated < 4096, "allocated {allocated} bytes");
    if expected == docx_edit::ResidentFontError::SizeLimit {
        assert_eq!(allocated, 0);
    }
    assert_eq!(result.unwrap_err(), expected);
}

#[test]
fn oversized_table_preserves_memory_and_editing_session() {
    let doc = docx_edit::EditingDoc::new(51);
    doc.create_story("body", "Text", "Normal", "left").unwrap();
    let before = doc.encode_state_as_update_v1();
    let bytes = woff1(&[(64, 1, 65 * 1024 * 1024)]);
    assert_eq!(bytes.len(), 65);
    reject_without_allocating(&bytes, docx_edit::ResidentFontError::SizeLimit);
    assert_eq!(doc.encode_state_as_update_v1(), before);
    assert!(
        doc.set_section_margins(
            &docx_edit::EditCtx::local("", ""),
            "final",
            r#"{"marginLeft":2880}"#
        )
        .unwrap()
    );
    assert_eq!(doc.paragraphs("body").unwrap()[0].text, "Text");
}

#[test]
fn table_sum_is_checked_before_decompression() {
    reject_without_allocating(
        &woff1(&[(84, 1, 33 * 1024 * 1024), (85, 1, 33 * 1024 * 1024)]),
        docx_edit::ResidentFontError::SizeLimit,
    );
}

#[test]
fn overlapping_tables_are_rejected_before_decompression() {
    reject_without_allocating(
        &woff1(&[(84, 1, 1), (84, 1, 1)]),
        docx_edit::ResidentFontError::OverlappingTables,
    );
}

#[test]
fn table_bounds_are_checked_before_decompression() {
    for bytes in [
        woff1(&[(64, 2, 2)]),
        woff1(&[(44, 1, 1)]),
        woff1(&[(64, 2, 1)]),
    ] {
        reject_without_allocating(&bytes, docx_edit::ResidentFontError::TableBounds);
    }
    reject_without_allocating(
        &woff1(&[(u32::MAX, 1, 1)]),
        docx_edit::ResidentFontError::SizeLimit,
    );
}

fn base128(mut value: u32) -> Vec<u8> {
    let mut bytes = vec![(value & 127) as u8];
    value >>= 7;
    while value > 0 {
        bytes.push((value & 127) as u8 | 128);
        value >>= 7;
    }
    bytes.reverse();
    bytes
}

fn woff2(tables: &[(u8, u32, Option<u32>)]) -> Vec<u8> {
    let mut bytes = vec![0; 48];
    bytes[..4].copy_from_slice(b"wOF2");
    bytes[4..8].copy_from_slice(&0x10000_u32.to_be_bytes());
    bytes[12..14].copy_from_slice(&(tables.len() as u16).to_be_bytes());
    bytes[16..20].copy_from_slice(&28_u32.to_be_bytes());
    bytes[20..24].copy_from_slice(&1_u32.to_be_bytes());
    for (flags, original, transformed) in tables {
        bytes.push(*flags);
        bytes.extend(base128(*original));
        if let Some(length) = transformed {
            bytes.extend(base128(*length));
        }
    }
    bytes.push(0);
    let length = bytes.len() as u32;
    bytes[8..12].copy_from_slice(&length.to_be_bytes());
    bytes
}

#[test]
fn woff2_table_and_transform_sizes_are_checked_before_decompression() {
    for tables in [
        vec![(0, 65 * 1024 * 1024, None)],
        vec![(0, 33 * 1024 * 1024, None), (1, 33 * 1024 * 1024, None)],
        vec![(10, 12, Some(65 * 1024 * 1024))],
        vec![
            (10, 12, Some(33 * 1024 * 1024)),
            (67, 12, Some(33 * 1024 * 1024)),
        ],
    ] {
        reject_without_allocating(&woff2(&tables), docx_edit::ResidentFontError::SizeLimit);
    }
    let mut bytes = woff2(&[(0, 12, None)]);
    bytes[20..24].copy_from_slice(&100_u32.to_be_bytes());
    reject_without_allocating(&bytes, docx_edit::ResidentFontError::TableBounds);
}
