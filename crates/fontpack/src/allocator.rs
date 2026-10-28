use alloc::alloc::{Layout, alloc, dealloc};
use core::ptr;

#[cfg(feature = "wasm")]
use lol_alloc::{AssumeSingleThreaded, FreeListAllocator};

// The standalone wasm has no shared memory or callbacks; allocator calls cannot overlap.
#[cfg(feature = "wasm")]
#[global_allocator]
static ALLOCATOR: AssumeSingleThreaded<FreeListAllocator> =
    unsafe { AssumeSingleThreaded::new(FreeListAllocator::new()) };

const ALIGN: usize = 16;
const HEADER: usize = ALIGN;

fn layout(size: usize) -> Option<Layout> {
    Layout::from_size_align(size.max(1).checked_add(HEADER)?, ALIGN).ok()
}

#[unsafe(no_mangle)]
pub extern "C" fn malloc(size: usize) -> *mut u8 {
    let Some(layout) = layout(size) else {
        return ptr::null_mut();
    };
    let base = unsafe { alloc(layout) };
    if base.is_null() {
        return base;
    }
    unsafe {
        base.cast::<usize>().write(size);
        base.add(HEADER)
    }
}

#[unsafe(no_mangle)]
pub extern "C" fn calloc(count: usize, size: usize) -> *mut u8 {
    let Some(len) = count.checked_mul(size) else {
        return ptr::null_mut();
    };
    let bytes = malloc(len);
    if !bytes.is_null() {
        unsafe { ptr::write_bytes(bytes, 0, len) };
    }
    bytes
}

#[unsafe(no_mangle)]
/// # Safety
/// `bytes` must be null or a live allocation from this module.
pub unsafe extern "C" fn free(bytes: *mut u8) {
    if bytes.is_null() {
        return;
    }
    unsafe {
        let base = bytes.sub(HEADER);
        if let Some(layout) = layout(base.cast::<usize>().read()) {
            dealloc(base, layout);
        }
    }
}
