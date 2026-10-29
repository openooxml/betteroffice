#![allow(non_snake_case)]

use alloc::alloc::{Layout, alloc, dealloc};
use core::ptr;

#[cfg(feature = "wasm")]
use core::alloc::GlobalAlloc;
#[cfg(feature = "wasm")]
use lol_alloc::{AssumeSingleThreaded, FreeListAllocator};

#[cfg(feature = "wasm")]
const MAX_ALLOCATION: usize = 1 << 30;

#[cfg(feature = "wasm")]
struct Bounded(AssumeSingleThreaded<FreeListAllocator>);

#[cfg(feature = "wasm")]
unsafe impl GlobalAlloc for Bounded {
    unsafe fn alloc(&self, layout: Layout) -> *mut u8 {
        if layout.size() > MAX_ALLOCATION {
            return ptr::null_mut();
        }
        unsafe { self.0.alloc(layout) }
    }

    unsafe fn dealloc(&self, bytes: *mut u8, layout: Layout) {
        unsafe { self.0.dealloc(bytes, layout) }
    }
}

// The standalone wasm has no shared memory or callbacks; allocator calls cannot overlap.
#[cfg(feature = "wasm")]
#[global_allocator]
static ALLOCATOR: Bounded = Bounded(unsafe { AssumeSingleThreaded::new(FreeListAllocator::new()) });

const ALIGN: usize = 16;
const HEADER: usize = ALIGN;

fn layout(size: usize) -> Option<Layout> {
    Layout::from_size_align(size.max(1).checked_add(HEADER)?, ALIGN).ok()
}

#[unsafe(no_mangle)]
pub extern "C" fn FONTPACK_malloc(size: usize) -> *mut u8 {
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
pub extern "C" fn FONTPACK_calloc(count: usize, size: usize) -> *mut u8 {
    let Some(len) = count.checked_mul(size) else {
        return ptr::null_mut();
    };
    let bytes = FONTPACK_malloc(len);
    if !bytes.is_null() {
        unsafe { ptr::write_bytes(bytes, 0, len) };
    }
    bytes
}

#[unsafe(no_mangle)]
/// # Safety
/// `bytes` must be null or a live allocation from this module.
pub unsafe extern "C" fn FONTPACK_free(bytes: *mut u8) {
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
