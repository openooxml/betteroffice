use alloc::vec::Vec;
use core::cell::UnsafeCell;
use core::ptr;
use core::slice;

use crate::{DecodeError, allocator, decode};

#[cfg(target_feature = "atomics")]
compile_error!("fontpack's standalone decoder requires wasm without atomics");

struct Output(UnsafeCell<Option<Vec<u8>>>);

// Exports run sequentially on one thread, with no shared memory, imports or callbacks.
unsafe impl Sync for Output {}

static OUTPUT: Output = Output(UnsafeCell::new(None));

#[cfg(all(not(feature = "build"), not(test)))]
#[panic_handler]
fn panic(_info: &core::panic::PanicInfo<'_>) -> ! {
    core::arch::wasm32::unreachable()
}

#[unsafe(no_mangle)]
pub extern "C" fn fp_alloc(len: usize) -> *mut u8 {
    allocator::malloc(len)
}

#[unsafe(no_mangle)]
/// # Safety
/// `bytes` must be null or a live `fp_alloc` allocation.
pub unsafe extern "C" fn fp_free(bytes: *mut u8, _len: usize) {
    unsafe { allocator::free(bytes) };
}

fn valid_range(bytes: *const u8, len: usize) -> bool {
    len <= isize::MAX as usize
        && (len == 0 || !bytes.is_null())
        && (bytes as usize)
            .checked_add(len)
            .is_some_and(|end| end <= core::arch::wasm32::memory_size(0) * 65536)
}

#[unsafe(no_mangle)]
/// # Safety
/// Inputs must be live `fp_alloc` allocations, separate from decoder output.
pub unsafe extern "C" fn fp_decode(
    src_ptr: *const u8,
    src_len: usize,
    dict_ptr: *const u8,
    dict_len: usize,
    max_output: usize,
) -> i32 {
    fp_out_free();
    if !valid_range(src_ptr, src_len) || !valid_range(dict_ptr, dict_len) {
        return -8;
    }
    let src = if src_len == 0 {
        &[]
    } else {
        unsafe { slice::from_raw_parts(src_ptr, src_len) }
    };
    let dict = if dict_len == 0 {
        None
    } else {
        Some(unsafe { slice::from_raw_parts(dict_ptr, dict_len) })
    };
    match decode(src, dict, max_output) {
        Ok(bytes) => {
            unsafe { *OUTPUT.0.get() = Some(bytes) };
            0
        }
        Err(error) => match error {
            DecodeError::NotZstd => -1,
            DecodeError::MissingContentSize => -2,
            DecodeError::TooLarge { .. } => -3,
            DecodeError::ChecksumMismatch => -4,
            DecodeError::WrongDictionary => -5,
            DecodeError::AllocationFailed => -7,
            DecodeError::MissingChecksum => -9,
            DecodeError::Corrupt | DecodeError::Truncated | DecodeError::TrailingData => -6,
        },
    }
}

#[unsafe(no_mangle)]
pub extern "C" fn fp_out_ptr() -> *const u8 {
    let output = unsafe { &*OUTPUT.0.get() };
    match output {
        Some(bytes) if !bytes.is_empty() => bytes.as_ptr(),
        _ => ptr::null(),
    }
}

#[unsafe(no_mangle)]
pub extern "C" fn fp_out_len() -> usize {
    let output = unsafe { &*OUTPUT.0.get() };
    output.as_ref().map_or(0, Vec::len)
}

#[unsafe(no_mangle)]
pub extern "C" fn fp_out_free() {
    unsafe { *OUTPUT.0.get() = None };
}
