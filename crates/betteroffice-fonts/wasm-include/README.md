# Freestanding decoder headers

These headers supply only the types, limits and declarations needed by the
zstd 1.5.7 decoder on `wasm32-unknown-unknown`. They use compiler type macros
and do not require a libc sysroot or a system `stddef.h`.

Rust supplies `FONTPACK_malloc`, `FONTPACK_calloc` and `FONTPACK_free` through
the global allocator. Wasm-only C defines rename the declarations and calls
from `malloc`, `calloc` and `free` to these decoder-specific symbols.
Each allocation has a 16-byte size header and 16-byte alignment. The decoder
does not reference `realloc`. Memory operations resolve to Rust's
compiler-builtins. Assertions are disabled, matching the amalgamation's
`DEBUGLEVEL=0`. Native builds use the system headers and libc.

The standalone wasm feature uses `no_std`, `alloc` and the single-threaded
`lol_alloc` freelist allocator. Freed allocations are reused and adjacent free
regions are coalesced. The module has unshared memory and no callbacks.
