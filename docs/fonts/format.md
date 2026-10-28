# Font pack format 1

The internal pack builder writes one `.ttf.zst` or `.otf.zst` file per face.
Each file contains exactly one Zstandard frame, encoded at level 19 by default,
with the decoded content size and the content checksum enabled. The checksum is
the low 32 bits of XXH64 of the decoded bytes. Concatenated frames, skippable
frames and trailing data are not part of this format.

## Dictionaries

A family with two or more faces uses the decoded `<family>-Regular.ttf` or
`<family>-Regular.otf` as a raw-content dictionary for its other faces. The
Regular face itself has no dictionary. A multi-face family must have exactly
one Regular face. A single-face family uses no dictionary. Each CJK face is a
single-face family and has no dictionary. There are no trained or cross-family
dictionaries.

`dict` names the prerequisite's encoded file within the same pack and version.
Fetch the face and its prerequisite together, then decode Regular first and pass
its decoded bytes to the sibling decoder. Never use the compressed prerequisite
as the dictionary. Every prerequisite has `dict: null`, so dependencies have
at most one level. The compression window covers dictionary plus source bytes.

Raw-content dictionaries have no dictionary ID. A missing or wrong dictionary
can report a decompression or checksum error. The checksum validates the
decoded content; the manifest's SHA-256 additionally identifies the exact bytes.

## Decoding and ownership

Decoders sniff only `28 B5 2F FD`. Other input, including TTF (`00 01 00 00`),
OTF (`OTTO`), WOFF (`wOFF`) and WOFF2 (`wOF2`), passes through without decoding.
Decoding is a single operation, with no recursive attempt on the result.

The internal TypeScript decoder uses the built-in Node zstd decompressor when
available, otherwise the standalone wasm decoder. An explicit `wasmUrl` selects
the wasm decoder. The Rust decoder links the same decompression-only C source
on native and wasm targets.

Each manifest face must have `codec: "zstd"`. Callers pass that value as the
internal decoder's `codec` option. Any other value fails with an unsupported
font face codec error before decoding or passthrough. Omitting the option for
standalone byte decoding selects `zstd`. No other codec, including
`woff2t+zstd`, is currently implemented.

The frame must declare its content size before output allocation. The default
TypeScript `maxOutput` is 64 MiB; Rust callers supply their own limit. Oversized
frames fail before output allocation. Decoders reject missing checksums,
corruption, truncation and output lengths that disagree with the header.
Passthrough does not apply the decoded-output limit.

TypeScript always returns a fresh, independently owned `ArrayBuffer`, including
for passthrough. A caller may transfer that buffer. The wasm instance is reused,
but its output is copied and freed after each call. Rust passthrough returns a
borrowed `Cow`; decoded output is owned.

## Manifest schema

`manifest.json` is UTF-8 JSON. `faces` is sorted by `file`. File names are
relative to the pack directory. `family` is the source filename stem before its
last `-`, or the entire stem when there is no `-`.

```json
{
  "format": 1,
  "pack": "synthetic",
  "version": "1.2.3",
  "faces": [
    {
      "file": "A-Bold.ttf.zst",
      "source": "A-Bold.ttf",
      "family": "A",
      "codec": "zstd",
      "encodedLength": 123,
      "byteLength": 456,
      "sha256": "0123456789abcdef0123456789abcdef0123456789abcdef0123456789abcdef",
      "dict": "A-Regular.ttf.zst"
    },
    {
      "file": "A-Regular.ttf.zst",
      "source": "A-Regular.ttf",
      "family": "A",
      "codec": "zstd",
      "encodedLength": 234,
      "byteLength": 567,
      "sha256": "abcdef0123456789abcdef0123456789abcdef0123456789abcdef0123456789",
      "dict": null
    }
  ]
}
```

`version` is a semantic version. `encodedLength` is the complete frame size in
bytes; `byteLength` is the decoded size. `sha256` is the lowercase hexadecimal
SHA-256 of decoded bytes. The builder decodes and compares every face before
writing its manifest. Display family names, weights, styles, scripts and metric
compatibility remain in the TypeScript manifest and are not fields in schema 1.

## Raw wasm ABI

The wasm module exports `memory` and these functions, without wasm-bindgen:

| Export | Contract |
| --- | --- |
| `fp_alloc(len)` | Allocate input bytes, returning a pointer or zero on allocation failure. |
| `fp_free(ptr, len)` | Free one input allocation; use its original pointer and length exactly once. |
| `fp_decode(src_ptr, src_len, dict_ptr, dict_len, max_output)` | Decode one frame; zero dictionary length means no dictionary. Return zero on success. |
| `fp_out_ptr()` | Borrow the output pointer until the next decode or output free. |
| `fp_out_len()` | Read the decoded length. |
| `fp_out_free()` | Release output, including after failure. |

The caller owns input allocations and must keep them alive during decoding.
Inputs must come from `fp_alloc`; outputs must be released with `fp_out_free`.
Do not overlap a decode with another call on the same instance. Memory grows
on demand; reacquire views of `memory.buffer` after any allocating operation.
The standalone module requires one thread and unshared memory. Atomics-enabled
builds are rejected.

Negative return codes:

| Code | Meaning |
| --- | --- |
| `-1` | Not a zstd frame. |
| `-2` | Missing content size. |
| `-3` | Decoded content exceeds the limit. |
| `-4` | Content checksum mismatch, including a wrong raw dictionary. |
| `-5` | Dictionary error. |
| `-6` | Corrupt, truncated or multiple frames. |
| `-7` | Allocation failure. |
| `-8` | Invalid input memory range. |
| `-9` | Missing content checksum. |
