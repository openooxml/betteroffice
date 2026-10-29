# Font pack format 1

The internal pack builder writes one `.ttf.zst` or `.otf.zst` file per face.
Each file contains exactly one Zstandard frame, encoded at level 19 by default,
with the decoded content size and the content checksum enabled. The checksum is
the low 32 bits of XXH64 of the decoded bytes. Concatenated frames, skippable
frames and trailing data are not part of this format.

## Face codecs

The per-face `codec` is `zstd` or `woff2t+zstd`; the manifest format stays 1.
For `zstd`, the frame payload is the source sfnt. For `woff2t+zstd`, it is a
WOFF2 file with transformed `glyf` and `loca` tables. Its table stream is stored
uncompressed: `totalCompressedSize` equals the raw stream length and the brotli
step is the identity. This codec supports TTF glyf/loca faces only.

Development and release tooling uses pinned fontTools to write both this
identity-stream `.uw2` and a standard brotli `.woff2` companion; see
[`scripts/fonts/README.md`](../../scripts/fonts/README.md). `fontpack build
--codec woff2t` compresses the `.uw2`, verifies every reconstructed face against
wuff's decode of the standard `.woff2`, and checks `maxp.numGlyphs` against the
source TTF. For `woff2t`, each `<face>.source.sha256` sidecar lists the SHA-256
of the `.ttf`, `.uw2` and `.woff2` in `sha256sum` format, one line each in that
order; the builder rejects a face whose sidecar is missing or does not match all
three current files. Regenerate companions with `scripts/fonts/woff2-transform.py`.
The default `--codec zstd` preserves source bytes. File names stay
`<face>.ttf.zst` for both codecs.

WOFF2 reconstruction produces a valid sfnt but does not preserve the source TTF
bytes: `glyf` is re-encoded and `DSIG` is dropped, as with standard WOFF2.
`byteLength` and `sha256` describe the reconstructed sfnt for `woff2t+zstd`.

## Dictionaries

A family with two or more faces uses the zstd-decoded payload of its
`<family>-Regular.ttf.zst` or `<family>-Regular.otf.zst` as a raw-content
dictionary for its other faces. The
Regular face itself has no dictionary. A multi-face family must have exactly
one Regular face. A single-face family uses no dictionary. Each CJK face is a
single-face family and has no dictionary. There are no trained or cross-family
dictionaries.

`dict` names the prerequisite's encoded file within the same pack and version.
Fetch the face and its prerequisite together, then decode Regular first and pass
its payload to the sibling decoder. For `zstd` this is the sfnt; for
`woff2t+zstd` it is the transformed WOFF2 payload, never the reconstructed sfnt.
A face and its dictionary must share the codec; the builder rejects mismatches.
Never use the compressed prerequisite as the dictionary.
Every prerequisite has `dict: null`, so dependencies have
at most one level. The compression window covers dictionary plus source bytes.

Raw-content dictionaries have no dictionary ID. A missing or wrong dictionary
can report a decompression or checksum error. The checksum validates the
zstd payload; the manifest's SHA-256 additionally identifies the exact final
font bytes.

## Decoding and ownership

The published `betteroffice-fonts` Rust crate owns decoding. Every engine links
it and exposes `decodeFontFace(bytes, codec, dict?, maxOutput?)` in wasm and
Node bindings. Python exposes `decode_font_face`; Rust facades re-export the
crate as `fonts`. Engine calls return owned font bytes and the optional woff2t
payload. The standalone wasm decoder links the same crate. Its vendored C
decoder lives at `crates/betteroffice-fonts/vendor/zstd`.

Decoders sniff only `28 B5 2F FD`. Other input, including TTF (`00 01 00 00`),
OTF (`OTTO`), WOFF (`wOFF`) and WOFF2 (`wOF2`), passes through without decoding.
Decoding is a single operation, with no recursive attempt on the result.

For plain zstd, the internal TypeScript decoder uses the built-in Node zstd
decompressor when available, otherwise the standalone wasm decoder. An explicit
`wasmUrl` selects wasm; without it the decoder loads `dist/decoder/fontpack.wasm`
from the package root. `woff2t+zstd` always uses the full wasm pipeline,
including in Node. Both codecs share one module, which links the
decompression-only zstd C source and wuff 0.2.9 with default features disabled.
Wuff reconstructs the font with `decompress_woff2_with_custom_brotli` and an
identity callback; brotli and flate2 are absent from this wasm feature set. The
builder alone enables brotli for reference verification.

Each manifest face must have `codec: "zstd"` or `codec: "woff2t+zstd"`.
Callers pass that value as the
internal decoder's `codec` option. Any other value fails with an unsupported
font face codec error before decoding or passthrough. Omitting the option for
standalone byte decoding selects `zstd`.

The frame must declare its content size before output allocation. The default
TypeScript `maxOutput` and Rust `decode_font_face` limit are 64 MiB by default.
Explicit limits must be integers from 0 to 2,147,483,647. Oversized
frames fail before output allocation. Decoders reject missing checksums,
corruption, truncation and output lengths that disagree with the header.
For `woff2t+zstd`, `maxOutput` bounds both the zstd payload and the reconstructed
font, and is never more than 64 MiB whatever the caller passes: reconstruction
briefly holds the payload, the table stream and growing glyph buffers at once, and
that cap keeps the peak well inside the decoder's 1 GiB memory. The same 64 MiB
limit applies to the encoded frame and to the dictionary of a `woff2t+zstd` face,
checked before either is copied into the decoder. Before wuff runs, an allocation-free preflight validates the WOFF2 header,
directory and all transformed glyph streams. There may be at most 256 tables;
stored lengths must fit the raw table stream, and no original table length may
exceed eight times the payload length. The declared sfnt size and the sum of
padded original table lengths plus the sfnt directory must fit `maxOutput`.
Glyph counts use the format's 16-bit field. A glyph has at most 65,536 points;
the total point count must fit the original `glyf` length, and point and glyph
scratch sizes must fit the output budget. The preflight also calculates the
reconstructed glyph and font sizes before reconstruction. These validation
failures report a WOFF2 reconstruction error (`-10`).
Reconstruction errors are reported as errors, without a passthrough retry.
Passthrough does not apply the decoded-output limit.

TypeScript always returns a fresh, independently owned `ArrayBuffer`, including
for passthrough. A caller may transfer that buffer. The wasm instance is reused,
but its output is copied and freed after each call. Rust passthrough returns a
borrowed `Cow` through `decode_or_passthrough`; `decode_font_face` always returns
owned bytes.
If any wasm export throws, TypeScript discards that instance and reports
`Font face could not be decoded`. It makes no further calls into that instance;
the next decode loads a fresh module instance.

The internal `decodeFontFace` returns `{ font, payload? }`; `decodeFontBytes`
returns only the font. `payload` is present for decoded woff2t faces and has its
own independently owned buffer. The future manager can retain it as a sibling
dictionary without decoding twice. Rust `decode_face` returns `Decoded` with
the same ownership. Existing Rust `decode` and `decode_or_passthrough` keep
their plain zstd behavior.

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
bytes; `byteLength` is the final sfnt size. `sha256` is the lowercase hexadecimal
SHA-256 of the final sfnt, including reconstruction for `woff2t+zstd`.
The builder decodes and compares every face before writing its manifest. It
refuses an existing output directory, writes the whole pack into a staging
directory beside it and renames that into place, so a failed build never leaves
a partial or mixed pack. Display family names, weights, styles, scripts and
metric compatibility remain in the TypeScript manifest and are not fields in
schema 1.

## Raw wasm ABI

The wasm module exports `memory` and these functions, without wasm-bindgen:

| Export | Contract |
| --- | --- |
| `fp_alloc(len)` | Allocate input bytes, returning a pointer or zero on allocation failure. |
| `fp_free(ptr, len)` | Free one input allocation; use its original pointer and length exactly once. |
| `fp_decode(src_ptr, src_len, dict_ptr, dict_len, max_output)` | Decode one frame; zero dictionary length means no dictionary. Return zero on success. |
| `fp_decode_face(src_ptr, src_len, dict_ptr, dict_len, max_output, codec)` | Decode and reconstruct a face. Codec 0 is `zstd`, 1 is `woff2t+zstd`; any other value returns `-6`. |
| `fp_out_ptr()` | Borrow the output pointer until the next decode or output free. |
| `fp_out_len()` | Read the decoded length. |
| `fp_out_free()` | Release output, including after failure. |
| `fp_payload_ptr()` | Borrow the woff2t payload pointer until the next decode or payload free. Zero for plain zstd. |
| `fp_payload_len()` | Read the woff2t payload length. Zero for plain zstd. |
| `fp_payload_free()` | Release the separate payload, including after failure. |

The caller owns input allocations and must keep them alive during decoding.
Inputs must come from `fp_alloc`; outputs must be released with `fp_out_free`
and `fp_payload_free`. Each decode clears both previous outputs, including on
failure. Output pointers must not be passed as inputs; copy a prerequisite's
payload to a new `fp_alloc` allocation before using it as a dictionary.
Do not overlap a decode with another call on the same instance. Memory grows
on demand; reacquire views of `memory.buffer` after any allocating operation.
The standalone module requires one thread and unshared memory. Atomics-enabled
builds are rejected.
Linear memory is capped at 1 GiB and individual allocator requests are bounded.

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
| `-10` | WOFF2 reconstruction failure. |
