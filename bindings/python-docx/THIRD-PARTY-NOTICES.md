# Third-party notices

This wheel embeds no third-party fonts or other assets. The extension module
statically links the Rust crates this package depends on, each under its own
permissive license (Apache-2.0 or MIT); their notices travel with those crates
on crates.io. The package's own code is Apache-2.0.

Document text is measured and rasterized with fonts the caller registers
through `register_font`, so no face is compiled into the wheel.

The extension also links the Zstandard decoder (`facebook/zstd` v1.5.7, decompression
only) through the `betteroffice-fonts` crate to decode compressed font packs.
Copyright (c) Meta Platforms, Inc. and affiliates. It is distributed under the
BSD-3-Clause license in `licenses/BSD-3-Clause-Zstandard.txt`.

It also links `wuff` 0.2.9 for WOFF2 reconstruction of woff2t font packs.
Copyright (c) 2013-2017 by the WOFF2 Authors. Copyright (c) 2025 by Nico Burns.
It is distributed under the MIT license in `licenses/MIT-wuff.txt`.
