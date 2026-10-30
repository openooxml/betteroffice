# Zstandard decoder source

Upstream: `facebook/zstd`, version `v1.5.7`.
Release tarball: https://github.com/facebook/zstd/releases/download/v1.5.7/zstd-1.5.7.tar.gz

Tarball SHA-256: `eb33e51f49a15e023950cd7825ca74a4a2b43db8354825ac24fc1b7ee09e6fa3`
(matches the upstream `zstd-1.5.7.tar.gz.sha256` release asset).

Run `bash scripts/fonts/vendor-zstd.sh` to verify and extract the pinned release,
then run its `build/single_file_libs/create_single_file_decoder.sh` from that
directory. The script copies the generated, unmodified `zstddeclib.c` and the
upstream `LICENSE` here. The selected upstream licence is BSD-3-Clause.

The decoder build namespaces C identifiers in a temporary build artifact to
allow linking alongside the build tool's full zstd compressor.
