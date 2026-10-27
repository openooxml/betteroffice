# Local BetterOffice DOCX capture

Build and start the local viewer:

```sh
bun install --frozen-lockfile
bun run build:fonts
bun run build:docx
bunx playwright install chromium
QUALITY_PORT=4178 node scripts/docx-quality/server.mjs
```

In another terminal, capture a file you choose:

```sh
QUALITY_ENGINE_LABEL=my-revision node scripts/docx-quality/browser-task.mjs \
  /path/to/example.docx .source/office-quality/example/betteroffice \
  cdn http://127.0.0.1:4178
```

The output directory must be empty. Captures use 150 DPI and export all pages; the viewer waits up to 120 seconds plus 2 seconds per reference page (at most 480) for layout, a canvas that is still blank at the HTML default size (300 x 150, not painted yet) is not accepted until it paints or the grace below runs out, and a page whose canvas size differs from the reference page gets 15 more seconds to settle before the capture fails; `result.json` records the source hash, page count, font loads, and external requests. The deadline defaults to 600 seconds (`QUALITY_TIMEOUT_SECONDS`). Use `none` instead of `cdn` to check fallback-font completion. All document bytes stay local. The viewer disables automatic Google Fonts lookups and permits external requests only for pinned Latin/CJK font binaries, without bodies, cookies, or referrers.

For release comparisons, extract the desired DOCX and DOCX React npm tarballs into `.source/`, then set `QUALITY_PACKAGE_ROOT` and `QUALITY_REACT_ROOT` when starting a server on a different port. Label captures with the actual release version. Restart servers after builds; automatic reloads are disabled and each port uses its own cache.

Use the [Office reference and comparison scripts](../office-quality/README.md) to export from Word and inspect the local image diff.
