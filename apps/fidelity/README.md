# Benchmarks site

The benchmarks site charts the latest [Benchmarks](../../scripts/office-quality/README.md) run
and the latest green [end-to-end](../../e2e/README.md) run on `main`.

```sh
bun run dev:fidelity
```

Styling is Tailwind CSS v4, as in `apps/web`: `src/styles.css` holds the theme tokens and the chart
styles, and `bun run build` compiles it into `dist/` next to the bundled scripts.

- `/` is the overview: headline numbers, the README tables drawn to scale, per-document SSIM,
  DOCX page agreement, render and recalculation times, formula accuracy, and SDK call latency.
- `/compare` is the page viewer: pick a document, then compare Microsoft Office's page against ours
  by swiping a divider across them or by switching to a difference blend. `?doc=` and `?page=`
  open a document directly.

Both pages read `report.json` as measured. The overview aggregates it with the same rules as the
generated README tables, and a test pins the two together. Office reference pages come from the
public corpus. Our renders, the report and the end-to-end results come from the `RENDERS` R2 bucket
under `renders/` and `e2e/`. Without the bucket, serve workflow artifacts from the same origin and
pass their paths as `?report=` (and `?e2e=` for a directory of `docx.json`, `pptx.json` and
`xlsx.json`), or open a `report.json` in the page viewer's file picker. Overrides that point at
another origin are ignored, so a shared link always shows the published numbers.

The old `fidelity` host redirects permanently to this site; its root lands on `/compare`.
