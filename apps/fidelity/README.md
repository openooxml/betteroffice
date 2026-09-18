# Visual fidelity viewer

Browse the [visual fidelity](../../scripts/office-quality/README.md) results: pick a document, then
compare Microsoft Office's page against ours by swiping a divider across them or by switching to a
difference blend. Scores come from `report.json` as measured; the viewer never computes its own.

```sh
bun run dev:fidelity
```

Office reference pages are served from the public corpus, so scores and page navigation work as soon
as a report loads. Our renders come from the `RENDERS` R2 bucket the fidelity workflow uploads to;
pages without a published render show the Office page alone and say so. Without the bucket, open a
`report.json` from a workflow artifact directly.
