# End-to-end scenarios

Pinned corpus documents (`corpus.ts`: id, byte count and sha256 from
`corpus.betteroffice.dev`) run through editing scenarios per format, and every
operation is timed on its own: the latency across the boundary it crosses (a
wasm call, a Python round trip, an update exchange) plus the stage breakdown
the engine reports for it. Correctness is asserted after every step, so a run
is an end-to-end test and a per-operation performance profile at once.

## Running

```bash
bun run test:e2e            # build the wasm bundles, run, print the per-op summary
bun run test:e2e:record     # ... and write scripts/e2e/results/<format>.json
bun run test:e2e:compare    # ... and fail when a median regresses 25% and 10 ms+
bun scripts/e2e/python-env.ts   # once: venv with the Python bindings for the cross-SDK scenarios
bun scripts/e2e/report.ts [dir] [--json] | --diff <before> <after>
```

`BETTEROFFICE_E2E` (`1` | `record` | `compare`) is what gates the suites; unset,
`bun test scripts` skips them. Corpus bytes cache under
`~/.cache/betteroffice/e2e-assets` (`QUALITY_ASSET_CACHE` overrides), the Python
venv under `~/.cache/betteroffice/e2e-venv` (`BETTEROFFICE_E2E_VENV`). Without the
venv the `python` scenarios are reported as skipped with the reason.
`BETTEROFFICE_E2E_OUTPUT=<dir>` exports the run as JSON and, on GitHub Actions,
the tables land in the step summary. Recorded baselines are machine-specific:
record on a quiet machine, compare on the same one.

## Layout

```
harness.ts        ScenarioRecorder, per-op stats, record/compare, summaries
suite.ts          defineSuite(format, scenarios, { setup, context })
corpus.ts         pinned samples and loadSample
python.ts         PythonWorker: a long-lived interpreter, one op per call
python-env.ts     venv bootstrap (maturin develop --release for the bindings)
python/worker.py  the interpreter loop
<format>/context.ts          wasm init, handle lifetime, stage mappers, helpers
<format>/<scenario>.ts       one Scenario per file
<format>/index.ts            the ordered scenario list
<format>.e2e.test.ts         defineSuite wiring
results/<format>.json        recorded baseline (schemaVersion 2)
```

A scenario is `{ name, description, participants, samples?, requires?, run(ctx) }`.
Inside `run`, every engine call goes through `recorder.op(name, fn, stages?)`
(or `recorder.as('web:a').op(...)` for multi-editor runs, `PythonWorker.call`
for Python); `recorder.load(fn)` opens the document. Each scenario runs against
every pinned sample of its format.

## Scenarios

| xlsx | docx | pptx |
| --- | --- | --- |
| editing-session | editing-session | editing-session |
| formula-chain-cascade | typing-burst | typing-burst |
| bulk-paste-and-formats | pagination-pressure | deck-build-from-scratch |
| structural-storm | tables-deep | slide-reorder-and-delete |
| typing-latency | formatting-and-styles | formatting-sweep |
| viewport-scroll-render | search-and-replace-sweep | shapes-and-pictures |
| two-editors-converge | two-editors-converge | comments-thread |
| three-editors-mesh-with-undo | three-editors-suggesting | two-editors-converge |
| proposals-review | comments-and-anchors | three-editors-mesh-with-undo |
| python-roundtrip | python-roundtrip | proposals-review |
| python-web-live-collab | python-layout-parity | python-roundtrip |
| save-load-cycles | save-load-cycles | python-web-live-collab |
| | | layout-all-slides |

Multi-editor scenarios open two or three replicas of the same document, stamp
every operation with the replica that issued it, and exchange Yrs updates until
the replicas' fingerprints match. Cross-SDK scenarios drive the Python bindings
from the same run: for xlsx and pptx a web replica and a Python replica trade
updates live; for docx the web engine measures and hands its retained kernel to
the Python paginator, which must report the same page count.

## Results schema

```jsonc
{
  "schemaVersion": 2,
  "commit": "…", "recordedAt": "ISO-8601",
  "environment": { "platform", "arch", "cpu", "cpus", "bun" },
  "scenarios": [{
    "format": "xlsx", "scenario": "two-editors-converge", "sample": "betteroffice-workbook",
    "description": "…", "participants": ["web:a", "web:b"],
    "status": "passed" | "skipped", "reason?": "…",
    "loadMs": 57.4,
    "ops": [{ "op": "applyUpdate", "actor": "web:b", "e2eMs": 1.2,
              "internal?": { "validate": 0.0, "apply": 0.9 }, "detail?": { "bytes": 412 } }],
    "summary": {
      "opCount": 19, "totalMs": 165.3,
      "byOp": { "editCell:number": { "count": 2, "totalMs", "meanMs", "p50Ms", "p95Ms", "maxMs", "stagesMs?": {…} } },
      "byActor": { "web:a": 80.1, "web:b": 85.2 }
    }
  }]
}
```

`ops` is the raw sequence; `summary.byOp` is what `compare` and `report.ts`
read. Stage names are the engine's own boundaries (xlsx: validate, apply,
recalc, result / build, encode; docx: selection, edit, lower, measure,
paginate, displayInput, displayBuild, displayFinalize, encode; pptx: scope,
layout, serialize / parse, apply, serialize / undo, snapshot, serialize;
Python: whatever the script wraps in `with timed('…')`).
