# VSDX benchmarks and Visio parity log

Running state for the overnight loop. **Read this first each iteration; append to it
last.** Do not redo anything already recorded as done.

Operating rules are in `openspec/vsdx-overnight.md`.

## How to use this file

- One iteration does **one** thing, gates it, commits it, and updates this log.
- Record absolute numbers with the commit they were measured at, so a regression is
  attributable to a change.
- Log a parity gap **with evidence** before fixing it. Never implement from a hunch.

## Priority order

1. A failing test or known-broken behaviour listed under *Open* below.
2. A measured performance regression or hotspot.
3. The next unimplemented behaviour, highest corpus frequency first.
4. **If 1–3 are all empty, run a discovery pass (below). The queue never stays empty.**

---

## Discovery protocol

The queue is not a fixed list. When nothing is queued, **go find more work from
measured data** — never from a hunch, and never stop because the list ran out.

A discovery pass produces new *Open* entries, each with evidence and a frequency
count, ranked by how often the thing actually occurs in real diagrams. Discovery is
itself a valid iteration: finding and ranking ten real gaps is a good night's work
even if none are fixed yet.

### Instrumented sources — these already count things for you

| Source | What it tells you |
|---|---|
| `crates/vsdx-eval` corpus harness | Every formula bucketed: `evaluated`, `unsupported_known`, `unsupported_other`, `error`. The `unsupported_other` and `error` buckets are the richest mine — group by function or pattern and rank by count. |
| `GeometryIssue::UnsupportedRowType` (`vsdx-resolve/src/geometry.rs`) | Which geometry row types appear in the corpus and get dropped. Instrument a count per type. |
| `placeholder()` / `placeholder_at()` in `vsdx-render/src/lib.rs` | Which shapes fail to render and why. Count by reason across the corpus. |
| `@V` oracle disagreements | Where our evaluated value differs from Visio's cached value. Each disagreement is either our bug or a stale cache — investigate, don't assume. |
| Benchmark stage timings | Any stage growing superlinearly with a scaling knob. |
| Round-trip diffs | Any corpus file that does not re-serialize byte-identically when unmodified. |

**Add counters where they are missing.** If a code path silently swallows something
unsupported, making it countable is itself a worthwhile iteration — you cannot rank
what you cannot measure.

### Other sources

- **MS-VSDX spec** — walk sections not yet implemented; check each against corpus
  frequency before queueing, so effort follows real usage.
- **Real Visio behaviour** — where the spec is silent or ambiguous, record what Visio
  actually does and cite the file that shows it.
- **Property and fuzz testing** — generate valid documents, assert invariants (saves
  without error, reopens to what was rendered, round-trips byte-identically). Every
  failure is a new queue entry.
- **`TODO`/`FIXME` and `unsupported`/`unimplemented` strings** in the vsdx crates.

### Rules

- Rank by **measured corpus frequency**, not by what seems interesting.
- Record evidence before implementing: what Visio does, the proof, what we do instead.
- Respect the permanently excluded list. Discovering a VBA gap does not make VBA
  in scope.
- A capability that lands must update the public copy in the same merge.

---

## Benchmark harness

**Status: built at `15d97db`.** Run it with `bun run bench:vsdx`.

Six synthetic fixtures generated on demand under `./.scratch/` and deleted afterwards —
nothing large is committed and no path can reach the proprietary corpus. One warm-up, then
the median of five samples per stage. Results and the commit SHA are recorded in
`scripts/vsdx-bench-results.json`. A stage fails the run when it regresses by **both** more
than 30% and more than 5 ms — two conditions, so ordinary Windows scheduling noise does not
trip it while a real multi-millisecond regression still does.

Verified independently by re-running it after merge: numbers reproduce to within noise
(`text-heavy` render 1157.295 ms against the agent's reported 1157.523 ms), the working
tree stays clean afterwards, and `./.scratch/` is removed.

Requirements:

- Extend `scripts/create-vsdx-fixture.ts` (deterministic, fixed zip date) to emit
  complex synthetic diagrams. Scale independently: shape count, group nesting depth,
  connector/glue density, master and style inheritance depth, page count, text volume.
- Time each stage **separately** — parse, resolve, evaluate, render to display list,
  save — so a regression points at a stage. A single end-to-end number is not enough.
- Record absolute numbers plus commit SHA. Fail loudly on regression against the last
  recorded run.
- Small generated corpora may be committed; large ones are generated on demand.
- Never benchmark against the proprietary corpus in a committed artifact — timings
  derived from it may be recorded, but the files themselves must never enter the repo.

### Results

Measured 2026-08-15 at `15d97db`, median of 5 after warm-up, milliseconds.

| Fixture | Parse | Resolve | Eval | Render | Save |
|---|---:|---:|---:|---:|---:|
| deep-inheritance | 28.4 | 85.7 | 9.9 | 190.0 | 1.9 |
| dense-glue | 53.9 | 58.7 | 11.8 | 127.4 | 4.2 |
| many-pages | 33.1 | 41.4 | 12.0 | 97.2 | 1.4 |
| nested-groups | 55.9 | 189.4 | 19.9 | 383.6 | 3.1 |
| shape-heavy | 68.2 | 116.4 | 23.6 | 270.0 | 3.3 |
| text-heavy | 18.7 | 24.7 | 6.3 | **1157.5** | 3.0 |

### Text hotspot, attempt 1 at `711d6e4` — right inefficiency, wrong cause

`wrap_paragraph` looked up "which run contains character N" by rescanning the whole run
list per character, twice, and allocated a `String` per character to measure it. That is
O(n·r) and looked like the obvious source of the superlinearity. Replaced with a forward
cursor and `char::encode_utf8`.

**The hypothesis was wrong.** Measured before and after, scaling text at a fixed 320 shapes:

| textBytes | before | after | doubling ratio before → after |
|---:|---:|---:|---|
| 512 | 264.4 ms | 262.4 ms | |
| 1024 | 519.4 ms | 508.3 ms | 1.96 → 1.94 |
| 2048 | 1150.1 ms | 1119.4 ms | 2.21 → 2.20 |
| 4096 | 2931.1 ms | 2852.1 ms | 2.55 → **2.55** |

A 3% constant-factor win and **the exponent did not move**. The run rescan was real waste
but contributed almost nothing. The change is kept — it is strictly less work and adds a
run-boundary regression test — but it must not be described as fixing the hotspot.

Worth recording as method: the only reason this is known is that the scaling experiment was
re-run *after* the fix. Had the 3% been reported as "improved text rendering" without
re-measuring the exponent, the real defect would have been marked done and buried.

### Round-trip is now property-tested — `5079739`

`CLAUDE.md` calls lossless round-trip the core invariant, and it was exercised against real
diagrams until the corpus vanished from this machine. `crates/vsdx-parse/tests/roundtrip_property.rs`
rebuilds that assurance synthetically, which is corpus-independent and had never been done.

256 generated packages per property, fixed seed `0x5EED_C0DE_D15E_A5E5` so any failure is
exactly reproducible. No dependency added — `proptest`, `arbitrary`, `quickcheck` and `rand`
are absent from `Cargo.lock` and cargo runs offline here, so the generator uses a small
inline PRNG.

Properties: per-part **byte-identical** re-serialisation, parse idempotence, and no panic.
The generator varies nested groups, `Del` markers, duplicate cell names, `F`-only / `V`-only
/ both cells, randomised attribute order and quoting (single quotes as real Visio writes),
opaque unknown XML at every level, non-ASCII including non-BMP, entities and whitespace.

**No round-trip failures were found.** That is a real result rather than a vacuous pass, and
it was verified by fault injection rather than assumed:

| Injected fault | Expected | Observed |
|---|---|---|
| Flip a byte inside the serialised zip | detected | fails — "serialized package did not parse" (CRC) |
| Append a byte to the *expected* part | byte assertion fires | fails — **"round-trip diverged in [Content_Types].xml"** |

The second injection is the meaningful one: it exercises the byte-equality assertion itself
rather than the zip checksum. Both were reverted; the committed test is unmodified.

A property test that has only ever been seen passing is not known to work — the same reason
the benchmark gate was proven two-sided at `e57cf8f`.

### `text-heavy` render characterised — further work needs a profiler, not code reading

Measured rather than guessed, and **no change was made.** What is now known:

- **The cost is per character, not per shape.** `shape-heavy` (1200 shapes, 19 K chars)
  renders in 75 ms, implying ~56 µs/shape. Applying that to `text-heavy`'s 320 shapes
  accounts for ~18 ms of its 284 ms; the remaining ~265 ms is per-character, ~400 ns each.
- **Growth is linear** (ladder ratios 1.54/1.69/1.81), so there is no complexity defect left
  here — this is a constant.
- **The measure cache is working.** The fixture text is `'benchmark '.repeat(...)`, i.e. ten
  distinct characters, so the `(font, size bits, char)` cache hits essentially always.
  Shaping is not the cost.
- **The benchmark's `render` stage is `layout_page` only** — no JSON serialisation — so the
  number is genuine layout, not measurement artefact. Note `Renderer::default()` is
  constructed inside the timed region, which is a small constant charged to every fixture.
- **Per character the path builds one `CaretStop`**, then up to three separate passes
  re-traverse `&mut caret_stops` to adjust positions: first-line indent, line spacing, and
  vertical alignment. The vertical-alignment pass at `vsdx-render:1047` **always** runs, even
  when `dy == 0.0`, which is the common case. One `Vec<CaretStop>` is also allocated per line.

Those traversals are real waste but account for a few milliseconds, not 265. **Reading the
code has not identified where ~400 ns/char goes, and two previous attempts to diagnose this
path by inspection were wrong** — the run-rescan fix bought 3%, and the cell-index change
was a 17.6% regression. The next attempt should use a sampling profiler and be dispatched
only once the hot frame is known.

A cheap, certain, unrelated tidy while someone is in there: skip the vertical-alignment
traversal entirely when `dy == 0.0`. **Caution:** `-0.0 + 0.0` is `+0.0`, so skipping the
add preserves `-0.0` where today it becomes `+0.0`. That is numerically equal but can differ
in serialised output, so it needs checking rather than assuming.

### Benchmark gate tightened and proven to fire — `e57cf8f`

The gate that let a 17.6% regression through is fixed. Threshold derived from measurement,
not taste: four consecutive runs, each already a median of five, gave worst-case spread ~6%
(`deep-inheritance` parse 6.2%, render 5.5%), with the noisiest entries being the
small-absolute stages where a few percent is a fraction of a millisecond.

`percent: 30, minimumMs: 5` → **`percent: 12, minimumMs: 3`**, keeping the both-conditions
rule. That is ~2× headroom over observed noise. The absolute floor still matters: a real run
shows `shape-heavy save +9.7%` which is only +0.283 ms, correctly not flagged.

The run now prints a signed ms and percent delta for **every** stage plus the baseline commit,
so drift sitting just under the threshold is visible instead of silent. Previously a run
printed absolute times and said nothing unless it tripped.

**Verified two-sided, at realistic magnitudes**, by editing the recorded baseline and
restoring it:

| Case | simulated change | required | observed |
|---|---|---|---|
| A | +17.6% (the regression that slipped) | fail | **exit 1** |
| B | +6.1% (noise level) | pass | **exit 0** |

A gate that has only ever been seen passing is not known to work. Case B matters as much as
case A — a threshold tight enough to fire on noise would be abandoned within a week.

Full gate green: vsdx (22/26/58/57), wasm-gated (55/54/1), fmt, clippy, TS (96).

### REJECTED: indexing chain cells by name made resolve 6–18% slower

Attempted and **discarded, not merged.** `resolve_cell` finds each cell by linear scan over
each source in the chain — shape, every master, every style sheet, page, document — once per
cell name, which is O(k²·d) string comparisons per shape. Replacing those scans with a
`HashMap<&str, &Cell>` per source looked like a clear win.

It is a loss. Same-worktree A/B, change stashed and unstashed so the build environment is
identical:

| Fixture | control | with index | delta |
|---|---:|---:|---:|
| deep-inheritance resolve | 84.5 | 99.4 | **+17.6%** |
| nested-groups resolve | 74.5 | 79.8 | +7.1% |
| shape-heavy resolve | 94.6 | 99.9 | +5.6% |

Stable over three runs each; not noise.

**Why the hypothesis was wrong.** `k`, the number of cells on a shape, is small — tens, not
hundreds — so O(k²) with a tiny constant beats hashing plus allocation. Worse, the index was
built **per shape** for the page and document sheets too. A linear `find` stops at the first
match; building a full index costs the whole sheet every time. For large page or document
sheets that is a large new cost per shape, which is why `deep-inheritance` (longest chain,
so most sources indexed) regressed most.

**What would actually pay off:** hoisting the page and document indexes out of the per-shape
loop so they are built once per page. That was explicitly deferred out of this task as "a
larger change, its own commit" — it turns out to be the *only* part worth doing, and the
per-shape indexing must not come with it. Not yet attempted.

The rejected patch is kept at `scratchpad/cellindex-rejected.patch` for whoever tries the
hoisted version.

**The benchmark did not catch this.** A +17.6% resolve regression passed the gate, because
the rule requires **both** >30% and >5 ms. The 30% threshold was chosen to avoid tripping on
Windows scheduling noise, but it is loose enough to let a real regression of this size
through unnoticed. Two options worth considering: tighten the percentage now that the
median-of-five is demonstrably stable to ~1%, or fail on a smaller percentage when the
absolute delta is large. **The agent's own gate reported "benchmark passed" for a change
that was 17.6% slower** — this is the only reason to re-measure every performance claim
independently rather than trusting a green gate.

### Second O(n²) removed at `5889310` — render is now linear in nesting depth

Re-measuring the depth ladder after `cfe4024` showed render's exponent had got *worse*
(2.19/2.91 against the earlier 1.91/2.55) even though absolute times had halved. That is
what happens when a large linear component is removed: the superlinear term stops being
masked. Re-measuring rather than reusing the old figure is what exposed it.

The cause was the same defect as `156d816`, in a second place. `layout_shape`
(`vsdx-render:638`) resolved every shape **by id**:

```rust
let resolved = resolver.resolve_shape(page_part, shape.id)?;
```

`resolve_shape` calls `find_shape`, a recursive search of the whole page tree — O(n) per
shape, O(n²) per page — and it re-did inheritance resolution that render already had in
hand. Now it looks the shape up in the existing map, falling back to the by-id path only
when the map is unavailable. No clone.

Depth ladder, 160 leaves:

| groupDepth | parse | render before | render after |
|---:|---:|---:|---:|
| 6  | 33.0 | 60.9 | **29.6** |
| 12 | 57.4 | 133.5 | **48.1** |
| 24 | 104.8 | 388.0 | **87.0** |

**Render doubling ratios 2.19/2.91 → 1.63/1.81**, against parse's 1.74/1.83. Linear.

### The hotspot has inverted: resolve is now the slowest stage

Baseline at `e2361a30`:

| Fixture | parse | resolve | evaluate | render | save |
|---|---:|---:|---:|---:|---:|
| deep-inheritance | 28.8 | **82.5** | 10.1 | 54.3 | 1.7 |
| dense-glue | 54.1 | **51.5** | 12.0 | 43.1 | 3.9 |
| many-pages | 32.9 | **40.0** | 12.2 | 30.8 | 1.3 |
| nested-groups | 56.4 | **72.7** | 20.5 | 49.9 | 2.9 |
| shape-heavy | 68.0 | **90.4** | 23.9 | 74.3 | 2.9 |
| text-heavy | 18.6 | 23.5 | 6.5 | **282.9** | 2.8 |

Render was the dominant stage on all six fixtures when the harness was built. It is now
behind resolve on five of six, and behind parse on three. **Resolve is the new top target**,
except on `text-heavy` where render still dominates by 12×.

Cumulative render from the first baseline at `15d97db`:

| Fixture | then | now | change |
|---|---:|---:|---:|
| deep-inheritance | 190.0 | 54.3 | −71% |
| dense-glue | 127.4 | 43.1 | −66% |
| many-pages | 97.2 | 30.8 | −68% |
| nested-groups | 383.6 | 49.9 | **−87%** |
| shape-heavy | 270.0 | 74.3 | −72% |
| text-heavy | 1157.5 | 282.9 | −76% |

**Next targets, measured but not diagnosed:**

1. **Resolve**, now slowest on five of six. `deep-inheritance` resolve is 82.5 ms against a
   28.8 ms parse — 2.9× — and the inheritance-depth ladder showed resolve growing while
   parse stayed flat, so per-shape chain walking is the obvious place to look. Sublinear in
   depth, so this is a constant-factor problem rather than a complexity defect.
2. **`text-heavy` render, 282.9 ms**, still 12× its own resolve and 3.8× the next-worst
   render. Two rounds of text work have not closed it.

Neither is diagnosed. The pattern this run — three of five hypotheses were wrong or
incomplete until measured — argues for profiling before dispatching either.

### Render now resolves the page once — `cfe4024`

`resolve_page_connectivity` resolved the whole page for itself (`connectivity.rs:175`), so
after `9c6b8cf` render still resolved twice. A `resolve_page_connectivity_with(page_part,
&shapes)` variant now takes an already-resolved map; the original entry point stays as a
thin wrapper because four call sites outside the change depend on it
(`betteroffice-vsdx:415`, `vsdx-bench:72`, `vsdx-render:3779`, `vsdx-resolve/tests.rs`).

Render, all six fixtures:

| Fixture | before | after | change |
|---|---:|---:|---:|
| deep-inheritance | 138.5 | **94.9** | −31% |
| dense-glue | 92.5 | **66.4** | −28% |
| many-pages | 69.4 | **50.1** | −28% |
| nested-groups | 172.6 | **139.5** | −19% |
| shape-heavy | 168.0 | **124.3** | −26% |
| text-heavy | 315.2 | **304.6** | −3% |

Unlike the previous commit this improves every fixture, including `text-heavy` — every page
render was paying for one redundant full resolution regardless of what the page contained.

An equivalence test pins the two connectivity paths against each other so they cannot drift.

Baseline re-recorded at `8f36cd2e`.

### Cumulative effect of the performance run

From the first recorded baseline at `15d97db` to `8f36cd2e`:

| Fixture | render before | render now | resolve before | resolve now |
|---|---:|---:|---:|---:|
| deep-inheritance | 190.0 | **94.9** | 85.7 | 81.6 |
| dense-glue | 127.4 | **66.4** | 58.7 | 53.8 |
| many-pages | 97.2 | **50.1** | 41.4 | 40.2 |
| nested-groups | 383.6 | **139.5** | 189.4 | 72.7 |
| shape-heavy | 270.0 | **124.3** | 116.4 | 90.0 |
| text-heavy | 1157.5 | **304.6** | 24.7 | 23.6 |

Render is 48%–74% faster across the suite; `nested-groups` resolve is 62% faster. Two
genuine complexity defects were removed (quadratic line breaking, quadratic page
resolution) plus three redundancies (per-character font resolution and shaping, two
duplicate page resolutions).

**Open performance questions, none diagnosed:**

1. `text-heavy` render is still 304.6 ms, 2.2× the next worst, and its parse/resolve/evaluate
   are the *cheapest* in the suite. Text remains disproportionately expensive even after the
   80% reduction.
2. Render is still superlinear in group nesting depth (ratios 1.91/2.55 against parse's
   1.87 when last measured at `156d816`) — worth re-measuring now that three redundant
   resolutions are gone, since the earlier figure included them.
3. `deep-inheritance` and `shape-heavy` resolve barely moved (85.7→81.6, 116.4→90.0). The
   quadratic fix helped nesting; inheritance depth and raw shape count are untouched.

### Duplicate page resolution removed at `9c6b8cf`

`PageShapeReferences::new` (`vsdx-render:541`) already resolves every shape on the page —
internally `resolve_page_shapes` at `vsdx-eval:219` — and the next line resolved the whole
page again. Render did full page resolution twice and discarded one copy. It now reuses the
map `references` already holds, falling back to a direct resolve only when `references` is
`None`, with no clone of the map.

Render, all six fixtures:

| Fixture | before | after | change |
|---|---:|---:|---:|
| deep-inheritance | 181.0 | **138.5** | −23% |
| dense-glue | 113.0 | **92.5** | −18% |
| many-pages | 89.6 | **69.4** | −23% |
| nested-groups | 212.7 | **172.6** | −19% |
| shape-heavy | 218.0 | **168.0** | −23% |
| text-heavy | 326.0 | **315.2** | −3% |

`text-heavy` barely moves, which is the expected shape: it has only 320 shapes, so page
resolution was never a large part of its render — its cost is text, and that is already
addressed.

The `None` branch was the risk here — it is the untested path, and silently turning
"resolve directly" into "give up" would not have failed most tests. It is preserved and now
covered by `missing_page_sheet_preserves_page_dimension_error`.

**Still duplicated: `resolve_page_connectivity` (`vsdx-render:540`) independently calls
`resolve_page_shapes` as well.** So render still resolves the page twice, not three times.
Closing that needs a reusable-map API on the connectivity path, which is a change in
`vsdx-resolve` rather than render, and was correctly left out of this scoped commit. It is
the next certain win.

Baseline re-recorded at `46aa90d6`.

### O(n²) page resolution removed at `156d816`

Scaling group depth exposed it: with the leaf count fixed at 160, doubling depth multiplied
parse by 1.87, evaluate by 1.84 and save by 1.6 — all linear — while **resolve went ×3.37
and render ×3.40**.

`resolve_page_shape_tree` (`vsdx-resolve/src/inheritance.rs:102`) walked the shape tree and
for each shape called `resolve_shape(page_part, shape.id)`, which called `find_shape` — a
recursive search of the **entire** page tree. It already held the `&Shape`, discarded it,
and searched the whole tree to find the same shape again: O(n) per shape, O(n²) overall.

Fixed by resolving the in-hand `&Shape`, with the page-sheet fallback computed once per walk
instead of per shape. The subtle part is that `resolve_shape` does **not** use
`page_contents` as the inheritance host — it prefers the page *sheet* and only falls back —
so passing the wrong sheet would have silently changed inheritance rather than failing. An
equivalence test now pins a deep leaf resolved by the new path against the by-id path.

This also explains why render tracked resolve so closely: render calls `resolve_page_shapes`
(`vsdx-render:542`) **and** `PageShapeReferences::new` (`:541`), which resolves the page
again — so render paid the quadratic cost twice.

| groupDepth | resolve before | resolve after | render before | render after |
|---:|---:|---:|---:|---:|
| 6  | 76.0 | **40.1** | 153.2 | **104.6** |
| 12 | 194.4 | **68.5** | 386.8 | **199.9** |
| 24 | 655.7 | **123.7** | 1315.3 | **510.4** |

**Resolve is now linear**: doubling ratios 2.56/3.37 → 1.71/1.81, matching parse's 1.87.
At depth 24 it is 81% faster.

**Render is not yet linear.** It improved 61% at depth 24, but its doubling ratios are
1.91/2.55 against parse's 1.87 — the top of the ladder is still superlinear. Removing the
quadratic resolution removed most of the cost but not all of the growth, so **something in
render itself still scales worse than linearly with nesting depth.** Candidate worth
measuring, not assuming: per-shape scene-transform composition walking ancestors, which
would be O(n·d). Also still open is the double resolution at `vsdx-render:541-542`, now a
2× constant rather than a 2× quadratic.

Baseline re-recorded at `b37c4383`.

### Text hotspot resolved at `20b8b18` — and it is no longer the worst fixture

`measure` ran per character and did two expensive things every time: `font_for` built up to
12 candidate keys, each doing `family.into()` and thus **heap-allocating a `String` just to
probe a `BTreeMap`**, and `ooxml_text::shape` performed full shaping setup to return one
glyph advance. Both are now memoised in a page-local cache keyed
`(Option<FontId>, size_in.to_bits(), char)` — exact bits, no tolerance.

Cumulative effect across all three text commits, same scaling ladder:

| textBytes | original | `711d6e4` | `e39a43a` | `20b8b18` | total |
|---:|---:|---:|---:|---:|---:|
| 512 | 264.4 | 262.4 | 236.9 | **126.1** | −52% |
| 1024 | 519.4 | 508.3 | 423.4 | **194.2** | −63% |
| 2048 | 1150.1 | 1119.4 | 782.2 | **328.1** | −71% |
| 4096 | 2931.1 | 2852.1 | 1492.7 | **592.7** | **−80%** |

Doubling ratios 1.96/2.21/2.55 → 1.54/1.69/1.81; per-character cost 1.14 µs → 0.45 µs.
Growth stays linear and the constant fell by more than half.

**The hotspot has moved.** `text-heavy` render was 1157 ms and by far the worst; it is now
327.9 ms, and `nested-groups` at 390.0 ms is the slowest render in the suite. Ranking on the
re-recorded baseline at `e25df7c7`:

| Fixture | Render | Next largest stage |
|---|---:|---:|
| nested-groups | **390.0** | resolve 193.9 |
| text-heavy | 327.9 | resolve 25.6 |
| shape-heavy | 260.7 | resolve 119.8 |
| deep-inheritance | 189.5 | resolve 87.8 |
| dense-glue | 122.2 | resolve 59.9 |
| many-pages | 90.4 | resolve 41.5 |

Baseline re-recorded so regressions are caught against the improved numbers.

**Next candidates**, in order of what the data supports:

1. `nested-groups` — now the worst render *and* the worst resolve (193.9 ms, 1.6× the next).
   Group nesting is expensive in both stages, which suggests one shared cause rather than
   two, and that makes it the highest-value thing to look at next.
2. Render still runs 1.6×–13× its resolve stage on every fixture. There may be a systemic
   render cost independent of any one knob.

Neither has been investigated. Do not assume either diagnosis without measuring — the first
attempt at the text hotspot targeted a real inefficiency that turned out to contribute
almost nothing.

### FIXED at `e39a43a` — the quadratic term is gone

Attempt 2 bounded all three scans. Measured on the same scaling ladder, 320 shapes fixed:

| textBytes | original | after `711d6e4` | after `e39a43a` | total change |
|---:|---:|---:|---:|---:|
| 512 | 264.4 | 262.4 | **236.9** | −10% |
| 1024 | 519.4 | 508.3 | **423.4** | −18% |
| 2048 | 1150.1 | 1119.4 | **782.2** | −32% |
| 4096 | 2931.1 | 2852.1 | **1492.7** | **−49%** |

**Doubling ratios: 1.96 / 2.21 / 2.55 → 1.79 / 1.85 / 1.91.** That was the acceptance
criterion, and it is met: every ratio is now below 2.0 and rising *toward* it as fixed costs
amortise, which is the signature of linear scaling rather than quadratic. The implied
exponent fell from ~1.35 to ~0.93.

The confirming detail is that the saving **grows with input size** — 10% at 512 bytes, 49%
at 4096. A constant-factor win would have been flat across the ladder. That shape is what
distinguishes this from attempt 1.

The implementer used exact-order slice sums rather than prefix sums, so floating-point
accumulation order is unchanged, and no existing test expectation was modified.

Baseline re-recorded at `e0ceefab` so future regressions are measured against the improved
numbers rather than the old slow ones — otherwise a regression back to 1150 ms would not
trip the gate. `text-heavy` render baseline is now 783.9 ms.

Remaining: render still dominates every fixture, and `text-heavy` render is still ~4× the
other stages combined. The next candidate is `measure`, which calls `ooxml_text::shape`
once **per character**; memoising it on `(family, bold, italic, size, char)` is safe because
it is a pure function of those inputs. Not yet attempted.

### Original diagnosis: three whole-array scans in line breaking

Found by reading the breaking loop after the exponent refused to move. All three scan an
entire array to reach a contiguous range whose bounds are already known:

| Site | Cost | Note |
|---|---|---|
| `breaks.iter().find(...)` per character (`wrap_paragraph`) | **O(n·b)** | dominant — runs per character over a list that grows with text |
| `cursor` recomputed by filtering all `widths` after each break | O(n·lines) | |
| `fn line` filters all `widths` twice per line (`:1188`, `:1200`) | O(n·lines) | width sum and caret stops |

`breaks` is ordered by `byte_index` and the loop's `next` increases monotonically, so a
single advancing index makes the first O(1) amortised. `widths` is built in increasing
position order, so `[start, end)` is a contiguous slice reachable by index rather than by
filtering the whole array.

Dispatched as attempt 2. The acceptance criterion is the **exponent**, not the wall time:
the doubling ratio must fall toward ~2.0. A constant-factor win that leaves it at 2.55 has
not fixed this.

Float caution flagged to the implementer: summation order is observable in the assertions,
so `prefix[end] - prefix[start]` is not necessarily bit-identical to summing the range, and
exact-order slice sums are preferred if any expectation shifts.

### PRIORITY-2 ITEM: text rendering is the dominant hotspot

The harness earned its keep on the first run. **`text-heavy` spends 1157 ms in render** —
three times the next-worst fixture (`nested-groups`, 383 ms) and twelve times the cheapest
(`many-pages`, 97 ms).

What makes it damning is the rest of that row: `text-heavy` has the **cheapest** parse
(18.7 ms), resolve (24.7 ms) and evaluate (6.3 ms) of all six fixtures. So the cost is not
"this fixture is big" — it is specific to rendering text. Render is 95% of that fixture's
total pipeline time.

Cross-checking the other rows, render dominates *every* fixture, running 1.6×–2.3× its
resolve stage. Render is where the time goes across the board; text is where it goes
pathologically.

This is the first entry that priority 2 of the work order — "a measured regression or
hotspot from a benchmark" — has ever been able to point at. Next step is to profile the
text path in `vsdx-render` and find whether the cost is superlinear in text volume, which
the harness can answer directly by scaling the text knob and re-running.

Note the numbers above are synthetic-fixture numbers and say nothing about real diagrams;
that comparison needs the corpus, which is currently missing.

---

## Open — work queued

### Geometry completion (in scope as of phase 11)

`crates/vsdx-resolve/src/geometry.rs:20` skips five row types and **drops the
segment**, so affected shapes render incomplete:

| Row type | Notes |
|---|---|
| `PolylineTo` | Simplest — a point list. Start here. |
| `InfiniteLine` | A line extended to the shape bounds. |
| `SplineStart` / `SplineKnot` | Standard spline-to-bezier conversion. |
| `NURBSTo` | NURBS to bezier; documented in MS-VSDX. Rendering only — editing stays excluded. |

Evidence: `GeometryIssue::UnsupportedRowType` emitted at `geometry.rs:22-23`; test at
`geometry.rs:540-551` asserts the current skip behaviour and will need updating.

### QUEUE RESET — the ranked list below is void

The five ranked items were derived from the "2,246 unexplained formulas" premise, which is
false (see *Parity gaps found*). Items 1a, 1b, 2 and 4 all target documented non-goals or
already-correct behaviour, and none of them can raise coverage honestly. **Do not work
them.** They are struck through rather than deleted so nobody re-derives them.

What actually remains, and it is a short list:

1. ~~Correct the `error`/non-goal misclassification introduced by `00dc323`.~~ **Done at
   `73ec02e`** — verified against the corpus, see below.
2. **Report `oracle_no_cache` alongside any coverage figure**, so the `V='Themed'` blind
   spot cannot hide a wrong implementation behind a 100% agreement rate.
3. **`unit/dimension error: incompatible units` — 2 formulas.** The only bucket left that
   looks like a genuine evaluator bug rather than a policy exclusion. Worth one look.
4. **Geometry completion and VSDX redaction** (below) — still open, still real, and now the
   highest-value remaining work in the whole queue. Both are spec-driven rather than
   corpus-driven, which is no longer a mark against them: the corpus is close to exhausted
   as a source of honest coverage gains.

The broader conclusion: **evaluator coverage is near its honest ceiling.** Roughly 2,068 of
the 3,313 unevaluated formulas are documented exclusions and several hundred more are
correct-as-is. Continuing to push the 52.62% number is the wrong objective; the remaining
value in this codebase is in geometry, redaction, benchmarks and round-trip fidelity.

### ~~From the 2026-08-15 discovery pass, ranked by measured corpus frequency~~ (VOID)

Evidence and root-cause analysis for all of these is under *Parity gaps found* below.

1. ~~`Inh` inheritance host for sheet-level cells~~ — **done at `00dc323`, zero coverage
   gain.** Superseded by 1a and 1b below; see the correction under *Parity gaps found*.

1a. **Theme-terminated inheritance + `THEMEVAL` — ~492 formulas.** The 306 `THEMEVAL`
   failures and ~186 of the residual `Inh` cells are the same gap: sheet-level cells have
   no theme context. Largest remaining single piece of work, and the two halves should not
   be split across agents.

1b. **Built-in ShapeSheet defaults as the inheritance terminator — ~312 formulas.** When
   the chain is exhausted Visio substitutes a hardcoded per-cell default. Oracle-validated
   by the cached `@V`. **The 5 locale cells are not constants** — `de-DE` in one corpus
   file, `en-US` in the other — and must come from the document.

2. **Reclassify the 937 event cells from `unsupported_other` to `unsupported_known`.**
   Measurement honesty only — they are the permanently-excluded actions/events set, and
   counting them as "unexplained" overstates both the pool and the coverage denominator.
   Must not be reported as a coverage improvement.
3. **`THEMEVAL` theme-cell context — 306 formulas.** 234 lack host-cell context
   (`vsdx-eval/src/lib.rs:924`), 72 have no resolvable theme (`:929`). Needs a look at
   whether the theme is genuinely absent or merely not threaded through.
4. **Unresolved cross-sheet references — 163 formulas.** 136 are
   `Sheet.N!Connections.X*` (connection-point lookups across shapes), 27 are
   `ThePage!DrawingScale`. The latter is one page-sheet cell not exposed to cross-sheet
   resolution and is likely the cheapest win in the whole queue.
5. **`unit/dimension error: incompatible units` — 2 formulas.** Low count but this is the
   only bucket that smells like a genuine evaluator bug rather than a missing feature.
   Worth one investigation to confirm it is not the tip of something.

Items 3 and 4 touch `crates/vsdx-eval/src/lib.rs`, the same file as item 1 — they cannot
be dispatched in parallel with it without racing the fence. Sequence them.

### VSDX redaction (un-excluded as of phase 11)

`crates/ooxml-redact/src/lib.rs:24-26` implements DOCX/XLSX/PPTX behind a format enum.
Add VSDX as a fourth, following the existing pattern. `crates/ooxml-redact-cli` and
`apps/redact-worker` are already wired. Must handle text, metadata and embedded media.

---

## Parity gaps found

Log each as: what Visio does, evidence (spec section or corpus file), what we do
instead, and whether it is worth fixing.

### Geometry row-type coverage vs corpus frequency — measured 2026-08-15 at `d9ee0b2`

Surveyed every `<Row T='…'>` across all masters and pages of both corpus files.
Frequencies, against what `crates/vsdx-resolve/src/geometry.rs` implements today:

| Row type | Corpus count | Implemented |
|---|---|---|
| `RelLineTo` | 2404 | yes (`geometry.rs:90`) |
| `RelMoveTo` | 601 | yes (`geometry.rs:80`) |
| `LineTo` | 430 | yes (`geometry.rs:75`) |
| `ArcTo` | 92 | yes (`geometry.rs:100`) |
| `MoveTo` | 54 | yes (`geometry.rs:70`) |
| `EllipticalArcTo` | 2 | yes (`geometry.rs:108`) |
| `PolylineTo` | 0 | no — skipped at `geometry.rs:20` |
| `InfiniteLine` | 0 | no — skipped at `geometry.rs:20` |
| `SplineStart` / `SplineKnot` | 0 | no — skipped at `geometry.rs:20` |
| `NURBSTo` | 0 | no — skipped at `geometry.rs:20` |

**Finding: every row type the corpus actually uses is already implemented.** All five
skipped types have zero corpus occurrences. The geometry-completion queue is therefore
spec-driven, not corpus-driven — synthetic fixtures are the only coverage available
for it, and the "highest corpus frequency first" rule cannot rank these five. This
does not remove them from phase 11 scope; it does mean they are lower value than the
Open list implied, and that no corpus regression can catch a mistake in them.

Incidental: the corpus writes XML attributes **single-quoted** (`<Row T='RelLineTo'
IX='2'>`). Fixture generators and any string-matching test must not assume double
quotes.

### CORRECTION: the "2,246 unexplained formulas" were never unexplained — 2026-08-15

**Everything below this heading about a large unmined pool is wrong, and the two
iterations of discovery work it drove were largely misdirected.** Left in place because
the reasoning is instructive, but do not act on it.

`crates/vsdx-eval/README.md:17-30` — the eval crate's own documentation, which predates
this log — already lists these as *explicit, intentional non-goals*, with the same counts
this log spent two iterations "discovering":

| README non-goal | Count | Matches the bucket |
|---|---|---|
| Event cells | 937 | `event cell is outside the display evaluation profile` |
| `Inh` in raw catalog sheets | 500 | `Inh requires an inheritance host` |
| Missing `DocLangID` | 325 | `Inh has no concrete inherited value` |
| `THEMEVAL` without host context or a theme | 306 | both `THEMEVAL` buckets |
| `SHADE` / `LUMDIFF` | 1058 | semantics undocumented by Microsoft |

That is **2,068 of the 2,246**. The pool was documented, deliberate, and closed. The
premise that it was "by far the largest known pool of undiscovered work" was false.

The README is also explicit about *why* the 500 catalog-sheet `Inh` cells are excluded:
"catalog sheets have no inheritance graph, so resolving these values **would manufacture
results**." Queue item 1b as previously written — synthesising Visio built-in defaults so
those cells produce values — is precisely the thing that sentence forbids. It is removed
from the queue, not deferred.

**Read `crates/vsdx-eval/README.md` before mining the evaluator buckets again.** A bucket
being large is not evidence it is unmined.

### Consequence: the `00dc323` change made classification worse

Giving catalog-sheet cells an inheritance host moved 456 documented non-goals out of
`unsupported_other` and into `error`. They are not errors — they are intentional
exclusions, and `error` implies a failure that does not exist. The resolver half of that
change (style parents on `resolve_sheet`) was a genuine fix and stays; the reclassification
side-effect should be corrected so documented non-goals report as non-goals.

### Non-goal classification corrected — measured 2026-08-15 at `73ec02e`

Exhausted `Inh` now reports as a documented non-goal instead of an evaluation error,
matching what `crates/vsdx-eval/README.md` already said. Measured on the corpus:

| | `4ee08e9` | `00dc323` | `73ec02e` |
|---|---|---|---|
| `evaluated` | 3679 | 3679 | **3679** |
| coverage | 52.62% | 52.62% | **52.62%** |
| `unsupported_known` | 1067 | 1067 | **1892** |
| `unsupported_other` | 1756 | 1300 | **1256** |
| `error` | 490 | 946 | **165** |
| `@V` agreement / disagreement / no-cache | 3670 / 0 / 0 | 3670 / 0 / 0 | **3670 / 0 / 0** |

`evaluated` is deliberately unchanged — this moves formulas between non-evaluated buckets
and nothing else. That was the acceptance criterion, and it held.

The `error` bucket is now 5.7x smaller and **fully explained**: 163 unresolved cross-sheet
references (27 correct-as-is from a master with no page context, 136 that would fall
through to `PAR`/`PNT`) plus the 2 `unit/dimension error: incompatible units`. Those 2 are
now the only entry in the entire error bucket that might be a genuine evaluator bug.

Full gate run and seen passing on the merged result: rust core (22/26/52/41), wasm-gated
(55/54/1), `fmt`, `clippy -D warnings`, vsdx wasm build, TS packages (96), apps/web (27),
apps/demo (7), 13 typechecks. No capability change, so no public copy update was due.

### FIXED at `ae4aabb` — `RelLineTo`/`RelMoveTo` now shape-relative

Confirmed and fixed. The investigation the entry below asked for came back positive:
`bounds_affine` (`connectivity.rs:77`) sets `scale_x = scale_y = 1.0` when `child_extent`
is `None`, the normal leaf-shape case, so `transform.local` only rotates, flips and
translates to the pin. **Nothing downstream compensated** — realized geometry reached the
display list in local inches, unscaled, and the corpus rectangle really did render as a
staircase.

`realize_geometry` now takes the shape's bounds:

```rust
pub fn realize_geometry(section: &ResolvedSection, width: f64, height: f64) -> RealizedGeometry
```

`Rel*` rows compute `(x·width, y·height)` as an **absolute** point and *set* `current`
rather than accumulating; `PolylineTo` relative flags use the same fractions;
`InfiniteLine` clips to `[0,width]×[0,height]` instead of the unit-box stand-in. Non-finite
bounds emit a `GeometryIssue`; zero bounds produce finite degenerate geometry.
`MoveTo`, `LineTo`, `ArcTo`, `EllipticalArcTo` are untouched, and the three spline/NURBS
types remain skipped.

Regression test: the corpus rectangle at `Width=4, Height=3` realizes exactly
`(0,0) → (4,0) → (4,3) → (0,3)`. Also covered: `Rel*` non-accumulation (two identical rows
give the same point twice), non-finite bounds, zero bounds.

Correcting the record from `8b11077`: the claim that `realize_geometry`'s signature was
"pinned by read-only callers" was **false**. There is exactly one production caller
(`vsdx-render/src/lib.rs:653`) and the shape's `bounds` was already in scope there;
everything else was a test. Both limitations disclosed with that commit — the unit-box
`InfiniteLine` clip and the `PolylineTo` flag convention — were never actually forced, and
both are now resolved by the same bounds-threading change.

Full gate run and seen passing, corpus set: rust core (22/26/52/**55**, resolve up from 50),
wasm-gated (55/54/1), `fmt`, `clippy -D warnings`, wasm build, TS (96), web (27), demo (7),
13 typechecks, fixtures regenerate byte-identically. Corpus measurement unmoved
(`evaluated=3679`, 52.62%, `@V` 3670/0/0) — expected: this is geometry, which the formula
measurement does not observe at all.

**No public copy change due** — `llms.txt`, `content.ts` and `README.md` describe geometry
at crate level and never enumerate row types. Checked, not assumed.

### VSDX redaction landed at `8629565` — with a shape-data gap that must be closed

`Format::Vsdx` is now the fourth redaction format. Detection by
`application/vnd.ms-visio.drawing.main+xml` with a `visio/document.xml` fallback; text runs
inside `<Text>` redacted with `<cp/>`/`<pp/>`/`<tp/>` markers preserved; `docProps`
metadata scrubbed; `visio/media/*` and `docProps/thumbnail.emf` replaced. `.vsdm` is
**rejected** via `RedactError::MacroEnabledVisio` rather than redacted — note the
DOCX/XLSX/PPTX arms deliberately *accept* their `macroenabled` types, so this asymmetry is
intentional and must not be "fixed".

**Gap closed at `cb6b227`.** A `V` attribute is now redacted when — and only when — the
cell is `<Cell N="Value|Prompt|Label">` and some enclosing `<Section>` has
`N="Property"` or `N="User"`. Everything else is preserved, with assertions pinning both
directions: `<Cell N="SortKey" V="123"/>` and `<Cell N="PinX" V="4.25"/>` must survive
verbatim. Over-redaction would destroy the document, so that guard matters as much as the
redaction itself. Covered at page and master level, with XML validity and UTF-8 checked.

**`Name`/`NameU` deliberately left alone**, and this should not be revisited casually.
Visio formulas reference shapes, pages and masters by name, and `NameU` is the universal
form used in ShapeSheet cell references. Blanking them would break references and produce
an invalid document. Redacting them safely needs formula-aware, package-wide deterministic
renaming with reference rewriting — a separate feature, not an extension of this one.

Original finding, kept for context — **VSDX attributes were never redacted**: `attribute_is_redactable`
returns `Format::Vsdx => false`, and text redaction only fires inside `<Text>`. In Visio a
great deal of user data lives in attributes instead:

| Location | Example | Redacted today |
|---|---|---|
| Shape Data / custom properties | `<Section N='Property'><Row N='Owner'><Cell N='Value' V='…'/></Row></Section>` | **no** |
| User-defined cells | `<Section N='User'><Row N='X'><Cell N='Value' V='…'/></Row></Section>` | **no** |
| Shape names | `<Shape Name='…' NameU='…'>` | **no** |

Shape Data is a headline Visio feature — asset tags, owners, costs — so this is a real leak
vector in a security tool, not a cosmetic omission. The dispatch prompt said "when in doubt,
redact more, not less"; this went the other way. Queue this above everything else in the
redaction area, with a test fixture that actually contains a `Property` section.

Caveat on the evidence: the corpus was already gone when this was reviewed (see below), so
the *frequency* of `Property`/`User` sections in real diagrams could not be measured. One
page of one file had none, sampled earlier. Absence in a single sample is not evidence of
absence, and for a redaction feature the burden runs the other way.

**No public copy change was due, and none was made.** Redaction is not mentioned in
`apps/web/public/llms.txt`, `apps/web/app/content.ts` or `README.md` for *any* format —
DOCX, XLSX and PPTX redaction are equally unadvertised. Adding VSDX to an unadvertised
subsystem changes no public claim. Introducing a first-ever public redaction claim for VSDX
specifically, while shape data leaks, would manufacture an overclaim rather than fix one.
Copy should land when the gap above is closed and it can cover all four formats honestly.

Gate run and seen passing on the merged result: redact crates (14/3/4), vsdx crates
(22/26/53/55), wasm-gated (55/54/1), `fmt`, `clippy -D warnings`, wasm build, TS (96),
web (27), demo (7), 13 typechecks. **Run without the corpus** — see the blocker below — so
it proves less than a full gate.

### RESOLVED: corpus restored, and every corpus-free merge re-verified — 2026-08-15

Restored to `G:\Heinrich\Dokumente\vsdx-corpus`, outside the temp tree. Path updated in
`openspec/vsdx-overnight.md`.

**The restored files are content-identical to what the baselines were measured on.** Every
pinned number reproduces exactly:

```
evaluated=3679  unsupported_known=1892  unsupported_other=1256  error=165  total=6992
@V oracle: agreement=3670  disagreement=0  excluded-stale=9  no-cache-available=0
agreement-rate=100.00%   coverage 3679/6992 = 52.62%
```

The stale-oracle exclusions keyed to specific shape ids (`…:page1.xml:shape:387` and
friends) also still match, which is a stronger identity check than the aggregate counts.

**All ~30 commits merged during the outage are now verified against real diagrams**, including
the four round-trip tests that could not run:

- `saves_real_corpus_cells_without_rewriting_other_parts` — pass
- `preserves_external_corpus_parts_when_available` — pass
- `structural_edits_preserve_external_corpus_parts_when_available` — pass
- `model_serializer_matches_original_parts_and_reparse` — pass
- facade `saves_crdt_session_edits_for_every_corpus_file` — pass

So byte-identical round-trip over real files still holds after the geometry, redaction and
performance work. No corpus file lost round-trip; coverage and the `@V` agreement rate are
unmoved.

Note the filename hazard recorded in `vsdx-overnight.md`: the originals are named
`Lichtsysteme System.vsdx` / `SoundplanMPP68.vsdx`, while 17 code sites hardcode
`lichtsysteme.vsdx` / `soundplan.vsdx`. The directory currently carries copies under the
expected names.

### Historical: the corpus directory disappeared — 2026-08-15

`VSDX_CORPUS_DIR` as recorded in `openspec/vsdx-overnight.md` no longer exists:

```
.../04487b0e-bc3e-4073-9731-c8decb847e08/scratchpad/vsdx-corpus   <- gone
```

The sibling files in that same scratchpad survive, so it was not the whole directory tree —
only `vsdx-corpus` and the `/tmp/tmp.*` extraction copies vanished, which is consistent
with a temp-file cleaner reclaiming them. Nothing in this repo deleted it.

**Symptom to recognise:** with `VSDX_CORPUS_DIR` still *set* but pointing at the missing
directory, corpus tests do not skip — they **fail** with OS error 3, "Das System kann den
angegebenen Pfad nicht finden". Five tests fail this way and none of them indicate a code
defect:

- `betteroffice-vsdx` facade: `saves_crdt_session_edits_for_every_corpus_file`
- `vsdx-parse`: `model_serializer_matches_original_parts_and_reparse`,
  `preserves_external_corpus_parts_when_available`,
  `saves_real_corpus_cells_without_rewriting_other_parts`,
  `structural_edits_preserve_external_corpus_parts_when_available`

**Impact.** Until the corpus is restored, the loop loses: the formula coverage and `@V`
oracle measurement, byte-identical round-trip verification over real files, and every
frequency count that ranks the queue. Those are the numbers that make this log meaningful,
so **measured claims cannot be made in this state.** The last good measurement is at
`73ec02e`: `evaluated=3679`, coverage 52.62%, `@V` 3670/0/0, `unsupported_known=1892`,
`unsupported_other=1256`, `error=165`.

Tests still skip gracefully when the variable is *unset*, so the non-corpus gate remains
usable — but a green run in that state proves strictly less, and must be reported as such
rather than as a passing full gate.

**Action needed from the user:** restore the corpus to a stable path outside
`%LOCALAPPDATA%\Temp` (a temp directory was always going to be reclaimed eventually) and
update the path in `openspec/vsdx-overnight.md`.

### NEW GAP: the render suite has zero coverage of the two most common row types

Measured while reviewing `ae4aabb`:

- `crates/vsdx-render/src/lib.rs` contains **0** occurrences of `RelLineTo` or `RelMoveTo`.
- Of seven committed fixtures, only `foundation.vsdx` contains a `Rel*` row at all, and
  that shape never reaches the geometry path — it short-circuits to a
  `"unresolvable transform"` placeholder.

So the two most common geometry row types in real diagrams (2404 + 601 occurrences) had
**no render-level test coverage whatsoever**. That is the direct reason a wrong coordinate
convention survived: the unit tests in `geometry.rs` asserted the buggy behaviour, and no
end-to-end test contradicted them.

Queued: add a render-level fixture whose shapes use `RelLineTo`/`RelMoveTo` with non-unit
`Width`/`Height`, and assert the resulting display-list path. This is cheap and closes the
hole that allowed the defect.

### ~~TOP OF QUEUE: `RelLineTo` looks wrong, and it is the most common row type~~ — `8b11077`

**Evidence.** `crates/vsdx-resolve/src/geometry.rs` realizes `RelLineTo` as
`current + xy` — turtle-graphics, an offset from the previous point. The corpus says
otherwise. This canonical run appears throughout `lichtsysteme/visio/pages/page1.xml`:

```xml
<Row T='MoveTo'    IX='0'><Cell N='X' V='0'/><Cell N='Y' V='0'/></Row>
<Row T='RelLineTo' IX='2'><Cell N='X' V='1'/><Cell N='Y' V='0'/></Row>
<Row T='RelLineTo' IX='3'><Cell N='X' V='1'/><Cell N='Y' V='1'/></Row>
<Row T='RelLineTo' IX='4'><Cell N='X' V='0'/><Cell N='Y' V='1'/></Row>
```

Those are the four corners of a rectangle in **normalized shape coordinates** — Visio's
`RelLineTo` takes fractions of `Width`/`Height`, per MS-VSDX. Realized as offsets they
instead trace (0,0) → (1,0) → (2,1) → (2,2), a staircase. Supporting statistic: of 2126
`RelLineTo` `X` values in one page, 87% lie in [0,1], which fits fractions and not offsets.

**Frequency: 2404 occurrences — the single most common geometry row type in the corpus.**

**Why nothing caught it.** Geometry has no `@V` oracle; nothing in the corpus validates
realized paths. The render tests encode the same assumption, so they agree with the bug.
This is the geometry analogue of the `V='Themed'` blind spot: a whole subsystem with no
independent ground truth.

**Before fixing, confirm there is no downstream compensation** — a shape transform applied
later could in principle absorb the difference. Check the realized path against the display
list for a known rectangle shape end to end. If nothing compensates, this is a real and
long-standing rendering defect in the most-used path, and it outranks everything else in
this file.

**`RelMoveTo` (601 occurrences) uses the same `current + xy` form and is likely wrong the
same way.**

Note that the `PolylineTo` merged at `8b11077` deliberately follows the existing
`RelLineTo` convention, because the dispatch prompt told it to rather than invent a third
one. If `RelLineTo` is wrong then `PolylineTo` inherits the same error — consistently, and
fixing the convention fixes both.

### Geometry: `PolylineTo` and `InfiniteLine` realized — `8b11077`

Both row types now realize instead of dropping the segment. `SplineStart`, `SplineKnot` and
`NURBSTo` remain skipped, as scoped.

Full gate run and seen passing on the merged result, corpus set: rust core (22/26/52/**50**,
resolve up from 41), wasm-gated (55/54/1), `fmt`, `clippy -D warnings`, wasm build, TS (96),
web (27), demo (7), 13 typechecks. Corpus measurement unmoved — `evaluated=3679`, coverage
52.62%, `@V` 3670/0/0 — as expected, since neither row type occurs in the corpus.
Re-running `scripts/create-vsdx-fixture.ts` regenerates every fixture byte-identically, so
determinism holds.

**Two disclosed limitations — do not describe this as complete support:**

1. **`InfiniteLine` clips to the unit box `[0,1]²`, not the real shape bounds.**
   `realize_geometry(&ResolvedSection)` has no access to shape `Width`/`Height`, and that
   signature is pinned by read-only callers (`vsdx-render/src/lib.rs:653`). Exact only for
   unit-sized shapes; finite but geometrically wrong otherwise.
2. **`PolylineTo` relative flags follow the existing `RelLineTo` convention** rather than
   the spec's ×`Width`/`Height`, for the same reason — see the entry above.

Both limitations share one root cause: geometry realization cannot see shape bounds.
Threading bounds into `realize_geometry` would fix the `RelLineTo` question, the
`InfiniteLine` clip and the `PolylineTo` flags together, and is the natural shape of that
work.

**No public copy change was due**: `apps/web/public/llms.txt`, `apps/web/app/content.ts` and
`README.md` describe crates at the level of "geometry" and never enumerate row types, so
they neither overclaimed before nor underclaim now. Checked rather than assumed.

### The `@V` oracle is blind to `V='Themed'` cells

All 242 no-arg `THEMEVAL()` cells in the corpus cache the literal sentinel `V='Themed'`,
not a value. `record_oracle` (`vsdx-eval/src/lib.rs:2108`) cannot parse that into a
comparable value, so it increments `oracle_no_cache` and returns **without comparing**.

Therefore an implementation of no-arg `THEMEVAL` would raise `evaluated` by up to 242 —
coverage 52.62% → 56.08% — while `agreement` and `disagreement` stayed at 3670/0 and the
reported agreement rate stayed a reassuring **100.00%**. The headline metric cannot detect
wrong answers in these cells at all.

This is a hole in the loop's own guard ("coverage up without agreement dropping below
100%"). Coverage gains are only meaningful when the newly-evaluated cells are *comparable*.
Any future coverage claim must also report the change in `oracle_no_cache`.

### Cross-sheet references (163) yield no coverage either

- The 27 `ThePage!DrawingScale` all come from one formula in `masters/master3.xml`
  (`<Cell N='Value' V='1' F='ThePage!DrawingScale/ThePage!PageScale'/>`). A master
  definition is not on a page, so it has no `ThePage` context and "unresolved" is the
  **correct** answer — `vsdx-eval/src/lib.rs:1375` already asserts this deliberately.
- The 136 `Sheet.N!Connections.X*` are all inside `PAR(PNT(...))`. The referenced shapes do
  exist on the same page, so resolving `Connections` section cells is feasible — but every
  such formula would then fail on `PAR`/`PNT`, which are glue functions outside the display
  profile. Net coverage gain: zero.

### Superseded: the original claim of a large unmined pool

From the phase-8 measurement (re-measure rather than trusting these numbers):

```
parse_ast_ok=6992  static_known_unsupported=1116
evaluated=3679  unsupported_known=1067  unsupported_other=1756  error=490
@V oracle: agreement=3670  disagreement=0  excluded-stale=9  agreement-rate=100.00%
coverage: 3679/6992 = 52.62%
```

`unsupported_other=1756` plus `error=490` is **2,246 formulas we cannot explain** —
by far the largest known pool of undiscovered work, and unlike geometry it is
corpus-driven, so it ranks properly by frequency.

Next discovery pass should group those two buckets by function name or failure
pattern, emit a ranked table, and queue the top entries. Note `unsupported_known`
(1,067) is the *deliberate* non-goal set documented in `crates/vsdx-eval/README.md` —
do not queue those.

The metric that matters: coverage rising above 52.62% **without** the `@V` agreement
rate falling below 100%. Raising coverage by guessing at semantics would move the
first number and wreck the second.

### The 2,246 unexplained formulas, fully accounted — measured 2026-08-15 at `4ee08e9`

Re-ran the corpus harness (`cargo test -p betteroffice-vsdx-eval --release -- --nocapture
corpus`). Headline numbers are unchanged from phase 8:

```
parse_ast_ok=6992  static_known_unsupported=1116
evaluated=3679  unsupported_known=1067  unsupported_other=1756  error=490
@V oracle: agreement=3670  disagreement=0  excluded-stale=9  agreement-rate=100.00%
coverage: 3679/6992 = 52.62%
```

The harness already histograms both buckets, so no new counter was needed. Grouped by
root cause, the 2,246 account **exactly** — there is no unexplained residue left:

| Rank | Family | Count | % of pool | Nature |
|---|---|---|---|---|
| 1 | Event cells outside the display evaluation profile | 937 | 41.7% | **Not work** — see below |
| 2 | `Inh` inheritance (500 unsupported + 325 error) | 825 | 36.7% | Real gap, corpus-driven |
| 3 | `THEMEVAL` (234 no theme-cell context + 72 no resolvable theme) | 306 | 13.6% | Real gap |
| 4 | Unresolved cross-sheet refs (136 `Sheet.N!Connections.X*` + 27 `ThePage!DrawingScale`) | 163 | 7.3% | Real gap |
| 5 | `TheText` requires phase-4b text layout | 12 | 0.5% | Blocked on text layout |
| 6 | `unit/dimension error: incompatible units` | 2 | 0.1% | Possible real bug — investigate |
| 7 | `string values are not display numbers` | 1 | <0.1% | Noise |

**Finding 1 — the pool is 42% smaller than it looks.** The single largest bucket, 937
event cells, is the *deliberate* non-goal set: actions and events are on the permanently
excluded list. `crates/vsdx-eval/src/lib.rs:347` returns
`unsupported("event cell is outside the display evaluation profile")`, and because that
reason string is not in the known list it lands in `unsupported_other` rather than
`unsupported_known`. This inflates the "unexplained" pool by 937 and makes the 52.62%
coverage denominator dishonest — it counts formulas we have decided never to evaluate.
Reclassifying them is a measurement-honesty fix, not a capability change, and it must not
be presented as a coverage win. Queued below.

**Finding 2 — the top real gap is a wiring gap, not missing semantics.** The 825 `Inh`
formulas fail because the corpus harness evaluates document, style, page and master
*sheet-level* cells against a raw unresolved sheet (`sheet_references(sheet)`,
`crates/vsdx-eval/src/lib.rs:1794`) which has no inheritance chain at all. The page-*shape*
path at line 1842 uses the resolver and resolves `Inh` correctly. Corroborating count:
`visio/document.xml` carries 249 `F='Inh'` cells in each of the two corpus files (498
total), against a bucket of 500.

Compounding it, `Resolver::resolve_sheet` (`crates/vsdx-resolve/src/inheritance.rs:62`)
hardcodes `line_style/fill_style/text_style: None` on the synthetic shape it builds, so a
StyleSheet's own style parent is dropped and style-to-style chains are never walked. The
corpus stylesheets really do chain — `ID='1' LineStyle='3'`, `ID='3' LineStyle='6'` —
and `Sheet` has no typed style fields, so those attributes live in `other_attrs`.

This looked like the ideal shape of work: an existing inheritance mechanism simply not
wired into one path. **That prediction was wrong** — see the measured result below.

### Correction: wiring the host raised coverage by zero — measured 2026-08-15 at `00dc323`

The fix above was implemented and merged. Measured before and after, same corpus:

| | before (`4ee08e9`) | after (`00dc323`) |
|---|---|---|
| `evaluated` | 3679 | **3679** |
| coverage | 52.62% | **52.62%** |
| `@V` agreement / disagreement | 3670 / 0 | 3670 / 0 |
| `Inh requires an inheritance host` | 500 | 44 |
| `Inh has no concrete inherited value` | 325 | **781** |

Exactly 456 formulas moved from one unexplained bucket to the other and **not one became
evaluated**. Wiring the inheritance host was necessary but not sufficient: those cells now
reach the root of their chain, and the root still says `Inh`.

The change was kept — `resolve_sheet` dropping a sheet's own style parent was a real
resolver defect that also affected the render and save projections, the diagnosis it now
reports is accurate rather than misleading, and it added three tests. But it must not be
recorded as a coverage win, because it was not one. The lesson is narrower than "don't
guess": the *root cause* was correctly identified and the fix was still worthless for the
metric, because the chain terminates somewhere nobody had looked.

### What the residual 456 actually need — evidence from the corpus

Every one of the 249 `F='Inh'` cells in each corpus `document.xml` carries a cached `@V`,
and the split is **identical in both otherwise-unrelated documents**:

| Cached `V` | count per file | meaning |
|---|---|---|
| `Themed` | 93 | value comes from the **theme**, not from a style parent |
| `0` | 107 | Visio built-in ShapeSheet default |
| `2`, `1`, `100`, `0.05555555555555555` (=1/18) | 29 | built-in defaults |
| `de-DE` / `en-US` | 5 | **document locale — not a constant** |

So the residual is two separate causes, not one:

1. **~186 theme-terminated cells** (93 × 2). Same root cause as queue item 3 — sheet-level
   cells have no theme context. Combined with the 306 `THEMEVAL` failures this is one
   ~492-formula piece of work, and should be tackled as one.
2. **~312 built-in-default-terminated cells** (156 × 2). When inheritance is exhausted
   Visio falls back to a hardcoded per-cell ShapeSheet default. That the same values appear
   in two unrelated documents is the evidence they are document-independent constants —
   **except the 5 locale cells, which differ between the two files (`de-DE` vs `en-US`) and
   must be read from the document, not hardcoded.** Anyone implementing this must not treat
   the whole set as constants.

Both are oracle-validated: the cached `@V` on these very cells is the ground truth, so an
implementation can be checked rather than guessed. That is what makes them safe to queue.

### Dispatch environment — the auto-reject is `/tmp`, not just the corpus

Broader than the entry below, and it cost a second lost run on 2026-08-15.

opencode auto-rejects `external_directory` access to `C:\Users\Heinrich\AppData\Local\Temp\*`
and **terminates the run with exit 0 and no work done**. On this machine Git-Bash `/tmp`
resolves into exactly that tree, so **any agent that touches `/tmp` for any reason dies
silently** — the corpus has nothing to do with it.

The run lost this way was extracting a *committed repo fixture*
(`unzip crates/vsdx-parse/tests/fixtures/geometry-polyline-and-infinite-line.vsdx -d /tmp/fx1`),
which is an entirely legitimate thing to want to do.

Every dispatch prompt must therefore say: never use `/tmp`; if scratch space is needed,
create it inside the worktree and delete it before committing. Stating "do not read the
corpus" is not sufficient.

### Dispatch environment — opencode auto-rejects the corpus path

opencode refuses `external_directory` access to `C:\Users\Heinrich\AppData\Local\Temp\*`
and **terminates the run with exit 0 and no work done** — a silent failure that looks
like success. The first geometry dispatch was lost to this. Agents must not be pointed
at `VSDX_CORPUS_DIR`; the orchestrator surveys the corpus read-only and inlines the
evidence into the agent prompt instead.

---

## Permanently excluded — do not implement

Per `openspec/vsdx-plan.md`: NURBS *editing* (preservation only), data graphics, data
record sets, containers/lists, validation rules, actions/events, arbitrary stencil
fidelity.

**VBA is permanently excluded and not subject to review** — macro-bearing packages are
rejected by document kind. That is a security property, not a missing feature.

---

## Closed

| Date | What | Commit |
|---|---|---|
| 2026-08-15 | Phases 0–10 closed; VSDX live as the fourth format | `2f0e7fc` |
| 2026-08-15 | Cell caches recomputed or dropped when formulas change | `048b63e` |
| 2026-08-15 | Four vsdx workspaces added to `bun.lock` — CI `--frozen-lockfile` was broken | `f21b4bf` |
