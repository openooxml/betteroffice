# VSDX engine plan — an open-source Microsoft Visio alternative

> Reviewed by an independent high-tier codex agent against the live repo and the MS-VSDX spec; 14 must-fix corrections applied. Provenance of every format claim is tagged `[spec]`, `[fixture]` (must be confirmed against a real file) or `[oracle]` (desktop-Visio behaviour).

## Context

BetterOffice ships three Rust-native OOXML engines — DOCX, XLSX, PPTX — each with the same vertical shape: a `*-parse` crate over `ooxml-opc`, a `*-edit` crate holding a `yrs` CRDT, a `*-render` crate emitting a display list, a `*-wasm` boundary, a `@betteroffice/*` TS package, a `*-react` editor, i18n, optional Python bindings, and a demo route. There is no fourth format, and **zero** Visio code exists anywhere (verified: `\bvisio\b|vsdx|vsdm|vstx|VisioDocument` returns no matches).

This plan adds `.vsdx` as that fourth format: a real, editable, collaborative open-source Visio alternative, not a viewer.

Locked scope decisions:

1. **Full PPTX parity** — parse + CRDT edit + render + wasm + `@betteroffice/vsdx` + `-react` + `-i18n` + Python bindings + collaboration + demo route.
2. **Full ShapeSheet recalculation engine** — Visio's `Cell@F` formulas are first-class, not metadata.
3. **No pre-refactor** — new `vsdx-*` crates reuse the shared layer where it fits and duplicate the rest, exactly as PPTX did.

**Honest sizing, stated once.** The independent review puts this at roughly **95–190 person-weeks** — a multi-person-year programme, not a feature. Phase 7 (ShapeSheet compatibility) alone is 26–52+ weeks and is realistically its own programme rather than one phase of this one. The plan below delivers the full locked scope, sequenced so every phase is independently shippable and useful on its own. If the timeline ever needs to compress, the review's recommended cut order is: broad formula compatibility first, then editable stencil fidelity — shipping a lossless inspector/viewer plus narrowly editable native shapes before the "Visio alternative" claim is made publicly.

---

## Part 1 — What .vsdx actually is

The reverse-engineering result. Every design choice below follows from it. Claims are tagged; `[fixture]` items **must be confirmed in phase 0 and this section corrected where real files disagree.**

### Corpus verification status

Two real Visio-produced files (a 4-master single-page drawing with media, and a 3-master two-page drawing) were inspected in phase 0. **Every structural claim below was confirmed**, plus three findings that were not anticipated:

1. **Attribute quoting is producer-dependent.** Both files use single-quoted XML attributes (`FromSheet='818'`). `quick-xml` handles this transparently, but the **lexical span patcher (Part 6) must never assume a quote character** when rewriting a cell in place — it must reuse the span's original quoting.
2. **`Row T=` is not Geometry-exclusive.** `Row T='Connection'` appears inside `Section N='Connection'`. The row-type attribute is broader than the Geometry section; do not gate `T` parsing on section name.
3. **Formulas are pervasive, not occasional.** 6,206 cells carry `F=` across just these two files. This confirms `F` is first-class and validates rejecting any model that stores only resolved rectangles.

Observed geometry row types: `RelLineTo` (2404), `RelMoveTo` (601), `LineTo` (430), `Connection` (115), `ArcTo` (92), `MoveTo` (54), `EllipticalArcTo` (2). Observed sections: `Geometry`, `Character`, `Control`, `Connection`, `User`, `Layer`, `Scratch`, `Actions`. `Del='1'` occurs 22 times — the deleted-state modelling is load-bearing, not theoretical. Named rows (`Row N='TextPosition'`, `'visVersion'`, `'msvThemeOrder'`, …) coexist with indexed rows, as specified.

The corpus is **not committed** — the source files are the user's own work diagrams. They live outside the repo as a local dev corpus with recorded provenance; committed fixtures will be generated once the real structure is understood. **No `.vsdm` is available**, so the kind-rejection test must use a synthetic package built by rewriting the content type.

### Package graph `[verified]`

`.vsdx` is an OPC zip whose payload is **Visio XML, not DrawingML**. There is no `p:sp`, no `a:prstGeom`, no OOXML shape tree.

```
[Content_Types].xml
_rels/.rels                        → visio/document.xml
visio/document.xml                 ← contains VisioDocument/StyleSheets,
visio/_rels/document.xml.rels        DocumentSheet, colors, fonts INLINE
  ├─ visio/pages/pages.xml         → _rels → page1.xml, page2.xml, …
  ├─ visio/masters/masters.xml     → _rels → master1.xml, …
  ├─ visio/theme/theme1.xml        (optional)
  ├─ visio/windows.xml             (optional)
  ├─ visio/dataRecordSets.xml      (optional)
  ├─ visio/comments.xml            (optional)
  └─ visio/extensions/, solutions/, media, custom XML, thumbnails
```

- **There is no `visio/styles.xml`.** Style sheets live inside `visio/document.xml` under `VisioDocument/StyleSheets`, alongside `DocumentSheet`, `Colors` and `FaceNames`. A plan that reaches for a separate styles part will not find one.
- Root relationship type is the Visio family (`http://schemas.microsoft.com/visio/2010/relationships/document`), **not** `officeDocument`. `[verified]` — exact in both corpus files. Note `docProps/*` still use the standard OOXML relationship types, so the root `.rels` is mixed-family.
- `visio/theme/theme1.xml` carries the **standard** OOXML theme content type (`application/vnd.openxmlformats-officedocument.theme+xml`) `[verified]` — so `ooxml_drawingml::Theme` is a genuinely appropriate target, even though the XML parser for it still has to be written.
- **Resolve every part through relationships, never by assumed path.** Producers vary.
- Content types distinguish `application/vnd.ms-visio.drawing.main+xml` from `…macroEnabled.main+xml` (and stencil/template variants).

### Format-kind policy `[spec]`

Detection must return a **document kind**, not the string `"vsdx"`. Structural fallback on `visio/document.xml` alone would silently accept a `.vsdm` as a `.vsdx`. The rule: content type is authoritative; the VSDX API accepts only the non-macro drawing content type; VSDM/VSSX/VSTX are recognised and **rejected by kind** even when the structural fallback matches. Tests must cover contradictory content types, missing content types, and dangling relationship targets.

### The ShapeSheet model `[spec]`

Every Visio shape is a spreadsheet.

```xml
<PageContents>
  <Shapes>
    <Shape ID="1" NameU="Process" Type="Shape"
           Master="2" MasterShape="5"
           LineStyle="3" FillStyle="3" TextStyle="3">
      <Cell N="PinX"    V="4.25"/>                    <!-- singleton: direct child -->
      <Cell N="LocPinX" F="Width*0.5"  V="1"/>
      <Section N="Geometry" IX="0">
        <Row T="RelMoveTo" IX="1"><Cell N="X" F="Width*0" V="0"/>…</Row>
      </Section>
      <Section N="Character">
        <Row IX="0"><Cell N="Size" V="0.1666"/></Row>
      </Section>
      <Section N="User">
        <Row N="visVersion"><Cell N="Value" V="15"/></Row>   <!-- named row -->
      </Section>
      <Text>Step <fld IX="0"/> of <fld IX="1"/></Text>
    </Shape>
  </Shapes>
  <Connects>
    <Connect FromSheet="7" FromCell="BeginX" FromPart="9"
             ToSheet="1"  ToCell="Connections.X1" ToPart="100"/>
  </Connects>
</PageContents>
```

Load-bearing facts, corrected against the spec:

- **`F` and `V` are not symmetric.** `F` is the expression *when present*; `V` is its stored result — a cache. Neither is guaranteed to exist on a given cell. Both must survive round-trip untouched. Editing `V` while discarding `F` produces a file Visio recalculates into something else on open. This remains the single most important invariant.
- **Section/Row/Cell is *not* a universal container.** Singleton cells (`PinX`, `Width`, `Angle`, …) are **direct children of `Shape`**. Only sectioned data uses `Section`/`Row`.
- **Rows are keyed by `N` *or* `IX`, not always `IX`.** Named rows (`User`, `Property`, `Actions`) use `N`; indexed rows use `IX`. `T` (RowType) is **specific to Geometry rows**, not general.
- **`Del` is meaningful at shape, row and cell level** and must be modelled at each. The four-state provenance — absent / inherited / locally overridden / explicitly deleted — must be distinguishable or round-trip breaks.
- **Geometry row types** `[spec]`: `MoveTo`, `LineTo`, `ArcTo`, `EllipticalArcTo`, `Ellipse`, `InfiniteLine`, `NURBSTo`, `PolylineTo`, `SplineStart`, `SplineKnot`, and the relative variants `RelMoveTo`, `RelLineTo`, `RelCubBezTo`, `RelQuadBezTo`. A renderer may support a subset; a round-tripper must preserve **all** of them verbatim, plus unknown types.
- **`<Text>` is mixed content.** Literal text interleaves with `cp` (character run), `pp` (paragraph), `tp` (tab) and `fld` (field) markers, each carrying a **row index** into the corresponding `Character`/`Paragraph`/`Tabs`/`Field` section — which may itself be inherited. An ordered text-token model is required, with tests for marker-only, interleaved, inherited and field-bearing text.
- **`Connects` is a dual model, not geometry.** A `Connect` record carries `FromSheet`/`ToSheet` plus *optional* `FromCell`/`ToCell` **and** `FromPart`/`ToPart`. Glue is only correct if both the `Connect` records **and** the endpoint/connection-point ShapeSheet formulas are preserved together. Flattening a connector to a line loses the glue and the diagram stops being a diagram.
- **Master inheritance** flows through `Shape@Master` and `@MasterShape`; local cells override inherited ones.
- **Style inheritance is a second axis**, independently sliced into `LineStyle` / `FillStyle` / `TextStyle`, referencing StyleSheet records in `document.xml` that themselves inherit style-to-style. **The full resolution order is not asserted here** — the naive "shape → style → master → default" chain is too simple. Document sheet, page sheet, master, style and local deletion all participate. **A tested resolution algorithm derived from the spec and fixtures is a phase-3 deliverable**, not an assumption.
- Roughly **200 built-in cell names** exist. This must be a string-keyed map with typed accessors for supported cells — **never a closed Rust enum**. Unknown cells must survive round-trip.
- **`GUARD`, `SETATREF`, `SETATREFEXPR`, `SETATREFEVAL` are not ordinary pure functions.** They are edit-interception semantics: they redirect or protect what a user gesture is allowed to write. They belong to the mutation policy, tested separately, not to the evaluator's function table.

### Where VSDX diverges from PPTX

| | VSDX | PPTX |
|---|---|---|
| Geometry | ShapeSheet cells + Geometry rows | `a:prstGeom` / `a:custGeom` |
| Units | inches (Visio units) | EMU |
| Origin | bottom-left, **Y grows up** | top-left, Y down |
| Placement | Pin + LocPin + Angle, formula-driven | offset/extent + transform |
| Computation | dependency-graph recalc engine | small local guide evaluation |
| Connectors | 1D shapes, glue records + formulas, routing | ordinary shapes |
| Inheritance | master-shape + stylesheet, two axes | master/layout + placeholder |
| Layers | native layer table + `LayerMember` | none |

`ooxml_drawingml::GeometryPathCommand` (`crates/ooxml-drawingml/src/shape.rs:68`) has only `Move`/`Line`/`Quad`/`Cubic`/`Close`. A fine **render target**; it cannot represent NURBS, relative commands or formula-bound geometry. It is not the storage model.

---

## Part 2 — Ownership and scope boundary

Rust is authoritative for: OPC access, Visio XML parsing and saving, the ShapeSheet model and evaluator, the `yrs` diagram, edit operations and undo, text shaping, page layout, display-list emission, and hit testing. TypeScript only loads wasm, decodes the boundary, replays display lists, exposes public types, and supplies React chrome.

**Untouched — read, never modify:**

```
crates/docx-*/**        crates/xlsx-*/**        crates/pptx-*/**
crates/betteroffice-{docx,xlsx,pptx}/**
packages/{docx,xlsx,pptx}*/**
bindings/python-{xlsx,pptx}/**
apps/relay/src/**
```

**Cross-cutting files in scope, with justification.** There is no single-file exception — a fourth format is structurally cross-cutting. Every file outside the new `vsdx-*` tree that this plan touches is listed here and nowhere else:

| File | Why in scope |
|---|---|
| `crates/ooxml-opc/src/sanitize.rs:13,56,584` | `matches!(expected_format, "docx"\|"xlsx"\|"pptx")` and `detect_format()` reject a `.vsdx` before any other code runs. Additive only. Must return a document **kind** (see Part 1). |
| `Cargo.toml`, `scripts/rust-crates.mjs`, `package.json`, `.changeset/config.json`, `.github/workflows/ci.yml` | Build, release-train and CI registration. Purely additive. |
| `apps/demo/**` | New `/vsdx` route, fixture, seed, format registry entry. |
| `apps/web/**`, `apps/docs/**`, `README.md`, `AGENTS.md` | Phase 10 only, when VSDX is genuinely live. |

**Redaction is explicitly a separate, separately-approved follow-up.** `ooxml-redact`, `ooxml-redact-cli` and `apps/redact-worker` all carry closed three-format enumerations. Adding VSDX redaction means implementing and testing a Visio redaction policy — a real deliverable, not a checklist tick. This plan **does not** claim VSDX redaction support; those files stay unchanged and are marked intentionally-unchanged in the audit (Part 8).

**Sanitizer hazard, unproven — must be tested, not assumed.** `sanitize_package` rewrites every XML part, strips comments and PIs, and applies generic field/formula neutralization. It does not currently appear to delete `<fld>` specifically, so the "it will eat Visio fields" worry is a hypothesis. Phase 1 must test actual VSDX field and formula survival through the sanitizer and, if needed, add a VSDX-specific sanitizer policy. Do not assume generic sanitization is lossless just because `unzip_parts` is generic.

---

## Part 3 — Reuse map

| Area | Reuse as-is | VSDX-specific work |
|---|---|---|
| **OPC container** | `unzip_parts` / `rezip_parts`, the 512 MiB inflate budget, 5000-entry cap, traversal + duplicate-path rejection (`crates/ooxml-opc/src/lib.rs:22,63,110`) — genuinely format-neutral. | VSDX/VSDM kind detection; a VSDX sanitizer policy proven by test, not assumed. |
| **Relationships** | Nothing — `ooxml-opc` has no relationship graph. | Own `relationships.rs`, shaped after `crates/pptx-parse/src/relationships.rs`. Visio rel-type constants. All part resolution via relationships. |
| **Bounded XML** | Nothing directly. | Port the `ParseLimits`/`ParseBudget` pattern from `crates/pptx-parse/src/xml.rs` verbatim (DTD + entity rejection, depth/event/attribute budgets threaded `&mut` package-wide), adding `max_cells`, `max_sections`, `max_rows`, `max_shapes`, `max_formula_depth`. |
| **Theme & color** | Only the neutral `Theme` representation (`ooxml-drawingml/src/theme.rs:134`) and the final-RGB helpers. | **More new work than it looks.** `ooxml-drawingml` has **no generic Office-theme XML parser**, and its `ColorValue` (`color.rs:7`) is DrawingML-shaped. Visio theme parsing, palette/color-index resolution, and `QuickStyle*` cell resolution are all new. |
| **Geometry** | `GeometryPathCommand`, `Transform2D`, `Point2D`, `Size2D` as the **render target only**. | `preset_geometry_to_path` (`geometry.rs:8`) is a DrawingML preset whitelist, **not** a Visio stencil library — mapping stencil names onto it produces visually plausible, semantically wrong diagrams. Visio geometry realization is new. |
| **Formula evaluator** | Nothing reusable as grammar. | `crates/docx-parse/src/drawingml.rs:276,540,604` is a **small, non-general DrawingML guide evaluator** (`shape.rs:386` merely calls into it). Reuse its **defensive design** — bounded passes, finite-value checks, explicit unsupported ops, hostile-input rejection — and nothing else. |
| **Text** | `ooxml_text::{FontStore, shape, break_opportunities, bidi_paragraphs, single_line_box, CompatFlags}` — the neutral half. | **Avoid `ooxml-text::measure/*` and `word_metrics`** — Word `settings.xml` compat flags and twip semantics. Visio text-block layout (margins, vertical align, direction, fields, tab stops, fitting, inches→CSS) is new. |
| **Parsing/saving** | The `pptx-parse` structure as **precedent**: guarded entry reads, ordered opaque `parts` with `#[serde(skip)]`, and the byte-exact untouched-round-trip test at `crates/pptx-parse/src/package.rs:571`. (`PackagePart`/`part_bytes`/`replace_part` live in `pptx-parse/src/model.rs`.) | The persistence design in Part 6 — a lexical span patcher, not generic reserialization. |
| **Editing** | `pptx-edit` patterns: `DeckSession`/`EditCtx`/typed receipts, deterministic `BOOTSTRAP_CLIENT_ID` seeding, staged-clone `apply_update_v1` validation, local-origin-only undo, `MAX_UPDATE_BYTES`. | Different roots: pages, sheets, cells, formulas, glue. `pptx-edit`'s EMU `ShapeRect` model is incompatible — do not port it. |
| **Display list** | JSON, as `pptx-render`/`xlsx-render` do. (`FrameDelta` is a DOCX resident-pagination optimisation, not an inter-format default.) | `VsdxDisplayList` with `CONTRACT_VERSION: u32 = 1`, **golden contract fixtures**, version rejection, stable primitive IDs, explicit transform/clip/z-order/hit-test conventions, and a size budget. Media and large documents cross as bytes, not JSON. |
| **Wasm** | The `pptx-wasm` facade shape. `scripts/build-pptx-wasm.ts` copied verbatim with names swapped (pins wasm-pack `0.15.0`, rewrites the glue's implicit-URL fallback into a throw). | `scripts/build-vsdx-wasm.ts`. |
| **Collaboration** | `apps/relay/src` is format-agnostic — free. `shared/collaboration-limits.ts`. | Fourth copy of `protocol.ts` + `provider.ts` + `presence.ts`, ported from `packages/pptx/src/collaboration/`. Note `apps/relay/test/limits.test.ts:3-10` imports all three protocols and needs a fourth line. |
| **React / demo** | `apps/demo` shell, `formats.ts` registry, `DemoStage`, `apps/demo/app/collab/*`. | Diagram canvas, page navigator, stencil picker, connection-point hit testing, connector drag, fixture, `/vsdx` route. |

---

## Part 4 — Decomposition

```
crates/
  vsdx-parse        betteroffice-vsdx-parse    OPC graph, bounded XML, LOSSLESS sheet model
                                               (cells/sections/rows/Del/order/unknowns),
                                               pages/masters/stylesheets/theme,
                                               opaque part preservation, lexical save
  vsdx-resolve      betteroffice-vsdx-resolve  the RESOLVED view: two inheritance axes,
                                               units, dependency graph, formula parser +
                                               evaluator, geometry realization
  vsdx-edit         betteroffice-vsdx-edit     yrs diagram, ops, receipts, undo, glue ops, wasm feature
  vsdx-render       betteroffice-vsdx-render   resolved scene → display list + hit testing
  vsdx-wasm         (publish = false)          wasm boundary
  betteroffice-vsdx                            native facade (Diagram / Page / Shape / Connector)

packages/
  vsdx        @betteroffice/vsdx          wasm loader, types, canvas replay, collaboration
  vsdx-react  @betteroffice/vsdx-react    diagram editor
  vsdx-i18n   @betteroffice/vsdx-i18n     locale JSON

bindings/
  python-vsdx  betteroffice-vsdx (PyO3, abi3-py39, maturin)
```

**The crate boundary is "lossless storage vs resolved view", not "Visio vs DrawingML."** An earlier draft split out a `vsdx-shapesheet` crate purely because ShapeSheet differs from DrawingML; the review correctly rejected that — parser, serializer, resolver, evaluator, renderer and editor all need the *same* lossless sheet representation, so splitting it out creates ownership and cycle pressure. `vsdx-parse` therefore owns the package graph, lexical preservation **and** the sheet model. `vsdx-resolve` owns everything that computes over it. If a third consumer later needs a parser-independent semantic API, extract then.

Naming follows the workspace convention exactly: short alias in `[workspace.dependencies]`, long `package = "betteroffice-vsdx-*"`, `[lib] name = "vsdx_*"`, `publish = ["crates-io"]`, a `README.md`.

---

## Part 5 — Phases

Resequenced per review: **resolution and evaluation come before editing, and save ships atomically with the first edit capability.** Editing formulas before you can interpret `GUARD`/`SETATREF` means a drag can silently overwrite a protected or redirected formula.

Every phase ends at its gate (Part 9) and a local commit. **No phase pushes. No phase opens a PR.**

**0 — Corpus and oracle harness.** No `.vsdx` exists in the repo. Build a **committed** license-cleared fixture suite (small files, in git, so CI can gate on them) plus expected structural/render/formula artifacts: basic shapes; flowchart with dynamic connectors; multi-page; master-heavy stencil use; groups; images; themed; named rows; `Del` at shape/row/cell; text with `cp`/`pp`/`tp`/`fld`; glue with `FromPart`/`ToPart`; unsupported geometry (NURBS, splines); malformed zip/XML/relationship cases; a VSDM that must be rejected by kind. Large or restricted originals stay external with immutable provenance and a documented acquisition process. Add `scripts/create-demo-diagram.ts` producing a deterministic branded fixture (fixed zip date, mirroring `scripts/create-demo-deck.ts`). **Then correct Part 1 wherever fixtures disagree with it.**

**1 — Foundation and lossless read/write.** Extend `sanitize.rs` to return a document kind; prove the three existing formats are byte-identical through the sanitizer. `vsdx-parse`: `ParseLimits`/`ParseBudget`, guarded `parse_xml`, relationship-driven part graph, ordered opaque parts. **Ship the byte-exact untouched-round-trip test first.** Test VSDX field/formula survival through `sanitize_package`.

**2 — Lossless sheet model.** `Cell { formula: Option<String>, value: Option<String>, unit: Option<Unit>, del: bool, unknown_attrs }`; singleton cells as direct shape children; sections with rows keyed by `N` **or** `IX`, optional `T`, `Del`, `LocalName`, unknown attributes, and source order preserved. The four-state provenance is representable. Formulas carried untouched. **Round-trip must be lossless here or the phase does not close.**

**3 — Resolution.** `vsdx-resolve`: the tested inheritance algorithm (document sheet, page sheet, master, master-shape, the three independent style slices, local deletion) derived from spec + fixtures — this algorithm is the phase deliverable. Plus text-token resolution and geometry realization. Facade `betteroffice-vsdx` opens and inspects.

**4 — Baseline evaluator + viewer.** A bounded evaluator covering the formula profile the corpus actually needs for correct display (arithmetic, cell references, units, the common geometry expressions), with unsupported formulas falling back to the cached `@V` and being flagged, never silently. `vsdx-render`: inches→CSS px and the Y-up→Y-down flip applied **once, at the final paint transform**. Basic 2D shapes, groups, images, direct line/fill styles, `MoveTo`/`LineTo`/`ArcTo`/`Ellipse`, text via `ooxml-text` shaping. Unsupported geometry renders as a diagnostic placeholder. `VsdxDisplayList` + hit testing + golden contract fixtures.

**5 — Transactional save + constrained editing.** The persistence design (Part 6) and the first edit capability ship together — a save path is not a follow-up to editing. Mutation policy first: `GUARD` blocks, `SETATREF` redirects, protected cells refuse. Then ops: move, resize, set fill/stroke/font, text edit, add rect/ellipse/line/text, z-order, delete. **Ops write formulas through the mutation policy, never raw values into `V`.**

**6 — CRDT and collaboration core.** `vsdx-edit`: yrs roots (`vsdx:meta`, `vsdx:page-order`, `vsdx:pages`, `vsdx:sheets`, `vsdx:stories`), deterministic bootstrap seeding, `EditCtx`, typed receipts, local-origin undo, staged-clone update validation, `MAX_UPDATE_BYTES`. Two-peer convergence tests.

**7 — Glue and routing.** 1D shape model, the dual `Connect`-records-plus-formulas glue model including `FromPart`/`ToPart`, connection points, and a deterministic straight/orthogonal routing policy with rerouting on shape move. Routing fidelity is advertised honestly — a documented limitation beats a stale route.

**8 — Broad ShapeSheet compatibility.** The locked "full recalc engine" target, made measurable. Deliverables: a declared **supported formula profile** with explicit non-goals; preservation behaviour for every unsupported formula; cross-page/master/style references; the unit system; locale vs universal cell naming; error values and propagation; cycle detection; recalc scheduling; and a **Visio-produced expected-results corpus** as the oracle. Source `@V` values are *not* a valid oracle on their own — they can be stale. `GUARD`/`SETATREF*` stay in the phase-5 mutation policy with their own tests, not in the evaluator's function table. Function coverage is driven by measured corpus frequency, and the compatibility percentage against the oracle corpus is the phase's exit metric.

**9 — Productization.** `@betteroffice/vsdx` (wasm loader, types, canvas replay), `@betteroffice/vsdx-react` (diagram canvas, page navigator, stencil picker, connector drag, toolbar), `@betteroffice/vsdx-i18n`, the ported collaboration provider + two-browser proof, `bindings/python-vsdx` (inches and diagram concepts, **not** the EMU `Rect` that `bindings/python-pptx/src/lib.rs:211` documents), `apps/demo/app/vsdx/`, collaboration seed.

**10 — Go live.** Only now: flip `apps/demo/lib/formats.ts` to `status: "live"` and update `apps/web/app/content.ts`, `apps/web/public/llms.txt`, `apps/docs/content/docs/`, `README.md`, `AGENTS.md` scopes. `llms.txt` is a public capability contract — it must not overclaim NURBS editing, data graphics, or Visio-equivalent routing.

**Explicitly out of scope:** NURBS *editing* (preservation only), data graphics, data record sets, containers/lists, validation rules, actions/events, VBA, arbitrary stencil fidelity, and VSDX redaction.

---

## Part 6 — Persistence design

The naive "surgical serializer that rewrites only mutated cells" does not survive contact with the format: `<Text>` is mixed content, and adding or deleting a shape changes `Shapes`, page metadata, ID allocation, `Connects`, and possibly relationships and content types. **Generic XML reserialization cannot preserve unknown XML byte-for-byte.**

The design is a **lexical span patcher over an intact source part**:

- Parsing records a byte span for every shape, section, row, cell and text node.
- The CRDT carries an identity mapping from each entity to its source span.
- A cell-value or formula edit rewrites only that span; every untouched byte of the part is copied through.
- Structural edits (add/delete shape, reorder, new section) take a **documented fallback path** that re-emits the containing element while preserving all descendant spans it did not touch.

The byte-preservation promise is stated precisely, and only this: **unchanged parts are byte-identical; changed parts preserve unmodified lexical spans where feasible.** Not an unconditional unknown-XML guarantee.

---

## Part 7 — Agent policy and execution rules

### Tier split

A hard split, enforced per task:

| Tier | Models | May |
|---|---|---|
| **Writing** | Sonnet subagents, or low-tier `codex exec` (`-c model_reasoning_effort=low\|medium`) | Implement, edit files, write tests, run builds, generate fixtures, mechanical refactors |
| **Reading** | Opus, or high-tier codex (`-c model_reasoning_effort=high`) | Explore, review diffs, assess architecture, give feedback, gate phases, approve designs |

- **All production code is written by low-tier models only.** Sonnet subagents or low/medium-effort codex. The orchestrator does not write engine code itself.
- **All review is done by high-tier models in read-only mode.** A high-tier agent **never edits a file** — no exceptions, not even a typo. Its output is a review, a design, or a verdict. Invoke codex reviewers with `--sandbox read-only`.
- A low-tier agent **never approves its own work.** Every phase gate is reviewed by a high-tier agent that did not write the code. (Role-neutral restatement, so this survives any model change: **author/reviewer separation is mandatory, and phase acceptance is measured against the criteria in Part 9, not against reviewer opinion.**)

### Orchestration

- **Codex first, Claude after.** The orchestrator dispatches every unit of work to codex while quota lasts, and falls back to Claude Code subagents only once codex hits its limit. This applies to both tiers: low-tier codex for writing, high-tier codex for review; Sonnet subagents and Opus review are the fallback path. Orchestration itself stays in Claude Code throughout.
- Codex non-interactive invocation must close stdin (`codex exec … - < prompt.md`), or it blocks forever.
- **Parallelism is encouraged, aggressively.** Fan out as wide as the phase allows: independent parsers, independent cell/function groups, independent TS packages, independent fixture generators. The only hard rule is **never batch across a phase boundary** — a phase gate is a synchronisation point.

### Branching and worktrees

- Work on dedicated branches off `main`, one per phase: `feat/vsdx-<phase>-<slug>` (e.g. `feat/vsdx-01-foundation`).
- **Use git worktrees** whenever parallel agents would otherwise contend for the same checkout — one worktree per concurrently-running agent. This is the default for wide fan-out, not an exception.
- Merge worktree branches back into the phase branch locally; the phase branch is what the high-tier reviewer gates.

### Commits

- **No `git push`. No pull requests.** Local commits and local branches only, until explicitly authorised.
- **No Claude co-author trailer.** Commits carry no `Co-Authored-By: Claude …` line and no generated-by footer.
- **Commit messages in English**, scoped conventional titles per `AGENTS.md`, new scope `vsdx`, imperative and concise: `feat(vsdx): parse page contents`, `test(vsdx): add formula oracle corpus`, `fix(vsdx): preserve deleted rows on save`.
- Every commit names explicit paths. Never `git add -A`.

### Boundaries

- The list in Part 2 is the complete set of files outside `vsdx-*` that may change.
- Plan document lives at `openspec/vsdx-plan.md`, screenshots at `openspec/evidence/vsdx-*.png`, mirroring the PPTX convention.

---

## Part 8 — Closed-enumeration audit

Maintained as a **generated `rg`-backed audit with expected hits**, not a hand list — every site is either updated for VSDX or explicitly marked intentionally-unchanged with a reason.

**Must change (phases 1–9):**
- `crates/ooxml-opc/src/sanitize.rs:13,56,584`
- `crates/ooxml-opc/README.md:3`
- `Cargo.toml:13-31` — `[workspace.dependencies]` aliases
- `scripts/rust-crates.mjs:7` — `RUST_CRATES`, in strict dependency order (`validateRustTrain` fails otherwise)
- `package.json:25-46` — `build:vsdx-wasm`, `build:vsdx`, `build:packages`, and the `pretest`/`test` wasm asymmetry (today `pretest` builds only PPTX wasm; `test` builds only XLSX+DOCX)
- `.changeset/config.json:5` — fourth `fixed` group
- `.github/workflows/ci.yml:40,47` — fixture diff gate + minimal-wasm-features check
- `apps/demo/lib/formats.ts:4,11` — `Format.id` union + array
- `apps/demo/middleware.ts:13` — route matcher
- `apps/demo/package.json:8,22`
- `apps/demo/scripts/build-collaboration-seeds.ts:36-104`, `check-collaboration-seeds.ts:156`
- `apps/relay/test/limits.test.ts:3-10` — imports all three collaboration protocols
- `scripts/python-bindings.mjs` — register `bindings/python-vsdx`

**Must change at phase 10 (product-facing):**
- `apps/web/app/content.ts:15,50`, `apps/web/app/page.tsx:59`, `apps/web/app/layout.tsx:21`
- `apps/web/public/llms.txt:3,11,27,44,57`
- `apps/docs/content/docs/index.mdx:3`, `collaboration.mdx:3,6,46-73`, `apps/docs/app/layout.tsx:18`
- `README.md:23-33`, `AGENTS.md:9`, `RELEASING.md:9,24`

**Intentionally unchanged — VSDX redaction is a separate follow-up (Part 2):**
- `crates/ooxml-redact/src/lib.rs:11,20,109`, `src/xml.rs:136,160,233-314`, `Cargo.toml:21`, `src/tests.rs:35`
- `crates/ooxml-redact-cli/src/main.rs:12,131`, `src/lib.rs:81-87`
- `apps/redact-worker/src/handler.ts:11,26,197`, `test/handler.test.ts:7-11,86-100,149-157`, `test/wasm.test.ts:11`, `README.md:3`

---

## Part 9 — Verification

**Gates are phase-appropriate.** A command is only added to the gate once its target exists — `vsdx-wasm` and `build:vsdx-wasm` do not exist before phase 4.

Always, from phase 1:

```bash
cargo fmt --check
cargo clippy --all-targets -- -D warnings
cargo test
```

From phase 4 (wasm exists):

```bash
cargo check -p vsdx-wasm --target wasm32-unknown-unknown --no-default-features
bun run build:vsdx-wasm
```

From phase 9 (TS packages exist):

```bash
bun run --filter './packages/*' typecheck
bun run test                     # after adding vsdx wasm to pretest/test
bun run build:packages
git diff --exit-code -- apps/demo/public/
```

Release-train and policy, from phase 1:

```bash
node scripts/publish-crates.mjs --dry-run   # runs validateRustTrain (publish-crates.mjs:116) before the dry-run branch
```

`cargo deny` is **not** a repo script — CI runs it via `EmbarkStudios/cargo-deny-action@v2.0.20`. Locally it requires `cargo install cargo-deny`; document the prerequisite rather than pretending `cargo deny check ...` is a project command.

Non-negotiable correctness gates:

- **Round-trip (phase 1+):** for every committed corpus file, `unzip_parts(write(parse(f))) == unzip_parts(f)` on untouched parts.
- **Sanitizer regression (phase 1):** the three shipping formats produce **byte-identical sanitizer output** before and after the `sanitize.rs` change. (Not "parse/render identical" — normal parse and render paths never call `sanitize_package`.)
- **Kind rejection (phase 1):** VSDM/VSSX/VSTX are rejected by kind even when the structural fallback matches.
- **Lossless model (phase 2):** every corpus file survives parse→model→write with `Del`, named rows, unknown cells, unknown attributes and source order intact.
- **Display-list contract (phase 4):** golden fixtures; a bumped `CONTRACT_VERSION` is rejected by the TS decoder.
- **Formula fidelity (phase 8):** recalculated values match the **Visio-produced oracle corpus** within tolerance; the pass rate is the phase's exit metric. Source `@V` is a cross-check, not the oracle.
- **Convergence (phase 6, 9):** two-peer yrs exchange converges; two-browser proof with screenshots into `openspec/evidence/`.
- **End-to-end (phase 9):** `bun run dev:demo` → `/vsdx` → open the fixture → edit a shape → save → **reopen in Microsoft Visio**. LibreOffice Draw is a smoke test, *not* a Visio fidelity proxy.
