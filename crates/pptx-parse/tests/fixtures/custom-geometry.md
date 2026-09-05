`custom-geometry.pptx` is a generated, text-free regression deck licensed with this repository.

Slide 1 contains two targeted shapes and three controls. Slide 2 contains only presets.

| Shape | Before | After |
| --- | --- | --- |
| Mixed paths, box `(20,20,200,100)` px | Red rectangle with blue outline | Four independent paths, with line, cubic, quadratic and close commands |
| Quarter ellipse, box `(260,20,200,100)` px | Green rectangle with blue outline | Elliptical arc from `(440,40)` to `(360,70)` px, then line and close |
| Missing path list / unresolved guide | Amber rectangles | Identical fallback rectangles |
| Preset ellipse | Purple ellipse | Identical preset path |

The mixed shape's first path uses a `200 × 100` coordinate space against a `1905000 × 952500` EMU transform. Its first points normalize to `(0.1,0.1)` and `(0.9,0.1)`, which render at `(40,30)` and `(200,30)` px. Its remaining paths use `400 × 200`, `100 × 50`, and `200 × 100` spaces. Their paints are respectively stroke only, fill only, and neither. The fill is `#DC2626`; the stroke is `#2563EB`, 2 px wide.

The quarter ellipse uses radii `(80,30)` in a `200 × 100` path space and a 90-degree sweep. Its normalized cubic controls are `(0.9,0.3656854249492381)` and `(0.7209138999323174,0.5)`, with endpoint `(0.5,0.5)`.

![Before](../../../pptx-render/tests/fixtures/custom-geometry-before.png)

![After](../../../pptx-render/tests/fixtures/custom-geometry-after.png)

The before image was generated using `origin/main` at `387f2392`.
The version-2 and version-6 update fixtures were regenerated with current
`origin/main` at `54fdaa00c8242d58db61418ac3bc3b2ad6d50cb4`, using client ID 285.
They are stored at `crates/pptx-edit/tests/fixtures/deck-custom-schema-v2.update.bin`
and `deck-custom-schema-v6.update.bin`. The fixture generator uses main’s legacy
parser and defaults before stamping v2; its normal writer produces v6.
Both contain the parsed model without custom paths and exercise migration to
schema 7. Restoring an old update without source bytes retains its historical
fallback geometry; attaching the original deck reparses the custom paths.
See the [generator and compatibility details](../../../pptx-edit/tests/fixtures/README.md).

The parser uses the same normalized command representation and 2,048-command bound as the DOCX custom-geometry parser. PPTX preserves individual path painting and rejects an unsupported path as a whole. It converts DrawingML polar angles to ellipse parameters before producing cubic curves; see [Apache POI's DrawingML angle convention](https://github.com/apache/poi/blob/trunk/poi/src/main/java/org/apache/poi/sl/draw/geom/ArcToCommand.java). Guide formulas still use the rectangle fallback.

Isolation against current main (`54fdaa00`): 14 tracked PPTX fixtures; 40 slides; 39 byte-identical off-target; only the two intended shapes on custom-geometry slide 1 change; 14 snapshot JSON payloads and 13 off-target package JSON payloads byte-identical; 239 no-edit ZIP part payloads byte-identical on both revisions.

Current main’s schema-6 documents load and migrate to schema 7. Main rejects
both fresh and migrated schema-7 documents, including full and differential
updates, without changing the receiving session. Snapshot JSON remains readable
by main. Legacy v2 commits migrations 3, 4, 5, 6, and 7 in separate transactions;
hidden keys first appear at 6 and survive 7. Reopening schema 7 is idempotent.

The schema-7 collision checks mutate skipping migration 7, skipping main’s
migration 6, reverting the schema constant, and accepting schema 8. Each fails
running migration tests and returns to green after restoring production code.

Demo collaboration seeds were regenerated after diagnosing the CI check:

The schema-6 seed loads through schema 7’s two metadata writes (`packageJson`
and `schemaVersion`), adding client 2 at clock 2 and growing the loaded state
vector from 11 to 13 bytes. A fresh schema-7 seed needs no migration and keeps
an 11-byte state vector. Both retain 5 metadata keys, 1,045 shape-map keys,
65 shapes, 55 stories, and 3 slides; their snapshots are exactly equal.
The raw old and fresh PPTX seeds both contain 160,445 bytes and differ only
at offset 40, where the schema byte changes from 6 to 7. This is the expected
schema upgrade, with no dropped data.

Original custom-geometry mutation verification at `79dc4056` (each failed a running test, then passed after restoration):

| Mutation | Test filter | Result |
| --- | --- | --- |
| `normalization` | `a_custom_geometry_becomes_a_path_normalised_to_the_shape` | Red → restored green |
| `cubic-control` | `a_custom_geometry_becomes_a_path_normalised_to_the_shape` | Red → restored green |
| `quadratic-control` | `paths_use_independent_spaces_and_paint_flags` | Red → restored green |
| `arc-fallback` | `arcs_preserve_ellipse_angles_direction_and_current_point` | Red → restored green |
| `ellipse-angle` | `arcs_preserve_ellipse_angles_direction_and_current_point` | Red → restored green |
| `arc-direction` | `arcs_preserve_ellipse_angles_direction_and_current_point` | Red → restored green |
| `close-current` | `arcs_preserve_ellipse_angles_direction_and_current_point` | Red → restored green |
| `path-fill` | `paths_use_independent_spaces_and_paint_flags` | Red → restored green |
| `path-stroke` | `paths_use_independent_spaces_and_paint_flags` | Red → restored green |
| `partial-fallback` | `invalid_paths_fall_back_as_a_whole_and_expansion_is_bounded` | Red → restored green |
| `command-budget` | `invalid_paths_fall_back_as_a_whole_and_expansion_is_bounded` | Red → restored green |
| `preset-parse` | `presets_keep_priority_over_custom_geometry` | Red → restored green |
| `snapshot-dispatch` | `custom_paths_keep_coordinates_paints_and_fallbacks` | Red → restored green |
| `inherited-dispatch` | `layout_and_master_paths_survive_snapshot_hydration` | Red → restored green |
| `render-fill` | `custom_paths_keep_coordinates_paints_and_fallbacks` | Red → restored green |
| `render-stroke` | `custom_paths_keep_coordinates_paints_and_fallbacks` | Red → restored green |
| `render-budget` | `custom_paths_share_the_slide_shape_budget` | Red → restored green |
| `serde-default` | `custom_paths_persist_without_changing_legacy_json_or_source_parts` | Red → restored green |
| `serde-skip` | `custom_paths_persist_without_changing_legacy_json_or_source_parts` | Red → restored green |
| `schema-version` | `released_v` | Red → restored green |
| `schema-migration` | `released_v2_snapshot_migrates_once_without_losing_shapes` | Red → restored green |
| `multiple-paths` | `paths_use_independent_spaces_and_paint_flags` | Red → restored green |
| `preset-render` | `custom_paths_keep_coordinates_paints_and_fallbacks` | Red → restored green |
| `normalized-finite` | `invalid_paths_fall_back_as_a_whole_and_expansion_is_bounded` | Red → restored green |
| `arc-budget` | `invalid_paths_fall_back_as_a_whole_and_expansion_is_bounded` | Red → restored green |
| `guide-fallback` | `a_custom_geometry_becomes_a_path_normalised_to_the_shape` | Red → restored green |
| `future-schema` | `unmigratable_schema_versions_stay_rejected` | Red → restored green |
| `migration-convergence-version` | `two_clients_migrating_the_same_v1_snapshot_converge` | Red → restored green |
