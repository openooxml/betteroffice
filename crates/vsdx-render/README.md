# betteroffice-vsdx-render

Dynamic connectors resolve their ShapeSheet endpoint values and page glue records at layout time.
The routing policy is deterministic: a single-subpath filed `Geometry` supplies the waypoints,
anchored to the resolved endpoints, otherwise the shape's `ShapeRouteStyle` — or the page's
`RouteStyle` — selects a direct run or one horizontal-first or vertical-first orthogonal bend.
Every 1-D shape takes this path, plain lines included, so a shape that carries no route style at
all runs straight rather than assuming a connector's right angle. An unresolved route becomes a
placeholder instead of a guessed line.

Connector crossings render Visio line jumps at layout time: the more-horizontal leg bridges the
more-vertical one under `LineJumpCode` 1 (mirrored for 2, z-order for 4 and 5), sized by
`LineJumpFactorX/Y` times `LineToLineX/Y`, with per-connector `ConLineJumpCode` overrides.
Jumps draw in the arc and gap styles, and crossing detection is bounded by a candidate and a
placed-jump budget.

VSDX resolved-scene to display-list compiler and hit tester.

Geometry sections keep their individual `IX` identities. Enabled or unevaluable
section-level `NoFill`, `NoLine`, and `NoShow` controls produce an explicit
unsupported placeholder; their paint semantics are not yet implemented. Disabled
controls remain renderable. The original section XML remains preserved on save.
