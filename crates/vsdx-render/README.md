# betteroffice-vsdx-render

Dynamic connectors resolve their ShapeSheet endpoint values and page glue records at layout time.
The routing policy is deterministic: ordinary connectors are straight, while nonzero `RoutStyle`
uses a horizontal-first orthogonal bend. It does not emulate Visio obstacle avoidance, line jumps,
or manually edited route geometry; an unresolved route becomes a placeholder instead of a guessed line.

VSDX resolved-scene to display-list compiler and hit tester.
