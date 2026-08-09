# betteroffice-vsdx-eval

Bounded ShapeSheet evaluator for the display baseline. It evaluates pure,
display-critical formulas with resolution-aware references, inheritance, package
themes and mutation policy checks. Unsupported formulas never fall back to cached
values.

## Supported formula profile

The evaluator supports bounded parsing and evaluation of numeric arithmetic,
comparisons, references, conditional and boolean expressions, common numeric and
trigonometric functions, units, `GUARD`, `SETATREF`, RGB colours and documented
colour transforms. It resolves shape, page, document and sheet references through
the VSDX resolver, and evaluates `THEMEVAL` when the required host context and
theme are present.

## Explicit non-goals

The following residual categories are intentionally outside this profile:

- Event cells (937): event and recalculation plumbing is not display evaluation.
- `Inh` in raw catalog sheets (500): catalog sheets have no inheritance graph, so
  resolving these values would manufacture results.
- Missing `DocLangID` (325): locale-sensitive evaluation is not attempted without
  the required document language context.
- `THEMEVAL` without host context or a theme (306): theme values require both to be
  meaningful.
- `SHADE` and `LUMDIFF`: their Visio semantics are undocumented, so they remain
  unsupported rather than guessed.

The corpus harness compares evaluated formulas with the cell `@V` cache. Those
values are produced by Visio, but may be stale; the reported agreement rate is an
imperfect compatibility signal, not proof of exact Visio compatibility.
