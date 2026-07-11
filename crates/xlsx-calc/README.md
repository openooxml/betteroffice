# xlsx-calc

The formula engine: lexer → parser → `Expr` AST → tree-walking evaluator, plus
per-formula reference extraction and (from the graph work) dependency tracking.
It reads cells exclusively through `xlsx_model::CellProvider` — see
`docs/decisions/0001-formula-engine.md`.

Written spec-first from ECMA-376 Part 1 §18.17 and public Excel function
semantics. No GPL/AGPL or proprietary spreadsheet source was consulted.

## Architecture

- `lexer.rs` / `parser.rs` — source → position-tagged tokens → `Expr`.
- `eval.rs` — the evaluator and coercion machinery (`to_number`, `to_text`,
  `to_bool`, `cmp_values`), range/`Area` access, and error propagation. It owns
  no functions; `Expr::FuncCall` dispatches into `functions::lookup`.
- `functions/` — the builtin library. `mod.rs` holds the name → implementation
  registry and shared argument collectors; one module per category
  (`math`, `stats`, `text`, `datetime`, `logical`, `lookups`, `info`) plus
  `criteria.rs` (the shared Excel criteria-string parser and the *IF/*IFS
  driver).

Every builtin has the signature `fn(&[Expr], &EvalContext) -> CellValue` and
receives its arguments **unevaluated**, so control-flow functions (`IF`, `IFS`,
`SWITCH`, `IFERROR`, `IFNA`, `CHOOSE`, `AND`, `OR`) evaluate only the branches
they take.

## Registry

Function names resolve **case-insensitively** (`sum`, `SUM`, `Sum` are one
function). Aliases map to a single implementation: `CONCAT`/`CONCATENATE`,
`MODE`/`MODE.SNGL`, `STDEV`/`STDEV.S`, `STDEVP`/`STDEV.P`, `VAR`/`VAR.S`,
`VARP`/`VAR.P`, `RANK`/`RANK.EQ`.

## Supported functions

### Math

| Name | Notes / deviations |
|---|---|
| `SUM` | Numeric cells only from references; literals coerce; errors propagate. |
| `SUMIF(range, criteria, [sum_range])` | `sum_range` is anchored at its top-left with the criteria shape. |
| `SUMIFS(sum_range, crit_range, crit, …)` | All ranges must share dimensions. |
| `SUMPRODUCT(array1, [array2], …)` | Element-wise product summed; non-numeric cells = 0; arrays must match length. |
| `PRODUCT` | No numbers → 0. |
| `ABS`, `SIGN` | — |
| `ROUND` | Half away from zero. |
| `ROUNDUP` / `ROUNDDOWN` | Directional (away from / toward zero). |
| `MROUND(n, m)` | Nearest multiple, half away from zero; opposite signs → `#NUM!`; `m=0` → 0. |
| `CEILING(n, s)` / `FLOOR(n, s)` | Multiple of `s`; positive `n` with negative `s` → `#NUM!`; `s=0` → 0. |
| `INT` | Floor. |
| `TRUNC(n, [digits])` | Toward zero. |
| `MOD(n, d)` | Sign follows divisor; `d=0` → `#DIV/0!`. |
| `POWER`, `SQRT`, `EXP` | Non-finite result → `#NUM!`; `SQRT` of a negative → `#NUM!`. |
| `LN`, `LOG10`, `LOG(n, [base])` | Non-positive input → `#NUM!`; `LN`/`LOG10` use the dedicated libm routine. |
| `PI` | — |

### Statistics

| Name | Notes / deviations |
|---|---|
| `AVERAGE` | No numbers → `#DIV/0!`. |
| `COUNT` / `COUNTA` / `COUNTBLANK` | Numeric / non-empty / empty-or-`""`. |
| `COUNTIF`, `COUNTIFS` | Criteria via the shared parser. |
| `AVERAGEIF(range, criteria, [avg_range])`, `AVERAGEIFS` | No matches → `#DIV/0!`. |
| `MIN` / `MAX` | No numbers → 0. |
| `MEDIAN` | — |
| `MODE` (`MODE.SNGL`) | Earliest-appearing value wins ties; no repeats → `#N/A`. |
| `STDEV`/`STDEV.S`, `STDEVP`/`STDEV.P` | Sample needs ≥2 values, population ≥1, else `#DIV/0!`. |
| `VAR`/`VAR.S`, `VARP`/`VAR.P` | As above. |
| `LARGE(array, k)` / `SMALL(array, k)` | `k` out of range → `#NUM!`. |
| `RANK`/`RANK.EQ(n, ref, [order])` | Order 0/omitted = descending; ties share the best rank; absent → `#N/A`. |

### Text

| Name | Notes / deviations |
|---|---|
| `LEN`, `UPPER`, `LOWER`, `TRIM` | `TRIM` collapses runs of the ASCII space only. |
| `LEFT`/`RIGHT(text, [n])`, `MID(text, start, count)` | 1-based; positions counted in Unicode scalar values (see below). |
| `FIND` / `SEARCH` | `FIND` case-sensitive, `SEARCH` case-insensitive; not found → `#VALUE!`. **No wildcards in `SEARCH`.** |
| `SUBSTITUTE(text, old, new, [instance])` | Empty `old` returns the text unchanged. |
| `REPLACE(old, start, count, new)` | Positional. |
| `REPT`, `EXACT`, `PROPER`, `CLEAN` | `CLEAN` strips control characters. |
| `T` | Text passes through, everything else → `""`. |
| `CHAR(n)` / `CODE(text)` | `CHAR` for code points 1..=255; `CODE` returns the first char's code point (Unicode, not a code page). |
| `VALUE`, `NUMBERVALUE(text, [dec], [grp])` | `VALUE` handles a trailing `%`. |
| `TEXT(value, format)` | **Minimal**: only `0`, `0.00`, `#,##0`, `#,##0.00`, `0%`; any other code → `#VALUE!`. |
| `TEXTJOIN(delim, ignore_empty, …)` | `ignore_empty` also skips empty strings; ranges flatten row-major. |
| `CONCAT` / `CONCATENATE` | Ranges flatten row-major. |

### Date & time

| Name | Notes / deviations |
|---|---|
| `DATE(y, m, d)` | Months roll into years, day is an offset (overflow rolls); years 0..=1899 → `1900+y`; result < 1 → `#NUM!`. |
| `YEAR` / `MONTH` / `DAY` | Serial < 0 → `#NUM!`; serial 0 renders as 1900-01-00. |
| `WEEKDAY(serial, [type])` | Types 1 (default), 2, 3, and 11..17. |
| `EDATE` / `EOMONTH(start, months)` | `EDATE` clamps the day to the target month length. |
| `TODAY` / `NOW` | Read `ctx.now_serial`; **absent clock → `#VALUE!`** (documented pure-engine boundary). |
| `HOUR` / `MINUTE` / `SECOND` | Time portion rounded to the nearest second. |
| `TIME(h, m, s)` | Fraction of a day in [0, 1); values beyond a day wrap. |
| `DATEDIF(start, end, unit)` | Units `Y`, `M`, `D`, `YM`, `YD`, `MD`; `end < start` → `#NUM!`. |

Serial ↔ calendar math is the Excel **1900 system including the deliberate leap
bug** (serial 60 = the phantom 1900-02-29), matching `xlsx_model::date`. The
workbook date system is not reachable through `CellProvider`, so the 1904 epoch
is not yet wired — a follow-up.

### Logical

| Name | Notes |
|---|---|
| `IF(cond, then, [else])` | Omitted else → `FALSE`. |
| `IFERROR(v, fallback)` | Fallback replaces any error. |
| `IFNA(v, fallback)` | Fallback replaces only `#N/A`. |
| `IFS(cond, val, …)` | First true condition; none → `#N/A`. |
| `SWITCH(expr, match, result, …, [default])` | Trailing odd argument is the default; no match/default → `#N/A`. |
| `AND` / `OR` / `NOT` / `XOR` | Blanks/text ignored; at least one logical required. |

### Lookup & reference

| Name | Notes / deviations |
|---|---|
| `VLOOKUP` / `HLOOKUP(value, table, index, [range_lookup])` | `range_lookup` defaults to TRUE (approximate on a sorted first column/row); index out of range → `#REF!`. **No wildcards in exact mode.** |
| `MATCH(value, area, [type])` | Types 1 (default, ascending), 0 (exact), -1 (descending). **No wildcards in type 0.** |
| `INDEX(area, row, [col])` | Single-row/column areas accept one index; out of range → `#REF!`. |
| `XLOOKUP(value, lookup, return, [if_not_found], …)` | **Exact match only**; match/search modes beyond exact are not yet implemented. |
| `CHOOSE(index, …)` | Only the chosen argument is evaluated. |
| `ROW` / `COLUMN([ref])` | **A reference is required** — the evaluator has no notion of the calling cell, so the no-arg form is `#VALUE!`. |
| `ROWS` / `COLUMNS(area)` | Dimension counts. |

### Information

| Name | Notes |
|---|---|
| `ISBLANK`, `ISNUMBER`, `ISTEXT`, `ISLOGICAL` | Type predicates; never propagate errors. |
| `ISERROR`, `ISERR`, `ISNA` | `ISERR` excludes `#N/A`. |
| `NA` | The `#N/A` literal. |
| `N` | Numbers/bools → numbers, text → 0, errors pass through. |

## Criteria strings

Shared by the *IF/*IFS family (`criteria.rs`). A criterion is an optional
leading comparison (`>=`, `<=`, `<>`, `>`, `<`, `=`) followed by a value; with
no operator the comparison is equality. Numeric-looking values compare as
numbers; everything else compares as case-insensitive text. For `=`/`<>` the
text may contain the wildcards `*` (any run) and `?` (any single character),
with `~` escaping a literal `*`, `?`, or `~`.

## Deviations & boundaries (summary)

- **Text positions** are counted in Unicode scalar values; Excel counts UTF-16
  code units. This differs only for astral (supplementary-plane) characters.
- **`SEARCH` / VLOOKUP / MATCH / VLOOKUP exact mode** do not implement
  wildcards yet.
- **`TEXT`** implements only the five format codes listed above; the full
  §18.8.31 number-format interpreter is a separate PR.
- **1904 date system** is not yet wired (see Date & time).
- **`RAND` / `RANDBETWEEN`** are intentionally **not implemented** here — the
  engine is kept pure and deterministic; volatility is handled generically by
  the dependency graph.
- **`TODAY` / `NOW`** return `#VALUE!` when no clock is injected via
  `EvalContext::with_now`.
