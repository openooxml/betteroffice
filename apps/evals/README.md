# Evals site

Static per-task results from `eval/results/*.json`, built with Bun and Tailwind CSS v4.
The site shares the benchmarks site's palette, page frame and Geist fonts.

```sh
bun run dev:evals
bun run --filter '@betteroffice/evals' build
bun run --filter '@betteroffice/evals' test
```

The newest suite appears at `/`. Every suite remains available under `/suites/<suite>`.
Format pages show task prompts, verdicts, failure reasons, tool calls and times.
`/failures`, `/method` and `/reproduce` explain the results and how to repeat them.
