# VSDX synthetic benchmark

Run `bun scripts/vsdx-bench.ts` from the repository root to generate six synthetic
workloads and measure parsing, resolution, evaluation, rendering, and saving.
Each result is a median of five samples after one warm-up.

The renderer has no registered fonts, so its measurements cover fallback text
layout, not font shaping or visual fidelity. These workloads do not replace the
private Visio corpus or Microsoft Visio compatibility checks.

Use `--record` from a clean checkout to replace the baseline with measurements
at the recorded commit. Compare results on the same machine and configuration;
the checked-in historical baseline alone does not establish a performance gain.
