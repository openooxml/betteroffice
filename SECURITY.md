# Security Policy

## Reporting a Vulnerability

Report vulnerabilities privately via [GitHub Security Advisories](https://github.com/openooxml/betteroffice/security/advisories/new). Please do not open a public issue.

We aim to acknowledge reports within 5 business days and will keep you updated on the fix and disclosure timeline.

## Audit exemptions

`bun audit --audit-level=high` gates every branch. An advisory is exempted only
when it cannot be reached in shipped code and no upgrade path exists that does
not cause a worse problem. Each exemption is listed here with its reason, and
`bun audit` still reports it unfiltered so it stays visible.

### GHSA-5p4m-2wfm-xmqj — js-yaml quadratic CPU on `!!omap`

Reached only through `@changesets/cli › @manypkg/get-packages › read-yaml-file`,
which parses the changeset files in this repository. It is release tooling, is
never shipped to users, and never reads untrusted YAML.

`read-yaml-file@1.1.0` calls `yaml.safeLoad`, removed in js-yaml 4, so it cannot
move off the affected 3.x line. Bun resolves `overrides` globally and supports
neither nested nor path-scoped forms, so the only reachable alternatives are to
pin js-yaml 3.15.1 for every consumer — putting the packages that expect 4.x and
5.x onto js-yaml 3's unsafe-by-default loader — or to pin 5.x and break
changesets outright. Both are worse than the advisory.

Remove this exemption when `@manypkg/get-packages` drops `read-yaml-file`, or
when Bun supports scoping an override to one dependency path.

### GHSA-vfj7-8cjw-p6xm: braces stack exhaustion on deeply nested patterns

Reached only through build and dev tooling: `tailwindcss` and `shadcn` in the
private demo app (via `fast-glob › micromatch`), `fumadocs-mdx` in the private
docs site and `tsup` in every package's devDependencies (via `chokidar`), and
`@changesets/cli` (via `micromatch`). These expand glob patterns from this
repository's own configuration and never see user input. No published package
depends on braces at runtime.

No patched release exists: every braces version up to the latest, 3.0.3, is
affected.

Remove this exemption when a patched braces release is published.
