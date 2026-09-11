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
