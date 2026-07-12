# Contributing to BetterOffice

Thanks for your interest in contributing! This guide will help you get started.

## Prerequisites

- [Bun](https://bun.sh/) (v1.0+)
- [Rust](https://rustup.rs/) (stable, with `rustfmt` and `clippy`)

## Development Setup

```bash
# Clone the repo
git clone https://github.com/openooxml/betteroffice.git
cd betteroffice

# Install dependencies
bun install

# Start the web app
bun run dev
```

## Running Tests

```bash
# TypeScript
bun run typecheck
bun test

# Rust engines (fmt + clippy with -D warnings + tests)
bun run rust:check
```

## Contributor License Agreement

Contributors are required to sign our [Contributor License Agreement](CLA.md). The CLA assistant will leave a comment on your first pull request with signing instructions — one short comment, about 30 seconds. That signature covers all of your future contributions. Contributing as part of your job? Your employer can sign the [Corporate CLA](CCLA.md) instead; the bot comment explains both paths.

## Making Changes

1. **Fork** the repository and create a branch from `main`
2. **Read the code** before modifying it — match the conventions you find
3. **Make your changes** — keep them focused and minimal
4. **Add/update tests** for your changes
5. **Verify** everything works:
   ```bash
   bun run typecheck && bun test && bun run rust:check
   ```
6. **Submit a PR** against `main` — the CLA bot will prompt you on your first one

## Reporting Bugs

Open an issue at [github.com/openooxml/betteroffice/issues](https://github.com/openooxml/betteroffice/issues) with:

- Steps to reproduce
- Expected vs actual behavior
- Attach a repro file if relevant (remove sensitive content first)

## License

By contributing, you agree that your contributions will be licensed under the [Apache License 2.0](LICENSE).
