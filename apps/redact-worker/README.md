# Redaction upload worker

The CLI removes sensitive content locally. `POST /upload` receives only that redacted package, validates its declared DOCX/XLSX/PPTX format, sanitizes it again inside the `ooxml-opc` WASM trust boundary, and stores the result under an opaque ID. `GET /f/:id` returns the sanitized package without an original filename.

## R2 setup

Create the bucket once:

```sh
bunx wrangler r2 bucket create betteroffice-repros
```

Then replace the `replace-with-redacted-bucket` placeholder in `wrangler.jsonc` with `betteroffice-repros`. The binding name must remain `REDACTED_BUCKET`. Until that binding is configured, the worker builds and typechecks but returns `503` for storage routes.

## Commands

```sh
bun run build:wasm
bunx tsc --noEmit
bun test
bun run deploy
```

The WASM build writes generated glue beside the worker and an ignored `ooxml_opc_bg.wasm`; the binary must not be committed.
