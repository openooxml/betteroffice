# BetterOffice desktop editors

The macOS Docs, Sheets, and Slides apps embed the same React editors shipped by
`apps/demo`, including their toolbars, menus, dialogs, status bars, sheet tabs,
slide thumbnails, editing commands, and collaboration presence. The macOS shell
provides file dialogs, recent files, Finder opening, drag and drop, Save and
Save As, clipboard menus, printing, and protection for unsaved changes.

The start screen opens without a sample document. Editor JavaScript, WebAssembly,
and fonts are bundled in the app and served over a private loopback connection;
opening and editing files work offline. Collaboration connects only when the
user joins a room. Everyone in a room must open the same document.

Build the bundled editor from the repository root:

```sh
apps/native-viewer/macos/build-editor.sh
```

Run the UI in a browser for development:

```sh
bun run --filter @betteroffice/desktop dev
```

Use `?format=xlsx` or `?format=pptx` to preview the other apps. Browser previews
use browser file selection and downloads. Test the actual macOS host with:

```sh
swiftc -parse-as-library \
  apps/native-viewer/macos/DesktopServer.swift \
  apps/native-viewer/macos/Desktop.swift \
  -o /tmp/betteroffice-desktop
/tmp/betteroffice-desktop --format docx --resources "$PWD/apps/desktop/dist"
```

Validation:

```sh
bun run --filter @betteroffice/desktop typecheck
bun run --filter @betteroffice/desktop test
apps/native-viewer/macos/test-desktop.sh
bun apps/desktop/tests/smoke.ts
```

On a Mac with a desktop session, run `apps/native-viewer/macos/test-desktop.sh --ui`
to launch each native app, check its start screen, edit a temporary copy of the
demo file, save it through the native bridge, and verify the saved contents.

The shell requires macOS 13.3 or later. The Rust/Vello viewer remains available
through `cargo run --manifest-path apps/native-viewer/Cargo.toml` for rendering
diagnostics.
