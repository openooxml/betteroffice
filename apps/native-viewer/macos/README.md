# macOS app

`build-app.sh FORMAT OUTPUT_DIR VERSION BUILD_VERSION` builds one of `docx`,
`xlsx`, or `pptx` for both architectures, merges it into a universal binary,
and assembles the matching BetterOffice app with its icon and Welcome file.
`package-dmg.sh DOCS_APP SHEETS_APP SLIDES_APP DMG` packages the suite with an
`/Applications` symlink.

The `macOS app` workflow runs for `@betteroffice/docx@<version>` tags, pull
requests and manual dispatches. Every run uploads an artifact suffixed
`-unsigned`. Non-PR runs can pass the `apple-signing` environment gate to also
sign, notarize, and staple all three apps and the disk image, then upload a
signed artifact. Tagged signed runs attach the disk image to that GitHub
release. A prerelease tag such as `@betteroffice/docx@0.2.0-rc.1` uses `0.2.0`
as the numeric macOS bundle version and keeps the full version in artifact
names.

## Secrets

Create an `apple-signing` environment with required reviewers, then create
these as environment secrets. All six are required for the signed path; if any
is missing the workflow retains only the unsigned artifact.

| Secret | What it is |
| --- | --- |
| `APPLE_CERTIFICATE_P12` | Base64 of a Developer ID Application certificate exported as `.p12` |
| `APPLE_CERTIFICATE_PASSWORD` | The password chosen during that export |
| `APPLE_TEAM_ID` | The ten-character team id, shown in the Apple Developer account page |
| `APPLE_API_KEY_ID` | Key id of an App Store Connect API key |
| `APPLE_API_ISSUER_ID` | Issuer id shown above the key list |
| `APPLE_API_KEY_P8` | Base64 of the `.p8` file, which downloads exactly once |

To export the certificate: request a Developer ID Application certificate in
Keychain Access, install it, then right-click it under My Certificates and
export as `.p12`. Encode both it and the `.p8` with `base64 -i FILE | pbcopy`.

An App Store Connect API key is used instead of an app-specific password so no
individual's Apple ID is embedded in CI. Create one under Users and Access >
Integrations with the Developer role.
