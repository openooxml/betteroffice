#!/usr/bin/env bash
# Install the pinned veraPDF CLI without the interactive IzPack installer.
#
# The release zip contains an IzPack installer jar; its "veraPDF CLI" pack is
# already the runnable fat jar. This script downloads the release, verifies the
# pinned SHA-256, extracts that pack and writes a small launcher, so CI needs
# only curl, unzip and a JRE.
#
# Usage:
#   bash scripts/office-quality/install-verapdf.sh [INSTALL_DIR]
#
# Environment:
#   VERAPDF_VERSION        release to install (default: 1.30.4)
#   VERAPDF_SHA256         expected installer SHA-256 (default: 1.30.4 value)
#   VERAPDF_DOWNLOAD_URL   override the release URL
#   VERAPDF_INSTALL_DIR    default when INSTALL_DIR is omitted
#
# The launcher is INSTALL_DIR/bin/verapdf.

set -euo pipefail

version="${VERAPDF_VERSION:-1.30.4}"
default_sha256="04415ef3f149526a3048b9889723e81320ddf013bd891bcc7f0f297f3eadcc23"
sha256="${VERAPDF_SHA256:-$default_sha256}"
destination="${1:-${VERAPDF_INSTALL_DIR:-}}"

if [ -z "$destination" ]; then
    echo "usage: $0 <install-dir>" >&2
    exit 2
fi
if [ "$version" != "1.30.4" ] && [ "$sha256" = "$default_sha256" ]; then
    echo "install-verapdf: VERAPDF_VERSION=$version needs an explicit VERAPDF_SHA256" >&2
    exit 2
fi
for tool in curl unzip java sha256sum; do
    if ! command -v "$tool" >/dev/null 2>&1; then
        echo "install-verapdf: missing required tool: $tool" >&2
        exit 2
    fi
done

release="${version%.*}"
archive="verapdf-greenfield-${version}-installer.zip"
url="${VERAPDF_DOWNLOAD_URL:-https://software.verapdf.org/rel/${release}/${archive}}"
mkdir -p "$destination"
destination="$(cd "$destination" && pwd)"

temporary="$(mktemp -d)"
trap 'rm -rf "$temporary"' EXIT

echo "install-verapdf: downloading $url"
curl --fail --silent --show-error --location \
    --retry 3 --retry-delay 2 \
    --output "$temporary/$archive" "$url"

actual_sha256="$(sha256sum "$temporary/$archive" | cut -d' ' -f1)"
if [ "$actual_sha256" != "$sha256" ]; then
    echo "install-verapdf: checksum mismatch for $archive" >&2
    echo "  expected $sha256" >&2
    echo "  actual   $actual_sha256" >&2
    exit 1
fi

installer_jar="$temporary/verapdf-greenfield-${version}/verapdf-izpack-installer-${version}.jar"
unzip -q "$temporary/$archive" "verapdf-greenfield-${version}/verapdf-izpack-installer-${version}.jar" -d "$temporary"
test -f "$installer_jar"
cli_jar="$destination/cli-${version}.jar"
unzip -p "$installer_jar" 'resources/packs/pack-veraPDF CLI' > "$cli_jar"
mkdir -p "$destination/bin"

cat > "$destination/bin/verapdf" <<EOF
#!/usr/bin/env bash
set -euo pipefail
here="\$(cd "\$(dirname "\${BASH_SOURCE[0]}")/.." && pwd)"
exec java -cp "\$here/cli-${version}.jar" org.verapdf.apps.GreenfieldCliWrapper "\$@"
EOF
chmod +x "$destination/bin/verapdf"

reported="$("$destination/bin/verapdf" --version 2>&1 | sed -n '1p')"
case "$reported" in
    *"veraPDF $version"*) ;;
    *)
        echo "install-verapdf: launcher reported '$reported', expected veraPDF $version" >&2
        exit 1
        ;;
esac
echo "install-verapdf: veraPDF $version at $destination/bin/verapdf"
