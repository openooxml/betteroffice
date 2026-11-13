set -euo pipefail

source "$(dirname "${BASH_SOURCE[0]}")/libreoffice-pin.sh"
lo_archive="$LO_ARCHIVE_PATH"
echo "$LO_SHA256  $lo_archive" | sha256sum --check
mkdir "${RUNNER_TEMP}/libreoffice-install"
tar -xzf "$lo_archive" -C "${RUNNER_TEMP}/libreoffice-install"
sudo apt-get update
sudo apt-get install -y --no-install-recommends \
  fontconfig libxinerama1 libx11-xcb1 libcairo2 libcups2t64 libdbus-glib-1-2 libsm6 libxrender1 libxext6 libnss3 \
  "${RUNNER_TEMP}"/libreoffice-install/LibreOffice_*/DEBS/*.deb
