set -euo pipefail

lo_archive="${RUNNER_TEMP}/libreoffice.tar.gz"
curl --fail --location --retry 3 --max-time 300 --output "$lo_archive" \
  https://downloadarchive.documentfoundation.org/libreoffice/old/26.2.3.2/deb/x86_64/LibreOffice_26.2.3.2_Linux_x86-64_deb.tar.gz
echo "18838cb9d028b664a9d0e966cd4c8ca47ca3ea363c393b41d1b5124740b121a5  $lo_archive" | sha256sum --check
mkdir "${RUNNER_TEMP}/libreoffice-install"
tar -xzf "$lo_archive" -C "${RUNNER_TEMP}/libreoffice-install"
sudo apt-get update
sudo apt-get install -y --no-install-recommends \
  fontconfig libxinerama1 libx11-xcb1 libcairo2 libcups2t64 libdbus-glib-1-2 libsm6 libxrender1 libxext6 libnss3 \
  "${RUNNER_TEMP}"/libreoffice-install/LibreOffice_*/DEBS/*.deb
