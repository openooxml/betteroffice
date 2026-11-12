LO_VERSION=26.2.3.2
LO_SHA256=18838cb9d028b664a9d0e966cd4c8ca47ca3ea363c393b41d1b5124740b121a5
LO_ARCHIVE="LibreOffice_${LO_VERSION}_Linux_x86-64_deb.tar.gz"
LO_ARCHIVE_PATH="$HOME/.cache/libreoffice-archive/$LO_ARCHIVE"
LO_MIRRORS=(
  "https://download.documentfoundation.org/libreoffice/stable/${LO_VERSION%.*}/deb/x86_64/$LO_ARCHIVE"
  "https://downloadarchive.documentfoundation.org/libreoffice/old/$LO_VERSION/deb/x86_64/$LO_ARCHIVE"
)
