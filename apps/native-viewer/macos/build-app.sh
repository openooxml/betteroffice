#!/usr/bin/env bash

set -euo pipefail

if [[ $# -ne 4 ]]; then
    echo "usage: $0 FORMAT OUTPUT_DIR VERSION BUILD_VERSION" >&2
    exit 2
fi

format=$1
output_dir=$2
version=$3
build_version=$4

case "$format" in
    docx)
        app_name="BetterOffice Docs"
        bundle_id=dev.betteroffice.docs
        extension=docx
        icon_variant=doc
        type_name="Word document"
        content_type=org.openxmlformats.wordprocessingml.document
        welcome=betteroffice-demo.docx
        ;;
    xlsx)
        app_name="BetterOffice Sheets"
        bundle_id=dev.betteroffice.sheets
        extension=xlsx
        icon_variant=sheet
        type_name="Excel workbook"
        content_type=org.openxmlformats.spreadsheetml.sheet
        welcome=sample.xlsx
        ;;
    pptx)
        app_name="BetterOffice Slides"
        bundle_id=dev.betteroffice.slides
        extension=pptx
        icon_variant=deck
        type_name="PowerPoint presentation"
        content_type=org.openxmlformats.presentationml.presentation
        welcome=betteroffice-demo.pptx
        ;;
    *)
        echo "FORMAT must be docx, xlsx, or pptx" >&2
        exit 2
        ;;
esac

script_dir=$(cd "$(dirname "$0")" && pwd)
repo_root=$(cd "$script_dir/../../.." && pwd)
target_dir=${CARGO_TARGET_DIR:-"$repo_root/apps/native-viewer/target"}
deployment_target=${MACOSX_DEPLOYMENT_TARGET:-12.0}
binary_name=betteroffice-native-viewer
app_path="$output_dir/$app_name.app"

if [[ ! $version =~ ^[0-9]+([.][0-9]+){0,2}$ ]]; then
    echo "VERSION must contain one to three numeric components" >&2
    exit 2
fi

if [[ ! $build_version =~ ^[0-9]+([.][0-9]+)*$ ]]; then
    echo "BUILD_VERSION must contain numeric components" >&2
    exit 2
fi

mkdir -p "$output_dir"
rm -rf "$app_path"

export CARGO_TARGET_DIR="$target_dir"
export MACOSX_DEPLOYMENT_TARGET="$deployment_target"

for target in aarch64-apple-darwin x86_64-apple-darwin; do
    cargo build \
        --locked \
        --release \
        --no-default-features \
        --features "$format" \
        --target "$target" \
        --manifest-path "$repo_root/apps/native-viewer/Cargo.toml"
done

mkdir -p "$app_path/Contents/MacOS" "$app_path/Contents/Resources"
xcrun lipo -create \
    "$target_dir/aarch64-apple-darwin/release/$binary_name" \
    "$target_dir/x86_64-apple-darwin/release/$binary_name" \
    -output "$app_path/Contents/MacOS/$binary_name"
chmod 755 "$app_path/Contents/MacOS/$binary_name"

cp "$script_dir/Info.plist" "$app_path/Contents/Info.plist"
plutil -replace CFBundleDisplayName -string "$app_name" "$app_path/Contents/Info.plist"
plutil -replace CFBundleName -string "$app_name" "$app_path/Contents/Info.plist"
plutil -replace CFBundleIdentifier -string "$bundle_id" "$app_path/Contents/Info.plist"
plutil -replace CFBundleShortVersionString -string "$version" "$app_path/Contents/Info.plist"
plutil -replace CFBundleVersion -string "$build_version" "$app_path/Contents/Info.plist"
plutil -replace LSMinimumSystemVersion -string "$deployment_target" "$app_path/Contents/Info.plist"
plutil -replace CFBundleDocumentTypes.0.CFBundleTypeExtensions -json "[\"$extension\"]" "$app_path/Contents/Info.plist"
plutil -replace CFBundleDocumentTypes.0.CFBundleTypeName -string "$type_name" "$app_path/Contents/Info.plist"
plutil -replace CFBundleDocumentTypes.0.LSItemContentTypes -json "[\"$content_type\"]" "$app_path/Contents/Info.plist"

icon_dir=$(mktemp -d "${TMPDIR:-/tmp}/betteroffice-icon.XXXXXX")
trap 'rm -rf "$icon_dir"' EXIT

swift "$script_dir/generate-icon.swift" "$icon_variant" "$icon_dir/icon.png"
icon_set="$icon_dir/AppIcon.iconset"
mkdir "$icon_set"
for spec in \
    icon_16x16.png:16 \
    icon_16x16@2x.png:32 \
    icon_32x32.png:32 \
    icon_32x32@2x.png:64 \
    icon_128x128.png:128 \
    icon_128x128@2x.png:256 \
    icon_256x256.png:256 \
    icon_256x256@2x.png:512 \
    icon_512x512.png:512 \
    icon_512x512@2x.png:1024
do
    name=${spec%%:*}
    pixels=${spec##*:}
    sips -z "$pixels" "$pixels" "$icon_dir/icon.png" --out "$icon_set/$name" >/dev/null
done
swift "$script_dir/package-icns.swift" \
    "$icon_set" "$app_path/Contents/Resources/AppIcon.icns"

cp "$repo_root/apps/demo/public/$welcome" "$app_path/Contents/Resources/Welcome.$extension"
plutil -lint "$app_path/Contents/Info.plist" >/dev/null
xcrun lipo "$app_path/Contents/MacOS/$binary_name" -verify_arch arm64 x86_64
printf '%s\n' "$app_path"
