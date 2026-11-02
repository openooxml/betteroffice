"""Builds the betteroffice-fonts-{latin,intl,cjk} wheels from packages/fonts-<pack>/assets.

Each wheel is pure data: the pack's font pack format 1 files under
betteroffice_fonts_packs/<pack>/<version>/, the layout the Rust font manager reads
from a local base directory. The format bindings depend on all three and point
their default font source at the installed directory.

usage: python3 -I scripts/python_font_packs.py --out DIR [--pack latin|intl|cjk ...] [--version-of PACK]
"""

import argparse
import base64
import hashlib
import json
import sys
import zipfile
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
PACKS = ("latin", "intl", "cjk")
MODULE = "betteroffice_fonts_packs"
SUMMARY = {
    "latin": "Latin, Greek, Cyrillic and Noto RTL faces for the BetterOffice Python bindings.",
    "intl": "Noto faces for international scripts for the BetterOffice Python bindings.",
    "cjk": "Noto CJK faces for the BetterOffice Python bindings.",
}
EPOCH = (1980, 1, 1, 0, 0, 0)


def manifest(pack):
    path = ROOT / "packages" / f"fonts-{pack}" / "assets" / "manifest.json"
    data = json.loads(path.read_text(encoding="utf-8"))
    if data.get("format") != 1 or data.get("pack") != pack or not data.get("version"):
        raise SystemExit(f"{path}: not a format 1 manifest for the {pack} pack")
    return data


def project(pack):
    return f"betteroffice-fonts-{pack}"


def readme(pack, version):
    return f"""# {project(pack)}

{SUMMARY[pack]}

This package holds data only: the {pack} font pack {version} that
`betteroffice-docx`, `betteroffice-pptx` and `betteroffice-xlsx` install as a
dependency and read from disk when they lay out and render documents. Nothing
is downloaded at run time. Install it through those packages rather than on
its own.

The fonts keep their original licenses (SIL Open Font License 1.1), with the
full texts in the wheel's license files.
"""


def metadata(pack, version, licenses):
    lines = [
        "Metadata-Version: 2.4",
        f"Name: {project(pack)}",
        f"Version: {version}",
        f"Summary: {SUMMARY[pack]}",
        "Author: OpenOOXML",
        "License-Expression: Apache-2.0 AND OFL-1.1",
        *[f"License-File: {name}" for name in licenses],
        "Requires-Python: >=3.9",
        "Classifier: Intended Audience :: Developers",
        "Classifier: Programming Language :: Python :: 3",
        "Classifier: Topic :: Text Processing :: Fonts",
        "Project-URL: Homepage, https://betteroffice.dev",
        "Project-URL: Documentation, https://docs.betteroffice.dev/docs/python",
        "Project-URL: Source, https://github.com/openooxml/betteroffice",
        "Description-Content-Type: text/markdown",
        "",
        readme(pack, version),
    ]
    return "\n".join(lines).encode("utf-8")


def record_hash(data):
    digest = base64.urlsafe_b64encode(hashlib.sha256(data).digest()).rstrip(b"=")
    return "sha256=" + digest.decode("ascii")


def build(pack, out):
    data = manifest(pack)
    version = data["version"]
    package = ROOT / "packages" / f"fonts-{pack}"
    assets = package / "assets"
    files = {"manifest.json": (assets / "manifest.json").read_bytes()}
    for face in data["faces"]:
        payload = (assets / face["file"]).read_bytes()
        if len(payload) != face["encodedLength"]:
            raise SystemExit(f"{face['file']}: {len(payload)} bytes, manifest says {face['encodedLength']}")
        files[face["file"]] = payload
    licenses = {"LICENSE": (package / "LICENSE").read_bytes()}
    for path in sorted((package / "LICENSES").glob("OFL-*.txt")):
        licenses[f"LICENSES/{path.name}"] = path.read_bytes()
    if len(licenses) == 1:
        raise SystemExit(f"{package}/LICENSES holds no OFL-*.txt font license")

    dist = f"betteroffice_fonts_{pack}-{version}"
    info = f"{dist}.dist-info"
    entries = [(f"{MODULE}/{pack}/{version}/{name}", files[name]) for name in sorted(files)]
    entries += [(f"{info}/licenses/{name}", licenses[name]) for name in sorted(licenses)]
    entries.append((f"{info}/METADATA", metadata(pack, version, sorted(licenses))))
    entries.append((
        f"{info}/WHEEL",
        b"Wheel-Version: 1.0\nGenerator: betteroffice python_font_packs.py\nRoot-Is-Purelib: true\nTag: py3-none-any\n",
    ))
    record = "".join(f"{name},{record_hash(payload)},{len(payload)}\n" for name, payload in entries)
    entries.append((f"{info}/RECORD", (record + f"{info}/RECORD,,\n").encode("utf-8")))

    out.mkdir(parents=True, exist_ok=True)
    wheel = out / f"{dist}-py3-none-any.whl"
    with zipfile.ZipFile(wheel, "w", zipfile.ZIP_DEFLATED, compresslevel=9) as archive:
        for name, payload in entries:
            item = zipfile.ZipInfo(name, EPOCH)
            item.compress_type = zipfile.ZIP_DEFLATED
            item.external_attr = 0o644 << 16
            archive.writestr(item, payload)
    return wheel


def main(argv):
    parser = argparse.ArgumentParser(description=__doc__.splitlines()[0])
    parser.add_argument("--out", type=Path)
    parser.add_argument("--pack", action="append", choices=PACKS)
    parser.add_argument("--version-of", choices=PACKS, help="print the pack's version and exit")
    args = parser.parse_args(argv)
    if args.version_of:
        print(manifest(args.version_of)["version"])
        return 0
    if args.out is None:
        parser.error("--out is required")
    for pack in args.pack or PACKS:
        wheel = build(pack, args.out)
        print(f"{wheel.name} {wheel.stat().st_size}")
    return 0


if __name__ == "__main__":
    sys.exit(main(sys.argv[1:]))
