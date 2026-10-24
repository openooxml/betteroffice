#!/usr/bin/env python3
"""Builds the public composite DOCX from public-composite.json: downloads each source, checks its sha256, composes.

python3 -I scripts/compare/build-public-composite.py [--out scripts/compare/docs/public-800pp.docx]
Writes the composite and its attribution list (<out stem>.PROVENANCE.txt), then checks the composite's sha256.
"""

import argparse
import hashlib
import json
import os
import subprocess
import sys
import urllib.request

HERE = os.path.dirname(os.path.abspath(__file__))


def sha256(path):
    h = hashlib.sha256()
    with open(path, "rb") as f:
        for chunk in iter(lambda: f.read(1 << 20), b""):
            h.update(chunk)
    return h.hexdigest()


def fetch(url, path):
    request = urllib.request.Request(url, headers={"User-Agent": "betteroffice-compare/1"})
    with urllib.request.urlopen(request, timeout=300) as response, open(path + ".part", "wb") as f:
        while chunk := response.read(1 << 20):
            f.write(chunk)
    os.replace(path + ".part", path)


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--out", default=os.path.join(HERE, "docs", "public-800pp.docx"))
    ap.add_argument("--corpus", default=os.path.join(HERE, "docs", "corpus"))
    a = ap.parse_args()
    with open(os.path.join(HERE, "public-composite.json"), encoding="utf-8") as f:
        spec = json.load(f)
    raw = os.path.join(a.corpus, "raw", "docx")
    os.makedirs(raw, exist_ok=True)
    rows = []
    for source in spec["sources"]:
        path = os.path.join(raw, f"{source['id']}.docx")
        if not os.path.exists(path) or sha256(path) != source["sha256"]:
            print(f"fetch {source['id']}", file=sys.stderr)
            fetch(source["source_url"], path)
        if sha256(path) != source["sha256"]:
            raise SystemExit(f"{source['id']}: sha256 mismatch for {source['source_url']}")
        rows.append({**source, "path": f"raw/docx/{source['id']}.docx"})
    with open(os.path.join(a.corpus, "manifest.jsonl"), "w", encoding="utf-8") as f:
        for row in rows:
            f.write(json.dumps(row, ensure_ascii=False) + "\n")
    os.makedirs(os.path.dirname(os.path.abspath(a.out)), exist_ok=True)
    subprocess.run([sys.executable, "-I", os.path.join(HERE, "compose.py"), "--corpus", a.corpus, "--out", a.out,
                    *(row["id"] for row in rows)], check=True, stdout=subprocess.DEVNULL)
    digest = sha256(a.out)
    print(f"{os.path.basename(a.out)} {os.path.getsize(a.out)} bytes sha256 {digest}")
    if digest != spec["sha256"]:
        raise SystemExit(f"composite sha256 {digest} differs from the measured {spec['sha256']}")


if __name__ == "__main__":
    main()
