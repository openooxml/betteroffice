import json
from pathlib import Path
import shutil
import subprocess
import sys
import tempfile
import xml.etree.ElementTree as ET
from zipfile import ZipFile

scripts = Path(__file__).resolve().parent
root = scripts.parents[3]
resources = root / "apps/desktop/dist"
assert (resources / "index.html").is_file(), "Build the desktop editor first"

with tempfile.TemporaryDirectory(prefix="betteroffice-native-smoke-") as directory:
    temporary = Path(directory)

    def run(format, script, document=None):
        output = temporary / f"{format}-{script}"
        command = [
            sys.argv[1], "--format", format, "--resources", str(resources),
            "--ui-test", str(scripts / f"{script}.js"), str(output),
        ]
        if document:
            command += ["--document", str(document)]
        subprocess.run(command, check=True, timeout=120)
        result = json.loads(output.with_suffix(".json").read_text())
        assert "error" not in result, result
        assert result["result"]["format"] == format, result
        return result["result"]

    for format in ("docx", "xlsx", "pptx"):
        assert run(format, "start-smoke")["start"]
        fixture = "showcase.xlsx" if format == "xlsx" else f"betteroffice-demo.{format}"
        document = temporary / f"native.{format}"
        source = root / "apps/demo/public" / fixture
        shutil.copyfile(source, document)
        assert run(format, "editor-smoke", document)["saved"]
        with ZipFile(document) as saved:
            if format == "docx":
                assert b"Native smoke" in saved.read("word/document.xml")
            elif format == "xlsx":
                sheet = ET.fromstring(saved.read("xl/worksheets/sheet1.xml"))
                cell = sheet.find(".//{*}c[@r='A1']")
                assert cell is not None and "N" in "".join(cell.itertext())
            else:
                slides = lambda archive: len(ET.fromstring(archive.read("ppt/presentation.xml")).findall(".//{*}sldId"))
                with ZipFile(source) as original:
                    assert slides(saved) == slides(original) + 1
        print(f"{format}: native start, edit, save, and saved contents passed", flush=True)
