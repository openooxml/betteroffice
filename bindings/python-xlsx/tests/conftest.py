import os
import subprocess
import sys

WITHOUT_FONT_PACKS = os.environ.get("BETTEROFFICE_TEST_WITHOUT_FONT_PACKS") == "1"
if WITHOUT_FONT_PACKS:
    sys.modules["betteroffice_fonts_packs"] = None

from pathlib import Path

import pytest

FIXTURES = Path(__file__).resolve().parents[3] / "apps" / "demo" / "public"


@pytest.fixture(scope="session")
def sample_path() -> Path:
    path = FIXTURES / "sample.xlsx"
    if not path.exists():
        pytest.skip(f"fixture missing: {path}")
    return path


@pytest.fixture(scope="session")
def sample_bytes(sample_path: Path) -> bytes:
    return sample_path.read_bytes()


def pytest_configure(config):
    config.addinivalue_line(
        "markers",
        "without_font_packs: runs only in the child process that cannot import the font pack wheels",
    )


def pytest_collection_modifyitems(config, items):
    if WITHOUT_FONT_PACKS:
        return
    skip = pytest.mark.skip(reason="runs in test_font_packs.py without the font pack wheels")
    for item in items:
        if "without_font_packs" in item.keywords:
            item.add_marker(skip)


def run_without_font_packs() -> subprocess.CompletedProcess:
    return subprocess.run(
        [sys.executable, "-m", "pytest", "-q", "-p", "no:cacheprovider", "-m", "without_font_packs",
         str(Path(__file__).parent)],
        env={**os.environ, "BETTEROFFICE_TEST_WITHOUT_FONT_PACKS": "1"},
        capture_output=True,
        text=True,
        timeout=900,
    )
