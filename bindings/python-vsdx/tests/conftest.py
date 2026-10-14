from pathlib import Path

import pytest

ROOT = Path(__file__).resolve().parents[3]

@pytest.fixture(scope="session")
def foundation_path() -> Path:
    return ROOT / "crates" / "vsdx-parse" / "tests" / "fixtures" / "foundation.vsdx"
