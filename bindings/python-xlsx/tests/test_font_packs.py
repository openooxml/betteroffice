from pathlib import Path

import pytest

from betteroffice_xlsx import _betteroffice_xlsx as native
from conftest import run_without_font_packs


def test_the_installed_font_packs_are_the_default_font_source():
    root, packs = native._font_packs()
    assert packs == ["latin", "intl", "cjk"]
    for pack in packs:
        assert [path.name for path in (Path(root) / pack).iterdir()]
        assert list((Path(root) / pack).glob("*/manifest.json"))


@pytest.mark.without_font_packs
def test_without_the_font_pack_wheels_no_pack_is_configured():
    assert native._font_packs() is None


def test_the_tests_without_font_packs_pass_in_a_child_process():
    result = run_without_font_packs()
    assert result.returncode == 0, result.stdout[-4000:] + result.stderr[-4000:]
    assert " passed" in result.stdout
