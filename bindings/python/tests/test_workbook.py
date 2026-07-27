import pytest

import betteroffice_xlsx as bo

PNG_MAGIC = b"\x89PNG\r\n\x1a\n"


def test_open_exposes_sheets(sample_bytes):
    wb = bo.Workbook.open(sample_bytes)
    assert wb.sheet_names == ["Budget", "Summary", "Styled"]
    assert wb.sheet_count == len(wb) == 3
    assert [sheet.name for sheet in wb] == wb.sheet_names


def test_open_path(sample_path):
    assert bo.Workbook.open_path(sample_path).sheet_count == 3


def test_sheet_lookup_by_name_and_index(sample_bytes):
    wb = bo.Workbook.open(sample_bytes)
    assert wb["Budget"].index == 0
    assert wb[0].name == "Budget"
    assert wb.sheet("Summary").index == 1


def test_value_and_formula_are_separate(sample_bytes):
    sheet = bo.Workbook.open(sample_bytes)["Budget"]
    assert sheet.formula("D3") == "B3+C3"
    assert sheet["D3"] == pytest.approx(157.0)
    assert sheet.formula("A1") is None
    assert sheet["A1"] == "Quarterly Budget Report"


def test_recalculated_open_agrees_with_authored_values(sample_bytes):
    cached = bo.Workbook.open(sample_bytes)
    fresh = bo.Workbook.open_recalculated(sample_bytes)
    for address in ("D3", "D4", "D5", "D6", "D7"):
        assert fresh.value("Budget", address) == pytest.approx(
            cached.value("Budget", address)
        )


def test_recalculate_reports_no_drift(sample_bytes):
    summary = bo.Workbook.open(sample_bytes).recalculate()
    assert summary.changed == 0
    assert summary.cycles == 0
    assert summary.limited == 0


def test_editing_an_input_recalculates_dependents(sample_bytes):
    wb = bo.Workbook.open(sample_bytes)
    sheet = wb["Budget"]
    addend = sheet["C3"]

    sheet["B3"] = 1000
    assert sheet["B3"] == pytest.approx(1000.0)
    assert sheet["D3"] == pytest.approx(1000.0 + addend)


def test_writing_a_formula_evaluates_it(sample_bytes):
    wb = bo.Workbook.open(sample_bytes)
    sheet = wb["Budget"]
    column_total = sum(sheet[f"B{row}"] for row in range(3, 11))

    sheet["H1"] = "=SUM(B3:B10)"
    assert sheet.formula("H1") == "SUM(B3:B10)"
    assert sheet["H1"] == pytest.approx(column_total)


def test_set_reports_whether_anything_changed(sample_bytes):
    wb = bo.Workbook.open(sample_bytes)
    assert wb.set("Budget", "B3", "1000") is True
    assert wb.set("Budget", "B3", "1000") is False


def test_value_types(sample_bytes):
    wb = bo.Workbook.open(sample_bytes)
    wb.set("Budget", "H6", "TRUE")
    wb.set("Budget", "H7", "hello")
    assert wb.value("Budget", "H6") is True
    assert wb.value("Budget", "H7") == "hello"
    assert wb.value("Budget", "Z99") is None


def test_error_values_are_typed(sample_bytes):
    wb = bo.Workbook.open(sample_bytes)
    wb.set("Budget", "H5", "=1/0")
    value = wb.value("Budget", "H5")
    assert isinstance(value, bo.CellError)
    assert value.code == "#DIV/0!"
    assert value == "#DIV/0!"
    assert str(value) == "#DIV/0!"


def test_numeric_and_text_coercion(sample_bytes):
    from decimal import Decimal

    wb = bo.Workbook.open(sample_bytes)
    sheet = wb["Budget"]
    sheet["H10"] = 7
    sheet["H11"] = 7.5
    sheet["H12"] = Decimal("2.25")
    sheet["H13"] = None
    assert sheet["H10"] == pytest.approx(7.0)
    assert sheet["H11"] == pytest.approx(7.5)
    assert sheet["H12"] == pytest.approx(2.25)
    assert sheet["H13"] is None


def test_typed_input_is_interpreted_like_excel(sample_bytes):
    sheet = bo.Workbook.open(sample_bytes)["Budget"]
    cases = {
        "=1+1": (2.0, "1+1"),
        "'=1+1": ("=1+1", None),
        "1e3": (1000.0, None),
        "TRUE": (True, None),
        "3.14.15": ("3.14.15", None),
        "inf": ("inf", None),
    }
    for index, (typed, (value, formula)) in enumerate(cases.items()):
        address = f"J{index + 1}"
        sheet[address] = typed
        assert sheet[address] == value, typed
        assert sheet.formula(address) == formula, typed


def test_dates_are_rejected_rather_than_silently_stringified(sample_bytes):
    import datetime

    sheet = bo.Workbook.open(sample_bytes)["Budget"]
    for value in (
        datetime.date(2026, 7, 26),
        datetime.datetime(2026, 7, 26, 12, 0),
        datetime.time(12, 0),
    ):
        with pytest.raises(TypeError, match="not supported yet"):
            sheet["H14"] = value


def test_non_finite_numbers_are_rejected(sample_bytes):
    from decimal import Decimal

    sheet = bo.Workbook.open(sample_bytes)["Budget"]
    for value in (float("nan"), float("inf"), float("-inf"), Decimal("Infinity")):
        with pytest.raises(ValueError, match="finite"):
            sheet["H14"] = value
    with pytest.raises(ValueError):
        sheet["H14"] = 10**400


def test_values_whose_text_form_is_not_a_number_are_rejected(sample_bytes):
    from fractions import Fraction

    sheet = bo.Workbook.open(sample_bytes)["Budget"]
    with pytest.raises(ValueError, match="serializes to"):
        sheet["M1"] = Fraction(1, 3)
    sheet["M1"] = Fraction(3, 1)
    assert sheet["M1"] == pytest.approx(3.0)


def test_untouched_formula_keeps_the_files_cached_value(sample_bytes):
    """open() trusts the cache; only recalculation re-evaluates."""
    import io
    import re
    import zipfile

    source = bo.Workbook.open(sample_bytes).save()
    out = io.BytesIO()
    with zipfile.ZipFile(io.BytesIO(source)) as src, zipfile.ZipFile(
        out, "w", zipfile.ZIP_DEFLATED
    ) as dst:
        for item in src.infolist():
            blob = src.read(item.filename)
            if item.filename.endswith("sheet1.xml"):
                patched, count = re.subn(
                    r'(<c r="D3"[^>]*>\s*<f>[^<]*</f>\s*<v>)[^<]*(</v>)',
                    r"\g<1>999\g<2>",
                    blob.decode(),
                )
                if count:
                    blob = patched.encode()
            dst.writestr(item, blob)
    stale = out.getvalue()

    assert bo.Workbook.open(stale).value("Budget", "D3") == pytest.approx(999.0)
    assert bo.Workbook.open_recalculated(stale).value("Budget", "D3") == pytest.approx(
        157.0
    )


def test_bool_is_not_a_sheet_index(sample_bytes):
    wb = bo.Workbook.open(sample_bytes)
    for value in (True, False):
        with pytest.raises(TypeError, match="not bool"):
            wb.sheet(value)


def test_out_of_range_int_keys_raise_index_error(sample_bytes):
    wb = bo.Workbook.open(sample_bytes)
    for value in (-1, 99, 10**100):
        with pytest.raises(IndexError):
            wb.sheet(value)


def test_missing_file_raises_file_not_found(tmp_path):
    missing = tmp_path / "nope.xlsx"
    with pytest.raises(FileNotFoundError) as caught:
        bo.Workbook.open_path(missing)
    assert caught.value.errno is not None
    assert caught.value.filename == str(missing)


def test_unwritable_save_path_raises_oserror(sample_bytes, tmp_path):
    target = tmp_path / "missing-dir" / "out.xlsx"
    with pytest.raises(OSError) as caught:
        bo.Workbook.open(sample_bytes).save_path(target)
    assert caught.value.filename == str(target)


def test_declared_version_matches_runtime(sample_bytes):
    import importlib.metadata as metadata

    assert metadata.version("betteroffice-xlsx") == bo.__version__


def test_error_value_hash_matches_its_string(sample_bytes):
    """Equal objects must hash equally, including against a plain str."""
    wb = bo.Workbook.open(sample_bytes)
    wb.set("Budget", "H5", "=1/0")
    value = wb.value("Budget", "H5")

    assert hash(value) == hash("#DIV/0!")
    assert {value: "found"}["#DIV/0!"] == "found"
    assert len({value, "#DIV/0!"}) == 1


def test_workbook_is_usable_from_another_thread(sample_bytes):
    """Not `unsendable`: worker threads must not panic."""
    import threading

    wb = bo.Workbook.open(sample_bytes)
    results: list = []

    def work() -> None:
        wb.set("Budget", "B3", "1000")
        results.append(wb.value("Budget", "D3"))

    thread = threading.Thread(target=work)
    thread.start()
    thread.join()

    assert results == [pytest.approx(wb.value("Budget", "D3"))]


def test_concurrent_renders_are_consistent(sample_bytes):
    from concurrent.futures import ThreadPoolExecutor

    wb = bo.Workbook.open(sample_bytes)
    with ThreadPoolExecutor(max_workers=4) as pool:
        results = list(pool.map(lambda _: wb.render_png("Budget", range="A1:D12"), range(8)))

    reference = results[0]
    assert all(png.bytes == reference.bytes for png in results)
    assert all(png.width == reference.width for png in results)


def test_concurrent_opens_are_consistent(sample_bytes):
    from concurrent.futures import ThreadPoolExecutor

    with ThreadPoolExecutor(max_workers=4) as pool:
        books = list(pool.map(lambda _: bo.Workbook.open(sample_bytes), range(8)))

    assert all(wb.sheet_names == ["Budget", "Summary", "Styled"] for wb in books)
    assert all(wb.value("Budget", "D3") == pytest.approx(157.0) for wb in books)


def test_render_png(sample_bytes):
    png = bo.Workbook.open(sample_bytes).render_png("Budget", range="A1:D12")
    assert png.bytes[:8] == PNG_MAGIC
    assert png.width > 0 and png.height > 0
    assert len(png) == len(png.bytes)


def test_render_png_scale_changes_dimensions(sample_bytes):
    wb = bo.Workbook.open(sample_bytes)
    small = wb.render_png("Budget", range="A1:D12", scale=1.0)
    large = wb.render_png("Budget", range="A1:D12", scale=2.0)
    assert large.width > small.width
    assert large.height > small.height


def test_render_png_write(sample_bytes, tmp_path):
    target = tmp_path / "sheet.png"
    bo.Workbook.open(sample_bytes).render_png("Budget", range="A1:D12").write(target)
    assert target.read_bytes()[:8] == PNG_MAGIC


def test_save_round_trip_preserves_edits(sample_bytes):
    wb = bo.Workbook.open(sample_bytes)
    wb.set("Budget", "B3", "1000")
    expected = wb.value("Budget", "D3")

    reopened = bo.Workbook.open(wb.save())
    assert reopened.sheet_names == wb.sheet_names
    assert reopened.value("Budget", "B3") == pytest.approx(1000.0)
    assert reopened.value("Budget", "D3") == pytest.approx(expected)


def test_save_path(sample_bytes, tmp_path):
    target = tmp_path / "out.xlsx"
    bo.Workbook.open(sample_bytes).save_path(target)
    assert bo.Workbook.open_path(target).sheet_count == 3


def test_unknown_sheet_name_raises_key_error(sample_bytes):
    with pytest.raises(KeyError):
        bo.Workbook.open(sample_bytes)["Nope"]


def test_sheet_index_out_of_range_raises_index_error(sample_bytes):
    with pytest.raises(IndexError):
        bo.Workbook.open(sample_bytes)[99]


def test_bad_sheet_key_type_raises_type_error(sample_bytes):
    with pytest.raises(TypeError):
        bo.Workbook.open(sample_bytes).value(1.5, "A1")


def test_malformed_address_raises_range_error(sample_bytes):
    with pytest.raises(bo.RangeError):
        bo.Workbook.open(sample_bytes).value(0, "NOTACELL")


def test_malformed_range_raises_range_error(sample_bytes):
    with pytest.raises(bo.RangeError):
        bo.Workbook.open(sample_bytes).render_png(0, range="ZZ")


def test_garbage_bytes_raise_parse_error():
    with pytest.raises(bo.ParseError):
        bo.Workbook.open(b"not a zip")


def test_error_hierarchy_rolls_up_to_xlsx_error():
    for subclass in (bo.ParseError, bo.RangeError, bo.RenderError):
        assert issubclass(subclass, bo.XlsxError)
    with pytest.raises(bo.XlsxError):
        bo.Workbook.open(b"not a zip")


def test_version_is_exposed():
    assert bo.__version__.count(".") == 2
