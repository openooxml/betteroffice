import io
import re
import xml.sax.saxutils
import zipfile

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


def test_a1_addresses_are_case_insensitive(sample_bytes):
    wb = bo.Workbook.open(sample_bytes)
    sheet = wb["Budget"]
    assert sheet["a1"] == sheet["A1"]
    sheet["h1"] = 42
    assert sheet["H1"] == pytest.approx(42.0)
    assert wb.set_style("Budget", "a1:d1", bold=True)


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
    assert wb.set("Budget", "B3", "1000").applied is True
    assert wb.set("Budget", "B3", "1000").applied is False


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
        results = list(pool.map(lambda _: wb.render_png({"sheet": 0, "range": "A1:D12"}), range(8)))

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
    png = bo.Workbook.open(sample_bytes).render_png({"sheet": 0, "range": "A1:D12"})
    assert png.bytes[:8] == PNG_MAGIC
    assert png.width > 0 and png.height > 0
    assert len(png) == len(png.bytes)


def test_render_png_scale_changes_dimensions(sample_bytes):
    wb = bo.Workbook.open(sample_bytes)
    small = wb.render_png({"sheet": 0, "range": "A1:D12"}, scale=1.0)
    large = wb.render_png({"sheet": 0, "range": "A1:D12"}, scale=2.0)
    assert large.width > small.width
    assert large.height > small.height


def test_render_png_write(sample_bytes, tmp_path):
    target = tmp_path / "sheet.png"
    bo.Workbook.open(sample_bytes).render_png({"sheet": 0, "range": "A1:D12"}).write(target)
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
        bo.Workbook.open(sample_bytes).render_png({"sheet": 0, "range": "ZZ"})


def test_garbage_bytes_raise_parse_error():
    with pytest.raises(bo.ParseError):
        bo.Workbook.open(b"not a zip")


def test_error_hierarchy_rolls_up_to_xlsx_error():
    for subclass in (
        bo.ParseError,
        bo.RangeError,
        bo.RenderError,
        bo.InvalidUpdateError,
        bo.CollaborativeStateError,
        bo.StaleProposalError,
        bo.NotCollaborativeError,
    ):
        assert issubclass(subclass, bo.XlsxError)
    with pytest.raises(bo.XlsxError):
        bo.Workbook.open(b"not a zip")


def test_version_is_exposed():
    assert bo.__version__.count(".") == 2


def test_replicas_converge(sample_bytes):
    left = bo.Workbook.open_collaborative(sample_bytes, client_id=101)
    right = bo.Workbook.open_collaborative(sample_bytes, client_id=202)
    assert left.is_collaborative and right.is_collaborative
    assert (left.client_id, right.client_id) == (101, 202)

    left["Budget"]["B3"] = 1000
    assert right.apply_update(left.diff(right.state_vector()))
    assert right.value("Budget", "D3") == pytest.approx(left.value("Budget", "D3"))

    right["Budget"]["C3"] = 500
    left.apply_update(right.diff(left.state_vector()))
    assert left.value("Budget", "C3") == pytest.approx(500.0)


def test_collaborative_open_generates_read_only_client_ids(sample_bytes):
    left = bo.Workbook.open_collaborative(sample_bytes)
    right = bo.Workbook.open_collaborative(sample_bytes)

    assert left.client_id != right.client_id
    assert 0 <= left.client_id < 2**53
    with pytest.raises(AttributeError):
        left.client_id = 99


@pytest.mark.parametrize("buffer_type", [bytearray, memoryview])
def test_collaboration_accepts_bytes_like_inputs(sample_bytes, buffer_type):
    assert bo.Workbook.open(buffer_type(sample_bytes)).sheet_count == 3
    left = bo.Workbook.open_collaborative(buffer_type(sample_bytes), client_id=401)
    right = bo.Workbook.open_collaborative(buffer_type(sample_bytes), client_id=402)

    left["Budget"]["B3"] = 1000
    update = left.diff(buffer_type(right.state_vector()))
    assert right.apply_update(buffer_type(update))


def test_oversized_update_is_rejected_before_copy(sample_bytes):
    class CopyGuard(bytearray):
        copied = False

        def __bytes__(self):
            self.copied = True
            raise AssertionError("oversized payload was copied")

    wb = bo.Workbook.open_collaborative(sample_bytes)
    update = CopyGuard(bo.MAX_COLLABORATION_BYTES + 1)

    with pytest.raises(bo.XlsxError, match="exceeds the .*byte limit"):
        wb.apply_update(update)
    assert update.copied is False


def test_collaboration_errors_are_actionable(sample_bytes):
    collaborative = bo.Workbook.open_collaborative(sample_bytes)
    with pytest.raises(bo.InvalidUpdateError):
        collaborative.apply_update(b"\xff")

    standalone = bo.Workbook.open(sample_bytes)
    with pytest.raises(bo.NotCollaborativeError):
        standalone.apply_update(b"\x00")


def test_a_fresh_replica_catches_up_from_one_update(sample_bytes):
    source = bo.Workbook.open_collaborative(sample_bytes, client_id=101)
    source["Budget"]["B3"] = 1000

    joiner = bo.Workbook.open_collaborative(sample_bytes, client_id=303)
    joiner.apply_update(source.state_as_update())
    assert joiner.value("Budget", "B3") == pytest.approx(1000.0)


def test_undo_covers_local_edits_but_not_remote_ones(sample_bytes):
    left = bo.Workbook.open_collaborative(sample_bytes, client_id=101)
    right = bo.Workbook.open_collaborative(sample_bytes, client_id=202)
    original = left.value("Budget", "B3")

    left["Budget"]["B3"] = 1000
    right["Budget"]["C3"] = 500
    left.apply_update(right.diff(left.state_vector()))

    left.undo()
    assert left.value("Budget", "B3") == pytest.approx(original)
    assert left.value("Budget", "C3") == pytest.approx(500.0)


def test_undo_redo_round_trip(sample_bytes):
    wb = bo.Workbook.open(sample_bytes)
    original = wb.value("Budget", "B3")
    wb["Budget"]["B3"] = 1000

    assert wb.can_undo and not wb.can_redo
    assert wb.history().undo_depth == 1

    wb.undo()
    assert wb.value("Budget", "B3") == pytest.approx(original)
    assert wb.can_redo

    wb.redo()
    assert wb.value("Budget", "B3") == pytest.approx(1000.0)


def test_set_many_is_one_undo_step(sample_bytes):
    wb = bo.Workbook.open(sample_bytes)
    assert wb.set_many("Budget", {"H1": 10, "H2": 20, "H3": "=H1+H2"})
    assert wb.value("Budget", "H3") == pytest.approx(30.0)

    wb.undo()
    assert wb.value("Budget", "H1") is None
    assert wb.value("Budget", "H3") is None


def test_mutation_lists_recalculated_dependents(sample_bytes):
    wb = bo.Workbook.open(sample_bytes)
    mutation = wb.set("Budget", "B3", "1000")
    assert mutation.applied and bool(mutation)
    assert "D3" in [name.split("!")[-1] for name in mutation.changed]


def test_a_proposal_does_not_apply_until_accepted(sample_bytes):
    wb = bo.Workbook.open(sample_bytes)
    proposal = wb.propose(
        "copilot", [("Budget", "H1", "=SUM(B3:B10)")], note="add a total"
    )

    assert proposal.agent_id == "copilot"
    assert proposal.note == "add a total"
    assert [edit.address for edit in proposal.edits] == ["H1"]
    assert proposal.edits[0].after
    assert len(wb.proposals()) == 1
    assert wb.value("Budget", "H1") is None

    wb.accept_proposal(proposal.id)
    assert wb.value("Budget", "H1") is not None
    assert wb.proposals() == []


def test_stale_proposal_exposes_cells_and_can_be_forced(sample_bytes):
    wb = bo.Workbook.open(sample_bytes)
    proposal = wb.propose("copilot", [("Budget", "H1", 10)])
    wb["Budget"]["H1"] = 999

    with pytest.raises(bo.StaleProposalError) as caught:
        wb.accept_proposal(proposal.id)
    assert caught.value.cells == ["H1"]

    assert wb.accept_proposal(proposal.id, force=True)
    assert wb["Budget"]["H1"] == pytest.approx(10.0)


def test_listed_proposals_name_their_stale_targets(sample_bytes):
    wb = bo.Workbook.open(sample_bytes)
    proposal = wb.propose("copilot", [("Budget", "H1", 10), ("Budget", "H2", 20)])
    assert proposal.stale_targets == []
    assert wb.list_proposals()[0].stale_targets == []

    wb["Budget"]["H2"] = 999

    assert [listed.id for listed in wb.list_proposals()] == [proposal.id]
    assert wb.list_proposals()[0].stale_targets == ["H2"]
    assert wb.proposals()[0].stale_targets == ["H2"]
    with pytest.raises(bo.StaleProposalError) as caught:
        wb.accept_proposal(proposal.id)
    assert caught.value.cells == wb.list_proposals()[0].stale_targets


def test_missing_proposal_raises_key_error(sample_bytes):
    with pytest.raises(KeyError, match="missing"):
        bo.Workbook.open(sample_bytes).accept_proposal("missing")


def test_a_rejected_proposal_leaves_the_sheet_alone(sample_bytes):
    wb = bo.Workbook.open(sample_bytes)
    proposal = wb.propose("copilot", [("Budget", "H2", 999)])

    assert wb.reject_proposal(proposal.id) is True
    assert wb.proposals() == []
    assert wb.value("Budget", "H2") is None
    assert wb.reject_proposal(proposal.id) is False


def test_sheet_metadata(sample_bytes):
    wb = bo.Workbook.open(sample_bytes)
    assert wb.active_sheet == 0
    wb.set_active_sheet("Summary")
    assert wb.active_sheet == 1
    assert wb.merged_ranges("Budget", "A1:H30") == ["A1:D1"]

    wb["Budget"]["B3"] = 1000
    assert wb.last_calculation().changed >= 1


def test_active_sheet_round_trips(sample_bytes):
    wb = bo.Workbook.open(sample_bytes)
    wb.set_active_sheet("Summary")

    saved = wb.save()
    with zipfile.ZipFile(io.BytesIO(saved)) as archive:
        workbook_xml = archive.read("xl/workbook.xml").decode()
    assert 'activeTab="1"' in workbook_xml
    assert bo.Workbook.open(saved).active_sheet == 1


def test_formatting(sample_bytes):
    wb = bo.Workbook.open(sample_bytes)
    assert wb.set_number_format("Budget", "B3:B10", "#,##0.00")
    assert wb.set_style("Budget", "A1:D1", bold=True, fill_color="#eeeeee")
    assert bo.Workbook.open(wb.save()).sheet_names == wb.sheet_names


def test_custom_number_formats_keep_their_case(sample_bytes):
    for pattern in ('0.00 "USD"', "[Red]#,##0.00", r"\K0.00"):
        wb = bo.Workbook.open(sample_bytes)
        assert wb.set_number_format("Budget", "B3:B4", pattern)
        with zipfile.ZipFile(io.BytesIO(wb.save())) as archive:
            styles = archive.read("xl/styles.xml").decode()
        codes = [
            xml.sax.saxutils.unescape(code, {"&quot;": '"'})
            for code in re.findall(r'formatCode="([^"]*)"', styles)
        ]
        assert pattern in codes, f"{pattern!r} was not written verbatim, saw {codes}"


def test_formatting_rejects_an_unknown_alignment(sample_bytes):
    wb = bo.Workbook.open(sample_bytes)
    with pytest.raises(ValueError, match="left, center, or right"):
        wb.set_style("Budget", "A1:D1", horizontal_alignment="sideways")


@pytest.mark.parametrize(
    "style",
    [
        {"font_size": 0},
        {"text_color": "red"},
        {"fill_color": "#12345g"},
    ],
)
def test_style_argument_errors_raise_value_error(sample_bytes, style):
    wb = bo.Workbook.open(sample_bytes)
    with pytest.raises(ValueError):
        wb.set_style("Budget", "A1:D1", **style)


def test_viewport_png_target_scales_and_validates_sheet(sample_bytes):
    workbook = bo.Workbook.open(sample_bytes)
    target = {"sheet": 0, "viewport": {"x": 0, "y": 0, "width": 128, "height": 64}}
    first = workbook.render_png(target)
    second = workbook.render_png(target, scale=2.0)
    assert (second.width, second.height) == (first.width * 2, first.height * 2)
    invalid_targets = [
        {"sheet": 99, "range": "A1:B2"},
        {"sheet": 99, "range": "used"},
        {"sheet": 99, "viewport": target["viewport"]},
    ]
    for invalid_target in invalid_targets:
        with pytest.raises(IndexError):
            workbook.render_png(invalid_target)


@pytest.mark.parametrize("target", [None, 0, {}, {"sheet": 0}, {"sheet": 0, "range": "A1", "viewport": {}}, {"sheet": 0, "viewport": {}}, {"sheet": 0, "range": 1}])
def test_rejects_invalid_png_targets(sample_bytes, target):
    with pytest.raises(bo.RangeError):
        bo.Workbook.open(sample_bytes).render_png(target)


def test_png_target_takes_a_sheet_name_or_index_and_any_range_form(sample_bytes):
    workbook = bo.Workbook.open(sample_bytes)
    by_index = workbook.render_png({"sheet": 0, "range": "A1:D12"})
    name = workbook.sheet_names[0]
    assert workbook.render_png({"sheet": name, "range": "A1:D12"}).bytes == by_index.bytes
    a1 = {"kind": "a1", "a1": "A1:D12"}
    assert workbook.render_png({"sheet": 0, "range": a1}).bytes == by_index.bytes
    row_col = {"kind": "rowCol", "start": {"row": 0, "col": 0}, "end": {"row": 11, "col": 3}}
    assert workbook.render_png({"sheet": 0, "range": row_col}).bytes == by_index.bytes


@pytest.mark.parametrize("range_address", [{}, {"kind": "a1"}, {"kind": "rowCol", "start": {"row": 0, "col": 0}}, {"kind": "rowCol", "start": {"row": 5, "col": 5}, "end": {"row": 0, "col": 0}}, {"kind": "rowCol", "start": {"row": 0, "col": 5}, "end": {"row": 5, "col": 0}}, {"kind": "other"}, 1])
def test_rejects_invalid_png_ranges(sample_bytes, range_address):
    with pytest.raises(bo.RangeError):
        bo.Workbook.open(sample_bytes).render_png({"sheet": 0, "range": range_address})


@pytest.mark.parametrize("scale,max_width", [(1.0, None), (1.5, 80)])
def test_png_used_range_matches_render_sheet(sample_bytes, scale, max_width):
    workbook = bo.Workbook.open(sample_bytes)
    target = {"sheet": 0, "range": "used"}
    for address in ("B3", "Z50"):
        workbook[0][address] = 42
        expected = workbook.render_sheet(0, scale=scale, max_width=max_width)
        actual = workbook.render_png(target, scale=scale, max_width=max_width)
        assert actual.bytes == expected.bytes
        assert (actual.width, actual.height) == (expected.width, expected.height)


def test_png_used_range_takes_a_sheet_name(sample_bytes):
    output = io.BytesIO()
    with zipfile.ZipFile(io.BytesIO(sample_bytes)) as source, zipfile.ZipFile(output, "w") as target:
        for part in source.infolist():
            contents = source.read(part.filename)
            if part.filename == "xl/workbook.xml":
                contents = contents.replace(b'name="Budget"', b'name="Sheet1"')
            target.writestr(part, contents)
    workbook = bo.Workbook.open(output.getvalue())
    assert workbook.sheet_names[0] == "Sheet1"
    assert workbook.render_png({"sheet": "Sheet1", "range": "used"}).bytes == workbook.render_sheet(0).bytes


@pytest.mark.parametrize("range_address", ["USED", "used "])
def test_png_used_range_requires_the_exact_keyword(sample_bytes, range_address):
    with pytest.raises(bo.RangeError, match="invalid range"):
        bo.Workbook.open(sample_bytes).render_png({"sheet": 0, "range": range_address})


def test_png_dimension_caps_apply_to_both_target_forms(sample_bytes):
    workbook = bo.Workbook.open(sample_bytes)
    viewport = {"sheet": 0, "viewport": {"x": 0, "y": 0, "width": 128, "height": 64}}
    for target in (viewport, {"sheet": 0, "range": "A1:D12"}):
        png = workbook.render_png(target, max_width=40, max_height=30)
        assert (png.width, png.height) == (40, 30)


def test_render_sheet_frames_the_used_range(sample_bytes):
    workbook = bo.Workbook.open(sample_bytes)
    framed = workbook.render_sheet(0)
    assert framed.width > 0 and framed.height > 0
    capped = workbook.render_sheet(0, max_width=40)
    assert capped.width == 40
    assert workbook.render_sheet(0, range="A1:D12").bytes == workbook.render_png({"sheet": 0, "range": "A1:D12"}).bytes


def test_png_target_keeps_the_sheet_key_errors(sample_bytes):
    workbook = bo.Workbook.open(sample_bytes)
    with pytest.raises(IndexError):
        workbook.render_png({"sheet": -1, "range": "A1"})
    with pytest.raises(TypeError):
        workbook.render_png({"sheet": True, "range": "A1"})
    with pytest.raises(KeyError):
        workbook.render_png({"sheet": "no-such-sheet", "range": "A1"})


@pytest.mark.parametrize("scale,cap", [(1.2, 52), (3.0, 10), (1.7, 33)])
def test_png_caps_hold_at_fractional_scales(sample_bytes, scale, cap):
    workbook = bo.Workbook.open(sample_bytes)
    viewport = {"sheet": 0, "viewport": {"x": 0, "y": 0, "width": 200, "height": 100}}
    for target in (viewport, {"sheet": 0, "range": "A1:D12"}):
        png = workbook.render_png(target, scale=scale, max_width=cap, max_height=cap)
        assert png.width <= cap and png.height <= cap


@pytest.mark.parametrize("width", [float("nan"), float("inf")])
def test_png_caps_do_not_hide_an_invalid_viewport(sample_bytes, width):
    target = {"sheet": 0, "viewport": {"x": 0, "y": 0, "width": width, "height": 64}}
    with pytest.raises(bo.RangeError):
        bo.Workbook.open(sample_bytes).render_png(target, max_width=40)
