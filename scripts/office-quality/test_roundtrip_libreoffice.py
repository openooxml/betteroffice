import unittest
from pathlib import Path
from types import SimpleNamespace
from unittest.mock import patch

from roundtrip_libreoffice import edit, load, verify


class Collection:
    def __init__(self, values):
        self.values = values
    def getCount(self):
        return len(self.values)
    def getByIndex(self, index):
        return self.values[index]


class Text:
    def __init__(self, value):
        self.value = value
    def getString(self):
        return self.value
    def getText(self):
        return self
    def getEnd(self):
        return self
    def createTextCursorByRange(self, end):
        if end is not self:
            raise ValueError('Wrong edit position')
        return SimpleNamespace(setString=lambda suffix: setattr(self, 'value', self.value + suffix))


class LibreOfficeEditTests(unittest.TestCase):
    def test_writer_appends_the_same_marker_and_requires_it_after_reopen(self):
        text = Text('original')
        document = SimpleNamespace(createSearchDescriptor=lambda: SimpleNamespace(),
                                   findAll=lambda query: Collection([text] if query.SearchString == text.value else []))
        probe = dict(old='original', new='original_marker')
        edit(document, 'docx', probe)
        self.assertEqual(text.value, probe['new'])
        verify(document, 'docx', probe)
        text.value = 'original'
        with self.assertRaises(ValueError):
            verify(document, 'docx', probe)

    def test_impress_finds_text_inside_grouped_shapes(self):
        text = Text('original')
        shape = SimpleNamespace(supportsService=lambda name: False, getText=lambda: text)
        group = Collection([shape])
        group.supportsService = lambda name: name == 'com.sun.star.drawing.GroupShape'
        document = SimpleNamespace(getDrawPages=lambda: Collection([Collection([group])]))
        probe = dict(old='original', new='original_marker')
        edit(document, 'pptx', probe)
        self.assertEqual(text.value, probe['new'])
        verify(document, 'pptx', probe)

    def test_impress_edits_table_cell_text(self):
        text = Text('original')
        table = SimpleNamespace(getRows=lambda: Collection([0]), getColumns=lambda: Collection([0]),
                                getCellByPosition=lambda column, row: text)
        shape = SimpleNamespace(supportsService=lambda name: name == 'com.sun.star.drawing.TableShape',
                                getPropertyValue=lambda name: table)
        document = SimpleNamespace(getDrawPages=lambda: Collection([Collection([shape])]))
        probe = dict(old='original', new='original_marker')
        edit(document, 'pptx', probe)
        verify(document, 'pptx', probe)

    def test_calc_edits_the_same_literal_and_rejects_formula_cells(self):
        cell = SimpleNamespace(value=7, kind='VALUE')
        cell.getType = lambda: SimpleNamespace(value=cell.kind)
        cell.getValue = lambda: cell.value
        cell.setValue = lambda value: setattr(cell, 'value', value)
        sheet = SimpleNamespace(getCellRangeByName=lambda name: cell)
        document = SimpleNamespace(getSheets=lambda: SimpleNamespace(getByName=lambda name: sheet))
        probe = dict(sheet='Data', cell='A1', old='7', new='0')
        edit(document, 'xlsx', probe)
        self.assertEqual(cell.value, 0)
        verify(document, 'xlsx', probe)
        cell.kind = 'FORMULA'
        with self.assertRaises(ValueError):
            verify(document, 'xlsx', probe)

    def test_load_is_hidden_and_disables_macros_and_external_updates(self):
        captured = {}
        document = SimpleNamespace(supportsService=lambda name: name == 'com.sun.star.text.TextDocument')
        def load_component(url, frame, flags, options):
            captured.update({prop.Name: prop.Value for prop in options})
            return document
        uno = SimpleNamespace(createUnoStruct=lambda name: SimpleNamespace(), getConstantByName=lambda name: name)
        with patch.dict('sys.modules', uno=uno):
            loaded = load(SimpleNamespace(loadComponentFromURL=load_component), Path('/tmp/source.docx'), 'docx')
        self.assertIs(loaded, document)
        self.assertTrue(captured['Hidden'])
        self.assertFalse(captured['ReadOnly'])
        self.assertTrue(captured['MacroExecutionMode'].endswith('NEVER_EXECUTE'))
        self.assertTrue(captured['UpdateDocMode'].endswith('NO_UPDATE'))


if __name__ == '__main__':
    unittest.main()
