import hashlib
import json
import os
import subprocess
import sys
import time
import uuid
from pathlib import Path

FILTERS = {'docx': 'Office Open XML Text', 'pptx': 'Impress MS PowerPoint 2007 XML', 'xlsx': 'Calc Office Open XML'}
SERVICES = {'docx': 'com.sun.star.text.TextDocument', 'pptx': 'com.sun.star.presentation.PresentationDocument',
            'xlsx': 'com.sun.star.sheet.SpreadsheetDocument'}


def properties(**values):
    import uno
    result = []
    for name, value in values.items():
        prop = uno.createUnoStruct('com.sun.star.beans.PropertyValue')
        prop.Name, prop.Value = name, value
        result.append(prop)
    return tuple(result)


def drawing_texts(shapes):
    for index in range(shapes.getCount()):
        shape = shapes.getByIndex(index)
        if shape.supportsService('com.sun.star.drawing.GroupShape'):
            yield from drawing_texts(shape)
        elif shape.supportsService('com.sun.star.drawing.TableShape'):
            table = shape.getPropertyValue('Model')
            for row in range(table.getRows().getCount()):
                for column in range(table.getColumns().getCount()):
                    yield table.getCellByPosition(column, row).getText()
        elif hasattr(shape, 'getText'):
            yield shape.getText()


def text_match(document, format, expected):
    if format == 'docx':
        search = document.createSearchDescriptor()
        search.SearchString = expected
        search.SearchCaseSensitive = True
        search.SearchRegularExpression = False
        matches = document.findAll(search)
        found = [matches.getByIndex(index) for index in range(matches.getCount())]
    else:
        pages = document.getDrawPages()
        found = [text for page in range(pages.getCount()) for text in drawing_texts(pages.getByIndex(page))
                 if text.getString() == expected]
    if len(found) != 1 or found[0].getString() != expected:
        raise ValueError('Text is missing or ambiguous')
    return found[0]


def numeric_cell(document, probe, expected):
    cell = document.getSheets().getByName(probe['sheet']).getCellRangeByName(probe['cell'])
    if cell.getType().value != 'VALUE' or cell.getValue() != float(expected):
        raise ValueError('Numeric literal differs from the probe')
    return cell


def edit(document, format, probe):
    if format == 'xlsx':
        numeric_cell(document, probe, probe['old']).setValue(float(probe['new']))
    else:
        found = text_match(document, format, probe['old'])
        if not probe['new'].startswith(probe['old']) or probe['new'] == probe['old']:
            raise ValueError('Invalid append edit')
        cursor = found.getText().createTextCursorByRange(found.getEnd())
        cursor.setString(probe['new'][len(probe['old']):])


def verify(document, format, probe):
    if format == 'xlsx':
        numeric_cell(document, probe, probe['new'])
    else:
        text_match(document, format, probe['new'])


def load(desktop, path, format):
    import uno
    options = properties(Hidden=True, ReadOnly=False,
        MacroExecutionMode=uno.getConstantByName('com.sun.star.document.MacroExecMode.NEVER_EXECUTE'),
        UpdateDocMode=uno.getConstantByName('com.sun.star.document.UpdateDocMode.NO_UPDATE'))
    document = desktop.loadComponentFromURL(path.as_uri(), '_blank', 0, options)
    if document is None or not document.supportsService(SERVICES[format]):
        raise ValueError('LibreOffice did not load the requested document type')
    return document


def main():
    import uno
    if sys.argv[1:] == ['--check']:
        properties(Hidden=True)
        uno.getComponentContext()
        return
    source, probe_path, saved, status = [Path(argument).resolve() for argument in sys.argv[1:]]
    format = source.suffix.lstrip('.')
    state = dict(source_sha256=hashlib.sha256(source.read_bytes()).hexdigest(), parse='failed', stage='parse')
    def checkpoint():
        temporary = status.with_suffix('.tmp')
        temporary.write_text(json.dumps(state, indent=2) + '\n')
        temporary.replace(status)
    checkpoint()
    endpoint = 'pipe,name=betteroffice_' + uuid.uuid4().hex
    command = [os.environ['SOFFICE'], '-env:UserInstallation=' + (saved.parent / 'profile').as_uri(),
               '--headless', '--nologo', '--nodefault', '--nofirststartwizard', '--norestore',
               '--accept=' + endpoint + ';urp;StarOffice.ServiceManager']
    document = None
    with subprocess.Popen(command, env={**os.environ, 'SAL_USE_VCLPLUGIN': 'svp'}) as process:
        try:
            local = uno.getComponentContext()
            resolver = local.ServiceManager.createInstanceWithContext('com.sun.star.bridge.UnoUrlResolver', local)
            deadline = time.monotonic() + 30
            while True:
                try:
                    remote = resolver.resolve('uno:' + endpoint + ';urp;StarOffice.ComponentContext')
                    break
                except Exception:
                    if process.poll() is not None or time.monotonic() >= deadline:
                        raise RuntimeError('Could not connect to the isolated LibreOffice process')
                    time.sleep(0.1)
            desktop = remote.ServiceManager.createInstanceWithContext('com.sun.star.frame.Desktop', remote)
            document = load(desktop, source, format)
            state.update(parse='ok', stage='edit')
            checkpoint()
            probe = json.loads(probe_path.read_text())
            if probe.get('status') != 'ok':
                raise ValueError('No eligible deterministic edit in the source')
            edit(document, format, probe)
            state['stage'] = 'save'
            checkpoint()
            document.storeAsURL(saved.as_uri(), properties(FilterName=FILTERS[format], Overwrite=True))
            document.close(True)
            document = None
            state.update(output_sha256=hashlib.sha256(saved.read_bytes()).hexdigest(), stage='reopen')
            checkpoint()
            document = load(desktop, saved, format)
            verify(document, format, probe)
            state.update(stage='complete', edit_verified=True)
            checkpoint()
        except Exception as error:
            state['error'] = str(error)
            checkpoint()
        finally:
            if process.poll() is None:
                process.kill()
            process.wait()


if __name__ == '__main__':
    main()
