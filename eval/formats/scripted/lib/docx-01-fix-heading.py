import json

from docx import Document


def call(name, function, *args):
    print(json.dumps({'type': 'script.call', 'name': name}), flush=True)
    return function(*args)


document = call('docx.Document', Document, 'input.docx')
for paragraph in document.paragraphs:
    for run in paragraph.runs:
        if 'Sumary' in run.text:
            call('run.text', setattr, run, 'text', run.text.replace('Sumary', 'Summary'))
call('document.save', document.save, 'output.docx')
