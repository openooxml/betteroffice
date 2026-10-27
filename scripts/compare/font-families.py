import json
import sys
import xml.etree.ElementTree as ET
from zipfile import ZipFile
w='{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'
names=set()
with ZipFile(sys.argv[1]) as archive:
    for name in archive.namelist():
        if name.startswith('word/') and name.endswith('.xml'):
            root=ET.fromstring(archive.read(name))
            for element in root.iter(w+'rFonts'):
                names.update(element.get(w+k) for k in ['ascii','hAnsi','eastAsia','cs'] if element.get(w+k))
            for element in root.iter(w+'font'):
                if element.get(w+'name'):names.add(element.get(w+'name'))
print(json.dumps(sorted(names),ensure_ascii=True))
