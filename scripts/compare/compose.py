#!/usr/bin/env python3
"""Deterministic docx composer: appends the bodies of several public corpus documents to the first one.

python3 -I scripts/compare/compose.py --corpus <corpus dir> --out <path.docx> <id> <id> ...
The corpus dir holds manifest.jsonl and raw/docx/<id>.docx. Writes <out> and <out stem>.PROVENANCE.txt,
then validates the result and prints per-feature counts.
"""

import argparse
import json
import os
import posixpath
import re
import sys
import zipfile
import xml.etree.ElementTree as ET

ALLOWED_LICENCES = {"CC-BY-4.0", "CC0-1.0", "OGL-UK-3.0", "OGL-Canada-2.0", "public-domain-us-gov"}
LICENCE_URLS = {
    "CC-BY-4.0": "https://creativecommons.org/licenses/by/4.0/",
    "CC0-1.0": "https://creativecommons.org/publicdomain/zero/1.0/",
    "OGL-UK-3.0": "https://www.nationalarchives.gov.uk/doc/open-government-licence/version/3/",
    "OGL-Canada-2.0": "https://open.canada.ca/en/open-government-licence-canada",
    "public-domain-us-gov": "https://www.usa.gov/government-copyright",
}
CHANGES = "combined with other documents; headers, footers and comments of appended parts removed; styles and numbering merged"

W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
O = "urn:schemas-microsoft-com:office:office"
WP = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
MC = "http://schemas.openxmlformats.org/markup-compatibility/2006"
PR = "http://schemas.openxmlformats.org/package/2006/relationships"
CT = "http://schemas.openxmlformats.org/package/2006/content-types"
REL_DOC = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/"
CT_FOOTNOTES = "application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml"
CT_ENDNOTES = "application/vnd.openxmlformats-officedocument.wordprocessingml.endnotes+xml"
CT_NUMBERING = "application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml"
CT_DGM_DATA = "application/vnd.openxmlformats-officedocument.drawingml.diagramData+xml"

SKIPPED_REL_TYPES = {
    "styles", "numbering", "settings", "webSettings", "fontTable", "theme", "footnotes", "endnotes",
    "comments", "commentsExtended", "commentsIds", "commentsExtensible", "people", "header", "footer",
    "customXml", "glossaryDocument", "stylesWithEffects", "documenttasks", "intelligence", "intelligence2",
}
REVISION_TAGS_WITHOUT_OFFSET = {"footnote", "endnote", "footnoteReference", "endnoteReference"}
COMMENT_TAGS = {"commentRangeStart", "commentRangeEnd", "commentReference"}
NOTE_SPECIAL_TYPES = {"separator", "continuationSeparator", "continuationNotice"}


def w(tag):
    return f"{{{W}}}{tag}"


def local(tag):
    return tag.rsplit("}", 1)[-1]


def ns_of(tag):
    return tag[1:].split("}", 1)[0] if tag.startswith("{") else ""


class Namespaces:
    def __init__(self):
        self.uri_by_prefix = {}

    def collect(self, data):
        head = re.search(rb"<(?![?!])[^>]*>", data).group(0)
        pattern = rb'xmlns:([A-Za-z_][\w.-]*)="([^"]*)"'
        root = {p.decode(): u.decode() for p, u in re.findall(pattern, head)}
        inner = {p.decode(): u.decode() for p, u in re.findall(pattern, data)}
        found = {}
        for prefix, uri in {**inner, **root}.items():
            known = self.uri_by_prefix.get(prefix)
            if known is not None and known != uri:
                if prefix in root:
                    raise SystemExit(f"namespace prefix {prefix} bound to both {known} and {uri}")
                continue
            found[prefix] = uri
            if known is None:
                self.uri_by_prefix[prefix] = uri
                ET.register_namespace(prefix, uri)
        return found

    def parse(self, data):
        self.collect(data)
        return ET.fromstring(data)

    def serialize(self, root, declared):
        body = ET.tostring(root, encoding="unicode")
        end = body.index(">")
        if body[end - 1] == "/":
            end -= 1
        tag = body[:end]
        extra = "".join(f' xmlns:{p}="{u}"' for p, u in sorted(declared.items()) if f"xmlns:{p}=" not in tag)
        out = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n' + tag + extra + body[end:]
        return out.encode("utf-8")


class Package:
    def __init__(self, path):
        self.path = path
        with zipfile.ZipFile(path) as z:
            self.order = [i.filename for i in z.infolist() if not i.is_dir()]
            self.parts = {n: z.read(n) for n in self.order}
        ct = ET.fromstring(self.parts["[Content_Types].xml"])
        self.defaults = {e.get("Extension").lower(): e.get("ContentType") for e in ct if local(e.tag) == "Default"}
        self.overrides = {e.get("PartName").lstrip("/"): e.get("ContentType") for e in ct if local(e.tag) == "Override"}
        self.main = next(
            resolve("", r["Target"]) for r in self.rels("") if r["Type"].endswith("/officeDocument")
        )

    def content_type(self, name):
        return self.overrides.get(name) or self.defaults.get(ext_of(name))

    def rels(self, part):
        data = self.parts.get(rels_name(part))
        if data is None:
            return []
        return [dict(e.attrib) for e in ET.fromstring(data) if local(e.tag) == "Relationship"]


def ext_of(name):
    base = posixpath.basename(name)
    return base.rsplit(".", 1)[1].lower() if "." in base else ""


def rels_name(part):
    d, b = posixpath.split(part)
    return posixpath.join(d, "_rels", b + ".rels")


def resolve(source_part, target):
    if target.startswith("/"):
        return target.lstrip("/")
    return posixpath.normpath(posixpath.join(posixpath.dirname(source_part), target))


def rels_xml(rels):
    out = ['<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n', f'<Relationships xmlns="{PR}">']
    for r in rels:
        attrs = "".join(f' {k}="{xml_attr(r[k])}"' for k in ("Id", "Type", "Target", "TargetMode") if k in r)
        out.append(f"<Relationship{attrs}/>")
    out.append("</Relationships>")
    return "".join(out).encode("utf-8")


def xml_attr(v):
    return v.replace("&", "&amp;").replace('"', "&quot;").replace("<", "&lt;").replace(">", "&gt;")


class Composer:
    def __init__(self, base):
        self.ns = Namespaces()
        self.out = base
        self.main = base.main
        self.declared = {}
        self.trees = {}
        self.part_rels = {}
        self.copied = {}
        self.dangling = []
        self.doc = self.load(self.main)
        self.body = self.doc.find(w("body"))
        self.styles = self.load_rel_part(self.main, "styles")
        self.numbering = self.load_rel_part(self.main, "numbering")
        self.font_table = self.load_rel_part(self.main, "fontTable")
        self.notes = {kind: self.load_rel_part(self.main, kind + "s") for kind in ("footnote", "endnote")}
        self.docpr_next = 1 + max((int(v) for v in self.all_values("docPr", "id")), default=0)
        self.annot_max = max(self.annotation_ids(self.doc), default=0)
        for kind in self.notes:
            if self.notes[kind] is not None:
                self.annot_max = max(self.annot_max, max(self.annotation_ids(self.notes[kind][1]), default=0))
        base_final = self.body.find(w("sectPr"))
        if base_final is not None:
            self.body.remove(base_final)
            self.body.append(section_paragraph(base_final))

    def load(self, name):
        tree = self.ns.parse(self.out.parts[name])
        self.trees[name] = tree
        self.declared.setdefault(name, {}).update(self.ns.collect(self.out.parts[name]))
        return tree

    def load_rel_part(self, part, rel_type):
        for r in self.rels_of(part):
            if r["Type"].rsplit("/", 1)[1] == rel_type and r.get("TargetMode") != "External":
                name = resolve(part, r["Target"])
                return name, self.load(name)
        return None

    def rels_of(self, part):
        if part not in self.part_rels:
            self.part_rels[part] = self.out.rels(part)
        return self.part_rels[part]

    def all_values(self, tag, attr):
        for name, data in self.out.parts.items():
            if name.startswith("word/") and name.endswith(".xml"):
                for m in re.finditer(rb"<[\w]+:" + tag.encode() + rb"\b[^>]*?\s" + attr.encode() + rb'="(\d+)"', data):
                    yield m.group(1).decode()

    @staticmethod
    def annotation_ids(root):
        for el in root.iter():
            if ns_of(el.tag) == W and local(el.tag) not in REVISION_TAGS_WITHOUT_OFFSET:
                v = el.get(w("id"))
                if v is not None and v.lstrip("-").isdigit():
                    yield int(v)

    def new_rid(self, part):
        used = {r["Id"] for r in self.rels_of(part)}
        n = len(used) + 1
        while f"rId{n}" in used:
            n += 1
        return f"rId{n}"

    def unique_name(self, name):
        d, b = posixpath.split(name)
        stem, ext = posixpath.splitext(b)
        stem = re.sub(r"\d+$", "", stem)
        n = 1
        # OPC part names are equivalent ignoring ASCII case, so image1.JPG collides with image1.jpg.
        taken = {p.lower() for p in self.out.parts}
        while posixpath.join(d, f"{stem}{n}{ext}").lower() in taken:
            n += 1
        return posixpath.join(d, f"{stem}{n}{ext}")

    def add_part(self, name, data, ctype, src_ext_default):
        self.out.parts[name] = data
        self.out.order.append(name)
        ext = ext_of(name)
        if src_ext_default and self.out.defaults.get(ext) is None:
            self.out.defaults[ext] = ctype
        elif self.out.defaults.get(ext) != ctype:
            self.out.overrides[name] = ctype

    def copy_part(self, src, name, doc_rid_map):
        key = (src.path, name)
        if key in self.copied:
            return self.copied[key]
        if name not in src.parts:
            self.dangling.append(name)
            return None
        new = self.unique_name(name)
        self.copied[key] = new
        ctype = src.content_type(name)
        if ctype is None:
            raise SystemExit(f"no content type for {name}")
        data = src.parts[name]
        if ctype == CT_DGM_DATA:
            data = re.sub(rb'relId="([^"]+)"', lambda m: b'relId="' + doc_rid_map.get(m.group(1).decode(), m.group(1).decode()).encode() + b'"', data)
        ext = ext_of(name)
        self.add_part(new, data, ctype, name not in src.overrides and src.defaults.get(ext) == ctype)
        rels = []
        for r in src.rels(name):
            r = dict(r)
            if r.get("TargetMode") != "External":
                child = self.copy_part(src, resolve(name, r["Target"]), doc_rid_map)
                if child is None:
                    continue
                r["Target"] = posixpath.relpath(child, posixpath.dirname(new))
            rels.append(r)
        if rels:
            self.add_part(rels_name(new), rels_xml(rels), "application/vnd.openxmlformats-package.relationships+xml", True)
        return new

    def import_rels(self, src, src_part, dst_part, skip=SKIPPED_REL_TYPES):
        mapping = {}
        pending = []
        for r in src.rels(src_part):
            if r["Type"].rsplit("/", 1)[1] in skip:
                continue
            rid = self.new_rid(dst_part)
            mapping[r["Id"]] = rid
            self.rels_of(dst_part).append({"Id": rid, "Type": r["Type"], "Target": r["Target"], **({"TargetMode": r["TargetMode"]} if "TargetMode" in r else {})})
            pending.append((r, self.rels_of(dst_part)[-1]))
        for r, new in pending:
            if r.get("TargetMode") == "External":
                continue
            copied = self.copy_part(src, resolve(src_part, r["Target"]), mapping)
            if copied is None:
                self.rels_of(dst_part).remove(new)
                del mapping[r["Id"]]
                continue
            new["Target"] = posixpath.relpath(copied, posixpath.dirname(dst_part))
        return mapping

    def ensure_note_part(self, kind, src, src_name):
        if self.notes[kind] is not None:
            return self.notes[kind]
        name = f"word/{kind}s.xml"
        self.out.parts[name] = src.parts[src_name]
        self.out.order.append(name)
        self.out.overrides[name] = CT_FOOTNOTES if kind == "footnote" else CT_ENDNOTES
        root = self.load(name)
        for el in list(root):
            if el.get(w("type")) not in NOTE_SPECIAL_TYPES:
                root.remove(el)
        self.rels_of(self.main).append({"Id": self.new_rid(self.main), "Type": REL_DOC + kind + "s", "Target": f"{kind}s.xml"})
        self.notes[kind] = (name, root)
        return self.notes[kind]

    def ensure_numbering(self, src_data):
        if self.numbering is not None:
            return self.numbering
        name = "word/numbering.xml"
        self.out.parts[name] = src_data
        self.out.order.append(name)
        self.out.overrides[name] = CT_NUMBERING
        root = self.load(name)
        for el in list(root):
            root.remove(el)
        self.rels_of(self.main).append({"Id": self.new_rid(self.main), "Type": REL_DOC + "numbering", "Target": "numbering.xml"})
        self.numbering = (name, root)
        return self.numbering

    def append(self, src, is_last):
        src_main = src.main
        src_doc_data = src.parts[src_main]
        src_doc = self.ns.parse(src_doc_data)
        self.declared[self.main].update(self.ns.collect(src_doc_data))
        merge_ignorable(self.doc, src_doc)
        rid_map = self.import_rels(src, src_main, self.main)
        num_map = self.merge_numbering(src)
        self.merge_styles(src, num_map)
        self.merge_fonts(src)
        annot_off = self.annot_max + 1
        note_maps = {}
        for kind in ("footnote", "endnote"):
            note_maps[kind] = self.merge_notes(src, kind, num_map, annot_off)
        src_body = src_doc.find(w("body"))
        final = src_body.find(w("sectPr"))
        if final is not None:
            src_body.remove(final)
        sections = list(src_body.iter(w("sectPr"))) + ([final] if final is not None else [])
        for s in sections:
            strip_section(s)
        self.rewrite(src_body, rid_map, num_map, annot_off, note_maps)
        if final is not None:
            self.rewrite(final, rid_map, num_map, annot_off, note_maps)
        if sections:
            first = sections[0]
            t = first.find(w("type"))
            if t is None:
                t = ET.Element(w("type"))
                at = sum(1 for c in first if local(c.tag) in ("footnotePr", "endnotePr"))
                first.insert(at, t)
            if t.get(w("val")) not in ("oddPage", "evenPage"):
                t.set(w("val"), "nextPage")
        self.annot_max = max([self.annot_max, *self.annotation_ids(src_body)])
        for child in list(src_body):
            self.body.append(child)
        if final is not None:
            if is_last:
                self.body.append(final)
            else:
                self.body.append(section_paragraph(final))

    def rewrite(self, root, rid_map, num_map, annot_off, note_maps):
        drop_comments(root)
        for el in root.iter():
            for k, v in list(el.attrib.items()):
                if ns_of(k) == R or k == f"{{{O}}}relid":
                    if v not in rid_map:
                        raise SystemExit(f"unresolved relationship {v} on {local(el.tag)}")
                    el.set(k, rid_map[v])
            t = el.tag
            if t == f"{{{WP}}}docPr":
                el.set("id", str(self.docpr_next))
                self.docpr_next += 1
            elif ns_of(t) != W:
                continue
            name = local(t)
            if name == "numId":
                v = el.get(w("val"))
                if v is not None and v != "0":
                    el.set(w("val"), num_map.get(v, v))
            elif name in ("footnoteReference", "endnoteReference"):
                m = note_maps[name[: -len("Reference")]]
                v = el.get(w("id"))
                if v in m:
                    el.set(w("id"), m[v])
            elif name not in REVISION_TAGS_WITHOUT_OFFSET:
                v = el.get(w("id"))
                if v is not None and v.lstrip("-").isdigit():
                    el.set(w("id"), str(int(v) + annot_off))
            if name == "bookmarkStart" and el.get(w("name")) == "_GoBack":
                el.set(w("name"), f"_GoBack{annot_off}")

    def merge_numbering(self, src):
        src_ref = next((r for r in src.rels(src.main) if r["Type"].endswith("/numbering")), None)
        if src_ref is None:
            return {}
        src_name = resolve(src.main, src_ref["Target"])
        data = src.parts[src_name]
        src_root = self.ns.parse(data)
        name, root = self.ensure_numbering(data)
        self.declared[name].update(self.ns.collect(data))
        merge_ignorable(root, src_root)
        rid_map = self.import_rels(src, src_name, name, skip=set())
        abs_off = 1 + max((int(e.get(w("abstractNumId"))) for e in root.iter(w("abstractNum"))), default=-1)
        num_off = max((int(e.get(w("numId"))) for e in root.iter(w("num"))), default=0)
        pic_off = 1 + max((int(e.get(w("numPicBulletId"))) for e in root.iter(w("numPicBullet"))), default=-1)
        num_map = {}
        groups = {"numPicBullet": [], "abstractNum": [], "num": []}
        for el in src_root:
            kind = local(el.tag)
            if kind not in groups:
                continue
            if kind == "numPicBullet":
                el.set(w("numPicBulletId"), str(int(el.get(w("numPicBulletId"))) + pic_off))
            elif kind == "abstractNum":
                el.set(w("abstractNumId"), str(int(el.get(w("abstractNumId"))) + abs_off))
                for b in el.iter(w("lvlPicBulletId")):
                    b.set(w("val"), str(int(b.get(w("val"))) + pic_off))
            else:
                old = el.get(w("numId"))
                num_map[old] = str(int(old) + num_off)
                el.set(w("numId"), num_map[old])
                a = el.find(w("abstractNumId"))
                a.set(w("val"), str(int(a.get(w("val"))) + abs_off))
            for sub in el.iter():
                for k, v in list(sub.attrib.items()):
                    if ns_of(k) == R or k == f"{{{O}}}relid":
                        sub.set(k, rid_map[v])
                if sub.tag == f"{{{WP}}}docPr":
                    sub.set("id", str(self.docpr_next))
                    self.docpr_next += 1
            groups[kind].append(el)
        existing = list(root)
        for el in existing:
            root.remove(el)
        order = ["numPicBullet", "abstractNum", "num"]
        for kind in order:
            for el in existing:
                if local(el.tag) == kind:
                    root.append(el)
            for el in groups[kind]:
                root.append(el)
        for el in existing:
            if local(el.tag) not in order:
                root.append(el)
        return num_map

    def merge_styles(self, src, num_map):
        src_ref = next((r for r in src.rels(src.main) if r["Type"].endswith("/styles")), None)
        if src_ref is None or self.styles is None:
            return
        data = src.parts[resolve(src.main, src_ref["Target"])]
        src_root = self.ns.parse(data)
        name, root = self.styles
        self.declared[name].update(self.ns.collect(data))
        merge_ignorable(root, src_root)
        have = {s.get(w("styleId")) for s in root.iter(w("style"))}
        for s in src_root.findall(w("style")):
            if s.get(w("styleId")) in have:
                continue
            for n in s.iter(w("numId")):
                v = n.get(w("val"))
                if v is not None and v != "0":
                    n.set(w("val"), num_map.get(v, v))
            root.append(s)
            have.add(s.get(w("styleId")))

    def merge_fonts(self, src):
        src_ref = next((r for r in src.rels(src.main) if r["Type"].endswith("/fontTable")), None)
        if src_ref is None or self.font_table is None:
            return
        data = src.parts[resolve(src.main, src_ref["Target"])]
        src_root = self.ns.parse(data)
        name, root = self.font_table
        self.declared[name].update(self.ns.collect(data))
        have = {f.get(w("name")) for f in root.findall(w("font"))}
        for f in src_root.findall(w("font")):
            if f.get(w("name")) in have:
                continue
            for e in list(f):
                if local(e.tag).startswith("embed"):
                    f.remove(e)
            root.append(f)
            have.add(f.get(w("name")))

    def merge_notes(self, src, kind, num_map, annot_off):
        src_ref = next((r for r in src.rels(src.main) if r["Type"].endswith("/" + kind + "s")), None)
        if src_ref is None:
            return {}
        src_name = resolve(src.main, src_ref["Target"])
        data = src.parts[src_name]
        src_root = self.ns.parse(data)
        normal = [n for n in src_root.findall(w(kind)) if n.get(w("type")) not in NOTE_SPECIAL_TYPES]
        if not normal:
            return {}
        name, root = self.ensure_note_part(kind, src, src_name)
        self.declared[name].update(self.ns.collect(data))
        merge_ignorable(root, src_root)
        rid_map = self.import_rels(src, src_name, name)
        next_id = 1 + max((int(n.get(w("id"))) for n in root.findall(w(kind))), default=0)
        mapping = {}
        for n in normal:
            mapping[n.get(w("id"))] = str(next_id)
            n.set(w("id"), str(next_id))
            next_id += 1
        for n in normal:
            for s in n.iter(w("sectPr")):
                strip_section(s)
            self.rewrite(n, rid_map, num_map, annot_off, {"footnote": {}, "endnote": {}})
            self.annot_max = max([self.annot_max, *self.annotation_ids(n)])
            root.append(n)
        return mapping

    def write(self, path):
        for name, tree in self.trees.items():
            self.out.parts[name] = self.ns.serialize(tree, self.declared.get(name, {}))
        for part, rels in self.part_rels.items():
            rn = rels_name(part)
            if rn not in self.out.parts:
                self.out.order.append(rn)
            self.out.parts[rn] = rels_xml(rels)
        ct = ['<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n', f'<Types xmlns="{CT}">']
        ct += [f'<Default Extension="{xml_attr(e)}" ContentType="{xml_attr(c)}"/>' for e, c in sorted(self.out.defaults.items())]
        ct += [f'<Override PartName="/{xml_attr(n)}" ContentType="{xml_attr(c)}"/>' for n, c in sorted(self.out.overrides.items()) if n in self.out.parts]
        ct.append("</Types>")
        self.out.parts["[Content_Types].xml"] = "".join(ct).encode("utf-8")
        order = ["[Content_Types].xml"] + [n for n in dict.fromkeys(self.out.order) if n != "[Content_Types].xml"]
        with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as z:
            for n in order:
                info = zipfile.ZipInfo(n, date_time=(1980, 1, 1, 0, 0, 0))
                info.compress_type = zipfile.ZIP_DEFLATED
                info.external_attr = 0o644 << 16
                z.writestr(info, self.out.parts[n])


def merge_ignorable(dst, src):
    key = f"{{{MC}}}Ignorable"
    tokens = (dst.get(key) or "").split()
    for t in (src.get(key) or "").split():
        if t not in tokens:
            tokens.append(t)
    if tokens:
        dst.set(key, " ".join(tokens))


def section_paragraph(sect):
    p = ET.Element(w("p"))
    ppr = ET.SubElement(p, w("pPr"))
    ppr.append(sect)
    return p


def strip_section(sect):
    for el in list(sect):
        if local(el.tag) in ("headerReference", "footerReference"):
            sect.remove(el)
        elif local(el.tag) == "pgNumType" and el.get(w("start")) is not None:
            del el.attrib[w("start")]


def drop_comments(root):
    for parent in list(root.iter()):
        for child in list(parent):
            if ns_of(child.tag) != W:
                continue
            if local(child.tag) in COMMENT_TAGS or local(child.tag) == "dataBinding":
                parent.remove(child)
            elif local(child.tag) == "r" and child.find(w("commentReference")) is not None:
                child.remove(child.find(w("commentReference")))
                if all(local(c.tag) == "rPr" for c in child):
                    parent.remove(child)


def load_manifest(corpus):
    rows = {}
    with open(os.path.join(corpus, "manifest.jsonl"), encoding="utf-8") as f:
        for line in f:
            d = json.loads(line)
            rows[d["id"]] = d
    return rows


def feature_counts(path):
    counts = dict.fromkeys(
        ["ins", "del", "moves", "vMerge", "gridSpan", "nested_tables", "tables", "fields", "toc_fields", "textboxes",
         "anchored_drawings", "inline_drawings", "footnotes", "endnotes", "rtl_runs", "rtl_chars", "cjk_chars",
         "images", "media_parts", "charts", "ole_objects"], 0)
    cjk = re.compile("[぀-ヿ㐀-鿿가-힯豈-﫿]")
    rtl = re.compile("[֐-ࣿיִ-﷿ﹰ-﻿]")
    with zipfile.ZipFile(path) as z:
        names = z.namelist()
        doc = ET.fromstring(z.read("word/document.xml"))
        stories = [doc]
        for kind in ("footnotes", "endnotes"):
            if f"word/{kind}.xml" in names:
                root = ET.fromstring(z.read(f"word/{kind}.xml"))
                stories.append(root)
                counts[kind] = sum(1 for n in root if n.get(w("type")) not in NOTE_SPECIAL_TYPES)
        for root in stories:
            for el in root.iter():
                t = el.tag
                if t == w("ins"):
                    counts["ins"] += 1
                elif t == w("del"):
                    counts["del"] += 1
                elif t in (w("moveFrom"), w("moveTo")):
                    counts["moves"] += 1
                elif t == w("vMerge"):
                    counts["vMerge"] += 1
                elif t == w("gridSpan") and el.get(w("val")) not in (None, "1"):
                    counts["gridSpan"] += 1
                elif t == w("tbl"):
                    counts["tables"] += 1
                elif t == w("tc"):
                    counts["nested_tables"] += sum(1 for c in el if c.tag == w("tbl"))
                elif t == w("fldChar") and el.get(w("fldCharType")) == "begin" or t == w("fldSimple"):
                    counts["fields"] += 1
                elif t == w("instrText") and el.text and el.text.strip().startswith("TOC"):
                    counts["toc_fields"] += 1
                elif t == w("txbxContent"):
                    counts["textboxes"] += 1
                elif t == f"{{{WP}}}anchor":
                    counts["anchored_drawings"] += 1
                elif t == f"{{{WP}}}inline":
                    counts["inline_drawings"] += 1
                elif t == w("rtl") and el.get(w("val")) in (None, "1", "true", "on"):
                    counts["rtl_runs"] += 1
                elif t in (w("t"), w("delText")) and el.text:
                    counts["cjk_chars"] += len(cjk.findall(el.text))
                    counts["rtl_chars"] += len(rtl.findall(el.text))
                elif local(t) == "blip" and el.get(f"{{{R}}}embed") or local(t) == "imagedata":
                    counts["images"] += 1
                elif t == f"{{{O}}}OLEObject":
                    counts["ole_objects"] += 1
        counts["media_parts"] = sum(1 for n in names if n.startswith("word/media/"))
        counts["charts"] = sum(1 for n in names if re.match(r"word/charts/chart\d+\.xml$", n))
    return counts


def validate(path):
    problems = []
    with zipfile.ZipFile(path) as z:
        names = set(z.namelist())
        ct = ET.fromstring(z.read("[Content_Types].xml"))
        defaults = {e.get("Extension").lower() for e in ct if local(e.tag) == "Default"}
        overrides = {e.get("PartName").lstrip("/") for e in ct if local(e.tag) == "Override"}
        xml_parts = 0
        for n in sorted(names):
            if n not in overrides and ext_of(n) not in defaults:
                problems.append(f"undeclared content type: {n}")
            if n.endswith((".xml", ".rels")):
                try:
                    ET.fromstring(z.read(n))
                    xml_parts += 1
                except ET.ParseError as e:
                    problems.append(f"xml parse error {n}: {e}")
        for n in sorted(names):
            if not n.endswith(".rels"):
                continue
            d, b = posixpath.split(n)
            source = posixpath.join(posixpath.dirname(d), b[: -len(".rels")])
            for r in ET.fromstring(z.read(n)):
                if r.get("TargetMode") == "External":
                    continue
                target = resolve(source, r.get("Target"))
                if target not in names:
                    problems.append(f"missing target {target} from {n}")
        rel_ids = {}
        for part in ("word/document.xml", "word/footnotes.xml", "word/endnotes.xml", "word/numbering.xml"):
            if part not in names:
                continue
            rn = rels_name(part)
            ids = {r.get("Id") for r in ET.fromstring(z.read(rn))} if rn in names else set()
            rel_ids[part] = ids
            root = ET.fromstring(z.read(part))
            unresolved = sum(1 for el in root.iter() for k, v in el.attrib.items() if (ns_of(k) == R or k == f"{{{O}}}relid") and v not in ids)
            if unresolved:
                problems.append(f"{unresolved} unresolved r: references in {part}")
        docpr = []
        bookmarks = []
        para_ids = []
        for n in sorted(names):
            if n.startswith("word/") and n.endswith(".xml") and posixpath.dirname(n) == "word":
                root = ET.fromstring(z.read(n))
                docpr += [el.get("id") for el in root.iter(f"{{{WP}}}docPr")]
                bookmarks += [el.get(w("name")) for el in root.iter(w("bookmarkStart"))]
                para_ids += [v for el in root.iter(w("p")) for k, v in el.attrib.items() if local(k) == "paraId"]
        dup_docpr = len(docpr) - len(set(docpr))
        if dup_docpr:
            problems.append(f"{dup_docpr} duplicate docPr ids")
        dup_bm = len(bookmarks) - len(set(bookmarks))
        dup_para = len(para_ids) - len(set(para_ids))
        media = sum(1 for n in names if n.startswith("word/media/"))
    return {
        "problems": problems, "parts": len(names), "xml_parts": xml_parts, "docpr_ids": len(docpr),
        "duplicate_bookmark_names": dup_bm, "duplicate_paraIds": dup_para, "media": media,
        "bytes": os.path.getsize(path),
    }


def provenance(rows, ids, out_name):
    lines = [
        f"{out_name}",
        "Composite of the public documents below, in this order. The first part is the base: its styles, numbering",
        "definitions, settings, theme, headers and footers apply to the whole document. Each further part starts on a new",
        "page (next-page section break) and keeps its own page setup.",
        "",
    ]
    total = 0
    for i, doc_id in enumerate(ids, 1):
        d = rows[doc_id]
        total += d.get("pages_or_slides") or 0
        prov = d.get("corpus_metadata", {}).get("provenance", {})
        lines += [
            f"[{i}] {doc_id}",
            f"    title: {d.get('title')}",
            f"    source_url: {d.get('source_url')}",
            f"    landing_page: {d.get('landing_page')}",
            f"    publisher: {d.get('publisher')}",
            f"    licence: {d['license']} {LICENCE_URLS[d['license']]}",
            f"    licence declaration: {d.get('license_url')}",
            f"    attribution: {prov.get('attribution') or d.get('attribution')}",
            f"    pages (manifest): {d.get('pages_or_slides')}",
            f"    sha256 of source: {d.get('sha256')}",
            f"    changes: {'base part; combined with the documents below, its body kept and its headers, footers, styles and numbering applied to the whole document' if i == 1 else CHANGES}",
        ]
        if d.get("_stress"):
            lines.append(f"    stress features: {d['_stress']}")
        lines.append("")
    lines.insert(5, f"Summed manifest pages: {total}")
    lines.insert(6, "")
    return "\n".join(lines) + "\n", total


def stress_summary(path):
    c = feature_counts(path)
    keys = [("ins", "ins"), ("del", "del"), ("vMerge", "vMerge"), ("gridSpan", "gridSpan"), ("nested_tables", "nested tables"),
            ("fields", "fields"), ("toc_fields", "TOC"), ("textboxes", "textboxes"), ("anchored_drawings", "anchored drawings"),
            ("footnotes", "footnotes"), ("endnotes", "endnotes"), ("rtl_chars", "RTL chars"), ("cjk_chars", "CJK chars"),
            ("images", "images"), ("charts", "charts"), ("ole_objects", "OLE objects")]
    return ", ".join(f"{label} {c[k]}" for k, label in keys if c[k])


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--corpus", default=os.environ.get("CMP_CORPUS"))
    ap.add_argument("--out", required=True)
    ap.add_argument("ids", nargs="+")
    a = ap.parse_args()
    if not a.corpus:
        ap.error("--corpus or CMP_CORPUS is required")
    rows = load_manifest(a.corpus)
    for doc_id in a.ids:
        if doc_id not in rows:
            raise SystemExit(f"unknown id {doc_id}")
        if rows[doc_id]["license"] not in ALLOWED_LICENCES:
            raise SystemExit(f"{doc_id}: licence {rows[doc_id]['license']} not allowed")
    paths = [os.path.join(a.corpus, rows[i]["path"]) for i in a.ids]
    for doc_id, p in zip(a.ids, paths):
        rows[doc_id]["_stress"] = stress_summary(p)
    composer = Composer(Package(paths[0]))
    for n, p in enumerate(paths[1:], 1):
        composer.append(Package(p), n == len(paths) - 1)
    composer.write(a.out)
    text, total = provenance(rows, a.ids, os.path.basename(a.out))
    prov_path = os.path.splitext(a.out)[0] + ".PROVENANCE.txt"
    with open(prov_path, "w", encoding="utf-8") as f:
        f.write(text)
    report = validate(a.out)
    report["summed_manifest_pages"] = total
    report["dangling_source_targets"] = sorted(set(composer.dangling))
    report["features"] = feature_counts(a.out)
    print(json.dumps(report, indent=2))
    sys.exit(1 if report["problems"] else 0)


if __name__ == "__main__":
    main()
