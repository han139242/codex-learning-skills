"""Read a DOCX without executing content or modifying the source; stdlib only."""
import argparse
import difflib
import hashlib
import io
import json
import re
from pathlib import Path
from zipfile import ZipFile
import xml.etree.ElementTree as ET

NS = {"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
      "m": "http://schemas.openxmlformats.org/officeDocument/2006/math"}
W, M = NS["w"], NS["m"]
def q(name):
    return "{" + W + "}" + name

def visible(element):
    if element.tag in (q("del"), q("moveFrom")):
        return ""
    if element.tag in (q("t"), "{" + M + "}t"):
        return element.text or ""
    if element.tag == q("tab"):
        return "\t"
    if element.tag in (q("br"), q("cr")):
        return "\n"
    return "".join(visible(child) for child in element)

def digest(data):
    return hashlib.sha256(data).hexdigest()

def scan(path):
    path = Path(path).resolve(strict=True)
    payload = path.read_bytes()
    with ZipFile(io.BytesIO(payload)) as archive:
        bad = archive.testzip()
        if bad:
            raise ValueError("ZIP CRC failure: " + bad)
        names = {}
        for raw in archive.namelist():
            name = raw.replace("\\", "/")
            if name in names:
                raise ValueError("Duplicate normalized ZIP member: " + name)
            names[name] = raw
        def read(name):
            return archive.read(names[name])
        document = ET.fromstring(read("word/document.xml"))
        styles = {}
        if "word/styles.xml" in names:
            for style in ET.fromstring(read("word/styles.xml")).findall("w:style", NS):
                sid = style.get(q("styleId"))
                title = style.find("w:name", NS)
                base = style.find("w:basedOn", NS)
                level = style.find("w:pPr/w:outlineLvl", NS)
                styles[sid] = {"name": title.get(q("val"), sid) if title is not None else sid,
                               "base": base.get(q("val")) if base is not None else None,
                               "outline": int(level.get(q("val"))) if level is not None else None}
        def outline(sid, seen=None):
            seen = set() if seen is None else seen
            if sid not in styles or sid in seen:
                return None
            seen.add(sid)
            style = styles[sid]
            return style["outline"] if style["outline"] is not None else outline(style["base"], seen)
        blocks, headings, codes = [], [], []
        body = document.find("w:body", NS)
        if body is None:
            raise ValueError("Missing Word body")
        for i, element in enumerate(body):
            if element.tag == q("p"):
                node = element.find("w:pPr/w:pStyle", NS)
                sid = node.get(q("val"), "") if node is not None else ""
                direct_level = element.find("w:pPr/w:outlineLvl", NS)
                level = int(direct_level.get(q("val"))) if direct_level is not None else outline(sid)
                block = {"index": i, "kind": "paragraph", "style_id": sid,
                         "style_name": styles.get(sid, {}).get("name", sid), "text": visible(element)}
                blocks.append(block)
                if level is not None and level < 9:
                    headings.append({"index": i, "level": level + 1, "text": block["text"]})
                if re.search(r"source\s*code|^code$|代码", block["style_name"], re.I):
                    codes.append({"index": i, "style": sid, "text": block["text"]})
            elif element.tag == q("tbl"):
                rows = []
                for row in element.findall("w:tr", NS):
                    cells = []
                    for cell in row.findall("w:tc", NS):
                        paragraphs = []
                        for paragraph in cell.findall("w:p", NS):
                            align = paragraph.find("w:pPr/w:jc", NS)
                            bold = paragraph.findall(".//w:rPr/w:b", NS)
                            paragraphs.append({"text": visible(paragraph),
                                               "direct_alignment": align.get(q("val")) if align is not None else None,
                                               "direct_bold_runs": sum(b.get(q("val"), "1") not in ("0", "false", "off") for b in bold)})
                        cells.append({"text": "\n".join(p["text"] for p in paragraphs), "paragraphs": paragraphs})
                    rows.append(cells)
                blocks.append({"index": i, "kind": "table", "rows": rows})
            else:
                content = visible(element)
                if content:
                    blocks.append({"index": i, "kind": element.tag.rsplit("}", 1)[-1], "text": content})
        fields, external = [], []
        for name in names:
            if name.startswith("word/") and name.endswith(".xml"):
                part = ET.fromstring(read(name))
                fields.extend({"part": name, "instruction": e.text or ""} for e in part.findall(".//w:instrText", NS))
                fields.extend({"part": name, "instruction": e.get(q("instr"), "")} for e in part.findall(".//w:fldSimple", NS))
            if name.endswith(".rels"):
                for rel in ET.fromstring(read(name)):
                    if rel.get("TargetMode") == "External":
                        external.append({"part": name, "type": rel.get("Type"), "target": rel.get("Target")})
        media = {n: digest(read(n)) for n in names if n.startswith("word/media/") and not n.endswith("/")}
        text = "\n".join(b.get("text", "\n".join(" | ".join(c["text"] for c in row) for row in b.get("rows", []))) for b in blocks)
        issues = re.findall(r"\{\{[^}]*\}\}|TODO|待填写|Error!|错误!|未定义书签", text)
        if digest(path.read_bytes()) != digest(payload):
            raise ValueError("Input changed during inspection; take a stable snapshot and retry")
        return {"input": str(path), "sha256": digest(payload),
                "scope": {"source_modified": False, "code_executed": False, "render_reviewed": False, "calculations_verified": False},
                "counts": {"paragraphs": len(document.findall(".//w:p", NS)), "tables": len(document.findall(".//w:tbl", NS)),
                           "equations": len(document.findall(".//m:oMath", NS)), "images": len(document.findall(".//w:drawing", NS)),
                           "insertions": len(document.findall(".//w:ins", NS)), "deletions": len(document.findall(".//w:del", NS)),
                           "comment_anchors": len(document.findall(".//w:commentReference", NS))},
                "headings": headings, "fields": fields, "media_sha256": media, "external_relationships": external,
                "code_paragraphs": codes, "suspected_placeholders": issues, "blocks": blocks, "text": text,
                "notes": ["Direct alignment may inherit a style when absent; this is not an alignment defect.",
                          "Flattened equation text is for comparison only; inspect OMML and rendered pages for notation."]}

def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("input", type=Path)
    parser.add_argument("--out", type=Path, required=True)
    parser.add_argument("--baseline", type=Path)
    args = parser.parse_args()
    source = args.input.resolve(strict=True)
    out = args.out.resolve()
    if out == source or (args.baseline and out == args.baseline.resolve()):
        parser.error("Output must not be an input document")
    if out.exists():
        parser.error("Output already exists; choose a new temporary filename")
    report = scan(source)
    if args.baseline:
        baseline = scan(args.baseline)
        old, new = baseline["media_sha256"], report["media_sha256"]
        report["comparison"] = {"baseline": baseline["input"], "baseline_sha256": baseline["sha256"],
                                "baseline_counts": baseline["counts"],
                                "media_changed": [n for n in sorted(set(old) | set(new)) if old.get(n) != new.get(n)],
                                "code_changed": [p["text"] for p in baseline["code_paragraphs"]] != [p["text"] for p in report["code_paragraphs"]],
                                "text_diff": "\n".join(difflib.unified_diff(baseline["text"].splitlines(), report["text"].splitlines(), fromfile=baseline["input"], tofile=report["input"], n=2))}
    with out.open("x", encoding="utf-8") as stream:
        json.dump(report, stream, ensure_ascii=False, indent=2)
    print(json.dumps({"audit": str(out), "counts": report["counts"], "scope": report["scope"]}, ensure_ascii=False))

if __name__ == "__main__":
    main()
