#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Minimal Markdown -> DOCX converter using only the Python standard library.
Targets the specific Markdown in this repo: headings (# ~ ####), paragraphs,
pipe tables, fenced code blocks, bold (**), inline code (`), blockquotes,
footnote-style references <sup>[n]</sup>, and horizontal rules (---).
Produces a .docx that opens in Word / WPS / LibreOffice.
"""
import os
import re
import zipfile
from xml.sax.saxutils import escape

SRC = "基于Gross理论的青少年非医疗化陶艺疗愈APP设计研究.md"
DST = "基于Gross理论的青少年非医疗化陶艺疗愈APP设计研究.docx"

W_NS = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'


def x(s: str) -> str:
    return escape(s, {'"': "&quot;", "'": "&apos;"})


# ---------- Inline formatting ----------

INLINE_RE = re.compile(
    r"(\*\*(?P<bold>[^*]+)\*\*)"
    r"|(`(?P<code>[^`]+)`)"
    r"|(<sup>(?P<sup>[^<]+)</sup>)"
)


def inline_runs(text: str) -> str:
    """Return XML for one paragraph's inline runs."""
    out = []
    pos = 0
    for m in INLINE_RE.finditer(text):
        if m.start() > pos:
            out.append(run(text[pos:m.start()]))
        if m.group("bold") is not None:
            out.append(run(m.group("bold"), bold=True))
        elif m.group("code") is not None:
            out.append(run(m.group("code"), mono=True))
        elif m.group("sup") is not None:
            out.append(run(m.group("sup"), sup=True))
        pos = m.end()
    if pos < len(text):
        out.append(run(text[pos:]))
    return "".join(out)


def run(text: str, bold: bool = False, mono: bool = False, sup: bool = False,
        size: int = 21) -> str:
    rpr = []
    if bold:
        rpr.append("<w:b/>")
    if mono:
        rpr.append('<w:rFonts w:ascii="Consolas" w:hAnsi="Consolas" w:eastAsia="宋体"/>')
    else:
        rpr.append('<w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman" w:eastAsia="宋体"/>')
    if sup:
        rpr.append('<w:vertAlign w:val="superscript"/>')
    rpr.append(f'<w:sz w:val="{size}"/><w:szCs w:val="{size}"/>')
    rpr_xml = f"<w:rPr>{''.join(rpr)}</w:rPr>"
    # preserve spaces
    return (
        f'<w:r>{rpr_xml}'
        f'<w:t xml:space="preserve">{x(text)}</w:t>'
        f"</w:r>"
    )


def para(text: str, style: str = None, align: str = None,
         first_line_indent: bool = True) -> str:
    ppr = []
    if style:
        ppr.append(f'<w:pStyle w:val="{style}"/>')
    if align:
        ppr.append(f'<w:jc w:val="{align}"/>')
    # Chinese-style body: 1.5 line, first-line indent 2 chars, space after 0
    if style is None:
        ppr.append('<w:spacing w:line="360" w:lineRule="auto" w:after="0"/>')
        if first_line_indent:
            ppr.append('<w:ind w:firstLineChars="200" w:firstLine="480"/>')
    ppr_xml = f"<w:pPr>{''.join(ppr)}</w:pPr>" if ppr else ""
    return f"<w:p>{ppr_xml}{inline_runs(text)}</w:p>"


def heading(text: str, level: int) -> str:
    # Map to built-in styles Heading1..Heading4; align left, bold, sized
    size_map = {1: 36, 2: 32, 3: 28, 4: 24}
    sz = size_map.get(level, 24)
    ppr = (
        f'<w:pPr>'
        f'<w:pStyle w:val="Heading{level}"/>'
        f'<w:spacing w:before="240" w:after="120" w:line="360" w:lineRule="auto"/>'
        f'<w:outlineLvl w:val="{level-1}"/>'
        f"</w:pPr>"
    )
    rpr = (
        f"<w:rPr>"
        f'<w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman" w:eastAsia="黑体"/>'
        f"<w:b/>"
        f'<w:sz w:val="{sz}"/><w:szCs w:val="{sz}"/>'
        f"</w:rPr>"
    )
    return (
        f"<w:p>{ppr}<w:r>{rpr}"
        f'<w:t xml:space="preserve">{x(text)}</w:t>'
        f"</w:r></w:p>"
    )


def hrule() -> str:
    return (
        "<w:p>"
        "<w:pPr><w:pBdr>"
        '<w:bottom w:val="single" w:sz="6" w:space="1" w:color="auto"/>'
        "</w:pBdr></w:pPr></w:p>"
    )


# ---------- Tables ----------

def build_table(rows):
    """rows: list of list of cells (strings). First row is header."""
    # compute grid
    cols = max(len(r) for r in rows)
    grid = "".join(f'<w:gridCol w:w="{int(9000/cols)}"/>' for _ in range(cols))
    tbl_pr = (
        "<w:tblPr>"
        '<w:tblW w:w="9000" w:type="dxa"/>'
        "<w:tblBorders>"
        '<w:top w:val="single" w:sz="4" w:color="auto"/>'
        '<w:left w:val="single" w:sz="4" w:color="auto"/>'
        '<w:bottom w:val="single" w:sz="4" w:color="auto"/>'
        '<w:right w:val="single" w:sz="4" w:color="auto"/>'
        '<w:insideH w:val="single" w:sz="4" w:color="auto"/>'
        '<w:insideV w:val="single" w:sz="4" w:color="auto"/>'
        "</w:tblBorders>"
        '<w:tblLayout w:type="fixed"/>'
        "</w:tblPr>"
        f"<w:tblGrid>{grid}</w:tblGrid>"
    )
    body = []
    for idx, cells in enumerate(rows):
        is_header = idx == 0
        body.append("<w:tr>")
        if is_header:
            body.append('<w:trPr><w:tblHeader/></w:trPr>')
        for c in cells:
            shd = '<w:shd w:val="clear" w:color="auto" w:fill="EFEFEF"/>' if is_header else ""
            width = int(9000 / cols)
            ppr = (
                "<w:pPr>"
                '<w:spacing w:line="300" w:lineRule="auto" w:after="0"/>'
                '<w:jc w:val="left"/>'
                "</w:pPr>"
            )
            body.append(
                f"<w:tc>"
                f'<w:tcPr><w:tcW w:w="{width}" w:type="dxa"/>{shd}</w:tcPr>'
                f"<w:p>{ppr}{inline_runs(c.strip())}</w:p>"
                f"</w:tc>"
            )
        body.append("</w:tr>")
    return f"<w:tbl>{tbl_pr}{''.join(body)}</w:tbl>"


# ---------- Code block ----------

def code_block(lines):
    out = []
    for ln in lines:
        ppr = (
            "<w:pPr>"
            '<w:spacing w:line="260" w:lineRule="auto" w:after="0"/>'
            '<w:shd w:val="clear" w:color="auto" w:fill="F4F4F4"/>'
            "</w:pPr>"
        )
        rpr = (
            '<w:rPr><w:rFonts w:ascii="Consolas" w:hAnsi="Consolas" w:eastAsia="Consolas"/>'
            '<w:sz w:val="20"/><w:szCs w:val="20"/></w:rPr>'
        )
        out.append(
            f"<w:p>{ppr}<w:r>{rpr}"
            f'<w:t xml:space="preserve">{x(ln)}</w:t>'
            f"</w:r></w:p>"
        )
    return "".join(out)


# ---------- Parser ----------

def parse_md(text):
    lines = text.splitlines()
    body_parts = []
    i = 0
    while i < len(lines):
        line = lines[i]

        # fenced code
        if line.strip().startswith("```"):
            j = i + 1
            buf = []
            while j < len(lines) and not lines[j].strip().startswith("```"):
                buf.append(lines[j])
                j += 1
            body_parts.append(code_block(buf))
            i = j + 1
            continue

        # blank
        if not line.strip():
            i += 1
            continue

        # horizontal rule
        if re.match(r"^-{3,}$", line.strip()):
            body_parts.append(hrule())
            i += 1
            continue

        # heading
        m = re.match(r"^(#{1,4})\s+(.*)$", line)
        if m:
            level = len(m.group(1))
            body_parts.append(heading(m.group(2).strip(), level))
            i += 1
            continue

        # table: line contains | and next line is divider
        if "|" in line and i + 1 < len(lines) and re.match(r"^\s*\|?[\s:\-|]+\|?\s*$", lines[i+1]):
            rows = []
            while i < len(lines) and "|" in lines[i]:
                if re.match(r"^\s*\|?[\s:\-|]+\|?\s*$", lines[i]):
                    i += 1
                    continue
                cells = [c.strip() for c in lines[i].strip().strip("|").split("|")]
                rows.append(cells)
                i += 1
            body_parts.append(build_table(rows))
            continue

        # blockquote
        if line.startswith(">"):
            buf = []
            while i < len(lines) and lines[i].startswith(">"):
                buf.append(lines[i].lstrip(">").strip())
                i += 1
            text_block = " ".join(buf)
            ppr = (
                "<w:pPr>"
                '<w:ind w:left="480"/>'
                '<w:spacing w:line="340" w:lineRule="auto" w:after="0"/>'
                '<w:pBdr><w:left w:val="single" w:sz="12" w:color="AAAAAA"/></w:pBdr>'
                "</w:pPr>"
            )
            body_parts.append(
                f'<w:p>{ppr}{inline_runs(text_block)}</w:p>'
            )
            continue

        # plain paragraph (may span until blank/special line)
        buf = [line]
        i += 1
        while i < len(lines):
            nxt = lines[i]
            if (not nxt.strip()
                    or nxt.startswith("#")
                    or nxt.startswith(">")
                    or nxt.strip().startswith("```")
                    or re.match(r"^-{3,}$", nxt.strip())
                    or ("|" in nxt and i + 1 < len(lines)
                        and re.match(r"^\s*\|?[\s:\-|]+\|?\s*$", lines[i+1]))):
                break
            buf.append(nxt)
            i += 1
        paragraph_text = " ".join(b.strip() for b in buf)
        # detect list-like lines starting with "（x）" or "[x]" – keep them as paragraphs
        body_parts.append(para(paragraph_text))

    return "".join(body_parts)


# ---------- DOCX packaging ----------

CONTENT_TYPES = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
<Default Extension="xml" ContentType="application/xml"/>
<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>
<Override PartName="/word/settings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/>
<Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>
<Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>
</Types>'''

RELS = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>
<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/>
</Relationships>'''

DOC_RELS = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings" Target="settings.xml"/>
</Relationships>'''

STYLES = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
<w:docDefaults>
  <w:rPrDefault><w:rPr>
    <w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman" w:eastAsia="宋体" w:cs="Times New Roman"/>
    <w:sz w:val="21"/><w:szCs w:val="21"/><w:lang w:val="en-US" w:eastAsia="zh-CN"/>
  </w:rPr></w:rPrDefault>
  <w:pPrDefault><w:pPr>
    <w:spacing w:line="360" w:lineRule="auto"/>
  </w:pPr></w:pPrDefault>
</w:docDefaults>
<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style>
<w:style w:type="paragraph" w:styleId="Heading1"><w:name w:val="heading 1"/><w:basedOn w:val="Normal"/><w:next w:val="Normal"/><w:pPr><w:outlineLvl w:val="0"/></w:pPr></w:style>
<w:style w:type="paragraph" w:styleId="Heading2"><w:name w:val="heading 2"/><w:basedOn w:val="Normal"/><w:next w:val="Normal"/><w:pPr><w:outlineLvl w:val="1"/></w:pPr></w:style>
<w:style w:type="paragraph" w:styleId="Heading3"><w:name w:val="heading 3"/><w:basedOn w:val="Normal"/><w:next w:val="Normal"/><w:pPr><w:outlineLvl w:val="2"/></w:pPr></w:style>
<w:style w:type="paragraph" w:styleId="Heading4"><w:name w:val="heading 4"/><w:basedOn w:val="Normal"/><w:next w:val="Normal"/><w:pPr><w:outlineLvl w:val="3"/></w:pPr></w:style>
</w:styles>'''

SETTINGS = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
<w:zoom w:percent="100"/>
<w:defaultTabStop w:val="420"/>
<w:characterSpacingControl w:val="compressPunctuation"/>
<w:compat><w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="15"/></w:compat>
</w:settings>'''

CORE = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">
<dc:title>基于Gross理论的青少年非医疗化陶艺疗愈APP设计研究</dc:title>
<dc:creator>Kiro</dc:creator>
<cp:lastModifiedBy>Kiro</cp:lastModifiedBy>
</cp:coreProperties>'''

APP = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties">
<Application>Kiro Markdown Converter</Application>
</Properties>'''


def build_docx(md_text: str, out_path: str):
    body = parse_md(md_text)
    doc_xml = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        f'<w:document {W_NS}>'
        "<w:body>"
        f"{body}"
        '<w:sectPr>'
        '<w:pgSz w:w="11906" w:h="16838"/>'
        '<w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" '
        'w:header="720" w:footer="720" w:gutter="0"/>'
        "</w:sectPr>"
        "</w:body>"
        "</w:document>"
    )
    with zipfile.ZipFile(out_path, "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr("[Content_Types].xml", CONTENT_TYPES)
        z.writestr("_rels/.rels", RELS)
        z.writestr("word/_rels/document.xml.rels", DOC_RELS)
        z.writestr("word/document.xml", doc_xml)
        z.writestr("word/styles.xml", STYLES)
        z.writestr("word/settings.xml", SETTINGS)
        z.writestr("docProps/core.xml", CORE)
        z.writestr("docProps/app.xml", APP)


if __name__ == "__main__":
    here = os.path.dirname(os.path.abspath(__file__))
    src_path = os.path.join(here, SRC)
    dst_path = os.path.join(here, DST)
    with open(src_path, "r", encoding="utf-8") as f:
        md = f.read()
    build_docx(md, dst_path)
    print(f"Wrote {dst_path} ({os.path.getsize(dst_path)} bytes)")
