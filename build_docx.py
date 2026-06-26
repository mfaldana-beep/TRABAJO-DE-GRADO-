#!/usr/bin/env python3
"""Conversor Markdown -> DOCX en Python puro (stdlib), con formato APA 7.
Genera un .docx valido (OOXML) sin dependencias externas.
"""
import re
import zipfile
import datetime

SRC = "Tesis_IA_Justicia_Penal_Garantista.md"
OUT = "Tesis_IA_Justicia_Penal_Garantista.docx"


def esc(s):
    return (s.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")
             .replace('"', "&quot;"))


# ---------- inline formatting ----------
def parse_inline(text, base_bold=False, base_italic=False, sz=None, font=None):
    """Return list of run XML strings from a line of markdown inline text."""
    # links [txt](url) -> "txt (url)"
    text = re.sub(r"\[([^\]]+)\]\((https?://[^)]+)\)", r"\1 (\2)", text)
    # checkbox
    text = text.replace("[ ]", "\u2610").replace("[x]", "\u2612")
    runs = []
    i = 0
    n = len(text)
    buf = ""
    bold = base_bold
    italic = base_italic
    code = False

    def flush():
        nonlocal buf
        if buf:
            runs.append(make_run(buf, bold, italic, code, sz, font))
            buf = ""

    while i < n:
        if text.startswith("**", i):
            flush(); bold = not bold; i += 2; continue
        if text[i] == "*":
            flush(); italic = not italic; i += 1; continue
        if text[i] == "`":
            flush(); code = not code; i += 1; continue
        buf += text[i]; i += 1
    flush()
    return runs


def make_run(t, bold, italic, code, sz, font):
    rpr = ""
    props = ""
    if font:
        props += f'<w:rFonts w:ascii="{font}" w:hAnsi="{font}" w:cs="{font}"/>'
    elif code:
        props += '<w:rFonts w:ascii="Courier New" w:hAnsi="Courier New" w:cs="Courier New"/>'
    if bold:
        props += "<w:b/>"
    if italic:
        props += "<w:i/>"
    if sz:
        props += f'<w:sz w:val="{sz}"/><w:szCs w:val="{sz}"/>'
    if props:
        rpr = f"<w:rPr>{props}</w:rPr>"
    return f'<w:r>{rpr}<w:t xml:space="preserve">{esc(t)}</w:t></w:r>'


def para(runs_xml, align=None, line=None, before=None, after=None,
         indent_first=None, keep=False, border_bottom=False):
    ppr = "<w:pPr>"
    spacing = '<w:spacing'
    spacing += f' w:before="{before}"' if before is not None else ' w:before="0"'
    spacing += f' w:after="{after}"' if after is not None else ' w:after="0"'
    if line is not None:
        spacing += f' w:line="{line}" w:lineRule="auto"'
    spacing += "/>"
    ppr += spacing
    if border_bottom:
        ppr += '<w:pBdr><w:bottom w:val="single" w:sz="6" w:space="1" w:color="999999"/></w:pBdr>'
    if align:
        ppr += f'<w:jc w:val="{align}"/>'
    if indent_first:
        ppr += f'<w:ind w:firstLine="{indent_first}"/>'
    if keep:
        ppr += "<w:keepNext/>"
    ppr += "</w:pPr>"
    return f"<w:p>{ppr}{runs_xml}</w:p>"


def page_break():
    return '<w:p><w:r><w:br w:type="page"/></w:r></w:p>'


HEAD = {
    1: dict(sz=32, bold=True, align="center", before=240, after=160),
    2: dict(sz=28, bold=True, align="left", before=200, after=120),
    3: dict(sz=26, bold=True, align="left", before=160, after=100),
    4: dict(sz=24, bold=True, italic=True, align="left", before=120, after=80),
    5: dict(sz=24, bold=True, italic=True, align="left", before=120, after=80),
    6: dict(sz=24, bold=True, italic=True, align="left", before=120, after=80),
}


def build_table(rows):
    # rows: list of list of cell-strings; first row is header
    ncol = max(len(r) for r in rows)
    grid = "".join('<w:gridCol w:w="%d"/>' % int(9360 / ncol) for _ in range(ncol))
    borders = ('<w:tblBorders>'
               '<w:top w:val="single" w:sz="4" w:space="0" w:color="666666"/>'
               '<w:left w:val="single" w:sz="4" w:space="0" w:color="666666"/>'
               '<w:bottom w:val="single" w:sz="4" w:space="0" w:color="666666"/>'
               '<w:right w:val="single" w:sz="4" w:space="0" w:color="666666"/>'
               '<w:insideH w:val="single" w:sz="4" w:space="0" w:color="999999"/>'
               '<w:insideV w:val="single" w:sz="4" w:space="0" w:color="999999"/>'
               '</w:tblBorders>')
    tblpr = ('<w:tblPr><w:tblW w:w="9360" w:type="dxa"/>'
             '<w:tblLayout w:type="fixed"/>' + borders +
             '<w:tblCellMar><w:top w:w="40" w:type="dxa"/><w:left w:w="80" w:type="dxa"/>'
             '<w:bottom w:w="40" w:type="dxa"/><w:right w:w="80" w:type="dxa"/></w:tblCellMar>'
             '</w:tblPr>')
    out = [f'<w:tbl>{tblpr}<w:tblGrid>{grid}</w:tblGrid>']
    for ri, row in enumerate(rows):
        cells = list(row) + [""] * (ncol - len(row))
        out.append("<w:tr>")
        for c in cells:
            header = ri == 0
            runs = parse_inline(c.strip(), base_bold=header, sz=22)
            shade = '<w:shd w:val="clear" w:color="auto" w:fill="E7E6E6"/>' if header else ""
            tcpr = f'<w:tcPr><w:tcW w:w="{int(9360/ncol)}" w:type="dxa"/>{shade}<w:vAlign w:val="center"/></w:tcPr>'
            cellpara = para("".join(runs) or "", line=240, after=20)
            out.append(f"<w:tc>{tcpr}{cellpara}</w:tc>")
        out.append("</w:tr>")
    out.append("</w:tbl>")
    # spacer paragraph after table
    out.append(para("", after=120, line=240))
    return "".join(out)


def convert(md):
    lines = md.split("\n")
    body = []
    i = 0
    n = len(lines)
    center = False
    in_comment = False
    while i < n:
        line = lines[i]
        stripped = line.strip()

        # HTML comments (single or multi-line)
        if not in_comment and stripped.startswith("<!--"):
            if "-->" in stripped:
                i += 1; continue
            in_comment = True; i += 1; continue
        if in_comment:
            if "-->" in stripped:
                in_comment = False
            i += 1; continue

        # div align center toggles
        if stripped.startswith('<div align="center"'):
            center = True; i += 1; continue
        if stripped.startswith("</div>"):
            center = False; i += 1; continue

        # page break
        if stripped == "\\newpage":
            body.append(page_break()); i += 1; continue

        # fenced code block
        if stripped.startswith("```"):
            i += 1
            code_lines = []
            while i < n and not lines[i].strip().startswith("```"):
                code_lines.append(lines[i]); i += 1
            i += 1  # skip closing fence
            for cl in code_lines:
                run = make_run(cl if cl else " ", False, False, True, 16, "Courier New")
                body.append(para(run, line=240, after=0))
            body.append(para("", after=80, line=240))
            continue

        # blank line
        if stripped == "":
            i += 1; continue

        # horizontal rule
        if stripped in ("---", "***", "___"):
            body.append(para("", border_bottom=True, after=120, line=240))
            i += 1; continue

        # table block
        if stripped.startswith("|") and i + 1 < n and re.match(r"^\|?[\s:|-]+\|?$", lines[i+1].strip()):
            tbl_lines = []
            while i < n and lines[i].strip().startswith("|"):
                tbl_lines.append(lines[i].strip()); i += 1
            rows = []
            for idx, tl in enumerate(tbl_lines):
                if idx == 1:
                    continue  # separator
                cells = tl.strip().strip("|").split("|")
                rows.append([c.strip() for c in cells])
            body.append(build_table(rows))
            continue

        # headings
        m = re.match(r"^(#{1,6})\s+(.*)$", stripped)
        if m:
            lvl = len(m.group(1))
            h = HEAD[lvl]
            runs = parse_inline(m.group(2), base_bold=h.get("bold", False),
                                base_italic=h.get("italic", False), sz=h["sz"])
            body.append(para("".join(runs), align=h["align"], line=240,
                             before=h["before"], after=h["after"], keep=True))
            i += 1; continue

        # blockquote
        if stripped.startswith(">"):
            qtext = stripped.lstrip(">").strip()
            runs = parse_inline(qtext, base_italic=True, sz=22)
            body.append(para("".join(runs), align="left", line=360, before=40,
                             after=120, indent_first=None))
            i += 1; continue

        # list item (ordered/unordered/checkbox)
        mlist = re.match(r"^(\d+)\.\s+(.*)$", stripped)
        mbul = re.match(r"^[-*]\s+(.*)$", stripped)
        if mlist:
            txt = mlist.group(1) + ". " + mlist.group(2)
            runs = parse_inline(txt)
            body.append(para("".join(runs), line=360, after=40,
                             indent_first=None))
            # manual hanging indent via leading
            body[-1] = body[-1].replace("<w:pPr>", '<w:pPr><w:ind w:left="360"/>')
            i += 1; continue
        if mbul:
            txt = "\u2022  " + mbul.group(1)
            runs = parse_inline(txt)
            body.append(para("".join(runs), line=360, after=40))
            body[-1] = body[-1].replace("<w:pPr>", '<w:pPr><w:ind w:left="360"/>')
            i += 1; continue

        # references section -> hanging indent paragraphs (detect by heading already passed)
        # normal paragraph
        runs = parse_inline(stripped)
        align = "center" if center else "both"
        body.append(para("".join(runs), align=align, line=480, after=120))
        i += 1
    return "".join(body)


def hanging_for_references(body_xml):
    return body_xml  # references handled inline; keep simple


# ---------- package assembly ----------
def main():
    with open(SRC, encoding="utf-8") as f:
        md = f.read()
    body = convert(md)

    sect = ('<w:sectPr>'
            '<w:headerReference w:type="default" r:id="rIdHdr"/>'
            '<w:pgSz w:w="12240" w:h="15840"/>'
            '<w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" '
            'w:header="720" w:footer="720" w:gutter="0"/>'
            '</w:sectPr>')

    document = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" '
        'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
        f'<w:body>{body}{sect}</w:body></w:document>'
    )

    styles = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
        '<w:docDefaults><w:rPrDefault><w:rPr>'
        '<w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman" w:cs="Times New Roman"/>'
        '<w:sz w:val="24"/><w:szCs w:val="24"/><w:lang w:val="es-CO"/>'
        '</w:rPr></w:rPrDefault>'
        '<w:pPrDefault><w:pPr><w:spacing w:after="0" w:line="480" w:lineRule="auto"/></w:pPr></w:pPrDefault>'
        '</w:docDefaults>'
        '<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style>'
        '</w:styles>'
    )

    header = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<w:hdr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
        '<w:p><w:pPr><w:jc w:val="right"/><w:spacing w:line="240" w:lineRule="auto"/></w:pPr>'
        '<w:r><w:fldChar w:fldCharType="begin"/></w:r>'
        '<w:r><w:instrText xml:space="preserve"> PAGE </w:instrText></w:r>'
        '<w:r><w:fldChar w:fldCharType="end"/></w:r></w:p></w:hdr>'
    )

    content_types = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
        '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
        '<Default Extension="xml" ContentType="application/xml"/>'
        '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>'
        '<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>'
        '<Override PartName="/word/header1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"/>'
        '<Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>'
        '<Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>'
        '</Types>'
    )

    rels = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>'
        '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>'
        '<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/>'
        '</Relationships>'
    )

    doc_rels = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>'
        '<Relationship Id="rIdHdr" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/header" Target="header1.xml"/>'
        '</Relationships>'
    )

    now = datetime.datetime.utcnow().strftime("%Y-%m-%dT%H:%M:%SZ")
    core = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" '
        'xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" '
        'xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">'
        '<dc:title>La inteligencia artificial en la administracion de justicia con enfoque garantista</dc:title>'
        '<dc:creator>Maria Fernanda Aldana Rojas; Jesus Esneider Prieto Soto</dc:creator>'
        f'<dcterms:created xsi:type="dcterms:W3CDTF">{now}</dcterms:created>'
        f'<dcterms:modified xsi:type="dcterms:W3CDTF">{now}</dcterms:modified>'
        '</cp:coreProperties>'
    )
    app = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties">'
        '<Application>Kiro DOCX Builder</Application></Properties>'
    )

    with zipfile.ZipFile(OUT, "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr("[Content_Types].xml", content_types)
        z.writestr("_rels/.rels", rels)
        z.writestr("word/document.xml", document)
        z.writestr("word/styles.xml", styles)
        z.writestr("word/header1.xml", header)
        z.writestr("word/_rels/document.xml.rels", doc_rels)
        z.writestr("docProps/core.xml", core)
        z.writestr("docProps/app.xml", app)
    print("OK ->", OUT)


if __name__ == "__main__":
    main()
