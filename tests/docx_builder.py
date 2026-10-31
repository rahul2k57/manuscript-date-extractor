"""Tiny helper that builds minimal .docx files for tests (standard library only)."""
import zipfile
from xml.sax.saxutils import escape

NS = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'


def para(text, split=False):
    """One paragraph. split=True puts 'Label:' and the value in separate runs."""
    if split and ":" in text:
        label, value = text.split(":", 1)
        return (f'<w:p><w:r><w:rPr><w:b/></w:rPr><w:t>{escape(label)}:</w:t></w:r>'
                f'<w:r><w:t xml:space="preserve">{escape(value)}</w:t></w:r></w:p>')
    return f'<w:p><w:r><w:t xml:space="preserve">{escape(text)}</w:t></w:r></w:p>'


def table(rows):
    xml = "<w:tbl>"
    for row in rows:
        xml += "<w:tr>" + "".join(f"<w:tc>{para(c)}</w:tc>" for c in row) + "</w:tr>"
    return xml + "</w:tbl>"


def make_docx(path, items, header=None, footer=None):
    """items: list of str | ("split", str) | ("table", [[cell, ...], ...])"""
    body = ""
    for it in items:
        if isinstance(it, str):
            body += para(it)
        elif it[0] == "split":
            body += para(it[1], split=True)
        elif it[0] == "table":
            body += table(it[1])
    with zipfile.ZipFile(path, "w") as z:
        z.writestr("[Content_Types].xml",
                   '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"/>')
        z.writestr("word/document.xml", f"<w:document {NS}><w:body>{body}</w:body></w:document>")
        if header:
            z.writestr("word/header1.xml", f"<w:hdr {NS}>{para(header)}</w:hdr>")
        if footer:
            z.writestr("word/footer1.xml", f"<w:ftr {NS}>{para(footer)}</w:ftr>")