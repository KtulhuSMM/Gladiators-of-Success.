#!/usr/bin/env python3
"""Build the editable A5 interior and its PDF from manuscript.md."""
from pathlib import Path
import uno
from com.sun.star.beans import PropertyValue
from com.sun.star.style.BreakType import PAGE_BEFORE
from com.sun.star.text.ControlCharacter import PARAGRAPH_BREAK

ROOT = Path(__file__).resolve().parents[1]
SOURCE = ROOT / "manuscript.md"
OUT = ROOT / "print" / "output"


def prop(name, value):
    item = PropertyValue()
    item.Name = name
    item.Value = value
    return item


def file_url(path):
    return uno.systemPathToFileUrl(str(path.resolve()))


def add_paragraph(doc, text, style):
    cursor = doc.Text.createTextCursor()
    cursor.gotoEnd(False)
    cursor.ParaStyleName = style
    doc.Text.insertString(cursor, text, False)
    doc.Text.insertControlCharacter(cursor, PARAGRAPH_BREAK, False)


def set_style(styles, name, *, font=None, size=None, before=None, after=None,
              page_before=False, keep=False, italic=False, bold=False,
              outline=0, align=None):
    style = styles.getByName(name)
    if font:
        style.CharFontName = font
    if size:
        style.CharHeight = size
    if before is not None:
        style.ParaTopMargin = before
    if after is not None:
        style.ParaBottomMargin = after
    if page_before:
        style.BreakType = PAGE_BEFORE
    if keep:
        style.ParaKeepTogether = True
        style.ParaSplit = False
    if italic:
        style.CharPosture = 2
    if bold:
        style.CharWeight = 150
    if outline:
        style.OutlineLevel = outline
    if align is not None:
        style.ParaAdjust = align


def main():
    OUT.mkdir(parents=True, exist_ok=True)
    local_ctx = uno.getComponentContext()
    resolver = local_ctx.ServiceManager.createInstanceWithContext(
        "com.sun.star.bridge.UnoUrlResolver", local_ctx)
    ctx = resolver.resolve(
        "uno:socket,host=localhost,port=2002;urp;StarOffice.ComponentContext")
    desktop = ctx.ServiceManager.createInstanceWithContext(
        "com.sun.star.frame.Desktop", ctx)
    doc = desktop.loadComponentFromURL("private:factory/swriter", "_blank", 0, ())

    page = doc.StyleFamilies.getByName("PageStyles").getByName("Default Page Style")
    page.Width, page.Height = 14800, 21000
    page.LeftMargin, page.RightMargin = 1800, 1500
    page.TopMargin, page.BottomMargin = 1800, 1800
    page.HeaderIsOn, page.FooterIsOn = True, True
    page.HeaderHeight, page.FooterHeight = 450, 550
    footer_cursor = page.FooterText.createTextCursor()
    footer_cursor.ParaAdjust = 3
    field = doc.createInstance("com.sun.star.text.TextField.PageNumber")
    field.NumberingType = 4
    page.FooterText.insertTextContent(footer_cursor, field, False)

    styles = doc.StyleFamilies.getByName("ParagraphStyles")
    set_style(styles, "Default Paragraph Style", font="Liberation Serif", size=10.5, after=160)
    set_style(styles, "Title", font="Liberation Serif", size=24, before=3200, after=500, bold=True, align=3)
    set_style(styles, "Subtitle", font="Liberation Serif", size=13, after=500, italic=True, align=3)
    set_style(styles, "Heading 1", font="Liberation Serif", size=16, after=380, page_before=True, keep=True, bold=True, outline=1)
    set_style(styles, "Heading 2", font="Liberation Serif", size=12, before=300, after=220, keep=True, bold=True, outline=2)

    body_started = False
    for raw in SOURCE.read_text(encoding="utf-8").splitlines():
        line = raw.strip()
        if not line or line.startswith("---"):
            continue
        if line.startswith("# "):
            add_paragraph(doc, line[2:], "Title")
        elif line.startswith("## "):
            add_paragraph(doc, line[3:], "Heading 1")
            body_started = True
        elif line.startswith("### "):
            add_paragraph(doc, line[4:], "Heading 2")
        elif not body_started and (line == "Авторское издание" or line.startswith("Москва")):
            add_paragraph(doc, line, "Subtitle")
        else:
            add_paragraph(doc, line, "Default Paragraph Style")

    props = doc.DocumentProperties
    props.Title = "Гладиаторы успеха"
    props.Author = "Николай Копылов"
    props.Subject = "Манифест против фальшивого успеха"
    doc.storeAsURL(file_url(OUT / "gladiatory-uspekha-interior.odt"), (prop("FilterName", "writer8"),))
    doc.storeToURL(file_url(OUT / "gladiatory-uspekha-interior.docx"), (prop("FilterName", "Office Open XML Text"),))
    doc.storeToURL(file_url(OUT / "gladiatory-uspekha-interior.pdf"), (prop("FilterName", "writer_pdf_Export"),))
    doc.close(True)


if __name__ == "__main__":
    main()
