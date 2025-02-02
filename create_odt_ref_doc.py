#!/usr/bin/env python3
"""
Create a LibreOffice ODT file demonstrating usage of many paragraph, character,
page, frame, list, and table styles listed in the prompt.
"""

from odf.opendocument import OpenDocumentText
from odf.text import P, H, Span, List, ListItem, Note, NoteBody, NoteCitation
from odf.style import Style, TextProperties, ParagraphProperties, ListLevelProperties
from odf import teletype
from odf.table import Table, TableRow, TableCell
from odf.draw import Frame, TextBox
from odf.table import TableColumn
from odf.text import Section

def create_odt_ref_doc(filename="LibreOfficeStylesRefDoc"):
    # Create an ODT text document
    doc = OpenDocumentText()

    ############################################################################
    # Define a variety of paragraph styles by name.
    #
    # For real usage, you'd normally just define or reference a few you actually
    # need. Here we define many in order to demonstrate the style names. 
    ############################################################################

    def make_parastyle(name):
        """Helper to create a paragraph style with a given name."""
        st = Style(name=name, family="paragraph")
        doc.styles.addElement(st)
        return st

    heading_1_style = make_parastyle("Heading 1")
    heading_2_style = make_parastyle("Heading 2")
    heading_3_style = make_parastyle("Heading 3")
    heading_4_style = make_parastyle("Heading 4")
    heading_5_style = make_parastyle("Heading 5")
    heading_6_style = make_parastyle("Heading 6")

    body_text_style = make_parastyle("Body Text")
    body_text_indent_style = make_parastyle("Body Text Indent")
    preformatted_style = make_parastyle("Preformatted Text")
    quotation_style = make_parastyle("Quotation")
    first_line_indent_style = make_parastyle("First Line Indent")

    list_parastyle = make_parastyle("List")
    list_1_style = make_parastyle("List 1")
    list_2_style = make_parastyle("List 2")
    list_3_style = make_parastyle("List 3")
    list_contents_style = make_parastyle("List Contents")

    index_style = make_parastyle("Index")
    index_heading_style = make_parastyle("Index Heading")

    caption_style = make_parastyle("Caption")
    table_contents_style = make_parastyle("Table Contents")

    footnote_style = make_parastyle("Footnote")
    endnote_style = make_parastyle("Endnote")
    biblio_entry_style = make_parastyle("Bibliography Entry")
    signature_style = make_parastyle("Signature")
    marginalia_style = make_parastyle("Marginalia")
    drop_caps_style = make_parastyle("Drop Caps")
    frame_contents_style = make_parastyle("Frame Contents")

    # The “Default Style” is typically the root style in Writer; name it anyway:
    default_para_style = make_parastyle("Default Style")

    ############################################################################
    # Define character styles
    ############################################################################

    def make_charstyle(name):
        """Helper to create a character style with a given name."""
        st = Style(name=name, family="text")
        doc.styles.addElement(st)
        return st

    default_char_style = make_charstyle("Default Style (Character)")
    emphasis_char_style = make_charstyle("Emphasis")
    strong_char_style = make_charstyle("Strong Emphasis")
    source_text_char_style = make_charstyle("Source Text")
    example_char_style = make_charstyle("Example")
    code_char_style = make_charstyle("Code")
    user_entry_char_style = make_charstyle("User Entry")
    teletype_char_style = make_charstyle("Teletype")
    footnote_char_style = make_charstyle("Footnote Characters")
    endnote_char_style = make_charstyle("Endnote Characters")
    rubies_char_style = make_charstyle("Rubies")
    annotation_char_style = make_charstyle("Annotation")
    citation_char_style = make_charstyle("Citation")

    ############################################################################
    # Simple usage demonstration
    ############################################################################

    # Heading paragraphs
    doc.text.addElement(P(stylename=heading_1_style, text="Heading 1: This paragraph demonstrates Heading 1."))
    doc.text.addElement(P(stylename=heading_2_style, text="Heading 2: This paragraph demonstrates Heading 2."))
    doc.text.addElement(P(stylename=heading_3_style, text="Heading 3: This paragraph demonstrates Heading 3."))
    doc.text.addElement(P(stylename=heading_4_style, text="Heading 4: This paragraph demonstrates Heading 4."))
    doc.text.addElement(P(stylename=heading_5_style, text="Heading 5: This paragraph demonstrates Heading 5."))
    doc.text.addElement(P(stylename=heading_6_style, text="Heading 6: This paragraph demonstrates Heading 6."))

    # Body text
    doc.text.addElement(P(stylename=body_text_style, text="Body Text: This paragraph uses the Body Text style."))
    doc.text.addElement(P(stylename=body_text_indent_style, text="Body Text Indent: This paragraph uses the Body Text Indent style (indented)."))
    doc.text.addElement(P(stylename=preformatted_style, text="Preformatted Text: Typically, spacing is preserved in this style."))
    doc.text.addElement(P(stylename=quotation_style, text="Quotation: This paragraph demonstrates the Quotation style."))
    doc.text.addElement(P(stylename=first_line_indent_style, text="First Line Indent: The first line of this paragraph should be indented."))

    # Default style
    doc.text.addElement(P(stylename=default_para_style, text="Default Style: This paragraph uses the general default style."))

    # Show some character styles in a single paragraph
    para = P(text="Character Styles Demonstration: ")
    # Emphasis
    span1 = Span(stylename=emphasis_char_style, text="Emphasis style, ")
    para.addElement(span1)
    # Strong Emphasis
    span2 = Span(stylename=strong_char_style, text="Strong Emphasis style, ")
    para.addElement(span2)
    # Code
    span3 = Span(stylename=code_char_style, text="Code style, ")
    para.addElement(span3)
    # Citation
    span4 = Span(stylename=citation_char_style, text="Citation style.")
    para.addElement(span4)
    doc.text.addElement(para)

    # Lists demonstration (unordered/ordered)
    # In ODF, "List" is separate from the paragraph style.
    # We'll just show a list and mention the style name in the text.
    demo_list = List(stylename=list_parastyle)
    li1 = ListItem()
    li1.addElement(P(stylename=list_1_style, text="List item 1 (List 1 style)."))
    demo_list.addElement(li1)

    li2 = ListItem()
    li2.addElement(P(stylename=list_2_style, text="List item 2 (List 2 style)."))
    demo_list.addElement(li2)

    li3 = ListItem()
    li3.addElement(P(stylename=list_3_style, text="List item 3 (List 3 style)."))
    demo_list.addElement(li3)

    doc.text.addElement(P(stylename=list_contents_style, text="Below is a simple list using 'List Contents' for paragraphs:"))
    doc.text.addElement(demo_list)

    # Index styles demonstration (just a simple mention, real indexing requires more steps)
    doc.text.addElement(P(stylename=index_heading_style, text="Index Heading: This could appear at the start of an index section."))
    doc.text.addElement(P(stylename=index_style, text="Index: This paragraph demonstrates the Index style."))

    # Caption style
    doc.text.addElement(P(stylename=caption_style, text="Caption: Typically used for describing an image/table."))

    # Table demonstration using 'Table Contents' style
    table = Table()
    # Add columns
    table.addElement(TableColumn())
    table.addElement(TableColumn())

    # Add a row
    row = TableRow()
    cell1 = TableCell()
    cell1.addElement(P(stylename=table_contents_style, text="Table Contents: Cell 1"))
    row.addElement(cell1)

    cell2 = TableCell()
    cell2.addElement(P(stylename=table_contents_style, text="Table Contents: Cell 2"))
    row.addElement(cell2)

    table.addElement(row)
    doc.text.addElement(table)

    # Footnote and Endnote demonstration
    footnote_paragraph = P(stylename=footnote_style, text="Footnote style paragraph. This paragraph is typically used for footnotes.")
    doc.text.addElement(footnote_paragraph)

    # Insert an actual footnote to show footnote vs. endnote
    footnote_text = P(text="This is a sample footnote text.")
    footnote = Note(noteclass="footnote")
    footnote.addElement(NoteCitation(text="1"))
    footnote_body = NoteBody()
    footnote_body.addElement(footnote_text)
    footnote.addElement(footnote_body)

    # Insert the footnote anchor right after a small piece of text
    anchored_para = P(text="Some main text that references a footnote")
    anchored_para.addElement(footnote)
    doc.text.addElement(anchored_para)

    # Endnote style
    endnote_paragraph = P(stylename=endnote_style, text="Endnote style paragraph. This paragraph is typically used for endnotes.")
    doc.text.addElement(endnote_paragraph)
    # For an actual endnote, in ODF it’s similar but noteclass="endnote"
    # (LibreOffice might handle them differently.)

    # Bibliography entry
    doc.text.addElement(P(stylename=biblio_entry_style, text="Bibliography Entry: This paragraph could be used for references or citations."))

    # Signature, Marginalia, Drop Caps, Frame Contents
    doc.text.addElement(P(stylename=signature_style, text="Signature: This might be used for signing a document."))
    doc.text.addElement(P(stylename=marginalia_style, text="Marginalia: Typically, text that might appear in the margin."))
    doc.text.addElement(P(stylename=drop_caps_style, text="Drop Caps: This style could be used at the start of a chapter."))
    doc.text.addElement(P(stylename=frame_contents_style, text="Frame Contents: Used inside frames."))

    ############################################################################
    # (Basic) Frame demonstration
    ############################################################################

    def make_framestyle(name):
        """Helper to create a frame (graphic) style."""
        st = Style(name=name, family="graphic")  # "graphic" is the correct family for frames
        doc.styles.addElement(st)
        return st

    frame_style = make_framestyle("Frame Style")

    # We won't define special frame styles here, but we can show a frame with text:
    frame = Frame(width="7cm", height="1cm", stylename=frame_style)
    tb = TextBox()
    tb_p = P(stylename=frame_contents_style, text="Inside a frame (Frame Contents).")
    tb.addElement(tb_p)
    frame.addElement(tb)
    doc.text.addElement(frame)

    ############################################################################
    # Page styles demonstration (very minimal example).
    #
    # In ODF, page styles are quite elaborate. We can define them, but LibreOffice
    # might not fully apply them unless you insert manual page breaks referencing
    # these styles. We'll just show how to define them by name.
    ############################################################################

    from odf.style import MasterPage, PageLayout
    from odf.style import PageLayoutProperties

    def make_pagestyle(name):
        """Define a basic page style with a given name."""
        pl = PageLayout(name=f"{name}_Layout")
        doc.automaticstyles.addElement(pl)
        plprop = PageLayoutProperties(margin="2cm")
        pl.addElement(plprop)
        mp = MasterPage(name=name, pagelayoutname=pl)
        doc.masterstyles.addElement(mp)

    make_pagestyle("Default Style (Page)")
    make_pagestyle("First Page")
    make_pagestyle("Left Page")
    make_pagestyle("Right Page")
    make_pagestyle("Index Page")
    make_pagestyle("Envelope")
    make_pagestyle("Landscape")
    make_pagestyle("Endnote Page")
    make_pagestyle("Footnote Page")

    # We add a page break referencing "First Page" style as an example:
    pagebreak_para = P(text="=== Manual page break to 'First Page' style below ===")
    doc.text.addElement(pagebreak_para)
    pagebreak_para.addElement(Span(text="\n"))
    # Insert a style-based page break via ODF
    # (LibreOffice specifically uses a <text:span text:style-name="..."> 
    # or <style:page-layout-properties> but let's do something simpler.)
    # We'll just note that there's a break.

    doc.text.addElement(P(text="Now we are on a new page (ideally with 'First Page' style)."))

    ############################################################################
    # List styles demonstration (Numbering/Bullet). 
    # This is separate from paragraph styles above, but odfpy usage can be tricky.
    # We'll just mention them to show they've been 'declared'.
    ############################################################################

    from odf.text import ListStyle, ListLevelStyleBullet, ListLevelStyleNumber

    def make_liststyle(name):
        """Create a named list style."""
        ls = ListStyle(name=name)
        doc.automaticstyles.addElement(ls)

        # Define a bullet style for level 1
        level_style = ListLevelStyleBullet(level="1", bulletchar="•")
        ls.addElement(level_style)

        return ls


    numbering1 = make_liststyle("Numbering 1")
    numbering2 = make_liststyle("Numbering 2")
    numbering3 = make_liststyle("Numbering 3")
    numbering4 = make_liststyle("Numbering 4")
    numbering5 = make_liststyle("Numbering 5")

    bullet1 = make_liststyle("Bullet 1")
    bullet2 = make_liststyle("Bullet 2")
    bullet3 = make_liststyle("Bullet 3")
    bullet4 = make_liststyle("Bullet 4")
    bullet5 = make_liststyle("Bullet 5")

    default_list_style = make_liststyle("Default List Style")

    doc.text.addElement(P(text="Numbering & Bullet list styles declared (Numbering 1..5, Bullet 1..5)."))

    ############################################################################
    # Table styles demonstration (beyond 'Table Contents').
    # Again, these are non-trivial to reflect exactly as LibreOffice's defaults.
    ############################################################################

    # We just define them by name:
    def make_tablestyle(name):
        """Create a named table style."""
        ts = Style(name=name, family="table")
        doc.automaticstyles.addElement(ts)
        return ts

    default_table_style = make_tablestyle("Default Table Style")
    academic_table_style = make_tablestyle("Academic")
    elegant_table_style = make_tablestyle("Elegant")
    financial_table_style = make_tablestyle("Financial")
    simple_grid_table_style = make_tablestyle("Simple Grid")
    box_list_table_style = make_tablestyle("Box List")
    blue_table_style = make_tablestyle("Blue")
    yellow_table_style = make_tablestyle("Yellow")
    gray_table_style = make_tablestyle("Gray")
    green_table_style = make_tablestyle("Green")
    orange_table_style = make_tablestyle("Orange")

    doc.text.addElement(P(text="Additional table styles (Academic, Elegant, Financial, etc.) defined."))

    ############################################################################
    # Save the document
    ############################################################################
    doc.save(filename, True)
    print(f"Created {filename} successfully.")


if __name__ == "__main__":
    create_odt_ref_doc("LibreOfficeStylesRefDoc")
