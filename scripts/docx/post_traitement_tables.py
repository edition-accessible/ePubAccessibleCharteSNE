from docx import Document
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

def set_cell_border(cell, **kwargs):
    """
    kwargs example:
    set_cell_border(cell,
        top={"sz": 8, "val": "single", "color": "000000"},
        bottom={"sz": 8, "val": "single", "color": "000000"},
    )
    """

    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()

    # remove existing borders
    for child in tcPr.findall(qn('w:tcBorders')):
        tcPr.remove(child)

    # create new borders element
    tcBorders = OxmlElement('w:tcBorders')

    for edge in ("top", "left", "bottom", "right"):
        if edge in kwargs:
            edge_data = kwargs.get(edge)
            element = OxmlElement("w:" + edge)
            element.set(qn("w:val"), edge_data.get("val", "single"))
            element.set(qn("w:sz"), str(edge_data.get("sz", 8)))
            element.set(qn("w:color"), edge_data.get("color", "000000"))
            tcBorders.append(element)

    tcPr.append(tcBorders)


def add_borders_to_docx(input_docx, output_docx):
    doc = Document(input_docx)

    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                set_cell_border(
                    cell,
                    top={"val": "single", "sz": 8, "color": "000000"},
                    bottom={"val": "single", "sz": 8, "color": "000000"},
                    left={"val": "single", "sz": 8, "color": "000000"},
                    right={"val": "single", "sz": 8, "color": "000000"},
                )

    doc.save(output_docx)
    print(f"✔ Bordures appliquées → {output_docx}")


if __name__ == "__main__":
    import sys
    if len(sys.argv) != 3:
        print("Usage : python post_traitement_tables.py <entrée.docx> <sortie.docx>")
        sys.exit(1)

    add_borders_to_docx(sys.argv[1], sys.argv[2])
