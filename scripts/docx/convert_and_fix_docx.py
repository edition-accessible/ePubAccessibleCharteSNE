import subprocess
import sys
from docx import Document
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from pathlib import Path
import os


# ---------------------------
# AJOUT DES BORDURES AUX TABLEAUX
# ---------------------------

def set_cell_border(cell, **kwargs):
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()

    # remove existing borders
    for child in tcPr.findall(qn('w:tcBorders')):
        tcPr.remove(child)

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


# ---------------------------
# CONVERSION HTML → DOCX (Pandoc)
# ---------------------------

def convert_html_to_docx(html_path, raw_docx_path):
    cmd = [
        "pandoc",
        html_path,
        "-o", raw_docx_path,
        "--embed-resources",
        "--standalone"
    ]

    print("🔄 Conversion HTML → DOCX via Pandoc…")
    subprocess.run(cmd, check=True)
    print(f"✔ Fichier DOCX brut créé → {raw_docx_path}")


# ---------------------------
# SCRIPT PRINCIPAL
# ---------------------------

if __name__ == "__main__":
    if len(sys.argv) != 2:
        print("Usage : python convert_and_fix_docx.py <chemin_complet_du_HTML>")
        sys.exit(1)

    html_path = Path(sys.argv[1])
    if not html_path.exists():
        print(f"ERREUR : fichier HTML introuvable : {html_path}")
        sys.exit(1)

    # Nom final = même nom que HTML mais .docx
    final_docx = html_path.with_suffix(".docx")

    # Temporaire
    temp_docx = html_path.with_name(html_path.stem + "_temp.docx")

    # Étape 1 : Conversion
    convert_html_to_docx(str(html_path), str(temp_docx))

    # Étape 2 : Bordures
    add_borders_to_docx(str(temp_docx), str(final_docx))

    # Supprimer fichier temporaire
    try:
        os.remove(temp_docx)
        print(f"🗑 Fichier temporaire supprimé : {temp_docx}")
    except Exception as e:
        print(f"⚠ Impossible de supprimer {temp_docx} : {e}")

    print("\n🎉 Conversion terminée !")
    print(f"➡️ Fichier final généré : {final_docx}")
