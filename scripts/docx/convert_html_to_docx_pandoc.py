import subprocess
import sys
from pathlib import Path


def convert_html_to_docx(input_html_path, output_docx_path):
    input_html = Path(input_html_path)
    output_docx = Path(output_docx_path)

    if not input_html.is_file():
        print(f"Fichier HTML introuvable : {input_html}")
        sys.exit(1)

    try:
        # Appel à Pandoc : HTML -> DOCX
        subprocess.run(
            [
                "pandoc",
                str(input_html),
                "-f", "html",
                "-t", "docx",
                "-o", str(output_docx),
            ],
            check=True,
        )
    except FileNotFoundError:
        print("Erreur : 'pandoc' n'est pas trouvé. Vérifie qu'il est bien installé et dans le PATH.")
        sys.exit(1)
    except subprocess.CalledProcessError as e:
        print(f"Pandoc a renvoyé une erreur (code {e.returncode}).")
        sys.exit(1)

    print(f"Conversion terminée : {output_docx}")


if __name__ == "__main__":
    if len(sys.argv) != 3:
        print("Usage :")
        print("  python convert_html_to_docx_pandoc.py <entrée.html> <sortie.docx>")
        print("Exemple :")
        print("  python convert_html_to_docx_pandoc.py ..\\charte.html ..\\charte.docx")
        sys.exit(1)

    convert_html_to_docx(sys.argv[1], sys.argv[2])
