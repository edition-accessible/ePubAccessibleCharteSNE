import sys
from pathlib import Path
from collections import OrderedDict
from bs4 import BeautifulSoup
import html


def extract_referentiels_from_table(table):
    """
    À partir d'un <table class="table-danger">,
    construit une structure :
    OrderedDict[
        nom_référentiel -> [
            {"href": ..., "text": ..., "level": ...},
            ...
        ]
    ]
    """
    refs = OrderedDict()

    # Récupère les lignes du tableau (tbody si présent, sinon tout)
    tbody = table.find("tbody")
    if tbody:
        rows = tbody.find_all("tr", recursive=False)
    else:
        rows = table.find_all("tr", recursive=False)

    current_ref = None

    for tr in rows:
        # Première cellule d'en-tête de ligne = nom du référentiel
        th = tr.find("th", scope="row")
        if th:
            # texte "nu" (ignore les <mark>, etc.)
            current_ref = th.get_text(strip=True)

        if not current_ref:
            # Ligne mal formée, on saute
            continue

        tds = tr.find_all("td")
        if len(tds) < 2:
            continue

        section_td = tds[0]
        level_td = tds[1]

        # Lien dans la cellule "Section / paragraphe"
        link = section_td.find("a")
        if link:
            href = link.get("href", "").strip()
            link_text = " ".join(link.stripped_strings)
        else:
            href = ""
            link_text = section_td.get_text(strip=True)

        level_text = level_td.get_text(strip=True)

        entry = {
            "href": href,
            "text": link_text,
            "level": level_text,
        }

        if current_ref not in refs:
            refs[current_ref] = []
        refs[current_ref].append(entry)

    return refs, len(rows)


def build_list_block_html(refs, row_count):
    """
    Construit le bloc HTML <h3> + <ul> à partir de la structure de données.
    - "Référentiel" si une seule ligne
    - "Référentiels" si plusieurs lignes
    - <br> entre plusieurs entrées pour le même référentiel
    """
    heading = "Référentiel" if row_count == 1 else "Référentiels"

    parts = []
    parts.append(f'<h3 class="h4 p-1 ps-2 mb-0 ref-titre-danger">{heading}</h3>')
    parts.append('<ul class="list-unstyled ref-danger">')

    for ref_name, entries in refs.items():
        # ouverture du <li>
        li_parts = []
        li_parts.append('<li class="mb-0">▸ ')
        li_parts.append(f"<strong>{html.escape(ref_name)}</strong> <br> ")

        # plusieurs lignes possibles pour un même réf (rowspan)
        first = True
        for e in entries:
            if not first:
                li_parts.append("<br>")
            first = False

            if e["href"]:
                link_html = (
                    f'<a href="{html.escape(e["href"], quote=True)}" '
                    f'class="theme-link" target="_blank">'
                    f'{html.escape(e["text"])}'
                    f"</a>"
                )
            else:
                link_html = html.escape(e["text"])

            level_html = html.escape(e["level"])

            li_parts.append(
                f"{link_html}"
                f' · <small class="text-muted">Niveau d’accessibilité : {level_html}</small>'
            )

        li_parts.append("</li>")
        parts.append("".join(li_parts))

    parts.append("</ul>")

    return "\n".join(parts)


def process_html(input_path, output_path):
    html_text = Path(input_path).read_text(encoding="utf-8")
    soup = BeautifulSoup(html_text, "html.parser")

    # Trouve tous les tableaux avec la classe "table-danger"
    tables = soup.find_all("table", class_="table-danger")

    print(f"{len(tables)} tableau(x) 'table-danger' trouvé(s).")

    for table in tables:
        refs, row_count = extract_referentiels_from_table(table)
        if not refs:
            continue

        block_html = build_list_block_html(refs, row_count)
        block_soup = BeautifulSoup(block_html, "html.parser")

        # On insère le bloc juste après le conteneur du tableau,
        # s'il est dans un <div class="table-responsive ...">, sinon juste après le <table>
        parent = table.parent
        insert_after_target = table

        if (
            parent
            and parent.name == "div"
            and "table-responsive" in (parent.get("class") or [])
        ):
            insert_after_target = parent

        insert_after_target.insert_after(block_soup)

    # Écrit le résultat
    Path(output_path).write_text(str(soup), encoding="utf-8")
    print(f"Fichier écrit : {output_path}")


if __name__ == "__main__":
    if len(sys.argv) != 3:
        print("Usage : python add_referentiel_lists.py input.html output.html")
        sys.exit(1)

    process_html(sys.argv[1], sys.argv[2])
