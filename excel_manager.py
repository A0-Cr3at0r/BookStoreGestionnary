import warnings

warnings.filterwarnings("ignore")

from io import BytesIO

from openpyxl import load_workbook

from pyscript import ffi, window


# =========================================================
# CONFIGURATION
# =========================================================

BOOKSTORE_HEADER_ROW = 2

BOOKSTORE_ARTICLE_HEADER = "Nom de l’article"

REPORT_HEADER_ROW = 1
REPORT_LIBRARY_CELL = "A2"

REPORT_ARTICLE_HEADER = "Articles"
REPORT_STOCK_HEADER = "StockRestant"


# =========================================================
# LECTURE DES FICHIERS
# =========================================================

async def _read_file(file):
    """
    Lit le contenu binaire d'un objet File JavaScript.

    Retourne :
        bytes
    """

    array_buffer = await file.arrayBuffer()

    return array_buffer.to_bytes()


def _load_workbook_from_bytes(data):
    """
    Charge un classeur Excel depuis des bytes.
    """

    return load_workbook(
        filename=BytesIO(data)
    )


# =========================================================
# RECHERCHE DES COLONNES
# =========================================================

def _find_column(ws, header, header_row):
    """
    Recherche une colonne dont l'en-tête correspond exactement
    à la valeur fournie.

    Retourne l'indice de colonne ou None.
    """

    for cell in ws[header_row]:

        if cell.value == header:
            return cell.column

    return None


def _find_all_columns(ws, header, header_row):
    """
    Recherche toutes les colonnes dont l'en-tête correspond
    exactement à la valeur fournie.

    Retourne une liste d'indices de colonnes.
    """

    columns = []

    for cell in ws[header_row]:

        if cell.value == header:
            columns.append(cell.column)

    return columns


def _find_quantity_columns(ws, library_name):
    """
    Recherche toutes les colonnes du bookstore correspondant
    à la librairie fournie par le rapport.

    Exemple :

        library_name = "EJP"

    recherche :

        Quantité actuelle *EJP*

    La ligne entière est parcourue. Les colonnes n'ont donc
    pas besoin d'être contiguës.
    """

    columns = []

    for cell in ws[BOOKSTORE_HEADER_ROW]:
        header = cell.value

        if not isinstance(header, str):
            continue

        if (
            "Quantité actuelle" in header
            and library_name in header
        ):
            columns.append(cell.column)

    return columns


# =========================================================
# INDEX DES ARTICLES
# =========================================================

def _build_article_index(ws, article_column):
    """
    Construit un index permettant de retrouver rapidement
    la ligne correspondant à un article.

    La correspondance est strictement exacte.

    Si plusieurs lignes du bookstore contiennent exactement
    le même article, la dernière ligne rencontrée est conservée.
    """
    article_index = {}

    for row in range(
        BOOKSTORE_HEADER_ROW + 1,
        ws.max_row + 1
    ):
        article = ws.cell(
            row=row,
            column=article_column
        ).value

        if article is None:
            continue

        if article not in article_index:
            article_index[article] = []

        article_index[article].append(row)

    return article_index

# =========================================================
# TRAITEMENT D'UN RAPPORT
# =========================================================

def _process_report(bookstore_ws, report_ws):
    """
    Applique un rapport au bookstore.

    Le rapport est traité uniquement si une ou plusieurs
    colonnes de quantité correspondent à la librairie
    indiquée dans A2.

    Retourne le nombre de mises à jour effectuées.
    """

    # -----------------------------------------------------
    # LIBRAIRIE
    # -----------------------------------------------------

    library_name = report_ws[
        REPORT_LIBRARY_CELL
    ].value

    if library_name is None:
        return 0

    # -----------------------------------------------------
    # COLONNES CORRESPONDANT À LA LIBRAIRIE
    # -----------------------------------------------------

    quantity_columns = _find_quantity_columns(
        bookstore_ws,
        library_name
    )

    # La librairie n'existe pas dans le bookstore.
    #
    # Dans ce cas, on ne poursuit PAS la lecture du rapport.
    if not quantity_columns:
        return 0

    # -----------------------------------------------------
    # COLONNES DU RAPPORT
    # -----------------------------------------------------

    article_column = _find_column(
        report_ws,
        REPORT_ARTICLE_HEADER,
        REPORT_HEADER_ROW
    )

    stock_column = _find_column(
        report_ws,
        REPORT_STOCK_HEADER,
        REPORT_HEADER_ROW
    )

    if article_column is None or stock_column is None:
        return 0

    # -----------------------------------------------------
    # COLONNE ARTICLE DU BOOKSTORE
    # -----------------------------------------------------

    bookstore_article_column = _find_column(
        bookstore_ws,
        BOOKSTORE_ARTICLE_HEADER,
        BOOKSTORE_HEADER_ROW
    )

    if bookstore_article_column is None:
        return 0

    # -----------------------------------------------------
    # INDEX DES ARTICLES
    # -----------------------------------------------------

    article_index = _build_article_index(
        bookstore_ws,
        bookstore_article_column
    )

    # -----------------------------------------------------
    # LECTURE DES LIGNES DU RAPPORT
    # -----------------------------------------------------

    updated_count = 0

    for row in range(
        REPORT_HEADER_ROW + 1,
        report_ws.max_row + 1
    ):

        article = report_ws.cell(
            row=row,
            column=article_column
        ).value

        stock = report_ws.cell(
            row=row,
            column=stock_column
        ).value

        # -------------------------------------------------
        # STOCK VIDE
        # -------------------------------------------------
        #
        # Une valeur vide signifie que cette entrée ne
        # définit pas de stock.
        #
        # On ne lance donc même pas la recherche de l'article
        # dans le bookstore.
        #

        if stock is None or stock == "":
            continue

        # -------------------------------------------------
        # ARTICLE ABSENT DU BOOKSTORE
        # -------------------------------------------------
        #
        # Le rapport est la source de vérité, mais le
        # bookstore doit tout de même contenir l'article
        # pour pouvoir être mis à jour.
        #

        if article not in article_index:
            continue

        bookstore_rows = article_index[article]

        # -------------------------------------------------
        # MISE À JOUR DE TOUTES LES COLONNES
        # -------------------------------------------------
        #
        # Si plusieurs colonnes portent exactement le nom
        # correspondant à la librairie, elles sont toutes
        # mises à jour.
        #

        for bookstore_row in bookstore_rows:
            for quantity_column in quantity_columns:
                bookstore_ws.cell(
                    row=bookstore_row,
                    column=quantity_column,
                    value=stock
                )

        updated_count += 1

    return updated_count


# =========================================================
# CRÉATION DU FICHIER DE SORTIE
# =========================================================

def _create_output_file(workbook, bookstore):
    """
    Transforme le classeur openpyxl en objet File JavaScript.
    """

    output_stream = BytesIO()

    workbook.save(output_stream)

    data = output_stream.getvalue()

    # -----------------------------------------------------
    # NOM DU FICHIER
    # -----------------------------------------------------

    original_name = bookstore.name

    if original_name.lower().endswith(".xlsx"):
        output_name = (
            original_name[:-5]
            + ".xlsx"
        )
    else:
        output_name = (
            original_name
            + ".xlsx"
        )

    # -----------------------------------------------------
    # TRANSFORMATION PYTHON → JAVASCRIPT
    # -----------------------------------------------------

    buffer = window.Uint8Array.new(len(data))

    buffer.assign(data)

    file_options = ffi.to_js({
        "type": (
            "application/"
            "vnd.openxmlformats-officedocument."
            "spreadsheetml.sheet"
        )
    })

    return window.File.new(
        [buffer],
        output_name,
        file_options
    )


# =========================================================
# TRAITEMENT PRINCIPAL
# =========================================================

async def process_excel(bookstore, reports):
    """
    Met à jour le bookstore à partir des rapports.

    Paramètres
    ----------
    bookstore :
        Objet File JavaScript correspondant au fichier
        Excel du bookstore.

    reports :
        Liste d'objets File JavaScript correspondant aux
        rapports Excel des librairies.

    Retour
    ------
    File JavaScript
        Nouveau fichier Excel généré.

    Le fichier bookstore original n'est jamais modifié.
    """

    # -----------------------------------------------------
    # VALIDATION
    # -----------------------------------------------------

    if bookstore is None:
        return None

    # -----------------------------------------------------
    # LECTURE DU BOOKSTORE
    # -----------------------------------------------------

    bookstore_data = await _read_file(bookstore)

    workbook = _load_workbook_from_bytes(
        bookstore_data
    )

    bookstore_ws = workbook.active

    # -----------------------------------------------------
    # VÉRIFICATION DU BOOKSTORE
    # -----------------------------------------------------

    bookstore_article_column = _find_column(
        bookstore_ws,
        BOOKSTORE_ARTICLE_HEADER,
        BOOKSTORE_HEADER_ROW
    )

    if bookstore_article_column is None:

        raise ValueError(
            f"La colonne "
            f"'{BOOKSTORE_ARTICLE_HEADER}' "
            f"est introuvable dans le bookstore."
        )

    # -----------------------------------------------------
    # TRAITEMENT DES RAPPORTS
    # -----------------------------------------------------

    total_updated = 0

    for report in reports:

        report_data = await _read_file(report)

        report_workbook = _load_workbook_from_bytes(
            report_data
        )
        report_ws = report_workbook.active

        updated = _process_report(
            bookstore_ws,
            report_ws
        )

        total_updated += updated

    # -----------------------------------------------------
    # CRÉATION DU NOUVEAU FICHIER
    # -----------------------------------------------------

    return _create_output_file(
        workbook,
        bookstore
    )
