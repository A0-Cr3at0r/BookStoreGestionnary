from pyscript import web, when, window

from excel_manager import process_excel
from save_manager import Save, load

# =========================================================
# APPLICATION STATE
# =========================================================

bookstore = None
reports = []


# =========================================================
# BOOKSTORE
# =========================================================

@when("click", "#bookstore-drop-zone")
def select_bookstore(event):
    """
    Ouvre le sélecteur de fichiers pour le bookstore.
    """

    web.page["bookstore-file"].click()


@when("change", "#bookstore-file")
def bookstore_selected(event):
    """
    Récupère le fichier sélectionné pour le bookstore.
    """

    files = event.target.files

    if len(files) == 0:
        return

    set_bookstore(files[0])


def set_bookstore(file):
    """
    Définit le fichier bookstore actuellement sélectionné.
    """

    global bookstore

    bookstore = file

    display_bookstore()


def display_bookstore():
    """
    Met à jour l'affichage du fichier bookstore.
    """

    container = web.page["bookstore-selected-file"]

    container.innerHTML = ""

    if bookstore is None:
        return

    file_element = web.div(
        web.span(bookstore.name),
        web.button(
            "×",
            classes=["remove-file"]
        ),
        classes=["selected-file-item"]
    )

    file_element.children[1].onclick = remove_bookstore

    container.append(file_element)


def remove_bookstore(event):
    """
    Retire le fichier bookstore actuellement sélectionné.
    """

    global bookstore

    bookstore = None

    web.page["bookstore-file"].value = ""

    display_bookstore()


# =========================================================
# REPORTS
# =========================================================

@when("click", "#reports-drop-zone")
def select_reports(event):
    """
    Ouvre le sélecteur de fichiers pour les reports.
    """

    web.page["report-file"].click()


@when("change", "#report-file")
def reports_selected(event):
    """
    Récupère les fichiers sélectionnés pour les reports.
    """

    files = event.target.files

    for file in files:
        add_report(file)

    web.page["report-file"].value = ""


def add_report(file):
    """
    Ajoute un fichier à la liste des reports.
    """

    reports.append(file)

    display_reports()


def remove_report(index):
    """
    Supprime un report de la liste.
    """

    del reports[index]

    display_reports()


def display_reports():
    """
    Reconstruit l'affichage de la liste des reports.
    """

    container = web.page["reports-list"]

    container.innerHTML = ""

    for index, report in enumerate(reports):

        remove_button = web.button(
            "×",
            classes=["remove-file"]
        )

        remove_button.onclick = (
            lambda event, i=index: remove_report(i)
        )

        report_element = web.div(
            web.span(report.name),
            remove_button,
            classes=["report-item"]
        )

        container.append(report_element)


# =========================================================
# UPDATE
# =========================================================

@when("click", "#update-button")
def update():
    """
    Lance le traitement Excel et affiche le fichier résultant.
    """

    result = process_excel(
        bookstore,
        reports
    )

    display_output(result)


def display_output(file):
    """
    Affiche le fichier Excel résultant avec un bouton
    permettant de le télécharger.
    """

    container = web.page["output-section"]

    container.innerHTML = ""

    if file is None:
        container.append(
            web.span("Aucun fichier de librairie fourni")
        )
        return

    download_url = window.URL.createObjectURL(file)

    file_element = web.div(
        web.span(file.name),
        web.a(
            "Télécharger",
            href=download_url,
            download=file.name,
            classes=["download-file"]
        ),
        classes=["output-file"]
    )

    container.append(file_element)

# =========================================================
# DRAG & DROP
# =========================================================

@when("dragover", "#bookstore-drop-zone")
def allow_bookstore_drop(event):
    """
    Autorise le dépôt dans la zone bookstore.
    """

    event.preventDefault()


@when("drop", "#bookstore-drop-zone")
def bookstore_drop(event):
    """
    Récupère le fichier déposé dans la zone bookstore.
    """

    event.preventDefault()

    files = event.dataTransfer.files

    if len(files) == 0:
        return

    # Le bookstore n'accepte qu'un seul fichier.
    set_bookstore(files[0])


@when("dragover", "#reports-drop-zone")
def allow_reports_drop(event):
    """
    Autorise le dépôt dans la zone reports.
    """

    event.preventDefault()


@when("drop", "#reports-drop-zone")
def reports_drop(event):
    """
    Récupère les fichiers déposés dans la zone reports.
    """

    event.preventDefault()

    files = event.dataTransfer.files

    for file in files:
        add_report(file)
