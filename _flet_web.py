"""UI Flet en mode web headless pour test visuel + screenshots. Jetable.
État piloté par la variable d'env UISTATE : empty | table | dark | dialog."""
import os
import flet as ft
import interface_flet as gui

STATE = os.environ.get("UISTATE", "empty")
SEED = [
    {"nom": "Martin", "prenom": "Alice", "identifiant": "A101",
     "type_bien": "Appartement", "adresse": "12 rue des Lilas, 75011 Paris",
     "loyer_nu": 800, "charges": 150, "part_caf": 200, "depot_garantie": 1600,
     "date_entree": "2024-01-15", "date_sortie": ""},
    {"nom": "Bernard", "prenom": "Sophie", "identifiant": "M2",
     "type_bien": "Maison", "adresse": "8 avenue du Général Leclerc, 92100 Boulogne",
     "loyer_nu": 1200, "charges": 200, "part_caf": 0, "depot_garantie": 2400,
     "date_entree": "2023-06-01", "date_sortie": "2024-12-31",
     "caution_rendue": True, "observation": "Fin de bail"},
    {"nom": "Nguyen", "prenom": "Thomas", "identifiant": "B3",
     "type_bien": "Studio", "adresse": "45 boulevard Voltaire, 75011 Paris",
     "loyer_nu": 650, "charges": 90, "part_caf": 350, "depot_garantie": 650,
     "date_entree": "2025-03-01", "date_sortie": ""},
]


def main(page):
    app = gui.AppLoyers(page)
    if STATE in ("table", "dark", "dialog", "tri", "search"):
        app.locataires = list(SEED)
        app._rafraichir_table()
    if STATE == "tri":
        app._tri_cle, app._tri_asc = "loyer", False  # loyer décroissant
        app._rafraichir_table()
    if STATE == "search":
        app._recherche = "paris"
        app._rafraichir_table()
    if STATE == "dark":
        app.apparence = "sombre"
        app._appliquer_apparence()
    if STATE == "dialog":
        import threading
        threading.Timer(1.2, lambda: gui.DialogueLocataire(
            page, app._modules_courants(), app._adresses_connues(), None,
            app._on_loc_valide(None)).ouvrir()).start()
    if STATE == "reglages":
        import threading
        threading.Timer(3.5, lambda: app.page.open(app._dlg_reglages)).start()


ft.app(main, view=None, host="0.0.0.0", port=8550)
