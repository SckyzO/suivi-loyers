"""Classeur ouvert dans Excel sous Windows : la génération doit échouer avant tout travail,
avec un message clair, sans créer de copie ni toucher au fichier.

Excel ouvre un classeur en refusant l'écriture aux autres programmes (partage en lecture
seule) : ce verrou est reproduit avec CreateFileW. Ne tourne que sous Windows (job CI
`windows`) ; ailleurs, le smoke test couvre le cas voisin du fichier en lecture seule.
"""

import sys
import tempfile
from pathlib import Path

import generer_suivi_loyers as g


def main() -> int:
    if sys.platform != "win32":
        print("VERROU WINDOWS ignoré (hors Windows)")
        return 0
    import ctypes
    from ctypes import wintypes

    k32 = ctypes.WinDLL("kernel32", use_last_error=True)
    k32.CreateFileW.argtypes = [wintypes.LPCWSTR, wintypes.DWORD, wintypes.DWORD,
                                wintypes.LPVOID, wintypes.DWORD, wintypes.DWORD,
                                wintypes.HANDLE]
    k32.CreateFileW.restype = wintypes.HANDLE
    k32.CloseHandle.argtypes = [wintypes.HANDLE]
    generic_read, file_share_read, open_existing = 0x80000000, 0x1, 3
    invalide = wintypes.HANDLE(-1).value

    cfg = g.valider_config({
        "bailleur": {"nom": "Verrou"}, "periode": {"annee_debut": 2025, "annee_fin": 2025},
        "modules": {"mode_charges": "sans", "caf": False, "depot_garantie": False,
                    "documents": False},
        "locataires": [{"nom": "Ana", "identifiant": "M1", "loyer": 100,
                        "date_entree": "2025-01-01"}]})
    tmp = Path(tempfile.mkdtemp())
    f = tmp / "verrou.xlsx"
    g.generer_workbook(cfg, f)
    avant = f.read_bytes()

    # Comme Excel : lecture autorisée aux autres, écriture refusée.
    h = k32.CreateFileW(str(f), generic_read, file_share_read, None, open_existing, 0, None)
    assert h not in (None, invalide), f"verrou impossible (erreur {ctypes.get_last_error()})"
    try:
        try:
            g.generer_workbook(cfg, f)
            raise AssertionError("la génération aurait dû échouer : classeur verrouillé")
        except PermissionError as exc:
            assert "fermez-le" in str(exc).lower(), exc
        assert f.read_bytes() == avant, "classeur modifié malgré le verrou"
    finally:
        k32.CloseHandle(h)
    assert not list(tmp.glob("*.bak-*")), "copie de sécurité créée pour rien"

    # Verrou levé (classeur fermé dans Excel) : la génération repasse.
    g.generer_workbook(cfg, f)
    assert list(tmp.glob("*.bak-*")), "copie de sécurité attendue après remplacement"
    print("VERROU WINDOWS OK")
    return 0


if __name__ == "__main__":
    sys.exit(main())
