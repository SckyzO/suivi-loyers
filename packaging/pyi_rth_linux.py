"""Runtime hook PyInstaller du binaire Linux (Dockerfile.linux).

Exécuté au démarrage du binaire, avant interface_flet. Ne modifie que
l'environnement hérité par le client Flutter (sous-processus) : le chargeur
dynamique du processus Python a déjà lu LD_LIBRARY_PATH à son lancement.
"""
import os

# 1. flet pack embarque l'archive du client Flet sous le nom de la distro du
#    BUILD, mais au lancement flet_desktop recalcule ce nom d'après la glibc de
#    l'HÔTE (ex. ubuntu24.04) : l'archive n'est pas trouvée et il tente un
#    téléchargement GitHub. On fige la distro du build (valeur injectée par
#    Dockerfile.linux ; setdefault : FLET_LINUX_DISTRO reste surchargeable).
os.environ.setdefault("FLET_LINUX_DISTRO", "@FLET_LINUX_DISTRO@")

# 2. Le bootloader PyInstaller place ses bibliothèques embarquées (libstdc++,
#    libz… de Debian, dont celles de Pillow) en tête de LD_LIBRARY_PATH. Le
#    client Flutter en hérite, et le pilote Mesa de l'hôte, chargé avec ces
#    versions, échoue : « No provider of eglGetPlatformDisplayEXT found », le
#    client s'arrête. On rend aux enfants le LD_LIBRARY_PATH d'origine.
if "LD_LIBRARY_PATH_ORIG" in os.environ:
    os.environ["LD_LIBRARY_PATH"] = os.environ["LD_LIBRARY_PATH_ORIG"]
else:
    os.environ.pop("LD_LIBRARY_PATH", None)
