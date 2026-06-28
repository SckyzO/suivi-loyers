# Image de génération des classeurs de suivi des loyers.
# Aucune installation locale : tout tourne dans ce conteneur.
FROM python:3.12-slim

WORKDIR /app

# Dépendances Python (couche cachée tant que requirements.txt ne change pas).
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

# Code du générateur (+ templates de style de graphique Microsoft).
COPY generer_suivi_loyers.py chart_style.py ./

# Les configs (entrée) et la sortie sont montées au runtime via -v.
#   configs/  -> /app/configs   (lecture)
#   sorties/  -> /app/sorties   (écriture)
ENTRYPOINT ["python", "generer_suivi_loyers.py"]
CMD ["configs/exemple.yaml"]
