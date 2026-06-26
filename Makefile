# Makefile — Suivi des loyers
#
# A lancer a chaque fin de modification :
#     make            # build + génération exemples + smoke test + sync Windows/Downloads
#
# Cibles utiles :
#     make build      # construit l'image Docker
#     make test       # smoke test du moteur
#     make sync       # source -> dossier Windows, xlsx -> ~/Downloads
#     make clean      # supprime les sorties générées

SHELL := /bin/bash
IMAGE := suivi-loyers:latest
COMPOSE := docker compose
DOWNLOADS := $(HOME)/Downloads
WIN_PROJECT := $(DOWNLOADS)/Suivi des loyers sur Excel

# Fichiers source synchronisés côté Windows (pour builder l'.exe via build.bat).
SRC := generer_suivi_loyers.py interface.py build.bat requirements.txt \
       Dockerfile docker-compose.yml Makefile README.md CLAUDE.md .gitignore

.DEFAULT_GOAL := all
.PHONY: all build gen test sync sync-win sync-downloads clean

all: build gen test sync ## Workflow complet de fin de modification

build: ## Construit l'image Docker
	docker build -t $(IMAGE) .

gen: build ## Génère les classeurs d'exemple dans sorties/
	@mkdir -p sorties
	$(COMPOSE) run --rm suivi configs/exemple.yaml
	$(COMPOSE) run --rm suivi configs/minimal.yaml

test: build ## Smoke test du moteur (structure, modularité, préservation)
	docker run --rm --entrypoint python -e PYTHONPATH=/app -v "$(CURDIR):/app" $(IMAGE) tests/smoke.py

sync: sync-win sync-downloads ## Synchronise vers Windows et Downloads

sync-win: ## Copie le code à jour dans le dossier Windows (pour build.bat)
	@if [ -d "$(DOWNLOADS)" ]; then \
	  mkdir -p "$(WIN_PROJECT)/configs"; \
	  cp $(SRC) "$(WIN_PROJECT)/"; \
	  cp configs/*.yaml "$(WIN_PROJECT)/configs/"; \
	  rm -rf "$(WIN_PROJECT)/__pycache__"; \
	  echo "Source synchronisée -> $(WIN_PROJECT)"; \
	else \
	  echo "Dossier Windows absent, sync-win ignorée"; \
	fi

sync-downloads: ## Copie les classeurs générés dans ~/Downloads (gère le verrou Excel)
	@shopt -s nullglob; \
	for f in sorties/*.xlsx; do \
	  base=$$(basename "$$f"); \
	  if cp "$$f" "$(DOWNLOADS)/$$base" 2>/dev/null; then \
	    echo "copié: $$base"; \
	  else \
	    cp "$$f" "$(DOWNLOADS)/$${base%.xlsx}_frais.xlsx" && \
	      echo "verrouillé (Excel ouvert) -> $${base%.xlsx}_frais.xlsx"; \
	  fi; \
	done

clean: ## Supprime les sorties générées (fichiers créés en root par Docker)
	@docker run --rm --entrypoint rm -v "$(CURDIR)/sorties:/app/sorties" $(IMAGE) -rf /app/sorties 2>/dev/null || true
	@rm -rf sorties && echo "sorties/ nettoyé"
