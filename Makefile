# Makefile — Suivi des loyers
#
# A lancer a chaque fin de modification :
#     make            # build + génération des exemples + smoke test + sync Windows
#
# Cibles utiles :
#     make build      # construit l'image Docker
#     make test       # smoke test du moteur
#     make gen        # génère les classeurs d'exemple dans exemples/
#     make sync       # source + exemples -> dossier Windows
#     make clean      # supprime les exemples générés

SHELL := /bin/bash
IMAGE := suivi-loyers:latest
COMPOSE := docker compose
DOWNLOADS := $(HOME)/Downloads
WIN_PROJECT := $(DOWNLOADS)/Suivi des loyers sur Excel
EXEMPLES := exemples

# Exécuter les conteneurs avec l'UID/GID courant : pas de fichiers créés en root.
USERFLAG := --user $(shell id -u):$(shell id -g)

# Fichiers source synchronisés côté Windows (pour builder l'.exe via build.bat).
SRC := generer_suivi_loyers.py interface.py build.bat requirements.txt \
       Dockerfile docker-compose.yml Makefile README.md CLAUDE.md .gitignore

.DEFAULT_GOAL := all
.PHONY: all build gen test sync sync-win sync-exemples clean

all: build gen test sync ## Workflow complet de fin de modification

build: ## Construit l'image Docker
	docker build -t $(IMAGE) .

gen: build ## Génère les classeurs d'exemple dans exemples/
	@mkdir -p $(EXEMPLES)
	$(COMPOSE) run --rm $(USERFLAG) suivi configs/exemple.yaml $(EXEMPLES)
	$(COMPOSE) run --rm $(USERFLAG) suivi configs/minimal.yaml $(EXEMPLES)

test: build ## Smoke test du moteur (structure, modularité, préservation)
	docker run --rm $(USERFLAG) --entrypoint python -e PYTHONPATH=/app -v "$(CURDIR):/app" $(IMAGE) tests/smoke.py

sync: sync-win sync-exemples ## Synchronise source + exemples vers Windows

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

sync-exemples: ## Copie les classeurs d'exemple dans le dossier Windows (exemples/)
	@shopt -s nullglob; \
	if [ -d "$(WIN_PROJECT)" ]; then \
	  mkdir -p "$(WIN_PROJECT)/$(EXEMPLES)"; \
	  for f in $(EXEMPLES)/*.xlsx; do \
	    base=$$(basename "$$f"); \
	    if cp "$$f" "$(WIN_PROJECT)/$(EXEMPLES)/$$base" 2>/dev/null; then \
	      echo "exemple: $$base"; \
	    else \
	      cp "$$f" "$(WIN_PROJECT)/$(EXEMPLES)/$${base%.xlsx}_frais.xlsx" && \
	        echo "verrouillé (Excel ouvert) -> $${base%.xlsx}_frais.xlsx"; \
	    fi; \
	  done; \
	else \
	  echo "Dossier Windows absent, sync-exemples ignorée"; \
	fi

clean: ## Supprime les exemples générés
	@rm -rf $(EXEMPLES) && echo "$(EXEMPLES)/ nettoyé"
