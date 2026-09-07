#!/bin/bash
# Corre el scraper. Si el proyecto está conectado a GitHub, además sincroniza:
# trae lo que subió la nube y sube lo que consiguió esta máquina (Paraná, que
# Cloudflare no deja leer desde los servidores de GitHub).
cd "$(dirname "$0")" || exit 1

if git rev-parse --git-dir >/dev/null 2>&1 && git remote get-url origin >/dev/null 2>&1; then
    echo "Sincronizando con GitHub..."
    git pull --rebase -q -X theirs origin main 2>&1 | tail -2
fi

./.venv/bin/python -u scraper.py "$@"
CODIGO=$?

if [ $CODIGO -eq 0 ] && git rev-parse --git-dir >/dev/null 2>&1 && git remote get-url origin >/dev/null 2>&1; then
    git add -A licitaciones.csv estado.json Licitaciones-Entre-Rios.xlsx 2>/dev/null
    if ! git diff --staged --quiet 2>/dev/null; then
        git -c user.name="Delfina" -c user.email="delfina@granssrl.com.ar" \
            commit -q -m "Licitaciones desde la Mac (incluye Paraná) - $(date '+%d/%m/%Y %H:%M')"
        git push -q origin main 2>&1 | tail -2 && echo "Subido a GitHub: la planilla ya tiene lo de Paraná."
    else
        echo "Sin cambios para subir."
    fi
fi
exit $CODIGO
