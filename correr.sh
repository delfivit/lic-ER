#!/bin/bash
# Corre el scraper. Si el proyecto está conectado a GitHub, además sincroniza:
# trae lo que subió la nube y sube lo que consiguió esta máquina (Paraná, que
# Cloudflare no deja leer desde los servidores de GitHub).
cd "$(dirname "$0")" || exit 1
set -o pipefail

# A las 8 la Mac suele estar dormida: macOS la despierta sin pantalla solo
# para correr esto, y el wifi tarda en reconectar. Si se arranca sin
# esperar, TODAS las fuentes fallan por falta de red (no porque los sitios
# estén caídos) y se pierden ~30 minutos reintentando igual sin suerte.
# Por eso acá se espera a que haya internet de verdad antes de seguir.
echo "Esperando conexión a internet..."
ESPERADO=0
TOPE_ESPERA=240
while ! curl -s --max-time 5 -o /dev/null https://www.google.com; do
    ESPERADO=$((ESPERADO+5))
    if [ $ESPERADO -ge $TOPE_ESPERA ]; then
        echo "Sin internet después de ${TOPE_ESPERA}s de espera. No se corrió hoy ($(date '+%d/%m/%Y %H:%M'))."
        exit 1
    fi
    sleep 5
done
echo "Internet OK (tardó ${ESPERADO}s). Arrancando."

if git rev-parse --git-dir >/dev/null 2>&1 && git remote get-url origin >/dev/null 2>&1; then
    echo "Sincronizando con GitHub..."
    git pull --rebase -q -X theirs origin main 2>&1 | tail -2
fi

./.venv/bin/python -u scraper.py "$@"
CODIGO=$?

if [ $CODIGO -eq 0 ] && git rev-parse --git-dir >/dev/null 2>&1 && git remote get-url origin >/dev/null 2>&1; then
    git add -A licitaciones.csv diagnostico.csv estado.json Licitaciones-Entre-Rios.xlsx 2>/dev/null
    if ! git diff --staged --quiet 2>/dev/null; then
        git -c user.name="Delfina" -c user.email="delfina@granssrl.com.ar" \
            commit -q -m "Licitaciones desde la Mac (incluye Paraná) - $(date '+%d/%m/%Y %H:%M')"
        if git push -q origin main 2>&1 | tail -2; then
            echo "Subido a GitHub: la planilla ya tiene lo de Paraná."
        else
            echo "No se pudo subir a GitHub (sin conexión o conflicto)."
        fi
    else
        echo "Sin cambios para subir."
    fi
fi
exit $CODIGO
