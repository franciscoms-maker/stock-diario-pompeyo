#!/bin/bash
# Uso: ./publicar.sh /ruta/al/data.json
# Valida el JSON, lo copia al repo y lo sube si cambió. GitHub Pages lo publica solo.
set -euo pipefail
SRC="${1:?Uso: ./publicar.sh /ruta/al/data.json}"
cd "$(dirname "$0")"
python3 - "$SRC" <<'E'
import json, sys
d = json.load(open(sys.argv[1]))
assert isinstance(d, list) and d, "data.json vacío o no es una lista"
for r in d:
    assert {"fecha", "grupo", "sucursal", "estado"} <= r.keys(), f"fila inválida: {r}"
print(f"OK: {len(d)} filas")
E
[ "$SRC" -ef data.json ] || cp "$SRC" data.json
git add data.json
git diff --cached --quiet && { echo "Sin cambios."; exit 0; }
git commit -qm "Actualiza data.json $(TZ=America/Santiago date '+%d-%m-%Y %H:%M')"
git pull -q --rebase origin main
git push -q origin main
echo "Publicado."
