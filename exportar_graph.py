#!/usr/bin/env python3
"""Lee Tabla1 del Excel por Microsoft Graph (credenciales de aplicación) y publica data.json.

Solo LECTURA: usa GET; no escribe nada en el Excel. Sin dependencias (solo librería estándar).
Uso: python3 exportar_graph.py --env /ruta/.env [--sin-publicar]
El .env debe traer tenant, client id y secret (nombres aceptados abajo). No se guarda ni imprime el secreto.
"""
import datetime as dt
import json
import os
import subprocess
import sys
import urllib.parse
import urllib.request
from pathlib import Path

AQUI = Path(__file__).resolve().parent
UPN = "francisco.marambio@pompeyo.cl"
ITEM_ID = "01CFBOCQQ2ENMHBBNO35ALPFGRXCJ2NMLF"
TABLA = "Tabla1"
NOMBRES = {
    "tenant": ["AZURE_TENANT_ID", "TENANT_ID", "MS_TENANT_ID", "GRAPH_TENANT_ID"],
    "client": ["AZURE_CLIENT_ID", "CLIENT_ID", "MS_CLIENT_ID", "GRAPH_CLIENT_ID"],
    "secret": ["AZURE_CLIENT_SECRET", "CLIENT_SECRET", "MS_CLIENT_SECRET", "GRAPH_CLIENT_SECRET"],
}


def cargar_env(ruta):
    env = dict(os.environ)
    for linea in Path(ruta).read_text().splitlines():
        linea = linea.strip()
        if linea and not linea.startswith("#") and "=" in linea:
            k, v = linea.split("=", 1)
            env[k.strip().removeprefix("export ").strip()] = v.strip().strip('"').strip("'")
    return env


def credencial(env, clave):
    for n in NOMBRES[clave]:
        if env.get(n):
            return env[n]
    sys.exit(f"Falta en el .env una variable para '{clave}' (acepto: {', '.join(NOMBRES[clave])})")


def http(url, data=None, token=None):
    req = urllib.request.Request(url, data=data, headers={"Authorization": f"Bearer {token}"} if token else {})
    with urllib.request.urlopen(req, timeout=60) as r:
        return json.load(r)


def obtener_token(env):
    cuerpo = urllib.parse.urlencode({
        "grant_type": "client_credentials", "client_id": credencial(env, "client"),
        "client_secret": credencial(env, "secret"), "scope": "https://graph.microsoft.com/.default",
    }).encode()
    return http(f"https://login.microsoftonline.com/{credencial(env, 'tenant')}/oauth2/v2.0/token", cuerpo)["access_token"]


def serial(v):
    return dt.datetime(1899, 12, 30) + dt.timedelta(days=v)


def fecha(v):
    if isinstance(v, (int, float)):
        return serial(v).strftime("%Y-%m-%d")
    s = str(v).strip()
    return f"{s[6:10]}-{s[3:5]}-{s[0:2]}" if len(s) >= 10 and s[2] in "-/" else s[:10]


def hora(v):
    if v in ("", None):
        return ""
    if isinstance(v, (int, float)):
        m = round((v % 1) * 1440)
        return f"{m // 60 % 24:02d}:{m % 60:02d}"
    return str(v).strip()[:5]


def convertir(valores):
    """valores: lista de filas [Fecha, Grupo, Sucursal, Estado, Registrado, Hora llegada, Puntualidad]."""
    out = []
    for v in valores:
        v = list(v) + [""] * 7
        if v[0] in ("", None) or not v[1] or not v[2]:
            continue
        reg = serial(v[4]).strftime("%Y-%m-%d %H:%M") if isinstance(v[4], (int, float)) else str(v[4]).strip()
        out.append({"fecha": fecha(v[0]), "grupo": str(v[1]).strip(), "sucursal": str(v[2]).strip(),
                    "estado": str(v[3] or "").strip(), "registrado": reg,
                    "horaLlegada": hora(v[5]), "puntualidad": str(v[6] or "").strip()})
    return out


def leer_filas(token):
    url = (f"https://graph.microsoft.com/v1.0/users/{UPN}/drive/items/{ITEM_ID}"
           f"/workbook/tables/{TABLA}/rows?$top=1000")
    filas = []
    while url:
        j = http(url, token=token)
        filas += [x["values"][0] for x in j["value"]]
        url = j.get("@odata.nextLink")
    return filas


if __name__ == "__main__":
    if "--env" not in sys.argv:
        sys.exit(__doc__)
    env = cargar_env(sys.argv[sys.argv.index("--env") + 1])
    datos = convertir(leer_filas(obtener_token(env)))
    if not datos:
        sys.exit("La tabla no devolvió filas; no se publica nada.")
    destino = AQUI / "data.json"
    destino.write_text(json.dumps(datos, ensure_ascii=False), encoding="utf-8")
    print(f"{len(datos)} filas leídas")
    if "--sin-publicar" not in sys.argv:
        subprocess.run([str(AQUI / "publicar.sh"), str(destino)], check=True)
