#!/usr/bin/env python3
"""
Publicación en una Página de Facebook (Graph API)
=================================================
La API de Facebook NO permite publicar en perfiles personales, solo en
Páginas. bot_meteo.py llama a publicar_facebook() con el mismo texto que
manda a Telegram (resumen de las 21:00 y curiosidades); aquí se convierte el
HTML de Telegram a texto plano y se publica en el muro de la Página.

Variables de entorno (/etc/bot-meteo.env). Si faltan, no se publica nada y
el bot sigue funcionando igual que siempre:
  FB_PAGE_ID        - ID numérico de la Página
  FB_PAGE_TOKEN     - Token de acceso de la Página (de larga duración)
  FB_API_VERSION    - Opcional, por defecto v25.0
  FB_PIE            - Opcional, texto que se añade al final de cada post
                      (ej: "📍 Estación meteorológica de Cartagena")

Uso manual:
  python facebook.py comprobar
      Comprueba que el token funciona y cuándo caduca (no publica nada)
  python facebook.py token APP_ID APP_SECRET TOKEN_USUARIO_CORTO
      Convierte el token de usuario del Graph API Explorer en un token de
      Página que no caduca y lo imprime para pegarlo en /etc/bot-meteo.env
"""

import os
import re
import sys
import logging
from html import unescape

import requests

log = logging.getLogger("bot_meteo")

FB_PAGE_ID = os.environ.get("FB_PAGE_ID", "")
FB_PAGE_TOKEN = os.environ.get("FB_PAGE_TOKEN", "")
FB_API_VERSION = os.environ.get("FB_API_VERSION", "v25.0")
FB_PIE = os.environ.get("FB_PIE", "")
GRAPH = f"https://graph.facebook.com/{FB_API_VERSION}"


def facebook_activo():
    return bool(FB_PAGE_ID and FB_PAGE_TOKEN)


def html_a_texto(html):
    """Facebook no admite HTML ni negritas: quita etiquetas y desescapa."""
    texto = re.sub(r"<br\s*/?>", "\n", html, flags=re.I)
    texto = re.sub(r"<[^>]+>", "", texto)
    texto = unescape(texto)
    # Comandos del bot (/lluvia, /records…) no tienen sentido en Facebook
    texto = "\n".join(l for l in texto.split("\n") if not l.strip().startswith("/"))
    return re.sub(r"\n{3,}", "\n\n", texto).strip()


def publicar_facebook(html, pie=True):
    """Publica en la Página. Devuelve el id del post o None. Nunca lanza
    excepciones: un fallo de Facebook no debe tumbar el resumen de Telegram."""
    if not facebook_activo():
        return None
    texto = html_a_texto(html)
    if pie and FB_PIE:
        texto += "\n\n" + FB_PIE
    for intento in range(3):
        try:
            r = requests.post(
                f"{GRAPH}/{FB_PAGE_ID}/feed",
                data={"message": texto, "access_token": FB_PAGE_TOKEN},
                timeout=30,
            )
            datos = r.json()
            if r.ok and "id" in datos:
                log.info("Publicado en Facebook: %s", datos["id"])
                return datos["id"]
            err = datos.get("error", {})
            log.warning("Facebook rechazó la publicación (intento %d): %s",
                        intento + 1, err.get("message", r.text))
            # Token caducado / sin permisos: reintentar no sirve
            if err.get("code") in (10, 190, 200):
                break
        except Exception as e:
            log.warning("Error publicando en Facebook (intento %d): %s", intento + 1, e)
    return None


# ---------------------------------------------------------------------------
# Utilidades de configuración
# ---------------------------------------------------------------------------
def comprobar():
    if not facebook_activo():
        print("Faltan FB_PAGE_ID y/o FB_PAGE_TOKEN en el entorno")
        return 1
    r = requests.get(f"{GRAPH}/{FB_PAGE_ID}",
                     params={"fields": "name,link", "access_token": FB_PAGE_TOKEN}, timeout=30)
    print("Página:", r.json())
    r = requests.get(f"{GRAPH}/debug_token",
                     params={"input_token": FB_PAGE_TOKEN, "access_token": FB_PAGE_TOKEN},
                     timeout=30)
    d = r.json().get("data", {})
    caduca = d.get("expires_at")
    print("Token válido:", d.get("is_valid"))
    print("Caduca:", "nunca" if caduca == 0 else caduca)
    print("Permisos:", ", ".join(d.get("scopes", [])))
    return 0 if d.get("is_valid") else 1


def token_largo(app_id, app_secret, token_corto):
    # 1) token de usuario corto (1-2 h) -> token de usuario largo (60 días)
    r = requests.get(f"{GRAPH}/oauth/access_token", params={
        "grant_type": "fb_exchange_token", "client_id": app_id,
        "client_secret": app_secret, "fb_exchange_token": token_corto,
    }, timeout=30)
    largo = r.json().get("access_token")
    if not largo:
        print("Error al alargar el token:", r.json())
        return 1
    # 2) con el token largo, los tokens de Página que devuelve /me/accounts no caducan
    r = requests.get(f"{GRAPH}/me/accounts",
                     params={"fields": "id,name,access_token", "access_token": largo},
                     timeout=30)
    paginas = r.json().get("data", [])
    if not paginas:
        print("No se encontró ninguna Página:", r.json())
        return 1
    for p in paginas:
        print(f"\n# {p['name']}")
        print(f"FB_PAGE_ID={p['id']}")
        print(f"FB_PAGE_TOKEN={p['access_token']}")
    return 0


if __name__ == "__main__":
    logging.basicConfig(level=logging.INFO)
    if len(sys.argv) >= 2 and sys.argv[1] == "comprobar":
        sys.exit(comprobar())
    if len(sys.argv) == 5 and sys.argv[1] == "token":
        sys.exit(token_largo(*sys.argv[2:5]))
    print(__doc__)
    sys.exit(1)
