#!/usr/bin/env python3
"""
Bot Meteo Listener - Comandos de Telegram bajo demanda
======================================================
Servicio permanente (systemd, Restart=always) que escucha el chat de Telegram
mediante long polling y responde a comandos:

  /tiempo   - Condiciones actuales de la estación (última observación de WU)
  /hoy      - Extremos del día hasta ahora (máx/mín, racha, lluvia...)
  /records  - Récords históricos guardados en records.json
  /ayuda    - Lista de comandos

Usa las mismas variables de entorno que bot_meteo.py (/etc/bot-meteo.env).
Solo responde a mensajes del chat TELEGRAM_CHAT_ID.
"""

import os
import sys
import json
import time
import logging
from pathlib import Path

import requests

# Reutilizamos la lógica de cálculo del script principal
sys.path.insert(0, str(Path(__file__).resolve().parent))
from bot_meteo import (  # noqa: E402
    descargar_observaciones, calcular_resumen, _flt, _fmt,
)

BASE_DIR = Path(__file__).resolve().parent
RECORDS_PATH = BASE_DIR / "records.json"
LOG_PATH = BASE_DIR / "bot_meteo_listener.log"

WU_API_KEY = os.environ.get("WU_API_KEY", "")
WU_STATION_ID = os.environ.get("WU_STATION_ID", "")
TELEGRAM_TOKEN = os.environ.get("TELEGRAM_TOKEN", "")
TELEGRAM_CHAT_ID = os.environ.get("TELEGRAM_CHAT_ID", "")

WU_CURRENT_URL = (
    "https://api.weather.com/v2/pws/observations/current"
    "?stationId={station}&format=json&units=m&numericPrecision=decimal&apiKey={key}"
)
API = f"https://api.telegram.org/bot{TELEGRAM_TOKEN}"

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    handlers=[logging.FileHandler(LOG_PATH), logging.StreamHandler()],
)
log = logging.getLogger("bot_meteo_listener")


# ---------------------------------------------------------------------------
# Datos actuales
# ---------------------------------------------------------------------------
def obtener_observacion_actual():
    """Última observación de la estación (endpoint 'current'). None si falla."""
    url = WU_CURRENT_URL.format(station=WU_STATION_ID, key=WU_API_KEY)
    for intento in range(3):
        try:
            r = requests.get(url, timeout=20)
            r.raise_for_status()
            obs = r.json().get("observations", [])
            if obs:
                return obs[0]
        except Exception as e:
            log.warning("Error obteniendo observación actual (intento %d): %s",
                        intento + 1, e)
    return None


RUMBOS = ["N", "NNE", "NE", "ENE", "E", "ESE", "SE", "SSE",
          "S", "SSO", "SO", "OSO", "O", "ONO", "NO", "NNO"]


def rumbo_viento(grados):
    g = _flt(grados)
    if g is None:
        return ""
    return RUMBOS[int((g + 11.25) // 22.5) % 16]


def mensaje_tiempo(obs):
    m = obs.get("metric", {}) or {}
    hora = obs.get("obsTimeLocal", "")[11:16] or "—"

    temp = _flt(m.get("temp"))
    sensacion = _flt(m.get("heatIndex")) if (temp or 0) >= 21 else _flt(m.get("windChill"))
    rocio = _flt(m.get("dewpt"))
    humedad = _flt(obs.get("humidity"))
    viento = _flt(m.get("windSpeed"))
    racha = _flt(m.get("windGust"))
    rumbo = rumbo_viento(obs.get("winddir"))
    presion = _flt(m.get("pressure"))
    lluvia_dia = _flt(m.get("precipTotal"))
    lluvia_ritmo = _flt(m.get("precipRate"))
    solar = _flt(obs.get("solarRadiation"))
    uv = _flt(obs.get("uv"))

    lineas = [
        f"⛅️ <b>Ahora mismo — {WU_STATION_ID}</b> (obs. de las {hora})",
        "",
        f"🌡 Temperatura: <b>{_fmt(temp, '°C')}</b>",
    ]
    if sensacion is not None and temp is not None and abs(sensacion - temp) >= 0.5:
        lineas.append(f"🥵 Sensación: <b>{_fmt(sensacion, '°C')}</b>")
    lineas += [
        f"💧 Humedad: <b>{_fmt(humedad, '%')}</b>  ·  Rocío: {_fmt(rocio, '°C')}",
        f"💨 Viento: <b>{_fmt(viento, ' km/h')}</b> {rumbo}"
        + (f"  ·  Racha: {_fmt(racha, ' km/h')}" if racha is not None else ""),
        f"🔵 Presión: <b>{_fmt(presion, ' hPa')}</b>",
    ]
    if lluvia_ritmo is not None and lluvia_ritmo > 0:
        lineas.append(f"🌧 Lloviendo ahora: <b>{lluvia_ritmo:.1f} mm/h</b>")
    lineas.append(f"🌧 Lluvia hoy: <b>{_fmt(lluvia_dia, ' mm')}</b>")
    if solar is not None:
        lineas.append(f"☀️ Radiación: <b>{_fmt(solar, ' W/m²')}</b>")
    if uv is not None:
        lineas.append(f"🕶 UV: <b>{_fmt(uv)}</b>")
    return "\n".join(lineas)


def mensaje_hoy():
    """Extremos del día en curso (reutiliza la lógica del resumen diario)."""
    observaciones = descargar_observaciones()
    if not observaciones:
        return "⚠️ No se pudieron obtener las observaciones del día."
    r = calcular_resumen(observaciones)
    lineas = [
        "📈 <b>Extremos de hoy (hasta ahora)</b>",
        "",
        f"🌡 Máxima: <b>{_fmt(r['temp_max'], '°C')}</b> ({r['hora_temp_max'] or '—'})",
        f"❄️ Mínima: <b>{_fmt(r['temp_min'], '°C')}</b> ({r['hora_temp_min'] or '—'})",
        f"💨 Racha máx: <b>{_fmt(r['racha_max'], ' km/h')}</b> ({r['hora_racha'] or '—'})",
        f"🌧 Lluvia: <b>{r['lluvia']:.1f} mm</b>",
        f"🔵 Presión: {_fmt(r['presion_min'])} – {_fmt(r['presion_max'])} hPa",
        f"💧 Humedad: {_fmt(r['humedad_min'], '%')} – {_fmt(r['humedad_max'], '%')}",
    ]
    if r["solar_max"] is not None:
        lineas.append(f"☀️ Radiación máx: <b>{_fmt(r['solar_max'], ' W/m²')}</b>")
    if r["uv_max"] is not None:
        lineas.append(f"🕶 UV máx: <b>{_fmt(r['uv_max'])}</b>")
    return "\n".join(lineas)


ETIQUETAS_RECORDS = {
    "temp_max": ("🔥 Temperatura máxima", "°C"),
    "temp_min": ("🧊 Temperatura mínima", "°C"),
    "racha_max": ("💨 Racha de viento", "km/h"),
    "lluvia_dia": ("🌧 Lluvia en un día", "mm"),
    "presion_max": ("🔵 Presión máxima", "hPa"),
    "presion_min": ("🔵 Presión mínima", "hPa"),
    "solar_max": ("☀️ Radiación solar", "W/m²"),
    "uv_max": ("🕶 Índice UV", ""),
}


def mensaje_records():
    if not RECORDS_PATH.exists():
        return "Todavía no hay récords guardados."
    try:
        records = json.loads(RECORDS_PATH.read_text(encoding="utf-8"))
    except Exception:
        return "⚠️ No se pudo leer records.json."
    lineas = ["🏆 <b>Récords históricos</b>", ""]
    for clave, (etiqueta, unidad) in ETIQUETAS_RECORDS.items():
        rec = records.get(clave)
        if not rec:
            continue
        hora = f" a las {rec['hora']}" if rec.get("hora") else ""
        lineas.append(
            f"{etiqueta}: <b>{rec['valor']} {unidad}</b>".rstrip()
            + f" ({rec['fecha']}{hora})"
        )
    return "\n".join(lineas)


AYUDA = (
    "🤖 <b>Bot Meteo — comandos</b>\n\n"
    "/tiempo — condiciones actuales de la estación\n"
    "/hoy — extremos del día hasta ahora\n"
    "/records — récords históricos\n"
    "/ayuda — este mensaje"
)


# ---------------------------------------------------------------------------
# Telegram
# ---------------------------------------------------------------------------
def enviar(texto):
    try:
        r = requests.post(
            f"{API}/sendMessage",
            json={"chat_id": TELEGRAM_CHAT_ID, "text": texto, "parse_mode": "HTML"},
            timeout=30,
        )
        r.raise_for_status()
    except Exception as e:
        log.warning("Error enviando mensaje: %s", e)


def procesar_comando(texto):
    cmd = texto.split()[0].split("@")[0].lower()  # admite /tiempo@MiBot
    if cmd == "/tiempo":
        obs = obtener_observacion_actual()
        enviar(mensaje_tiempo(obs) if obs else
               "⚠️ No se pudo obtener la observación actual de la estación.")
    elif cmd == "/hoy":
        enviar(mensaje_hoy())
    elif cmd == "/records":
        enviar(mensaje_records())
    elif cmd in ("/ayuda", "/help", "/start"):
        enviar(AYUDA)


def main():
    faltan = [
        n for n, v in [
            ("WU_API_KEY", WU_API_KEY), ("WU_STATION_ID", WU_STATION_ID),
            ("TELEGRAM_TOKEN", TELEGRAM_TOKEN), ("TELEGRAM_CHAT_ID", TELEGRAM_CHAT_ID),
        ] if not v
    ]
    if faltan:
        log.error("Faltan variables de entorno: %s", ", ".join(faltan))
        sys.exit(1)

    log.info("Listener iniciado; escuchando comandos de Telegram...")
    offset = None
    while True:
        try:
            params = {"timeout": 50, "allowed_updates": ["message"]}
            if offset is not None:
                params["offset"] = offset
            r = requests.get(f"{API}/getUpdates", params=params, timeout=60)
            r.raise_for_status()
            for update in r.json().get("result", []):
                offset = update["update_id"] + 1
                msg = update.get("message") or {}
                chat_id = str((msg.get("chat") or {}).get("id", ""))
                texto = (msg.get("text") or "").strip()
                if chat_id != str(TELEGRAM_CHAT_ID) or not texto.startswith("/"):
                    continue
                log.info("Comando recibido: %s", texto)
                procesar_comando(texto)
        except KeyboardInterrupt:
            break
        except Exception as e:
            log.warning("Error en el bucle principal: %s", e)
            time.sleep(10)  # evita bucles agresivos si Telegram falla


if __name__ == "__main__":
    main()
