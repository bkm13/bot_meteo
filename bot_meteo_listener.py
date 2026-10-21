#!/usr/bin/env python3
"""
Bot Meteo Listener - Comandos de Telegram bajo demanda
======================================================
Servicio permanente (systemd, Restart=always) que escucha el chat de Telegram
mediante long polling y responde a comandos:

  /tiempo   - Condiciones actuales de la estación (última observación de WU)
  /hoy      - Extremos del día hasta ahora (máx/mín, racha, lluvia...)
  /records  - Récords históricos guardados en records.json
  /mensual  - Menú con botones: año -> mes -> resumen mensual completo
  /mes      - Resumen de un mes escrito a mano: /mes septiembre 2026,
              /mes 9 2026, /mes 09/2026 (también /septiembre 2026)
  /ayuda    - Lista de comandos

Usa las mismas variables de entorno que bot_meteo.py (/etc/bot-meteo.env).
Solo responde a mensajes del chat TELEGRAM_CHAT_ID.
"""

import os
import sys
import json
import time
import logging
from datetime import date
from pathlib import Path

import requests

# Reutilizamos la lógica de cálculo del script principal
sys.path.insert(0, str(Path(__file__).resolve().parent))
from bot_meteo import (  # noqa: E402
    descargar_observaciones, calcular_resumen, _flt, _fmt,
    leer_diario, meses_con_datos, resumen_mensual, NOMBRES_MES,
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


# ---------------------------------------------------------------------------
# Resumen mensual bajo demanda (/mensual con botones, /mes escrito)
# ---------------------------------------------------------------------------
ABREV_MES = ["", "Ene", "Feb", "Mar", "Abr", "May", "Jun",
             "Jul", "Ago", "Sep", "Oct", "Nov", "Dic"]


def teclado_anios(disponibles):
    botones = [{"text": str(a), "callback_data": f"y:{a}"} for a in disponibles]
    return {"inline_keyboard": [botones[i:i + 4] for i in range(0, len(botones), 4)]}


def teclado_meses(anio, meses, hay_mas_anios):
    botones = [
        {"text": ABREV_MES[m], "callback_data": f"m:{anio}:{m}"} for m in meses
    ]
    filas = [botones[i:i + 4] for i in range(0, len(botones), 4)]
    if hay_mas_anios:
        filas.append([{"text": "« Años", "callback_data": "a"}])
    return {"inline_keyboard": filas}


def menu_mensual(anio=None):
    """Devuelve (texto, teclado) del menú. Sin año: lista de años
    (o directamente los meses si solo hay un año con datos)."""
    disponibles = meses_con_datos()
    if not disponibles:
        return "Todavía no hay días registrados.", None
    if anio is None and len(disponibles) == 1:
        anio = next(iter(disponibles))
    if anio is None or anio not in disponibles:
        return "📅 <b>Resúmenes mensuales</b>\nElige un año:", teclado_anios(disponibles)
    return (
        f"📅 <b>Resúmenes mensuales — {anio}</b>\nElige un mes:",
        teclado_meses(anio, disponibles[anio], len(disponibles) > 1),
    )


def texto_resumen_mes(anio, mes):
    try:
        texto, _ = resumen_mensual(anio, mes, date.today(), leer_diario())
    except Exception as e:
        log.warning("Error calculando el resumen de %02d/%d: %s", mes, anio, e)
        return "⚠️ No se pudo leer el registro (¿se está guardando ahora mismo?). Prueba en un minuto."
    return texto or f"No hay datos registrados de {NOMBRES_MES[mes]} {anio}."


def interpretar_mes(args):
    """Convierte ['septiembre', '2026'], ['9', '2026'], ['09/2026'] o ['sept']
    en (año, mes). Sin año: el más reciente con ese mes que no sea futuro."""
    tokens = " ".join(args).replace("/", " ").replace("-", " ").lower().split()
    mes = anio = None
    for t in tokens:
        if t.isdigit():
            n = int(t)
            if n > 12:
                anio = n
            elif mes is None:
                mes = n
        else:
            for i, nombre in enumerate(NOMBRES_MES):
                if i and len(t) >= 3 and nombre.startswith(t):
                    mes = i
                    break
    if not mes or not 1 <= mes <= 12:
        return None
    if anio is None:
        hoy = date.today()
        anio = hoy.year if mes <= hoy.month else hoy.year - 1
    return anio, mes


AYUDA = (
    "🤖 <b>Bot Meteo — comandos</b>\n\n"
    "/tiempo — condiciones actuales de la estación\n"
    "/hoy — extremos del día hasta ahora\n"
    "/records — récords históricos\n"
    "/mensual — menú de resúmenes mensuales (botones)\n"
    "/mes septiembre 2026 — resumen de un mes concreto\n"
    "/ayuda — este mensaje"
)


# ---------------------------------------------------------------------------
# Telegram
# ---------------------------------------------------------------------------
def _api(metodo, payload):
    try:
        r = requests.post(f"{API}/{metodo}", json=payload, timeout=30)
        r.raise_for_status()
        return True
    except Exception as e:
        log.warning("Error en %s: %s", metodo, e)
        return False


def enviar(texto, teclado=None):
    payload = {"chat_id": TELEGRAM_CHAT_ID, "text": texto, "parse_mode": "HTML"}
    if teclado:
        payload["reply_markup"] = teclado
    _api("sendMessage", payload)


def editar(message_id, texto, teclado=None):
    payload = {"chat_id": TELEGRAM_CHAT_ID, "message_id": message_id,
               "text": texto, "parse_mode": "HTML"}
    if teclado:
        payload["reply_markup"] = teclado
    _api("editMessageText", payload)


def registrar_comandos():
    """Publica la lista de comandos para que aparezca en el menú '/' de Telegram."""
    _api("setMyCommands", {"commands": [
        {"command": "tiempo", "description": "Condiciones actuales"},
        {"command": "hoy", "description": "Extremos del día hasta ahora"},
        {"command": "mensual", "description": "Resúmenes mensuales (menú)"},
        {"command": "mes", "description": "Resumen de un mes: /mes septiembre 2026"},
        {"command": "records", "description": "Récords históricos"},
        {"command": "ayuda", "description": "Lista de comandos"},
    ]})


def procesar_callback(cq):
    """Pulsaciones de los botones del menú mensual."""
    msg = cq.get("message") or {}
    chat_id = str((msg.get("chat") or {}).get("id", ""))
    _api("answerCallbackQuery", {"callback_query_id": cq["id"]})  # quita el "cargando"
    if chat_id != str(TELEGRAM_CHAT_ID):
        return
    datos = cq.get("data") or ""
    log.info("Botón pulsado: %s", datos)
    try:
        if datos == "a":
            texto, teclado = menu_mensual()
            editar(msg["message_id"], texto, teclado)
        elif datos.startswith("y:"):
            texto, teclado = menu_mensual(int(datos[2:]))
            editar(msg["message_id"], texto, teclado)
        elif datos.startswith("m:"):
            _, anio, mes = datos.split(":")
            # El resumen va en un mensaje nuevo; el menú queda para seguir consultando
            enviar(texto_resumen_mes(int(anio), int(mes)))
    except Exception as e:
        log.warning("Error procesando botón %s: %s", datos, e)


def procesar_comando(texto):
    partes = texto.split()
    cmd = partes[0].split("@")[0].lower()  # admite /tiempo@MiBot
    args = partes[1:]
    # Atajo: /septiembre 2026 equivale a /mes septiembre 2026
    if cmd[1:] in NOMBRES_MES[1:]:
        args = [cmd[1:]] + args
        cmd = "/mes"
    if cmd == "/tiempo":
        obs = obtener_observacion_actual()
        enviar(mensaje_tiempo(obs) if obs else
               "⚠️ No se pudo obtener la observación actual de la estación.")
    elif cmd == "/hoy":
        enviar(mensaje_hoy())
    elif cmd == "/records":
        enviar(mensaje_records())
    elif cmd in ("/mensual", "/meses"):
        texto_menu, teclado = menu_mensual()
        enviar(texto_menu, teclado)
    elif cmd == "/mes":
        if not args:
            texto_menu, teclado = menu_mensual()
            enviar(texto_menu, teclado)
            return
        objetivo = interpretar_mes(args)
        if not objetivo:
            enviar("No entiendo el mes. Ejemplos: <code>/mes septiembre 2026</code>, "
                   "<code>/mes 9 2026</code>, <code>/mes 09/2026</code> — o usa /mensual.")
            return
        enviar(texto_resumen_mes(*objetivo))
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
    registrar_comandos()
    offset = None
    while True:
        try:
            # allowed_updates debe ir serializado en JSON; Telegram lo recuerda
            # entre llamadas, por eso se pide explícitamente callback_query
            params = {"timeout": 50,
                      "allowed_updates": json.dumps(["message", "callback_query"])}
            if offset is not None:
                params["offset"] = offset
            r = requests.get(f"{API}/getUpdates", params=params, timeout=60)
            r.raise_for_status()
            for update in r.json().get("result", []):
                offset = update["update_id"] + 1
                if update.get("callback_query"):
                    procesar_callback(update["callback_query"])
                    continue
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
