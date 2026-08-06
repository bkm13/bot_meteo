#!/usr/bin/env python3
"""
Bot Meteo - Resumen diario de estación meteorológica (Weather Underground)
==========================================================================
Se ejecuta una vez al día (systemd timer, 21:00 Europe/Madrid):
  1. Descarga todas las observaciones del día desde la API PWS de WU
  2. Calcula máximos/mínimos con hora exacta
  3. Envía resumen por Telegram
  4. Registra el día en registro_meteo.xlsx (hoja "Diario")
  5. Comprueba récords históricos (records.json); si se bate alguno,
     lo anuncia y lo guarda en la hoja "Récords" del Excel

Variables de entorno necesarias (/etc/bot-meteo.env):
  WU_API_KEY        - API key de Weather Underground
  WU_STATION_ID     - ID de la estación (ej: ICARTA12)
  TELEGRAM_TOKEN    - Token del bot de Telegram (BotFather)
  TELEGRAM_CHAT_ID  - Chat ID donde enviar los mensajes
"""

import os
import sys
import json
import logging
from datetime import datetime, date
from pathlib import Path

import requests
from openpyxl import Workbook, load_workbook

# ---------------------------------------------------------------------------
# Configuración
# ---------------------------------------------------------------------------
BASE_DIR = Path(__file__).resolve().parent
EXCEL_PATH = BASE_DIR / "registro_meteo.xlsx"
RECORDS_PATH = BASE_DIR / "records.json"
LOG_PATH = BASE_DIR / "bot_meteo.log"

WU_API_KEY = os.environ.get("532aea9c11f14d11aaea9c11f12d11c4", "")
WU_STATION_ID = os.environ.get("IMONTE1305", "")
TELEGRAM_TOKEN = os.environ.get("8600853013:AAFpFVZ2mz6Ku8IqMbSNUEL5EV1HstOPbnk", "")
TELEGRAM_CHAT_ID = os.environ.get("8986639509", "")

WU_URL = (
    "https://api.weather.com/v2/pws/observations/all/1day"
    "?stationId={station}&format=json&units=m&numericPrecision=decimal&apiKey={key}"
)

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    handlers=[logging.FileHandler(LOG_PATH), logging.StreamHandler()],
)
log = logging.getLogger("bot_meteo")

CABECERA_DIARIO = [
    "Fecha", "Temp Máx (°C)", "Hora Temp Máx", "Temp Mín (°C)", "Hora Temp Mín",
    "Racha Máx (km/h)", "Hora Racha", "Lluvia (mm)",
    "Presión Máx (hPa)", "Presión Mín (hPa)",
    "Humedad Máx (%)", "Humedad Mín (%)", "Nº Observaciones",
]
CABECERA_RECORDS = ["Tipo de récord", "Valor", "Fecha", "Hora", "Récord anterior", "Detalles"]


# ---------------------------------------------------------------------------
# Descarga y cálculo
# ---------------------------------------------------------------------------
def descargar_observaciones():
    """Descarga todas las observaciones del día. Devuelve lista o None."""
    url = WU_URL.format(station=WU_STATION_ID, key=WU_API_KEY)
    for intento in range(3):
        try:
            r = requests.get(url, timeout=30)
            r.raise_for_status()
            obs = r.json().get("observations", [])
            if obs:
                return obs
            log.warning("Respuesta sin observaciones (intento %d)", intento + 1)
        except Exception as e:
            log.warning("Error descargando WU (intento %d): %s", intento + 1, e)
    return None


def _hora(obs):
    """Extrae HH:MM de obsTimeLocal ('2026-08-06 14:35:00')."""
    try:
        return obs.get("obsTimeLocal", "")[11:16]
    except Exception:
        return "?"


def _flt(valor):
    try:
        return float(valor)
    except (TypeError, ValueError):
        return None


def calcular_resumen(observaciones):
    """Recorre las observaciones del día y extrae extremos con su hora."""
    resumen = {
        "temp_max": None, "hora_temp_max": "",
        "temp_min": None, "hora_temp_min": "",
        "racha_max": None, "hora_racha": "",
        "lluvia": 0.0,
        "presion_max": None, "presion_min": None,
        "humedad_max": None, "humedad_min": None,
        "n_obs": len(observaciones),
    }

    for obs in observaciones:
        m = obs.get("metric", {}) or {}

        t_hi = _flt(m.get("tempHigh"))
        if t_hi is not None and (resumen["temp_max"] is None or t_hi > resumen["temp_max"]):
            resumen["temp_max"] = t_hi
            resumen["hora_temp_max"] = _hora(obs)

        t_lo = _flt(m.get("tempLow"))
        if t_lo is not None and (resumen["temp_min"] is None or t_lo < resumen["temp_min"]):
            resumen["temp_min"] = t_lo
            resumen["hora_temp_min"] = _hora(obs)

        # Racha: usamos windgustHigh; si la estación no lo da, windspeedHigh
        racha = _flt(m.get("windgustHigh"))
        if racha is None:
            racha = _flt(m.get("windspeedHigh"))
        if racha is not None and (resumen["racha_max"] is None or racha > resumen["racha_max"]):
            resumen["racha_max"] = racha
            resumen["hora_racha"] = _hora(obs)

        p_hi = _flt(m.get("pressureMax"))
        if p_hi is not None and (resumen["presion_max"] is None or p_hi > resumen["presion_max"]):
            resumen["presion_max"] = p_hi

        p_lo = _flt(m.get("pressureMin"))
        if p_lo is not None and (resumen["presion_min"] is None or p_lo < resumen["presion_min"]):
            resumen["presion_min"] = p_lo

        h_hi = _flt(obs.get("humidityHigh"))
        if h_hi is not None and (resumen["humedad_max"] is None or h_hi > resumen["humedad_max"]):
            resumen["humedad_max"] = h_hi

        h_lo = _flt(obs.get("humidityLow"))
        if h_lo is not None and (resumen["humedad_min"] is None or h_lo < resumen["humedad_min"]):
            resumen["humedad_min"] = h_lo

        # precipTotal es acumulado del día: nos quedamos con el mayor visto
        lluvia = _flt(m.get("precipTotal"))
        if lluvia is not None and lluvia > resumen["lluvia"]:
            resumen["lluvia"] = lluvia

    return resumen


# ---------------------------------------------------------------------------
# Récords históricos
# ---------------------------------------------------------------------------
DEFINICION_RECORDS = [
    # (clave_resumen, clave_record, etiqueta, unidad, es_maximo, clave_hora)
    ("temp_max", "temp_max", "Temperatura máxima", "°C", True, "hora_temp_max"),
    ("temp_min", "temp_min", "Temperatura mínima", "°C", False, "hora_temp_min"),
    ("racha_max", "racha_max", "Racha de viento", "km/h", True, "hora_racha"),
    ("lluvia", "lluvia_dia", "Lluvia en un día", "mm", True, None),
    ("presion_max", "presion_max", "Presión máxima", "hPa", True, None),
    ("presion_min", "presion_min", "Presión mínima", "hPa", False, None),
]


def cargar_records():
    if RECORDS_PATH.exists():
        try:
            return json.loads(RECORDS_PATH.read_text(encoding="utf-8"))
        except Exception as e:
            log.warning("records.json corrupto, se reinicia: %s", e)
    return {}


def comprobar_records(resumen, hoy):
    """Compara el día contra los récords. Devuelve lista de récords batidos."""
    records = cargar_records()
    batidos = []

    for clave_res, clave_rec, etiqueta, unidad, es_max, clave_hora in DEFINICION_RECORDS:
        valor = resumen.get(clave_res)
        if valor is None:
            continue
        if clave_rec == "lluvia_dia" and valor <= 0:
            continue  # un día sin lluvia no compite

        anterior = records.get(clave_rec)
        es_nuevo = anterior is None or (
            valor > anterior["valor"] if es_max else valor < anterior["valor"]
        )
        if es_nuevo:
            hora = resumen.get(clave_hora, "") if clave_hora else ""
            # El primer registro solo siembra el histórico, no se anuncia
            if anterior is not None:
                batidos.append({
                    "etiqueta": etiqueta,
                    "valor": valor,
                    "unidad": unidad,
                    "hora": hora,
                    "anterior": anterior,
                })
            records[clave_rec] = {
                "valor": valor,
                "fecha": hoy.isoformat(),
                "hora": hora,
            }

    RECORDS_PATH.write_text(
        json.dumps(records, indent=2, ensure_ascii=False), encoding="utf-8"
    )
    return batidos


# ---------------------------------------------------------------------------
# Excel
# ---------------------------------------------------------------------------
def abrir_excel():
    if EXCEL_PATH.exists():
        wb = load_workbook(EXCEL_PATH)
    else:
        wb = Workbook()
        wb.active.title = "Diario"
        wb["Diario"].append(CABECERA_DIARIO)
    if "Diario" not in wb.sheetnames:
        wb.create_sheet("Diario").append(CABECERA_DIARIO)
    if "Récords" not in wb.sheetnames:
        wb.create_sheet("Récords").append(CABECERA_RECORDS)
    return wb

def registrar_en_excel(resumen, batidos, hoy):
    wb = abrir_excel()

    ws = wb["Diario"]
    # Evitar fila duplicada si el script se ejecuta dos veces el mismo día
    fecha_str = hoy.strftime("%d/%m/%Y")
    ultima = ws.cell(row=ws.max_row, column=1).value
    if ultima == fecha_str:
        log.info("El día %s ya estaba registrado; se sobrescribe la fila", fecha_str)
        ws.delete_rows(ws.max_row)

    ws.append([
        fecha_str,
        resumen["temp_max"], resumen["hora_temp_max"],
        resumen["temp_min"], resumen["hora_temp_min"],
        resumen["racha_max"], resumen["hora_racha"],
        resumen["lluvia"],
        resumen["presion_max"], resumen["presion_min"],
        resumen["humedad_max"], resumen["humedad_min"],
        resumen["n_obs"],
    ])

    ws_rec = wb["Récords"]
    for b in batidos:
        ant = b["anterior"]
        ant_txt = (
            f"{ant['valor']} {b['unidad']} ({ant['fecha']})" if ant else "— (primer registro)"
        )
        ws_rec.append([
            b["etiqueta"],
            f"{b['valor']} {b['unidad']}",
            fecha_str,
            b["hora"] or "—",
            ant_txt,
            f"TempMax {resumen['temp_max']}°C / TempMin {resumen['temp_min']}°C / "
            f"Racha {resumen['racha_max']} km/h / Lluvia {resumen['lluvia']} mm / "
            f"Presión {resumen['presion_min']}-{resumen['presion_max']} hPa",
        ])

    wb.save(EXCEL_PATH)
    log.info("Excel actualizado: %s", EXCEL_PATH)


# ---------------------------------------------------------------------------
# Telegram
# ---------------------------------------------------------------------------
def _fmt(valor, unidad=""):
    return f"{valor:.1f}{unidad}" if valor is not None else "n/d"


def construir_mensaje(resumen, batidos, hoy):
    lineas = [
        f"🌦 <b>Resumen meteo — {hoy.strftime('%d/%m/%Y')}</b>",
        "",
        f"🌡 Máxima: <b>{_fmt(resumen['temp_max'], '°C')}</b> ({resumen['hora_temp_max'] or '—'})",
        f"❄️ Mínima: <b>{_fmt(resumen['temp_min'], '°C')}</b> ({resumen['hora_temp_min'] or '—'})",
        f"💨 Racha máx: <b>{_fmt(resumen['racha_max'], ' km/h')}</b> ({resumen['hora_racha'] or '—'})",
    ]
    if resumen["lluvia"] and resumen["lluvia"] > 0:
        lineas.append(f"🌧 Lluvia: <b>{resumen['lluvia']:.1f} mm</b>")
    lineas.append(
        f"🔵 Presión: <b>{_fmt(resumen['presion_min'])} – {_fmt(resumen['presion_max'])} hPa</b>"
    )
    if resumen["humedad_max"] is not None:
        lineas.append(
            f"💧 Humedad: {_fmt(resumen['humedad_min'], '%')} – {_fmt(resumen['humedad_max'], '%')}"
        )

    if batidos:
        lineas += ["", "🏆 <b>¡RÉCORD BATIDO!</b>"]
        for b in batidos:
            hora_txt = f" a las {b['hora']}" if b["hora"] else ""
            ant = b["anterior"]
            ant_txt = f" (anterior: {ant['valor']} {b['unidad']}, {ant['fecha']})" if ant else ""
            lineas.append(f"• {b['etiqueta']}: <b>{b['valor']} {b['unidad']}</b>{hora_txt}{ant_txt}")

    return "\n".join(lineas)


def enviar_telegram(texto):
    url = f"https://api.telegram.org/bot{TELEGRAM_TOKEN}/sendMessage"
    payload = {"chat_id": TELEGRAM_CHAT_ID, "text": texto, "parse_mode": "HTML"}
    for intento in range(3):
        try:
            r = requests.post(url, json=payload, timeout=30)
            r.raise_for_status()
            log.info("Mensaje enviado a Telegram")
            return True
        except Exception as e:
            log.warning("Error enviando Telegram (intento %d): %s", intento + 1, e)
    return False


# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------
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

    hoy = date.today()
    log.info("=== Ejecución diaria %s ===", hoy.isoformat())

    observaciones = descargar_observaciones()
    if not observaciones:
        enviar_telegram(
            f"⚠️ <b>Bot Meteo</b>: no se pudieron obtener datos de la estación "
            f"{WU_STATION_ID} para el {hoy.strftime('%d/%m/%Y')}."
        )
        sys.exit(1)

    resumen = calcular_resumen(observaciones)
    batidos = comprobar_records(resumen, hoy)
    registrar_en_excel(resumen, batidos, hoy)
    enviar_telegram(construir_mensaje(resumen, batidos, hoy))
    log.info("=== Fin de la ejecución ===")


if __name__ == "__main__":
    main()
