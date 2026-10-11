#!/usr/bin/env python3
"""
Bot Meteo - Resumen diario de estación meteorológica (Weather Underground)
==========================================================================
Modos de ejecución (systemd timers):
  resumen   21:00  - Mensaje informativo del día + snapshot de lluvia/viento
  registro  23:57  - Registro definitivo del día (Excel, récords, GitHub), silencioso
  mensual   día 1  - Medias del mes anterior en medias_mensuales.xlsx + GitHub

El resumen de las 21:00 incluye la lluvia acumulada del año hidrológico
(1 de septiembre – 31 de agosto), calculada sumando la hoja Diario.

Variables de entorno necesarias (/etc/bot-meteo.env):
  WU_API_KEY        - API key de Weather Underground
  WU_STATION_ID     - ID de la estación (ej: ICARTA12)
  TELEGRAM_TOKEN    - Token del bot de Telegram (BotFather)
  TELEGRAM_CHAT_ID  - Chat ID donde enviar los mensajes
"""

import os
import sys
import csv
import json
import logging
import subprocess
from datetime import datetime, date, timedelta
from pathlib import Path

import requests
from openpyxl import Workbook, load_workbook

# ---------------------------------------------------------------------------
# Configuración
# ---------------------------------------------------------------------------
BASE_DIR = Path(__file__).resolve().parent
EXCEL_PATH = BASE_DIR / "registro_meteo.xlsx"
CSV_DIARIO_PATH = BASE_DIR / "registro_diario.csv"
CSV_RECORDS_PATH = BASE_DIR / "records.csv"
MEDIAS_XLSX_PATH = BASE_DIR / "medias_mensuales.xlsx"
MEDIAS_CSV_PATH = BASE_DIR / "medias_mensuales.csv"
RECORDS_PATH = BASE_DIR / "records.json"
SNAPSHOT_PATH = BASE_DIR / "snapshot_2100.json"
LOG_PATH = BASE_DIR / "bot_meteo.log"

WU_API_KEY = os.environ.get("WU_API_KEY", "")
WU_STATION_ID = os.environ.get("WU_STATION_ID", "")
TELEGRAM_TOKEN = os.environ.get("TELEGRAM_TOKEN", "")
TELEGRAM_CHAT_ID = os.environ.get("TELEGRAM_CHAT_ID", "")

WU_URL = (
    "https://api.weather.com/v2/pws/observations/all/1day"
    "?stationId={station}&format=json&units=m&numericPrecision=decimal&apiKey={key}"
)

# Año hidrológico: 1 de septiembre – 31 de agosto
MES_INICIO_HIDRO = 9

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
    "Humedad Máx (%)", "Humedad Mín (%)",
    "Radiación Máx (W/m²)", "Hora Radiación", "UV Máx",
    "Nº Observaciones",
]
CABECERA_RECORDS = ["Tipo de récord", "Valor", "Fecha", "Hora", "Récord anterior", "Detalles"]
CABECERA_MENSUAL = [
    "Mes", "Días con datos",
    "Media Máximas (°C)", "Media Mínimas (°C)", "Temp Media (°C)",
    "Máx absoluta (°C)", "Fecha Máx", "Mín absoluta (°C)", "Fecha Mín",
    "Media Rachas (km/h)", "Racha Máx (km/h)", "Fecha Racha",
    "Lluvia Total (mm)", "Días de lluvia",
    "Media Presión (hPa)",
    "Media Humedad Máx (%)", "Media Humedad Mín (%)",
    "Media Radiación Máx (W/m²)", "Media UV Máx",
]


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
        "solar_max": None, "hora_solar": "",
        "uv_max": None,
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

        # Radiación solar y UV (van a nivel de observación, no en "metric")
        sol = _flt(obs.get("solarRadiationHigh"))
        if sol is not None and (resumen["solar_max"] is None or sol > resumen["solar_max"]):
            resumen["solar_max"] = sol
            resumen["hora_solar"] = _hora(obs)

        uv = _flt(obs.get("uvHigh"))
        if uv is not None and (resumen["uv_max"] is None or uv > resumen["uv_max"]):
            resumen["uv_max"] = uv

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
    ("solar_max", "solar_max", "Radiación solar", "W/m²", True, "hora_solar"),
    ("uv_max", "uv_max", "Índice UV", "", True, None),
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
# Excel diario
# ---------------------------------------------------------------------------
def _sincronizar_cabecera(hoja, cabecera):
    """Asegura que la fila 1 de la hoja coincide con la cabecera actual
    (necesario al añadir columnas nuevas como radiación/UV)."""
    for col, titulo in enumerate(cabecera, start=1):
        hoja.cell(row=1, column=col, value=titulo)


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
    _sincronizar_cabecera(wb["Diario"], CABECERA_DIARIO)
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
        resumen["solar_max"], resumen["hora_solar"], resumen["uv_max"],
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
            f"{b['valor']} {b['unidad']}".strip(),
            fecha_str,
            b["hora"] or "—",
            ant_txt,
            f"TempMax {resumen['temp_max']}°C / TempMin {resumen['temp_min']}°C / "
            f"Racha {resumen['racha_max']} km/h / Lluvia {resumen['lluvia']} mm / "
            f"Presión {resumen['presion_min']}-{resumen['presion_max']} hPa",
        ])

    wb.save(EXCEL_PATH)
    log.info("Excel actualizado: %s", EXCEL_PATH)

    # CSVs espejo (GitHub los renderiza como tabla en el navegador)
    exportar_csv(wb["Diario"], CSV_DIARIO_PATH)
    exportar_csv(wb["Récords"], CSV_RECORDS_PATH)


def exportar_csv(hoja, destino):
    try:
        with open(destino, "w", newline="", encoding="utf-8-sig") as f:
            writer = csv.writer(f)
            for fila in hoja.iter_rows(values_only=True):
                writer.writerow(fila)
        log.info("CSV exportado: %s", destino)
    except Exception as e:
        log.warning("No se pudo exportar %s: %s", destino, e)


# ---------------------------------------------------------------------------
# Medias mensuales (modo "mensual", se ejecuta el día 1 de cada mes)
# ---------------------------------------------------------------------------
def _media(valores):
    valores = [v for v in valores if v is not None]
    return round(sum(valores) / len(valores), 1) if valores else None


def leer_dias_del_mes(anio, mes):
    """Lee la hoja Diario y devuelve los días del mes indicado como dicts."""
    if not EXCEL_PATH.exists():
        return []
    wb = load_workbook(EXCEL_PATH, read_only=True, data_only=True)
    if "Diario" not in wb.sheetnames:
        return []
    ws = wb["Diario"]

    filas = ws.iter_rows(values_only=True)
    cabecera = next(filas, None)
    if not cabecera:
        return []
    idx = {str(nombre): i for i, nombre in enumerate(cabecera) if nombre}

    def celda(fila, nombre):
        i = idx.get(nombre)
        if i is None or i >= len(fila):
            return None
        return fila[i]

    dias = []
    for fila in filas:
        fecha_raw = celda(fila, "Fecha")
        if not fecha_raw:
            continue
        try:
            f = datetime.strptime(str(fecha_raw).strip(), "%d/%m/%Y").date()
        except ValueError:
            continue
        if f.year != anio or f.month != mes:
            continue
        dias.append({
            "fecha": f,
            "temp_max": _flt(celda(fila, "Temp Máx (°C)")),
            "temp_min": _flt(celda(fila, "Temp Mín (°C)")),
            "racha": _flt(celda(fila, "Racha Máx (km/h)")),
            "lluvia": _flt(celda(fila, "Lluvia (mm)")),
            "presion_max": _flt(celda(fila, "Presión Máx (hPa)")),
            "presion_min": _flt(celda(fila, "Presión Mín (hPa)")),
            "hum_max": _flt(celda(fila, "Humedad Máx (%)")),
            "hum_min": _flt(celda(fila, "Humedad Mín (%)")),
            "solar": _flt(celda(fila, "Radiación Máx (W/m²)")),
            "uv": _flt(celda(fila, "UV Máx")),
        })
    return dias


# ---------------------------------------------------------------------------
# Año hidrológico (1 sep – 31 ago)
# ---------------------------------------------------------------------------
def anio_hidrologico(d):
    """Devuelve (fecha_inicio, etiqueta) del año hidrológico al que pertenece d."""
    anio_ini = d.year if d.month >= MES_INICIO_HIDRO else d.year - 1
    return date(anio_ini, MES_INICIO_HIDRO, 1), f"{anio_ini}-{str(anio_ini + 1)[-2:]}"


def lluvia_acumulada_hidro(hoy, incluir_hoy=True):
    """Suma la columna 'Lluvia (mm)' de la hoja Diario desde el 1 de septiembre
    del año hidrológico en curso hasta hoy. Devuelve float, o None si falla.
    Se calcula desde el Excel, así que re-ejecutar un día no duplica la suma."""
    inicio, _ = anio_hidrologico(hoy)
    if not EXCEL_PATH.exists():
        return 0.0
    total = 0.0
    try:
        wb = load_workbook(EXCEL_PATH, read_only=True, data_only=True)
        if "Diario" not in wb.sheetnames:
            wb.close()
            return 0.0
        filas = wb["Diario"].iter_rows(values_only=True)
        cabecera = next(filas, None) or ()
        idx = {str(nombre): i for i, nombre in enumerate(cabecera) if nombre}
        i_fecha, i_lluvia = idx.get("Fecha"), idx.get("Lluvia (mm)")
        if i_fecha is None or i_lluvia is None:
            wb.close()
            log.warning("Hoja Diario sin columnas Fecha/Lluvia; no se calcula el año hidrológico")
            return None
        for fila in filas:
            if not fila or len(fila) <= max(i_fecha, i_lluvia) or not fila[i_fecha]:
                continue
            try:
                f = datetime.strptime(str(fila[i_fecha]).strip(), "%d/%m/%Y").date()
            except ValueError:
                continue
            if f < inicio or f > hoy or (f == hoy and not incluir_hoy):
                continue
            total += _flt(fila[i_lluvia]) or 0.0
        wb.close()
    except Exception as e:
        log.warning("No se pudo calcular la lluvia del año hidrológico: %s", e)
        return None
    return total


def linea_hidro(total, hoy):
    if total is None:
        return None
    _, etiqueta = anio_hidrologico(hoy)
    return f"🪣 Año hidrológico {etiqueta}: <b>{total:.1f} L/m²</b>"


def calcular_medias_mes(dias):
    """A partir de la lista de días, calcula el resumen mensual."""
    def maximo(clave):
        candidatos = [(d[clave], d["fecha"]) for d in dias if d[clave] is not None]
        return max(candidatos) if candidatos else (None, None)

    def minimo(clave):
        candidatos = [(d[clave], d["fecha"]) for d in dias if d[clave] is not None]
        return min(candidatos) if candidatos else (None, None)

    temp_medias = [
        (d["temp_max"] + d["temp_min"]) / 2
        for d in dias if d["temp_max"] is not None and d["temp_min"] is not None
    ]
    presion_medias = [
        (d["presion_max"] + d["presion_min"]) / 2
        for d in dias if d["presion_max"] is not None and d["presion_min"] is not None
    ]

    max_abs, fecha_max = maximo("temp_max")
    min_abs, fecha_min = minimo("temp_min")
    racha_abs, fecha_racha = maximo("racha")
    lluvias = [d["lluvia"] for d in dias if d["lluvia"] is not None]

    return {
        "dias": len(dias),
        "media_max": _media([d["temp_max"] for d in dias]),
        "media_min": _media([d["temp_min"] for d in dias]),
        "temp_media": _media(temp_medias),
        "max_abs": max_abs, "fecha_max": fecha_max,
        "min_abs": min_abs, "fecha_min": fecha_min,
        "media_racha": _media([d["racha"] for d in dias]),
        "racha_abs": racha_abs, "fecha_racha": fecha_racha,
        "lluvia_total": round(sum(lluvias), 1) if lluvias else 0.0,
        "dias_lluvia": sum(1 for v in lluvias if v > 0),
        "media_presion": _media(presion_medias),
        "media_hum_max": _media([d["hum_max"] for d in dias]),
        "media_hum_min": _media([d["hum_min"] for d in dias]),
        "media_solar": _media([d["solar"] for d in dias]),
        "media_uv": _media([d["uv"] for d in dias]),
    }


def registrar_medias(stats, anio, mes):
    """Escribe/actualiza la fila del mes en medias_mensuales.xlsx y su CSV."""
    if MEDIAS_XLSX_PATH.exists():
        wb = load_workbook(MEDIAS_XLSX_PATH)
    else:
        wb = Workbook()
        wb.active.title = "Mensual"
        wb["Mensual"].append(CABECERA_MENSUAL)
    if "Mensual" not in wb.sheetnames:
        wb.create_sheet("Mensual").append(CABECERA_MENSUAL)
    ws = wb["Mensual"]
    _sincronizar_cabecera(ws, CABECERA_MENSUAL)

    mes_str = f"{mes:02d}/{anio}"
    # Si el mes ya existe (re-ejecución), se sobrescribe su fila
    for fila in range(ws.max_row, 1, -1):
        if ws.cell(row=fila, column=1).value == mes_str:
            log.info("El mes %s ya estaba registrado; se sobrescribe", mes_str)
            ws.delete_rows(fila)

    def fch(f):
        return f.strftime("%d/%m/%Y") if f else "—"

    ws.append([
        mes_str, stats["dias"],
        stats["media_max"], stats["media_min"], stats["temp_media"],
        stats["max_abs"], fch(stats["fecha_max"]),
        stats["min_abs"], fch(stats["fecha_min"]),
        stats["media_racha"], stats["racha_abs"], fch(stats["fecha_racha"]),
        stats["lluvia_total"], stats["dias_lluvia"],
        stats["media_presion"],
        stats["media_hum_max"], stats["media_hum_min"],
        stats["media_solar"], stats["media_uv"],
    ])

    wb.save(MEDIAS_XLSX_PATH)
    log.info("Excel mensual actualizado: %s", MEDIAS_XLSX_PATH)
    exportar_csv(ws, MEDIAS_CSV_PATH)


NOMBRES_MES = [
    "", "enero", "febrero", "marzo", "abril", "mayo", "junio",
    "julio", "agosto", "septiembre", "octubre", "noviembre", "diciembre",
]


def construir_mensaje_mensual(stats, anio, mes):
    nombre = NOMBRES_MES[mes].capitalize()
    lineas = [
        f"📊 <b>Resumen mensual — {nombre} {anio}</b>",
        f"({stats['dias']} días con datos)",
        "",
        f"🌡 Media de máximas: <b>{_fmt(stats['media_max'], '°C')}</b>",
        f"❄️ Media de mínimas: <b>{_fmt(stats['media_min'], '°C')}</b>",
        f"🌡 Temperatura media: <b>{_fmt(stats['temp_media'], '°C')}</b>",
        f"🔥 Máxima absoluta: <b>{_fmt(stats['max_abs'], '°C')}</b>"
        + (f" ({stats['fecha_max'].strftime('%d/%m')})" if stats["fecha_max"] else ""),
        f"🧊 Mínima absoluta: <b>{_fmt(stats['min_abs'], '°C')}</b>"
        + (f" ({stats['fecha_min'].strftime('%d/%m')})" if stats["fecha_min"] else ""),
        f"💨 Media de rachas: <b>{_fmt(stats['media_racha'], ' km/h')}</b>",
        f"🌪 Racha máxima: <b>{_fmt(stats['racha_abs'], ' km/h')}</b>"
        + (f" ({stats['fecha_racha'].strftime('%d/%m')})" if stats["fecha_racha"] else ""),
        f"🌧 Lluvia total: <b>{stats['lluvia_total']:.1f} mm</b> "
        f"({stats['dias_lluvia']} días de lluvia)",
        f"🔵 Presión media: <b>{_fmt(stats['media_presion'], ' hPa')}</b>",
        f"💧 Humedad media: {_fmt(stats['media_hum_min'], '%')} – {_fmt(stats['media_hum_max'], '%')}",
    ]
    if stats["media_solar"] is not None:
        lineas.append(f"☀️ Media radiación máx: <b>{_fmt(stats['media_solar'], ' W/m²')}</b>")
    if stats["media_uv"] is not None:
        lineas.append(f"🕶 Media UV máx: <b>{_fmt(stats['media_uv'])}</b>")
    return "\n".join(lineas)


def ejecutar_mensual(hoy):
    """Calcula y registra las medias del mes anterior."""
    ultimo_dia_mes_anterior = hoy.replace(day=1) - timedelta(days=1)
    anio, mes = ultimo_dia_mes_anterior.year, ultimo_dia_mes_anterior.month
    log.info("Calculando medias de %02d/%d", mes, anio)

    dias = leer_dias_del_mes(anio, mes)
    if not dias:
        log.warning("No hay días registrados para %02d/%d", mes, anio)
        enviar_telegram(
            f"⚠️ <b>Bot Meteo</b>: no hay datos registrados de "
            f"{NOMBRES_MES[mes]} {anio}, no se pueden calcular medias.",
        )
        return

    stats = calcular_medias_mes(dias)
    registrar_medias(stats, anio, mes)
    enviar_telegram(construir_mensaje_mensual(stats, anio, mes))
    subir_a_github(hoy, f"Medias mensuales {mes:02d}/{anio}")


# ---------------------------------------------------------------------------
# Sincronización con GitHub
# ---------------------------------------------------------------------------
def subir_a_github(hoy, mensaje=None):
    """Commit y push de los ficheros de datos. Si falla, solo lo registra."""
    archivos = [
        "registro_meteo.xlsx", "registro_diario.csv", "records.csv", "records.json",
        "medias_mensuales.xlsx", "medias_mensuales.csv",
    ]
    existentes = [a for a in archivos if (BASE_DIR / a).exists()]
    mensaje = mensaje or f"Registro diario {hoy.strftime('%d/%m/%Y')}"
    base = ["git", "-C", str(BASE_DIR),
            "-c", "user.name=bot-meteo", "-c", "user.email=bot-meteo@vm"]
    try:
        subprocess.run(base + ["add"] + existentes, check=True, capture_output=True, timeout=60)
        commit = subprocess.run(
            base + ["commit", "-m", mensaje],
            capture_output=True, text=True, timeout=60,
        )
        if commit.returncode != 0:
            if "nothing to commit" in (commit.stdout + commit.stderr):
                log.info("Sin cambios que subir a GitHub")
                return
            raise RuntimeError(commit.stderr.strip() or commit.stdout.strip())
        subprocess.run(base + ["push"], check=True, capture_output=True, timeout=120)
        log.info("Datos subidos a GitHub")
    except Exception as e:
        log.warning("No se pudo subir a GitHub (el bot sigue funcionando): %s", e)


# ---------------------------------------------------------------------------
# Snapshot de las 21:00 y novedades nocturnas
# ---------------------------------------------------------------------------
def guardar_snapshot(resumen, hoy):
    """Guarda la foto de lluvia/viento en el momento del resumen de las 21:00."""
    datos = {
        "fecha": hoy.isoformat(),
        "lluvia": resumen["lluvia"],
        "racha_max": resumen["racha_max"],
    }
    try:
        SNAPSHOT_PATH.write_text(
            json.dumps(datos, ensure_ascii=False), encoding="utf-8"
        )
        log.info("Snapshot 21:00 guardado: %s", datos)
    except Exception as e:
        log.warning("No se pudo guardar el snapshot: %s", e)


def novedades_desde_snapshot(resumen, hoy, total_hidro=None):
    """Compara el día completo con la foto de las 21:00.
    Devuelve lista de líneas si llovió más o el viento superó la racha."""
    if not SNAPSHOT_PATH.exists():
        return []
    try:
        snap = json.loads(SNAPSHOT_PATH.read_text(encoding="utf-8"))
    except Exception as e:
        log.warning("snapshot_2100.json ilegible: %s", e)
        return []
    if snap.get("fecha") != hoy.isoformat():
        log.info("Snapshot de otro día (%s), se ignora", snap.get("fecha"))
        return []

    lineas = []

    lluvia_2100 = snap.get("lluvia") or 0.0
    lluvia_final = resumen["lluvia"] or 0.0
    if lluvia_final > lluvia_2100 + 0.01:
        delta = lluvia_final - lluvia_2100
        lineas.append(
            f"🌧 Lluvia: <b>{lluvia_final:.1f} mm</b> en el día "
            f"(+{delta:.1f} mm desde las 21:00)"
        )
        # El acumulado anual dado a las 21:00 ha cambiado: se actualiza
        linea = linea_hidro(total_hidro, hoy)
        if linea:
            lineas.append(linea)

    racha_2100 = snap.get("racha_max")
    racha_final = resumen["racha_max"]
    if racha_final is not None and (racha_2100 is None or racha_final > racha_2100 + 0.01):
        hora_txt = f" a las {resumen['hora_racha']}" if resumen["hora_racha"] else ""
        ant_txt = f" (a las 21:00 era {racha_2100:.1f} km/h)" if racha_2100 is not None else ""
        lineas.append(
            f"💨 Racha máx: <b>{racha_final:.1f} km/h</b>{hora_txt}{ant_txt}"
        )

    return lineas


def _fmt(valor, unidad=""):
    return f"{valor:.1f}{unidad}" if valor is not None else "n/d"


def construir_mensaje(resumen, batidos, hoy, total_hidro=None):
    lineas = [
        f"🌦 <b>Resumen meteo — {hoy.strftime('%d/%m/%Y')}</b>",
        "",
        f"🌡 Máxima: <b>{_fmt(resumen['temp_max'], '°C')}</b> ({resumen['hora_temp_max'] or '—'})",
        f"❄️ Mínima: <b>{_fmt(resumen['temp_min'], '°C')}</b> ({resumen['hora_temp_min'] or '—'})",
        f"💨 Racha máx: <b>{_fmt(resumen['racha_max'], ' km/h')}</b> ({resumen['hora_racha'] or '—'})",
    ]
    if resumen["lluvia"] and resumen["lluvia"] > 0:
        lineas.append(f"🌧 Lluvia: <b>{resumen['lluvia']:.1f} mm</b>")
    linea = linea_hidro(total_hidro, hoy)
    if linea:
        lineas.append(linea)
    lineas.append(
        f"🔵 Presión: <b>{_fmt(resumen['presion_min'])} – {_fmt(resumen['presion_max'])} hPa</b>"
    )
    if resumen["humedad_max"] is not None:
        lineas.append(
            f"💧 Humedad: {_fmt(resumen['humedad_min'], '%')} – {_fmt(resumen['humedad_max'], '%')}"
        )
    if resumen["solar_max"] is not None:
        hora_txt = f" ({resumen['hora_solar']})" if resumen["hora_solar"] else ""
        lineas.append(f"☀️ Radiación máx: <b>{_fmt(resumen['solar_max'], ' W/m²')}</b>{hora_txt}")
    if resumen["uv_max"] is not None:
        lineas.append(f"🕶 UV máx: <b>{_fmt(resumen['uv_max'])}</b>")

    if batidos:
        lineas += ["", "🏆 <b>¡RÉCORD BATIDO!</b>"]
        for b in batidos:
            hora_txt = f" a las {b['hora']}" if b["hora"] else ""
            ant = b["anterior"]
            ant_txt = f" (anterior: {ant['valor']} {b['unidad']}, {ant['fecha']})" if ant else ""
            lineas.append(f"• {b['etiqueta']}: <b>{b['valor']} {b['unidad']}</b>{hora_txt}{ant_txt}")

    return "\n".join(lineas)


def enviar_telegram(texto, silencioso=False):
    url = f"https://api.telegram.org/bot{TELEGRAM_TOKEN}/sendMessage"
    payload = {
        "chat_id": TELEGRAM_CHAT_ID,
        "text": texto,
        "parse_mode": "HTML",
        "disable_notification": silencioso,
    }
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

    # Modos: "resumen" (21:00), "registro" (23:57, silencioso), "mensual" (día 1)
    modo = sys.argv[1] if len(sys.argv) > 1 else "resumen"
    if modo not in ("resumen", "registro", "mensual"):
        log.error("Modo desconocido: %s (usa 'resumen', 'registro' o 'mensual')", modo)
        sys.exit(1)

    hoy = date.today()
    log.info("=== Ejecución diaria %s (modo: %s) ===", hoy.isoformat(), modo)

    if modo == "mensual":
        # No necesita descargar nada: trabaja sobre el registro ya guardado
        ejecutar_mensual(hoy)
        log.info("=== Fin de la ejecución (modo: mensual) ===")
        return

    observaciones = descargar_observaciones()
    if not observaciones:
        enviar_telegram(
            f"⚠️ <b>Bot Meteo</b>: no se pudieron obtener datos de la estación "
            f"{WU_STATION_ID} para el {hoy.strftime('%d/%m/%Y')}.",
            silencioso=(modo == "registro"),
        )
        sys.exit(1)

    resumen = calcular_resumen(observaciones)

    if modo == "resumen":
        # 21:00 — mensaje informativo + foto de lluvia/viento para comparar luego
        # Acumulado hidrológico: días ya registrados + lo que lleva hoy
        total_hidro = lluvia_acumulada_hidro(hoy, incluir_hoy=False)
        if total_hidro is not None:
            total_hidro += resumen["lluvia"] or 0.0
        enviar_telegram(construir_mensaje(resumen, [], hoy, total_hidro))
        guardar_snapshot(resumen, hoy)
    else:
        # 23:57 — registro definitivo del día completo, sin despertar a nadie
        batidos = comprobar_records(resumen, hoy)
        registrar_en_excel(resumen, batidos, hoy)
        total_hidro = lluvia_acumulada_hidro(hoy, incluir_hoy=True)
        log.info("Lluvia año hidrológico: %s mm", total_hidro)

        bloques = []

        novedades = novedades_desde_snapshot(resumen, hoy, total_hidro)
        if novedades:
            bloques.append(
                "\n".join([f"🌙 <b>Novedades desde las 21:00 — {hoy.strftime('%d/%m/%Y')}</b>"]
                          + novedades)
            )

        if batidos:
            lineas = [f"🏆 <b>Récord batido el {hoy.strftime('%d/%m/%Y')}</b>"]
            for b in batidos:
                hora_txt = f" a las {b['hora']}" if b["hora"] else ""
                ant = b["anterior"]
                ant_txt = (
                    f" (anterior: {ant['valor']} {b['unidad']}, {ant['fecha']})" if ant else ""
                )
                lineas.append(
                    f"• {b['etiqueta']}: <b>{b['valor']} {b['unidad']}</b>{hora_txt}{ant_txt}"
                )
            bloques.append("\n".join(lineas))

        if bloques:
            enviar_telegram("\n\n".join(bloques), silencioso=True)

        subir_a_github(hoy)

    log.info("=== Fin de la ejecución (modo: %s) ===", modo)


if __name__ == "__main__":
    main()
