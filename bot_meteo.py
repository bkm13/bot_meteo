#!/usr/bin/env python3
"""
Bot Meteo - Resumen diario de estación meteorológica (Weather Underground)
==========================================================================
Modos de ejecución (systemd timers):
  resumen   21:00  - Mensaje informativo del día + snapshot de lluvia/viento
  registro  23:57  - Registro definitivo del día (Excel, récords, GitHub), silencioso
  mensual   día 1  - Medias del mes anterior en medias_mensuales.xlsx + GitHub

El resumen de las 21:00 incluye la lluvia acumulada del mes en curso y la del
año hidrológico (1 de septiembre – 31 de agosto), calculadas sumando la hoja
Diario más lo que lleva el día. El resumen mensual incluye además el desglose
mes a mes del año hidrológico. resumen_mensual() se reutiliza desde el
listener para consultar cualquier mes bajo demanda (/mensual, /mes).

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
import calendar
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
    # Columnas nuevas SIEMPRE al final: _sincronizar_cabecera reescribe la
    # fila 1 por posición y una inserción en medio desplaza las filas antiguas
    "Inicio Lluvia", "Fin Lluvia", "Intensidad Máx (mm/h)", "Hora Intensidad",
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
# Variables globales para control de alertas y evitar spam
_alerta_enviada = False
def enviar_alerta_telegram(mensaje):
    """Envía un aviso proactivo a tu Telegram si la Pi muere."""
    token = os.environ.get("TELEGRAM_BOT_TOKEN", "TU_TOKEN")
    chat_id = os.environ.get("TELEGRAM_CHAT_ID", "TU_CHAT_ID")
    url = f"https://api.telegram.org/bot{token}/sendMessage"
    try:
        requests.post(url, json={"chat_id": chat_id, "text": mensaje, "parse_mode": "Markdown"}, timeout=10)
    except Exception as e:
        print(f"Error enviando alerta: {e}")

def verificar_salud_y_descargar():
    """Descarga de GitHub, vigila la salud de la Pi y activa respaldos si falla."""
    global _alerta_enviada
    url_github = "https://raw.githubusercontent.com/bkm13/bot_meteo/main/observaciones_hoy.json"
    
    # 1. Intentar leer los datos procesados por la Raspberry Pi vía GitHub
    try:
        r = requests.get(url_github, timeout=10)
        if r.status_code == 200:
            obs = r.json()
            if obs:
                # Comprobar la antigüedad de la última lectura para ver si la Pi sigue viva
                ultima_obs = obs[-1]
                fecha_str = ultima_obs.get("date") # Ajusta al campo de fecha real de tu JSON
                
                if fecha_str:
                    tiempo_lectura = datetime.strptime(fecha_str, "%Y-%m-%d %H:%M:%S").replace(tzinfo=timezone.utc)
                    if datetime.now(timezone.utc) - tiempo_lectura > timedelta(hours=2):
                        if not _alerta_enviada:
                            enviar_alerta_telegram(
                                "🚨 **¡ALERTA DE SISTEMA!**\n\n"
                                f"La Raspberry Pi lleva más de 2 horas sin actualizar los datos.\n"
                                f"Última lectura: `{fecha_str}`.\n"
                                "⚠️ *Posible fallo de hardware, tarjeta SD o red.*"
                            )
                            _alerta_enviada = True
                    else:
                        if _alerta_enviada:
                            enviar_alerta_telegram("✅ **¡Sistema recuperado!** La Raspberry Pi ha vuelto a enviar datos.")
                            _alerta_enviada = False
                            
                return obs
    except Exception as e:
        print(f"No se pudo conectar con GitHub: {e}")

    # 2. Plan B: Si la Pi ha muerto o GitHub falla, usar Weather Underground directamente
    print("⚠️ Pi inaccesible. Usando respaldo de Weather Underground...")
    url_wu = WU_URL.format(station=IMONTE1305, key=ElJt9IYx)
    try:
        r = requests.get(url_wu, timeout=15)
        r.raise_for_status()
        return r.json().get("observations", [])
    except Exception as e:
        print(f"Fallo crítico en el respaldo de WU: {e}")
        
    return None

# ---------------------------------------------------------------------------
# Descarga y cálculo
# ---------------------------------------------------------------------------
def descargar_observaciones():
    """Descarga las observaciones intentando primero la Pi vía GitHub.
       Si la Pi está rota o no actualiza, recurre automáticamente a Weather Underground."""
    
    url_github = "https://raw.githubusercontent.com/bkm13/bot_meteo/main/observaciones_hoy.json"
    
    # 1. Intentar leer los datos locales procesados por la Raspberry Pi
    try:
        r = requests.get(url_github, timeout=10)
        if r.status_code == 200:
            obs = r.json()
            if obs:
                # Opcional: Validar que el archivo no esté obsoleto (ej. última lectura reciente)
                log.info("Datos obtenidos correctamente desde la Raspberry Pi (GitHub)")
                return obs
    except Exception as e:
        log.warning("La Raspberry Pi / GitHub no responde o falló: %s", e)

    # 2. Plan B: Si la Pi se ha roto o no hay conexión, tirar de la API oficial directamente
    log.warning("⚠️ Alerta: Raspberry Pi inaccesible. Activando respaldo externo de Weather Underground...")
    url_wu = WU_URL.format(station=WU_STATION_ID, key=WU_API_KEY)
    
    try:
        r = requests.get(url_wu, timeout=15)
        r.raise_for_status()
        datos_wu = r.json().get("observations", [])
        if datos_wu:
            log.info("Datos recuperados con éxito mediante el respaldo de Weather Underground.")
            return datos_wu
    except Exception as e:
        log.error("Fallo crítico: Tampoco se pudo conectar con Weather Underground: %s", e)
        
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
        "lluvia_ini": "", "lluvia_fin": "",
        "intensidad_max": None, "hora_intensidad": "",
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

    resumen.update(calcular_lluvia_detalle(observaciones))
    return resumen


def calcular_lluvia_detalle(observaciones):
    """Hora de inicio/fin de la lluvia del día e intensidad máxima (mm/h).
    Inicio/fin: primera y última observación en que sube precipTotal.
    Intensidad: máximo de precipRate; si la estación no lo da, se estima
    con la subida de precipTotal entre observaciones consecutivas."""
    ordenadas = sorted(observaciones, key=lambda o: o.get("obsTimeLocal", ""))
    ini = fin = ""
    int_max, hora_int = None, ""
    est_max, hora_est = None, ""
    acumulado = 0.0
    t_prev = None

    for obs in ordenadas:
        m = obs.get("metric", {}) or {}
        hora = _hora(obs)
        try:
            t = datetime.strptime(obs.get("obsTimeLocal", "")[:19], "%Y-%m-%d %H:%M:%S")
        except ValueError:
            t = None

        total = _flt(m.get("precipTotal"))
        if total is not None and total > acumulado + 0.001:
            delta = total - acumulado
            if not ini:
                ini = hora
            fin = hora
            if t is not None and t_prev is not None:
                minutos = (t - t_prev).total_seconds() / 60
                if 0 < minutos <= 30:
                    estimada = delta / minutos * 60
                    if est_max is None or estimada > est_max:
                        est_max, hora_est = estimada, hora
            acumulado = total
        if t is not None:
            t_prev = t

        ritmo = _flt(m.get("precipRate"))
        if ritmo is not None and ritmo > 0 and (int_max is None or ritmo > int_max):
            int_max, hora_int = ritmo, hora

    if int_max is None and est_max is not None:
        int_max, hora_int = round(est_max, 1), hora_est
    if not ini:  # sin lluvia: no se registra intensidad
        int_max, hora_int = None, ""
    return {
        "lluvia_ini": ini, "lluvia_fin": fin,
        "intensidad_max": int_max, "hora_intensidad": hora_int,
    }


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
    _reparar_filas_sin_radiacion(wb["Diario"])
    return wb


def _reparar_filas_sin_radiacion(ws):
    """Corrige las filas anteriores a la columna de radiación/UV (agosto 2026):
    al insertar esas columnas, su 'Nº Observaciones' quedó en 'Radiación Máx'.
    Solo toca filas con radiación pero sin hora de radiación, UV ni Nº obs."""
    col = {str(c.value): c.column for c in ws[1] if c.value}
    c_rad, c_hora = col.get("Radiación Máx (W/m²)"), col.get("Hora Radiación")
    c_uv, c_obs = col.get("UV Máx"), col.get("Nº Observaciones")
    if None in (c_rad, c_hora, c_uv, c_obs):
        return
    for fila in range(2, ws.max_row + 1):
        rad = ws.cell(row=fila, column=c_rad).value
        if (rad is not None and ws.cell(row=fila, column=c_hora).value in (None, "")
                and ws.cell(row=fila, column=c_uv).value is None
                and ws.cell(row=fila, column=c_obs).value is None):
            ws.cell(row=fila, column=c_obs, value=rad)
            ws.cell(row=fila, column=c_rad).value = None  # cell(value=None) no borra
            log.info("Fila %s reparada: Nº observaciones movido desde radiación",
                     ws.cell(row=fila, column=1).value)


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
        resumen["lluvia_ini"], resumen["lluvia_fin"],
        resumen["intensidad_max"], resumen["hora_intensidad"],
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


NOMBRES_MES = [
    "", "enero", "febrero", "marzo", "abril", "mayo", "junio",
    "julio", "agosto", "septiembre", "octubre", "noviembre", "diciembre",
]

# Columna de la hoja Diario -> clave interna
CAMPOS_DIARIO = {
    "temp_max": "Temp Máx (°C)",
    "temp_min": "Temp Mín (°C)",
    "racha": "Racha Máx (km/h)",
    "lluvia": "Lluvia (mm)",
    "presion_max": "Presión Máx (hPa)",
    "presion_min": "Presión Mín (hPa)",
    "hum_max": "Humedad Máx (%)",
    "hum_min": "Humedad Mín (%)",
    "solar": "Radiación Máx (W/m²)",
    "uv": "UV Máx",
    "intensidad": "Intensidad Máx (mm/h)",
}
CAMPOS_TEXTO_DIARIO = {
    "lluvia_ini": "Inicio Lluvia",
    "lluvia_fin": "Fin Lluvia",
    "hora_intensidad": "Hora Intensidad",
}


def leer_diario():
    """Lee la hoja Diario completa y devuelve una lista de dicts por día
    (clave 'fecha' como date + las claves de CAMPOS_DIARIO)."""
    if not EXCEL_PATH.exists():
        return []
    wb = load_workbook(EXCEL_PATH, read_only=True, data_only=True)
    try:
        if "Diario" not in wb.sheetnames:
            return []
        filas = wb["Diario"].iter_rows(values_only=True)
        cabecera = next(filas, None) or ()
        idx = {str(nombre): i for i, nombre in enumerate(cabecera) if nombre}
        i_fecha = idx.get("Fecha")
        if i_fecha is None:
            return []
        dias = []
        for fila in filas:
            if not fila or len(fila) <= i_fecha or not fila[i_fecha]:
                continue
            try:
                f = datetime.strptime(str(fila[i_fecha]).strip(), "%d/%m/%Y").date()
            except ValueError:
                continue
            dia = {"fecha": f}
            for clave, columna in CAMPOS_DIARIO.items():
                i = idx.get(columna)
                dia[clave] = _flt(fila[i]) if i is not None and i < len(fila) else None
            for clave, columna in CAMPOS_TEXTO_DIARIO.items():
                i = idx.get(columna)
                valor = fila[i] if i is not None and i < len(fila) else None
                dia[clave] = str(valor).strip() if valor not in (None, "") else ""
            dias.append(dia)
        return dias
    finally:
        wb.close()


def leer_dias_del_mes(anio, mes, dias=None):
    """Días del mes indicado (lee la hoja Diario si no se le pasan)."""
    dias = leer_diario() if dias is None else dias
    return [d for d in dias if d["fecha"].year == anio and d["fecha"].month == mes]


def meses_con_datos(dias=None):
    """{año: [meses con al menos un día registrado]} ordenado."""
    dias = leer_diario() if dias is None else dias
    resultado = {}
    for d in dias:
        resultado.setdefault(d["fecha"].year, set()).add(d["fecha"].month)
    return {a: sorted(m) for a, m in sorted(resultado.items())}


# ---------------------------------------------------------------------------
# Año hidrológico (1 sep – 31 ago) y lluvia acumulada
# ---------------------------------------------------------------------------
def anio_hidrologico(d):
    """Devuelve (fecha_inicio, etiqueta) del año hidrológico al que pertenece d."""
    anio_ini = d.year if d.month >= MES_INICIO_HIDRO else d.year - 1
    return date(anio_ini, MES_INICIO_HIDRO, 1), f"{anio_ini}-{str(anio_ini + 1)[-2:]}"


def _suma_lluvia(dias, desde, hasta, incluir_hasta=True):
    return sum(
        d["lluvia"] or 0.0 for d in dias
        if desde <= d["fecha"] <= hasta and (incluir_hasta or d["fecha"] != hasta)
    )


def acumulados_lluvia(hoy, incluir_hoy=True):
    """Lluvia acumulada (mes en curso, año hidrológico) sumando la hoja Diario.
    Devuelve (total_mes, total_hidro), o (None, None) si falla la lectura.
    Se calcula desde el Excel, así que re-ejecutar un día no duplica la suma."""
    try:
        dias = leer_diario()
    except Exception as e:
        log.warning("No se pudo leer la hoja Diario para la lluvia acumulada: %s", e)
        return None, None
    inicio_hidro, _ = anio_hidrologico(hoy)
    total_mes = _suma_lluvia(dias, hoy.replace(day=1), hoy, incluir_hoy)
    total_hidro = _suma_lluvia(dias, inicio_hidro, hoy, incluir_hoy)
    return total_mes, total_hidro


def lluvia_acumulada_hidro(hoy, incluir_hoy=True):
    """Compatibilidad: solo el acumulado del año hidrológico."""
    return acumulados_lluvia(hoy, incluir_hoy)[1]


def linea_mes(total, hoy):
    if total is None:
        return None
    return f"🗓 Lluvia en {NOMBRES_MES[hoy.month]}: <b>{total:.1f} L/m²</b>"


def linea_hidro(total, hoy):
    if total is None:
        return None
    _, etiqueta = anio_hidrologico(hoy)
    return f"🪣 Año hidrológico {etiqueta}: <b>{total:.1f} L/m²</b>"


def bloque_lluvia_hidro(anio, mes, dias, hasta=None):
    """Desglose mes a mes de la lluvia del año hidrológico al que pertenece
    el mes indicado, desde el 1 de septiembre hasta el final de ese mes
    (o hasta 'hasta' si el mes está en curso)."""
    fin = hasta or date(anio, mes, calendar.monthrange(anio, mes)[1])
    inicio, etiqueta = anio_hidrologico(fin)
    del_periodo = [d for d in dias if inicio <= d["fecha"] <= fin]
    if not del_periodo:
        return []

    por_mes = {}
    for d in del_periodo:
        clave = (d["fecha"].year, d["fecha"].month)
        por_mes[clave] = por_mes.get(clave, 0.0) + (d["lluvia"] or 0.0)
    total = sum(por_mes.values())

    hasta_txt = f"hasta el {hasta.strftime('%d/%m')}" if hasta else f"a fin de {NOMBRES_MES[mes]}"
    lineas = [f"🪣 <b>Año hidrológico {etiqueta}</b> ({hasta_txt}): <b>{total:.1f} L/m²</b>"]
    for (a, m), mm in sorted(por_mes.items()):
        lineas.append(f"   • {NOMBRES_MES[m].capitalize()} {a}: {mm:.1f} L/m²")

    primer_dia = min(d["fecha"] for d in del_periodo)
    if primer_dia > inicio:
        lineas.append(f"   <i>(registro desde el {primer_dia.strftime('%d/%m/%Y')})</i>")
    return lineas


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


def construir_mensaje_mensual(stats, anio, mes, dias_todos=None, en_curso_hasta=None):
    """Texto del resumen mensual. Si se pasa dias_todos (hoja Diario completa)
    se añade el desglose de lluvia del año hidrológico. en_curso_hasta marca
    un mes todavía sin terminar (último día registrado)."""
    nombre = NOMBRES_MES[mes].capitalize()
    titulo = f"📊 <b>Resumen mensual — {nombre} {anio}</b>"
    if en_curso_hasta:
        titulo += f"\n<i>Mes en curso, datos hasta el {en_curso_hasta.strftime('%d/%m')}</i>"
    lineas = [
        titulo,
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
        f"🌧 Lluvia del mes: <b>{stats['lluvia_total']:.1f} L/m²</b> "
        f"({stats['dias_lluvia']} {'día' if stats['dias_lluvia'] == 1 else 'días'} de lluvia)",
        f"🔵 Presión media: <b>{_fmt(stats['media_presion'], ' hPa')}</b>",
        f"💧 Humedad media: {_fmt(stats['media_hum_min'], '%')} – {_fmt(stats['media_hum_max'], '%')}",
    ]
    if stats["media_solar"] is not None:
        lineas.append(f"☀️ Media radiación máx: <b>{_fmt(stats['media_solar'], ' W/m²')}</b>")
    if stats["media_uv"] is not None:
        lineas.append(f"🕶 Media UV máx: <b>{_fmt(stats['media_uv'])}</b>")

    if dias_todos:
        hidro = bloque_lluvia_hidro(anio, mes, dias_todos, hasta=en_curso_hasta)
        if hidro:
            lineas += [""] + hidro
    return "\n".join(lineas)


def texto_detalle_lluvia(ini, fin, intensidad, hora_int):
    """' · de 17:05 a 18:40 · intensidad máx 12.4 mm/h (17:20)' o '' si no hay datos."""
    partes = []
    if ini and fin and ini != fin:
        partes.append(f"de {ini} a {fin}")
    elif ini:
        partes.append(f"a las {ini}")
    if intensidad is not None:
        hora_txt = f" ({hora_int})" if hora_int else ""
        partes.append(f"intensidad máx {intensidad:.1f} mm/h{hora_txt}")
    return "".join(f" · {p}" for p in partes)


def teclado_lluvia(anio, mes):
    """Botón que se adjunta a cualquier resumen mensual (lo atiende el listener)."""
    return {"inline_keyboard": [[
        {"text": "🌧 Ver días de lluvia", "callback_data": f"l:{anio}:{mes}"}
    ]]}


def construir_mensaje_lluvia_mes(anio, mes, dias_todos=None):
    """Lista de días con lluvia del mes, con hora e intensidad cuando se registraron."""
    dias = leer_dias_del_mes(anio, mes, dias_todos)
    nombre = f"{NOMBRES_MES[mes].capitalize()} {anio}"
    if not dias:
        return f"No hay datos registrados de {NOMBRES_MES[mes]} {anio}."
    lluviosos = [d for d in dias if (d["lluvia"] or 0) > 0]
    if not lluviosos:
        return f"☀️ <b>{nombre}</b>: sin lluvia registrada ({len(dias)} días con datos)."

    lineas = [f"🌧 <b>Días de lluvia — {nombre}</b>", ""]
    for d in lluviosos:
        detalle = texto_detalle_lluvia(
            d.get("lluvia_ini"), d.get("lluvia_fin"), d.get("intensidad"), d.get("hora_intensidad")
        )
        lineas.append(f"• {d['fecha'].strftime('%d/%m')}: <b>{d['lluvia']:.1f} L/m²</b>{detalle}")

    total = sum(d["lluvia"] for d in lluviosos)
    mas = max(lluviosos, key=lambda d: d["lluvia"])
    n = len(lluviosos)
    lineas += [
        "",
        f"Total: <b>{total:.1f} L/m²</b> en {n} {'día' if n == 1 else 'días'}",
        f"Día más lluvioso: {mas['fecha'].strftime('%d/%m')} ({mas['lluvia']:.1f} L/m²)",
    ]
    con_int = [d for d in lluviosos if d.get("intensidad") is not None]
    if con_int:
        fuerte = max(con_int, key=lambda d: d["intensidad"])
        hora_txt = f" a las {fuerte['hora_intensidad']}" if fuerte.get("hora_intensidad") else ""
        lineas.append(
            f"Mayor intensidad: {fuerte['intensidad']:.1f} mm/h "
            f"({fuerte['fecha'].strftime('%d/%m')}{hora_txt})"
        )
    if len(con_int) < n:
        lineas.append("<i>Los días sin hora ni intensidad son anteriores a que se empezaran a registrar.</i>")
    return "\n".join(lineas)


def resumen_mensual(anio, mes, hoy=None, dias_todos=None):
    """Resumen de cualquier mes calculado al vuelo desde la hoja Diario.
    Devuelve (texto, stats) o (None, None) si no hay días de ese mes.
    Lo usan el modo 'mensual' y el listener (/mensual, /mes)."""
    hoy = hoy or date.today()
    dias_todos = leer_diario() if dias_todos is None else dias_todos
    dias = leer_dias_del_mes(anio, mes, dias_todos)
    if not dias:
        return None, None
    stats = calcular_medias_mes(dias)
    en_curso = (anio, mes) == (hoy.year, hoy.month)
    ultimo = max(d["fecha"] for d in dias)
    texto = construir_mensaje_mensual(
        stats, anio, mes, dias_todos, en_curso_hasta=ultimo if en_curso else None
    )
    return texto, stats


def ejecutar_mensual(hoy):
    """Calcula y registra las medias del mes anterior."""
    ultimo_dia_mes_anterior = hoy.replace(day=1) - timedelta(days=1)
    anio, mes = ultimo_dia_mes_anterior.year, ultimo_dia_mes_anterior.month
    log.info("Calculando medias de %02d/%d", mes, anio)

    texto, stats = resumen_mensual(anio, mes, hoy)
    if not stats:
        log.warning("No hay días registrados para %02d/%d", mes, anio)
        enviar_telegram(
            f"⚠️ <b>Bot Meteo</b>: no hay datos registrados de "
            f"{NOMBRES_MES[mes]} {anio}, no se pueden calcular medias.",
        )
        return

    registrar_medias(stats, anio, mes)
    enviar_telegram(texto, teclado=teclado_lluvia(anio, mes))
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


def novedades_desde_snapshot(resumen, hoy, total_hidro=None, total_mes=None):
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
            + texto_detalle_lluvia(
                resumen.get("lluvia_ini"), resumen.get("lluvia_fin"),
                resumen.get("intensidad_max"), resumen.get("hora_intensidad"),
            )
        )
        # Los acumulados dados a las 21:00 han cambiado: se actualizan
        for linea in (linea_mes(total_mes, hoy), linea_hidro(total_hidro, hoy)):
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


def construir_mensaje(resumen, batidos, hoy, total_hidro=None, total_mes=None):
    lineas = [
        f"🌦 <b>Resumen meteo — {hoy.strftime('%d/%m/%Y')}</b>",
        "",
        f"🌡 Máxima: <b>{_fmt(resumen['temp_max'], '°C')}</b> ({resumen['hora_temp_max'] or '—'})",
        f"❄️ Mínima: <b>{_fmt(resumen['temp_min'], '°C')}</b> ({resumen['hora_temp_min'] or '—'})",
        f"💨 Racha máx: <b>{_fmt(resumen['racha_max'], ' km/h')}</b> ({resumen['hora_racha'] or '—'})",
    ]
    if resumen["lluvia"] and resumen["lluvia"] > 0:
        detalle = texto_detalle_lluvia(
            resumen.get("lluvia_ini"), resumen.get("lluvia_fin"),
            resumen.get("intensidad_max"), resumen.get("hora_intensidad"),
        )
        lineas.append(f"🌧 Lluvia: <b>{resumen['lluvia']:.1f} mm</b>{detalle}")
    for linea in (linea_mes(total_mes, hoy), linea_hidro(total_hidro, hoy)):
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


def enviar_telegram(texto, silencioso=False, teclado=None):
    url = f"https://api.telegram.org/bot{TELEGRAM_TOKEN}/sendMessage"
    payload = {
        "chat_id": TELEGRAM_CHAT_ID,
        "text": texto,
        "parse_mode": "HTML",
        "disable_notification": silencioso,
    }
    if teclado:
        payload["reply_markup"] = teclado
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
        # Acumulados de mes y año hidrológico: días ya registrados + lo que lleva hoy
        total_mes, total_hidro = acumulados_lluvia(hoy, incluir_hoy=False)
        hoy_mm = resumen["lluvia"] or 0.0
        if total_mes is not None:
            total_mes += hoy_mm
        if total_hidro is not None:
            total_hidro += hoy_mm
        enviar_telegram(construir_mensaje(resumen, [], hoy, total_hidro, total_mes))
        guardar_snapshot(resumen, hoy)
    else:
        # 23:57 — registro definitivo del día completo, sin despertar a nadie
        batidos = comprobar_records(resumen, hoy)
        registrar_en_excel(resumen, batidos, hoy)
        total_mes, total_hidro = acumulados_lluvia(hoy, incluir_hoy=True)
        log.info("Lluvia mes: %s mm / año hidrológico: %s mm", total_mes, total_hidro)

        bloques = []

        novedades = novedades_desde_snapshot(resumen, hoy, total_hidro, total_mes)
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
