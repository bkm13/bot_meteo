#!/usr/bin/env python3
"""
Bot Meteo - Previsión y barómetro
=================================
Modos (systemd timers):
  prevision   12:00      - Previsión de mañana (Open-Meteo) corregida con lo
                           que mide la estación + lectura del barómetro
  vigilancia  cada 15 min - Guarda la presión; si baja deprisa avisa de que
                           puede llover pronto, a cualquier hora
  calibrar    a mano     - Reconstruye las previsiones pasadas (desde que hay
                           estación) para calcular la corrección desde el día 1

Previsión: modelo de Open-Meteo (sin clave) para las coordenadas de la
estación. Cada día se guarda en predicciones.csv lo que dijo el modelo y se
compara con la máxima/mínima medidas: la media de esos errores en los
últimos 30 días es la corrección que se aplica ("en mi casa la máxima sale
+1.2 °C sobre el modelo").

Barómetro: algoritmo Zambretti (presión, tendencia en 3 h, viento y época
del año). La tendencia se corrige de la marea barométrica: la presión baja
sola ~1-2 hPa entre las 10 y las 16 h (y entre las 22 y las 4 h) todos los
días; sin corregirla habría falsos avisos casi cada tarde.

Variables de entorno (opcionales, además de las de /etc/bot-meteo.env):
  PRED_LAT, PRED_LON        coordenadas (por defecto, las de la estación)
  AVISO_CAIDA_FUERTE        hPa en 3 h para avisar siempre (2.5)
  AVISO_CAIDA               hPa en 3 h para avisar si el modelo o el
                            barómetro también ven lluvia (1.6)
  MAREA_HPA                 amplitud de la marea barométrica (1.0)
"""

import os
import sys
import csv
import json
import math
import time
import logging
from datetime import datetime, date, timedelta
from pathlib import Path

import requests

sys.path.insert(0, str(Path(__file__).resolve().parent))
from bot_meteo import (  # noqa: E402
    leer_diario, enviar_telegram, _flt, _leer_wu,
    WU_API_KEY, WU_STATION_ID, NOMBRES_MES,
)

BASE_DIR = Path(__file__).resolve().parent
PRED_CSV_PATH = BASE_DIR / "predicciones.csv"
PRESION_PATH = BASE_DIR / "presion_historial.json"
ALERTA_PATH = BASE_DIR / "alerta_presion.json"

LAT = float(os.environ.get("PRED_LAT", "36.9983046"))
LON = float(os.environ.get("PRED_LON", "-5.5739743"))
CAIDA_FUERTE = float(os.environ.get("AVISO_CAIDA_FUERTE", "2.5"))
CAIDA = float(os.environ.get("AVISO_CAIDA", "1.6"))
MAREA_HPA = float(os.environ.get("MAREA_HPA", "1.0"))
HORAS_ENTRE_AVISOS = 6
DIAS_CORRECCION = 30
MIN_DIAS_CORRECCION = 5

OM_URL = "https://api.open-meteo.com/v1/forecast"
OM_PREV_URL = "https://previous-runs-api.open-meteo.com/v1/forecast"
WU_CURRENT_URL = (
    "https://api.weather.com/v2/pws/observations/current"
    "?stationId={station}&format=json&units=m&numericPrecision=decimal&apiKey={key}"
)

log = logging.getLogger("bot_meteo")


def ahora_madrid():
    """Hora local de Madrid sin tzinfo (las horas de Open-Meteo vienen así),
    aunque el reloj de la VM esté en UTC."""
    try:
        from zoneinfo import ZoneInfo
        return datetime.now(ZoneInfo("Europe/Madrid")).replace(tzinfo=None)
    except Exception:
        return datetime.now()


def hoy_madrid():
    return ahora_madrid().date()

DIAS_SEMANA = ["lunes", "martes", "miércoles", "jueves", "viernes", "sábado", "domingo"]
RUMBOS = ["N", "NNE", "NE", "ENE", "E", "ESE", "SE", "SSE",
          "S", "SSO", "SO", "OSO", "O", "ONO", "NO", "NNO"]

CABECERA_PRED = [
    "Fecha", "Emitida", "Tmax modelo (°C)", "Tmin modelo (°C)",
    "Tmax corregida (°C)", "Tmin corregida (°C)", "Lluvia modelo (mm)",
    "Prob lluvia (%)", "Racha modelo (km/h)", "Código tiempo",
]

CODIGOS_TIEMPO = {
    0: ("☀️", "Despejado"), 1: ("🌤", "Casi despejado"), 2: ("⛅️", "Intervalos nubosos"),
    3: ("☁️", "Cubierto"), 45: ("🌫", "Niebla"), 48: ("🌫", "Niebla con escarcha"),
    51: ("🌦", "Llovizna débil"), 53: ("🌦", "Llovizna"), 55: ("🌦", "Llovizna intensa"),
    56: ("🌧", "Llovizna helada"), 57: ("🌧", "Llovizna helada intensa"),
    61: ("🌧", "Lluvia débil"), 63: ("🌧", "Lluvia"), 65: ("🌧", "Lluvia fuerte"),
    66: ("🌧", "Lluvia helada"), 67: ("🌧", "Lluvia helada fuerte"),
    71: ("🌨", "Nieve débil"), 73: ("🌨", "Nieve"), 75: ("🌨", "Nieve fuerte"),
    77: ("🌨", "Granos de nieve"), 80: ("🌦", "Chubascos débiles"), 81: ("🌦", "Chubascos"),
    82: ("⛈", "Chubascos fuertes"), 85: ("🌨", "Chubascos de nieve"),
    86: ("🌨", "Chubascos de nieve fuertes"), 95: ("⛈", "Tormenta"),
    96: ("⛈", "Tormenta con granizo"), 99: ("⛈", "Tormenta fuerte con granizo"),
}


def rumbo(grados):
    g = _flt(grados)
    return "" if g is None else RUMBOS[int((g + 11.25) // 22.5) % 16]


def _fmt1(v, unidad=""):
    return "n/d" if v is None else f"{v:.1f}{unidad}"


# ---------------------------------------------------------------------------
# Open-Meteo
# ---------------------------------------------------------------------------
def _get_json(url, params, intentos=3):
    for i in range(intentos):
        try:
            r = requests.get(url, params=params, timeout=30)
            r.raise_for_status()
            return r.json()
        except Exception as e:
            log.warning("Open-Meteo (intento %d): %s", i + 1, e)
            time.sleep(2)
    return None


def pedir_pronostico(dias=3):
    """Pronóstico diario y horario. None si falla."""
    return _get_json(OM_URL, {
        "latitude": LAT, "longitude": LON, "timezone": "Europe/Madrid",
        "forecast_days": dias,
        "daily": ",".join([
            "weather_code", "temperature_2m_max", "temperature_2m_min",
            "precipitation_sum", "precipitation_probability_max",
            "wind_speed_10m_max", "wind_gusts_10m_max", "wind_direction_10m_dominant",
        ]),
        "hourly": "precipitation_probability,precipitation",
    })


def pronostico_dia(datos, fecha):
    """Extrae del JSON de Open-Meteo los datos del día 'fecha'."""
    if not datos or "daily" not in datos:
        return None
    d = datos["daily"]
    try:
        i = d["time"].index(fecha.isoformat())
    except (KeyError, ValueError):
        return None

    def v(clave):
        lista = d.get(clave) or []
        return lista[i] if i < len(lista) else None

    p = {
        "fecha": fecha,
        "codigo": v("weather_code"),
        "tmax": _flt(v("temperature_2m_max")),
        "tmin": _flt(v("temperature_2m_min")),
        "lluvia": _flt(v("precipitation_sum")) or 0.0,
        "prob": _flt(v("precipitation_probability_max")),
        "viento": _flt(v("wind_speed_10m_max")),
        "racha": _flt(v("wind_gusts_10m_max")),
        "dir": v("wind_direction_10m_dominant"),
        "franjas": franjas_lluvia(datos, fecha),
    }
    return p


def franjas_lluvia(datos, fecha):
    """Tramos horarios del día con lluvia probable: [(hora_ini, hora_fin), ...]."""
    h = (datos or {}).get("hourly") or {}
    horas = h.get("time") or []
    probs = h.get("precipitation_probability") or []
    mm = h.get("precipitation") or []
    prefijo = fecha.isoformat()
    lluviosas = []
    for k, t in enumerate(horas):
        if not t.startswith(prefijo):
            continue
        pr = _flt(probs[k]) if k < len(probs) else None
        x = _flt(mm[k]) if k < len(mm) else None
        if (pr is not None and pr >= 50) or (x is not None and x >= 0.3):
            lluviosas.append(int(t[11:13]))
    tramos = []
    for hora in lluviosas:
        if tramos and hora <= tramos[-1][1] + 2:  # une huecos de una hora
            tramos[-1][1] = hora
        else:
            tramos.append([hora, hora])
    return [(a, b + 1) for a, b in tramos]


def texto_franjas(franjas):
    partes = []
    for a, b in franjas:
        if a == 0 and b == 24:
            partes.append("todo el día")
        elif b == 24:
            partes.append(f"desde las {a} h")
        else:
            partes.append(f"de {a} a {b} h")
    if len(partes) > 1:
        return ", ".join(partes[:-1]) + " y " + partes[-1]
    return partes[0] if partes else ""


def lluvia_proximas_horas(datos, ahora, horas=6):
    """(prob. máxima %, mm totales, primera hora con lluvia probable) en las
    próximas 'horas' horas, según el pronóstico horario."""
    h = (datos or {}).get("hourly") or {}
    tiempos = h.get("time") or []
    probs = h.get("precipitation_probability") or []
    mm = h.get("precipitation") or []
    desde = ahora.replace(minute=0, second=0, microsecond=0)
    hasta = desde + timedelta(hours=horas)
    pmax, total, primera = None, 0.0, None
    for k, t in enumerate(tiempos):
        try:
            th = datetime.strptime(t, "%Y-%m-%dT%H:%M")
        except ValueError:
            continue
        if not desde <= th <= hasta:
            continue
        pr = _flt(probs[k]) if k < len(probs) else None
        x = _flt(mm[k]) if k < len(mm) else None
        if pr is not None:
            pmax = pr if pmax is None else max(pmax, pr)
        total += x or 0.0
        if primera is None and ((pr or 0) >= 50 or (x or 0) >= 0.3):
            primera = th
    return pmax, total, primera


# ---------------------------------------------------------------------------
# Registro de previsiones y corrección con la estación
# ---------------------------------------------------------------------------
def leer_predicciones():
    """{fecha: dict} con lo guardado en predicciones.csv."""
    res = {}
    if not PRED_CSV_PATH.exists():
        return res
    with open(PRED_CSV_PATH, newline="", encoding="utf-8-sig") as f:
        for fila in csv.DictReader(f):
            try:
                fecha = datetime.strptime(fila["Fecha"], "%d/%m/%Y").date()
            except (KeyError, ValueError):
                continue
            res[fecha] = {
                "emitida": fila.get("Emitida", ""),
                "tmax": _flt(fila.get("Tmax modelo (°C)")),
                "tmin": _flt(fila.get("Tmin modelo (°C)")),
                "tmax_corr": _flt(fila.get("Tmax corregida (°C)")),
                "tmin_corr": _flt(fila.get("Tmin corregida (°C)")),
                "lluvia": _flt(fila.get("Lluvia modelo (mm)")),
                "prob": _flt(fila.get("Prob lluvia (%)")),
                "racha": _flt(fila.get("Racha modelo (km/h)")),
                "codigo": fila.get("Código tiempo", ""),
            }
    return res


def escribir_predicciones(preds):
    with open(PRED_CSV_PATH, "w", newline="", encoding="utf-8-sig") as f:
        w = csv.writer(f)
        w.writerow(CABECERA_PRED)
        for fecha in sorted(preds):
            p = preds[fecha]
            w.writerow([
                fecha.strftime("%d/%m/%Y"), p.get("emitida", ""),
                _r(p.get("tmax")), _r(p.get("tmin")),
                _r(p.get("tmax_corr")), _r(p.get("tmin_corr")),
                _r(p.get("lluvia")), _r(p.get("prob")), _r(p.get("racha")),
                p.get("codigo", "") if p.get("codigo") is not None else "",
            ])


def _r(v):
    return "" if v is None else round(v, 1)


def guardar_prediccion(p, corr):
    preds = leer_predicciones()
    preds[p["fecha"]] = {
        "emitida": ahora_madrid().strftime("%d/%m/%Y %H:%M"),
        "tmax": p["tmax"], "tmin": p["tmin"],
        "tmax_corr": corr["tmax"], "tmin_corr": corr["tmin"],
        "lluvia": p["lluvia"], "prob": p["prob"], "racha": p["racha"],
        "codigo": p["codigo"],
    }
    escribir_predicciones(preds)


def es_lluvioso(pred):
    """Día de lluvia según lo que dijo el modelo (lo único que se sabe al prever)."""
    return (pred.get("lluvia") or 0) >= 1


def evaluar(dias_diario=None, ultimos=DIAS_CORRECCION, hasta=None, filtro=None):
    """Compara previsiones guardadas con lo medido.
    Devuelve dict con sesgo y errores de tmax/tmin, factor de rachas y
    aciertos de lluvia. 'filtro' elige solo días parecidos (p. ej. lluviosos)."""
    dias_diario = leer_diario() if dias_diario is None else dias_diario
    medido = {d["fecha"]: d for d in dias_diario}
    preds = leer_predicciones()
    hasta = hasta or hoy_madrid()
    fechas = sorted(f for f in preds if f in medido and f < hasta
                    and (filtro is None or filtro(preds[f])))
    fechas = fechas[-ultimos:] if ultimos else fechas
    res = {"n": len(fechas), "desde": fechas[0] if fechas else None}
    for clave in ("tmax", "tmin"):
        col = "temp_max" if clave == "tmax" else "temp_min"
        errores = [(medido[f][col] - preds[f][clave]) for f in fechas
                   if medido[f].get(col) is not None and preds[f].get(clave) is not None]
        if errores:
            sesgo = sum(errores) / len(errores)
            res[clave] = {
                "n": len(errores),
                "sesgo": sesgo,
                "error_modelo": sum(abs(e) for e in errores) / len(errores),
                # error que habría tenido aplicando la corrección media
                "error_corregido": sum(abs(e - sesgo) for e in errores) / len(errores),
            }
    rachas = [(medido[f]["racha"], preds[f]["racha"]) for f in fechas
              if medido[f].get("racha") and preds[f].get("racha")]
    if rachas:
        res["racha"] = {"n": len(rachas),
                        "factor": sum(m for m, _ in rachas) / sum(p for _, p in rachas)}
    lluvia = [(preds[f].get("lluvia") or 0, medido[f].get("lluvia") or 0) for f in fechas]
    if lluvia:
        dijo = [p >= 1 for p, _ in lluvia]
        llovio = [o >= 1 for _, o in lluvia]
        res["lluvia"] = {
            "n": len(lluvia),
            "aciertos": sum(1 for a, b in zip(dijo, llovio) if a == b),
            "dias_lluvia": sum(llovio),
            "avisados": sum(1 for a, b in zip(dijo, llovio) if a and b),
            "falsas": sum(1 for a, b in zip(dijo, llovio) if a and not b),
        }
    return res


def evaluar_parecidos(p, dias_diario=None):
    """Evaluación con los días que más se parecen al previsto: si el modelo
    da lluvia, solo días de lluvia (nubes y viento: casi no hay inversión
    nocturna); si no, los días secos de los últimos 30. Si no hay bastantes
    días parecidos, se usan todos y en días de lluvia la corrección se reduce
    a la mitad. Devuelve (evaluación, descripción, factor)."""
    lluvioso = es_lluvioso(p)
    if lluvioso:
        ev = evaluar(dias_diario, ultimos=None, filtro=es_lluvioso)
        desc = "días de lluvia"
    else:
        ev = evaluar(dias_diario, filtro=lambda r: not es_lluvioso(r))
        desc = "días secos"
    if all((ev.get(c) or {}).get("n", 0) >= MIN_DIAS_CORRECCION for c in ("tmax", "tmin")):
        return ev, desc, 1.0
    return evaluar(dias_diario), "todos los días", 0.5 if lluvioso else 1.0


def corregir(p, ev, factor=1.0):
    """Aplica la corrección media si hay bastantes días evaluados."""
    corr = {"tmax": p["tmax"], "tmin": p["tmin"], "racha": p.get("racha"),
            "aplicada": False, "racha_corr": False}
    for clave in ("tmax", "tmin"):
        e = ev.get(clave)
        if e and e["n"] >= MIN_DIAS_CORRECCION and p[clave] is not None:
            corr[clave] = p[clave] + e["sesgo"] * factor
            corr["aplicada"] = True
    e = ev.get("racha")
    if e and e["n"] >= MIN_DIAS_CORRECCION and p.get("racha") is not None:
        corr["racha"] = p["racha"] * e["factor"]
        corr["racha_corr"] = True
    return corr


def calibrar_historico(dias_diario=None):
    """Reconstruye las previsiones de los días pasados con datos de la
    estación usando la API de ejecuciones anteriores de Open-Meteo
    (lo que el modelo pronosticaba el día antes). No toca días ya guardados."""
    dias_diario = leer_diario() if dias_diario is None else dias_diario
    if not dias_diario:
        return 0
    inicio = min(d["fecha"] for d in dias_diario)
    fin = hoy_madrid() - timedelta(days=1)
    datos = _get_json(OM_PREV_URL, {
        "latitude": LAT, "longitude": LON, "timezone": "Europe/Madrid",
        "start_date": inicio.isoformat(), "end_date": fin.isoformat(),
        "hourly": "temperature_2m_previous_day1,precipitation_previous_day1,"
                  "wind_gusts_10m_previous_day1",
    })
    h = (datos or {}).get("hourly") or {}
    tiempos = h.get("time") or []
    temps = h.get("temperature_2m_previous_day1") or []
    lluvias = h.get("precipitation_previous_day1") or []
    rachas = h.get("wind_gusts_10m_previous_day1") or []
    por_dia = {}
    for k, t in enumerate(tiempos):
        f = date.fromisoformat(t[:10])
        tv = _flt(temps[k]) if k < len(temps) else None
        lv = _flt(lluvias[k]) if k < len(lluvias) else None
        dia = por_dia.setdefault(f, {"t": [], "ll": 0.0, "n_ll": 0, "r": []})
        rv = _flt(rachas[k]) if k < len(rachas) else None
        if rv is not None:
            dia["r"].append(rv)
        if tv is not None:
            dia["t"].append(tv)
        if lv is not None:
            dia["ll"] += lv
            dia["n_ll"] += 1
    preds = leer_predicciones()
    nuevos = 0
    for f, dia in por_dia.items():
        racha = max(dia["r"]) if len(dia["r"]) >= 20 else None
        if f in preds:
            # Filas reconstruidas antes de guardar rachas: se completan
            if preds[f].get("racha") is None and racha is not None:
                preds[f]["racha"] = racha
                nuevos += 1
            continue
        if len(dia["t"]) < 20:
            continue
        preds[f] = {
            "emitida": "reconstruida",
            "tmax": max(dia["t"]), "tmin": min(dia["t"]),
            "tmax_corr": None, "tmin_corr": None,
            "lluvia": dia["ll"] if dia["n_ll"] >= 20 else None,
            "prob": None, "racha": racha, "codigo": None,
        }
        nuevos += 1
    if nuevos:
        escribir_predicciones(preds)
    log.info("Calibración: %d días reconstruidos o completados", nuevos)
    return nuevos


# ---------------------------------------------------------------------------
# Barómetro: marea barométrica y Zambretti
# ---------------------------------------------------------------------------
def marea(epoch):
    """Componente semidiurna de la presión (hPa) en ese instante.
    Máximos hacia las 10 y 22 h solares, mínimos hacia las 4 y 16 h."""
    hora_utc = (epoch % 86400) / 3600
    hora_solar = (hora_utc + LON / 15) % 24
    return MAREA_HPA * math.cos(2 * math.pi * (hora_solar - 10) / 12)


ZAMBRETTI = {
    # Z: (texto, nivel de lluvia 0 nada · 1 posible · 2 probable · 3 mucha)
    1: ("Buen tiempo estable", 0), 2: ("Buen tiempo", 0),
    3: ("Bueno, volviéndose inestable", 0), 4: ("Bastante bueno, chubascos más tarde", 1),
    5: ("Chubascos, cada vez más inestable", 2), 6: ("Inestable, lluvia más tarde", 2),
    7: ("Lluvia a ratos, empeorando", 2), 8: ("Lluvia a ratos, muy inestable", 3),
    9: ("Muy inestable, lluvia", 3),
    10: ("Buen tiempo estable", 0), 11: ("Buen tiempo", 0),
    12: ("Bueno, posibles chubascos", 1), 13: ("Bastante bueno, chubascos probables", 1),
    14: ("Chubascos con claros", 2), 15: ("Variable, algo de lluvia", 2),
    16: ("Inestable, lluvia a ratos", 2), 17: ("Lluvia frecuente", 3),
    18: ("Muy inestable, lluvia", 3), 19: ("Tormentoso, mucha lluvia", 3),
    20: ("Buen tiempo estable", 0), 21: ("Buen tiempo", 0), 22: ("Mejorando", 0),
    23: ("Bastante bueno, mejorando", 0), 24: ("Bastante bueno, posibles chubascos al principio", 1),
    25: ("Chubascos al principio, mejorando", 1), 26: ("Variable, mejorando", 1),
    27: ("Algo inestable, despejando más tarde", 1), 28: ("Inestable, probablemente mejorando", 1),
    29: ("Inestable, con ratos de buen tiempo", 2), 30: ("Muy inestable, mejorando a ratos", 2),
    31: ("Tormentoso, posiblemente mejorando", 2), 32: ("Tormentoso, mucha lluvia", 3),
}
# Corrección de la presión según el viento (hPa). Los vientos del S/SO traen
# las borrascas atlánticas; el N/NE y el levante son secos.
AJUSTE_VIENTO = {
    "N": 6, "NNE": 5, "NE": 5, "ENE": 2, "E": -0.5, "ESE": -2, "SE": -5, "SSE": -8.5,
    "S": -12, "SSO": -10, "SO": -6, "OSO": -4.5, "O": -3, "ONO": -0.5, "NO": 1.5, "NNO": 3,
}


def zambretti(presion, tendencia, dir_viento=None, viento=None, mes=None):
    """Pronóstico Zambretti (12 h). tendencia = cambio en 3 h (hPa, ya sin marea).
    Devuelve (Z, texto, nivel_lluvia, estado: 'bajando'|'estable'|'subiendo')."""
    mes = mes or hoy_madrid().month
    p = presion
    r = rumbo(dir_viento)
    if r and (viento is None or viento >= 3):
        p += AJUSTE_VIENTO.get(r, 0)
    p = min(max(p, 950), 1050)
    if tendencia <= -1.6:
        estado, z = "bajando", round(127 - 0.12 * p)
        z = z + 1 if mes in (10, 11, 12, 1, 2, 3) else z
        z = min(max(z, 1), 9)
    elif tendencia >= 1.6:
        estado, z = "subiendo", round(185 - 0.16 * p)
        z = z - 1 if mes in (4, 5, 6, 7, 8, 9) else z
        z = min(max(z, 20), 32)
    else:
        estado, z = "estable", round(144 - 0.13 * p)
        z = min(max(z, 10), 19)
    texto, nivel = ZAMBRETTI[z]
    return z, texto, nivel, estado


def cargar_historial():
    try:
        return json.loads(PRESION_PATH.read_text(encoding="utf-8"))
    except Exception:
        return []


def guardar_historial(hist):
    limite = time.time() - 12 * 3600
    hist = [h for h in hist if h["epoch"] >= limite]
    PRESION_PATH.write_text(json.dumps(hist), encoding="utf-8")
    return hist


def _lectura(obs):
    m = obs.get("metric", {}) or {}
    epoch = obs.get("epoch")
    if epoch is None:
        try:
            epoch = datetime.strptime(obs.get("obsTimeUtc", "")[:19], "%Y-%m-%dT%H:%M:%S")
            epoch = (epoch - datetime(1970, 1, 1)).total_seconds()
        except ValueError:
            return None
    p = _flt(m.get("pressure"))
    if p is None:  # observaciones del endpoint "all": máx/mín del intervalo
        hi, lo = _flt(m.get("pressureMax")), _flt(m.get("pressureMin"))
        p = (hi + lo) / 2 if hi is not None and lo is not None else None
    if p is None:
        return None
    return {
        "epoch": int(epoch), "p": round(p, 2),
        "dir": obs.get("winddir") if obs.get("winddir") is not None else obs.get("winddirAvg"),
        "viento": _flt(m.get("windSpeed")) if m.get("windSpeed") is not None else _flt(m.get("windspeedAvg")),
        "hum": _flt(obs.get("humidity")) if obs.get("humidity") is not None else _flt(obs.get("humidityAvg")),
        "ritmo": _flt(m.get("precipRate")),
    }


def obs_actual():
    url = WU_CURRENT_URL.format(station=WU_STATION_ID, key=WU_API_KEY)
    for i in range(3):
        try:
            r = requests.get(url, timeout=20)
            r.raise_for_status()
            obs = r.json().get("observations", [])
            if obs:
                return obs[0]
        except Exception as e:
            log.warning("WU current (intento %d): %s", i + 1, e)
    return None


def actualizar_historial(obs=None):
    """Añade la lectura actual al historial (lo rellena con el día de hoy si
    está vacío) y lo devuelve con la última lectura."""
    hist = cargar_historial()
    if len(hist) < 3:
        for o in _leer_wu() or []:
            lec = _lectura(o)
            if lec:
                hist.append(lec)
    obs = obs or obs_actual()
    actual = _lectura(obs) if obs else None
    if actual and not any(h["epoch"] == actual["epoch"] for h in hist):
        hist.append(actual)
    hist = sorted({h["epoch"]: h for h in hist}.values(), key=lambda h: h["epoch"])
    hist = guardar_historial(hist)
    return hist, actual or (hist[-1] if hist else None)


def tendencia(hist, actual, horas=3):
    """Cambio de presión en 'horas' horas: (bruto, sin marea). None si no hay
    lectura de hace ~3 h (±40 min)."""
    if not actual:
        return None
    objetivo = actual["epoch"] - horas * 3600
    cerca = [h for h in hist if abs(h["epoch"] - objetivo) <= 2400]
    if not cerca:
        return None
    antes = min(cerca, key=lambda h: abs(h["epoch"] - objetivo))
    bruto = actual["p"] - antes["p"]
    sin_marea = bruto - (marea(actual["epoch"]) - marea(antes["epoch"]))
    return bruto, sin_marea


def lectura_barometro(hist=None, actual=None):
    """Dict con presión, tendencias y Zambretti; None si no hay datos."""
    if hist is None:
        hist, actual = actualizar_historial()
    if not actual:
        return None
    t = tendencia(hist, actual)
    bruto, neta = t if t else (None, 0.0)
    z, texto, nivel, estado = zambretti(actual["p"], neta, actual.get("dir"), actual.get("viento"))
    return {"p": actual["p"], "bruto": bruto, "neta": neta if t else None, "z": z,
            "texto": texto, "nivel": nivel, "estado": estado, "actual": actual}


def linea_barometro(b):
    flecha = {"bajando": "↘️", "subiendo": "↗️", "estable": "➡️"}[b["estado"]]
    tend = f" ({b['neta']:+.1f} hPa en 3 h)" if b["neta"] is not None else ""
    return f"🧭 Barómetro: <b>{b['p']:.1f} hPa</b> {flecha} {b['estado']}{tend} → <i>{b['texto']}</i>"


def mensaje_barometro():
    b = lectura_barometro()
    if not b:
        return "⚠️ No se pudo leer la presión de la estación."
    lineas = ["🧭 <b>Barómetro de la estación</b>", "",
              f"🔵 Presión: <b>{b['p']:.1f} hPa</b>"]
    if b["neta"] is not None:
        lineas.append(f"📉 Últimas 3 h: {b['bruto']:+.1f} hPa medido · <b>{b['neta']:+.1f} hPa</b> "
                      f"sin la marea diaria → {b['estado']}")
    else:
        lineas.append("📉 Tendencia: aún no hay lecturas de hace 3 h")
    a = b["actual"]
    if a.get("dir") is not None:
        lineas.append(f"💨 Viento: {rumbo(a['dir'])} {_fmt1(a.get('viento'), ' km/h')}")
    lineas += ["", f"🔮 Próximas 12 h (Zambretti): <b>{b['texto']}</b>",
               "<i>Método clásico de barómetro: orientativo, se basa solo en tu estación.</i>"]
    return "\n".join(lineas)


# ---------------------------------------------------------------------------
# Mensajes de previsión
# ---------------------------------------------------------------------------
def mensaje_prevision(fecha=None, guardar=False, dias_diario=None):
    """Previsión de 'fecha' (por defecto mañana) corregida con la estación."""
    fecha = fecha or hoy_madrid() + timedelta(days=1)
    datos = pedir_pronostico()
    p = pronostico_dia(datos, fecha)
    if not p:
        return None
    dias_diario = leer_diario() if dias_diario is None else dias_diario
    ev, desc_ev, factor = evaluar_parecidos(p, dias_diario)
    corr = corregir(p, ev, factor)
    if guardar:
        guardar_prediccion(p, corr)

    emoji, desc = CODIGOS_TIEMPO.get(p["codigo"], ("🌡", "—"))
    lineas = [f"🔮 <b>Previsión para mañana — {DIAS_SEMANA[fecha.weekday()]} "
              f"{fecha.strftime('%d/%m')}</b>", "", f"{emoji} <b>{desc}</b>",
              f"🌡 Máxima: <b>{_fmt1(corr['tmax'], '°C')}</b> · Mínima: <b>{_fmt1(corr['tmin'], '°C')}</b>"]
    if corr["aplicada"]:
        partes = []
        for clave, nombre in (("tmax", "máx"), ("tmin", "mín")):
            e = ev.get(clave)
            if e and e["n"] >= MIN_DIAS_CORRECCION:
                partes.append(f"{nombre} {e['sesgo'] * factor:+.1f}")
        base = f"{ev['n']} {desc_ev}" + (", corrección a medias" if factor < 1 else "")
        lineas.append(f"   <i>Ajustada a tu estación ({', '.join(partes)} °C según {base}). "
                      f"Modelo: {_fmt1(p['tmax'])}/{_fmt1(p['tmin'])}</i>")
    if p["lluvia"] >= 0.1 or (p["prob"] or 0) >= 20:
        txt = f"🌧 Lluvia: <b>{p['lluvia']:.1f} L/m²</b>"
        if p["prob"] is not None:
            txt += f" · probabilidad {p['prob']:.0f}%"
        if p["franjas"]:
            txt += " · " + texto_franjas(p["franjas"])
        lineas.append(txt)
    else:
        lineas.append("🌂 Sin lluvia prevista")
    if p["viento"] is not None:
        txt = f"💨 Viento hasta {p['viento']:.0f} km/h del {rumbo(p['dir'])}".rstrip()
        if p["racha"] is not None:
            if corr["racha_corr"]:
                txt += (f" · rachas ≈{corr['racha']:.0f} km/h en tu estación "
                        f"(modelo {p['racha']:.0f})")
            else:
                txt += f" · rachas {p['racha']:.0f} km/h"
        lineas.append(txt)
    try:
        b = lectura_barometro()
        if b:
            lineas += ["", linea_barometro(b)]
    except Exception as e:
        log.warning("Barómetro no disponible: %s", e)
    return "\n".join(lineas)


def mensaje_aciertos(dias_diario=None):
    ev = evaluar(dias_diario, ultimos=None)
    if not ev["n"]:
        return ("🎯 Todavía no hay previsiones que comparar.\n"
                "Se guardan cada día a las 12:00 (o ejecuta <code>prediccion.py calibrar</code>).")
    ev30 = evaluar(dias_diario)
    lineas = [f"🎯 <b>Aciertos de la previsión</b>",
              f"<i>{ev['n']} días comparados desde el {ev['desde'].strftime('%d/%m/%Y')}</i>", ""]
    for clave, nombre in (("tmax", "Máxima"), ("tmin", "Mínima")):
        e = ev.get(clave)
        if not e:
            continue
        sentido = "más alta" if e["sesgo"] > 0 else "más baja"
        lineas.append(f"🌡 {nombre}: en tu estación sale de media <b>{abs(e['sesgo']):.1f} °C {sentido}</b> "
                      f"que el modelo")
        lineas.append(f"   Error medio: modelo {e['error_modelo']:.1f} °C → corregido "
                      f"{e['error_corregido']:.1f} °C")
    r = ev.get("racha")
    if r:
        lineas.append(f"💨 Rachas: tu estación mide de media el <b>{r['factor'] * 100:.0f}%</b> "
                      f"de lo que da el modelo ({r['n']} días)")
    l = ev.get("lluvia")
    if l:
        lineas.append(f"🌧 Lluvia (≥1 L/m²): acierta {l['aciertos'] * 100 // l['n']}% de los días")
        if l["dias_lluvia"]:
            lineas.append(f"   Avisó de {l['avisados']} de los {l['dias_lluvia']} días que llovió"
                          f" · {l['falsas']} falsas alarmas")
    if ev30.get("tmax") and ev30["n"] < ev["n"]:
        lineas += ["", f"Corrección actual (últimos {ev30['n']} días): máx "
                       f"{ev30['tmax']['sesgo']:+.1f} °C · mín {ev30.get('tmin', {}).get('sesgo', 0):+.1f} °C"]
    return "\n".join(lineas)


# ---------------------------------------------------------------------------
# Vigilancia de presión
# ---------------------------------------------------------------------------
def cargar_estado():
    try:
        return json.loads(ALERTA_PATH.read_text(encoding="utf-8"))
    except Exception:
        return {"activa": False, "ultima": 0}


def evaluar_aviso(b, modelo=None):
    """Decide si hay que avisar. modelo = (prob_max, mm, primera_hora) o None.
    Devuelve (avisar, motivo)."""
    if not b or b["neta"] is None:
        return False, "sin tendencia"
    caida = -b["neta"]
    if caida >= CAIDA_FUERTE:
        return True, "caída fuerte"
    if caida >= CAIDA:
        prob = (modelo or (None,))[0]
        if b["nivel"] >= 2 or (prob is not None and prob >= 50):
            return True, "caída + confirmación"
    return False, "sin caída suficiente"


def ejecutar_vigilancia():
    hist, actual = actualizar_historial()
    b = lectura_barometro(hist, actual)
    if not b:
        log.warning("Vigilancia: sin lectura de presión")
        return
    estado = cargar_estado()
    ahora = time.time()
    log.info("Vigilancia: %.1f hPa, tendencia 3h bruta %s / sin marea %s, Z=%s",
             b["p"], b["bruto"], b["neta"], b["z"])

    # Fin del episodio: la presión deja de bajar
    if estado.get("activa") and b["neta"] is not None and b["neta"] > -0.5:
        estado["activa"] = False
        ALERTA_PATH.write_text(json.dumps(estado), encoding="utf-8")

    if b["neta"] is None or -b["neta"] < CAIDA:
        return
    if estado.get("activa") and ahora - estado.get("ultima", 0) < HORAS_ENTRE_AVISOS * 3600:
        return
    if (b["actual"].get("ritmo") or 0) > 0:
        log.info("Ya está lloviendo: no se avisa")
        return

    datos = pedir_pronostico(dias=2)
    modelo = lluvia_proximas_horas(datos, ahora_madrid()) if datos else None
    avisar, motivo = evaluar_aviso(b, modelo)
    log.info("Vigilancia: avisar=%s (%s), modelo=%s", avisar, motivo, modelo)
    if not avisar:
        return
    if enviar_telegram(mensaje_aviso(b, modelo)):
        estado.update({"activa": True, "ultima": ahora})
        ALERTA_PATH.write_text(json.dumps(estado), encoding="utf-8")


def mensaje_aviso(b, modelo=None):
    lineas = ["⚠️ <b>La presión está bajando: puede llover pronto</b>", "",
              f"🔵 {b['p']:.1f} hPa · <b>{b['neta']:+.1f} hPa</b> en las últimas 3 h"
              + (f" (medido {b['bruto']:+.1f}, sin la marea diaria)" if b["bruto"] is not None
                 and abs(b["bruto"] - b["neta"]) >= 0.3 else ""),
              f"🧭 Barómetro: <i>{b['texto']}</i>"]
    if modelo and modelo[0] is not None:
        prob, mm, primera = modelo
        txt = f"🛰 El modelo da un {prob:.0f}% de lluvia en las próximas 6 h"
        if mm >= 0.1:
            txt += f" (≈{mm:.1f} L/m²"
            txt += f", desde las {primera.strftime('%H')} h)" if primera else ")"
        lineas.append(txt)
    a = b["actual"]
    extra = []
    if a.get("hum") is not None:
        extra.append(f"humedad {a['hum']:.0f}%")
    if a.get("dir") is not None:
        extra.append(f"viento {rumbo(a['dir'])} {_fmt1(a.get('viento'), ' km/h')}")
    if extra:
        texto = " · ".join(extra)
        lineas.append("💧 " + texto[0].upper() + texto[1:])
    return "\n".join(lineas)


# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------
def main():
    modo = sys.argv[1] if len(sys.argv) > 1 else ""
    if modo == "prevision":
        manana = hoy_madrid() + timedelta(days=1)
        try:
            calibrar_historico()
        except Exception as e:
            log.warning("Calibración fallida: %s", e)
        texto = mensaje_prevision(manana, guardar=True)
        if texto:
            enviar_telegram(texto)
        else:
            enviar_telegram("⚠️ <b>Bot Meteo</b>: no se pudo obtener la previsión de mañana.",
                            silencioso=True)
    elif modo == "vigilancia":
        ejecutar_vigilancia()
    elif modo == "calibrar":
        n = calibrar_historico()
        print(f"{n} días reconstruidos")
        print(mensaje_aciertos().replace("<b>", "").replace("</b>", "")
              .replace("<i>", "").replace("</i>", ""))
    else:
        print("Uso: prediccion.py prevision | vigilancia | calibrar")
        sys.exit(1)


if __name__ == "__main__":
    main()
