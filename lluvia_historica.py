#!/usr/bin/env python3
"""
Lluvia histórica - serie diaria desde el 01/10/1993
====================================================
Une la serie manual (lluvia_historica.csv, hasta el 05/08/2026) con el
registro de la estación (hoja Diario de registro_meteo.xlsx, desde el
06/08/2026). Si un día está en los dos sitios, manda la estación.

En el CSV, una celda vacía significa "sin dato" (no cuenta para medias ni
rachas) y un 0 significa "no llovió". Un mes, año o año hidrológico solo
entra en medias y rankings si tiene dato de todos sus días.

Este módulo no importa bot_meteo (evita importaciones circulares): quien lo
usa le pasa los días de la hoja Diario con cargar_serie(leer_diario()).
"""

import csv
import random
import calendar
from datetime import date, datetime, timedelta
from pathlib import Path

BASE_DIR = Path(__file__).resolve().parent
HIST_CSV_PATH = BASE_DIR / "lluvia_historica.csv"
CABECERA_HIST = ["Fecha", "Lluvia (L/m²)", "Observaciones"]
MES_INICIO_HIDRO = 9

NOMBRES_MES = [
    "", "enero", "febrero", "marzo", "abril", "mayo", "junio",
    "julio", "agosto", "septiembre", "octubre", "noviembre", "diciembre",
]
ABREV_MES = ["", "Ene", "Feb", "Mar", "Abr", "May", "Jun",
             "Jul", "Ago", "Sep", "Oct", "Nov", "Dic"]
DIAS_SEMANA = ["lunes", "martes", "miércoles", "jueves", "viernes", "sábado", "domingo"]


def plural_mes(m):
    return "abriles" if m == 4 else NOMBRES_MES[m] + "s"


def _mm(v):
    return f"{v:.1f}"


def _f(d):
    return d.strftime("%d/%m/%Y")


def _num(valor):
    """'12,5' / '12.5' / 12 -> float; vacío -> None."""
    if valor is None:
        return None
    if isinstance(valor, (int, float)):
        return float(valor)
    s = str(valor).strip().replace(",", ".")
    if not s:
        return None
    return float(s)


def _fecha(valor):
    if isinstance(valor, datetime):
        return valor.date()
    if isinstance(valor, date):
        return valor
    s = str(valor or "").strip()
    for fmt in ("%d/%m/%Y", "%Y-%m-%d", "%d-%m-%Y", "%Y-%m-%d %H:%M:%S"):
        try:
            return datetime.strptime(s, fmt).date()
        except ValueError:
            pass
    return None


def _ultimo_dia(a, m):
    return calendar.monthrange(a, m)[1]


def _barra(v, maximo, ancho=10):
    if not maximo or v is None:
        return ""
    n = round(v / maximo * ancho)
    return "▰" * n + "▱" * (ancho - n)


def detalle_estacion(det):
    """' · de 17:05 a 18:40 · intensidad máx 12.4 mm/h (17:20)' para días de la estación."""
    if not det:
        return ""
    ini, fin = det.get("lluvia_ini"), det.get("lluvia_fin")
    inten, hora = det.get("intensidad"), det.get("hora_intensidad")
    partes = []
    if ini and fin and ini != fin:
        partes.append(f"de {ini} a {fin}")
    elif ini:
        partes.append(f"a las {ini}")
    if inten is not None:
        partes.append(f"intensidad máx {inten:.1f} mm/h" + (f" ({hora})" if hora else ""))
    return "".join(f" · {p}" for p in partes)


# ---------------------------------------------------------------------------
# Lectura
# ---------------------------------------------------------------------------
def leer_historico(path=HIST_CSV_PATH):
    """Devuelve ({fecha: mm o None}, {fecha: observación})."""
    valores, notas = {}, {}
    if not Path(path).exists():
        return valores, notas
    with open(path, newline="", encoding="utf-8-sig") as f:
        muestra = f.read(2048)
        f.seek(0)
        sep = ";" if muestra.count(";") > muestra.count(",") else ","
        filas = csv.reader(f, delimiter=sep)
        next(filas, None)
        for fila in filas:
            if not fila:
                continue
            d = _fecha(fila[0])
            if d is None:
                continue
            try:
                valores[d] = _num(fila[1]) if len(fila) > 1 else None
            except ValueError:
                valores[d] = None
            if len(fila) > 2 and fila[2].strip():
                notas[d] = fila[2].strip()
    return valores, notas


def cargar_serie(dias_diario=None, path=HIST_CSV_PATH):
    """Serie completa: histórico + días de la estación (que mandan)."""
    valores, notas = leer_historico(path)
    detalle = {}
    for d in dias_diario or []:
        f = d.get("fecha")
        if f is None:
            continue
        valores[f] = d.get("lluvia") if d.get("lluvia") is not None else 0.0
        detalle[f] = d
    return SerieLluvia(valores, notas, detalle)


# ---------------------------------------------------------------------------
# Serie
# ---------------------------------------------------------------------------
class SerieLluvia:
    def __init__(self, valores, notas=None, detalle=None):
        self.notas = notas or {}
        self.detalle = detalle or {}
        conocidos = [d for d, v in valores.items() if v is not None]
        if not conocidos:
            self.inicio = self.fin = None
            self.v = {}
            self.conocidos = []
            return
        self.inicio, self.fin = min(conocidos), max(conocidos)
        self.v = {}
        d = self.inicio
        while d <= self.fin:
            self.v[d] = valores.get(d)
            d += timedelta(days=1)
        self.conocidos = [(d, v) for d, v in sorted(self.v.items()) if v is not None]
        self._cache = {}

    # -- básicos ------------------------------------------------------------
    def get(self, d):
        return self.v.get(d)

    def huecos(self):
        """Lista de (desde, hasta) sin dato dentro de la serie."""
        res, ini, prev = [], None, None
        for d, v in sorted(self.v.items()):
            if v is None:
                if ini is None:
                    ini = d
                prev = d
            elif ini is not None:
                res.append((ini, prev))
                ini = None
        if ini is not None:
            res.append((ini, prev))
        return res

    def periodo(self, desde, hasta):
        """Estadística de un periodo (incluye los dos extremos)."""
        total, con_dato, lluviosos, maximo, n = 0.0, 0, 0, None, 0
        d = desde
        while d <= hasta:
            n += 1
            v = self.v.get(d)
            if v is not None:
                con_dato += 1
                total += v
                if v > 0:
                    lluviosos += 1
                    if maximo is None or v > maximo[0]:
                        maximo = (v, d)
            d += timedelta(days=1)
        return {"total": round(total, 2), "con_dato": con_dato, "dias": n,
                "lluviosos": lluviosos, "max": maximo,
                "completo": con_dato == n and n > 0, "desde": desde, "hasta": hasta}

    def _agrupado(self, clave):
        if clave in self._cache:
            return self._cache[clave]
        grupos = {}
        if self.inicio:
            if clave == "mes":
                a, m = self.inicio.year, self.inicio.month
                while (a, m) <= (self.fin.year, self.fin.month):
                    grupos[(a, m)] = self.periodo(date(a, m, 1), date(a, m, _ultimo_dia(a, m)))
                    a, m = (a + 1, 1) if m == 12 else (a, m + 1)
            elif clave == "anio":
                for a in range(self.inicio.year, self.fin.year + 1):
                    grupos[a] = self.periodo(date(a, 1, 1), date(a, 12, 31))
            elif clave == "hidro":
                a0 = inicio_hidro(self.inicio).year
                a1 = inicio_hidro(self.fin).year
                for a in range(a0, a1 + 1):
                    grupos[a] = self.periodo(date(a, 9, 1), date(a + 1, 8, 31))
        self._cache[clave] = grupos
        return grupos

    def meses(self):
        return self._agrupado("mes")

    def anios(self):
        return self._agrupado("anio")

    def hidros(self):
        return self._agrupado("hidro")

    def meses_completos(self, m=None, excluir=None):
        return {k: p for k, p in self.meses().items()
                if p["completo"] and (m is None or k[1] == m) and k != excluir}

    def media_mes(self, m, excluir=None):
        tot = [p["total"] for p in self.meses_completos(m, excluir).values()]
        return (sum(tot) / len(tot), len(tot)) if tot else (None, 0)

    def media_anual(self):
        tot = [p["total"] for p in self.anios().values() if p["completo"]]
        return (sum(tot) / len(tot), len(tot)) if tot else (None, 0)

    def media_hidro(self):
        tot = [p["total"] for p in self.hidros().values() if p["completo"]]
        return (sum(tot) / len(tot), len(tot)) if tot else (None, 0)

    def media_hidro_a_fecha(self, fecha):
        """Media de lo acumulado desde el 1 de septiembre hasta el mismo día
        y mes de 'fecha', en los años hidrológicos completos."""
        sumas = []
        for a, p in self.hidros().items():
            if not p["completo"]:
                continue
            ini = date(a, 9, 1)
            anio_fin = a if fecha.month >= MES_INICIO_HIDRO else a + 1
            fin = date(anio_fin, fecha.month, min(fecha.day, _ultimo_dia(anio_fin, fecha.month)))
            sumas.append(self.periodo(ini, fin)["total"])
        return (sum(sumas) / len(sumas), len(sumas)) if sumas else (None, 0)

    def rachas(self, seca=True):
        """Lista de (longitud, desde, hasta) de días seguidos secos (o lluviosos)."""
        res, ini, n, prev = [], None, 0, None
        for d, v in sorted(self.v.items()):
            cumple = v is not None and ((v == 0) if seca else (v > 0))
            if cumple:
                if n == 0:
                    ini = d
                n += 1
                prev = d
            else:
                if n:
                    res.append((n, ini, prev))
                n = 0
        if n:
            res.append((n, ini, prev))
        return res

    def racha_seca_actual(self):
        n, d = 0, self.fin
        while d and self.v.get(d) == 0:
            n += 1
            d -= timedelta(days=1)
        return n

    def ranking_dia(self, mm, excluir=None):
        """Puesto que ocuparía un día con 'mm' entre todos los días de la serie."""
        return 1 + sum(1 for d, v in self.conocidos if v > mm and d != excluir)


def inicio_hidro(d):
    return date(d.year if d.month >= MES_INICIO_HIDRO else d.year - 1, MES_INICIO_HIDRO, 1)


def etiqueta_hidro(a):
    return f"{a}-{str(a + 1)[-2:]}"


def _rango_anios(claves):
    anios = sorted({k[0] if isinstance(k, tuple) else k for k in claves})
    return f"{anios[0]}-{anios[-1]}" if anios else ""


def _aviso_incompleto(p, hoy_fin=None):
    if p["completo"]:
        return None
    faltan = p["dias"] - p["con_dato"]
    if hoy_fin and p["hasta"] > hoy_fin:
        return f"<i>Periodo en curso, datos hasta el {hoy_fin.strftime('%d/%m/%Y')}</i>"
    return f"<i>⚠️ Faltan datos de {faltan} {'día' if faltan == 1 else 'días'}</i>"


def _pct(v, media):
    if not media:
        return ""
    return f"{v / media * 100:.0f}%"


# ---------------------------------------------------------------------------
# Comparaciones (resumen de las 21:00, registro, resumen mensual)
# ---------------------------------------------------------------------------
def comparar_dia(serie, mm, fecha, rnd=random, publico=False):
    """Líneas de 'en la historia' para un día con 'mm' litros.
    publico=True: frases explicadas para Facebook."""
    if mm is None or mm < 1 or not serie.conocidos:
        return []
    lineas = []
    otros = [(d, v) for d, v in serie.conocidos if v > 0 and d != fecha]
    tol = max(1.0, mm * 0.1)
    cercanos = [x for x in otros if abs(x[1] - mm) <= tol]
    iguales = [x for x in cercanos if abs(x[1] - mm) < 0.05]
    if iguales:
        d, v = rnd.choice(iguales)
        extra = f" (y {len(iguales) - 1} días más)" if len(iguales) > 1 else ""
        if publico:
            extra = (f". Ha pasado otras {len(iguales) - 1} veces más desde {serie.inicio.year}"
                     if len(iguales) > 1 else "")
            lineas.append(f"🔁 Ha caído exactamente lo mismo que el {_f(d)}: {_mm(v)} L/m²{extra}.")
        else:
            lineas.append(f"🔁 Lo mismo que el {_f(d)}: {_mm(v)} L/m²{extra}")
    elif cercanos:
        d, v = rnd.choice(cercanos)
        if publico:
            n = len(cercanos)
            lineas.append(f"🔁 Una lluvia parecida a la del {_f(d)}, cuando cayeron {_mm(v)} L/m². "
                          f"Desde {serie.inicio.year} ha habido "
                          f"{'1 día' if n == 1 else f'{n} días'} con una cantidad similar.")
        else:
            lineas.append(f"🔁 Parecido al {_f(d)}: {_mm(v)} L/m² "
                          f"({len(cercanos)} {'día parecido' if len(cercanos) == 1 else 'días parecidos'} en la serie)")
    elif otros:
        d, v = min(otros, key=lambda x: abs(x[1] - mm))
        if publico:
            lineas.append(f"🔁 El día más parecido desde {serie.inicio.year} fue el {_f(d)}, "
                          f"con {_mm(v)} L/m².")
        else:
            lineas.append(f"🔁 El día más parecido de la serie: {_f(d)} con {_mm(v)} L/m²")

    anteriores = [(d, v) for d, v in otros if v >= mm and d < fecha]
    if not [x for x in otros if x[1] >= mm]:
        lineas.append(f"🏆 <b>¡El día más lluvioso desde {serie.inicio.year}!</b>")
    elif anteriores:
        d, v = anteriores[-1]
        if (fecha - d).days > 60:
            if publico:
                lineas.append(f"📈 Es la mayor lluvia en un solo día desde el {_f(d)}, "
                              f"cuando cayeron {_mm(v)} L/m².")
            else:
                lineas.append(f"📈 No llovía tanto en un día desde el {_f(d)} ({_mm(v)} L/m²)")
    puesto = serie.ranking_dia(mm, excluir=fecha)
    if puesto <= 100:
        if publico:
            lineas.append(f"🥇 Es el día número {puesto} con más lluvia de todos los "
                          f"registrados desde {serie.inicio.year}.")
        else:
            lineas.append(f"🥇 Puesto {puesto} entre los días más lluviosos desde {serie.inicio.year}")
    return lineas


def _mes_parecido(serie, total, a, m):
    completos = serie.meses_completos(excluir=(a, m))
    if not completos:
        return None
    tol = max(2.0, total * 0.1)
    mismos = [(k, p) for k, p in completos.items() if k[1] == m and abs(p["total"] - total) <= tol]
    if mismos:
        return min(mismos, key=lambda x: abs(x[1]["total"] - total))
    return min(completos.items(), key=lambda x: abs(x[1]["total"] - total))


def comparar_mes(serie, total, a, m, en_curso=False, publico=False):
    """Compara la lluvia de un mes (terminado o en curso) con la serie."""
    lineas = []
    nombre = NOMBRES_MES[m]
    media, n = serie.media_mes(m, excluir=(a, m))
    if media:
        if en_curso:
            lineas.append(f"📊 {nombre.capitalize()} lleva el {_pct(total, media)} de lo que llueve "
                          f"de media en todo el mes ({_mm(media)} L/m², {n} años)")
        else:
            lineas.append(f"📊 Media de {nombre}: {_mm(media)} L/m² ({n} años) → "
                          f"este ha sido el {_pct(total, media)} de lo normal")
    completos = serie.meses_completos(m, excluir=(a, m))
    if completos:
        superados = sum(1 for p in completos.values() if total > p["total"])
        if en_curso:
            if superados and publico:
                lineas.append(f"📈 En lo que va de mes ya ha llovido más que en {superados} de los "
                              f"{len(completos)} meses de {nombre} registrados desde {serie.inicio.year}.")
            elif superados:
                lineas.append(f"📈 Ya ha llovido más que en {superados} de los {len(completos)} "
                              f"{plural_mes(m)} registrados")
        else:
            puesto = len(completos) - superados + 1
            if superados == len(completos):
                lineas.append(f"🏆 <b>¡El {nombre} más lluvioso desde {serie.inicio.year}!</b>")
            elif total < min(p["total"] for p in completos.values()):
                lineas.append(f"🏜 El {nombre} más seco de la serie")
            else:
                lineas.append(f"🏅 Puesto {puesto} de {len(completos) + 1} {plural_mes(m)} más lluviosos")
    if total >= 1:
        parecido = _mes_parecido(serie, total, a, m)
        if parecido:
            (pa, pm), p = parecido
            if publico and en_curso:
                lineas.append(f"🔁 Lo que llevamos de mes ya es parecido a lo que llovió en todo "
                              f"{NOMBRES_MES[pm]} de {pa} ({_mm(p['total'])} L/m²).")
            elif publico:
                lineas.append(f"🔁 Un mes parecido fue {NOMBRES_MES[pm]} de {pa}, "
                              f"con {_mm(p['total'])} L/m².")
            else:
                lineas.append(f"🔁 Un mes parecido: {NOMBRES_MES[pm]} de {pa} con {_mm(p['total'])} L/m²")
    return lineas


def comparar_hidro(serie, total, fecha, corto=False, publico=False):
    media, n = serie.media_hidro_a_fecha(fecha)
    if not media:
        return []
    if publico:
        pct = total / media * 100
        estado = ("por debajo de lo normal" if pct < 90 else
                  "por encima de lo normal" if pct > 110 else "dentro de lo normal")
        return [f"🪣 Desde el 1 de septiembre (inicio del año hidrológico) lo normal a "
                f"estas alturas es haber recogido unos {_mm(media)} L/m². Llevamos el "
                f"{pct:.0f}% de esa cantidad: vamos {estado}."]
    if corto:  # cuando el total ya se ha dicho en el mismo mensaje
        return [f"🪣 A {fecha.day} de {NOMBRES_MES[fecha.month]} lo normal es llevar "
                f"{_mm(media)} L/m² en el año hidrológico (vamos al {_pct(total, media)})"]
    ini = inicio_hidro(fecha)
    return [f"🪣 Año hidrológico {etiqueta_hidro(ini.year)}: {_mm(total)} L/m² — "
            f"a {fecha.day} de {NOMBRES_MES[fecha.month]} lo normal es llevar {_mm(media)} L/m² "
            f"(vamos al {_pct(total, media)})"]


def cierre_periodo(serie, a, hidro=False):
    """Comparativa al terminar un año natural (dic) o hidrológico (ago)."""
    grupos = serie.hidros() if hidro else serie.anios()
    p = grupos.get(a)
    if not p:
        return []
    nombre = f"año hidrológico {etiqueta_hidro(a)}" if hidro else f"año {a}"
    media, n = serie.media_hidro() if hidro else serie.media_anual()
    lineas = [f"🎆 <b>Cierre del {nombre}: {_mm(p['total'])} L/m²</b> en {p['lluviosos']} días de lluvia"]
    if media:
        lineas.append(f"📊 Media ({n} años completos): {_mm(media)} L/m² → {_pct(p['total'], media)}")
    completos = sorted((x["total"] for k, x in grupos.items() if x["completo"] and k != a), reverse=True)
    if p["completo"] and completos:
        puesto = 1 + sum(1 for t in completos if t > p["total"])
        lineas.append(f"🏅 Puesto {puesto} de {len(completos) + 1}")
    return lineas


# ---------------------------------------------------------------------------
# Consultas bajo demanda (/lluvia, /efemerides)
# ---------------------------------------------------------------------------
def mensaje_dia(serie, d, hoy=None):
    hoy = hoy or date.today()
    if d >= hoy:
        return "El día de hoy aún no ha terminado: usa /hoy para ver lo que lleva."
    if not serie.inicio or d < serie.inicio:
        return f"La serie empieza el {_f(serie.inicio)}." if serie.inicio else "No hay datos de lluvia."
    v = serie.get(d)
    titulo = f"🌧 <b>Lluvia del {DIAS_SEMANA[d.weekday()]} {_f(d)}</b>"
    if v is None:
        return f"{titulo}\n\nNo hay dato de ese día (hueco en la serie)."
    lineas = [titulo, ""]
    if v > 0:
        lineas.append(f"Cayeron <b>{_mm(v)} L/m²</b>{detalle_estacion(serie.detalle.get(d))}")
    else:
        lineas.append("☀️ No llovió.")
    if d in serie.notas:
        lineas.append(f"📝 {serie.notas[d]}")

    mes = serie.meses().get((d.year, d.month))
    if mes:
        lineas += ["", f"🗓 {NOMBRES_MES[d.month].capitalize()} de {d.year}: {_mm(mes['total'])} L/m² "
                       f"en {mes['lluviosos']} días de lluvia"]
    if v > 0:
        puesto = serie.ranking_dia(v)
        if puesto <= 100:
            lineas.append(f"🥇 Puesto {puesto} entre los días más lluviosos de la serie")
    # Mismo día en otros años
    mismos = [(a, serie.get(date(a, d.month, d.day)))
              for a in range(serie.inicio.year, serie.fin.year + 1)
              if a != d.year and _existe(a, d.month, d.day)]
    mismos = [(a, x) for a, x in mismos if x is not None]
    if mismos:
        llovio = [(a, x) for a, x in mismos if x > 0]
        txt = f"📅 Un {d.day} de {NOMBRES_MES[d.month]} ha llovido en {len(llovio)} de {len(mismos)} años más"
        if llovio:
            a, x = max(llovio, key=lambda t: t[1])
            txt += f"; el que más, {a} ({_mm(x)} L/m²)"
        lineas.append(txt)
    return "\n".join(lineas)


def _existe(a, m, d):
    return d <= _ultimo_dia(a, m)


def mensaje_mes(serie, a, m):
    """Total del mes, comparativa y lista de días de lluvia."""
    p = serie.meses().get((a, m))
    nombre = f"{NOMBRES_MES[m].capitalize()} {a}"
    if not p or p["con_dato"] == 0:
        return f"No hay datos de lluvia de {NOMBRES_MES[m]} {a}."
    lineas = [f"🌧 <b>Lluvia — {nombre}</b>", ""]
    n = p["lluviosos"]
    lineas.append(f"Total: <b>{_mm(p['total'])} L/m²</b> en {n} {'día' if n == 1 else 'días'} de lluvia")
    aviso = _aviso_incompleto(p, serie.fin)
    if aviso:
        lineas.append(aviso)
    en_curso = p["hasta"] > serie.fin
    if p["completo"] or en_curso:
        lineas += comparar_mes(serie, p["total"], a, m, en_curso=en_curso)

    dias = [(d, v) for d, v in serie.conocidos if d.year == a and d.month == m and v > 0]
    if dias:
        lineas += ["", "<b>Días de lluvia:</b>"]
        maximo = max(v for _, v in dias)
        for d, v in dias:
            nota = f" 📝 {serie.notas[d]}" if d in serie.notas else ""
            marca = " ⭐" if v == maximo and len(dias) > 1 else ""
            lineas.append(f"• {d.strftime('%d/%m')}: <b>{_mm(v)}</b>{marca}"
                          f"{detalle_estacion(serie.detalle.get(d))}{nota}")
    return "\n".join(lineas)


def _bloque_meses(serie, claves):
    meses = serie.meses()
    datos = [(k, meses.get(k)) for k in claves]
    maximo = max((p["total"] for _, p in datos if p and p["con_dato"]), default=0)
    lineas = []
    for (a, m), p in datos:
        if not p or p["con_dato"] == 0:
            lineas.append(f"<code>{ABREV_MES[m]}</code>  sin datos")
            continue
        marca = "" if p["completo"] else " *"
        lineas.append(f"<code>{ABREV_MES[m]} {_barra(p['total'], maximo)}</code> {_mm(p['total'])}{marca}")
    return lineas


def _resumen_periodo(serie, p, titulo, claves_meses, grupos, media, n_media):
    if not p or p["con_dato"] == 0:
        return None
    lineas = [titulo, ""]
    lineas.append(f"Total: <b>{_mm(p['total'])} L/m²</b> en {p['lluviosos']} días de lluvia")
    aviso = _aviso_incompleto(p, serie.fin)
    if aviso:
        lineas.append(aviso)
    if media:
        lineas.append(f"📊 Media ({n_media} años completos): {_mm(media)} L/m² → {_pct(p['total'], media)}")
    completos = {k: x for k, x in grupos.items() if x["completo"]}
    if p["completo"] and len(completos) > 1:
        orden = sorted(completos.values(), key=lambda x: -x["total"])
        puesto = 1 + sum(1 for x in orden if x["total"] > p["total"])
        extra = " — ¡el más lluvioso!" if puesto == 1 else (" — el más seco" if puesto == len(orden) else "")
        lineas.append(f"🏅 Puesto {puesto} de {len(orden)}{extra}")
    if p["max"]:
        v, d = p["max"]
        lineas.append(f"🌊 Día más lluvioso: {_f(d)} ({_mm(v)} L/m²)")
    meses = serie.meses()
    con_dato = [(k, meses[k]) for k in claves_meses if k in meses and meses[k]["completo"]]
    if con_dato:
        k, x = max(con_dato, key=lambda t: t[1]["total"])
        lineas.append(f"📅 Mes más lluvioso: {NOMBRES_MES[k[1]]} ({_mm(x['total'])} L/m²)")
    lineas += [""] + _bloque_meses(serie, claves_meses)
    if any(k in meses and not meses[k]["completo"] and meses[k]["con_dato"] for k in claves_meses):
        lineas.append("<i>* mes incompleto</i>")
    return lineas


def mensaje_anio(serie, a):
    p = serie.anios().get(a)
    media, n = serie.media_anual()
    lineas = _resumen_periodo(serie, p, f"🌧 <b>Lluvia del año {a}</b>",
                              [(a, m) for m in range(1, 13)], serie.anios(), media, n)
    if not lineas:
        return f"No hay datos de lluvia de {a}.", None
    return "\n".join(lineas), teclado_meses_lluvia([(a, m) for m in range(1, 13)], serie)


def mensaje_hidro(serie, a):
    p = serie.hidros().get(a)
    media, n = serie.media_hidro()
    claves = [(a, m) for m in range(9, 13)] + [(a + 1, m) for m in range(1, 9)]
    lineas = _resumen_periodo(serie, p, f"🪣 <b>Año hidrológico {etiqueta_hidro(a)}</b> "
                                        f"(1/9/{a} – 31/8/{a + 1})",
                              claves, serie.hidros(), media, n)
    if not lineas:
        return f"No hay datos del año hidrológico {etiqueta_hidro(a)}.", None
    return "\n".join(lineas), teclado_meses_lluvia(claves, serie)


def teclado_meses_lluvia(claves, serie):
    meses = serie.meses()
    botones = [{"text": f"{ABREV_MES[m]} {str(a)[-2:]}", "callback_data": f"lm:{a}:{m}"}
               for a, m in claves if (a, m) in meses and meses[(a, m)]["con_dato"]]
    if not botones:
        return None
    return {"inline_keyboard": [botones[i:i + 4] for i in range(0, len(botones), 4)]}


def mensaje_general(serie, hoy=None):
    hoy = hoy or date.today()
    if not serie.inicio:
        return "No hay datos de lluvia."
    fin = serie.fin
    lineas = [f"🌧 <b>Lluvia — serie desde el {_f(serie.inicio)}</b>",
              f"<i>Datos hasta el {_f(fin)}</i>", ""]
    ini_h = inicio_hidro(fin)
    tot_h = serie.periodo(ini_h, fin)["total"]
    lineas += comparar_hidro(serie, tot_h, fin)
    p_mes = serie.periodo(fin.replace(day=1), fin)
    media_m, _ = serie.media_mes(fin.month, excluir=(fin.year, fin.month))
    txt = f"🗓 {NOMBRES_MES[fin.month].capitalize()} {fin.year}: {_mm(p_mes['total'])} L/m²"
    if media_m:
        txt += f" (media del mes completo: {_mm(media_m)})"
    lineas.append(txt)
    p_anio = serie.periodo(date(fin.year, 1, 1), fin)
    lineas.append(f"📅 Año {fin.year}: {_mm(p_anio['total'])} L/m²")
    seca = serie.racha_seca_actual()
    if seca:
        lineas.append(f"☀️ Días seguidos sin llover: {seca}")
    lineas += ["", "🏆 <b>Récords</b>"] + lineas_records(serie)
    lineas += ["", "<b>Consultas</b>",
               "/lluvia 2000 — un año",
               "/lluvia 2000-01 — año hidrológico",
               "/lluvia marzo 2000 — un mes",
               "/lluvia 15/03/2000 — un día (también: ayer)",
               "/lluvia top — rankings",
               "/efemerides — tal día como hoy",
               "/curiosidad — un dato curioso"]
    return "\n".join(lineas)


def lineas_records(serie):
    lineas = []
    if serie.conocidos:
        d, v = max(serie.conocidos, key=lambda x: x[1])
        lineas.append(f"🌊 Día más lluvioso: <b>{_mm(v)} L/m²</b> ({_f(d)})")
    meses = serie.meses_completos()
    if meses:
        (a, m), p = max(meses.items(), key=lambda x: x[1]["total"])
        lineas.append(f"🗓 Mes más lluvioso: <b>{_mm(p['total'])} L/m²</b> ({NOMBRES_MES[m]} {a})")
    anios = {a: p for a, p in serie.anios().items() if p["completo"]}
    if anios:
        a, p = max(anios.items(), key=lambda x: x[1]["total"])
        a2, p2 = min(anios.items(), key=lambda x: x[1]["total"])
        lineas.append(f"📅 Año más lluvioso: <b>{_mm(p['total'])} L/m²</b> ({a})")
        lineas.append(f"🏜 Año más seco: <b>{_mm(p2['total'])} L/m²</b> ({a2})")
    secas = serie.rachas(seca=True)
    if secas:
        n, d1, d2 = max(secas)
        lineas.append(f"☀️ Racha más larga sin llover: <b>{n} días</b> ({_f(d1)} – {_f(d2)})")
    return lineas


def mensaje_ranking(serie):
    if not serie.conocidos:
        return "No hay datos de lluvia."
    lineas = [f"🏆 <b>Rankings de lluvia desde {serie.inicio.year}</b>", "", "<b>Días más lluviosos</b>"]
    medallas = ["🥇", "🥈", "🥉"]
    for i, (d, v) in enumerate(sorted(serie.conocidos, key=lambda x: -x[1])[:10]):
        lineas.append(f"{medallas[i] if i < 3 else f'{i + 1}.'} {_f(d)}: <b>{_mm(v)}</b>")
    meses = sorted(serie.meses_completos().items(), key=lambda x: -x[1]["total"])
    lineas += ["", "<b>Meses más lluviosos</b>"]
    for i, ((a, m), p) in enumerate(meses[:5]):
        lineas.append(f"{i + 1}. {NOMBRES_MES[m].capitalize()} {a}: <b>{_mm(p['total'])}</b>")
    anios = sorted(((a, p) for a, p in serie.anios().items() if p["completo"]), key=lambda x: -x[1]["total"])
    if anios:
        lineas += ["", "<b>Años más lluviosos</b>"]
        lineas += [f"{i + 1}. {a}: <b>{_mm(p['total'])}</b>" for i, (a, p) in enumerate(anios[:5])]
        lineas += ["", "<b>Años más secos</b>"]
        lineas += [f"{i + 1}. {a}: <b>{_mm(p['total'])}</b>" for i, (a, p) in enumerate(anios[::-1][:5])]
    hidros = sorted(((a, p) for a, p in serie.hidros().items() if p["completo"]), key=lambda x: -x[1]["total"])
    if hidros:
        lineas += ["", "<b>Años hidrológicos</b>",
                   f"Más lluvioso: {etiqueta_hidro(hidros[0][0])} ({_mm(hidros[0][1]['total'])})",
                   f"Más seco: {etiqueta_hidro(hidros[-1][0])} ({_mm(hidros[-1][1]['total'])})"]
    secas = sorted(serie.rachas(seca=True), reverse=True)[:3]
    lluviosas = sorted(serie.rachas(seca=False), reverse=True)[:1]
    if secas:
        lineas += ["", "<b>Rachas</b>"]
        for n, d1, d2 in secas:
            lineas.append(f"☀️ {n} días sin llover ({_f(d1)} – {_f(d2)})")
        for n, d1, d2 in lluviosas:
            lineas.append(f"🌧 {n} días seguidos lloviendo ({_f(d1)} – {_f(d2)})")
    huecos = [h for h in serie.huecos() if (h[1] - h[0]).days >= 0]
    if huecos:
        lineas += ["", "<i>Sin datos: " + ", ".join(f"{_f(a)} – {_f(b)}" for a, b in huecos) + "</i>"]
    return "\n".join(lineas)


def mensaje_efemerides(serie, dia, mes):
    if not serie.inicio:
        return "No hay datos de lluvia."
    if not _existe(2000, mes, dia):
        return "Esa fecha no existe."
    datos = []
    for a in range(serie.inicio.year, serie.fin.year + 1):
        if _existe(a, mes, dia):
            v = serie.get(date(a, mes, dia))
            if v is not None:
                datos.append((a, v))
    titulo = f"📅 <b>Un {dia} de {NOMBRES_MES[mes]} en la serie</b>"
    if not datos:
        return f"{titulo}\n\nNo hay datos de ese día."
    llovio = [(a, v) for a, v in datos if v > 0]
    lineas = [titulo, "", f"Ha llovido en <b>{len(llovio)} de {len(datos)}</b> años"]
    if llovio:
        media = sum(v for _, v in llovio) / len(llovio)
        lineas.append(f"Cuando llueve, caen de media {_mm(media)} L/m²")
        lineas.append("")
        maximo = max(v for _, v in llovio)
        for a, v in llovio:
            nota = f" 📝 {serie.notas[date(a, mes, dia)]}" if date(a, mes, dia) in serie.notas else ""
            lineas.append(f"• {a}: <b>{_mm(v)}</b>{' ⭐' if v == maximo else ''}{nota}")
    # Lo más fuerte en ±3 días
    cerca = []
    for a in range(serie.inicio.year, serie.fin.year + 1):
        if not _existe(a, mes, dia):
            continue
        base = date(a, mes, dia)
        for k in range(-3, 4):
            d = base + timedelta(days=k)
            v = serie.get(d)
            if k and v:
                cerca.append((v, d))
    if cerca:
        v, d = max(cerca)
        lineas += ["", f"🌊 En los días de alrededor, lo más fuerte: {_f(d)} con {_mm(v)} L/m²"]
    return "\n".join(lineas)


# ---------------------------------------------------------------------------
# Curiosidades aleatorias
# ---------------------------------------------------------------------------
def _cur_tal_dia(s, hoy, rnd):
    datos = [(a, s.get(date(a, hoy.month, hoy.day))) for a in range(s.inicio.year, hoy.year)
             if _existe(a, hoy.month, hoy.day)]
    datos = [(a, v) for a, v in datos if v is not None]
    if len(datos) < 5:
        return None
    llovio = [(a, v) for a, v in datos if v > 0]
    fecha_txt = f"{hoy.day} de {NOMBRES_MES[hoy.month]}"
    if not llovio:
        return (f"tal_dia-{hoy:%m-%d}",
                f"📅 En {len(datos)} años de registros, <b>nunca ha llovido un {fecha_txt}</b>.")
    a, v = max(llovio, key=lambda x: x[1])
    return (f"tal_dia-{hoy:%m-%d}",
            f"📅 Tal día como hoy, en <b>{a}</b>, cayeron <b>{_mm(v)} L/m²</b>. "
            f"En {len(datos)} años ha llovido {len(llovio)} "
            f"{'vez' if len(llovio) == 1 else 'veces'} un {fecha_txt}.")


def _cur_semana(s, hoy, rnd):
    cand = []
    for a in range(s.inicio.year, hoy.year):
        if not _existe(a, hoy.month, hoy.day):
            continue
        base = date(a, hoy.month, hoy.day)
        for k in range(1, 8):
            v = s.get(base + timedelta(days=k))
            if v:
                cand.append((v, base + timedelta(days=k)))
    if not cand:
        return None
    v, d = max(cand)
    return (f"semana-{hoy:%m-%d}",
            f"🔮 Mirando atrás: en la semana que viene de otros años, el día más lluvioso fue "
            f"el <b>{_f(d)}</b> con <b>{_mm(v)} L/m²</b>.")


def _cur_racha_actual(s, hoy, rnd):
    n = s.racha_seca_actual()
    if n < 10:
        return None
    secas = s.rachas(seca=True)
    larga, d1, d2 = max(secas)
    superadas = sum(1 for r in secas if r[0] < n)
    return (f"racha_actual-{n // 10}",
            f"☀️ Llevamos <b>{n} días sin llover</b>. Más largo que el {superadas * 100 // len(secas)}% "
            f"de las rachas secas desde {s.inicio.year}. El récord: {larga} días "
            f"({_f(d1)} – {_f(d2)}).")


def _cur_rachas(s, hoy, rnd):
    if rnd.random() < 0.5:
        n, d1, d2 = max(s.rachas(seca=True))
        return ("racha_seca", f"☀️ La racha más larga sin llover duró <b>{n} días</b>: "
                              f"del {_f(d1)} al {_f(d2)}.")
    n, d1, d2 = max(s.rachas(seca=False))
    return ("racha_lluvia", f"🌧 La racha más larga lloviendo fue de <b>{n} días seguidos</b>: "
                            f"del {_f(d1)} al {_f(d2)}.")


def _cur_top_dia(s, hoy, rnd):
    top = sorted(s.conocidos, key=lambda x: -x[1])[:10]
    i = rnd.randrange(len(top))
    d, v = top[i]
    puesto = "el día más lluvioso" if i == 0 else f"el {i + 1}º día más lluvioso"
    hace = hoy.year - d.year
    return (f"top-{d}", f"🌊 El {DIAS_SEMANA[d.weekday()]} <b>{_f(d)}</b> cayeron <b>{_mm(v)} L/m²</b>, "
                        f"{puesto} desde {s.inicio.year}" + (f" (hace {hace} años)." if hace > 1 else "."))


def _cur_anios(s, hoy, rnd):
    anios = {a: p for a, p in s.anios().items() if p["completo"]}
    if len(anios) < 3:
        return None
    a1, p1 = max(anios.items(), key=lambda x: x[1]["total"])
    a2, p2 = min(anios.items(), key=lambda x: x[1]["total"])
    veces = p1["total"] / p2["total"] if p2["total"] else 0
    return ("anios", f"📅 En <b>{a1}</b> cayeron <b>{_mm(p1['total'])} L/m²</b>, "
                     f"{veces:.1f} veces más que en {a2}, el año más seco ({_mm(p2['total'])} L/m²).")


def _cur_mes_historia(s, hoy, rnd):
    m = hoy.month
    meses = s.meses_completos(m)
    if len(meses) < 5:
        return None
    (a1, _), p1 = max(meses.items(), key=lambda x: x[1]["total"])
    (a2, _), p2 = min(meses.items(), key=lambda x: x[1]["total"])
    media, n = s.media_mes(m)
    return (f"mes-{m}", f"🗓 Un {NOMBRES_MES[m]} normal deja <b>{_mm(media)} L/m²</b>. "
                        f"El más lluvioso fue el de {a1} ({_mm(p1['total'])}) y el más seco "
                        f"el de {a2} ({_mm(p2['total'])}).")


def _cur_total(s, hoy, rnd):
    total = sum(v for _, v in s.conocidos)
    tejado = total * 100  # litros en 100 m²
    return ("total", f"💧 Desde {s.inicio.year} se han medido <b>{_miles(total)} L/m²</b>: "
                     f"una columna de agua de {total / 1000:.1f} metros. En un tejado de 100 m² "
                     f"serían {_miles(tejado / 1000)} m³ de agua, unas {_miles(tejado / 150)} "
                     f"bañeras llenas.")


def _miles(n):
    return f"{n:,.0f}".replace(",", ".")


def _cur_hace_anios(s, hoy, rnd):
    opciones = []
    for n in (5, 10, 15, 20, 25, 30):
        a = hoy.year - n
        p = s.meses().get((a, hoy.month))
        if p and p["completo"]:
            opciones.append((n, a, p))
    if not opciones:
        return None
    n, a, p = rnd.choice(opciones)
    media, _ = s.media_mes(hoy.month)
    comp = ""
    if media:
        comp = " (más de lo normal)" if p["total"] > media * 1.2 else (
            " (menos de lo normal)" if p["total"] < media * 0.8 else " (lo normal)")
    txt = (f"⏳ Hace {n} años, en {NOMBRES_MES[hoy.month]} de {a}, cayeron "
           f"<b>{_mm(p['total'])} L/m²</b> en {p['lluviosos']} días de lluvia{comp}.")
    if p["max"]:
        txt += f" El día más fuerte, el {p['max'][1].strftime('%d/%m')} ({_mm(p['max'][0])})."
    return (f"hace-{n}-{hoy.month}", txt)


def _cur_decadas(s, hoy, rnd):
    grupos = {}
    for a, p in s.anios().items():
        if p["completo"]:
            grupos.setdefault(a // 10 * 10, []).append(p["total"])
    grupos = {k: v for k, v in grupos.items() if len(v) >= 2}
    if len(grupos) < 2:
        return None
    partes = [f"{str(k)[-2:]}s: <b>{_mm(sum(v) / len(v))}</b>" for k, v in sorted(grupos.items())]
    return ("decadas", "📆 Lluvia media anual por décadas (años completos): " + " · ".join(partes) + " L/m².")


def _cur_fuertes(s, hoy, rnd):
    fuertes = [(d, v) for d, v in s.conocidos if v >= 50]
    if not fuertes:
        return None
    d, v = fuertes[-1]
    return ("fuertes", f"⛈ Desde {s.inicio.year} ha habido <b>{len(fuertes)} días con 50 L/m² o más</b>. "
                       f"El último, el {_f(d)} ({_mm(v)} L/m²).")


def _cur_hidro(s, hoy, rnd):
    fin = s.fin
    total = s.periodo(inicio_hidro(fin), fin)["total"]
    lineas = comparar_hidro(s, total, fin)
    if not lineas:
        return None
    return (f"hidro-{fin.isocalendar()[1]}", lineas[0])


def _cur_climatologia(s, hoy, rnd):
    medias = [(m, s.media_mes(m)[0]) for m in range(1, 13)]
    medias = [(m, v) for m, v in medias if v is not None]
    if len(medias) < 12:
        return None
    m1, v1 = max(medias, key=lambda x: x[1])
    m2, v2 = min(medias, key=lambda x: x[1])
    seco = (f"el más seco <b>{NOMBRES_MES[m2]}</b>, que ronda los {_mm(v2)} L/m²" if v2 >= 0.05
            else f"el más seco <b>{NOMBRES_MES[m2]}</b>, en el que prácticamente no llueve nunca")
    return ("climatologia", f"📊 El mes más lluvioso de media es <b>{NOMBRES_MES[m1]}</b> "
                            f"({_mm(v1)} L/m²) y {seco}.")


def _cur_meses_secos(s, hoy, rnd):
    """Meses del año en los que nunca ha caído ni una gota."""
    nunca = []
    for m in range(1, 13):
        completos = s.meses_completos(m)
        if len(completos) >= 10 and all(p["total"] == 0 for p in completos.values()):
            nunca.append((m, len(completos)))
    if not nunca:
        return None
    m, n = rnd.choice(nunca)
    return (f"nunca-{m}", f"🏜 En los {n} {plural_mes(m)} registrados desde {s.inicio.year} "
                          f"<b>no ha llovido ni un solo día</b>.")


def _cur_frecuencia(s, hoy, rnd):
    total = len(s.conocidos)
    lluvia = sum(1 for _, v in s.conocidos if v > 0)
    anios = [p["lluviosos"] for p in s.anios().values() if p["completo"]]
    txt = f"🎲 Llueve <b>{lluvia / total * 100:.0f}% de los días</b>"
    if anios:
        txt += f": unos {sum(anios) / len(anios):.0f} días de lluvia al año"
    return ("frecuencia", txt + ".")


def _cur_dia_semana(s, hoy, rnd):
    tot = [0.0] * 7
    for d, v in s.conocidos:
        tot[d.weekday()] += v
    i = max(range(7), key=lambda k: tot[k])
    j = min(range(7), key=lambda k: tot[k])
    finde = (tot[5] + tot[6]) / 2
    entre = sum(tot[:5]) / 5
    return ("dia_semana", f"🗓 ¿Llueve más en fin de semana? Desde {s.inicio.year} el día que más agua "
                          f"ha recogido es el <b>{DIAS_SEMANA[i]}</b> y el que menos el {DIAS_SEMANA[j]}. "
                          f"Un sábado/domingo medio recoge el {finde / entre * 100:.0f}% que un día laborable.")


def _cur_estaciones(s, hoy, rnd):
    nombres = {"invierno": (12, 1, 2), "primavera": (3, 4, 5), "verano": (6, 7, 8), "otoño": (9, 10, 11)}
    medias = {}
    for nombre, meses in nombres.items():
        vals = [s.media_mes(m)[0] for m in meses]
        if None in vals:
            return None
        medias[nombre] = sum(vals)
    total = sum(medias.values())
    partes = [f"{k} {v / total * 100:.0f}%" for k, v in sorted(medias.items(), key=lambda x: -x[1])]
    return ("estaciones", "🍂 Reparto de la lluvia del año: " + " · ".join(partes) + ".")


def _cur_nota(s, hoy, rnd):
    if not s.notas:
        return None
    d = rnd.choice(sorted(s.notas))
    v = s.get(d)
    extra = f" ({_mm(v)} L/m²)" if v else ""
    return (f"nota-{d}", f"📝 El {_f(d)}{extra}: <i>{s.notas[d]}</i>.")


CURIOSIDADES = [
    _cur_tal_dia, _cur_semana, _cur_racha_actual, _cur_rachas, _cur_top_dia, _cur_anios,
    _cur_mes_historia, _cur_total, _cur_hace_anios, _cur_decadas, _cur_fuertes, _cur_hidro,
    _cur_climatologia, _cur_frecuencia, _cur_dia_semana, _cur_estaciones, _cur_nota,
    _cur_meses_secos,
]


def curiosidad_aleatoria(serie, hoy=None, excluir=(), rnd=random):
    """Devuelve (clave, texto) de una curiosidad que no esté en 'excluir'."""
    hoy = hoy or date.today()
    if not serie.conocidos:
        return None
    funciones = CURIOSIDADES[:]
    rnd.shuffle(funciones)
    reserva = None
    for fn in funciones:
        try:
            res = fn(serie, hoy, rnd)
        except Exception:
            continue
        if not res:
            continue
        if res[0] not in excluir:
            return res
        reserva = reserva or res
    return reserva
