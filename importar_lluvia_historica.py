#!/usr/bin/env python3
"""
Importa la lluvia histórica desde el Excel (o CSV) rellenado a mano
====================================================================
Uso:
    python importar_lluvia_historica.py lluvia_historica_1993.xlsx

Lee las columnas Fecha / Lluvia / Observaciones, valida los datos y
(re)escribe lluvia_historica.csv, que es lo que usa el bot. Se puede
volver a ejecutar cuando se rellenen huecos: siempre reemplaza el CSV
entero con el contenido del archivo.

  - Celda vacía = sin dato · 0 = no llovió
  - Acepta decimales con coma o punto
  - Los días desde el inicio de la estación (06/08/2026) se ignoran y se
    conservan los que ya hay en el CSV: esos los rellena el bot cada noche
    desde la hoja Diario de registro_meteo.xlsx
"""

import csv
import sys
from datetime import date
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))
from lluvia_historica import (  # noqa: E402
    HIST_CSV_PATH, _fecha, _num, cargar_serie, leer_historico, escribir_historico,
)

INICIO_ESTACION = date(2026, 8, 6)
MAX_RAZONABLE = 300  # L/m² en un día: por encima se avisa (¿error de tecleo?)


def leer_origen(ruta):
    ruta = Path(ruta)
    if ruta.suffix.lower() in (".xlsx", ".xlsm"):
        from openpyxl import load_workbook
        wb = load_workbook(ruta, read_only=True, data_only=True)
        filas = list(wb.active.iter_rows(values_only=True))
        wb.close()
    else:
        with open(ruta, newline="", encoding="utf-8-sig") as f:
            muestra = f.read(2048)
            f.seek(0)
            sep = ";" if muestra.count(";") > muestra.count(",") else ","
            filas = list(csv.reader(f, delimiter=sep))
    return filas[1:]  # sin cabecera


def main():
    if len(sys.argv) != 2:
        print(__doc__)
        sys.exit(1)
    filas = leer_origen(sys.argv[1])

    datos, errores, avisos = {}, [], []
    for n, fila in enumerate(filas, start=2):
        if not fila or all(c in (None, "") for c in fila):
            continue
        d = _fecha(fila[0])
        if d is None:
            errores.append(f"Fila {n}: fecha no válida ({fila[0]!r})")
            continue
        if d >= INICIO_ESTACION:
            continue
        try:
            mm = _num(fila[1]) if len(fila) > 1 else None
        except ValueError:
            errores.append(f"Fila {n} ({d:%d/%m/%Y}): cantidad no válida ({fila[1]!r})")
            continue
        if mm is not None and mm < 0:
            errores.append(f"Fila {n} ({d:%d/%m/%Y}): valor negativo ({mm})")
            continue
        if mm is not None and mm > MAX_RAZONABLE:
            avisos.append(f"{d:%d/%m/%Y}: {mm} L/m² (¿seguro?)")
        if d in datos:
            errores.append(f"Fila {n}: fecha repetida {d:%d/%m/%Y}")
        nota = str(fila[2]).strip() if len(fila) > 2 and fila[2] not in (None, "") else ""
        datos[d] = (mm, nota)

    if errores:
        print("❌ Hay errores, no se ha escrito nada:")
        print("\n".join(errores[:50]))
        sys.exit(1)

    # Días de la estación ya volcados por el bot: se conservan tal cual
    previos, notas_previas = leer_historico()
    valores = {d: mm for d, (mm, _) in datos.items()}
    notas = {d: nota for d, (_, nota) in datos.items() if nota}
    for d, mm in previos.items():
        if d >= INICIO_ESTACION:
            valores[d] = mm
            if notas_previas.get(d):
                notas[d] = notas_previas[d]
    escribir_historico(valores, notas)

    serie = cargar_serie()
    print(f"✅ {HIST_CSV_PATH.name} escrito: {len(datos)} días "
          f"({serie.inicio:%d/%m/%Y} – {serie.fin:%d/%m/%Y})")
    print(f"   Días con dato: {len(serie.conocidos)} · "
          f"total: {sum(v for _, v in serie.conocidos):.1f} L/m²")
    for a, b in serie.huecos():
        print(f"   Hueco sin datos: {a:%d/%m/%Y} – {b:%d/%m/%Y} ({(b - a).days + 1} días)")
    if avisos:
        print("⚠️ Valores muy altos, revísalos:")
        print("\n".join(avisos))
    print("\nRecuerda subirlo:  git add lluvia_historica.csv && git commit -m 'Lluvia histórica' && git push")


if __name__ == "__main__":
    main()
