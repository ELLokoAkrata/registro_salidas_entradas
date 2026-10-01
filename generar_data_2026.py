# -*- coding: utf-8 -*-
"""
Auditoría de registros 2026 y generación de data complementaria (GENERADA).

Base: registros_all.zip (export de Firestore al 05/08/2026, semanas W12 a W32).

Reglas acordadas:
  - Jornada: lunes a sábado, 8 h diarias exactas ("sin más ni menos").
  - Feriados nacionales de Perú 2026 excluidos (fuente: gob.pe/feriados, RPP).
  - La data generada NUNCA toca Firestore: se entrega en Excels aparte dentro
    de data_generada/, con columna "Origen" que identifica cada fila.

Salidas del script:
  - auditoria_2026.xlsx            -> Resumen + Detalle día a día por trabajador
  - data_generada/registro_YYYY_Wnn_generado.xlsx  -> solo las filas generadas de esa semana
  - data_generada/registro_2026_completo_generado.xlsx -> consolidado real + generado
"""

import hashlib
import io
import re
import zipfile
from datetime import date, datetime, timedelta
from statistics import median

import pandas as pd
from openpyxl.styles import Border, Side

ZIP_PATH = "registros_all.zip"
INICIO = date(2026, 3, 16)   # lunes de la W12, primer día con registros reales
FIN = date(2026, 9, 30)      # último día transcurrido del año
JORNADA = timedelta(hours=8)

# Feriados nacionales de Perú 2026 (aplican aunque caigan domingo; el filtro L-S los ignora)
FERIADOS = {
    date(2026, 1, 1),                                        # Año Nuevo
    date(2026, 4, 2), date(2026, 4, 3),                      # Jueves y Viernes Santo
    date(2026, 5, 1),                                        # Día del Trabajo
    date(2026, 6, 7),                                        # Batalla de Arica
    date(2026, 6, 29),                                       # San Pedro y San Pablo
    date(2026, 7, 23),                                       # Día de la Fuerza Aérea
    date(2026, 7, 28), date(2026, 7, 29),                    # Fiestas Patrias
    date(2026, 8, 6),                                        # Batalla de Junín
    date(2026, 8, 30),                                       # Santa Rosa de Lima
    date(2026, 10, 8),                                       # Combate de Angamos
    date(2026, 11, 1),                                       # Todos los Santos
    date(2026, 12, 8), date(2026, 12, 9),                    # Inmaculada / Ayacucho
    date(2026, 12, 25),                                      # Navidad
}

FMT = "%d/%m/%Y %I:%M:%S %p"
DIAS = ["Lun", "Mar", "Mié", "Jue", "Vie", "Sáb", "Dom"]

ENTRADA_MIN = 7 * 60 + 30   # 07:30
ENTRADA_MAX = 10 * 60 + 45  # 10:45


def fmt_td(td):
    """Timedelta a 'H:MM:SS' (mismo formato que usa la app)."""
    total = int(td.total_seconds())
    return f"{total // 3600}:{(total % 3600) // 60:02d}:{total % 60:02d}"


def cargar_reales():
    """Lee el zip y devuelve dict {(nombre, fecha): fila}, más el set de trabajadores."""
    reales = {}
    trabajadores = []
    with zipfile.ZipFile(ZIP_PATH) as z:
        for name in sorted(z.namelist()):
            if not re.match(r"registro_\d{4}_W\d+\.xlsx", name):
                continue
            df = pd.read_excel(io.BytesIO(z.read(name)), sheet_name="Registros")
            for _, row in df.iterrows():
                fecha = datetime.strptime(str(row["Fecha"]), "%Y-%m-%d").date()
                nombre = str(row["Nombre"]).strip()
                if nombre not in trabajadores:
                    trabajadores.append(nombre)

                def _parse(v):
                    try:
                        return datetime.strptime(str(v), FMT)
                    except Exception:
                        return None

                entrada = _parse(row["Entrada"])
                salida = row["Salida"]
                if pd.isna(salida) or str(salida).strip() in ("", "No marcó salida"):
                    salida_dt = None
                else:
                    salida_dt = _parse(salida)
                reales[(nombre, fecha)] = {"entrada": entrada, "salida": salida_dt}
    return reales, trabajadores


def minuto_aleatorio(nombre, dia, low, high):
    """Segundo pseudoaleatorio estable por (trabajador, fecha) en [low, high]."""
    h = hashlib.md5(f"{nombre}|{dia.isoformat()}".encode()).hexdigest()
    return low + int(h[:8], 16) % (high - low + 1)


def entrada_generada(nombre, dia, bases):
    """Hora de entrada 'natural': base histórica del trabajador + jitter, acotada 07:30-10:45."""
    minuto = int(bases[nombre]) + minuto_aleatorio(nombre, dia, -40, 40)
    minuto = max(ENTRADA_MIN, min(ENTRADA_MAX, minuto))
    dt = datetime.combine(dia, datetime.min.time()) + timedelta(minutes=minuto)
    dt += timedelta(seconds=minuto_aleatorio(nombre, dia + timedelta(days=1), 0, 59))
    return dt


def estilo(ws):
    """Anchos de columna y bordes finos, igual que hace la app."""
    thin = Border(left=Side(style="thin"), right=Side(style="thin"),
                  top=Side(style="thin"), bottom=Side(style="thin"))
    for col_cells in ws.columns:
        letter = col_cells[0].column_letter
        max_len = max((len(str(c.value)) for c in col_cells if c.value is not None), default=0)
        ws.column_dimensions[letter].width = max_len + 2
    for row in ws.iter_rows():
        for cell in row:
            cell.border = thin


def escribir_excel(path, hojas):
    """hojas: lista de (nombre_hoja, DataFrame)."""
    with pd.ExcelWriter(path, engine="openpyxl") as writer:
        for nombre_hoja, df in hojas:
            df.to_excel(writer, sheet_name=nombre_hoja, index=False)
            estilo(writer.sheets[nombre_hoja])


def main():
    reales, trabajadores = cargar_reales()

    # Base histórica de entrada por trabajador (mediana en minutos del día)
    bases = {}
    for nombre in trabajadores:
        minutos = [r["entrada"].hour * 60 + r["entrada"].minute
                   for (n, _), r in reales.items()
                   if n == nombre and r["entrada"]]
        bases[nombre] = median(minutos) if minutos else 8 * 60 + 30

    # ---- Recorrido día por día -------------------------------------------
    detalle = []            # auditoría
    generadas = []          # filas generadas (complemento)
    consolidado = []        # día completo real + generado

    for nombre in trabajadores:
        dia = INICIO
        while dia <= FIN:
            es_laborable = dia.weekday() <= 5          # lunes a sábado
            es_feriado = dia in FERIADOS
            if not es_laborable:
                dia += timedelta(days=1)
                continue

            real = reales.get((nombre, dia))
            fila_base = {"Nombre": nombre, "Fecha": dia.isoformat(),
                         "Día": DIAS[dia.weekday()]}

            if es_feriado and real is None:
                detalle.append({**fila_base, "Estado": "FERIADO",
                                "Entrada": "", "Salida": "", "Horas": "", "Origen": ""})
                dia += timedelta(days=1)
                continue

            if real is None:
                ent = entrada_generada(nombre, dia, bases)
                sal = ent + JORNADA
                origen = "GENERADO_DIA_COMPLETO"
                detalle.append({**fila_base, "Estado": "DÍA GENERADO",
                                "Entrada": ent.strftime(FMT), "Salida": sal.strftime(FMT),
                                "Horas": fmt_td(JORNADA), "Origen": origen})
                generadas.append({**fila_base, "Entrada": ent.strftime(FMT),
                                  "Salida": sal.strftime(FMT),
                                  "Horas Trabajadas": fmt_td(JORNADA), "Origen": origen})
                consolidado.append({**fila_base, "Entrada": ent.strftime(FMT),
                                    "Salida": sal.strftime(FMT),
                                    "Horas Trabajadas": fmt_td(JORNADA), "Origen": origen})
            elif real["entrada"] and real["salida"]:
                horas = real["salida"] - real["entrada"]
                detalle.append({**fila_base, "Estado": "COMPLETO (REAL)",
                                "Entrada": real["entrada"].strftime(FMT),
                                "Salida": real["salida"].strftime(FMT),
                                "Horas": fmt_td(horas), "Origen": "REAL"})
                consolidado.append({**fila_base, "Entrada": real["entrada"].strftime(FMT),
                                    "Salida": real["salida"].strftime(FMT),
                                    "Horas Trabajadas": fmt_td(horas), "Origen": "REAL"})
            elif real["entrada"] and real["salida"] is None:
                sal = real["entrada"] + JORNADA
                detalle.append({**fila_base, "Estado": "SIN SALIDA -> GENERADA",
                                "Entrada": real["entrada"].strftime(FMT),
                                "Salida": sal.strftime(FMT),
                                "Horas": fmt_td(JORNADA), "Origen": "GENERADO_SALIDA"})
                generadas.append({**fila_base, "Entrada": real["entrada"].strftime(FMT),
                                  "Salida": sal.strftime(FMT),
                                  "Horas Trabajadas": fmt_td(JORNADA),
                                  "Origen": "GENERADO_SALIDA"})
                consolidado.append({**fila_base, "Entrada": real["entrada"].strftime(FMT),
                                    "Salida": sal.strftime(FMT),
                                    "Horas Trabajadas": fmt_td(JORNADA),
                                    "Origen": "GENERADO_SALIDA"})
            elif real["salida"] and real["entrada"] is None:
                ent = real["salida"] - JORNADA
                detalle.append({**fila_base, "Estado": "SIN ENTRADA -> GENERADA",
                                "Entrada": ent.strftime(FMT),
                                "Salida": real["salida"].strftime(FMT),
                                "Horas": fmt_td(JORNADA), "Origen": "GENERADO_ENTRADA"})
                generadas.append({**fila_base, "Entrada": ent.strftime(FMT),
                                  "Salida": real["salida"].strftime(FMT),
                                  "Horas Trabajadas": fmt_td(JORNADA),
                                  "Origen": "GENERADO_ENTRADA"})
                consolidado.append({**fila_base, "Entrada": ent.strftime(FMT),
                                    "Salida": real["salida"].strftime(FMT),
                                    "Horas Trabajadas": fmt_td(JORNADA),
                                    "Origen": "GENERADO_ENTRADA"})
            dia += timedelta(days=1)

    det_df = pd.DataFrame(detalle)
    gen_df = pd.DataFrame(generadas)
    con_df = pd.DataFrame(consolidado).sort_values(["Nombre", "Fecha"]).reset_index(drop=True)

    # ---- Resumen de auditoría --------------------------------------------
    resumen = []
    for nombre in trabajadores:
        d = det_df[det_df["Nombre"] == nombre]
        g = gen_df[gen_df["Nombre"] == nombre]
        c = con_df[con_df["Nombre"] == nombre]

        def _td(v):
            h, m, s = map(int, str(v).split(":"))
            return timedelta(hours=h, minutes=m, seconds=s)

        laborables = len(d[d["Estado"] != "FERIADO"])
        horas_reales = sum((_td(v) for v in c[c["Origen"] == "REAL"]["Horas Trabajadas"]),
                           timedelta())
        horas_gen = sum((_td(v) for v in c[c["Origen"] != "REAL"]["Horas Trabajadas"]),
                        timedelta())
        resumen.append({
            "Nombre": nombre,
            "Días laborables (L-S)": laborables,
            "Días completos (REAL)": int((d["Estado"] == "COMPLETO (REAL)").sum()),
            "Salidas generadas": int((d["Estado"] == "SIN SALIDA -> GENERADA").sum()),
            "Entradas generadas": int((d["Estado"] == "SIN ENTRADA -> GENERADA").sum()),
            "Días completos generados": int((d["Estado"] == "DÍA GENERADO").sum()),
            "Feriados excluidos": int((d["Estado"] == "FERIADO").sum()),
            "Horas reales": fmt_td(horas_reales),
            "Horas generadas": fmt_td(horas_gen),
            "Horas totales (real + generado)": fmt_td(horas_reales + horas_gen),
        })
    res_df = pd.DataFrame(resumen)

    import os
    os.makedirs("data_generada", exist_ok=True)

    escribir_excel("auditoria_2026.xlsx",
                   [("Resumen", res_df), ("Detalle", det_df)])
    print("OK auditoria_2026.xlsx")

    # ---- Un excel por semana con solo lo generado -------------------------
    gen_df["semana"] = gen_df["Fecha"].apply(
        lambda f: (lambda iso: f"{iso[0]}_W{iso[1]}")(
            date.fromisoformat(f).isocalendar()))
    for semana, g in gen_df.groupby("semana"):
        g = g.drop(columns=["semana"])
        res = (g.assign(td=g["Horas Trabajadas"].apply(_td))
                 .groupby("Nombre")["td"].sum()
                 .reset_index())
        res.columns = ["Nombre", "Total Horas Generadas"]
        res["Total Horas Generadas"] = res["Total Horas Generadas"].apply(fmt_td)
        res = res.sort_values("Nombre").reset_index(drop=True)
        path = f"data_generada/registro_{semana}_generado.xlsx"
        escribir_excel(path, [("Registros", g), ("Resumen", res)])
        print(f"OK {path} ({len(g)} filas)")

    # ---- Consolidado real + generado --------------------------------------
    resumen_total = c2 = con_df.assign(td=con_df["Horas Trabajadas"].apply(_td))
    resumen_total = (c2.groupby(["Nombre", c2["Origen"].apply(
        lambda o: "REAL" if o == "REAL" else "GENERADO")])["td"].sum()
        .reset_index().rename(columns={"Origen": "Tipo", "td": "Horas"}))
    resumen_total["Horas"] = resumen_total["Horas"].apply(fmt_td)
    pivot = resumen_total.pivot(index="Nombre", columns="Tipo", values="Horas")
    for col in ("REAL", "GENERADO"):
        if col not in pivot.columns:
            pivot[col] = "0:00:00"
    pivot = pivot.reset_index()[["Nombre", "REAL", "GENERADO"]]
    pivot.columns = ["Nombre", "Horas reales", "Horas generadas"]

    path = "data_generada/registro_2026_completo_generado.xlsx"
    escribir_excel(path, [("Registros", con_df), ("Resumen", pivot)])
    print(f"OK {path} ({len(con_df)} filas)")

    # ---- Consola -----------------------------------------------------------
    print("\n== RESUMEN ==")
    print(res_df.to_string(index=False))


if __name__ == "__main__":
    main()
