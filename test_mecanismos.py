# -*- coding: utf-8 -*-
"""
Prueba offline del fix de zona horaria y del mecanismo anti-olvidos.
NO toca Firestore: se stubbean streamlit y firebase_admin al importar la app.
Ejecutar:  python test_mecanismos.py
"""
import sys
from unittest import mock

import pandas as pd
from datetime import datetime

SECRETS = {
    "firebase": {k: "x" for k in [
        "type", "project_id", "private_key_id", "private_key", "client_email",
        "client_id", "auth_uri", "token_uri", "auth_provider_x509_cert_url",
        "client_x509_cert_url"]},
    "user_passwords": {},
}

fake_st = mock.MagicMock()
fake_st.secrets = SECRETS
fake_fb = mock.MagicMock()
fake_fb._apps = {}
fake_fb.credentials = mock.MagicMock()
fake_fb.firestore = mock.MagicMock()

sys.modules["streamlit"] = fake_st
sys.modules["firebase_admin"] = fake_fb

import app_registros as app  # noqa: E402

FALLOS = []

def check(cond, msg):
    print(("OK  " if cond else "FALLA") + " - " + msg)
    if not cond:
        FALLOS.append(msg)

# ---- Congelar el 'ahora': 1 de octubre de 2026, 9:00 AM Lima ----
app.ahora_lima = lambda: datetime(2026, 10, 1, 9, 0)

# 1) Semana y feriados según hora de Lima
check(app.get_week_filename() == "registro_2026_W40.xlsx", "get_week_filename() usa hora de Lima (W40)")
check(app.es_dia_laborable(datetime(2026, 8, 6).date()) is False, "feriado 6 de agosto excluido")
check(app.es_dia_laborable(datetime(2026, 9, 26).date()) is True, "sábado es día laborable")
check(app.es_dia_laborable(datetime(2026, 9, 27).date()) is False, "domingo no es día laborable")

# 2) Anti-olvidos con datos de prueba
cols = ["Nombre", "Fecha", "Entrada", "Salida", "Horas Trabajadas", "Origen"]
w40 = pd.DataFrame([
    ["Nelida Ruiz", "2026-09-28", "28/09/2026 08:16:31 AM", "No marcó salida", "No marcó salida", "REAL"],
    ["Paula Lecaros", "2026-09-29", "29/09/2026 09:02:00 AM", "29/09/2026 05:02:00 PM", "8:00:00", "REAL"],
], columns=cols)
w39 = pd.DataFrame(columns=cols)  # semana anterior sin ningún registro
data = {"registro_2026_W40.xlsx": w40, "registro_2026_W39.xlsx": w39}
saved = {}
app.load_week_data = lambda f: data.get(f, pd.DataFrame(columns=cols)).copy()
app.save_week_data_and_upload = lambda df, f: saved.__setitem__(f, df.copy())

out = app.completar_olvidados()

w39_out = saved.get("registro_2026_W39.xlsx")
w40_out = saved.get("registro_2026_W40.xlsx")
gen40 = w40_out[w40_out["Origen"] == "GENERADO_OLVIDO"]

check(len(w39_out) == 18, f"W39 vacía -> 18 días generados (6 laborables x 3 trabajadores), hay {0 if w39_out is None else len(w39_out)}")
check(len(w40_out) == 9, f"W40 -> 2 reales + 7 días completos generados (9-2 existentes), hay {len(w40_out)}")
check(len(out) == 26, f"total completados reportados = 26, hay {len(out)}")

nel28 = w40_out[(w40_out["Nombre"] == "Nelida Ruiz") & (w40_out["Fecha"] == "2026-09-28")].iloc[0]
check(nel28["Salida"] == "28/09/2026 04:16:31 PM", f"salida de Nelida = entrada + 8h exactas ({nel28['Salida']})")
check(nel28["Origen"] == "GENERADO_OLVIDO", "fila tocada por anti-olvidos queda con Origen=GENERADO_OLVIDO")

paula29 = w40_out[(w40_out["Nombre"] == "Paula Lecaros") & (w40_out["Fecha"] == "2026-09-29")].iloc[0]
check(paula29["Origen"] == "REAL" and paula29["Salida"] == "29/09/2026 05:02:00 PM",
      "día real completo NO se toca")

gen = w40_out[w40_out["Origen"] == "GENERADO_OLVIDO"]
check((gen["Horas Trabajadas"] == "8:00:00").all(), "toda fila generada tiene 8:00:00 exactas")
check(not (gen["Fecha"] >= "2026-10-01").any(), "no se genera nada de hoy ni del futuro")
check(not gen["Fecha"].str.contains("2026-09-27").any(), "domingo excluido")

# 3) Determinismo: mismas entradas en dos corridas
out2 = app.completar_olvidados()
check(out == out2, "el mecanismo es determinista (dos corridas iguales)")

# 4) Ventana de entradas generadas 07:30-10:45
ok_ventana = True
for v in gen["Entrada"]:
    dt = datetime.strptime(v, app.FMT_FECHA_HORA)
    m = dt.hour * 60 + dt.minute
    if not (7 * 60 + 30 <= m <= 10 * 60 + 45):
        ok_ventana = False
check(ok_ventana, "entradas generadas dentro de la ventana 07:30-10:45")

# 5) Log de auditoria en Firestore (mock): cada revision con completados queda registrada
log_calls = [c for c in app.db.method_calls
             if c[0] == "collection" and c[1] and c[1][0] == "sistema"]
check(len(log_calls) >= 1, "revision con completados escribe log de auditoria")

# 6) Aviso personalizado: marcas GENERADO_OLVIDO recientes por trabajador
# Ventana de 15 dias desde el 1-oct: cubre W39 (6 dias generados por trabajador)
# y W40 (Nelida 3, Paula 2, Ricardo 3)
app.load_week_data = lambda f: saved.get(f, pd.DataFrame(columns=cols)).copy()
n_nelida = len(app.marcas_generadas_recientes("Nelida Ruiz"))
check(n_nelida == 9, f"aviso de Nelida: 6 (W39) + 3 (W40), hay {n_nelida}")
n_paula = len(app.marcas_generadas_recientes("Paula Lecaros"))
check(n_paula == 8, f"aviso de Paula: 6 (W39) + 2 (W40), hay {n_paula}")
n_ricardo = len(app.marcas_generadas_recientes("Ricardo Adrian Ruiz"))
check(n_ricardo == 9, f"aviso de Ricardo: 6 (W39) + 3 (W40), hay {n_ricardo}")
check(app.marcas_generadas_recientes("Trabajador Inexistente") == [],
      "trabajador sin marcas generadas -> aviso vacio")

print()
if FALLOS:
    print(f"{len(FALLOS)} PRUEBAS FALLARON")
    sys.exit(1)
print("TODO OK")
