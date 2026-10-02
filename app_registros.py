import os

# Firestore usa gRPC. En algunos entornos de despliegue un proxy HTTP altera
# los metadatos gRPC y convierte el ID predeterminado `(default)` en
# `%28default%29`, que Firestore rechaza. Excluimos solo este destino del
# proxy antes de que se carguen las bibliotecas de Google.
_grpc_no_proxy = {
    value.strip()
    for value in os.environ.get("no_grpc_proxy", "").split(",")
    if value.strip()
}
_grpc_no_proxy.add("firestore.googleapis.com")
os.environ["no_grpc_proxy"] = ",".join(sorted(_grpc_no_proxy))

import streamlit as st
import pandas as pd
from datetime import date, datetime, timedelta
from statistics import median
import pytz
import io
import openpyxl
from openpyxl.utils import get_column_letter
from openpyxl.styles import Border, Side

# NUEVOS IMPORTS ---------------------------------------------
import zipfile
import re
import hashlib
# ------------------------------------------------------------

# ---------------------------
# INICIALIZACIÓN DE FIREBASE
# ---------------------------
import firebase_admin
from firebase_admin import credentials, firestore

firebase_secrets = st.secrets["firebase"]

if not firebase_admin._apps:
    cred = credentials.Certificate({
        "type": firebase_secrets["type"],
        "project_id": firebase_secrets["project_id"],
        "private_key_id": firebase_secrets["private_key_id"],
        "private_key": firebase_secrets["private_key"],
        "client_email": firebase_secrets["client_email"],
        "client_id": firebase_secrets["client_id"],
        "auth_uri": firebase_secrets["auth_uri"],
        "token_uri": firebase_secrets["token_uri"],
        "auth_provider_x509_cert_url": firebase_secrets["auth_provider_x509_cert_url"],
        "client_x509_cert_url": firebase_secrets["client_x509_cert_url"]
    })
    firebase_admin.initialize_app(cred)

db = firestore.client()

# ---------------------------
# CONFIGURACIÓN DE HORARIOS
# ---------------------------
# El servidor de Streamlit Cloud corre en UTC. NUNCA se usa datetime.now()
# directo: toda la hora de negocio pasa por ahora_lima() para que el día
# cambie a medianoche de Lima y no a las 19:00.
TZ_LIMA = pytz.timezone("America/Lima")
FMT_FECHA_HORA = "%d/%m/%Y %I:%M:%S %p"

ENTRADA_LIMITE = 11    # Se permite marcar entrada solo hasta las 11:00 AM
SALIDA_LIMITE = 23     # Se permite marcar salida hasta las 23:59 (antes era 18:00)
JORNADA = timedelta(hours=8)   # Jornada estándar del mecanismo anti-olvidos

# Feriados nacionales de Perú 2026 (fuente: gob.pe/feriados). ACTUALIZAR CADA AÑO.
FERIADOS = {
    date(2026, 1, 1),                      # Año Nuevo
    date(2026, 4, 2), date(2026, 4, 3),    # Jueves y Viernes Santo
    date(2026, 5, 1),                      # Día del Trabajo
    date(2026, 6, 7),                      # Batalla de Arica
    date(2026, 6, 29),                     # San Pedro y San Pablo
    date(2026, 7, 23),                     # Día de la Fuerza Aérea
    date(2026, 7, 28), date(2026, 7, 29),  # Fiestas Patrias
    date(2026, 8, 6),                      # Batalla de Junín
    date(2026, 8, 30),                     # Santa Rosa de Lima
    date(2026, 10, 8),                     # Combate de Angamos
    date(2026, 11, 1),                     # Todos los Santos
    date(2026, 12, 8), date(2026, 12, 9),  # Inmaculada / Ayacucho
    date(2026, 12, 25),                    # Navidad
}

# ---------------------------
# FUNCIONES AUXILIARES EXISTENTES
# ---------------------------
def ahora_lima():
    """'Ahora' de negocio, SIEMPRE en hora de Lima.

    En la nube el reloj del servidor está en UTC: usarlo directo hace que la
    fecha (y con ella la semana) cambie a las 19:00 de Lima. Cualquier código
    que necesite la fecha/hora actual debe llamar a esta función.
    """
    return datetime.now(TZ_LIMA)

def es_feriado(fecha):
    """True si la fecha (date) es feriado nacional de Perú."""
    return fecha in FERIADOS

def es_dia_laborable(fecha):
    """Lunes a sábado, excluyendo feriados."""
    return fecha.weekday() <= 5 and not es_feriado(fecha)

def get_week_filename():
    """Genera el nombre del archivo según el año y la semana actual (hora de Lima)."""
    year, week, _ = ahora_lima().isocalendar()
    return f"registro_{year}_W{week}.xlsx"

def get_week_id(filename=None):
    """Obtiene el ID de semana para Firestore (ej. '2026_W11') a partir del filename o la fecha actual (Lima)."""
    if filename:
        match = re.match(r"registro_(\d{4}_W\d+)\.xlsx", filename)
        if match:
            return match.group(1)
    year, week, _ = ahora_lima().isocalendar()
    return f"{year}_W{week}"

def format_datetime(dt):
    """Formatea un datetime (ya en hora de Lima) a cadena."""
    return dt.strftime(FMT_FECHA_HORA)

def parse_timedelta(td_str):
    """Convierte una cadena tipo 'H:MM:SS' a un objeto timedelta."""
    try:
        h, m, s = map(int, td_str.split(':'))
        return timedelta(hours=h, minutes=m, seconds=s)
    except Exception:
        return timedelta()

def create_summary_df(df):
    """Crea un DataFrame resumen con el total de horas trabajadas por cada trabajador."""
    summary = {}
    for name in df["Nombre"].unique():
        valid = df[(df["Nombre"] == name) & (df["Horas Trabajadas"].notna())]
        total = timedelta()
        for _, row in valid.iterrows():
            total += parse_timedelta(row["Horas Trabajadas"])
        summary[name] = str(total)
    summary_df = pd.DataFrame(list(summary.items()), columns=["Nombre", "Total Horas Trabajadas"])
    summary_df = summary_df.sort_values("Nombre")
    return summary_df

def generate_excel_bytes(df):
    """
    Genera un archivo Excel en memoria con dos hojas: 'Registros' y 'Resumen',
    ajusta columnas y añade bordes. Retorna los bytes del archivo.
    """
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, sheet_name='Registros', index=False)
        ws = writer.sheets['Registros']
        for col_cells in ws.columns:
            max_length = 0
            col_letter = col_cells[0].column_letter
            for cell in col_cells:
                if cell.value:
                    cell_length = len(str(cell.value))
                    if cell_length > max_length:
                        max_length = cell_length
            ws.column_dimensions[col_letter].width = max_length + 2

        summary_df = create_summary_df(df)
        summary_df.to_excel(writer, sheet_name='Resumen', index=False)
        ws_summary = writer.sheets['Resumen']
        for col_cells in ws_summary.columns:
            max_length = 0
            col_letter = col_cells[0].column_letter
            for cell in col_cells:
                if cell.value:
                    cell_length = len(str(cell.value))
                    if cell_length > max_length:
                        max_length = cell_length
            ws_summary.column_dimensions[col_letter].width = max_length + 2

        thin_border = Border(
            left=Side(style="thin"),
            right=Side(style="thin"),
            top=Side(style="thin"),
            bottom=Side(style="thin")
        )
        for ws_sheet in [writer.sheets['Registros'], writer.sheets['Resumen']]:
            for row in ws_sheet.iter_rows():
                for cell in row:
                    cell.border = thin_border

    output.seek(0)
    return output.read()

def save_week_data_and_upload(df, filename):
    """
    Guarda el DataFrame en Firestore bajo la colección 'semanas/{week_id}/registros'.
    Cada fila se guarda como un documento con ID '{Nombre}_{Fecha}'.
    El documento padre semanas/{week_id} se crea explícitamente para que
    list_week_files() pueda listarlo con .stream().
    """
    week_id = get_week_id(filename)
    week_ref = db.collection("semanas").document(week_id)
    # Crear/actualizar el documento padre para que sea visible en .stream()
    week_ref.set({"week_id": week_id}, merge=True)
    registros_ref = week_ref.collection("registros")

    for _, row in df.iterrows():
        doc_id = f"{row['Nombre']}_{row['Fecha']}"
        registros_ref.document(doc_id).set(row.to_dict())

def load_week_data(filename):
    """
    Carga el DataFrame desde Firestore.
    Incluye la columna 'Origen' (REAL para marcas humanas, GENERADO_OLVIDO
    para las completadas por el mecanismo anti-olvidos).
    Si no existen registros para la semana, retorna un DataFrame vacío.
    """
    week_id = get_week_id(filename)
    registros_ref = db.collection("semanas").document(week_id).collection("registros")
    docs = list(registros_ref.stream())

    cols = ["Nombre", "Fecha", "Entrada", "Salida", "Horas Trabajadas", "Origen"]
    if docs:
        rows = [doc.to_dict() for doc in docs]
        df = pd.DataFrame(rows).reindex(columns=cols)
        df["Origen"] = df["Origen"].fillna("REAL")
        return df
    else:
        return pd.DataFrame(columns=cols)

def update_firestore(worker, data):
    """
    Actualiza Firestore en la colección 'registros'.
    Cada trabajador tendrá un documento cuyo ID es su nombre, y se guardan los registros.
    """
    doc_ref = db.collection("registros").document(worker)
    doc_ref.set(data)

def register_event(worker, event_type):
    """
    Registra una entrada o salida para un trabajador.
    Se actualiza Firestore y se actualiza el archivo Excel en Firebase Storage.
    Se convierte la hora de UTC a la hora de Lima.
    Se validan horarios:
      - Entrada: solo se permite hasta las 11:00 AM.
      - Salida: solo se permite hasta las 6:00 PM.
    Si no se marcó entrada, no se permite marcar salida.
    """
    filename = get_week_filename()
    df = load_week_data(filename)

    # Hora de negocio SIEMPRE en Lima: el servidor vive en UTC y cambia de
    # día a las 19:00 de Lima, lo que corrompía la fecha del registro.
    local_now = ahora_lima()
    today_str = local_now.date().isoformat()
    now_str = format_datetime(local_now)

    # Validar horario según tipo de evento
    if event_type == "entrada":
        if local_now.hour >= ENTRADA_LIMITE:
            return False, f"Fuera del horario permitido para marcar entrada (hasta las {ENTRADA_LIMITE}:00 AM)."
    elif event_type == "salida":
        if local_now.hour > SALIDA_LIMITE:
            return False, f"Fuera del horario permitido para marcar salida (hasta las {SALIDA_LIMITE}:59)."

    record = df[(df["Nombre"] == worker) & (df["Fecha"] == today_str)]

    if event_type == "entrada":
        if not record.empty and pd.notna(record.iloc[0]["Entrada"]):
            return False, "Ya se ha registrado una entrada hoy para este trabajador."
        if record.empty:
            new_row = {
                "Nombre": worker,
                "Fecha": today_str,
                "Entrada": now_str,
                "Salida": "No marcó salida",
                "Horas Trabajadas": "No marcó salida",
                "Origen": "REAL",
            }
            df = pd.concat([df, pd.DataFrame([new_row])], ignore_index=True)
        else:
            idx = record.index[0]
            df.at[idx, "Entrada"] = now_str
        save_week_data_and_upload(df, filename)
        update_firestore(worker, {"Fecha": today_str, "Evento": "entrada", "Timestamp": now_str})
        return True, f"Entrada registrada para {worker} a las {now_str}"

    elif event_type == "salida":
        if record.empty or pd.isna(record.iloc[0]["Entrada"]):
            return False, "No se ha registrado entrada hoy para este trabajador."
        if pd.notna(record.iloc[0]["Salida"]) and record.iloc[0]["Salida"] != "No marcó salida":
            return False, "Ya se ha registrado una salida hoy para este trabajador."

        idx = record.index[0]
        df.at[idx, "Salida"] = now_str
        try:
            entry_time = datetime.strptime(df.at[idx, "Entrada"], FMT_FECHA_HORA)
            exit_time = datetime.strptime(now_str, FMT_FECHA_HORA)
            worked = exit_time - entry_time
            df.at[idx, "Horas Trabajadas"] = str(worked)
        except Exception as e:
            return False, f"Error al calcular las horas trabajadas: {e}"
        save_week_data_and_upload(df, filename)
        update_firestore(worker, {"Fecha": today_str, "Evento": "salida", "Timestamp": now_str})
        return True, f"Salida registrada para {worker} a las {now_str}"

    return False, "Evento desconocido."

def get_worker_week_hours(worker):
    """Obtiene la suma de las horas trabajadas en la semana para un trabajador."""
    filename = get_week_filename()
    df = load_week_data(filename)
    records = df[(df["Nombre"] == worker) & (df["Horas Trabajadas"].notna())]
    total = timedelta()
    for _, row in records.iterrows():
        total += parse_timedelta(row["Horas Trabajadas"])
    return total

# ---------------------------
# MECANISMO ANTI-OLVIDOS
# ---------------------------
def _parse_fecha_hora(valor):
    try:
        return datetime.strptime(str(valor), FMT_FECHA_HORA)
    except Exception:
        return None

def _minuto_aleatorio(nombre, fecha, low, high):
    """Número estable por (trabajador, fecha): el mismo día siempre produce la misma hora."""
    h = hashlib.md5(f"{nombre}|{fecha.isoformat()}".encode()).hexdigest()
    return low + int(h[:8], 16) % (high - low + 1)

def _entrada_simulada(nombre, fecha, base_min):
    """Hora de entrada 'natural': base histórica del trabajador + jitter, acotada 07:30-10:45."""
    minuto = int(base_min) + _minuto_aleatorio(nombre, fecha, -40, 40)
    minuto = max(7 * 60 + 30, min(10 * 60 + 45, minuto))
    dt = datetime.combine(fecha, datetime.min.time()) + timedelta(minutes=minuto)
    return dt + timedelta(seconds=_minuto_aleatorio(nombre, fecha + timedelta(days=1), 0, 59))

def completar_olvidados():
    """
    Anti-olvidos: revisa los días YA TERMINADOS de la semana actual y la
    anterior, y completa lo que el trabajador no marcó:
      - entrada sin salida -> salida = entrada + 8 h
      - día sin registro   -> día completo (entrada estimada + 8 h)
    Solo días laborables pasados (L-S, no feriado); NUNCA toca el día en
    curso ni el futuro. Toda fila creada/modificada queda con
    Origen = GENERADO_OLVIDO para no mezclarse con las marcas reales.
    """
    hoy = ahora_lima().date()
    completados = []

    for delta in (7, 0):  # semana anterior y semana actual
        anio, semana, _ = (hoy - timedelta(days=delta)).isocalendar()
        filename = f"registro_{anio}_W{semana}.xlsx"
        df = load_week_data(filename)

        # Base de entrada por trabajador: mediana de sus marcados reales
        bases = {}
        for nombre in df["Nombre"].unique():
            minutos = [
                dt.hour * 60 + dt.minute
                for dt in (_parse_fecha_hora(v) for v in df[df["Nombre"] == nombre]["Entrada"])
                if dt is not None
            ]
            bases[nombre] = median(minutos) if minutos else 8 * 60 + 30

        cambio = False
        for idx, row in df.iterrows():
            try:
                fecha = datetime.strptime(str(row["Fecha"]), "%Y-%m-%d").date()
            except Exception:
                continue
            if fecha >= hoy or not es_dia_laborable(fecha):
                continue
            nombre = row["Nombre"]
            entrada_dt = _parse_fecha_hora(row["Entrada"])
            salida_dt = _parse_fecha_hora(row["Salida"])

            if entrada_dt and salida_dt is None:
                df.at[idx, "Salida"] = format_datetime(entrada_dt + JORNADA)
                df.at[idx, "Horas Trabajadas"] = "8:00:00"
                df.at[idx, "Origen"] = "GENERADO_OLVIDO"
                cambio = True
                completados.append(f"{nombre} {fecha}: salida generada (entrada + 8 h)")
            elif salida_dt and entrada_dt is None:
                df.at[idx, "Entrada"] = format_datetime(salida_dt - JORNADA)
                df.at[idx, "Horas Trabajadas"] = "8:00:00"
                df.at[idx, "Origen"] = "GENERADO_OLVIDO"
                cambio = True
                completados.append(f"{nombre} {fecha}: entrada generada (salida - 8 h)")

        # Días sin NINGÚN registro -> fila completa generada
        existentes = {(str(r["Nombre"]), str(r["Fecha"])) for _, r in df.iterrows()}
        lunes = datetime.fromisocalendar(anio, semana, 1).date()
        filas_nuevas = []
        for offset in range(7):
            dia = lunes + timedelta(days=offset)
            if dia >= hoy or not es_dia_laborable(dia):
                continue
            for nombre in user_list:
                if (nombre, dia.isoformat()) in existentes:
                    continue
                ent = _entrada_simulada(nombre, dia, bases.get(nombre, 8 * 60 + 30))
                filas_nuevas.append({
                    "Nombre": nombre,
                    "Fecha": dia.isoformat(),
                    "Entrada": format_datetime(ent),
                    "Salida": format_datetime(ent + JORNADA),
                    "Horas Trabajadas": "8:00:00",
                    "Origen": "GENERADO_OLVIDO",
                })
                completados.append(f"{nombre} {dia}: día completo generado")
        if filas_nuevas:
            df = pd.concat([df, pd.DataFrame(filas_nuevas)], ignore_index=True)
            cambio = True

        if cambio:
            save_week_data_and_upload(df, filename)

    if completados:
        # Log de auditoria: queda constancia de cuando se genero cada marca
        db.collection("sistema").document("anti_olvidos_log").collection("revisiones").add({
            "generado_en": format_datetime(ahora_lima()),
            "cantidad": len(completados),
            "detalle": completados,
        })

    return completados

def revision_diaria_anti_olvidos():
    """Ejecuta completar_olvidados() máximo una vez por día (marca en Firestore)."""
    ref = db.collection("sistema").document("anti_olvidos")
    hoy = ahora_lima().date().isoformat()
    doc = ref.get()
    if doc.exists and doc.to_dict().get("ultima_revision") == hoy:
        return None
    ref.set({"ultima_revision": hoy}, merge=True)
    return completar_olvidados()

def marcas_generadas_recientes(worker, dias=15):
    """
    Marcas GENERADO_OLVIDO del trabajador en los ultimos N dias,
    para el aviso personalizado que se muestra al iniciar sesion.
    """
    hoy = ahora_lima().date()
    desde = hoy - timedelta(days=dias)
    semanas = sorted({
        (desde + timedelta(days=i)).isocalendar()[:2]
        for i in range(dias + 1)
    })
    avisos = []
    for anio, semana in semanas:
        df = load_week_data(f"registro_{anio}_W{semana}.xlsx")
        if df.empty:
            continue
        gen = df[(df["Nombre"] == worker) & (df["Origen"] == "GENERADO_OLVIDO")]
        for _, row in gen.iterrows():
            try:
                fecha = datetime.strptime(str(row["Fecha"]), "%Y-%m-%d").date()
            except Exception:
                continue
            if fecha < desde or fecha >= hoy:
                continue
            ent = _parse_fecha_hora(row["Entrada"])
            sal = _parse_fecha_hora(row["Salida"])
            detalle = str(row["Fecha"])
            if ent and sal:
                detalle = (f"{fecha.strftime('%d/%m/%Y')} - "
                           f"entrada {ent.strftime('%I:%M %p')}, "
                           f"salida {sal.strftime('%I:%M %p')} (8 h)")
            avisos.append(detalle)
    return sorted(set(avisos))

# ---------------------------
# NUEVA FUNCIÓN: GENERAR ARCHIVO MENSUAL
# ---------------------------
def generate_monthly_file(selected_year, selected_month):
    """
    Reúne todos los registros semanales almacenados en Firestore correspondientes
    al mes y año seleccionados. Genera un archivo Excel con dos hojas: 'Registros' y 'Resumen'.
    Retorna el contenido binario del archivo.
    """
    week_docs = list(db.collection("semanas").stream())
    monthly_dfs = []

    for week_doc in week_docs:
        week_id = week_doc.id
        if not re.match(r"\d{4}_W\d+", week_id):
            continue
        try:
            registros_ref = db.collection("semanas").document(week_id).collection("registros")
            docs = list(registros_ref.stream())
            if not docs:
                continue
            rows = [doc.to_dict() for doc in docs]
            df_week = pd.DataFrame(rows).reindex(
                columns=["Nombre", "Fecha", "Entrada", "Salida", "Horas Trabajadas"]
            )
            if not df_week.empty:
                df_week["Fecha_dt"] = pd.to_datetime(df_week["Fecha"], format="%Y-%m-%d", errors="coerce")
                mask = (
                    (df_week["Fecha_dt"].dt.year == selected_year) &
                    (df_week["Fecha_dt"].dt.month == selected_month)
                )
                df_filtered = df_week.loc[mask].drop(columns=["Fecha_dt"])
                if not df_filtered.empty:
                    monthly_dfs.append(df_filtered)
        except Exception as e:
            st.error(f"Error procesando la semana {week_id}: {e}")

    if not monthly_dfs:
        st.error("No se encontraron registros para el mes y año seleccionados.")
        return None

    df_month = pd.concat(monthly_dfs, ignore_index=True)
    resumen_month = create_summary_df(df_month)

    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df_month.to_excel(writer, sheet_name='Registros', index=False)
        ws = writer.sheets['Registros']
        for col_cells in ws.columns:
            max_length = 0
            col_letter = col_cells[0].column_letter
            for cell in col_cells:
                if cell.value:
                    cell_length = len(str(cell.value))
                    if cell_length > max_length:
                        max_length = cell_length
            ws.column_dimensions[col_letter].width = max_length + 2

        resumen_month.to_excel(writer, sheet_name='Resumen', index=False)
        ws_resumen = writer.sheets['Resumen']
        for col_cells in ws_resumen.columns:
            max_length = 0
            col_letter = col_cells[0].column_letter
            for cell in col_cells:
                if cell.value:
                    cell_length = len(str(cell.value))
                    if cell_length > max_length:
                        max_length = cell_length
            ws_resumen.column_dimensions[col_letter].width = max_length + 2

        thin_border = Border(
            left=Side(style="thin"),
            right=Side(style="thin"),
            top=Side(style="thin"),
            bottom=Side(style="thin")
        )
        for ws_sheet in [ws, ws_resumen]:
            for row in ws_sheet.iter_rows():
                for cell in row:
                    cell.border = thin_border

    output.seek(0)
    return output

# ---------------------------
# NUEVAS FUNCIONES AUXILIARES PARA DESCARGAS ----------------
_pattern_week = re.compile(r"registro_(\d{4})_W(\d{1,2})\.xlsx")

def list_week_files() -> list[str]:
    """Devuelve los nombres de archivo semanales disponibles en Firestore."""
    week_docs = list(db.collection("semanas").stream())
    week_ids = [doc.id for doc in week_docs if re.match(r"\d{4}_W\d+", doc.id)]
    return sorted([f"registro_{wid}.xlsx" for wid in week_ids])

def week_files_for_month(year: int, month: int, all_files: list[str]) -> list[str]:
    """Filtra las semanas cuyo primer día ISO-week cae en el mes/año dados."""
    target = []
    for fname in all_files:
        year_w, week = map(int, _pattern_week.match(fname).groups())
        first_day = datetime.fromisocalendar(year_w, week, 1)
        if first_day.year == year and first_day.month == month:
            target.append(fname)
    return target

def zip_blobs(blob_names: list[str]) -> io.BytesIO:
    """Crea un ZIP en memoria con archivos Excel generados desde Firestore."""
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as zf:
        for name in blob_names:
            df = load_week_data(name)
            excel_data = generate_excel_bytes(df)
            zf.writestr(name, excel_data)
    buf.seek(0)
    return buf

# ---------------------------
# INTERFAZ DE USUARIO CON STREAMLIT
# ---------------------------

user_list = ["Nelida Ruiz", "Ricardo Adrian Ruiz", "Paula Lecaros"]
user_passwords = st.secrets["user_passwords"]

st.title("Registro de Entradas y Salidas (CLOUD: Firestore)")

# --- Mecanismo anti-olvidos: revisión diaria al abrir la app ---------------
try:
    _completados_hoy = revision_diaria_anti_olvidos()
    if _completados_hoy:
        st.caption(
            f"Anti-olvidos: se completaron {len(_completados_hoy)} marcados de "
            "días anteriores (quedan marcados como GENERADO_OLVIDO)."
        )
except Exception as e:
    st.caption(f"Anti-olvidos no pudo ejecutarse: {e}")

with st.expander("Selecciona al Trabajador"):
    worker = st.selectbox("Elige tu nombre:", [""] + user_list)

if worker:
    password_input = st.text_input("Ingrese su contraseña:", type="password")

    if password_input:
        if password_input == user_passwords.get(worker, ""):
            if worker == "Ricardo Adrian Ruiz":
                st.info("Bienvenido, ADMIN.")
            else:
                st.info(f"Bienvenido, {worker}.")

            # Aviso personalizado: marcas que el anti-olvidos completo por el usuario
            try:
                generadas = marcas_generadas_recientes(worker)
                if generadas:
                    st.warning(
                        f"El sistema completo {len(generadas)} marcado(s) tuyo(s) por olvido "
                        "(quedan como GENERADO_OLVIDO). Si algun dato no corresponde, "
                        "avisale al administrador:"
                    )
                    for item in generadas:
                        st.caption(f"- {item}")
            except Exception as e:
                st.caption(f"No se pudo revisar tus marcas generadas: {e}")

            st.header(f"Registro para: {worker}")

            col1, col2 = st.columns(2)
            with col1:
                if st.button("Registrar Entrada"):
                    success, msg = register_event(worker, "entrada")
                    if success:
                        st.success(msg)
                    else:
                        st.warning(msg)
            with col2:
                if st.button("Registrar Salida"):
                    success, msg = register_event(worker, "salida")
                    if success:
                        st.success(msg)
                    else:
                        st.warning(msg)

            total_hours = get_worker_week_hours(worker)
            st.write("Total de horas trabajadas esta semana:", str(total_hours))

            if worker == "Ricardo Adrian Ruiz":
                st.subheader("Resumen Semanal General")
                if st.button("Mostrar resumen de horas por trabajador"):
                    filename = get_week_filename()
                    df = load_week_data(filename)
                    resumen = create_summary_df(df)
                    st.dataframe(resumen)
            else:
                if st.button("Mostrar mis registros semanales"):
                    filename = get_week_filename()
                    df = load_week_data(filename)
                    worker_records = df[df["Nombre"] == worker]
                    st.dataframe(worker_records)

            # --- Sección ADMIN: Descarga de registros semanales ---------------
            if worker == "Ricardo Adrian Ruiz":
                st.markdown("---")
                st.subheader("Descarga de registros semanales")

                # Valores de año y mes para filtros y descargas
                current_year = ahora_lima().year
                selected_year = st.session_state.get("selected_year", current_year)
                selected_month = st.session_state.get("selected_month_num", ahora_lima().month)

                week_files = list_week_files()

                if not week_files:
                    st.info("No hay archivos semanales en Firestore.")
                else:
                    selected_file = st.selectbox(
                        "Selecciona un archivo semanal para descargar:", [""] + week_files
                    )

                    col_dl_one, col_dl_month, col_dl_all = st.columns(3)

                    # Descargar archivo individual
                    with col_dl_one:
                        if selected_file:
                            df_selected = load_week_data(selected_file)
                            data = generate_excel_bytes(df_selected)
                            st.download_button(
                                label=f"Descargar {selected_file}",
                                data=data,
                                file_name=selected_file,
                                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                            )

                    # ZIP del mes seleccionado (usa año/mes elegidos arriba)
                    with col_dl_month:
                        if st.button("ZIP del mes seleccionado"):
                            month_weeks = week_files_for_month(
                                selected_year, selected_month, week_files
                            )
                            if month_weeks:
                                buf = zip_blobs(month_weeks)
                                zip_name = f"registros_{selected_year}_{selected_month:02d}.zip"
                                st.download_button(
                                    label=f"Descargar {zip_name}",
                                    data=buf,
                                    file_name=zip_name,
                                    mime="application/zip"
                                )
                            else:
                                st.warning("No hay semanas para ese mes.")

                    # ZIP con todo el histórico
                    with col_dl_all:
                        if st.button("ZIP con TODO"):
                            buf = zip_blobs(week_files)
                            st.download_button(
                                label="Descargar registros_all.zip",
                                data=buf,
                                file_name="registros_all.zip",
                                mime="application/zip"
                            )

                # --- Auditoría del anti-olvidos: cuándo se generó cada marca ---
                st.markdown("---")
                st.subheader("Auditoría del anti-olvidos")
                if st.button("Ver auditoría del anti-olvidos"):
                    log_docs = list(
                        db.collection("sistema")
                          .document("anti_olvidos_log")
                          .collection("revisiones")
                          .stream()
                    )
                    if not log_docs:
                        st.info("Aún no hay revisiones con marcados completados.")
                    else:
                        rows = []
                        for d in log_docs:
                            data = d.to_dict()
                            detalle = data.get("detalle", [])
                            rows.append({
                                "Revisión del": data.get("generado_en", ""),
                                "Marcados completados": data.get("cantidad", 0),
                                "Detalle": " | ".join(detalle) if isinstance(detalle, list) else str(detalle),
                            })
                        df_log = pd.DataFrame(rows).sort_values(
                            "Revisión del", ascending=False
                        ).reset_index(drop=True)
                        st.dataframe(df_log)

            # --- Sección ADMIN: Generar y descargar archivo mensual -----------
            if worker == "Ricardo Adrian Ruiz":
                st.markdown("---")
                st.subheader("Archivo Mensual (nuevo, generado al vuelo)")

                col_year, col_month = st.columns(2)
                with col_year:
                    current_year = ahora_lima().year
                    selected_year = st.number_input(
                        "Elige el año",
                        min_value=2000,
                        max_value=current_year + 1,
                        value=st.session_state.get("selected_year", current_year),
                        step=1,
                        key="selected_year",
                    )
                with col_month:
                    month_names = [
                        "Enero",
                        "Febrero",
                        "Marzo",
                        "Abril",
                        "Mayo",
                        "Junio",
                        "Julio",
                        "Agosto",
                        "Septiembre",
                        "Octubre",
                        "Noviembre",
                        "Diciembre",
                    ]
                    selected_month_name = st.selectbox(
                        "Elige el mes",
                        month_names,
                        index=st.session_state.get("selected_month_num", ahora_lima().month) - 1,
                        key="selected_month_name",
                    )
                    selected_month = month_names.index(selected_month_name) + 1
                    st.session_state["selected_month_num"] = selected_month

                if st.button("Generar archivo mensual"):
                    output_buffer = generate_monthly_file(selected_year, selected_month)
                    if output_buffer is not None:
                        filename = f"registro_{selected_year}_{selected_month:02d}.xlsx"
                        st.download_button(
                            label="Descargar archivo mensual",
                            data=output_buffer,
                            file_name=filename,
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                        )
        else:
            st.error("Contraseña incorrecta. Intente nuevamente.")
