# app_dashboard_full.py
# Dashboard Flet con:
# - Selección de archivo, consolidación multi-hoja
# - Filtro por SERVICIO: sólo "consulta externa" y "urgencia gral/gener(al)" (normalizado)
# - Separa filas por persona + servicio_norm + fecha_evento (validación o creación)
# - Vista previa con scroll horizontal y vertical
# - Exportación con barra de progreso
# Requisitos: pip install flet pandas openpyxl

from __future__ import annotations
import re, traceback
from datetime import datetime
from pathlib import Path
import unicodedata
import pandas as pd
import flet as ft

# ---------------- Compat icons/colors ----------------
try:
    ICONS = ft.icons
except Exception:
    try:
        from flet import icons as ICONS
    except Exception:
        ICONS = None

try:
    COLORS = ft.colors
except Exception:
    try:
        from flet import colors as COLORS
    except Exception:
        COLORS = None

def tone(name: str, default_hex: str) -> str:
    return getattr(COLORS, name, default_hex) if COLORS else default_hex

PRIMARY = tone("BLUE_600", "#2563EB")
GRAD_1 = tone("INDIGO_500", "#6366F1")
GRAD_2 = tone("BLUE_600", "#2563EB")
SURFACE = tone("GREY_50", "#F9FAFB")
SURFACE_ALT = tone("GREY_100", "#F3F4F6")
TEXT = tone("GREY_900", "#111827")
TEXT_MUTED = tone("GREY_700", "#374151")
DANGER = tone("RED_600", "#DC2626")
WHITE = tone("WHITE", "#FFFFFF")

# --------- Config preview ---------
MAX_ROWS_PREVIEW = 80
MAX_COLS_PREVIEW = 20
TRUNCATE_CELL_CHARS = 120
# ----------------------------------

# ====================== NORMALIZACIÓN / SINÓNIMOS ======================
def _unidecode_local(x: str) -> str:
    return "".join(c for c in unicodedata.normalize("NFKD", str(x)) if not unicodedata.combining(c))

def slug(s: str) -> str:
    if s is None: return ""
    s = _unidecode_local(str(s)).strip().lower()
    s = "".join(ch if ch.isalnum() else "_" for ch in s)
    return re.sub(r"_+", "_", s).strip("_")

# ---- Estudio (manteniendo lo tuyo + añadidos) ----
STUDY_SHORT = {
    "QUIMICA SANGUINEA 6 ELEMENTOS": "QS6",
    "QUIMICA SANGUINEA 3 ELEMENTOS" : "QS3",
    "QUIMICA SANGUINEA 5 ELEMENTOS" : "QS5",
    "GLUCOSA": "GLU",
    "UREA":"UREA",
    "CREATININA": "CREATININA",
    "ACIDO URICO": "ACH",
    "COLESTEROL": "COL",
    "TRIGLICERIDOS": "TRIG",
    "TASA DE FILTRACION GLOMERULAR":"TFG",
    "QUIMICA SANGUINEA 4": "QS4",
    "EXAMEN GENERAL DE ORINA": "EGO",
    "BIOMETRIA HEMATICA": "BH",
    "HEMOGLOBINA GLICOSILADA": "HbA1c",
    "PERFIL DE LIPIDOS": "PL",

    # AÑADIDOS (no quitan nada)
    "QUIMICA SANGUINEA 6": "QS6",
    "BIOQUIMICA SANGUINEA 6 ELEMENTOS": "QS6",
    "EXAMEN GENERAL DE ORINA (EGO)": "EGO",
    "ORINA EXAMEN GENERAL": "EGO",
    "BIOMETRIA HEMATICA B4": "BH",
}

# ---- Prueba (manteniendo lo tuyo + añadidos) ----
TEST_SYNONYMS = {
    "glucosa":"Glucosa",
    "creatinina serica":"Creatinina","creatinina":"Creatinina","creatinina sérica":"Creatinina",
    "urea":"Urea",
    "acido urico":"Ácido Úrico","ácido úrico":"Ácido Úrico",
    "bun":"BUN","colesterol":"Colesterol",
    "trigliceridos":"Triglicéridos","triglicéridos":"Triglicéridos",
    "hemoglobina glicosilada":"Hemoglobina Glicosilada",
    "tasa de filtracion glomerular":"Tasa de Filtración Glomerular",
    "tasa de filtración glomerular":"Tasa de Filtración Glomerular",
    "ph":"pH",
    "proteinas":"Proteínas","proteínas":"Proteínas",
    "nitrito":"Nitrito","nitritos":"Nitrito",
    "hemoglobina":"Hemoglobina",

    # AÑADIDOS
    "proteinas totales": "Proteínas",
    "proteina": "Proteínas",
    "tfg": "Tasa de Filtración Glomerular",
}

# Aceptar ambos separadores en las llaves (en-dash y guion corto)
SEP_KEYS = [" – ", " - "]
def key_variants(study: str, test: str) -> list[str]:
    return [f"{study}{sep}{test}" for sep in SEP_KEYS]

# ===== Columnas finales (mismas cabeceras, agregadas variantes de llave) =====
want = [
    # Glucosa
    ("Glucosa (EGO)", [
        "EGO – Glucosa", "EGO - Glucosa",
        "ORINA – Glucosa", "ORINA - Glucosa"
    ]),
    ("Glucosa (QS)", [
        "QS6 – Glucosa", "QS6 - Glucosa",
        "QS4 – Glucosa", "QS4 - Glucosa"
    ]),

    # Urea
    ("Urea (Orina)", [
        "EGO – Urea", "EGO - Urea",
        "ORINA – Urea", "ORINA - Urea"
    ]),
    ("Urea (QS)", [
        "QS6 – Urea", "QS6 - Urea",
        "QS4 – Urea", "QS4 - Urea"
    ]),

    # Creatinina sérica (QS)
    ("Creatinina sérica (QS)", [
        "QS6 – Creatinina", "QS6 - Creatinina",
        "QS4 – Creatinina", "QS4 - Creatinina"
    ]),

    # TFG (QS)
    ("Tasa de filtración glomerular (QS)", [
        "QS6 – Tasa de Filtración Glomerular", "QS6 - Tasa de Filtración Glomerular",
        "QS4 – Tasa de Filtración Glomerular", "QS4 - Tasa de Filtración Glomerular",
        "QS6 – TFG", "QS6 - TFG", "QS4 – TFG", "QS4 - TFG"
    ]),

    # Hemoglobina
    ("Hemoglobina (Biometría Hemática)", [
        "BH – Hemoglobina", "BH - Hemoglobina"
    ]),
    ("Hemoglobina (Sanguínea)", [
        "QS6 – Hemoglobina", "QS6 - Hemoglobina",
        "QS4 – Hemoglobina", "QS4 - Hemoglobina"
    ]),

    # HbA1c
    ("Hemoglobina glicosilada", [
        "HbA1c – Hemoglobina Glicosilada", "HbA1c - Hemoglobina Glicosilada",
        "HEMOGLOBINA GLICOSILADA – Hemoglobina Glicosilada",
        "HEMOGLOBINA GLICOSILADA - Hemoglobina Glicosilada"
    ]),

    # EGO: Proteínas y Nitritos
    ("Proteínas (EGO)", [
        "EGO – Proteínas", "EGO - Proteínas",
        "ORINA – Proteínas", "ORINA - Proteínas"
    ]),
    ("Nitritos (EGO)", [
        "EGO – Nitrito", "EGO - Nitrito",
        "EGO – Nitritos", "EGO - Nitritos",
        "ORINA – Nitrito", "ORINA - Nitrito",
        "ORINA – Nitritos", "ORINA - Nitritos"
    ]),

    # pH (Sanguínea) — QS; si no existe, EGO como respaldo
    ("pH (Sanguínea)", [
        "QS6 – pH", "QS6 - pH",
        "QS4 – pH", "QS4 - pH",
        "EGO – pH", "EGO - pH"
    ]),

    # Hemoglobina en orina (EGO)
    ("Hemoglobina (orina EGO)", [
        "EGO – Hemoglobina", "EGO - Hemoglobina",
        "ORINA – Hemoglobina", "ORINA - Hemoglobina"
    ]),
]

# ====================== PIPELINE ======================
def normalize_headers(cols: list[str]) -> list[str]:
    syn = {
        "nombres":"nombres","nombre":"nombres",
        "apellidop":"apellido_paterno","apellido_p":"apellido_paterno",
        "apellidom":"apellido_materno","apellido_m":"apellido_materno",
        "sexo":"sexo",
        "servicio":"servicio",
        "fecnacimiento":"fecha_nacimiento","fecnacimien":"fecha_nacimiento",
        "fec_nacimiento":"fecha_nacimiento","fecha_nac":"fecha_nacimiento",
        "estudio":"estudio",
        "prueba":"prueba","resultado":"resultado",
        "rangoinferior":"rango_inferior","rangosuperior":"rango_superior",
        "rangoalterno":"rango_alternativo","rangoaltern":"rango_alternativo",
        "ide":"ide","nss":"nss","codigo":"codigo","loinc":"loinc","p_loinc":"p_loinc",
        "fechacrea":"fecha_creacion","fechaval":"fecha_validacion","usrval":"usuario_validador",
        "fecha":"fecha",  # por si viene una fecha genérica
    }
    return [syn.get(slug(c), slug(c)) for c in cols]

def build_nombre(df: pd.DataFrame) -> pd.Series:
    if "apellido_p" in df.columns and "apellido_m" in df.columns:
        df = df.rename(columns={"apellido_p":"apellido_paterno","apellido_m":"apellido_materno"})
    if {"nombres","apellido_paterno","apellido_materno"}.issubset(df.columns):
        nombre = (
            df["nombres"].fillna("").str.strip()+" "+
            df["apellido_paterno"].fillna("").str.strip()+" "+
            df["apellido_materno"].fillna("").str.strip()
        ).str.replace(r"\s+"," ",regex=True).str.strip()
    elif "nombres" in df.columns:
        nombre = df["nombres"].fillna("").astype(str).str.strip()
    else:
        nombre = pd.Series([""]*len(df))
    return nombre

def parse_dob(s: str) -> str:
    if s is None or s=="" or pd.isna(s): return ""
    for fmt in ("%d/%m/%Y","%Y-%m-%d","%d-%m-%Y","%m/%d/%Y"):
        try:
            return datetime.strptime(str(s), fmt).strftime("%Y-%m-%d")
        except Exception:
            pass
    try:
        dt = pd.to_datetime(s, errors="coerce")
        if pd.notna(dt): return dt.strftime("%Y-%m-%d")
    except Exception:
        pass
    return ""

def edad_from_iso(iso: str) -> int | None:
    if not iso: return None
    try:
        dob = datetime.strptime(iso,"%Y-%m-%d")
    except Exception:
        return None
    today = datetime.today()
    return today.year - dob.year - ((today.month,today.day) < (dob.month,dob.day))

def canon_study(study_raw: str) -> str:
    if not study_raw: return ""
    key = _unidecode_local(study_raw).strip().upper()
    return STUDY_SHORT.get(key, study_raw.strip())

def canon_test(test_raw: str) -> str:
    if not test_raw: return ""
    key = _unidecode_local(test_raw).strip().lower()
    return TEST_SYNONYMS.get(key, test_raw.strip())

def first_nonempty(series: pd.Series) -> str:
    for v in series:
        if pd.notna(v) and str(v).strip() != "":
            return str(v)
    return ""

def normalize_servicio(s: str) -> str:
    """
    Normaliza 'servicio' y mapea variantes/abreviaturas/typos a categorías canónicas.
    Queremos conservar SOLO:
      - 'consulta externa'
      - 'urgencia general'  (incluye: 'urg. gral', 'urg gral', 'urgencias general', 'urgenciasl general', etc.)
    """
    if s is None or (isinstance(s, float) and pd.isna(s)):
        return ""

    raw = _unidecode_local(str(s)).lower().strip()

    # Normalización básica
    raw = raw.replace(".", " ").replace("-", " ").replace("/", " ")
    raw = re.sub(r"\s+", " ", raw)

    # Corrección de typos frecuentes
    raw = raw.replace("urgenciasl", "urgencias")   # 'urgenciasl general' -> 'urgencias general'
    raw = raw.replace("urgenc", "urgencia")        # 'urgenc gral' -> 'urgencia gral'
    raw = raw.replace("genral", "general")         # 'genral' -> 'general'
    raw = raw.replace("grl", "gral")               # 'grl' -> 'gral'

    # ====== DETECCIÓN CONSULTA EXTERNA ======
    if (
        re.search(r"\bconsulta\s*externa\b", raw)
        or re.search(r"\bconsulta\s*ext(erna)?\b", raw)
        or re.search(r"\bcons?(\s*)ext(\s*erna)?\b", raw)
        or re.search(r"\bc(\s*)externa\b", raw)
    ):
        return "consulta externa"

    # ====== DETECCIÓN URGENCIA GENERAL ======
    has_urg = (
        re.search(r"\burgencia(s)?\b", raw) is not None
        or re.search(r"\burg\b", raw) is not None
        or re.search(r"\burgs\b", raw) is not None
        or re.search(r"\burg\w*\b", raw) is not None  # urg., urg-
    )
    has_general = (
        re.search(r"\bgeneral\b", raw) is not None
        or re.search(r"\bgral\b", raw) is not None
    )

    if has_urg and has_general:
        return "urgencia general"

    # Si sólo dice 'urgencia' / 'urgencias' sin 'general', lo tratamos como urgencia general
    if re.search(r"\burgencia(s)?\b", raw):
        return "urgencia general"

    # Cualquier otro servicio queda tal cual (y luego será excluido por el filtro)
    return raw


def cargar_y_preparar_df(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()
    df.columns = normalize_headers(list(df.columns))
    # Asegurar columnas base
    for col in ["nombres","apellido_paterno","apellido_materno","sexo","servicio",
                "fecha_nacimiento","estudio","prueba","resultado",
                "fecha_creacion","fecha_validacion","fecha"]:
        if col not in df.columns:
            df[col] = ""

    # --- Filtro por SERVICIO (SOLO consulta externa y urgencia gral/general) ---
    df["servicio_norm"] = df["servicio"].apply(normalize_servicio)
    df = df[df["servicio_norm"].isin(["consulta externa", "urgencia general"])].copy()

    # Si todo quedó vacío tras el filtro, devolvemos DataFrame vacío con las columnas requeridas
    if df.empty:
        cols = ["nombre","sexo","edad","mayor_18","fecha_nacimiento","servicio_norm","fecha_evento"]
        return pd.DataFrame(columns=cols)

    # Datos de persona
    df["nombre"] = build_nombre(df)
    df["fecha_nacimiento"] = df["fecha_nacimiento"].apply(parse_dob)
    df["edad"] = df["fecha_nacimiento"].apply(edad_from_iso)
    df["mayor_18"] = df["edad"].apply(lambda e: "Sí" if e is not None and e >= 18 else "No")

    # --- FECHA DEL EVENTO (clave para separar filas por distintas atenciones) ---
    df["fecha_validacion"] = df["fecha_validacion"].apply(parse_dob)
    df["fecha_creacion"] = df["fecha_creacion"].apply(parse_dob)
    # Si existe una 'fecha' genérica, úsala si ambas anteriores están vacías
    df["fecha"] = df["fecha"].apply(parse_dob)

    # Preferir fecha_validacion > fecha_creacion > fecha
    df["fecha_evento"] = df["fecha_validacion"]
    df.loc[df["fecha_evento"] == "", "fecha_evento"] = df.loc[df["fecha_evento"] == "", "fecha_creacion"]
    df.loc[df["fecha_evento"] == "", "fecha_evento"] = df.loc[df["fecha_evento"] == "", "fecha"]

    # Canon estudio/prueba y llave
    df["study_canon"] = df["estudio"].fillna("").apply(canon_study)
    df["test_canon"] = df["prueba"].fillna("").apply(canon_test)
    df["col_key"] = (df["study_canon"].fillna("").astype(str).str.strip()
                     + df["test_canon"].apply(lambda s: " – "+s if str(s).strip()!="" else ""))

    return df

def pivot_por_persona_cols_estudio_prueba(df: pd.DataFrame) -> pd.DataFrame:
    # INCLUYE fecha_evento en el índice
    base_cols = ["fecha_nacimiento","nombre","sexo","edad","mayor_18","servicio_norm","fecha_evento"]
    df_base = df[base_cols + ["col_key","resultado"]].copy()
    if df_base.empty or (df_base["col_key"].replace("", pd.NA).isna().all() and df_base["resultado"].replace("", pd.NA).isna().all()):
        return pd.DataFrame(columns=["nombre","sexo","edad","mayor_18","fecha_nacimiento","servicio_norm","fecha_evento"])
    pivot = (df_base
             .pivot_table(index=base_cols, columns="col_key", values="resultado",
                          aggfunc=first_nonempty, fill_value="")
             .reset_index())
    pivot = pivot.sort_values(["fecha_nacimiento","nombre","fecha_evento"]).reset_index(drop=True)
    fixed = ["nombre","sexo","edad","mayor_18","fecha_nacimiento","servicio_norm","fecha_evento"]
    tests = [c for c in pivot.columns if c not in fixed]
    pivot = pivot[fixed + sorted(tests, key=lambda x: x.lower())]
    return pivot

def pick_first(df: pd.DataFrame, cols: list[str]) -> pd.Series:
    vals = []
    for _, row in df.iterrows():
        val = ""
        for c in cols:
            if c in df.columns:
                v = row[c]
                if v is not None and str(v).strip() != "":
                    val = str(v)
                    break
        vals.append(val)
    return pd.Series(vals, index=df.index)

def reducir_a_columnas_solicitadas(df: pd.DataFrame) -> pd.DataFrame:
    # df trae columnas "ESTUDIO – PRUEBA"
    base_cols = ["nombre","sexo","edad","mayor_18","fecha_nacimiento","servicio_norm","fecha_evento"]
    if df.empty:
        out = pd.DataFrame(columns=["id_trabajador"] + base_cols + [w[0] for w in want])
        return out

    out = df[base_cols].copy()
    for new_name, candidates in want:
        out[new_name] = pick_first(df, candidates)

    # ID incremental
    out.insert(0, "id_trabajador", range(1, len(out)+1))
    return out

def consolidar_todas_las_hojas(path_xlsx: str) -> pd.DataFrame:
    xls = pd.ExcelFile(path_xlsx, engine="openpyxl")
    tablas = []
    for sheet in xls.sheet_names:
        df_raw = pd.read_excel(xls, sheet_name=sheet, dtype=str, header=0)
        df_base = cargar_y_preparar_df(df_raw)
        tabla = pivot_por_persona_cols_estudio_prueba(df_base)
        if not tabla.empty:
            tablas.append(tabla)

    if not tablas:
        cols = ["id_trabajador","nombre","sexo","edad","mayor_18","fecha_nacimiento","servicio_norm","fecha_evento"] + [w[0] for w in want]
        return pd.DataFrame(columns=cols)

    unido = pd.concat(tablas, ignore_index=True, sort=True).fillna("")
    # OJO: ya NO usamos consulta_externa en 'fixed'
    fixed = ["nombre","sexo","edad","mayor_18","fecha_nacimiento","servicio_norm","fecha_evento"]
    for c in fixed:
        if c not in unido.columns:
            unido[c] = ""
    agg_dict = {col: first_nonempty for col in unido.columns}
    consolidado = (unido.groupby(fixed, dropna=False, as_index=False).agg(agg_dict))
    final = reducir_a_columnas_solicitadas(consolidado)
    return final

# ---------------- UI helpers ----------------
def truncate(s: str, n: int) -> str:
    if s is None: return ""
    s = str(s)
    return (s[: n-1] + "…") if len(s) > n else s

def df_to_datatable(df: pd.DataFrame, max_rows: int = 100, max_cols: int | None = None) -> ft.DataTable:
    if df is None or df.empty:
        return ft.DataTable(columns=[ft.DataColumn(ft.Text("Sin datos"))], rows=[])
    view = df.head(max_rows)
    if max_cols is not None and max_cols > 0:
        view = view.iloc[:, :max_cols]
    view = view.applymap(lambda v: truncate("" if pd.isna(v) else v, TRUNCATE_CELL_CHARS))
    columns = [ft.DataColumn(ft.Text(str(c))) for c in view.columns]
    rows = [
        ft.DataRow(cells=[ft.DataCell(ft.Text("" if pd.isna(v) else str(v))) for v in view.iloc[i]])
        for i in range(len(view))
    ]
    return ft.DataTable(
        columns=columns,
        rows=rows,
        heading_row_height=40,
        data_row_min_height=36,
        divider_thickness=0.6,
        column_spacing=28,
    )

# ============================ APP ============================
def main(page: ft.Page):
    page.title = "Reportes de Laboratorio — Dashboard"
    page.window_min_width = 1180
    page.window_min_height = 760
    page.theme_mode = "light"
    page.bgcolor = SURFACE_ALT

    # HERO
    hero = ft.Container(
        padding=20,
        border_radius=16,
        gradient=ft.LinearGradient(
            begin=ft.alignment.top_left,
            end=ft.alignment.bottom_right,
            colors=[GRAD_1, GRAD_2],
        ),
        shadow=ft.BoxShadow(blur_radius=20, spread_radius=1, color="#00000022"),
        content=ft.Row(
            [
                ft.Icon(ICONS.DASHBOARD, color=WHITE, size=32) if ICONS else ft.Container(),
                ft.Column(
                    [
                        ft.Text("Consolidador • Resultados de Laboratorio",
                                size=22, weight=ft.FontWeight.W_700, color=WHITE),
                        ft.Text("Filtro por servicio; filas separadas por fecha del evento (validación/creación).",
                                size=12, color=WHITE, italic=True),
                    ],
                    spacing=2,
                ),
            ],
            spacing=12,
            vertical_alignment=ft.CrossAxisAlignment.CENTER,
        ),
    )

    # Estado
    selected_file = {"path": None}
    df_result = {"df": None}

    # Controles
    btn_select = ft.ElevatedButton("Seleccionar Excel…")
    btn_process = ft.FilledButton("Procesar")
    btn_export = ft.OutlinedButton("Exportar a Excel", disabled=True)

    file_info = ft.Text("Archivo: (ninguno)", size=12, color=TEXT_MUTED, selectable=True)
    status_ok = ft.Text("Estado: esperando acción", size=12, color=TEXT_MUTED)
    status_err = ft.Text("", size=12, color=DANGER)

    # Barra de progreso
    progress_bar = ft.ProgressBar(width=240, visible=False)
    progress_text = ft.Text("", size=12, color=TEXT_MUTED)

    # PREVIEW con scroll H/V (compat)
    table_holder_inner = ft.Row([], scroll=ft.ScrollMode.ALWAYS)  # HORIZONTAL
    preview_panel = ft.Container(
        bgcolor=WHITE,
        border=ft.border.all(1, tone("GREY_200", "#E5E7EB")),
        border_radius=16,
        padding=16,
        content=ft.Column(
            [
                ft.Text("Vista previa", size=16, weight=ft.FontWeight.W_700, color=TEXT),
                ft.Text(
                    f"Se muestran {MAX_ROWS_PREVIEW} filas y {MAX_COLS_PREVIEW} columnas (máx.) para fluidez.",
                    size=12, color=TEXT_MUTED),
                ft.Container(
                    height=420,
                    content=ft.Column([table_holder_inner], scroll=ft.ScrollMode.ALWAYS),  # VERTICAL
                ),
            ],
            spacing=8,
        ),
    )

    # Panel acciones
    actions_panel = ft.Container(
        bgcolor=WHITE,
        border=ft.border.all(1, tone("GREY_200", "#E5E7EB")),
        border_radius=16,
        padding=16,
        content=ft.Column(
            [
                ft.Text("Acciones", size=16, weight=ft.FontWeight.W_700, color=TEXT),
                btn_select, btn_process, btn_export,
                ft.Row([progress_bar, progress_text], spacing=10),
                ft.Divider(),
                file_info, status_ok, status_err,
            ],
            spacing=10,
        ),
    )

    # Footer
    footer = ft.Container(
        padding=ft.padding.only(left=16, right=16, top=6, bottom=12),
        content=ft.Row(
            [ft.Text("powered by fley – python by Alfredo H Tellez.", size=12, italic=True, color=TEXT_MUTED)],
            alignment=ft.MainAxisAlignment.CENTER,
        ),
    )

    # File pickers
    fp_open = ft.FilePicker()
    fp_save = ft.FilePicker()
    page.overlay.extend([fp_open, fp_save])

    def on_file_selected(e: ft.FilePickerResultEvent):
        status_err.value = ""
        if e.files:
            selected_file["path"] = e.files[0].path
            file_info.value = f"Archivo: {selected_file['path']}"
            status_ok.value = "Listo para procesar."
            btn_export.disabled = True
            df_result["df"] = None
            table_holder_inner.controls = []
        page.update()
    fp_open.on_result = on_file_selected

    def set_processing(is_on: bool, msg: str = ""):
        btn_select.disabled = is_on
        btn_process.disabled = is_on
        btn_export.disabled = True if is_on else btn_export.disabled
        progress_bar.visible = is_on
        progress_text.value = msg if is_on else ""
        page.update()

    # -------- Procesar --------
    def do_select(e):
        fp_open.pick_files(
            allow_multiple=False,
            allowed_extensions=["xlsx", "xls"],
            dialog_title="Selecciona el Excel con pestañas"
        )

    def do_process(e):
        status_err.value = ""
        if not selected_file["path"]:
            status_err.value = "Primero selecciona un archivo Excel."
            page.update()
            return
        try:
            set_processing(True, "Procesando datos…")
            status_ok.value = "Procesando…"

            df_all = consolidar_todas_las_hojas(selected_file["path"])
            df_result["df"] = df_all

            set_processing(False)
            if df_all.empty:
                status_ok.value = "Procesado: no se encontraron datos útiles (tras filtro por servicio)."
                btn_export.disabled = True
                table_holder_inner.controls = []
            else:
                total_rows, total_cols = df_all.shape
                page.snack_bar = ft.SnackBar(ft.Text(f"Procesado: {total_rows} filas, {total_cols} columnas"), open=True)
                status_ok.value = f"Procesado: {total_rows} filas. (Preview: {MAX_ROWS_PREVIEW} filas / {MAX_COLS_PREVIEW} columnas)"
                btn_export.disabled = False
                dt = df_to_datatable(df_all, max_rows=MAX_ROWS_PREVIEW, max_cols=MAX_COLS_PREVIEW)
                table_holder_inner.controls = [dt]
        except Exception:
            set_processing(False)
            status_err.value = "Error procesando:\n" + traceback.format_exc()
        page.update()

    # -------- Exportar --------
    def perform_export(target_path: str):
        try:
            set_processing(True, "Exportando a Excel…")
            with pd.ExcelWriter(target_path, engine="openpyxl") as writer:
                df_result["df"].to_excel(writer, sheet_name="REPORTE", index=False)
            set_processing(False)
            status_ok.value = f"Exportado: {target_path}"
            status_err.value = ""
            page.snack_bar = ft.SnackBar(ft.Text("Archivo exportado correctamente."), open=True)
        except Exception:
            set_processing(False)
            status_err.value = "Error al exportar:\n" + traceback.format_exc()
        page.update()

    def on_save_selected(e: ft.FilePickerResultEvent):
        if df_result["df"] is None or df_result["df"].empty:
            status_err.value = "No hay datos para exportar."
            page.update()
            return
        if e and getattr(e, "path", None):
            perform_export(e.path)
            return
        try:
            base = Path(selected_file["path"]).with_suffix("")
            ts = datetime.now().strftime("%Y%m%d_%H%M%S")
            target = str(base.parent / f"{base.stem}_REPORTE_UNICO_{ts}.xlsx")
        except Exception:
            target = f"REPORTE_UNICO_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
        perform_export(target)

    fp_save.on_result = on_save_selected

    def do_export(e):
        if df_result["df"] is None or df_result["df"].empty:
            status_err.value = "No hay datos para exportar."
            page.update()
            return
        try:
            base = Path(selected_file["path"]).with_suffix("")
            suggested = f"{base.stem}_REPORTE_UNICO.xlsx"
        except Exception:
            suggested = "REPORTE_UNICO.xlsx"
        # Compatibilidad Flet antiguo
        try:
            fp_save.save_file(file_name=suggested)
        except TypeError:
            try:
                fp_save.save_file()
            except Exception:
                on_save_selected(ft.FilePickerResultEvent(path=None, files=None, file_name=None, action=None))

    # Vincular
    btn_select.on_click = do_select
    btn_process.on_click = do_process
    btn_export.on_click = do_export

    # Layout
    page.add(
        ft.Container(padding=12, content=hero),
        ft.Container(
            padding=16,
            content=ft.ResponsiveRow(
                controls=[
                    ft.Container(actions_panel, col={"xs": 12, "md": 4, "lg": 3}),
                    ft.Container(preview_panel, col={"xs": 12, "md": 8, "lg": 9}),
                ],
                columns=12, spacing=16, run_spacing=16,
            ),
        ),
        footer,
    )

if __name__ == "__main__":
    ft.app(target=main)
