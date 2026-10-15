import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import threading
import time
import json
import math
import random
import re
import shutil
import socket
import subprocess
from datetime import datetime
from pathlib import Path
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.comments import Comment
import ollama

BASE_DIR     = Path(__file__).parent
RESULTS_DIR  = BASE_DIR / "resultados"
CODEBOOK_DIR = BASE_DIR / "codebook"
PROFILE_DIR  = BASE_DIR / "chrome_profile"   # sesión persistente de X
PREFS_FILE   = BASE_DIR / "preferencias.json"

# ── MOTORES DE CODIFICACIÓN ───────────────────────────────────────────────── #
# Ollama corre en esta máquina; Claude sale a internet. El motor se elige por
# ronda y las dos listas de modelos nunca se mezclan.
CLAUDE_BIN     = shutil.which("claude") or shutil.which("claude.cmd")
CLAUDE_MODELS  = ("sonnet", "opus", "fable")          # alias que acepta el CLI
CLAUDE_EFFORTS = ("low", "medium", "high", "xhigh", "max")
CLAUDE_TIMEOUT = 600                                   # segundos por lote

# ── DISEÑO (paleta fría + tipografía) ─────────────────────────────────────── #
FUENTE      = "Segoe UI"
C_ACENTO    = "#3d5a80"   # encabezado y botones de acción primaria
C_SECUND    = "#98c1d9"   # selecciones, rayas de tabla, bordes
C_PANEL     = "#e0fbfc"   # recuadros informativos y zona de configuración
C_PELIGRO   = "#0a1128"   # EXCLUSIVO: cancelar y errores (Azul de Prusia)
C_TEXTO_SEC = "#666666"   # texto secundario


def detect_ram_gb():
    """Retorna la RAM total del sistema en GB. Retorna None si no se puede detectar."""
    try:
        import ctypes
        class MEMORYSTATUSEX(ctypes.Structure):
            _fields_ = [
                ("dwLength",                ctypes.c_ulong),
                ("dwMemoryLoad",            ctypes.c_ulong),
                ("ullTotalPhys",            ctypes.c_ulonglong),
                ("ullAvailPhys",            ctypes.c_ulonglong),
                ("ullTotalPageFile",        ctypes.c_ulonglong),
                ("ullAvailPageFile",        ctypes.c_ulonglong),
                ("ullTotalVirtual",         ctypes.c_ulonglong),
                ("ullAvailVirtual",         ctypes.c_ulonglong),
                ("ullAvailExtendedVirtual", ctypes.c_ulonglong),
            ]
        stat = MEMORYSTATUSEX()
        stat.dwLength = ctypes.sizeof(stat)
        ctypes.windll.kernel32.GlobalMemoryStatusEx(ctypes.byref(stat))
        raw = stat.ullTotalPhys / (1024 ** 3)
        # Redondear al estándar más cercano (RAM siempre viene en potencias de 2)
        for estandar in [4, 8, 16, 32, 64, 128]:
            if raw <= estandar + 0.5:
                return estandar
        return round(raw)
    except Exception:
        return None


def recommended_batch(ram_gb):
    """Retorna el batch size recomendado según la RAM disponible."""
    if ram_gb is None:
        return 5
    if ram_gb < 8:
        return 1
    if ram_gb < 16:
        return 3
    if ram_gb < 32:
        return 5
    return 8


def recommended_max_ctx(ram_gb):
    """Retorna el tope de contexto (num_ctx) seguro según la RAM disponible.

    Un num_ctx mayor consume más RAM (caché KV) y va más lento, así que se
    limita según el equipo para no provocar swap.
    """
    if ram_gb is None:
        return 4096
    if ram_gb < 8:
        return 2048
    if ram_gb < 16:
        return 4096
    if ram_gb < 32:
        return 8192
    return 16384


def choose_num_ctx(needed, max_ctx):
    """Elige la ventana de contexto estándar más pequeña que cubra 'needed',
    sin exceder 'max_ctx'.

    Retorna (num_ctx, fits) donde fits=False indica que 'needed' supera el
    tope del equipo (riesgo de truncamiento → conviene avisar al usuario).
    """
    windows = [2048, 4096, 8192, 16384]
    fits = needed <= max_ctx
    for w in windows:
        if w >= needed:
            return min(w, max_ctx), fits
    return max_ctx, fits


# ── CODEBOOK ─────────────────────────────────────────────────────────────── #

def scan_codebooks():
    """Return sorted list of (display_name, path) for every .xlsx in CODEBOOK_DIR."""
    if not CODEBOOK_DIR.exists():
        return []
    return sorted(
        [(p.stem, p) for p in CODEBOOK_DIR.glob("*.xlsx")],
        key=lambda x: x[0].lower(),
    )


def load_codebook(path):
    wb = openpyxl.load_workbook(path)
    ws = wb.active
    categorias = []
    # Normalizar encabezados a minúsculas para que cat.get('codigo') funcione
    # sin importar cómo el usuario los escribió en Excel (Codigo, CODIGO, etc.)
    headers = [
        str(cell.value).strip().lower() if cell.value is not None else ""
        for cell in ws[1]
    ]
    for row in ws.iter_rows(min_row=2, values_only=True):
        if row[0]:
            cat = dict(zip(headers, row))
            categorias.append(cat)
    wb.close()
    return categorias


def validate_codebook(categorias, nombre_archivo):
    """
    Valida estructura y contenido del codebook.
    Retorna (errores_criticos: list[str], advertencias: list[str]).
    Si errores_criticos no está vacío, el análisis no debe ejecutarse.
    """
    COLS_REQUERIDAS   = ["codigo", "nombre", "descripcion", "criterios"]
    COLS_RECOMENDADAS = ["ejemplos_positivos", "ejemplos_exclusion"]

    errores = []
    avisos  = []

    if not categorias:
        errores.append("El archivo no contiene ninguna categoría (filas vacías).")
        return errores, avisos

    # ── 1. Columnas requeridas ──────────────────────────────────────────────
    presentes = [str(k).lower().strip() for k in categorias[0].keys() if k]
    faltantes  = [c for c in COLS_REQUERIDAS if c not in presentes]
    if faltantes:
        errores.append(
            f"Faltan columnas obligatorias: {', '.join(faltantes)}\n"
            f"   Columnas encontradas: {', '.join(presentes)}"
        )
        return errores, avisos   # sin estas columnas no tiene sentido seguir

    # ── 2. Columnas recomendadas ────────────────────────────────────────────
    faltantes_rec = [c for c in COLS_RECOMENDADAS if c not in presentes]
    if faltantes_rec:
        avisos.append(
            f"Columnas recomendadas ausentes: {', '.join(faltantes_rec)}\n"
            "   El análisis funcionará, pero la calidad puede ser menor\n"
            "   sin ejemplos de inclusión/exclusión."
        )

    # ── 3. Códigos vacíos o duplicados ─────────────────────────────────────
    codigos = []
    for i, cat in enumerate(categorias, 1):
        cod = str(cat.get("codigo") or "").strip()
        if not cod:
            errores.append(f"La categoría en la fila {i + 1} no tiene código.")
        else:
            codigos.append(cod)

    duplicados = {c for c in codigos if codigos.count(c) > 1}
    if duplicados:
        errores.append(
            f"Códigos duplicados encontrados: {', '.join(sorted(duplicados))}\n"
            "   Cada categoría debe tener un código único."
        )

    # ── 4. Campos críticos vacíos por categoría ─────────────────────────────
    sin_nombre     = []
    sin_descripcion = []
    sin_criterios  = []
    for cat in categorias:
        cod  = str(cat.get("codigo") or "").strip() or "?"
        if not str(cat.get("nombre") or "").strip():
            sin_nombre.append(cod)
        if not str(cat.get("descripcion") or "").strip():
            sin_descripcion.append(cod)
        if not str(cat.get("criterios") or "").strip():
            sin_criterios.append(cod)

    if sin_nombre:
        avisos.append(
            f"Categorías sin nombre: {', '.join(sin_nombre)}")
    if sin_descripcion:
        avisos.append(
            f"Categorías sin descripción: {', '.join(sin_descripcion)}")
    if sin_criterios:
        errores.append(
            f"Categorías sin criterios de inclusión: {', '.join(sin_criterios)}\n"
            "   Los criterios son la columna más importante para el análisis.\n"
            "   Sin ellos el modelo no sabe cuándo aplicar la categoría."
        )

    return errores, avisos


# ── OLLAMA PROMPT / PARSING ───────────────────────────────────────────────── #

def build_batch_prompt(items_batch, categorias):
    prompt = (
        "Eres un codificador académico especializado en análisis del discurso. "
        "Tu ÚNICA guía es el codebook que se presenta a continuación. "
        "No uses conocimiento externo ni interpretaciones propias.\n\n"
        "INSTRUCCIONES OBLIGATORIAS:\n"
        "1. Lee los criterios de inclusión y exclusión de cada categoría con máxima atención.\n"
        "2. Basa tu decisión EXCLUSIVAMENTE en los criterios del codebook.\n"
        "3. El campo 'razon' es OBLIGATORIO y debe contener siempre una explicación en español:\n"
        "   - Si valor=1: cita las palabras exactas del texto que activan la categoría y explica qué criterio cumplen.\n"
        "   - Si valor=0: explica qué criterios del codebook NO se cumplen en el texto.\n"
        "   - Si valor=\"NC\": explica qué información falta o qué criterio es insuficiente para decidir.\n"
        "   - NUNCA dejes 'razon' vacío ni con texto genérico. Siempre referencia el texto analizado.\n"
        "4. Codifica 1 SOLO si el texto cumple los criterios de inclusión. "
        "Codifica 0 si no los cumple. Si el texto genuinamente no permite decidir, "
        "responde \"NC\" — úsalo con moderación, solo cuando ni 1 ni 0 sean defendibles.\n"
        "5. Codifica CADA texto de forma independiente. No mezcles ni compares entre ellos.\n\n"
        "=== CODEBOOK ===\n\n"
    )

    for i, cat in enumerate(categorias, 1):
        prompt += f"--- CATEGORÍA {i} ---\n"
        prompt += f"Código: {cat.get('codigo') or ''}\n"
        prompt += f"Nombre: {cat.get('nombre') or ''}\n"
        prompt += f"Descripción: {cat.get('descripcion') or ''}\n"
        prompt += f"Criterios de inclusión:\n{cat.get('criterios') or ''}\n"
        prompt += f"Ejemplos que SÍ aplican:\n{cat.get('ejemplos_positivos') or ''}\n"
        prompt += f"Ejemplos que NO aplican (criterios de exclusión):\n{cat.get('ejemplos_exclusion') or ''}\n\n"

    prompt += "=== TEXTOS A ANALIZAR ===\n\n"
    for item in items_batch:
        prompt += f"[ID: {item['id']}]\n{item['texto']}\n\n"

    prompt += "=== RESPUESTA ===\n"
    prompt += "Responde ÚNICAMENTE con un bloque JSON válido. Sin texto antes ni después.\n"
    prompt += "El campo 'valor' debe ser 0, 1 o \"NC\".\n"
    prompt += "IMPORTANTE: el campo 'razon' de cada categoría DEBE contener una explicación no vacía.\n"
    prompt += "Formato exacto (un objeto por cada ID):\n{\n"

    item_entries = []
    for item in items_batch:
        cat_lineas = [
            f'    "{cat.get("codigo") or ""}": {{"valor": 0, "razon": "escribe aquí la explicación obligatoria basada en el texto"}}'
            for cat in categorias
        ]
        entry = f'  "{item["id"]}": {{\n' + ",\n".join(cat_lineas) + "\n  }"
        item_entries.append(entry)

    prompt += ",\n".join(item_entries) + "\n}\n"
    return prompt


def _try_parse_json(text):
    def clean(t):
        return re.sub(r',\s*([\}\]])', r'\1', t)
    for candidate in [text.strip(), clean(text.strip())]:
        try:
            return json.loads(candidate)
        except Exception:
            pass
    match = re.search(r'\{[\s\S]*\}', text)
    if match:
        for candidate in [match.group(), clean(match.group())]:
            try:
                return json.loads(candidate)
            except Exception:
                pass
    return None


def parse_batch_response(response_text, items_batch, categorias):
    parsed  = _try_parse_json(response_text)
    results = {}
    for item in items_batch:
        iid = item["id"]
        if parsed and iid in parsed and isinstance(parsed[iid], dict):
            item_coding  = parsed[iid]
            clean_coding = {}
            for cat in categorias:
                codigo = cat["codigo"]
                entry  = item_coding.get(codigo, {})
                if not isinstance(entry, dict):
                    entry = {"valor": entry, "razon": ""}
                raw = entry.get("valor", 0)
                # "NC" = no concluyente (ambigüedad genuina). None queda
                # reservado para error técnico (solo lo generan los except).
                if isinstance(raw, str) and raw.strip().upper() in ("NC", "N/C"):
                    valor = "NC"
                elif raw is None:
                    valor = "NC"
                else:
                    valor = int(bool(raw))
                razon = str(entry.get("razon", "")).strip()
                if not razon:
                    razon = "El modelo no proporcionó justificación para esta categoría."
                clean_coding[codigo] = {"valor": valor, "razon": razon}
            results[iid] = clean_coding
        else:
            # valor=None → dato faltante, NO una observación negativa
            results[iid] = {
                cat["codigo"]: {"valor": None, "razon": "Sin respuesta del modelo."}
                for cat in categorias
            }
    return results


# ── MOTOR CLAUDE (CLI en modo headless) ──────────────────────────────────── #

class ClaudeTerminal(Exception):
    """Fallo que no tiene sentido reintentar: cuota agotada, sesión caducada,
    modelo fuera del plan. Detiene la ronda en vez de rellenar el resto del
    corpus con datos faltantes, que produciría un Excel con aspecto de completo.
    """


# Fragmentos que identifican un fallo terminal en el mensaje del CLI.
_CLAUDE_TERMINAL = (
    "usage limit", "limit reached", "rate limit", "quota", "credit balance",
    "authenticat", "unauthorized", "oauth", "session expired", "log in",
    "invalid api key", "not found", "not available", "does not have access",
)


def _mensaje_claude(salida, stderr=""):
    """Saca el mensaje legible de una salida del CLI.

    El CLI devuelve el error dentro del campo 'result' de su envoltorio JSON,
    así que quedarse con el JSON crudo escondería la única línea útil.
    """
    try:
        d = json.loads((salida or "").strip())
        if isinstance(d, dict) and d.get("result"):
            return str(d["result"])
    except Exception:
        pass
    return (stderr or salida or "").strip()


def _error_claude(msg):
    """Traduce el mensaje del CLI a la excepción que corresponde."""
    corto = (msg or "").strip()[:220] or "falló la llamada a Claude."
    if any(k in corto.lower() for k in _CLAUDE_TERMINAL):
        return ClaudeTerminal(corto)
    return RuntimeError(corto)


def claude_generate(prompt, model, effort, timeout=CLAUDE_TIMEOUT):
    """Codifica un lote llamando a Claude Code en modo headless.

    Devuelve (texto_crudo, coste_usd). El texto lo digiere parse_batch_response
    igual que con Ollama. Lanza ClaudeTerminal si no tiene sentido reintentar.

    Los tres flags de aislamiento no son opcionales: sin ellos, el CLAUDE.md
    del usuario, sus skills y sus hooks entrarían en el contexto del
    codificador y contaminarían el dato.
    """
    if not CLAUDE_BIN:
        raise ClaudeTerminal(
            "No se encuentra el ejecutable 'claude' en el PATH de este equipo.")
    cmd = [
        CLAUDE_BIN, "-p", prompt,
        "--model", model,
        "--effort", effort,
        "--output-format", "json",
        "--no-session-persistence",   # nada persiste entre lotes
        "--setting-sources", "",      # sin settings del usuario
        "--allowedTools", "",         # sin herramientas: solo texto
    ]
    kw = {}
    if hasattr(subprocess, "CREATE_NO_WINDOW"):   # Windows: sin consola negra
        kw["creationflags"] = subprocess.CREATE_NO_WINDOW
    try:
        p = subprocess.run(
            cmd, capture_output=True, text=True, encoding="utf-8",
            errors="replace", timeout=timeout, cwd=str(BASE_DIR), **kw)
    except subprocess.TimeoutExpired:
        raise RuntimeError(
            f"Claude no respondió en {timeout} s. Prueba con un lote más pequeño.")

    salida = (p.stdout or "").strip()
    if p.returncode != 0:
        raise _error_claude(_mensaje_claude(salida, p.stderr))
    try:
        envoltorio = json.loads(salida)
    except Exception:
        return salida, 0.0       # sin envoltorio reconocible: que lo intente el parser
    if not isinstance(envoltorio, dict) or "result" not in envoltorio:
        return salida, 0.0       # es JSON, pero no el envoltorio del CLI
    if envoltorio.get("is_error"):
        raise _error_claude(str(envoltorio.get("result", "")))
    # El coste de cada llamada viene medido en el envoltorio. No es lo que el
    # usuario paga (tiene suscripción), pero es la única cifra exacta y actual
    # del consumo: el porcentaje de cuota que Claude Code cachea en disco puede
    # llevar días sin refrescarse.
    try:
        coste = float(envoltorio.get("total_cost_usd") or 0.0)
    except (TypeError, ValueError):
        coste = 0.0
    return str(envoltorio["result"]), coste


def celda_matriz(coding, codigo):
    """Valor para la matriz: 0/1, "N/C" (no concluyente) o None (celda vacía).

    None = fallo técnico (red caída, JSON roto) → dato faltante.
    "N/C" = el modelo no pudo decidir con los criterios → visible en la matriz
    (el SUM de tiene_codigo ignora texto) y excluido del κ como faltante.
    Escribir 0 en cualquiera de los dos casos sesgaría la prevalencia y el κ.
    """
    v = coding.get(codigo, {}).get("valor")
    if v is None:
        return None
    if v == "NC":
        return "N/C"
    return int(bool(v))


# ── CONFIABILIDAD / KAPPA ────────────────────────────────────────────────── #

META_COLS_HABIC = frozenset({
    "id", "fecha", "cuenta", "texto", "likes", "reposts", "respuestas",
    "url", "origen", "tiene_codigo", "notas_codificador",
})


def load_result_file(path):
    """
    Carga un resultado HABIC (.xlsx) o codificación manual en el mismo formato.
    Retorna (data_dict, cat_codes, meta_rows) donde:
      data_dict = {id: {codigo: int, ...}, ...}
      cat_codes = [str, ...]
      meta_rows = {id: {col: valor, ...}, ...}  — columnas no-categoría (ej. texto)
    """
    wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    ws = wb["Corpus"] if "Corpus" in wb.sheetnames else wb.active
    rows = list(ws.iter_rows(values_only=True))
    wb.close()
    if not rows:
        raise ValueError("El archivo no tiene datos.")
    raw_headers = [
        str(c).strip().lower() if c is not None else "" for c in rows[0]
    ]
    if "id" not in raw_headers:
        raise ValueError("El archivo no tiene columna 'id'.")
    cat_codes = [h for h in raw_headers if h and h not in META_COLS_HABIC]
    if not cat_codes:
        raise ValueError("No se encontraron columnas de categorías en el archivo.")
    id_idx    = raw_headers.index("id")
    cat_idx   = {c: raw_headers.index(c) for c in cat_codes}
    meta_cols = [h for h in raw_headers if h and h in META_COLS_HABIC and h != "id"]
    meta_idx  = {h: raw_headers.index(h) for h in meta_cols}
    data = {}
    meta = {}
    n_headers = len(raw_headers)
    for row in rows[1:]:
        if not row:
            continue
        # En modo read_only openpyxl puede recortar las celdas vacías del final
        # de una fila; rellenar evita IndexError al leer columnas posteriores.
        if len(row) < n_headers:
            row = tuple(row) + (None,) * (n_headers - len(row))
        if row[id_idx] is None:
            continue
        rid = str(row[id_idx]).strip()
        if not rid:
            continue
        data[rid] = {}
        for code, ci in cat_idx.items():
            val = row[ci]
            # Celda vacía o ilegible → None (dato faltante), nunca 0
            try:
                data[rid][code] = int(float(str(val))) if val is not None else None
            except (ValueError, TypeError):
                data[rid][code] = None
        meta[rid] = {h: row[ci] for h, ci in meta_idx.items()}
    return data, cat_codes, meta


def kappa_for_pair(pairs):
    """
    Cohen's κ para una lista de (val_a, val_b) binarios.
    Retorna (kappa, po, pe, n) o (None, None, None, 0) si n == 0.
    """
    n = len(pairs)
    if n == 0:
        return None, None, None, 0
    n_agree = sum(1 for a, b in pairs if a == b)
    po      = n_agree / n
    n_1a    = sum(a for a, b in pairs)
    n_1b    = sum(b for a, b in pairs)
    pe = (n_1a / n) * (n_1b / n) + ((n - n_1a) / n) * ((n - n_1b) / n)
    if pe >= 1.0:
        kappa = 1.0 if po >= 1.0 else 0.0
    else:
        kappa = (po - pe) / (1 - pe)
    return round(kappa, 4), round(po, 4), round(pe, 4), n


def lacy_riffe_n(N, P=0.80, error=0.05, z=1.96):
    """Tamaño de la muestra de fiabilidad según Lacy & Riffe (1996).

    n = N·P·Q / ((N−1)·SE² + P·Q), con Q = 1−P y SE = error/z

    z = 1.96 es el valor crítico de dos colas al 95%, que es lo que se entiende
    por "95% de confianza" sin más explicación. Una cola (z = 1.64) daría una
    muestra menor y es defendible en fiabilidad, donde solo importa el límite
    inferior, pero obliga a declararlo en cada informe. P conservador (más
    bajo) → muestra mayor.
    """
    if N <= 1:
        return N
    Q  = 1 - P
    SE = error / z
    n  = (N * P * Q) / ((N - 1) * SE**2 + P * Q)
    return min(N, math.ceil(n))


def kappa_interpretation(k):
    """Etiqueta de Landis & Koch (1977) para el valor κ."""
    if k is None:
        return "—"
    if k < 0:
        return "Sin acuerdo"
    if k < 0.20:
        return "Pobre"
    if k < 0.40:
        return "Leve"
    if k < 0.60:
        return "Moderado"
    if k < 0.80:
        return "Sustancial"
    return "Casi perfecto"


# ── APP ───────────────────────────────────────────────────────────────────── #

class HABICApp:
    def __init__(self, root):
        self.root = root
        self.root.title("HABIC — Herramienta de Análisis Basada en IA para la Codificación")
        self.root.geometry("600x620")
        self.root.minsize(600, 620)
        self.root.resizable(True, True)

        self.tweets          = []
        self.corpus_items    = []
        self.corpus_filepath = None
        self.codebooks       = scan_codebooks()
        self._cancel_event   = threading.Event()
        self._process_running = False   # True mientras cualquier proceso corre
        self._closing        = False   # True cuando el usuario confirmó cerrar
        self._ollama_models  = []      # lista de modelos disponibles en Ollama
        self._max_ctx        = 4096    # tope de num_ctx según RAM (se fija en setup_ui)
        self._ctx_warned     = False   # evita repetir el aviso de contexto insuficiente
        self._ronda_info     = {}      # motor/modelo/esfuerzo de la ronda en curso
        self._ronda_coste    = 0.0     # coste acumulado de la ronda (solo Claude)

        self.kappa_data_a   = None   # (data_dict, cat_codes, meta_rows)
        self.kappa_data_b   = None
        self.val_data       = None   # resultado cargado para validación humana

        try:
            self.prefs = json.loads(PREFS_FILE.read_text(encoding="utf-8"))
        except Exception:
            self.prefs = {}

        self.setup_ui()
        self.root.protocol("WM_DELETE_WINDOW", self.on_close_request)

        if not self.prefs.get("ocultar_bienvenida"):
            self.root.after(200, self._mostrar_bienvenida)

    def _mostrar_bienvenida(self):
        win = tk.Toplevel(self.root)
        win.title("Acerca de HABIC")
        win.resizable(False, False)
        win.transient(self.root)
        win.grab_set()

        banda = tk.Frame(win, bg=C_ACENTO, pady=8)
        banda.pack(fill=tk.X)
        tk.Label(banda, text="Soy HABIC", font=(FUENTE, 16, "bold"),
                 bg=C_ACENTO, fg="white").pack()
        tk.Label(banda, text="Herramienta de Análisis Basada en IA para la Codificación",
                 bg=C_ACENTO, fg="white", font=(FUENTE, 9)).pack()

        cuerpo = tk.Frame(win, padx=20, pady=12)
        cuerpo.pack(fill=tk.BOTH, expand=True)

        def bloque(titulo, texto):
            tk.Label(cuerpo, text=titulo, font=(FUENTE, 10, "bold"),
                     fg=C_ACENTO, anchor="w").pack(fill=tk.X, pady=(8, 0))
            tk.Label(cuerpo, text=texto, font=(FUENTE, 9), justify="left",
                     anchor="w", wraplength=460).pack(fill=tk.X)

        tk.Label(cuerpo, font=(FUENTE, 9), justify="left", wraplength=460,
                 text=("Mi propósito es aplicar un codebook definido por el "
                       "investigador a un corpus de textos: codificación deductiva "
                       "(análisis de contenido dirigido), reproducible y auditable."),
                 ).pack(fill=tk.X)
        bloque("Lo que hago",
               "Codificación binaria por categorías con modelos locales (Ollama) · "
               "una justificación por celda · κ de Cohen para validar el acuerdo.")
        bloque("Lo que no hago",
               "Codificación inductiva. No genero categorías: el codebook es la "
               "teoría del investigador y su calidad define la del análisis. Las "
               "«razones» que registro son ayuda de auditoría, no prueba "
               "de razonamiento.")
        bloque("Flujo recomendado",
               "1. Construir el codebook  →  2. Codificar el corpus  →  "
               "3. Validar con una muestra codificada a mano (pestaña "
               "Confiabilidad)  →  4. Usar los resultados si el κ es adecuado.")

        pie = tk.Frame(win, padx=20, pady=10)
        pie.pack(fill=tk.X)
        no_mostrar = tk.BooleanVar(value=False)
        tk.Checkbutton(pie, text="No volver a mostrar", variable=no_mostrar,
                       font=(FUENTE, 9)).pack(side=tk.LEFT)

        def cerrar():
            if no_mostrar.get():
                self.prefs["ocultar_bienvenida"] = True
                try:
                    PREFS_FILE.write_text(json.dumps(self.prefs),
                                          encoding="utf-8")
                except Exception:
                    pass
            win.destroy()

        tk.Button(pie, text="Entendido", command=cerrar,
                  font=(FUENTE, 10, "bold"), bg=C_ACENTO, fg="white",
                  relief=tk.FLAT, padx=18, pady=5,
                  cursor="hand2").pack(side=tk.RIGHT)
        win.protocol("WM_DELETE_WINDOW", cerrar)

    # ── UI PRINCIPAL ─────────────────────────────────────────────────────── #

    def setup_ui(self):
        # Estilo de las pestañas del Notebook
        style = ttk.Style()
        style.configure("TNotebook.Tab", font=(FUENTE, 10), padding=[14, 6])

        # Encabezado
        header = tk.Frame(self.root, bg=C_ACENTO, pady=8)
        header.pack(fill=tk.X)
        tk.Label(header, text="HABIC", font=(FUENTE, 20, "bold"),
                 bg=C_ACENTO, fg="white").pack()
        tk.Label(header, text="Herramienta de Análisis Basada en IA para la Codificación",
                 font=(FUENTE, 9), bg=C_ACENTO, fg="white").pack()

        # Zona de configuración unificada (Codebook / Modelo / Lote):
        # un solo panel con grid para que los tres selectores queden alineados.
        config = tk.Frame(self.root, padx=16, pady=8, bg=C_PANEL)
        config.pack(fill=tk.X)

        def _fila_config(row, etiqueta):
            tk.Label(config, text=etiqueta, font=(FUENTE, 10, "bold"),
                     bg=C_PANEL, width=10, anchor="w").grid(
                row=row, column=0, sticky="w", pady=3)
            fila = tk.Frame(config, bg=C_PANEL)
            fila.grid(row=row, column=1, sticky="w", pady=3)
            return fila

        # ── Fila 0: Codebook ────────────────────────────────────────────
        fila_cb = _fila_config(0, "Codebook:")
        self.codebook_var   = tk.StringVar()
        cb_names            = [name for name, _ in self.codebooks]
        self.codebook_combo = ttk.Combobox(
            fila_cb, textvariable=self.codebook_var,
            values=cb_names, state="readonly", width=36, font=(FUENTE, 10),
        )
        self.codebook_combo.pack(side=tk.LEFT, padx=(0, 4))
        if cb_names:
            self.codebook_combo.current(0)
        else:
            self.codebook_var.set("(sin codebooks — coloca .xlsx en HABIC/codebook/)")

        tk.Button(
            fila_cb, text="↺", font=(FUENTE, 12), relief=tk.FLAT,
            cursor="hand2", bg=C_PANEL, command=self.refresh_codebooks,
        ).pack(side=tk.LEFT)

        tk.Button(
            fila_cb, text="Gestionar Codebook",
            font=(FUENTE, 9), relief=tk.FLAT, bg=C_PANEL,
            fg=C_ACENTO, cursor="hand2",
            command=self.open_codebook_manager,
        ).pack(side=tk.LEFT, padx=(10, 0))

        # ── Fila 1: Motor ───────────────────────────────────────────────
        # Ollama corre aquí; Claude sale a internet. La lista de modelos de
        # abajo cambia entera según cuál esté pulsado: nunca se mezclan.
        fila_motor = _fila_config(1, "Motor:")
        self.motor_var = tk.StringVar(value="ollama")
        for texto, valor in (("Ollama  (local)", "ollama"),
                             ("Claude  (en línea)", "claude")):
            tk.Radiobutton(
                fila_motor, text=texto, value=valor, variable=self.motor_var,
                indicatoron=False, width=17, font=(FUENTE, 9),
                bg="#FFFFFF", fg=C_ACENTO,
                selectcolor=C_SECUND, relief=tk.RIDGE, cursor="hand2",
                command=self._aplicar_motor,
            ).pack(side=tk.LEFT, padx=(0, 4))

        self.lbl_esfuerzo = tk.Label(fila_motor, text="Esfuerzo:",
                                     font=(FUENTE, 9), bg=C_PANEL)
        self.effort_var   = tk.StringVar(
            value=self.prefs.get("claude_esfuerzo", "low"))
        self.effort_combo = ttk.Combobox(
            fila_motor, textvariable=self.effort_var,
            values=list(CLAUDE_EFFORTS), state="readonly",
            width=8, font=(FUENTE, 9),
        )

        # ── Fila 2: Modelo ──────────────────────────────────────────────
        fila_mod = _fila_config(2, "Modelo:")
        self.model_var   = tk.StringVar()
        self.model_combo = ttk.Combobox(
            fila_mod, textvariable=self.model_var,
            values=[], state="readonly", width=36, font=(FUENTE, 10),
        )
        self.model_combo.pack(side=tk.LEFT, padx=(0, 4))

        self.btn_refresh_models = tk.Button(
            fila_mod, text="↺", font=(FUENTE, 12), relief=tk.FLAT,
            cursor="hand2", bg=C_PANEL, command=self.refresh_models,
        )
        self.btn_refresh_models.pack(side=tk.LEFT)

        self.btn_compatibles = tk.Button(
            fila_mod, text="Ver modelos compatibles",
            font=(FUENTE, 9), relief=tk.FLAT, bg=C_PANEL,
            fg=C_ACENTO, cursor="hand2",
            command=self.show_compatible_models,
        )
        self.btn_compatibles.pack(side=tk.LEFT, padx=(10, 0))

        self.refresh_models()
        if not self._ollama_models and CLAUDE_BIN:
            self.motor_var.set("claude")
        self._aplicar_motor()

        # ── Fila 3: Tamaño de lote ──────────────────────────────────────
        ram_gb   = detect_ram_gb()
        rec      = recommended_batch(ram_gb)
        ram_info = f"{ram_gb} GB" if ram_gb else "no detectada"
        self._max_ctx = recommended_max_ctx(ram_gb)

        BATCH_OPTIONS = [
            (1,  "1 — Máxima precisión  (textos largos o codebooks complejos)"),
            (3,  "3 — Preciso  (recomendado para codebooks detallados)"),
            (5,  "5 — Estándar"),
            (8,  "8 — Rápido  (textos cortos y modelos potentes)"),
            (10, "10 — Máxima velocidad  (riesgo de errores con codebooks grandes)"),
        ]
        # Marcar la opción recomendada según RAM
        self._batch_values = []
        for size, desc in BATCH_OPTIONS:
            label = f"{desc}  — Recomendado para tu equipo" if size == rec else desc
            self._batch_values.append((size, label))

        fila_lote = _fila_config(3, "Lote:")
        self.batch_var = tk.StringVar()
        self.batch_combo = ttk.Combobox(
            fila_lote, textvariable=self.batch_var,
            values=[label for _, label in self._batch_values],
            state="readonly", width=54, font=(FUENTE, 9),
        )
        self.batch_combo.pack(side=tk.LEFT, padx=(0, 10))

        # Seleccionar la opción recomendada por defecto
        rec_labels = [label for size, label in self._batch_values if size == rec]
        if rec_labels:
            self.batch_var.set(rec_labels[0])

        tk.Label(fila_lote, text=f"RAM: {ram_info}",
                 font=(FUENTE, 8), fg=C_TEXTO_SEC, bg=C_PANEL).pack(side=tk.LEFT)

        # Notebook
        self.notebook = ttk.Notebook(self.root)
        self.notebook.pack(fill=tk.BOTH, expand=True, padx=8, pady=(4, 8))

        tab_twitter = tk.Frame(self.notebook, padx=16, pady=10)
        tab_corpus  = tk.Frame(self.notebook, padx=16, pady=10)
        tab_kappa   = tk.Frame(self.notebook, padx=16, pady=10)

        self.notebook.add(tab_twitter, text="  Twitter / X  ")
        self.notebook.add(tab_corpus,  text="  Cargar Excel  ")
        self.notebook.add(tab_kappa,   text="  Confiabilidad  ")

        self.setup_twitter_tab(tab_twitter)
        self.setup_excel_tab(tab_corpus)
        self.setup_kappa_tab(tab_kappa)

    # ── CIERRE SEGURO ────────────────────────────────────────────────────── #

    def on_close_request(self):
        if self._process_running:
            resp = messagebox.askyesno(
                "Proceso en curso",
                "Hay un proceso activo en este momento.\n\n"
                "Si cierras ahora, Chrome podría quedar abierto en segundo "
                "plano y el análisis en curso se perderá.\n\n"
                "¿Deseas cancelar el proceso y cerrar el programa?",
                icon="warning",
            )
            if resp:
                self._closing = True
                self._cancel_event.set()
                self.root.after(600, self.root.destroy)
        else:
            self.root.destroy()

    # ── VERIFICACIONES PREVIAS ────────────────────────────────────────────#

    def _check_ollama(self):
        """Verifica que Ollama esté corriendo. Muestra error y retorna False si no."""
        try:
            with socket.create_connection(("localhost", 11434), timeout=3):
                pass
            return True
        except Exception:
            messagebox.showerror(
                "Ollama no disponible",
                "No se pudo conectar con Ollama.\n\n"
                "Para iniciarlo:\n"
                "  1. Abre una terminal (cmd o PowerShell)\n"
                "  2. Ejecuta:  ollama serve\n\n"
                "Una vez iniciado vuelve a intentarlo.",
            )
            return False

    # ── MODELOS OLLAMA ────────────────────────────────────────────────────── #

    # Modelos conocidos y recomendados para codificación cualitativa
    COMPATIBLE_MODELS = [
        ("llama3.1:8b",  "Recomendado — equilibrio entre velocidad y calidad"),
        ("llama3.1:70b", "Alta calidad — requiere equipo potente (≥32 GB RAM)"),
        ("llama3.2:3b",  "Rápido y liviano — menor calidad de razonamiento"),
        ("llama3.2:1b",  "Muy rápido — solo para pruebas rápidas"),
        ("llama3.3:70b", "Alta calidad — requiere equipo potente (≥32 GB RAM)"),
        ("mistral:7b",   "Buena alternativa a llama — buen razonamiento"),
        ("mistral:latest","Buena alternativa a llama — buen razonamiento"),
        ("gemma2:9b",    "Buen rendimiento en español"),
        ("gemma2:27b",   "Alta calidad en español — requiere ≥16 GB RAM"),
        ("phi3:mini",    "Muy liviano — calidad limitada para análisis complejos"),
        ("phi4:14b",     "Buen razonamiento — requiere ≥10 GB RAM"),
        ("qwen2.5:7b",   "Buen soporte multilingüe incluyendo español"),
        ("qwen2.5:14b",  "Excelente soporte multilingüe — requiere ≥10 GB RAM"),
    ]
    # Modelos NO compatibles (solo embeddings o visión sin texto)
    INCOMPATIBLE_KEYWORDS = ["embed", "vision", "clip", "whisper", "stable"]

    def _aplicar_motor(self):
        """Cambia la lista de modelos al motor pulsado. Las dos nunca se mezclan."""
        if self.motor_var.get() == "claude":
            if not CLAUDE_BIN:
                messagebox.showerror(
                    "Claude no disponible",
                    "No se encuentra el comando 'claude' en este equipo.\n\n"
                    "Instala Claude Code y comprueba que esté en el PATH.")
                self.motor_var.set("ollama")
                return self._aplicar_motor()
            self.model_combo["values"] = list(CLAUDE_MODELS)
            self.model_var.set(self.prefs.get("claude_modelo")
                               or CLAUDE_MODELS[0])
            self.lbl_esfuerzo.pack(side=tk.LEFT, padx=(12, 4))
            self.effort_combo.pack(side=tk.LEFT)
            # Ambos botones son de Ollama: sin sentido con Claude
            self.btn_refresh_models.config(state=tk.DISABLED)
            self.btn_compatibles.config(state=tk.DISABLED)
        else:
            self.lbl_esfuerzo.pack_forget()
            self.effort_combo.pack_forget()
            self.btn_refresh_models.config(state=tk.NORMAL)
            self.btn_compatibles.config(state=tk.NORMAL)
            self.refresh_models()

    def refresh_models(self):
        """Consulta Ollama y actualiza el dropdown de modelos disponibles."""
        try:
            raw = [m.model for m in ollama.list().models]
        except Exception:
            raw = []

        # Filtrar modelos incompatibles (embeddings, visión pura, etc.)
        self._ollama_models = [
            m for m in raw
            if not any(k in m.lower() for k in self.INCOMPATIBLE_KEYWORDS)
        ]

        if self.motor_var.get() == "claude":
            return   # la lista visible es la de Claude; no pisarla

        if self._ollama_models:
            prev = self.model_var.get()
            self.model_combo["values"] = self._ollama_models
            if prev in self._ollama_models:
                self.model_var.set(prev)
            else:
                # Preferir llama3.1:8b si está disponible, si no el primero
                preferred = next(
                    (m for m in self._ollama_models if "llama3.1:8b" in m),
                    self._ollama_models[0],
                )
                self.model_var.set(preferred)
        else:
            self.model_combo["values"] = []
            self.model_var.set("(sin modelos — ejecuta: ollama pull llama3.1:8b)")

    def show_compatible_models(self):
        """Muestra una ventana con los modelos recomendados para HABIC."""
        win = tk.Toplevel(self.root)
        win.title("Modelos compatibles con HABIC")
        win.geometry("580x420")
        win.resizable(False, False)
        win.transient(self.root)
        win.grab_set()

        tk.Label(win, text="Modelos compatibles con HABIC",
                 font=(FUENTE, 13, "bold"), pady=10).pack()
        tk.Label(
            win, font=(FUENTE, 9), fg=C_TEXTO_SEC, justify="left",
            text=(
                "HABIC necesita modelos de texto generativo (LLM). "
                "Los marcados como recomendados son los mejores\n"
                "para análisis cualitativo riguroso. "
                "Para descargar uno: ollama pull <nombre>"
            ),
        ).pack(padx=16, pady=(0, 8))

        # Frame con scroll
        container = tk.Frame(win)
        container.pack(fill=tk.BOTH, expand=True, padx=16, pady=(0, 8))
        canvas    = tk.Canvas(container, borderwidth=0, bg="white")
        scrollbar = ttk.Scrollbar(container, orient="vertical",
                                  command=canvas.yview)
        inner = tk.Frame(canvas, bg="white")
        inner.bind("<Configure>",
                   lambda e: canvas.configure(
                       scrollregion=canvas.bbox("all")))
        canvas.create_window((0, 0), window=inner, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)
        canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

        # Encabezados
        tk.Label(inner, text="Modelo", font=(FUENTE, 10, "bold"),
                 bg=C_ACENTO, fg="white", width=22, anchor="w",
                 padx=6, pady=4).grid(row=0, column=0, sticky="ew", padx=1, pady=1)
        tk.Label(inner, text="Descripción", font=(FUENTE, 10, "bold"),
                 bg=C_ACENTO, fg="white", anchor="w",
                 padx=6, pady=4).grid(row=0, column=1, sticky="ew", padx=1, pady=1)

        installed = set(self._ollama_models)
        for i, (nombre, desc) in enumerate(self.COMPATIBLE_MODELS, 1):
            bg     = C_PANEL if i % 2 == 0 else "white"
            estado = "Instalado" if any(nombre in m for m in installed) else "No instalado"
            color  = C_ACENTO if "Instalado" in estado else C_TEXTO_SEC
            tk.Label(inner, text=nombre, font=(FUENTE, 9, "bold"),
                     bg=bg, anchor="w", padx=6, pady=5, width=22).grid(
                row=i, column=0, sticky="ew", padx=1, pady=1)
            tk.Label(inner, text=f"{desc}  —  {estado}",
                     font=(FUENTE, 9), fg=color, bg=bg,
                     anchor="w", padx=6, pady=5).grid(
                row=i, column=1, sticky="ew", padx=1, pady=1)

        tk.Label(
            win, font=(FUENTE, 8), fg=C_TEXTO_SEC,
            text="Nota: modelos de solo embeddings o visión (sin generación de texto) no son compatibles.",
        ).pack(pady=(0, 4))
        tk.Button(win, text="Cerrar", command=win.destroy,
                  font=(FUENTE, 10), relief=tk.FLAT,
                  bg=C_ACENTO, fg="white", padx=20, pady=5,
                  cursor="hand2").pack(pady=(0, 10))

    # ── CODEBOOK ─────────────────────────────────────────────────────────── #

    def refresh_codebooks(self):
        self.codebooks = scan_codebooks()
        names = [name for name, _ in self.codebooks]
        self.codebook_combo["values"] = names
        if names:
            prev = self.codebook_var.get()
            if prev in names:
                self.codebook_var.set(prev)   # conserva la selección actual
            else:
                self.codebook_combo.current(0)
        else:
            self.codebook_var.set("(sin codebooks — coloca .xlsx en HABIC/codebook/)")

    # ── GESTOR DE CODEBOOK ────────────────────────────────────────────────── #

    def open_codebook_manager(self):
        win = tk.Toplevel(self.root)
        win.title("Gestionar Codebook")
        win.geometry("860x560")
        win.resizable(True, True)
        win.transient(self.root)
        win.grab_set()

        tk.Label(win, text="Gestionar Codebook",
                 font=(FUENTE, 13, "bold"), pady=8).pack()

        # ── Sección 1: Previsualización ────────────────────────────────────
        sec1 = tk.LabelFrame(win, text="  Previsualización del codebook activo  ",
                             font=(FUENTE, 10, "bold"), padx=10, pady=8,
                             fg=C_ACENTO)
        sec1.pack(fill=tk.BOTH, expand=True, padx=16, pady=(0, 8))

        # Nombre del codebook activo
        nombre_activo = self.codebook_var.get()
        tk.Label(sec1, text=f"Codebook: {nombre_activo}",
                 font=(FUENTE, 9, "italic"), fg=C_TEXTO_SEC).pack(anchor="w", pady=(0, 6))

        # Treeview con scroll vertical Y horizontal
        cols = ("codigo", "nombre", "descripcion",
                "criterios", "ejemplos_positivos", "ejemplos_exclusion")
        headers = ("Código", "Nombre", "Descripción",
                   "Criterios de inclusión", "Ejemplos (sí)", "Ejemplos (no)")
        anchos  = (80, 130, 180, 220, 160, 160)

        tree_frame = tk.Frame(sec1)
        tree_frame.pack(fill=tk.BOTH, expand=True)

        vsb = ttk.Scrollbar(tree_frame, orient="vertical")
        hsb = ttk.Scrollbar(tree_frame, orient="horizontal")

        tree = ttk.Treeview(
            tree_frame, columns=cols, show="headings",
            yscrollcommand=vsb.set, xscrollcommand=hsb.set,
        )
        vsb.config(command=tree.yview)
        hsb.config(command=tree.xview)

        for col, header, ancho in zip(cols, headers, anchos):
            tree.heading(col, text=header)
            tree.column(col, width=ancho, minwidth=60, anchor="w")

        # Estilo alterno de filas
        tree.tag_configure("par",   background=C_PANEL)
        tree.tag_configure("impar", background="white")

        # Cargar datos del codebook activo
        path = self.get_codebook_path()
        n_cats = 0
        if path:
            try:
                categorias = load_codebook(path)
                for i, cat in enumerate(categorias):
                    tag = "par" if i % 2 == 0 else "impar"
                    tree.insert("", "end", tag=tag, values=tuple(
                        str(cat.get(c) or "").replace("\n", " ")
                        for c in cols
                    ))
                n_cats = len(categorias)
            except Exception:
                tree.insert("", "end", values=("Error al cargar el codebook", "", "", "", "", ""))
        else:
            tree.insert("", "end", values=("Sin codebook seleccionado", "", "", "", "", ""))

        tree.grid(row=0, column=0, sticky="nsew")
        vsb.grid(row=0, column=1, sticky="ns")
        hsb.grid(row=1, column=0, sticky="ew")
        tree_frame.rowconfigure(0, weight=1)
        tree_frame.columnconfigure(0, weight=1)

        tk.Label(sec1, text=f"{n_cats} categoría(s) cargada(s).",
                 font=(FUENTE, 8), fg=C_TEXTO_SEC).pack(anchor="e", pady=(4, 0))

        # ── Sección 2: Desde Excel ──────────────────────────────────────────
        sec2 = tk.LabelFrame(win, text="  Desde Excel  ",
                             font=(FUENTE, 10, "bold"), padx=14, pady=10,
                             fg=C_ACENTO)
        sec2.pack(fill=tk.X, padx=16, pady=(0, 12))

        tk.Label(
            sec2, font=(FUENTE, 9), fg=C_TEXTO_SEC,
            text="Descarga la plantilla, llénala en Excel y luego impórtala aquí.",
        ).grid(row=0, column=0, columnspan=2, sticky="w", pady=(0, 8))

        tk.Button(sec2, text="Descargar plantilla Excel",
                  font=(FUENTE, 10), relief=tk.FLAT, bg=C_ACENTO, fg="white",
                  padx=10, pady=6, cursor="hand2",
                  command=self._cb_download_template).grid(
            row=1, column=0, padx=(0, 10), pady=4)

        tk.Button(sec2, text="Importar codebook desde Excel",
                  font=(FUENTE, 10), relief=tk.FLAT, bg=C_ACENTO, fg="white",
                  padx=10, pady=6, cursor="hand2",
                  command=lambda: self._cb_import(win)).grid(
            row=1, column=1, pady=4)

    def _cb_download_template(self):
        dest = filedialog.asksaveasfilename(
            title="Guardar plantilla",
            defaultextension=".xlsx",
            filetypes=[("Excel", "*.xlsx")],
            initialfile="plantilla_codebook.xlsx",
        )
        if not dest:
            return
        columns = [
            ("codigo", 14, "Código único (ej: V01)", "V01"),
            ("nombre", 24, "Nombre corto de la categoría", "Violencia verbal"),
            ("descripcion", 36, "Qué mide esta categoría",
             "Identifica presencia de lenguaje agresivo o insultos."),
            ("criterios", 50, "Cuándo se aplica — MUY IMPORTANTE",
             "Se codifica 1 cuando el texto contiene insultos, amenazas o "
             "lenguaje despectivo dirigido a personas o grupos."),
            ("ejemplos_positivos", 40, "Textos que SÍ deben codificarse con 1",
             "«Eres un idiota» · «Te voy a destruir»"),
            ("ejemplos_exclusion", 40, "Textos que NO deben codificarse (exclusiones)",
             "«No estoy de acuerdo» (crítica sin agresión)"),
        ]
        try:
            self._write_simple_template(dest, "Codebook", columns)
            messagebox.showinfo("Plantilla descargada",
                f"Plantilla guardada en:\n{dest}\n\n"
                "Escribe tus categorías a partir de la fila 2.\n"
                "Pasa el cursor sobre cada encabezado para ver la guía y un ejemplo.\n"
                "Luego impórtala desde 'Gestionar Codebook → Importar'.")
        except PermissionError:
            messagebox.showerror("Error",
                "No se pudo guardar el archivo porque está abierto en Excel.")

    def _write_simple_template(self, dest, sheet_title, columns):
        """Escribe una plantilla Excel con una sola fila de encabezados; la guía
        y el ejemplo de cada columna van en un comentario de celda (se ven al
        pasar el cursor en Excel).

        No se escriben filas de instrucción ni de ejemplo dentro de los datos,
        para que al cargar la plantilla ya llena no se cuelen filas basura.
        columns: lista de (nombre, ancho, nota, ejemplo).
        Propaga PermissionError si el archivo está abierto en Excel.
        """
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = sheet_title

        h_font = Font(name="Arial", bold=True, color="FFFFFF", size=11)
        h_fill = PatternFill("solid", fgColor="3d5a80")
        h_alig = Alignment(horizontal="center", vertical="center", wrap_text=True)
        borde  = Border(left=Side(style="thin"), right=Side(style="thin"),
                        top=Side(style="thin"), bottom=Side(style="thin"))

        ws.row_dimensions[1].height = 30
        for ci, (nombre, ancho, nota, ejemplo) in enumerate(columns, 1):
            cell = ws.cell(row=1, column=ci, value=nombre)
            cell.font = h_font; cell.fill = h_fill
            cell.alignment = h_alig; cell.border = borde
            ws.column_dimensions[get_column_letter(ci)].width = ancho

            c = Comment(f"{nota}\n\nEjemplo: {ejemplo}", "HABIC")
            c.width, c.height = 240, 120
            cell.comment = c

        ws.freeze_panes = "A2"
        wb.save(dest)

    def _xl_download_template(self):
        dest = filedialog.asksaveasfilename(
            title="Guardar plantilla de corpus",
            defaultextension=".xlsx",
            filetypes=[("Excel", "*.xlsx")],
            initialfile="plantilla_corpus.xlsx",
        )
        if not dest:
            return
        columns = [
            ("id",     14, "Identificador único de cada texto (ej: 1, 2, 3…)", "1"),
            ("texto",  60, "El texto a codificar (trino, respuesta, párrafo…)",
             "Aquí va el texto completo que el modelo analizará."),
            ("origen", 24, "De dónde viene el texto (opcional pero recomendado)",
             "Entrevista 1"),
        ]
        try:
            self._write_simple_template(dest, "Corpus", columns)
            messagebox.showinfo("Plantilla descargada",
                f"Plantilla guardada en:\n{dest}\n\n"
                "Escribe tus datos a partir de la fila 2.\n"
                "Pasa el cursor sobre cada encabezado para ver la guía y un ejemplo.\n"
                "Luego cárgala con el botón 'Seleccionar…'.")
        except PermissionError:
            messagebox.showerror("Error",
                "No se pudo guardar el archivo porque está abierto en Excel.")

    def _kappa_download_template(self):
        dest = filedialog.asksaveasfilename(
            title="Guardar plantilla de codificación",
            defaultextension=".xlsx",
            filetypes=[("Excel", "*.xlsx")],
            initialfile="plantilla_codificacion.xlsx",
        )
        if not dest:
            return

        # Intentar derivar las categorías del codebook seleccionado
        codigos = []
        path = self.get_codebook_path()
        if path:
            try:
                cats = load_codebook(path)
                codigos = [str(c.get("codigo") or "").strip()
                           for c in cats if str(c.get("codigo") or "").strip()]
            except Exception:
                codigos = []

        columns = [
            ("id",    14, "Mismo id que el corpus, para poder comparar", "1"),
            ("texto", 50, "El texto codificado (opcional, como referencia)",
             "Aquí va el texto que se codificó."),
        ]
        if codigos:
            for cod in codigos:
                columns.append((cod, 12, "1 si aplica, 0 si no", "0"))
            nota_extra = (f"Se generó una columna por cada código del codebook "
                          f"'{self.codebook_var.get()}'.")
        else:
            for cod in ("codigo1", "codigo2"):
                columns.append(
                    (cod, 12, "Renómbrala con tu código (1 si aplica, 0 si no)", "0"))
            nota_extra = ("No había codebook seleccionado: se usaron columnas "
                          "genéricas (codigo1, codigo2). Renómbralas con tus códigos.")

        try:
            self._write_simple_template(dest, "Codificacion", columns)
            messagebox.showinfo("Plantilla descargada",
                f"Plantilla guardada en:\n{dest}\n\n"
                f"{nota_extra}\n\n"
                "Escribe tus datos a partir de la fila 2.\n"
                "Cada categoría se marca con 1 (aplica) o 0 (no aplica).")
        except PermissionError:
            messagebox.showerror("Error",
                "No se pudo guardar el archivo porque está abierto en Excel.")

    def _cb_import(self, win):
        path = filedialog.askopenfilename(
            title="Importar codebook desde Excel",
            filetypes=[("Excel", "*.xlsx"), ("Todos los archivos", "*.*")],
        )
        if not path:
            return
        path = Path(path)
        try:
            try:
                categorias = load_codebook(path)
            except PermissionError:
                messagebox.showerror("Archivo en uso",
                    f"'{path.name}' está abierto en Excel. Ciérralo e intenta de nuevo.",
                    parent=win)
                return

            errores, avisos = validate_codebook(categorias, path.name)
            if errores:
                detalle = "\n\n".join(f"• {e}" for e in errores)
                messagebox.showerror("Codebook inválido",
                    f"'{path.name}' tiene errores:\n\n{detalle}", parent=win)
                return
            if avisos:
                detalle = "\n".join(f"• {a}" for a in avisos)
                if not messagebox.askyesno("Advertencias",
                    f"'{path.name}' tiene advertencias:\n\n{detalle}\n\n"
                    "¿Importar de todas formas?", icon="warning", parent=win):
                    return

            CODEBOOK_DIR.mkdir(parents=True, exist_ok=True)
            dest = CODEBOOK_DIR / path.name
            if dest.exists() and not messagebox.askyesno(
                "Archivo existente",
                f"Ya existe '{path.name}' en codebook/.\n¿Reemplazarlo?", parent=win):
                return

            shutil.copy2(path, dest)
            self.refresh_codebooks()
            nombre = path.stem
            if nombre in [n for n, _ in self.codebooks]:
                self.codebook_var.set(nombre)
            messagebox.showinfo("Importado",
                f"'{path.name}' importado correctamente\ny seleccionado como codebook activo.",
                parent=win)
            win.destroy()

        except Exception as e:
            messagebox.showerror("Error al importar", str(e), parent=win)

    def get_codebook_path(self):
        selected = self.codebook_var.get()
        for name, path in self.codebooks:
            if name == selected:
                return path
        return None

    def load_selected_codebook(self):
        path = self.get_codebook_path()
        if path is None:
            raise Exception(
                "No hay codebook seleccionado.\n"
                "Coloca archivos .xlsx en la carpeta HABIC/codebook/ y pulsa ↺."
            )
        try:
            categorias = load_codebook(path)
        except PermissionError:
            raise Exception(
                f"No se puede leer '{path.name}' porque está abierto en Excel.\n\n"
                "Cierra el archivo en Excel y vuelve a intentarlo."
            )

        errores, avisos = validate_codebook(categorias, path.name)

        if errores:
            detalle = "\n\n".join(f"• {e}" for e in errores)
            raise Exception(
                f"El codebook '{path.name}' tiene errores que impiden el análisis:\n\n"
                f"{detalle}\n\n"
                "Corrige el archivo y vuelve a intentarlo."
            )

        if avisos:
            detalle = "\n".join(f"• {a}" for a in avisos)
            continuar = messagebox.askyesno(
                "Advertencias en el codebook",
                f"El codebook '{path.name}' se cargó con advertencias:\n\n"
                f"{detalle}\n\n"
                "El análisis puede ejecutarse, pero la calidad podría verse afectada.\n\n"
                "¿Deseas continuar de todas formas?",
                icon="warning",
            )
            if not continuar:
                raise Exception("Análisis cancelado por el usuario.")

        return categorias

    # ── AYUDA DESPLEGABLE ─────────────────────────────────────────────────#

    def _seccion_ayuda(self, parent, row, texto):
        """Botón '▸ Cómo usar esta pestaña' que muestra/oculta una guía breve."""
        btn = tk.Button(parent, text="▸ Cómo usar esta pestaña",
                        font=(FUENTE, 9), relief=tk.FLAT, fg=C_ACENTO,
                        cursor="hand2", anchor="w")
        btn.grid(row=row, column=0, columnspan=2, sticky="w", pady=(10, 0))
        lbl = tk.Label(parent, text=texto, font=(FUENTE, 9), fg=C_TEXTO_SEC,
                       justify="left")
        lbl.grid(row=row + 1, column=0, columnspan=2, sticky="ew")
        lbl.bind("<Configure>",
                 lambda e, l=lbl: l.config(wraplength=max(100, e.width - 4)))
        lbl.grid_remove()

        def alternar():
            if lbl.winfo_viewable():
                lbl.grid_remove()
                btn.config(text="▸ Cómo usar esta pestaña")
            else:
                lbl.grid()
                btn.config(text="▾ Cómo usar esta pestaña")
        btn.config(command=alternar)

    # ── TAB 1: TWITTER ───────────────────────────────────────────────────── #

    def setup_twitter_tab(self, parent):
        tk.Label(parent, text="Cuenta en X:", font=(FUENTE, 11)).grid(
            row=0, column=0, sticky="w", pady=8)
        cuenta_frame = tk.Frame(parent)
        cuenta_frame.grid(row=0, column=1, sticky="w", pady=8, padx=10)
        tk.Label(cuenta_frame, text="@", font=(FUENTE, 11)).pack(side=tk.LEFT)
        self.cuenta_var = tk.StringVar()
        tk.Entry(cuenta_frame, textvariable=self.cuenta_var,
                 font=(FUENTE, 11), width=22).pack(side=tk.LEFT)

        tk.Label(parent, text="Desde:", font=(FUENTE, 11)).grid(
            row=1, column=0, sticky="w", pady=8)
        desde_frame = tk.Frame(parent)
        desde_frame.grid(row=1, column=1, sticky="w", pady=8, padx=10)
        self.desde_var   = tk.StringVar()
        self.desde_entry = tk.Entry(desde_frame, textvariable=self.desde_var,
                                    font=(FUENTE, 11), width=14)
        self.desde_entry.pack(side=tk.LEFT)
        self.desde_entry.bind("<KeyRelease>",
            lambda e: self.autoformat_date(self.desde_var, self.desde_entry))
        tk.Label(desde_frame, text="  DD/MM/AAAA",
                 font=(FUENTE, 9), fg=C_TEXTO_SEC).pack(side=tk.LEFT)

        tk.Label(parent, text="Hasta:", font=(FUENTE, 11)).grid(
            row=2, column=0, sticky="w", pady=8)
        hasta_frame = tk.Frame(parent)
        hasta_frame.grid(row=2, column=1, sticky="w", pady=8, padx=10)
        self.hasta_var   = tk.StringVar()
        self.hasta_entry = tk.Entry(hasta_frame, textvariable=self.hasta_var,
                                    font=(FUENTE, 11), width=14)
        self.hasta_entry.pack(side=tk.LEFT)
        self.hasta_entry.bind("<KeyRelease>",
            lambda e: self.autoformat_date(self.hasta_var, self.hasta_entry))
        tk.Label(hasta_frame, text="  DD/MM/AAAA",
                 font=(FUENTE, 9), fg=C_TEXTO_SEC).pack(side=tk.LEFT)

        parent.columnconfigure(1, weight=1)

        self.tw_status_var = tk.StringVar(value="Listo para iniciar.")
        _tw_st = tk.Label(parent, textvariable=self.tw_status_var, font=(FUENTE, 10),
                          fg=C_TEXTO_SEC, justify="left")
        _tw_st.grid(row=3, column=0, columnspan=2, pady=(12, 2), sticky="ew")
        _tw_st.bind("<Configure>",
                    lambda e, l=_tw_st: l.config(wraplength=max(100, e.width - 4)))

        self.tw_progress = ttk.Progressbar(parent, mode="indeterminate")
        self.tw_progress.grid(row=4, column=0, columnspan=2, pady=4, sticky="ew")

        self.tw_count_var = tk.StringVar(value="")
        _tw_ct = tk.Label(parent, textvariable=self.tw_count_var, font=(FUENTE, 9),
                          fg=C_ACENTO, justify="left")
        _tw_ct.grid(row=5, column=0, columnspan=2, pady=2, sticky="ew")
        _tw_ct.bind("<Configure>",
                    lambda e, l=_tw_ct: l.config(wraplength=max(100, e.width - 4)))

        btn_frame = tk.Frame(parent)
        btn_frame.grid(row=6, column=0, columnspan=2, pady=(8, 0), sticky="w")

        self.btn_start = tk.Button(
            btn_frame, text="▶  Iniciar scraping",
            font=(FUENTE, 11, "bold"), bg=C_ACENTO, fg="white",
            padx=16, pady=7, cursor="hand2", relief=tk.FLAT,
            command=self.start_scraping,
        )
        self.btn_start.pack(side=tk.LEFT, padx=(0, 10))

        self.btn_code = tk.Button(
            btn_frame, text="Codificar y Exportar",
            font=(FUENTE, 11, "bold"), bg=C_ACENTO, fg="white",
            padx=16, pady=7, cursor="hand2", relief=tk.FLAT,
            state=tk.DISABLED, command=self.start_twitter_coding,
        )
        self.btn_code.pack(side=tk.LEFT, padx=(0, 10))

        self.btn_cancel = tk.Button(
            btn_frame, text="■  Cancelar proceso",
            font=(FUENTE, 11, "bold"), bg=C_PELIGRO, fg="white",
            padx=12, pady=7, cursor="hand2", relief=tk.FLAT,
            state=tk.DISABLED, command=self._cancel_process,
        )
        self.btn_cancel.pack(side=tk.LEFT)

        self._seccion_ayuda(parent, 7,
            "Recolecta trinos de una cuenta pública de X en un rango de fechas y "
            "los codifica con el codebook seleccionado arriba. Pasos: escribe la "
            "cuenta (sin @), el rango DD/MM/AAAA y pulsa Iniciar scraping — la "
            "primera vez deberás iniciar sesión en X en la ventana de Chrome; las "
            "siguientes la sesión queda guardada. Al terminar, Codificar y "
            "Exportar corre el análisis (puedes elegir varias rondas con modelos "
            "distintos) y guarda el Excel en resultados/. Recomendado: valida el "
            "resultado en la pestaña Confiabilidad antes de usarlo.")

    # ── TAB 2: CARGAR EXCEL ──────────────────────────────────────────────── #

    def setup_excel_tab(self, parent):
        # Instrucciones
        info = tk.Frame(parent, bg=C_PANEL, padx=12, pady=10,
                        relief=tk.GROOVE, bd=1)
        info.grid(row=0, column=0, columnspan=2, sticky="ew", pady=(0, 12))
        tk.Label(
            info, bg=C_PANEL, justify="left", font=(FUENTE, 9),
            text=(
                "Carga un Excel (.xlsx) con las columnas:  id  |  texto  |  origen\n"
                "HABIC codificará cada fila usando el codebook seleccionado arriba.\n"
                "Ideal para entrevistas, transcripciones de grupos focales, encuestas, etc."
            ),
        ).pack(anchor="w")

        # Selector de archivo
        tk.Label(parent, text="Archivo:", font=(FUENTE, 11)).grid(
            row=1, column=0, sticky="w", pady=8)
        file_frame = tk.Frame(parent)
        file_frame.grid(row=1, column=1, sticky="w", pady=8, padx=10)

        self.xl_file_var = tk.StringVar(value="(ninguno seleccionado)")
        tk.Label(file_frame, textvariable=self.xl_file_var, font=(FUENTE, 9),
                 fg=C_TEXTO_SEC, width=34, anchor="w").pack(side=tk.LEFT)
        tk.Button(
            file_frame, text="Seleccionar…",
            font=(FUENTE, 10), relief=tk.FLAT, bg="#E2E8F0",
            cursor="hand2", padx=8, pady=4,
            command=self.browse_corpus,
        ).pack(side=tk.LEFT, padx=(8, 0))
        tk.Button(
            file_frame, text="Plantilla",
            font=(FUENTE, 10), relief=tk.FLAT, bg=C_ACENTO, fg="white",
            cursor="hand2", padx=8, pady=4,
            command=self._xl_download_template,
        ).pack(side=tk.LEFT, padx=(8, 0))

        # Vista previa
        self.xl_preview_var = tk.StringVar(value="")
        tk.Label(parent, textvariable=self.xl_preview_var,
                 font=(FUENTE, 9), fg=C_ACENTO, justify="left").grid(
            row=2, column=0, columnspan=2, pady=(0, 2))

        parent.columnconfigure(1, weight=1)

        # Estado y progreso
        self.xl_status_var = tk.StringVar(
            value="Selecciona un archivo Excel para comenzar.")
        _xl_st = tk.Label(parent, textvariable=self.xl_status_var, font=(FUENTE, 10),
                          fg=C_TEXTO_SEC, justify="left")
        _xl_st.grid(row=3, column=0, columnspan=2, pady=(8, 2), sticky="ew")
        _xl_st.bind("<Configure>",
                    lambda e, l=_xl_st: l.config(wraplength=max(100, e.width - 4)))

        self.xl_progress = ttk.Progressbar(parent, mode="indeterminate")
        self.xl_progress.grid(row=4, column=0, columnspan=2, pady=4, sticky="ew")

        self.xl_count_var = tk.StringVar(value="")
        _xl_ct = tk.Label(parent, textvariable=self.xl_count_var, font=(FUENTE, 9),
                          fg=C_ACENTO, justify="left")
        _xl_ct.grid(row=5, column=0, columnspan=2, pady=2, sticky="ew")
        _xl_ct.bind("<Configure>",
                    lambda e, l=_xl_ct: l.config(wraplength=max(100, e.width - 4)))

        xl_btn_frame = tk.Frame(parent)
        xl_btn_frame.grid(row=6, column=0, columnspan=2, pady=(8, 0), sticky="w")

        self.btn_xl_code = tk.Button(
            xl_btn_frame, text="Codificar y Exportar",
            font=(FUENTE, 11, "bold"), bg=C_ACENTO, fg="white",
            padx=16, pady=7, cursor="hand2", relief=tk.FLAT,
            state=tk.DISABLED, command=self.start_corpus_coding,
        )
        self.btn_xl_code.pack(side=tk.LEFT, padx=(0, 10))

        self.btn_xl_cancel = tk.Button(
            xl_btn_frame, text="■  Cancelar proceso",
            font=(FUENTE, 11, "bold"), bg=C_PELIGRO, fg="white",
            padx=12, pady=7, cursor="hand2", relief=tk.FLAT,
            state=tk.DISABLED, command=self._cancel_process,
        )
        self.btn_xl_cancel.pack(side=tk.LEFT)

        self._seccion_ayuda(parent, 7,
            "Codifica cualquier corpus propio: entrevistas, transcripciones, "
            "encuestas abiertas… Pasos: descarga la Plantilla, llénala desde la "
            "fila 2 (columnas id, texto, origen), cárgala con Seleccionar y pulsa "
            "Codificar y Exportar. El resultado queda en resultados/ con una hoja "
            "de valores 0/1 y otra con las justificaciones. Recomendado: valida "
            "una muestra en la pestaña Confiabilidad antes de usar los números.")

    def browse_corpus(self):
        path = filedialog.askopenfilename(
            title="Seleccionar corpus",
            filetypes=[("Excel", "*.xlsx"), ("Todos los archivos", "*.*")],
        )
        if path:
            self.load_corpus(Path(path))

    def load_corpus(self, path):
        try:
            try:
                wb = openpyxl.load_workbook(path)
            except PermissionError:
                messagebox.showerror(
                    "Archivo en uso",
                    f"No se puede abrir '{path.name}' porque está abierto en Excel.\n\n"
                    "Cierra el archivo en Excel y vuelve a intentarlo.",
                )
                return
            ws      = wb.active
            headers = [
                str(cell.value).strip().lower() if cell.value else ""
                for cell in ws[1]
            ]

            required = {"id", "texto", "origen"}
            missing  = required - set(headers)
            if missing:
                wb.close()
                messagebox.showerror(
                    "Formato incorrecto",
                    f"El archivo no tiene las columnas requeridas.\n"
                    f"Faltan: {', '.join(sorted(missing))}\n\n"
                    f"Columnas encontradas: {', '.join(h for h in headers if h)}",
                )
                return

            idx   = {h: headers.index(h) for h in required}
            items = []
            for row_num, row in enumerate(ws.iter_rows(min_row=2, values_only=True), 2):
                texto = row[idx["texto"]]
                if not texto or not str(texto).strip():
                    continue
                raw_id = row[idx["id"]]
                item_id = (
                    str(raw_id).strip()
                    if raw_id is not None
                    else f"fila_{row_num}"
                )
                items.append({
                    "id":     item_id,
                    "texto":  str(texto).strip(),
                    "origen": str(row[idx["origen"]]).strip()
                              if row[idx["origen"]] is not None else "",
                })
            wb.close()

            if not items:
                messagebox.showwarning("Archivo vacío",
                                       "El archivo no contiene filas con datos.")
                return

            self.corpus_items    = items
            self.corpus_filepath = path
            short = path.name if len(path.name) <= 38 else path.name[:35] + "…"
            self.xl_file_var.set(short)
            self.xl_preview_var.set(
                f"{len(items)} registros cargados correctamente.")
            if not self._process_running:        # no habilitar ni pisar mensaje si hay proceso activo
                self.xl_status_var.set(
                    "Archivo listo. Selecciona un codebook y haz clic en Codificar.")
                self.btn_xl_code.config(state=tk.NORMAL)

        except Exception as e:
            messagebox.showerror("Error al cargar",
                                 f"No se pudo leer el archivo:\n{e}")

    # ── CANCELACIÓN ──────────────────────────────────────────────────────────#

    def _cancel_process(self):
        self._cancel_event.set()
        # Feedback inmediato en el botón para que el usuario sepa que se registró
        if str(self.btn_cancel["state"]) != "disabled":
            self.btn_cancel.config(text="■  Cancelando…", state=tk.DISABLED)
        if str(self.btn_xl_cancel["state"]) != "disabled":
            self.btn_xl_cancel.config(text="■  Cancelando…", state=tk.DISABLED)

    # ── CODIFICACIÓN (GENÉRICA) ───────────────────────────────────────────── #

    def _confirmar_ui(self, titulo, mensaje):
        """askyesno desde un hilo de trabajo: pregunta en el hilo de la UI
        y espera la respuesta."""
        listo = threading.Event()
        resp  = {"ok": False}

        def preguntar():
            resp["ok"] = messagebox.askyesno(titulo, mensaje)
            listo.set()

        self.root.after(0, preguntar)
        listo.wait()
        return resp["ok"]

    def _generar_raw(self, motor, model, effort, prompt, n_items, n_cats,
                     update_status_fn, run_prefix):
        """Una llamada al motor elegido. Devuelve el texto crudo de la respuesta.

        Es el único punto del programa que sabe qué motor está en uso: todo lo
        demás (prompt, parseo, exportación, κ) trabaja igual con los dos.
        """
        if motor == "claude":
            texto, coste = claude_generate(prompt, model, effort)
            self._ronda_coste += coste
            return texto

        # ── Ollama: num_ctx dinámico para que el prompt completo quepa sin
        #    truncar el codebook (limitado por la RAM del equipo) ──
        needed = int((len(prompt) // 3 + n_items * n_cats * 50) * 1.2)
        num_ctx, fits = choose_num_ctx(needed, self._max_ctx)
        if not fits and not self._ctx_warned:
            self._ctx_warned = True
            update_status_fn(
                f"{run_prefix}Atención: el codebook + lote superan el "
                f"contexto disponible para tu RAM. Reduce el tamaño "
                f"del lote para evitar pérdida de precisión."
            )
        return ollama.generate(
            model=model,
            prompt=prompt,
            format="json",                 # salida JSON válida garantizada
            options={
                "temperature": 0,          # determinista → reproducible
                "seed": 42,
                "num_ctx": num_ctx,
            },
        ).response

    def _batch_size(self):
        """Tamaño de lote elegido en la pestaña."""
        return next((size for size, label in self._batch_values
                     if label == self.batch_var.get()), 5)

    def _confirmar_control(self, hechos, total, t0, run_prefix):
        """Punto de control: enseña lo que va gastado y pregunta si sigue.

        El usuario no fija un presupuesto por adelantado, porque nadie sabe
        dimensionar eso antes de empezar. Decide con lo que ya ocurrió delante.
        """
        minutos = (time.time() - t0) / 60
        coste   = self._ronda_coste
        falta   = total - hechos
        proy    = coste * total / hechos
        min_falta = minutos * falta / hechos
        return self._confirmar_ui(
            "Punto de control",
            f"{run_prefix}Llevo {hechos} de {total} textos "
            f"({minutos:.0f} min).\n\n"
            f"Esta ronda ha consumido ${coste:.2f}.\n"
            f"A este ritmo, terminar costará unos ${proy:.2f} en total "
            f"y {min_falta:.0f} min más.\n\n"
            f"¿Continúo con los {falta} que faltan?\n\n"
            f"Si respondes No, la ronda se detiene y se guarda lo codificado "
            f"hasta aquí."
        )

    def _code_items(self, items, categorias, motor, model, effort,
                    update_status_fn, update_count_fn,
                    progress_widget, run_prefix=""):
        """Codifica 'items' con 'model' de forma síncrona (en el hilo llamante).

        Retorna (resultados, motivo) donde motivo es:
          None        → se codificó todo
          "cancelado" → el usuario pulsó Cancelar (no se exporta nada)
          "parado"    → el usuario paró en un punto de control; los textos que
                        no se alcanzaron van con valor None (dato faltante) para
                        que el archivo conserve todas las filas.
        'run_prefix' se antepone a los mensajes de estado (ej. ronda N/M).
        Propaga ClaudeTerminal: ese fallo detiene la ronda, no se reintenta.
        """
        t0 = time.time()

        def parar(motivo):
            """Rellena lo que falta como pendiente y devuelve el parcial."""
            for it in items[len(results):]:
                results.append({"item": it, "coding": {
                    cat["codigo"]: {
                        "valor": None,
                        "razon": "Pendiente: la ronda se detuvo antes de "
                                 "llegar a este texto.",
                    } for cat in categorias}})
            return results, motivo
        results   = []
        total     = len(items)
        processed = 0
        batch_size = self._batch_size()
        batches   = [items[i:i + batch_size] for i in range(0, total, batch_size)]
        n_batches = len(batches)

        # Puntos de control: solo con Claude (Ollama no consume cuota) y solo
        # si el usuario los quiere. Corpus corto, uno a la mitad; largo, tres:
        # a un cuarto, a la mitad y antes del último cuarto. Preguntar más que
        # eso cansa, y un aviso que cansa se deja de leer.
        if (motor == "claude" and self.prefs.get("control_activo", True)
                and total >= 10):
            fracciones = [0.5] if total < 100 else [0.25, 0.5, 0.75]
            controles  = [int(total * f) for f in fracciones]
        else:
            controles = []

        for batch_idx, batch in enumerate(batches, 1):
            # ── Punto de cancelación al inicio de cada lote ─────────────── #
            if self._cancel_event.is_set():
                return None, "cancelado"

            # ── Punto de control: enseñar el gasto y preguntar ──────────── #
            if controles and processed >= controles[0] and processed < total:
                # Un lote grande puede saltarse varios de golpe: se descartan
                # para no encadenar dos preguntas seguidas.
                while controles and processed >= controles[0]:
                    controles.pop(0)
                if not self._confirmar_control(processed, total, t0, run_prefix):
                    update_status_fn(
                        f"{run_prefix}Detenido a petición tuya en el texto "
                        f"{processed} de {total}. Guardando lo codificado…")
                    return parar("parado")

            b_start = processed + 1
            b_end   = processed + len(batch)
            update_status_fn(
                f"{run_prefix}Codificando lote {batch_idx} de {n_batches}  "
                f"(registros {b_start}–{b_end} de {total})…"
            )
            preview = "  |  ".join(
                t["texto"][:35].replace("\n", " ") for t in batch)
            update_count_fn(f"[Lote {batch_idx}/{n_batches}]  {preview}…")

            batch_coding = None
            for attempt in range(2):   # 0 = intento normal · 1 = reintento
                try:
                    prompt = build_batch_prompt(batch, categorias)
                    raw    = self._generar_raw(
                        motor, model, effort, prompt,
                        len(batch), len(categorias),
                        update_status_fn, run_prefix)
                    batch_coding = parse_batch_response(raw, batch, categorias)
                    break   # éxito → salir del bucle de reintentos
                except ClaudeTerminal:
                    raise   # cuota, sesión o plan: parar, no rellenar con faltantes
                except Exception as e:
                    err_lower = str(e).lower()
                    is_conn   = any(k in err_lower for k in (
                        "connection", "connect", "remote", "protocol",
                        "network", "refused", "reset", "broken pipe",
                        "no respondió", "timeout", "temporarily",
                    ))
                    if is_conn and attempt == 0:
                        # Posible suspensión del equipo — esperar y reintentar
                        quien = "Claude" if motor == "claude" else "Ollama"
                        update_status_fn(
                            f"{run_prefix}Atención: conexión con {quien} interrumpida "
                            f"(lote {batch_idx}/{n_batches}). "
                            f"Reintentando en 20 s…"
                        )
                        for _ in range(20):
                            if self._cancel_event.is_set():
                                break
                            time.sleep(1)
                        if self._cancel_event.is_set():
                            return None, "cancelado"
                        update_status_fn(
                            f"{run_prefix}Reintentando lote {batch_idx} de {n_batches}  "
                            f"(registros {b_start}–{b_end} de {total})…"
                        )
                        # continúa al intento 1
                    else:
                        # Error no de conexión, o segundo intento también falló.
                        # valor=None → dato faltante, NO una observación negativa
                        batch_coding = {
                            item["id"]: {
                                cat["codigo"]: {
                                    "valor": None,
                                    "razon": f"Error al codificar: {str(e)[:120]}",
                                }
                                for cat in categorias
                            }
                            for item in batch
                        }
                        break

            if batch_coding is None:   # salvaguarda: nunca debería ocurrir
                batch_coding = {
                    item["id"]: {
                        cat["codigo"]: {"valor": None, "razon": "Error inesperado."}
                        for cat in categorias
                    }
                    for item in batch
                }

            for item in batch:
                coding = batch_coding.get(item["id"], {
                    cat["codigo"]: {"valor": None, "razon": "Sin codificación recibida."}
                    for cat in categorias
                })
                results.append({"item": item, "coding": coding})
                processed += 1
                self.root.after(
                    0, lambda v=processed: progress_widget.config(value=v))

        # ── Segundo pase: casos no concluyentes, uno por uno ─────────────
        # Con Ollama y lote 1 el reintento sería el mismo prompt (temperature 0)
        # → misma respuesta: se salta y los NC quedan como N/C. Claude no es
        # determinista, así que ahí sí vale la pena repreguntar aunque el lote
        # ya fuera de 1.
        dudosos = [r for r in results
                   if any(e.get("valor") == "NC" for e in r["coding"].values())]
        if (dudosos and (batch_size > 1 or motor == "claude")
                and not self._cancel_event.is_set()):
            if self._confirmar_ui(
                    "Casos no concluyentes",
                    f"{run_prefix}{len(dudosos)} texto(s) quedaron con categorías "
                    f"no concluyentes (N/C).\n\n"
                    "¿Reprocesarlos uno por uno para mayor precisión?\n"
                    "(Más lento; lo que siga sin decidirse quedará como N/C.)"):
                for k, r in enumerate(dudosos, 1):
                    if self._cancel_event.is_set():
                        break
                    update_status_fn(
                        f"{run_prefix}Reprocesando casos no concluyentes "
                        f"({k}/{len(dudosos)})…")
                    try:
                        prompt = build_batch_prompt([r["item"]], categorias)
                        raw    = self._generar_raw(
                            motor, model, effort, prompt, 1, len(categorias),
                            update_status_fn, run_prefix)
                        nuevo = parse_batch_response(
                            raw, [r["item"]], categorias)[r["item"]["id"]]
                    except ClaudeTerminal:
                        raise      # cuota agotada: parar, no seguir gastando
                    except Exception:
                        continue   # fallo puntual → esas celdas siguen N/C
                    # Solo se rellenan las categorías que estaban NC;
                    # las decisiones confiadas del primer pase no se pisan.
                    for cod, entry in r["coding"].items():
                        if entry.get("valor") != "NC":
                            continue
                        cand = nuevo.get(cod, {})
                        if cand.get("valor") in (0, 1):
                            r["coding"][cod] = cand
                        else:
                            entry["razon"] = (str(entry.get("razon") or "")
                                + " (persiste tras segundo intento individual)")
                if self._cancel_event.is_set():
                    return None, "cancelado"   # cancelar es cancelar

        return results, None

    def _coding_worker(self, items, categorias, rondas,
                       update_status_fn, update_count_fn, progress_widget,
                       export_fn, on_all_done_fn, on_cancel_fn):
        """Ejecuta la codificación una vez por cada ronda, exportando tras cada
        una. Corre en un hilo de fondo.

        'rondas' es una lista de tuplas (motor, modelo, esfuerzo).
        Llama a on_all_done_fn(files, error, n_nc) al terminar (error=None si OK).
        """
        total    = len(items)
        n        = len(rondas)
        files    = []
        total_nc = 0
        for idx, (motor, model, effort) in enumerate(rondas, 1):
            if self._cancel_event.is_set():
                self.root.after(0, on_cancel_fn)
                return
            # Reiniciar la barra de progreso para esta ronda
            self.root.after(
                0, lambda t=total: progress_widget.config(
                    mode="determinate", maximum=t, value=0))
            etiqueta   = f"Claude {model}" if motor == "claude" else model
            run_prefix = f"Ronda {idx}/{n} · {etiqueta} — " if n > 1 else ""
            # La hoja Método lee de aquí la procedencia de esta ronda
            self._ronda_info  = {"motor": motor, "modelo": model,
                                 "esfuerzo": effort}
            self._ronda_coste = 0.0
            try:
                results, motivo = self._code_items(
                    items, categorias, motor, model, effort,
                    update_status_fn, update_count_fn, progress_widget,
                    run_prefix=run_prefix,
                )
            except ClaudeTerminal as e:
                self.root.after(
                    0, on_all_done_fn, files,
                    f"Claude detuvo la codificación: {e}", total_nc)
                return
            if motivo == "cancelado":
                self.root.after(0, on_cancel_fn)
                return
            self._ronda_info["parcial"] = (motivo == "parado")
            self._ronda_info["coste"]   = self._ronda_coste
            # Coste por texto de esta ronda: sirve para estimar la siguiente
            # antes de gastar nada.
            if motor == "claude" and motivo is None and items:
                self.prefs["coste_por_texto"] = self._ronda_coste / len(items)
                self._save_prefs()
            total_nc += sum(1 for r in results
                            for e in r["coding"].values()
                            if e.get("valor") == "NC")
            update_status_fn(f"{run_prefix}Generando Excel…")
            try:
                # model siempre (procedencia en la hoja Método);
                # run_idx solo con varias rondas (sufijo del nombre)
                fname = export_fn(
                    results, categorias,
                    model,
                    idx if n > 1 else None,
                )
            except Exception as e:
                self.root.after(0, on_all_done_fn, files, str(e), total_nc)
                return
            files.append(fname)
            if motivo == "parado":
                # Parar es parar: no se lanzan las rondas siguientes.
                self.root.after(0, on_all_done_fn, files, None, total_nc)
                return

        self.root.after(0, on_all_done_fn, files, None, total_nc)

    # ── CONFIGURACIÓN DE RONDAS ───────────────────────────────────────────── #

    def _save_prefs(self):
        """Guarda las preferencias. Si falla, no pasa nada: son comodidades."""
        try:
            PREFS_FILE.write_text(json.dumps(self.prefs), encoding="utf-8")
        except Exception:
            pass

    def _model_available(self, motor, name):
        """True si el modelo se puede usar con ese motor."""
        if not name:
            return False
        if motor == "claude":
            # No hay forma fiable de preguntarle al CLI qué modelos cubre el
            # plan del usuario: si el ejecutable está, se deja intentar y el
            # fallo, si lo hay, llega en la primera llamada con su mensaje.
            return bool(CLAUDE_BIN)
        try:
            available = [m.model for m in ollama.list().models]
            return any(name in a for a in available)
        except Exception:
            return True   # si no se puede comprobar, dejar que el proceso intente

    def _texto_estimacion(self, n_items):
        """Qué se sabe del gasto ANTES de empezar, que es cuando aún se puede
        decidir codificar 100 textos en vez de 800."""
        lote   = self._batch_size()
        lotes  = math.ceil(n_items / lote)
        linea  = (f"Corpus: {n_items} textos en {lotes} lote(s) de {lote}.")
        por_texto = self.prefs.get("coste_por_texto")
        if por_texto:
            linea += (f"\nSi codificas con Claude, según tu última ronda "
                      f"costará unos ${por_texto * n_items:.2f}.")
        else:
            linea += ("\nAún no hay referencia de coste para Claude: la "
                      "primera ronda te la dará.")
        if n_items >= 10:
            cuantos = "una vez, a la mitad" if n_items < 100 else "tres veces"
            linea += (f"\nCon Claude te preguntaré {cuantos} si quieres seguir, "
                      f"enseñándote lo que lleva gastado.")
        return linea

    def _aviso_claude(self, parent):
        """Aviso de tratamiento de datos, una sola vez. True si se puede seguir.

        Solo afirma hechos comprobables y remite a los términos vigentes: si el
        programa resumiera esas políticas y mañana cambiaran, quien queda
        expuesto ante un comité de ética es el investigador.
        """
        if self.prefs.get("aviso_claude_aceptado"):
            return True

        win = tk.Toplevel(parent)
        win.title("Claude es un servicio externo")
        win.resizable(False, False)
        win.transient(parent)
        win.grab_set()
        resp = {"ok": False}

        tk.Label(win, text="Claude es un servicio externo",
                 font=(FUENTE, 13, "bold"), pady=10).pack()
        cuerpo = tk.Frame(win, bg=C_PANEL, padx=16, pady=12)
        cuerpo.pack(padx=16, fill=tk.X)
        tk.Label(
            cuerpo, bg=C_PANEL, font=(FUENTE, 10), justify="left",
            text=(
                "•  El texto de tu corpus sale de este equipo y se envía a los\n"
                "    servidores de Anthropic a través de Claude Code.\n\n"
                "•  Su tratamiento queda sujeto a los términos de servicio y a la\n"
                "    política de privacidad de Anthropic vigentes en cada momento,\n"
                "    y depende del tipo de cuenta y de la configuración de tu perfil:\n"
                "        anthropic.com/legal/consumer-terms\n"
                "        anthropic.com/legal/privacy\n\n"
                "•  HABIC no puede comprobar ni garantizar esas condiciones por ti.\n\n"
                "•  Si tu material es sensible (entrevistas, sujetos protegidos,\n"
                "    datos personales), usa Ollama: se ejecuta en esta máquina y\n"
                "    el corpus no sale de ella.\n\n"
                "Los archivos que generes dejarán constancia en la hoja Método de\n"
                "que se usó un servicio externo."
            ),
        ).pack(anchor="w")

        def aceptar():
            resp["ok"] = True
            self.prefs["aviso_claude_aceptado"] = True
            self._save_prefs()
            win.destroy()

        btns = tk.Frame(win)
        btns.pack(pady=(12, 14))
        tk.Button(btns, text="Entiendo y continúo", command=aceptar,
                  font=(FUENTE, 10, "bold"), bg=C_ACENTO, fg="white",
                  relief=tk.FLAT, padx=16, pady=5,
                  cursor="hand2").pack(side=tk.LEFT, padx=6)
        tk.Button(btns, text="Cancelar", command=win.destroy,
                  font=(FUENTE, 10), relief=tk.FLAT, bg="#E2E8F0",
                  padx=16, pady=5, cursor="hand2").pack(side=tk.LEFT, padx=6)

        win.wait_window()
        return resp["ok"]

    def _open_run_config(self, default_model, n_items=0):
        """Diálogo modal: el usuario elige cuántas rondas y, en cada una, con qué
        motor y modelo codifica.

        Retorna [(motor, modelo, esfuerzo), …] o None si se canceló.
        El esfuerzo solo aplica a Claude; con Ollama va None.
        """
        win = tk.Toplevel(self.root)
        win.title("Configurar rondas de codificación")
        win.resizable(False, False)
        win.transient(self.root)
        win.grab_set()

        result = {"rondas": None}

        tk.Label(win, text="Configurar rondas de codificación",
                 font=(FUENTE, 13, "bold"), pady=8).pack()
        tk.Label(
            win, font=(FUENTE, 9), fg=C_TEXTO_SEC, justify="left",
            text=("Cada ronda codifica el MISMO corpus con el motor y el modelo\n"
                  "que elijas. Ollama corre en este equipo y es determinista\n"
                  "(temperature 0); Claude es un servicio externo y no lo es, así\n"
                  "que dos rondas suyas pueden diferir.\n"
                  "Usa varias rondas para comparar luego los resultados en la\n"
                  "pestaña Confiabilidad (Kappa)."),
        ).pack(padx=16, pady=(0, 8))

        top = tk.Frame(win)
        top.pack(padx=16, pady=(0, 6), anchor="w")
        tk.Label(top, text="Número de rondas:", font=(FUENTE, 10)).pack(side=tk.LEFT)
        n_var = tk.IntVar(value=1)
        tk.Spinbox(top, from_=1, to=6, width=4, textvariable=n_var,
                   state="readonly", font=(FUENTE, 10),
                   command=lambda: rebuild_rows()).pack(side=tk.LEFT, padx=(6, 0))

        cab = tk.Frame(win)
        cab.pack(padx=16, anchor="w")
        tk.Label(cab, font=(FUENTE, 9), fg=C_TEXTO_SEC, justify="left",
                 text="Motor                      Modelo                    "
                      "Esfuerzo (solo Claude)").pack(anchor="w", padx=(72, 0))

        if n_items:
            panel = tk.Frame(win, bg=C_PANEL, padx=10, pady=6)
            panel.pack(padx=16, pady=(8, 0), fill=tk.X)
            tk.Label(panel, font=(FUENTE, 9), justify="left", bg=C_PANEL,
                     text=self._texto_estimacion(n_items)).pack(anchor="w")
            control_var = tk.BooleanVar(
                value=bool(self.prefs.get("control_activo", True)))
            tk.Checkbutton(
                panel, variable=control_var, bg=C_PANEL, font=(FUENTE, 9),
                activebackground=C_PANEL, cursor="hand2", anchor="w",
                text=("Preguntarme por el camino cuánto llevo gastado "
                      "(solo con Claude)"),
            ).pack(anchor="w", pady=(4, 0))
        else:
            control_var = None

        rows_frame = tk.Frame(win)
        rows_frame.pack(padx=16, pady=4, fill=tk.X)
        filas = []      # [{motor, modelo, esfuerzo}, …]

        def _default_ollama():
            if default_model in self._ollama_models:
                return default_model
            return self._ollama_models[0] if self._ollama_models else ""

        def rebuild_rows():
            for w in rows_frame.winfo_children():
                w.destroy()
            filas.clear()
            for i in range(1, n_var.get() + 1):
                row = tk.Frame(rows_frame)
                row.pack(anchor="w", pady=3)
                tk.Label(row, text=f"Ronda {i}:", font=(FUENTE, 10),
                         width=9, anchor="w").pack(side=tk.LEFT)

                # Cada ronda hereda lo elegido en la pestaña; desde aquí se
                # puede cambiar para comparar motores entre rondas.
                motor_var  = tk.StringVar(value=self.motor_var.get())
                modelo_var = tk.StringVar()
                esf_var    = tk.StringVar(value=self.effort_var.get())

                radios = []
                for txt, val, activo in (
                        ("Ollama", "ollama", bool(self._ollama_models)),
                        ("Claude", "claude", bool(CLAUDE_BIN))):
                    rb = tk.Radiobutton(
                        row, text=txt, value=val, variable=motor_var,
                        font=(FUENTE, 10),
                        state=tk.NORMAL if activo else tk.DISABLED)
                    rb.pack(side=tk.LEFT)
                    radios.append(rb)

                combo = ttk.Combobox(row, textvariable=modelo_var,
                                     state="readonly", width=22,
                                     font=(FUENTE, 10))
                combo.pack(side=tk.LEFT, padx=(10, 0))
                esf = ttk.Combobox(row, textvariable=esf_var, state="readonly",
                                   width=8, font=(FUENTE, 10),
                                   values=list(CLAUDE_EFFORTS))
                esf.pack(side=tk.LEFT, padx=(8, 0))

                def cambiar_motor(mv=motor_var, md=modelo_var,
                                  cb=combo, ef=esf):
                    if mv.get() == "claude":
                        cb["values"] = list(CLAUDE_MODELS)
                        md.set(self.model_var.get()
                               if self.motor_var.get() == "claude"
                               else (self.prefs.get("claude_modelo")
                                     or CLAUDE_MODELS[0]))
                        ef.config(state="readonly")
                    else:
                        cb["values"] = self._ollama_models
                        md.set(_default_ollama())
                        ef.config(state=tk.DISABLED)

                for rb in radios:
                    rb.config(command=cambiar_motor)
                cambiar_motor()
                filas.append({"motor": motor_var, "modelo": modelo_var,
                              "esfuerzo": esf_var})

        rebuild_rows()

        def on_start():
            rondas = []
            for f in filas:
                motor  = f["motor"].get()
                modelo = f["modelo"].get().strip()
                if not modelo:
                    messagebox.showerror(
                        "Falta modelo", "Selecciona un modelo para cada ronda.",
                        parent=win)
                    return
                if not self._model_available(motor, modelo):
                    messagebox.showerror(
                        "Modelo no disponible",
                        f"'{modelo}' no está disponible para el motor "
                        f"seleccionado.",
                        parent=win)
                    return
                rondas.append((motor, modelo,
                               f["esfuerzo"].get() if motor == "claude" else None))

            if control_var is not None:
                self.prefs["control_activo"] = control_var.get()

            usa_claude = [r for r in rondas if r[0] == "claude"]
            if usa_claude:
                if not self._aviso_claude(win):
                    return
                self.prefs["claude_modelo"]   = usa_claude[-1][1]
                self.prefs["claude_esfuerzo"] = usa_claude[-1][2]
            self._save_prefs()

            result["rondas"] = rondas
            win.destroy()

        btns = tk.Frame(win)
        btns.pack(pady=(10, 12))
        tk.Button(btns, text="▶  Iniciar", command=on_start,
                  font=(FUENTE, 10, "bold"), bg=C_ACENTO, fg="white",
                  relief=tk.FLAT, padx=16, pady=5, cursor="hand2").pack(side=tk.LEFT, padx=6)
        tk.Button(btns, text="Cancelar", command=win.destroy,
                  font=(FUENTE, 10), relief=tk.FLAT, bg="#E2E8F0",
                  padx=16, pady=5, cursor="hand2").pack(side=tk.LEFT, padx=6)

        win.wait_window()
        return result["rondas"]

    # ── TWITTER: CODIFICACIÓN ────────────────────────────────────────────── #

    def _sin_motores(self):
        """True (y avisa) si no hay ningún motor con el que codificar."""
        if self._ollama_models or CLAUDE_BIN:
            return False
        messagebox.showerror(
            "Sin motores de codificación",
            "No hay ningún modelo disponible.\n\n"
            "Opciones:\n"
            "  • Ollama: descarga un modelo (ej: ollama pull llama3.1:8b) y pulsa ↺.\n"
            "  • Claude: instala Claude Code y comprueba que el comando 'claude'\n"
            "    esté en el PATH.")
        return True

    def _ollama_listo(self, rondas):
        """Comprueba Ollama solo si alguna ronda lo usa."""
        if not any(motor == "ollama" for motor, _, _ in rondas):
            return True
        return self._check_ollama()

    @staticmethod
    def _nombre_motores(rondas):
        motores = {motor for motor, _, _ in rondas}
        if motores == {"ollama"}:
            return "Ollama"
        if motores == {"claude"}:
            return "Claude"
        return "Ollama y Claude"

    def _start_coding(self, items, vacio, unidad, barra, estado_var, cuenta_var,
                      estado_fn, cuenta_fn, exportar, terminado, cancelado,
                      apagar, cancelar, otra_var, otra_msg):
        """Arranque común de las dos pestañas.

        Cambia poco entre ellas: qué datos se codifican, qué botones se apagan,
        cuál es el de cancelar y qué aviso queda en la pestaña de enfrente.
        """
        if not items:
            messagebox.showwarning("Aviso", vacio)
            return
        if self._sin_motores():
            return
        try:
            categorias = self.load_selected_codebook()
        except Exception as e:
            messagebox.showerror("Error", str(e))
            return

        rondas = self._open_run_config(self.model_var.get(), len(items))
        if not rondas:
            return
        if not self._ollama_listo(rondas):
            return

        total = len(items)
        self._cancel_event.clear()
        self._ctx_warned = False
        self._process_running = True
        for b in apagar:
            b.config(state=tk.DISABLED)
        cancelar.config(state=tk.NORMAL, text="■  Cancelar proceso")
        otra_var.set(otra_msg)
        barra.config(mode="determinate", maximum=total, value=0)
        plural = f"{len(rondas)} rondas" if len(rondas) > 1 else "1 ronda"
        estado_var.set(f"Iniciando {plural} de {total} {unidad} con "
                       f"{self._nombre_motores(rondas)}…")
        cuenta_var.set("")

        threading.Thread(
            target=self._coding_worker,
            args=(items, categorias, rondas, estado_fn, cuenta_fn,
                  barra, exportar, terminado, cancelado),
            daemon=True,
        ).start()

    def start_twitter_coding(self):
        self._start_coding(
            self.tweets, "No hay trinos para codificar.", "trinos",
            self.tw_progress, self.tw_status_var, self.tw_count_var,
            self.update_tw_status, self.update_tw_count,
            self.export_twitter_excel,
            self.on_twitter_all_done, self.on_twitter_coding_cancelled,
            apagar=(self.btn_start, self.btn_code, self.btn_xl_code),
            cancelar=self.btn_cancel,
            otra_var=self.xl_status_var,
            otra_msg=("Hay un proceso activo en la pestaña Twitter. "
                      "Espera a que termine antes de codificar."),
        )

    def _restaurar_twitter(self):
        """Estado de los botones al terminar una codificación de Twitter."""
        self._process_running = False
        self.tw_count_var.set("")
        self.btn_start.config(state=tk.NORMAL)
        self.btn_code.config(state=tk.NORMAL)
        self.btn_cancel.config(state=tk.DISABLED, text="■  Cancelar proceso")
        if self.corpus_items:
            self.btn_xl_code.config(state=tk.NORMAL)
            self.xl_status_var.set(
                "Archivo listo. Selecciona un codebook y haz clic en Codificar.")
        else:
            self.xl_status_var.set("Selecciona un archivo Excel para comenzar.")

    def _restaurar_corpus(self):
        """Estado de los botones al terminar una codificación de corpus."""
        self._process_running = False
        self.xl_count_var.set("")
        self.btn_xl_code.config(state=tk.NORMAL)
        self.btn_xl_cancel.config(state=tk.DISABLED, text="■  Cancelar proceso")
        self.btn_start.config(state=tk.NORMAL)
        if self.tweets:
            self.btn_code.config(state=tk.NORMAL)
        self.tw_status_var.set("Listo para iniciar.")

    def _codificacion_terminada(self, restaurar, avisar, files, error, n_nc):
        """Cierre común de las dos pestañas: restaura botones y avisa.

        'restaurar' deja los botones como toque en esa pestaña; 'avisar' escribe
        en su barra de estado. Todo lo demás (errores, listado de archivos,
        recuento de N/C) es idéntico en las dos.
        """
        restaurar()
        if error is not None:
            avisar("Error al exportar.")
            if not self._closing:
                guardados = ("\n\nRondas guardadas antes del error:\n"
                             + "\n".join(f"• {f}" for f in files)) if files else ""
                messagebox.showerror("Error al exportar", f"{error}{guardados}")
            return

        avisar(f"¡Listo! {len(files)} archivo(s) guardado(s) en resultados/")
        if not self._closing:
            listado = "\n".join(f"• {f}" for f in files)
            extra = ("\n\nPuedes compararlos en la pestaña Confiabilidad."
                     if len(files) > 1 else "")
            nc_msg = (f"\n\n{n_nc} celda(s) quedaron como N/C (no concluyente): "
                      "revisa sus razones — pueden señalar criterios débiles "
                      "del codebook." if n_nc else "")
            messagebox.showinfo(
                "¡Análisis completado!",
                f"{len(files)} archivo(s) generado(s):\n{listado}{extra}{nc_msg}\n\n"
                f"Ubicación:\nDocuments/HABIC/resultados/",
            )

    def _codificacion_cancelada(self, restaurar, barra):
        barra.config(value=0)
        restaurar()
        if not self._closing:
            messagebox.showinfo("Cancelado",
                                "El proceso de codificación fue cancelado.")

    def on_twitter_all_done(self, files, error=None, n_nc=0):
        self._codificacion_terminada(self._restaurar_twitter,
                                     self.update_tw_status,
                                     files, error, n_nc)

    def on_twitter_coding_cancelled(self):
        self._codificacion_cancelada(self._restaurar_twitter, self.tw_progress)

    # ── CORPUS: CODIFICACIÓN ─────────────────────────────────────────────── #

    def start_corpus_coding(self):
        self._start_coding(
            self.corpus_items, "No hay datos cargados para codificar.",
            "registros",
            self.xl_progress, self.xl_status_var, self.xl_count_var,
            self.update_xl_status, self.update_xl_count,
            self.export_corpus_excel,
            self.on_corpus_all_done, self.on_corpus_coding_cancelled,
            apagar=(self.btn_xl_code, self.btn_start, self.btn_code),
            cancelar=self.btn_xl_cancel,
            otra_var=self.tw_status_var,
            otra_msg=("Hay un proceso activo en la pestaña Corpus. "
                      "Espera a que termine antes de hacer scraping."),
        )

    def on_corpus_all_done(self, files, error=None, n_nc=0):
        self._codificacion_terminada(self._restaurar_corpus,
                                     self.update_xl_status,
                                     files, error, n_nc)

    def on_corpus_coding_cancelled(self):
        self._codificacion_cancelada(self._restaurar_corpus, self.xl_progress)

    # ── EXPORTACIÓN ───────────────────────────────────────────────────────── #

    def _excel_styles(self):
        return dict(
            h_font       = Font(name="Arial", bold=True, color="FFFFFF", size=11),
            # Sobre aguamarina (98c1d9) el blanco no contrasta: encabezado negro
            h_font_dark  = Font(name="Arial", bold=True, color="000000", size=11),
            d_font       = Font(name="Arial", size=10),
            bold_font    = Font(name="Arial", bold=True, size=10),
            fill_azul    = PatternFill("solid", fgColor="3d5a80"),
            fill_verde   = PatternFill("solid", fgColor="98c1d9"),
            fill_naranja = PatternFill("solid", fgColor="3d5a80"),
            fill_morado  = PatternFill("solid", fgColor="98c1d9"),
            center       = Alignment(horizontal="center", vertical="center",
                                     wrap_text=True),
            left         = Alignment(horizontal="left", vertical="top",
                                     wrap_text=True),
            borde        = Border(
                left=Side(style="thin"), right=Side(style="thin"),
                top=Side(style="thin"),  bottom=Side(style="thin"),
            ),
        )

    def _write_razones_sheet(self, ws2, results, categorias, s):
        """Write 'Razones' sheet — works for both Twitter and corpus results."""
        cols = (
            [("id", 18, s["fill_azul"])]
            + [(f"{cat['codigo']}_razon", 55, s["fill_morado"])
               for cat in categorias]
            + [("notas_codificador", 55, s["fill_naranja"])]
        )
        ws2.row_dimensions[1].height = 36
        for ci, (nombre, ancho, fill) in enumerate(cols, 1):
            cell = ws2.cell(row=1, column=ci, value=nombre)
            cell.font = (s["h_font_dark"]
                         if fill in (s["fill_verde"], s["fill_morado"])
                         else s["h_font"])
            cell.fill = fill
            cell.alignment = s["center"]; cell.border = s["borde"]
            ws2.column_dimensions[cell.column_letter].width = ancho

        for ri, r in enumerate(results, 2):
            it     = r["item"]
            coding = r["coding"]
            fila   = (
                [it["id"]]
                + [coding.get(cat["codigo"], {}).get("razon", "")
                   for cat in categorias]
                + [""]
            )
            ws2.row_dimensions[ri].height = 90
            for ci, valor in enumerate(fila, 1):
                cell = ws2.cell(row=ri, column=ci, value=valor)
                cell.font = s["d_font"]; cell.border = s["borde"]
                cell.alignment = s["center"] if ci == 1 else s["left"]

        ws2.freeze_panes = "A2"

    def _filas_motor(self):
        """Filas de procedencia del motor con el que se codificó la ronda.

        La hoja Método existe para que el archivo se describa a sí mismo, así
        que aquí no se puede mentir: si no hubo control de temperatura, se dice.
        """
        info  = getattr(self, "_ronda_info", None) or {}
        filas = []
        if info.get("parcial"):
            filas.append((
                "Estado",
                "PARCIAL. La ronda se detuvo antes de terminar y las filas sin "
                "codificar quedaron vacías. No confundir con celdas vacías por "
                "fallo puntual: aquí el modelo nunca llegó a ver esos textos."))
        if info.get("motor") == "claude":
            return filas + [
                ("Coste de la ronda",
                 f"${info.get('coste', 0):.4f} (equivalente medido por Claude "
                 f"Code; con suscripción no es un cargo directo)"),
                ("Parámetros de inferencia",
                 f"Claude Code en modo headless · modelo {info.get('modelo')} · "
                 f"esfuerzo {info.get('esfuerzo')} · format json. Sin control de "
                 f"temperatura ni de semilla: el resultado NO es determinista y "
                 f"dos rondas del mismo corpus pueden diferir."),
                ("Servicio externo",
                 "Sí. El texto del corpus se envió a los servidores de Anthropic "
                 "a través de Claude Code, sujeto a los términos y a la política "
                 "de privacidad de Anthropic vigentes en la fecha de este "
                 "archivo."),
            ]
        return filas + [
            ("Parámetros de inferencia",
             "temperature 0 · seed 42 · format json · num_ctx dinámico"),
            ("Servicio externo",
             "No. Ollama se ejecutó en la máquina local; el corpus no salió "
             "del equipo."),
        ]

    def _hoja_metodo(self, wb, contexto, extra=None, metodo=None):
        """Hoja 'Método': el resultado se describe a sí mismo (procedencia).

        'metodo' sustituye la descripción por defecto, que es la de un archivo
        de codificación: una muestra de validación no tiene hoja de razones ni
        la codifica un modelo, así que ahí describirla igual sería mentir.
        """
        s  = self._excel_styles()
        ws = wb.create_sheet(title="Método")
        ws.column_dimensions["A"].width = 26
        ws.column_dimensions["B"].width = 90

        ws.merge_cells("A1:B1")
        cell = ws["A1"]
        cell.value = "HABIC — Procedencia y método de este archivo"
        cell.font  = Font(name="Arial", bold=True, color="FFFFFF", size=12)
        cell.fill  = s["fill_azul"]
        cell.alignment = Alignment(horizontal="center", vertical="center")
        ws.row_dimensions[1].height = 24

        filas = [
            ("Programa", "HABIC — Herramienta de Análisis Basada en IA "
                         "para la Codificación"),
            ("Método", metodo or
                       "Codificación deductiva dirigida por codebook (análisis "
                       "de contenido dirigido). El modelo aplica los criterios "
                       "del codebook sin conocimiento externo. 1 = aplica, "
                       "0 = no aplica, celda vacía = sin dato por fallo de "
                       "codificación. La hoja 'Razones' es ayuda de auditoría, "
                       "no prueba del razonamiento del modelo."),
        ] + list(contexto.items()) + list(extra or []) + [
            ("Fecha", datetime.now().strftime("%Y-%m-%d %H:%M")),
        ]
        for ri, (lbl, val) in enumerate(filas, 3):
            c1 = ws.cell(row=ri, column=1, value=lbl)
            c1.font = s["bold_font"]; c1.alignment = s["left"]
            c1.fill = PatternFill("solid", fgColor="e0fbfc")
            c2 = ws.cell(row=ri, column=2, value=val)
            c2.font = s["d_font"]; c2.alignment = s["left"]
        ws.row_dimensions[4].height = 75   # fila "Método", texto largo

    @staticmethod
    def _safe_name(text):
        """Strip characters that are invalid in Windows filenames."""
        return re.sub(r'[<>:"/\\|?*\s]+', '_', text).strip('_') or "sin_nombre"

    def _nombre_archivo(self, prefijo, medio, model, run_idx):
        """HABIC_<prefijo>_<medio>_<codebook>[_modelo_cN][_PARCIAL]_<sello>.xlsx"""
        cb_name   = self._safe_name(self.codebook_var.get())
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        # Sufijo solo con varias rondas: con una, el nombre queda como siempre
        # (el modelo real va en la hoja Método de todos modos).
        sufijo = f"_{self._safe_name(model)}_c{run_idx}" if run_idx else ""
        # Un parcial se reconoce sin abrirlo
        if (getattr(self, "_ronda_info", None) or {}).get("parcial"):
            sufijo += "_PARCIAL"
        return f"HABIC_{prefijo}_{medio}_{cb_name}{sufijo}_{timestamp}.xlsx"

    def _export_excel(self, filename, results, categorias,
                      meta_cols, fila_meta, col_texto, contexto):
        """Escribe el .xlsx de resultados: matriz, razones y procedencia.

        Entre la pestaña de Twitter y la de corpus solo cambian tres cosas, y
        son los tres argumentos del medio: qué columnas meta hay, cómo se arma
        cada fila con ellas, y en cuál va el texto (la única alineada a la
        izquierda). El resto del archivo es idéntico en las dos.
        """
        filepath = RESULTS_DIR / filename
        RESULTS_DIR.mkdir(parents=True, exist_ok=True)
        s  = self._excel_styles()
        wb = openpyxl.Workbook()

        # Hoja 1 — Corpus
        ws1 = wb.active
        ws1.title = "Corpus"

        cat_cols = [(cat["codigo"], 16, s["fill_verde"]) for cat in categorias]
        fin_cols = [("tiene_codigo", 14, s["fill_naranja"])]
        todas    = ([(n, ancho, s["fill_azul"]) for n, ancho in meta_cols]
                    + cat_cols + fin_cols)

        n_meta      = len(meta_cols)
        n_cat       = len(cat_cols)
        col_cat_ini = n_meta + 1
        col_cat_fin = n_meta + n_cat
        col_fin     = n_meta + n_cat + 1

        ws1.row_dimensions[1].height = 36
        for ci, (nombre, ancho, fill) in enumerate(todas, 1):
            cell = ws1.cell(row=1, column=ci, value=nombre)
            cell.font = (s["h_font_dark"]
                         if fill in (s["fill_verde"], s["fill_morado"])
                         else s["h_font"])
            cell.fill = fill
            cell.alignment = s["center"]; cell.border = s["borde"]
            ws1.column_dimensions[cell.column_letter].width = ancho

        letra_ini = get_column_letter(col_cat_ini)
        letra_fin = get_column_letter(col_cat_fin)

        for ri, r in enumerate(results, 2):
            coding = r["coding"]
            fila   = (
                fila_meta(r["item"])
                + [celda_matriz(coding, cat["codigo"]) for cat in categorias]
                + [f"=IF(SUM({letra_ini}{ri}:{letra_fin}{ri})>=1,1,0)"]
            )
            ws1.row_dimensions[ri].height = 60
            for ci, valor in enumerate(fila, 1):
                cell = ws1.cell(row=ri, column=ci, value=valor)
                cell.font = s["d_font"]; cell.border = s["borde"]
                if ci == col_texto:
                    cell.alignment = s["left"]
                elif ci == col_fin:
                    cell.alignment = s["center"]
                    cell.font = s["bold_font"]
                else:
                    cell.alignment = s["center"]

        ws1.freeze_panes = "A2"

        # Hoja 2 — Razones
        ws2 = wb.create_sheet(title="Razones")
        self._write_razones_sheet(ws2, results, categorias, s)

        n_nc = sum(1 for r in results for e in r["coding"].values()
                   if e.get("valor") == "NC")
        self._hoja_metodo(wb, dict(contexto, **{
            "Registros": str(len(results)),
            "Celdas no concluyentes (N/C)": str(n_nc),
        }), extra=self._filas_motor())

        try:
            wb.save(filepath)
        except PermissionError:
            raise Exception(
                f"No se pudo guardar '{filename}' porque está abierto en Excel.\n\n"
                "Cierra el archivo en Excel y vuelve a intentarlo.\n"
                f"Ubicación: HABIC/resultados/{filename}"
            )
        return filename

    def export_twitter_excel(self, results, categorias, model=None, run_idx=None):
        cuenta = self._safe_name(self.cuenta_var.get().strip().lstrip("@"))
        desde  = self.desde_var.get().replace("/", "-")
        hasta  = self.hasta_var.get().replace("/", "-")
        return self._export_excel(
            self._nombre_archivo("TW", f"{cuenta}_{desde}_{hasta}",
                                 model, run_idx),
            results, categorias,
            [("id", 18), ("fecha", 14), ("cuenta", 18), ("texto", 62),
             ("likes", 10), ("reposts", 10), ("respuestas", 13), ("url", 42)],
            lambda t: [t["id"], t.get("fecha", ""), t.get("cuenta", ""),
                       t["texto"], t.get("likes", ""), t.get("reposts", ""),
                       t.get("respuestas", ""), t.get("url", "")],
            4,                                   # columna "texto"
            {"Fuente":   f"Twitter/X — @{self.cuenta_var.get().strip().lstrip('@')}",
             "Rango":    f"{self.desde_var.get()} a {self.hasta_var.get()}",
             "Codebook": self.codebook_var.get(),
             "Modelo":   model or "(no registrado)"},
        )

    def export_corpus_excel(self, results, categorias, model=None, run_idx=None):
        src_name = self._safe_name(
            self.corpus_filepath.stem if self.corpus_filepath else "corpus")
        return self._export_excel(
            self._nombre_archivo("CORPUS", src_name, model, run_idx),
            results, categorias,
            [("id", 18), ("texto", 62), ("origen", 28)],
            lambda it: [it["id"], it["texto"], it.get("origen", "")],
            2,                                   # columna "texto"
            {"Fuente":   f"Corpus Excel — {src_name}",
             "Codebook": self.codebook_var.get(),
             "Modelo":   model or "(no registrado)"},
        )

    # ── TWITTER: SCRAPING ────────────────────────────────────────────────── #

    def autoformat_date(self, var, entry):
        val    = var.get()
        digits = "".join(c for c in val if c.isdigit())[:8]
        fmt    = digits
        if len(digits) > 4:
            fmt = digits[:2] + "/" + digits[2:4] + "/" + digits[4:]
        elif len(digits) > 2:
            fmt = digits[:2] + "/" + digits[2:]
        if fmt != val:
            var.set(fmt)
            entry.icursor(tk.END)

    def validate_inputs(self):
        cuenta = self.cuenta_var.get().strip().lstrip("@")
        desde  = self.desde_var.get().strip()
        hasta  = self.hasta_var.get().strip()
        if not cuenta:
            messagebox.showerror("Error", "Ingresa la cuenta de X.")
            return None
        try:
            desde_dt = datetime.strptime(desde, "%d/%m/%Y")
            hasta_dt = datetime.strptime(hasta, "%d/%m/%Y")
        except ValueError:
            messagebox.showerror("Error",
                                 "Formato de fecha incorrecto.\nUsa DD/MM/AAAA")
            return None
        if desde_dt >= hasta_dt:
            messagebox.showerror("Error",
                                 "La fecha 'Desde' debe ser anterior a 'Hasta'.")
            return None
        return cuenta, desde_dt, hasta_dt

    def start_scraping(self):
        inputs = self.validate_inputs()
        if not inputs:
            return
        self.tweets = []
        self.tw_count_var.set("")
        self._cancel_event.clear()
        self._process_running = True
        self.btn_start.config(state=tk.DISABLED)
        self.btn_code.config(state=tk.DISABLED)
        self.btn_cancel.config(state=tk.NORMAL, text="■  Cancelar proceso")
        self.btn_xl_code.config(state=tk.DISABLED)
        self.xl_status_var.set(
            "Hay un proceso activo en la pestaña Twitter. "
            "Espera a que termine antes de codificar.")
        self.tw_progress.config(mode="indeterminate")
        self.tw_progress.start()
        cuenta, desde_dt, hasta_dt = inputs
        threading.Thread(
            target=self.scrape, args=(cuenta, desde_dt, hasta_dt), daemon=True
        ).start()

    def scrape(self, cuenta, desde_dt, hasta_dt):
        driver = None
        try:
            self.update_tw_status("Verificando conexión a internet…")
            try:
                with socket.create_connection(("x.com", 443), timeout=5):
                    pass
            except (socket.gaierror, socket.timeout, OSError, ConnectionRefusedError):
                raise Exception(
                    "Sin conexión a internet.\n\n"
                    "HABIC necesita conexión activa para acceder a X.\n"
                    "Verifica tu WiFi o cable de red e intenta de nuevo."
                )

            self.update_tw_status("Abriendo navegador…")
            options = Options()
            options.add_argument("--disable-notifications")
            options.add_argument("--disable-blink-features=AutomationControlled")
            options.add_experimental_option("excludeSwitches", ["enable-automation"])
            options.add_experimental_option("useAutomationExtension", False)
            # Perfil persistente: conserva la sesión de X entre ejecuciones,
            # así el usuario solo inicia sesión la primera vez.
            PROFILE_DIR.mkdir(parents=True, exist_ok=True)
            options.add_argument(f"--user-data-dir={PROFILE_DIR}")
            try:
                driver = webdriver.Chrome(options=options)
            except Exception as e:
                err = str(e).lower()
                if "user data directory" in err or "in use" in err:
                    raise Exception(
                        "La sesión anterior de Chrome sigue abierta.\n\n"
                        "Cierra esa ventana de Chrome e intenta de nuevo."
                    )
                raise

            def _is_logged(url):
                return ("x.com" in url
                        and "/login" not in url
                        and "/i/flow" not in url)

            # ── Detectar sesión existente ─────────────────────────────────
            self.update_tw_status("Verificando sesión de X…")
            driver.get("https://x.com/home")
            logged_in = False
            for _ in range(5):          # ~5 s para que X redirija si no hay sesión
                if self._cancel_event.is_set():
                    break
                time.sleep(1)
                try:
                    if _is_logged(driver.current_url):
                        logged_in = True
                        break
                except Exception:
                    pass

            if not logged_in and not self._cancel_event.is_set():
                # No hay sesión guardada → login manual (solo la primera vez)
                self.update_tw_status("Inicia sesión en X en la ventana de Chrome…")
                try:
                    driver.get("https://x.com/login")
                except Exception:
                    pass
                deadline = time.time() + 180
                while time.time() < deadline and not self._cancel_event.is_set():
                    try:
                        if _is_logged(driver.current_url):
                            logged_in = True
                            break
                    except Exception:
                        pass
                    time.sleep(2)

            if self._cancel_event.is_set():
                driver.quit()
                self.root.after(0, self.on_scrape_cancelled)
                return

            if not logged_in:
                raise Exception(
                    "Tiempo de espera de login agotado (3 minutos).\n\n"
                    "Posibles causas:\n"
                    "• La computadora entró en suspensión durante la espera.\n"
                    "• No se completó el inicio de sesión en Chrome a tiempo.\n\n"
                    "Solución: cierra la ventana de Chrome que quedó abierta "
                    "e intenta de nuevo."
                )

            self.update_tw_status("Sesión activa. Buscando trinos…")
            desde_str  = desde_dt.strftime("%Y-%m-%d")
            hasta_str  = hasta_dt.strftime("%Y-%m-%d")
            search_url = (
                f"https://x.com/search?q=from%3A{cuenta}"
                f"+since%3A{desde_str}+until%3A{hasta_str}"
                f"+-filter%3Aretweets&src=typed_query&f=live"
            )
            tweets   = []
            seen_ids = set()

            for pasada in range(1, 4):
                if self._cancel_event.is_set():
                    break
                self.update_tw_status(f"Recolectando — pasada {pasada} de 3…")
                driver.get(search_url)
                time.sleep(4)
                last_height = 0
                no_change   = 0

                while no_change < 5:
                    if self._cancel_event.is_set():
                        break
                    for article in driver.find_elements(
                            By.CSS_SELECTOR, 'article[data-testid="tweet"]'):
                        try:
                            data = self.extract_tweet(article, cuenta)
                            if data and data["id"] not in seen_ids:
                                seen_ids.add(data["id"])
                                tweets.append(data)
                                self.update_tw_count(
                                    f"Trinos recolectados: {len(tweets)}")
                        except Exception:
                            continue

                    driver.execute_script(
                        "window.scrollTo(0, document.body.scrollHeight);")
                    time.sleep(2.5)
                    new_height = driver.execute_script(
                        "return document.body.scrollHeight")
                    if new_height == last_height:
                        no_change += 1
                    else:
                        no_change = 0
                    last_height = new_height

            # ── Fase 2: recuperar el texto completo de los trinos largos ──
            # "Mostrar más" del timeline no expande el texto: el completo solo
            # está en la página individual del trino.
            pendientes = [t for t in tweets
                          if t.get("_truncado") and t.get("url")]
            n_pend = len(pendientes)
            for k, t in enumerate(pendientes, 1):
                if self._cancel_event.is_set():
                    break
                self.update_tw_status(
                    f"Recuperando textos completos ({k}/{n_pend})…")
                try:
                    driver.get(t["url"])
                    time.sleep(3)
                    # En la página de detalle puede haber varios tweetText
                    # (si el trino es una respuesta, el padre aparece arriba).
                    # Identificar el correcto por su inicio: el texto truncado
                    # del timeline es el comienzo del texto completo.
                    huella = t["texto"][:40]
                    completo = None
                    for el in driver.find_elements(
                            By.CSS_SELECTOR, '[data-testid="tweetText"]'):
                        try:
                            txt = el.text
                        except Exception:
                            continue
                        if not huella or txt.startswith(huella):
                            completo = txt
                            break
                    if completo and len(completo) > len(t["texto"]):
                        t["texto"] = completo
                except Exception:
                    continue   # trino borrado o error puntual → conservar parcial

            for t in tweets:
                t.pop("_truncado", None)

            driver.quit()
            if self._cancel_event.is_set():
                self.root.after(0, self.on_scrape_cancelled)
                return
            self.tweets = tweets
            self.root.after(0, self.on_scrape_done, len(tweets))

        except Exception as e:
            try:
                driver.quit()
            except Exception:
                pass
            err_str = str(e).lower()
            if any(k in err_str for k in (
                "no such window", "disconnected", "chrome not reachable",
                "session deleted", "target closed",
            )):
                msg = (
                    "La ventana de Chrome fue cerrada durante el scraping.\n\n"
                    "Para recolectar trinos es necesario mantener Chrome abierto "
                    "hasta que el proceso termine."
                )
            else:
                msg = str(e)
            self.root.after(0, self.on_scrape_error, msg)

    def extract_tweet(self, article, cuenta):
        tweet_url, tweet_id = "", ""
        for link in article.find_elements(
                By.CSS_SELECTOR, 'a[href*="/status/"]'):
            href = link.get_attribute("href") or ""
            if "/status/" in href and cuenta.lower() in href.lower():
                tweet_url = href
                tweet_id  = href.split("/status/")[-1].split("?")[0].split("/")[0]
                break
        if not tweet_id:
            return None

        try:
            texto = article.find_element(
                By.CSS_SELECTOR, '[data-testid="tweetText"]').text
        except Exception:
            texto = ""

        try:
            fecha = article.find_element(
                By.CSS_SELECTOR, "time").get_attribute("datetime")[:10]
        except Exception:
            fecha = ""

        def metric(testid):
            try:
                return (article.find_element(
                    By.CSS_SELECTOR, f'[data-testid="{testid}"]'
                ).text.strip() or "0")
            except Exception:
                return "0"

        # Trino largo cortado en el timeline ("Mostrar más"): el texto completo
        # solo está en la página individual — se recupera en una segunda fase.
        truncado = bool(article.find_elements(
            By.CSS_SELECTOR, '[data-testid="tweet-text-show-more-link"]'))

        return {
            "id":         tweet_id,
            "fecha":      fecha,
            "cuenta":     f"@{cuenta}",
            "texto":      texto,
            "likes":      metric("like"),
            "reposts":    metric("retweet"),
            "respuestas": metric("reply"),
            "url":        tweet_url,
            "_truncado":  truncado,
        }

    def _restore_after_scrape(self):
        """Restaura el estado de ambas pestañas tras finalizar el scraping."""
        self._process_running = False
        self.tw_progress.stop()
        self.btn_start.config(state=tk.NORMAL)
        self.btn_cancel.config(state=tk.DISABLED, text="■  Cancelar proceso")
        self.tw_count_var.set("")
        if self.corpus_items:
            self.btn_xl_code.config(state=tk.NORMAL)
            self.xl_status_var.set(
                "Archivo listo. Selecciona un codebook y haz clic en Codificar.")
        else:
            self.xl_status_var.set(
                "Selecciona un archivo Excel para comenzar.")

    def on_scrape_done(self, total):
        self._restore_after_scrape()
        if total:
            self.tw_status_var.set(
                f"¡Completado! {total} trinos recolectados. Listo para codificar.")
            self.btn_code.config(state=tk.NORMAL)
        else:
            self.tw_status_var.set(
                "Búsqueda completada — no se encontraron trinos en ese período.")

    def on_scrape_cancelled(self):
        self._restore_after_scrape()
        self.tw_status_var.set("Proceso cancelado.")
        if not self._closing:
            messagebox.showinfo(
                "Cancelado",
                "El scraping fue cancelado.\n"
                "Los trinos recolectados hasta ese momento no se guardaron.")

    def on_scrape_error(self, msg):
        self._restore_after_scrape()
        self.tw_status_var.set("Error durante el scraping.")
        if not self._closing:
            messagebox.showerror("Error", f"Ocurrió un error:\n{msg}")

    # ── TAB 3: CONFIABILIDAD (KAPPA) ─────────────────────────────────────── #

    def setup_kappa_tab(self, parent):
        parent.columnconfigure(0, weight=1)

        # ── Sección 1: preparar la validación humana ──────────────────────
        sec1 = tk.LabelFrame(parent, text="  1 · Preparar la validación humana  ",
                             font=(FUENTE, 10, "bold"), fg=C_ACENTO,
                             padx=12, pady=8)
        sec1.grid(row=0, column=0, sticky="ew", pady=(0, 10))
        sec1.columnconfigure(1, weight=1)

        tk.Label(sec1, text="Resultado de HABIC:", font=(FUENTE, 10)).grid(
            row=0, column=0, sticky="w")
        fila_val = tk.Frame(sec1)
        fila_val.grid(row=0, column=1, sticky="w", padx=(8, 0))
        self.val_file_var = tk.StringVar(value="(ninguno)")
        tk.Label(fila_val, textvariable=self.val_file_var, font=(FUENTE, 9),
                 fg=C_TEXTO_SEC, width=30, anchor="w").pack(side=tk.LEFT)
        tk.Button(fila_val, text="Seleccionar…",
                  font=(FUENTE, 10), relief=tk.FLAT, bg="#E2E8F0",
                  cursor="hand2", padx=8, pady=4,
                  command=self._browse_validacion).pack(side=tk.LEFT, padx=(8, 0))

        fila_p = tk.Frame(sec1)
        fila_p.grid(row=1, column=0, columnspan=2, sticky="w", pady=(6, 0))
        tk.Label(fila_p, text="Acuerdo esperado (P):",
                 font=(FUENTE, 10)).pack(side=tk.LEFT)
        self.val_p_var = tk.StringVar(value="0.80")
        tk.Spinbox(fila_p, from_=0.50, to=0.95, increment=0.05, width=6,
                   textvariable=self.val_p_var, state="readonly",
                   font=(FUENTE, 10), format="%.2f",
                   command=self._recalc_validacion).pack(side=tk.LEFT, padx=(6, 14))
        self.val_n_var = tk.StringVar(value="")
        tk.Label(fila_p, textvariable=self.val_n_var,
                 font=(FUENTE, 10, "bold"), fg=C_ACENTO).pack(side=tk.LEFT)

        self.btn_val_gen = tk.Button(
            sec1, text="Generar archivos de validación",
            font=(FUENTE, 10, "bold"), bg=C_ACENTO, fg="white",
            padx=12, pady=5, cursor="hand2", relief=tk.FLAT,
            state=tk.DISABLED, command=self._generar_validacion,
        )
        self.btn_val_gen.grid(row=2, column=0, columnspan=2, sticky="w", pady=(8, 2))
        tk.Label(sec1, font=(FUENTE, 9), fg=C_TEXTO_SEC, justify="left",
                 text=("Genera dos Excel con la misma muestra aleatoria: HUMANO (vacío, "
                       "codifícalo SIN mirar el IA)\ny IA (con los códigos del modelo). "
                       "Muestra según Lacy & Riffe (1996), 95% de confianza "
                       "(dos colas, z = 1,96), margen ±0,05."),
                 ).grid(row=3, column=0, columnspan=2, sticky="w")

        # ── Sección 2: calcular el acuerdo ────────────────────────────────
        sec2 = tk.LabelFrame(parent, text="  2 · Calcular el acuerdo (κ de Cohen)  ",
                             font=(FUENTE, 10, "bold"), fg=C_ACENTO,
                             padx=12, pady=8)
        sec2.grid(row=1, column=0, sticky="ew")
        sec2.columnconfigure(1, weight=1)

        # Resultado A
        tk.Label(sec2, text="Resultado A:", font=(FUENTE, 10, "bold"),
                 fg=C_ACENTO).grid(row=0, column=0, sticky="w", pady=(4, 0))
        frame_a = tk.Frame(sec2)
        frame_a.grid(row=0, column=1, sticky="w", pady=(4, 0), padx=(10, 0))
        self.kappa_file_a_var = tk.StringVar(value="(ninguno)")
        tk.Label(frame_a, textvariable=self.kappa_file_a_var,
                 font=(FUENTE, 9), fg=C_TEXTO_SEC, width=30, anchor="w").pack(side=tk.LEFT)
        tk.Button(frame_a, text="Seleccionar…",
                  font=(FUENTE, 10), relief=tk.FLAT, bg="#E2E8F0",
                  cursor="hand2", padx=8, pady=4,
                  command=lambda: self._browse_kappa_file("a")).pack(side=tk.LEFT, padx=(8, 0))
        self.kappa_prev_a_var = tk.StringVar(value="")
        tk.Label(sec2, textvariable=self.kappa_prev_a_var,
                 font=(FUENTE, 9), fg=C_ACENTO, justify="left").grid(
            row=1, column=0, columnspan=2, sticky="w", pady=(0, 4))

        # Resultado B
        tk.Label(sec2, text="Resultado B:", font=(FUENTE, 10, "bold"),
                 fg=C_ACENTO).grid(row=2, column=0, sticky="w", pady=(4, 0))
        frame_b = tk.Frame(sec2)
        frame_b.grid(row=2, column=1, sticky="w", pady=(4, 0), padx=(10, 0))
        self.kappa_file_b_var = tk.StringVar(value="(ninguno)")
        tk.Label(frame_b, textvariable=self.kappa_file_b_var,
                 font=(FUENTE, 9), fg=C_TEXTO_SEC, width=30, anchor="w").pack(side=tk.LEFT)
        tk.Button(frame_b, text="Seleccionar…",
                  font=(FUENTE, 10), relief=tk.FLAT, bg="#E2E8F0",
                  cursor="hand2", padx=8, pady=4,
                  command=lambda: self._browse_kappa_file("b")).pack(side=tk.LEFT, padx=(8, 0))
        self.kappa_prev_b_var = tk.StringVar(value="")
        tk.Label(sec2, textvariable=self.kappa_prev_b_var,
                 font=(FUENTE, 9), fg=C_ACENTO, justify="left").grid(
            row=3, column=0, columnspan=2, sticky="w", pady=(0, 8))

        # Botón calcular + acciones secundarias
        fila_btn = tk.Frame(sec2)
        fila_btn.grid(row=4, column=0, columnspan=2, sticky="w", pady=(0, 4))
        self.btn_kappa_calc = tk.Button(
            fila_btn, text="Calcular Kappa",
            font=(FUENTE, 11, "bold"), bg=C_ACENTO, fg="white",
            padx=16, pady=7, cursor="hand2", relief=tk.FLAT,
            state=tk.DISABLED, command=self.calculate_kappa,
        )
        self.btn_kappa_calc.pack(side=tk.LEFT, padx=(0, 12))
        tk.Button(
            fila_btn, text="Descargar plantilla de codificación",
            font=(FUENTE, 9), relief=tk.FLAT, bg="#E2E8F0",
            cursor="hand2", padx=8, pady=4,
            command=self._kappa_download_template,
        ).pack(side=tk.LEFT)

        tk.Label(sec2, font=(FUENTE, 9), fg=C_TEXTO_SEC, justify="left",
                 text=("Acepta resultados de HABIC o Excel con columnas id | código1 | … "
                       "Los resultados abren en ventana aparte."),
                 ).grid(row=5, column=0, columnspan=2, sticky="w")

        self._seccion_ayuda(parent, 2,
            "Mide qué tanto puedes confiar en la codificación. Flujo: en la "
            "sección 1 carga un resultado de HABIC — te digo cuántos textos "
            "codificar a mano (fórmula de Lacy & Riffe) y genero dos archivos "
            "con la misma muestra: codifica el HUMANO sin mirar el IA "
            "(codificación ciega). En la sección 2 carga IA como A y HUMANO "
            "como B y pulsa Calcular Kappa. Umbral usual: κ ≥ 0.60 (acuerdo "
            "sustancial) para usar la codificación de la IA; si sale menor, "
            "revisa los criterios del codebook. También sirve para comparar dos "
            "modelos o dos codificadores humanos.")

    # ── VALIDACIÓN HUMANA (sección 1) ─────────────────────────────────────#

    def _browse_validacion(self):
        path = filedialog.askopenfilename(
            title="Seleccionar resultado de HABIC a validar",
            filetypes=[("Excel", "*.xlsx"), ("Todos los archivos", "*.*")],
        )
        if not path:
            return
        path = Path(path)
        try:
            data, cat_codes, meta = load_result_file(path)
        except Exception as e:
            messagebox.showerror("Error al cargar", str(e))
            return
        if not any(m.get("texto") for m in meta.values()):
            messagebox.showerror(
                "Sin columna de texto",
                "El archivo no contiene la columna 'texto'.\n\n"
                "Para la validación humana se necesita el texto de cada registro "
                "(el codificador debe poder leerlo).")
            return
        self.val_data = (data, cat_codes, meta, path)
        short = path.name if len(path.name) <= 32 else path.name[:29] + "…"
        self.val_file_var.set(short)
        self._recalc_validacion()
        self.btn_val_gen.config(state=tk.NORMAL)

    def _recalc_validacion(self):
        if not self.val_data:
            return
        data, _, _, _ = self.val_data
        N = len(data)
        n = lacy_riffe_n(N, P=float(self.val_p_var.get()))
        self.val_n_var.set(f"N = {N}  →  codificar a mano: n = {n}")

    def _generar_validacion(self):
        if not self.val_data:
            return
        data, cat_codes, meta, src = self.val_data
        N   = len(data)
        P   = float(self.val_p_var.get())
        n   = lacy_riffe_n(N, P=P)
        ids = sorted(random.sample(list(data.keys()), n))

        s  = self._excel_styles()
        ts = datetime.now().strftime("%Y%m%d_%H%M%S")
        RESULTS_DIR.mkdir(parents=True, exist_ok=True)
        archivos = []
        try:
            # (sufijo, con_valores): HUMANO vacío para codificación ciega; IA lleno
            for sufijo, con_valores in (("HUMANO", False), ("IA", True)):
                wb = openpyxl.Workbook()
                ws = wb.active
                ws.title = "Corpus"
                cols = ([("id", 18, s["fill_azul"]), ("texto", 62, s["fill_azul"])]
                        + [(c, 14, s["fill_verde"]) for c in cat_codes])
                ws.row_dimensions[1].height = 30
                for ci, (nombre, ancho, fill) in enumerate(cols, 1):
                    cell = ws.cell(row=1, column=ci, value=nombre)
                    cell.font = (s["h_font_dark"] if fill is s["fill_verde"]
                                 else s["h_font"])
                    cell.fill = fill
                    cell.alignment = s["center"]; cell.border = s["borde"]
                    ws.column_dimensions[get_column_letter(ci)].width = ancho
                for ri, rid in enumerate(ids, 2):
                    fila = [rid, str(meta[rid].get("texto") or "")]
                    if con_valores:
                        fila += [data[rid].get(c) for c in cat_codes]  # None → vacía
                    else:
                        fila += [None] * len(cat_codes)
                    ws.row_dimensions[ri].height = 60
                    for ci, valor in enumerate(fila, 1):
                        cell = ws.cell(row=ri, column=ci, value=valor)
                        cell.font = s["d_font"]; cell.border = s["borde"]
                        cell.alignment = s["left"] if ci == 2 else s["center"]
                ws.freeze_panes = "A2"
                # La muestra declara sus propios parámetros: dentro de seis
                # meses nadie recuerda con qué P se sacó.
                self._hoja_metodo(wb, {
                    "Contenido": (
                        "Muestra de fiabilidad para codificación humana. El "
                        "archivo HUMANO va vacío a propósito: se codifica sin "
                        "mirar el archivo IA, y después se comparan los dos en "
                        "la pestaña Confiabilidad para obtener el κ de Cohen."),
                    "Archivo de origen":    str(src),
                    "Universo (N)":         str(N),
                    "Muestra (n)":          str(n),
                    "Acuerdo esperado (P)": f"{P:.2f}",
                    "Margen de error":      "±0,05",
                    "Nivel de confianza":   "95% (dos colas, z = 1,96)",
                    "Fórmula":              ("Lacy & Riffe (1996): "
                                             "n = N·P·Q / ((N−1)·SE² + P·Q)"),
                    "Selección":            "Aleatoria simple sin reemplazo",
                    "Codebook":             self.codebook_var.get(),
                }, metodo=(
                    "Muestra aleatoria para medir la fiabilidad entre "
                    "codificadores. Se codifica a mano con las mismas reglas "
                    "que aplicó el modelo: 1 si el texto cumple los criterios "
                    "del codebook, 0 si no los cumple, N/C si el texto no "
                    "permite decidir. Este archivo no lleva hoja de razones."))
                fname = f"HABIC_VALIDACION_{sufijo}_{ts}.xlsx"
                wb.save(RESULTS_DIR / fname)
                archivos.append(fname)
        except PermissionError:
            messagebox.showerror("Error",
                "No se pudo guardar porque un archivo está abierto en Excel.")
            return

        messagebox.showinfo(
            "Archivos de validación generados",
            f"Muestra aleatoria: {n} de {N} textos (Lacy & Riffe, P="
            f"{self.val_p_var.get()}).\n\n"
            f"• {archivos[0]}  ← codifícalo a mano, SIN mirar el archivo IA\n"
            f"• {archivos[1]}  ← códigos del modelo (no abrir hasta terminar)\n\n"
            "Aplica la misma regla que la IA: 1 solo si el texto cumple los "
            "criterios; 0 si no los cumple; N/C si el texto no permite decidir.\n\n"
            "Cuando termines: carga IA como Resultado A y HUMANO como "
            "Resultado B en la sección 2 y pulsa Calcular Kappa.\n\n"
            f"Ubicación:\nDocuments/HABIC/resultados/",
        )

    def _browse_kappa_file(self, which):
        path = filedialog.askopenfilename(
            title=f"Seleccionar Resultado {which.upper()}",
            filetypes=[("Excel", "*.xlsx"), ("Todos los archivos", "*.*")],
        )
        if not path:
            return
        path = Path(path)
        try:
            data, cat_codes, meta = load_result_file(path)
            n_ids  = len(data)
            n_cats = len(cat_codes)
            short  = path.name if len(path.name) <= 36 else path.name[:33] + "…"
            preview = (
                f"{n_ids} registros  ·  {n_cats} categoría(s): "
                + ", ".join(cat_codes[:5])
                + ("…" if n_cats > 5 else "")
            )
            if which == "a":
                self.kappa_data_a = (data, cat_codes, meta)
                self.kappa_file_a_var.set(short)
                self.kappa_prev_a_var.set(preview)
            else:
                self.kappa_data_b = (data, cat_codes, meta)
                self.kappa_file_b_var.set(short)
                self.kappa_prev_b_var.set(preview)
            if self.kappa_data_a is not None and self.kappa_data_b is not None:
                self.btn_kappa_calc.config(state=tk.NORMAL)
        except Exception as e:
            messagebox.showerror("Error al cargar", str(e))

    def calculate_kappa(self):
        if self.kappa_data_a is None or self.kappa_data_b is None:
            return
        data_a, cats_a, meta_a = self.kappa_data_a
        data_b, cats_b, meta_b = self.kappa_data_b

        shared_ids = sorted(set(data_a.keys()) & set(data_b.keys()))
        if not shared_ids:
            messagebox.showerror(
                "Sin IDs en común",
                "Los dos archivos no tienen ningún ID de fila en común.\n\n"
                "Asegúrate de que ambos sean resultados del mismo corpus.",
            )
            return

        cats_b_set  = set(cats_b)
        shared_cats = [c for c in cats_a if c in cats_b_set]
        if not shared_cats:
            messagebox.showerror(
                "Sin categorías en común",
                "Los dos archivos no comparten ninguna columna de categoría.",
            )
            return

        per_cat      = []
        all_pairs    = []
        disagree_rows = []
        n_faltantes  = 0   # celdas excluidas por dato faltante (error de codificación)

        for cat in shared_cats:
            pairs = []
            for rid in shared_ids:
                a_val = data_a[rid].get(cat)
                b_val = data_b[rid].get(cat)
                # Dato faltante: no es acuerdo ni desacuerdo → fuera del cálculo
                if a_val is None or b_val is None:
                    n_faltantes += 1
                    continue
                pairs.append((a_val, b_val))
                all_pairs.append((a_val, b_val))
                if a_val != b_val:
                    disagree_rows.append({
                        "id":        rid,
                        "categoria": cat,
                        "valor_a":   a_val,
                        "valor_b":   b_val,
                        "texto":     (meta_a.get(rid, {}).get("texto") or
                                      meta_b.get(rid, {}).get("texto") or ""),
                    })
            kappa, po, pe, n = kappa_for_pair(pairs)
            per_cat.append({
                "cat":     cat,
                "kappa":   kappa,
                "po":      po,
                "pe":      pe,
                "n_agree": sum(1 for a, b in pairs if a == b),
                "n_total": n,
            })

        k_global, po_global, pe_global, _ = kappa_for_pair(all_pairs)

        # Tres cifras, no una. El agregado junta todas las decisiones y una
        # categoría frecuente puede tapar a una rara; el mínimo es el que
        # decide si el instrumento se sostiene.
        validos = [c for c in per_cat if c["kappa"] is not None]
        k_prom  = (sum(c["kappa"] for c in validos) / len(validos)
                   if validos else None)
        peor    = min(validos, key=lambda c: c["kappa"]) if validos else None

        self._show_kappa_results_window(
            k_global, po_global, pe_global,
            shared_ids, shared_cats, per_cat, disagree_rows, n_faltantes,
            k_prom, peor)

    def _show_kappa_results_window(self, k_global, po_global, pe_global,
                                    shared_ids, shared_cats, per_cat, disagree_rows,
                                    n_faltantes=0, k_prom=None, peor=None):
        win = tk.Toplevel(self.root)
        win.title("Resultados — Kappa de Cohen")
        win.geometry("720x540")
        win.minsize(600, 420)
        win.resizable(True, True)
        win.transient(self.root)

        win.columnconfigure(0, weight=1)
        win.rowconfigure(2, weight=1)   # fila del Treeview se expande

        # ── Banda superior: las tres cifras ───────────────────────────────
        kgf = tk.Frame(win, bg=C_PANEL, relief=tk.GROOVE, bd=1)
        kgf.grid(row=0, column=0, sticky="ew", padx=12, pady=(12, 6))
        kgf.columnconfigure(2, weight=1)

        cifras = tk.Frame(kgf, bg=C_PANEL)
        cifras.grid(row=0, column=0, columnspan=3, sticky="w", padx=10, pady=(8, 2))
        k_min = peor["kappa"] if peor else None
        for col, (etiqueta, valor, extra) in enumerate((
                ("κ agregado", k_global, ""),
                ("κ promedio", k_prom, ""),
                ("κ mínimo", k_min, f"  ({peor['cat']})" if peor else ""))):
            caja = tk.Frame(cifras, bg=C_PANEL)
            caja.grid(row=0, column=col, sticky="w", padx=(0, 28))
            tk.Label(caja, text=etiqueta, font=(FUENTE, 10, "bold"),
                     bg=C_PANEL, fg=C_ACENTO).pack(anchor="w")
            tk.Label(caja, text=f"{valor:.3f}" if valor is not None else "—",
                     font=(FUENTE, 20, "bold"), bg=C_PANEL,
                     fg="black").pack(anchor="w")
            tk.Label(caja, text=kappa_interpretation(valor) + extra,
                     font=(FUENTE, 10), bg=C_PANEL, fg=C_ACENTO).pack(anchor="w")

        po_pct = f"{po_global * 100:.1f}%" if po_global is not None else "—"
        pe_pct = f"{pe_global * 100:.1f}%" if pe_global is not None else "—"
        falt = (f"  ·  {n_faltantes} celda(s) sin dato (excluidas)"
                if n_faltantes else "")
        tk.Label(kgf,
                 text=(f"Po = {po_pct}  ·  Pe = {pe_pct}  ·  "
                       f"{len(shared_ids)} IDs en común  "
                       f"·  {len(shared_cats)} categoría(s){falt}"),
                 font=(FUENTE, 10), bg=C_PANEL, fg=C_TEXTO_SEC).grid(
            row=1, column=0, columnspan=3, sticky="w", padx=10)
        tk.Label(kgf, justify="left", wraplength=660,
                 text=("El agregado junta todas las decisiones de todas las "
                       "categorías: una categoría frecuente puede tapar a una "
                       "rara. El promedio da el mismo peso a cada categoría, y "
                       "el mínimo es el que decide si el instrumento se sostiene."),
                 font=(FUENTE, 9), bg=C_PANEL, fg=C_TEXTO_SEC).grid(
            row=2, column=0, columnspan=3, sticky="w", padx=10, pady=(2, 8))

        # ── Tabla por categoría ───────────────────────────────────────────
        tk.Label(win, text="Por categoría:", font=(FUENTE, 10, "bold")).grid(
            row=1, column=0, sticky="w", padx=14, pady=(4, 2))

        tree_frame = tk.Frame(win)
        tree_frame.grid(row=2, column=0, sticky="nsew", padx=12, pady=(0, 6))
        tree_frame.rowconfigure(0, weight=1)
        tree_frame.columnconfigure(0, weight=1)

        vsb = ttk.Scrollbar(tree_frame, orient="vertical")
        tree = ttk.Treeview(
            tree_frame,
            columns=("cat", "kappa", "interp", "po", "pe", "n_acuerdo", "n_total"),
            show="headings",
            yscrollcommand=vsb.set,
        )
        vsb.config(command=tree.yview)
        tree.heading("cat",       text="Categoría")
        tree.heading("kappa",     text="κ")
        tree.heading("interp",    text="Interpretación")
        tree.heading("po",        text="Acuerdo (Po)")
        tree.heading("pe",        text="Esperado (Pe)")
        tree.heading("n_acuerdo", text="Coinciden")
        tree.heading("n_total",   text="Total IDs")
        tree.column("cat",       width=110, anchor="w")
        tree.column("kappa",     width=70,  anchor="center")
        tree.column("interp",    width=140, anchor="w")
        tree.column("po",        width=100, anchor="center")
        tree.column("pe",        width=100, anchor="center")
        tree.column("n_acuerdo", width=90,  anchor="center")
        tree.column("n_total",   width=80,  anchor="center")
        tree.tag_configure("bueno", background="#F0FDF4")
        tree.tag_configure("medio", background="#FFFBEB")
        tree.tag_configure("malo",  background="#FEF2F2")

        for r in per_cat:
            k = r["kappa"]
            tag    = ("bueno" if (k is not None and k >= 0.60) else
                      "medio" if (k is not None and k >= 0.40) else "malo")
            po_str = f"{r['po'] * 100:.1f}%" if r["po"] is not None else "—"
            pe_str = f"{r['pe'] * 100:.1f}%" if r["pe"] is not None else "—"
            k_str  = f"{k:.3f}" if k is not None else "—"
            tree.insert("", "end", tag=tag, values=(
                r["cat"], k_str, kappa_interpretation(k),
                po_str, pe_str, r["n_agree"], r["n_total"],
            ))

        tree.grid(row=0, column=0, sticky="nsew")
        vsb.grid(row=0, column=1, sticky="ns")

        # Foto inmediata de los datos de ESTA ventana, para que sus botones
        # sigan siendo correctos aunque el usuario calcule otro Kappa después.
        k_snapshot = {
            "k_global":    k_global,
            "po_global":   po_global,
            "pe_global":   pe_global,
            "shared_ids":  shared_ids,
            "shared_cats": shared_cats,
            "per_cat":     per_cat,
            "n_faltantes": n_faltantes,
            "k_prom":      k_prom,
            "peor":        peor,
        }

        # ── Botones ───────────────────────────────────────────────────────
        btn_frame = tk.Frame(win)
        btn_frame.grid(row=3, column=0, sticky="w", padx=12, pady=(0, 12))

        tk.Button(
            btn_frame,
            text=f"Ver desacuerdos ({len(disagree_rows)})",
            font=(FUENTE, 10), bg=C_ACENTO, fg="white",
            padx=12, pady=5, cursor="hand2", relief=tk.FLAT,
            command=lambda d=disagree_rows: self.show_kappa_disagreements(d),
        ).pack(side=tk.LEFT, padx=(0, 10))

        tk.Button(
            btn_frame, text="Exportar HABIC_KAPPA_*.xlsx",
            font=(FUENTE, 10), bg=C_ACENTO, fg="white",
            padx=12, pady=5, cursor="hand2", relief=tk.FLAT,
            command=lambda r=k_snapshot, d=disagree_rows:
                self.export_kappa_excel(r, d),
        ).pack(side=tk.LEFT)

    def show_kappa_disagreements(self, disagree):
        if not disagree:
            messagebox.showinfo(
                "Sin desacuerdos",
                "¡Los dos archivos tienen acuerdo total en todos los IDs compartidos!",
            )
            return
        win = tk.Toplevel(self.root)
        win.title(f"Desacuerdos — {len(disagree)} caso(s)")
        win.geometry("780x460")
        win.resizable(True, True)
        win.transient(self.root)

        tk.Label(win,
                 text=f"Desacuerdos entre Resultado A y B  "
                      f"({len(disagree)} caso(s))",
                 font=(FUENTE, 12, "bold"), pady=8).pack()
        tk.Label(win, font=(FUENTE, 9), fg=C_TEXTO_SEC,
                 text="Filas donde el valor asignado por A y B difiere para esa categoría.",
                 ).pack(pady=(0, 8))

        cols  = ("id", "categoria", "valor_a", "valor_b", "texto")
        heads = ("ID", "Categoría", "A", "B", "Texto (extracto)")
        anchs = (120, 90, 50, 50, 400)
        tf = tk.Frame(win)
        tf.pack(fill=tk.BOTH, expand=True, padx=12, pady=(0, 8))
        vsb = ttk.Scrollbar(tf, orient="vertical")
        hsb = ttk.Scrollbar(tf, orient="horizontal")
        tree = ttk.Treeview(tf, columns=cols, show="headings",
                            yscrollcommand=vsb.set, xscrollcommand=hsb.set)
        vsb.config(command=tree.yview)
        hsb.config(command=tree.xview)
        for col, head, anch in zip(cols, heads, anchs):
            tree.heading(col, text=head)
            tree.column(col, width=anch, minwidth=40, anchor="w")
        tree.tag_configure("par", background=C_PANEL)
        for i, row in enumerate(disagree):
            tag = "par" if i % 2 == 0 else ""
            texto = str(row.get("texto") or "")[:80].replace("\n", " ")
            tree.insert("", "end", tag=tag, values=(
                row["id"], row["categoria"],
                row["valor_a"], row["valor_b"], texto,
            ))
        tree.grid(row=0, column=0, sticky="nsew")
        vsb.grid(row=0, column=1, sticky="ns")
        hsb.grid(row=1, column=0, sticky="ew")
        tf.rowconfigure(0, weight=1)
        tf.columnconfigure(0, weight=1)

        tk.Button(win, text="Cerrar", command=win.destroy,
                  font=(FUENTE, 10), relief=tk.FLAT,
                  bg=C_ACENTO, fg="white", padx=20, pady=5,
                  cursor="hand2").pack(pady=(0, 10))

    def export_kappa_excel(self, k_res, disagree):
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        filename  = f"HABIC_KAPPA_{timestamp}.xlsx"
        filepath  = RESULTS_DIR / filename
        RESULTS_DIR.mkdir(parents=True, exist_ok=True)

        s    = self._excel_styles()
        wb   = openpyxl.Workbook()

        # Hoja 1 — Resumen
        ws1 = wb.active
        ws1.title = "Resumen"
        ws1.merge_cells("A1:G1")
        cell = ws1["A1"]
        cell.value     = "HABIC — Reporte de Confiabilidad entre Codificadores (κ de Cohen)"
        cell.font      = Font(name="Arial", bold=True, color="FFFFFF", size=12)
        cell.fill      = s["fill_naranja"]
        cell.alignment = Alignment(horizontal="center", vertical="center")
        ws1.row_dimensions[1].height = 26

        resumen_filas = [
            ("κ agregado (todas las decisiones juntas)",
             f"{k_res['k_global']:.4f}" if k_res["k_global"] is not None else "—",
             kappa_interpretation(k_res["k_global"])
             + " — una categoría frecuente puede tapar a una rara"),
            ("κ promedio (cada categoría pesa igual)",
             f"{k_res['k_prom']:.4f}" if k_res.get("k_prom") is not None else "—",
             kappa_interpretation(k_res.get("k_prom"))),
            ("κ mínimo (la categoría más débil)",
             f"{k_res['peor']['kappa']:.4f}" if k_res.get("peor") else "—",
             (kappa_interpretation(k_res["peor"]["kappa"])
              + f" — categoría {k_res['peor']['cat']}") if k_res.get("peor")
             else ""),
            ("Acuerdo observado (Po)",
             f"{k_res['po_global'] * 100:.1f}%"
             if k_res["po_global"] is not None else "—", ""),
            ("Acuerdo esperado por azar (Pe)",
             f"{k_res['pe_global'] * 100:.1f}%"
             if k_res.get("pe_global") is not None else "—", ""),
            ("IDs compartidos",     str(len(k_res["shared_ids"])),  ""),
            ("Categorías comparadas", str(len(k_res["shared_cats"])), ""),
            ("Celdas sin dato (excluidas)", str(k_res.get("n_faltantes", 0)),
             "Codificaciones fallidas: no cuentan como 0"),
        ]
        for ri, (lbl, val, note) in enumerate(resumen_filas, 3):
            ws1.cell(row=ri, column=1, value=lbl).font  = s["bold_font"]
            ws1.cell(row=ri, column=1).fill             = PatternFill("solid", fgColor="e0fbfc")
            ws1.cell(row=ri, column=2, value=val).font  = s["d_font"]
            ws1.cell(row=ri, column=3, value=note).font = s["d_font"]

        cat_headers = [
            ("Categoría",    18, s["fill_naranja"]),
            ("κ",            10, s["fill_naranja"]),
            ("Interpretación", 20, s["fill_naranja"]),
            ("Acuerdo (Po)", 14, s["fill_naranja"]),
            ("Esperado (Pe)", 14, s["fill_naranja"]),
            ("Coinciden",    12, s["fill_naranja"]),
            ("Total IDs",    10, s["fill_naranja"]),
        ]
        ws1.row_dimensions[10].height = 24
        for ci, (h, ancho, fill) in enumerate(cat_headers, 1):
            cell = ws1.cell(row=10, column=ci, value=h)
            cell.font = (s["h_font_dark"]
                         if fill in (s["fill_verde"], s["fill_morado"])
                         else s["h_font"])
            cell.fill = fill
            cell.alignment = s["center"]; cell.border = s["borde"]
            ws1.column_dimensions[get_column_letter(ci)].width = ancho

        for ri, r in enumerate(k_res["per_cat"], 11):
            k      = r["kappa"]
            po_str = f"{r['po'] * 100:.1f}%" if r["po"] is not None else "—"
            pe_str = f"{r['pe'] * 100:.1f}%" if r.get("pe") is not None else "—"
            k_str  = f"{k:.4f}" if k is not None else "—"
            fila   = [r["cat"], k_str, kappa_interpretation(k),
                      po_str, pe_str, r["n_agree"], r["n_total"]]
            for ci, valor in enumerate(fila, 1):
                cell = ws1.cell(row=ri, column=ci, value=valor)
                cell.font = s["d_font"]; cell.border = s["borde"]
                cell.alignment = s["left"] if ci == 1 else s["center"]
        ws1.freeze_panes = "A11"

        # Hoja 2 — Desacuerdos
        ws2 = wb.create_sheet(title="Desacuerdos")
        dis_h = [
            ("ID",              20, s["fill_azul"]),
            ("Categoría",       18, s["fill_naranja"]),
            ("Valor A",         10, s["fill_verde"]),
            ("Valor B",         10, s["fill_verde"]),
            ("Texto (extracto)", 70, s["fill_azul"]),
        ]
        ws2.row_dimensions[1].height = 28
        for ci, (h, ancho, fill) in enumerate(dis_h, 1):
            cell = ws2.cell(row=1, column=ci, value=h)
            cell.font = (s["h_font_dark"]
                         if fill in (s["fill_verde"], s["fill_morado"])
                         else s["h_font"])
            cell.fill = fill
            cell.alignment = s["center"]; cell.border = s["borde"]
            ws2.column_dimensions[get_column_letter(ci)].width = ancho

        for ri, row in enumerate(disagree, 2):
            texto = str(row.get("texto") or "")[:200].replace("\n", " ")
            fila  = [row["id"], row["categoria"],
                     row["valor_a"], row["valor_b"], texto]
            for ci, valor in enumerate(fila, 1):
                cell = ws2.cell(row=ri, column=ci, value=valor)
                cell.font = s["d_font"]; cell.border = s["borde"]
                cell.alignment = s["left"] if ci in (1, 2, 5) else s["center"]
        ws2.freeze_panes = "A2"

        k_str = (f"{k_res['k_global']:.4f}"
                 if k_res["k_global"] is not None else "—")
        self._hoja_metodo(wb, {
            "Contenido": "Reporte de confiabilidad entre codificadores "
                         "(κ de Cohen) — comparación de dos codificaciones "
                         "del mismo corpus.",
            "κ global":  f"{k_str}  ({kappa_interpretation(k_res['k_global'])})",
            "IDs comparados": str(len(k_res["shared_ids"])),
            "Celdas sin dato (excluidas)": str(k_res.get("n_faltantes", 0)),
        })

        try:
            wb.save(filepath)
        except PermissionError:
            messagebox.showerror("Error",
                f"No se pudo guardar '{filename}' porque está abierto en Excel.")
            return

        messagebox.showinfo(
            "Exportado",
            f"Reporte Kappa guardado:\n{filename}\n\n"
            f"Hoja 1 — Resumen: κ global y por categoría\n"
            f"Hoja 2 — Desacuerdos: {len(disagree)} caso(s)\n\n"
            f"Ubicación:\nDocuments/HABIC/resultados/",
        )

    # ── HELPERS UI ────────────────────────────────────────────────────────── #

    def update_tw_status(self, msg):
        self._set(self.tw_status_var, msg)

    def _set(self, var, msg):
        """Escribe en una StringVar desde cualquier hilo."""
        self.root.after(0, lambda: var.set(msg))

    def update_tw_count(self, msg):
        self._set(self.tw_count_var, str(msg))

    def update_xl_status(self, msg):
        self._set(self.xl_status_var, msg)

    def update_xl_count(self, msg):
        self._set(self.xl_count_var, str(msg))


if __name__ == "__main__":
    root = tk.Tk()
    app  = HABICApp(root)
    root.mainloop()
