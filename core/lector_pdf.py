"""
lector_pdf.py
=============
Lee facturas PDF (en realidad ZIPs con .txt y .jpeg por página)
y extrae todos los campos relevantes usando proveedores.json.

FLUJO:
    1. Detecta si el archivo es un ZIP con páginas (formato actual) o un PDF clásico
    2. Junta el texto de todas las páginas
    3. Identifica el proveedor por sus firmas (texto único)
    4. Extrae los campos con los regex del proveedor
    5. Valida que el CUPS del PDF coincide con el CUPS registrado para esa comunidad
    6. Devuelve un dict limpio con todos los datos, listo para insertar_factura()

TIPOS DE RESULTADO:
    {'ok': True, 'tipo': 'FACTURA', 'datos': {...}}
    {'ok': True, 'tipo': 'LECTURA_METRIGEST', 'datos': [...]}   # lista de vecinos
    {'ok': False, 'motivo': 'PROVEEDOR_NO_IDENTIFICADO'}
    {'ok': False, 'motivo': 'CUPS_NO_COINCIDE', 'cups_pdf': ..., 'cups_esperado': ...}
    {'ok': False, 'motivo': 'CUARENTENA', 'detalle': ...}
"""

import os
import re
import json
import zipfile
import hashlib
import calendar
import shutil
import sys
from dataclasses import dataclass
from functools import lru_cache
from datetime import datetime
from pathlib import Path


# ---------------------------------------------------------------------------
# CARGA DE CONFIGURACIÓN
# ---------------------------------------------------------------------------

@lru_cache(maxsize=None)
def cargar_proveedores(ruta_json: str | None = None) -> dict:
    """Carga proveedores.json desde la configuración del proyecto.

    Se conserva la antigua ubicación junto al lector como compatibilidad para
    instalaciones ya configuradas, pero las instalaciones actuales guardan el
    archivo en ``config/``.
    """
    if ruta_json is None:
        legacy_path = os.path.join(os.path.dirname(__file__), "proveedores.json")
        project_config = os.path.join(
            os.path.dirname(os.path.dirname(__file__)), "config", "proveedores.json"
        )
        ruta_json = legacy_path if os.path.exists(legacy_path) else project_config
    with open(ruta_json, encoding="utf-8") as f:
        return json.load(f)


# ---------------------------------------------------------------------------
# UTILIDADES
# ---------------------------------------------------------------------------

def _limpiar_numero(texto: str, puntos_millares: bool = False) -> float:
    """Convierte un importe o consumo textual a ``float``.

    ``puntos_millares`` se activa sólo desde el perfil que confirma ese
    formato. Así ``4.011`` puede leerse como 4011 kWh en una factura española
    sin convertir los decimales con punto de otros proveedores.
    """
    if not texto:
        return 0.0
    t = texto.strip().replace("€", "").replace(" ", "")
    if puntos_millares and "," not in t and re.fullmatch(r"\d{1,3}(?:\.\d{3})+", t):
        t = t.replace(".", "")
    elif "." in t and "," in t:
        if t.index(".") < t.index(","):
            t = t.replace(".", "")
    t = t.replace(",", ".")
    try:
        return float(t)
    except ValueError:
        return 0.0


def _limpiar_importe_decorativo(texto: str, patron: str) -> float:
    """
    Extrae importe de texto con puntos suspensivos intercalados entre los dígitos.
    Ej: 'TOTAL FACTURA .....1....4..0..3..,.7..9...€' → 1403.79

    Si el patrón define un grupo con nombre "resto", solo se leen los dígitos
    de ese grupo (útil cuando delante del importe hay otro número, como el
    "10" de "I.V.A. 10% ....127,62€", que no debe contarse como decorativo).
    """
    m = re.search(patron, texto, re.IGNORECASE)
    if not m:
        return 0.0
    try:
        raw = m.group("resto")
    except IndexError:
        raw = m.group(0)
    # Extraer solo dígitos y la última coma (separador decimal español)
    solo_digitos = re.findall(r'\d', raw)
    if not solo_digitos:
        return 0.0
    # Los últimos 2 dígitos son siempre los decimales
    if len(solo_digitos) >= 3:
        enteros  = ''.join(solo_digitos[:-2])
        decimales = ''.join(solo_digitos[-2:])
        try:
            return float(f"{enteros}.{decimales}")
        except ValueError:
            return 0.0
    return 0.0


def _normalizar_fecha(texto: str) -> str:
    """Convierte fechas en distintos formatos a YYYY-MM-DD."""
    texto = texto.strip()
    meses = {
        "enero": "01", "febrero": "02", "marzo": "03", "abril": "04",
        "mayo": "05", "junio": "06", "julio": "07", "agosto": "08",
        "septiembre": "09", "octubre": "10", "noviembre": "11", "diciembre": "12"
    }
    # DD.MM.YYYY, DD/MM/YYYY o DD-MM-YYYY, con año de 2 o 4 dígitos
    # (el año de 4 dígitos se prueba primero: si no, "2026" solo aportaría "20")
    m = re.match(r"^(\d{2})[./-](\d{2})[./-](\d{4}|\d{2})$", texto)
    if m:
        anio = m.group(3)
        if len(anio) == 2:
            anio = f"20{anio}"
        return f"{anio}-{m.group(2)}-{m.group(1)}"
    # DD de mes de YYYY (con espacios)
    m = re.match(r"(\d{1,2})\s+de\s+(\w+)\s+de\s+(\d{4})", texto, re.IGNORECASE)
    if m:
        mes = meses.get(m.group(2).lower(), "00")
        return f"{m.group(3)}-{mes}-{int(m.group(1)):02d}"
    # DDdeMESDEYYYY (texto fusionado sin espacios: "07demayode2026", "31demarzode2026")
    m = re.match(r"(\d{1,2})(de)(\w+?)(de)(\d{4})", texto, re.IGNORECASE)
    if m:
        mes = meses.get(m.group(3).lower(), "00")
        return f"{m.group(5)}-{mes}-{int(m.group(1)):02d}"
    return texto


def _md5_archivo(ruta: str) -> str:
    h = hashlib.md5()
    with open(ruta, 'rb') as f:
        for bloque in iter(lambda: f.read(65536), b''):
            h.update(bloque)
    return h.hexdigest()


def extraer_cif_pdf(texto: str) -> str | None:
    """
    Extrae el CIF/NIF de la comunidad del texto del PDF.
    El CIF aparece en todas las facturas, generalmente como:
      'NIF: H99258139', 'CIF / NIF: H99258139', 'H99258139'
    Formato CIF de comunidad: letra + 8 dígitos (ej: H99258139).
    """
    patrones = [
        r'NIF\s*/?\s*CIF[:\s]+([A-Z][0-9]{8})',
        r'CIF\s*/?\s*NIF[:\s]+([A-Z][0-9]{8})',
        r'NIF[:\s]+([A-Z][0-9]{8})',
        r'CIF[:\s]+([A-Z][0-9]{8})',
        r'\b([HEJG][0-9]{8})\b',  # H=cooperativa/comunidad, E=entidad, J/G=asociación
    ]
    for patron in patrones:
        m = re.search(patron, texto, re.IGNORECASE)
        if m:
            return m.group(1).upper()
    return None
    h = hashlib.md5()
    with open(ruta, "rb") as f:
        for bloque in iter(lambda: f.read(65536), b""):
            h.update(bloque)
    return h.hexdigest()


# ---------------------------------------------------------------------------
# EXTRACCIÓN LOCAL EN DOS FASES — resiliencia sin APIs ni dependencias externas
# ---------------------------------------------------------------------------

# ── Fase 1: anclas regex con keywords de alta especificidad ─────────────────
# Soportan separadores variados entre keyword y número (espacios, puntos, etc.)
_ANCLAS_REGEX_CAPA_B: dict[str, list[str]] = {
    "importe_total": [
        r"(?i)total\s+a\s+pagar[\s:\.•·]+(?P<valor>[\d]{1,7}(?:[.,][\d]{3})*[.,][\d]{2})",
        r"(?i)importe\s+total[\s:\.•·]+(?P<valor>[\d]{1,7}(?:[.,][\d]{3})*[.,][\d]{2})",
        r"(?i)total\s+factura[\s:\.•·]+(?P<valor>[\d]{1,7}(?:[.,][\d]{3})*[.,][\d]{2})",
        r"(?i)total\s+importe(?:\s+factura)?[\s:\.•·]+(?P<valor>[\d]{1,7}(?:[.,][\d]{3})*[.,][\d]{2})",
        r"(?i)(?:importe\s+a\s+)?(?:pagar|abonar)[\s:\.•·]+(?P<valor>[\d]{1,7}(?:[.,][\d]{3})*[.,][\d]{2})",
    ],
    "consumo_kwh": [
        r"(?i)consumo\s+(?:total|activ[ao]|facturado)[\s:]*(?P<valor>[\d]{1,7}[.,]?[\d]{0,3})\s*kwh",
        r"(?i)energ[íi]a\s+(?:activa|consumida|total)[\s:]*(?P<valor>[\d]{1,7}[.,]?[\d]{0,3})\s*kwh",
        r"(?i)total\s+consumido[\s:]*(?P<valor>[\d]{1,7}[.,]?[\d]{0,3})\s*kwh",
    ],
    "consumo_m3": [
        r"(?i)consumo\s+(?:real|total|gas|agua)[\s:]*(?P<valor>[\d]{1,7}[.,]?[\d]{0,3})\s*m[3³]",
        r"(?i)total\s+consumido[\s:]*(?P<valor>[\d]{1,7}[.,]?[\d]{0,3})\s*m[3³]",
        r"(?i)gas\s+natural[\s:]*(?P<valor>[\d]{1,7}[.,]?[\d]{0,3})\s*m[3³]",
    ],
    "iva": [
        r"(?i)iva\s*\(\s*\d{1,2}\s*%\s*\)[\s:\.•·]*(?:de\s+[\d.,]+[\s:\.•·]*)?(?P<valor>[\d]{1,7}(?:[.,][\d]{3})*[.,][\d]{2})",
        r"(?i)cuota\s+iva[\s:\.•·]+(?P<valor>[\d]{1,7}(?:[.,][\d]{3})*[.,][\d]{2})",
        r"(?i)total\s+iva[\s:\.•·]+(?:de\s+[\d.,]+[\s:\.•·]*)?(?P<valor>[\d]{1,7}(?:[.,][\d]{3})*[.,][\d]{2})",
    ],
    "termino_fijo": [
        r"(?i)t[ée]rmino\s+fijo[\s:\.•·\"»)]*\)?[\s:\.•·]*(?P<valor>[\d]{1,7}(?:[.,][\d]{3})*[.,][\d]{2})",
        r"(?i)cuota\s+fija[\s:\.•·]+(?P<valor>[\d]{1,7}(?:[.,][\d]{3})*[.,][\d]{2})",
        r"(?i)potencia[\s:\.•·]+(?P<valor>[\d]{1,7}(?:[.,][\d]{3})*[.,][\d]{2})",
        r"(?i)disponibilidad[\s:\.•·]+(?P<valor>[\d]{1,7}(?:[.,][\d]{3})*[.,][\d]{2})",
    ],
    "termino_variable": [
        r"(?i)t[ée]rmino\s+variable[\s:\.•·\"»)]*\)?[\s:\.•·]*(?P<valor>[\d]{1,7}(?:[.,][\d]{3})*[.,][\d]{2})",
        r"(?i)cuota\s+variable[\s:\.•·]+(?P<valor>[\d]{1,7}(?:[.,][\d]{3})*[.,][\d]{2})",
        r"(?i)energ[íi]a[\s:\.•·]+(?P<valor>[\d]{1,7}(?:[.,][\d]{3})*[.,][\d]{2})",
    ],
}

# ── Fase 2: tesauro legal ────────────────────────────────────────────────────
# Solo keywords multi-palabra para evitar falsos positivos: "total" solo
# también aparece en "consumo total", "total kWh", etc.
_TESAURO_LEGAL: dict[str, list[str]] = {
    "importe_total": [
        "total a pagar",
        "importe total",
        "total factura",
        "total importe",
        "a pagar",
        "a abonar",
        "importe a pagar",
        "importe a abonar",
        "importe neto factura",
        "importe neto",
        "liquido a pagar",
        "total del periodo",
        "total periodo",
    ],
    "consumo_kwh": [
        "consumo total",
        "consumo activo",
        "consumo facturado",
        "energia activa",
        "energia consumida",
        "total kwh",
        "total consumido",
    ],
    "consumo_m3": [
        "consumo real",
        "consumo total",
        "consumo gas",
        "consumo agua",
        "total m3",
        "total consumido",
        "metros cubicos",
    ],
    "iva": [
        "cuota iva",
        "total iva",
        "iva normal",
        "iva repercutido",
    ],
    "termino_fijo": [
        "termino fijo",
        "cuota fija",
        "cuota de servicio",
        "disponibilidad",
        "potencia contratada",
    ],
    "termino_variable": [
        "termino variable",
        "cuota variable",
        "consumo facturado",
        "energia consumida",
    ],
}

# Patrón para extraer candidatos numéricos de una línea ya limpiada
_PAT_NUM_LINEA = re.compile(
    r"(?<!\d)"
    r"("
    r"\d{1,3}(?:\.\d{3})+,\d{2}"   # ES con miles: 1.234,56
    r"|"
    r"\d{1,7},\d{2,3}"              # ES sin miles: 145,20
    r"|"
    r"\d{1,3}(?:,\d{3})+\.\d{2}"   # US con miles: 1,234.56
    r"|"
    r"\d{1,7}\.\d{2,3}"             # US sin miles: 145.20
    r"|"
    r"\d{2,7}"                       # entero (min 2 digitos)
    r")"
    r"(?!\d)"
)


def _num_linea_a_float(s: str) -> float:
    """
    Convierte string numérico de línea a float para comparación heurística.
    Resuelve la ambigüedad del punto en '1.234' (miles ES = 1234 vs decimal = 1.234):
    si hay exactamente 3 dígitos tras el punto y no hay coma, se trata como miles.
    """
    t = s.strip()
    if "." in t and "," in t:
        if t.index(".") < t.index(","):
            t = t.replace(".", "")   # miles ES: "1.234,56" -> "1234,56"
        t = t.replace(",", ".")
    elif "," in t:
        t = t.replace(",", ".")      # decimal ES: "145,20" -> "145.20"
    elif "." in t:
        partes = t.split(".")
        if len(partes) == 2 and len(partes[1]) == 3:
            t = t.replace(".", "")   # miles sin decimal: "1.234" -> "1234"
    try:
        return float(t)
    except ValueError:
        return 0.0


def _extraer_numero_linea(linea: str) -> str | None:
    """
    Extrae el número más probable de una línea candidata (ya identificada por tesauro).

    Estrategia:
      1. Elimina secuencias de 3+ caracteres decorativos consecutivos iguales
         ('........', '--------', '=========') reemplazándolos por espacio.
         Resuelve formatos como:
           'Importe neto factura del periodo......... 145,20 Euros'
      2. Extrae todos los candidatos numéricos con _PAT_NUM_LINEA.
      3. Prioriza candidatos con exactamente 2 decimales (formato EUR estándar).
      4. Devuelve el de mayor valor (descarta precios unitarios, fechas
         fragmentadas como '01.01' = 1.01, porcentajes menores que el total).
    """
    limpia = re.sub(r'[.\-=_]{3,}', ' ', linea)
    numeros = _PAT_NUM_LINEA.findall(limpia)
    if not numeros:
        return None
    con_2dec = [n for n in numeros if re.search(r'[,.](\d{2})$', n)]
    candidatos = con_2dec if con_2dec else numeros
    return max(candidatos, key=_num_linea_a_float)


def _extraer_campo_capa_b(texto: str, campo: str) -> str | None:
    """
    Capa B local en dos fases — sin APIs ni dependencias externas.

    Fase 1 (anclas regex): patrones directos con keywords de alta especificidad.
      Ejemplo: 'TOTAL A PAGAR.....1.234,56 euros'
        -> el separador absorbe los puntos -> extrae '1.234,56'.

    Fase 2 (tesauro legal + segmentacion por lineas):
      Escanea el texto linea a linea. Cuando una linea contiene algun termino
      del tesauro del campo (keywords multi-palabra para evitar falsos positivos),
      limpia sus separadores decorativos y extrae el numero mayor con 2 decimales.
      Ejemplo: 'Importe neto factura del periodo......... 145,20 Euros'
        -> keyword 'importe neto factura' -> limpia '.........' -> extrae '145,20'.
    """
    # -- Fase 1: anclas regex --------------------------------------------------
    for patron in _ANCLAS_REGEX_CAPA_B.get(campo, []):
        m = re.search(patron, texto, re.IGNORECASE | re.MULTILINE)
        if m:
            try:
                return m.group("valor").strip()
            except IndexError:
                continue

    # -- Fase 2: tesauro + segmentacion por lineas ----------------------------
    keywords = _TESAURO_LEGAL.get(campo, [])
    # Intentar primero las keywords mas largas (mas especificas)
    keywords_ord = sorted(keywords, key=len, reverse=True)

    for linea in texto.splitlines():
        linea_lower = linea.lower()
        for kw in keywords_ord:
            if kw in linea_lower:
                resultado = _extraer_numero_linea(linea)
                if resultado:
                    return resultado
                break  # keyword encontrada pero sin numero valido -> siguiente linea

    return None


# ---------------------------------------------------------------------------
# EXTRACCIÓN DE TEXTO
# ---------------------------------------------------------------------------

def _extraer_texto_pdf_con_columnas(ruta) -> str:
    """
    Extrae texto de un PDF reconstruyendo las líneas desde palabras individuales.
    Necesario para PDFs de doble columna (ej. facturas Endesa, Energía XXI) donde
    pdfplumber.extract_text() fusiona columnas adyacentes sin espacios.

    Estrategia:
      1. Extrae palabras con sus coordenadas x/y.
      2. Agrupa por y (misma línea) y ordena por x dentro de cada línea.
      3. Reconstruye el texto con espacios reales entre palabras.
    Si falla, cae al extract_text() clásico como respaldo.
    """
    import pdfplumber
    texto_total = []
    with pdfplumber.open(ruta) as pdf:
        for pag in pdf.pages:
            try:
                words = pag.extract_words(
                    x_tolerance=3, y_tolerance=3, keep_blank_chars=False
                )
                if not words:
                    # Página vacía o imagen — intentar extract_text
                    t = pag.extract_text()
                    if t:
                        texto_total.append(t)
                    continue

                # Agrupar palabras por línea (misma coordenada y)
                lineas: dict[float, list] = {}
                for w in words:
                    y = round(w["top"], 0)
                    lineas.setdefault(y, []).append((w["x0"], w["text"]))

                lineas_texto = []
                for y in sorted(lineas.keys()):
                    palabras_ord = sorted(lineas[y], key=lambda p: p[0])
                    lineas_texto.append("  ".join(p[1] for p in palabras_ord))
                texto_total.append("\n".join(lineas_texto))

            except Exception:
                # Fallback página a página
                t = pag.extract_text()
                if t:
                    texto_total.append(t)

    return "\n".join(texto_total)


def _colapsar_espacios(texto: str) -> str:
    """
    Colapsa espacios/tabs repetidos a uno solo, preservando saltos de línea.

    _extraer_texto_pdf_con_columnas reconstruye cada línea uniendo palabras
    con DOBLE espacio ("  ".join(...)) para separar visualmente columnas.
    Los regex de proveedores.json que usan "\\s+" toleran esto sin problema,
    pero los que tienen un espacio literal simple entre palabras clave
    (ej. "Fecha Factura[:\\s]+...") nunca matchean contra "Fecha  Factura:",
    y el campo queda silenciosamente en None. Se normaliza aquí, una vez,
    para que todo el texto que ve el resto del módulo tenga espaciado simple.
    """
    return re.sub(r"[ \t]{2,}", " ", texto)


def _normalizar_decimales_ocr(texto: str) -> str:
    """Recupera céntimos cuando OCR sustituye la coma por un espacio.

    Solo se toca una cifra seguida del símbolo de euro, por ejemplo
    ``6.279 49€``. No se alteran espacios normales ni escalas de gráficas.
    """
    return re.sub(
        r"(\d{1,3}(?:\.\d{3})*)\s+(\d{2}\s*(?:€|�))",
        r"\1,\2",
        texto,
    )


def _pdf_probablemente_escaneado(ruta: Path) -> bool:
    """Detecta rápidamente PDFs de imagen sin capa tipográfica utilizable.

    Pdfplumber puede tardar mucho intentando descomponer páginas escaneadas
    complejas. Un PDF que contiene imágenes pero no fuentes ni mapa Unicode se
    envía directamente a OCR; la comprobación es binaria y no interpreta el
    documento completo.
    """
    if ruta.suffix.lower() != ".pdf" or zipfile.is_zipfile(ruta):
        return False
    try:
        with ruta.open("rb") as archivo:
            contenido = archivo.read(2 * 1024 * 1024)
    except OSError:
        return False
    return b"/Image" in contenido and b"/Font" not in contenido and b"/ToUnicode" not in contenido


def extraer_texto(ruta_archivo: str) -> str:
    """
    Extrae todo el texto de un archivo.
    Soporta:
      - ZIP con páginas N.txt (formato Meditrade actual)
      - PDF clásico con columnas: usa extracción por palabras para preservar espacios
      - PDF clásico simple: fallback a extract_text()
    """
    ruta = Path(ruta_archivo)

    # — Formato ZIP con páginas .txt —
    if zipfile.is_zipfile(ruta):
        texto_total = []
        with zipfile.ZipFile(ruta) as z:
            nombres = sorted(
                [n for n in z.namelist() if n.endswith(".txt")],
                key=lambda x: int(x.replace(".txt", "")) if x.replace(".txt", "").isdigit() else 99
            )
            for nombre in nombres:
                with z.open(nombre) as f:
                    texto_total.append(f.read().decode("utf-8", errors="replace"))
        return _colapsar_espacios("\n".join(texto_total))

    # — PDF: primero intentar extracción por palabras (mejor para columnas) —
    try:
        import pdfplumber
        texto = _extraer_texto_pdf_con_columnas(ruta)
        if texto.strip():
            return _colapsar_espacios(texto)
    except Exception:
        pass

    # — Fallback: extract_text() clásico —
    try:
        import pdfplumber
        texto_total = []
        with pdfplumber.open(ruta) as pdf:
            for pag in pdf.pages:
                t = pag.extract_text()
                if t:
                    texto_total.append(t)
        return _colapsar_espacios("\n".join(texto_total))
    except Exception:
        pass

    return ""


# ---------------------------------------------------------------------------
# IDENTIFICACIÓN DE PROVEEDOR
# ---------------------------------------------------------------------------

@dataclass(frozen=True)
class OCRResult:
    """Resultado legible por la interfaz del intento de OCR de una fuente."""

    text: str
    status: str
    detail: str
    language: str | None = None


def _tesseract_executable() -> str | None:
    """Localiza un motor existente sin exigir instalaciones al usuario."""
    candidates = [
        shutil.which("tesseract"),
        r"C:\Program Files\Tesseract-OCR\tesseract.exe",
        r"C:\Program Files (x86)\Tesseract-OCR\tesseract.exe",
    ]
    return next((candidate for candidate in candidates if candidate and os.path.exists(candidate)), None)


def _poppler_path() -> str | None:
    """Devuelve la carpeta de Poppler si Winget la instaló fuera del PATH."""
    if sys.platform != "win32":
        return None
    import glob

    candidates = glob.glob(os.path.expandvars(
        r"%LOCALAPPDATA%\Microsoft\WinGet\Packages\oschwartz10612.Poppler_*\poppler-*\Library\bin"
    ))
    return candidates[0] if candidates else None


def extraer_texto_ocr_con_diagnostico(ruta_archivo: str) -> OCRResult:
    """Intenta OCR de portada y conserva una causa concreta cuando no puede.

    La aplicación no pide al gestor que instale nada: el diagnóstico se usa
    para decidir si revisar el documento, corregir una imagen ilegible o
    comunicar una incidencia técnica al soporte del despacho.
    """
    try:
        from pdf2image import convert_from_path
        import pytesseract
    except ImportError:
        return OCRResult(
            text="", status="dependencies_unavailable",
            detail="El componente de lectura de documentos escaneados no está disponible en esta instalación.",
        )

    engine = _tesseract_executable()
    if not engine:
        return OCRResult(
            text="", status="engine_unavailable",
            detail="No se encontró el motor OCR local para leer esta imagen escaneada.",
        )
    pytesseract.pytesseract.tesseract_cmd = engine

    try:
        languages = set(pytesseract.get_languages(config=""))
    except Exception as exc:
        return OCRResult(
            text="", status="engine_unavailable",
            detail=f"No se pudo iniciar el motor OCR local: {exc}",
        )
    language = "spa" if "spa" in languages else "eng" if "eng" in languages else None
    if language is None:
        return OCRResult(
            text="", status="language_unavailable",
            detail="El motor OCR no dispone de un idioma de lectura compatible para este documento.",
        )

    try:
        # Los datos decisivos de una factura están en portada. Limitar el OCR
        # evita que un anexo de muchas páginas bloquee la interfaz durante
        # minutos; si la portada no basta, el documento queda para revisión.
        imagenes = convert_from_path(
            ruta_archivo, dpi=200, poppler_path=_poppler_path(), first_page=1, last_page=1,
        )
        partes = []
        for img in imagenes:
            # Los recibos municipales y de suministros suelen tener varias
            # cajas de importes. PSM 6 conserva etiqueta y cifra en la misma
            # línea, a diferencia del modo automático que separa columnas.
            ocr_config = "--psm 6"
            texto = pytesseract.image_to_string(img, lang=language, config=ocr_config)
            partes.append(texto)
    except Exception as exc:
        return OCRResult(
            text="", status="render_or_ocr_failed",
            detail=f"No se pudo leer la portada escaneada: {exc}", language=language,
        )

    text = _normalizar_decimales_ocr("\n".join(partes))
    if not text.strip():
        return OCRResult(
            text="", status="no_text",
            detail="El OCR no obtuvo texto legible de la primera página.", language=language,
        )
    return OCRResult(text=text, status="ok", detail="OCR aplicado a la portada.", language=language)


def extraer_texto_ocr(ruta_archivo: str) -> str:
    """Compatibilidad: devuelve sólo el texto del OCR de portada."""
    return extraer_texto_ocr_con_diagnostico(ruta_archivo).text


def identificar_proveedor(texto: str, nombre_archivo: str, proveedores: dict) -> tuple[str, dict] | tuple[None, None]:
    """
    Devuelve (clave_proveedor, config_proveedor) o (None, None).

    Lógica especial por familia:
    - METRIGEST_ACS vs METRIGEST_CALEF: por tipo de lectura en el texto
    - ENDESA_LUZ_COMUNIDAD: acepta CUPS con sufijos RR0F o PS0F
    - ENERGIAXXI_BOMBA: solo CUPS con sufijo DY0F
    - Resto: primera firma encontrada, respetando firma_exclusion si existe
    """
    texto_up = texto.upper()

    for clave, config in proveedores["proveedores"].items():
        # Firma de exclusión: si el texto contiene este patrón, saltar este proveedor
        exclusion = config.get("firma_exclusion")
        if exclusion and re.search(exclusion, texto, re.IGNORECASE):
            continue

        firmas = config.get("firmas_identificacion", [])
        for firma in firmas:
            if not re.search(firma, texto, re.IGNORECASE):
                continue
            firmas_requeridas = config.get("firmas_requeridas", [])
            if any(not re.search(requerida, texto, re.IGNORECASE)
                   for requerida in firmas_requeridas):
                continue

            # — Metrigest: ACS vs Calefacción —
            if "METRIGEST" in clave:
                primera_pagina = texto_up[:700]
                if "CALEF" in clave and "CALEFACCION" in primera_pagina:
                    return clave, config
                if "ACS" in clave and "ACS" in primera_pagina and "CALEFACCION" not in primera_pagina:
                    return clave, config
                continue

            # — Endesa luz (mercado libre): CUPS sufijo RR0F o PS0F —
            if "ENDESA_LUZ" in clave and config.get("cups_sufijos_validos"):
                m = re.search(r"CUPS[):\s]+(?P<cups>ES\w+)", texto, re.IGNORECASE)
                if m:
                    cups = m.group("cups").upper()
                    sufijos = config.get("cups_sufijos_validos", ["RR0F", "PS0F"])
                    if any(s in cups for s in sufijos):
                        return clave, config
                continue

            # — Energía XXI (bomba incendios): CUPS sufijo DY0F —
            if "ENERGIAXXI" in clave:
                m = re.search(r"CUPS[):\s]+(?P<cups>ES\w+)", texto, re.IGNORECASE)
                if m and "DY0F" in m.group("cups").upper():
                    return clave, config
                continue

            # — Caso general —
            return clave, config

    return None, None


# ---------------------------------------------------------------------------
# EXTRACCIÓN DE CAMPOS — FACTURAS
# ---------------------------------------------------------------------------

def _extraer_campo(texto: str, patron: str) -> str | None:
    """Aplica un regex con grupo named 'valor' y devuelve el match o None."""
    try:
        m = re.search(patron, texto, re.IGNORECASE | re.MULTILINE)
        if m:
            return m.group("valor").strip()
    except Exception:
        pass
    return None


def _extraer_campo_preferido(texto: str, patron: str, alternativos: list[str]) -> str | None:
    """Extrae primero los patrones más fiables configurados por proveedor.

    Algunos PDF incluyen gráficas entre una etiqueta y el importe. Un patrón
    específico que exige un valor monetario completo evita interpretar la
    escala de esa gráfica como el total de la factura, sin alterar el rescate
    genérico que usan los formatos simples.
    """
    for alternativo in alternativos:
        valor = _extraer_campo(texto, alternativo)
        if valor:
            return valor
    return _extraer_campo(texto, patron)


_PAT_RANGO_FECHAS_CAPA_B = [
    # "del 1 al 30 de Abril de 2026" (Lebal y similares)
    re.compile(
        r"(?i)del\s+(?P<d1>\d{1,2})\s+al\s+(?P<d2>\d{1,2})\s+de\s+(?P<mes>\w+)\s+de\s+(?P<anio>\d{4})"
    ),
    # "Periodo facturado ... 25/07/2025 - 27/08/2025" (rango explicito con contexto)
    re.compile(
        r"(?i)(?:periodo|facturad[oa]|consumo|desde)[^\n]{0,40}?"
        r"(?P<f1>\d{2}[./]\d{2}[./]\d{2,4})\s*(?:-|a|al)\s*(?P<f2>\d{2}[./]\d{2}[./]\d{2,4})"
    ),
]


def _extraer_rango_fechas_capa_b(texto: str) -> tuple[str, str] | None:
    """
    Rescate genérico de fecha_inicio/fecha_fin cuando el proveedor no tiene
    regex propio o el regex propio no matcheó. Cubre el formato en prosa
    'del D al D2 de MES de AAAA' y rangos explícitos 'DD/MM/AAAA - DD/MM/AAAA'
    cerca de una palabra clave de periodo (evita falsos positivos con
    números de factura o cuentas bancarias que también tienen guiones).
    """
    m = _PAT_RANGO_FECHAS_CAPA_B[0].search(texto)
    if m:
        mes, anio = m.group("mes"), m.group("anio")
        ini = _normalizar_fecha(f"{m.group('d1')} de {mes} de {anio}")
        fin = _normalizar_fecha(f"{m.group('d2')} de {mes} de {anio}")
        return ini, fin

    m = _PAT_RANGO_FECHAS_CAPA_B[1].search(texto)
    if m:
        return _normalizar_fecha(m.group("f1")), _normalizar_fecha(m.group("f2"))

    return None


def _extraer_rango_mes_facturado(texto: str, patron: str) -> tuple[str, str] | None:
    """Convierte un mes facturado explícito en el rango completo del mes.

    El patrón pertenece al perfil del proveedor y debe aportar los grupos
    ``anio`` y ``mes``. Así sólo se infiere un período cuando el documento
    realmente declara el mes que se está facturando.
    """
    try:
        match = re.search(patron, texto, re.IGNORECASE)
    except re.error:
        return None
    if match is None:
        return None
    anio = match.group("anio")
    mes = match.group("mes").lower()
    mes_num = _MESES_NUM.get(mes)
    if not mes_num:
        return None
    ultimo_dia = calendar.monthrange(int(anio), mes_num)[1]
    return f"{anio}-{mes_num:02d}-01", f"{anio}-{mes_num:02d}-{ultimo_dia:02d}"


def _fecha_iso_valida(valor: object) -> bool:
    try:
        datetime.strptime(str(valor), "%Y-%m-%d")
    except (TypeError, ValueError):
        return False
    return True


def extraer_datos_factura(texto: str, config: dict) -> dict:
    """
    Extrae los campos de una factura con estrategia en tres capas:
      Capa A — regex del proveedor (proveedores.json)      → intento específico
      Capa B — anclas genéricas por palabras clave legales → si A falla en críticos
      Capa C — LLM Gemini API (GEMINI_API_KEY en entorno)  → último recurso
    """
    regex = config.get("regex", {})
    regex_preferidos = config.get("regex_preferidos", {})
    campos_con_puntos_millares = set(config.get("campos_con_puntos_millares", []))
    tiene_decorativos = config.get("importe_tiene_puntos_decorativos", False)
    datos: dict = {}

    # ── Capa A: regex del proveedor ──────────────────────────────────────────
    for campo, patron in regex.items():
        if campo.startswith("linea_"):
            continue  # campos especiales de Ríos, se tratan aparte
        datos[campo] = _extraer_campo_preferido(
            texto, patron, regex_preferidos.get(campo, [])
        )

    # ── Capas B y C: solo para campos críticos que Capa A no encontró ────────
    # Si el proveedor usa puntos decorativos, importe_total se maneja
    # con _limpiar_importe_decorativo más abajo — excluirlo aquí.
    _criticos = [c for c in ["importe_total", "consumo_kwh", "consumo_m3",
                              "iva", "termino_fijo", "termino_variable"]
                 if not (c == "importe_total" and tiene_decorativos)]

    _fallos_a = [c for c in _criticos if not datos.get(c)]
    if _fallos_a:
        for campo in _fallos_a:
            val_b = _extraer_campo_capa_b(texto, campo)
            if val_b:
                datos[campo] = val_b

    # ── Normalizar fechas ─────────────────────────────────────────────────────
    for campo_fecha in ["fecha_factura", "fecha_inicio", "fecha_fin"]:
        if datos.get(campo_fecha):
            datos[campo_fecha] = _normalizar_fecha(datos[campo_fecha])

    # ── Rescate genérico de periodo si fecha_inicio/fecha_fin siguen vacías ──
    if not datos.get("fecha_inicio") or not datos.get("fecha_fin"):
        rango = _extraer_rango_fechas_capa_b(texto)
        if rango:
            if not datos.get("fecha_inicio"):
                datos["fecha_inicio"] = rango[0]
            if not datos.get("fecha_fin"):
                datos["fecha_fin"] = rango[1]

    # Algunos servicios de lectura sólo imprimen el mes facturado. El perfil
    # declara su ancla para no confundirlo con cualquier mes citado en el PDF.
    if not datos.get("fecha_inicio") or not datos.get("fecha_fin"):
        patron_mes = config.get("regex_periodo_mes_facturado")
        rango = _extraer_rango_mes_facturado(texto, patron_mes) if patron_mes else None
        if rango:
            datos["fecha_inicio"] = datos.get("fecha_inicio") or rango[0]
            datos["fecha_fin"] = datos.get("fecha_fin") or rango[1]

    # Los perfiles de gastos sin período de prestación se imputan por fecha de
    # factura, pero sólo cuando la configuración lo declara explícitamente.
    if (
        config.get("periodo_por_fecha_factura")
        and _fecha_iso_valida(datos.get("fecha_factura"))
    ):
        datos["fecha_inicio"] = datos.get("fecha_inicio") or datos["fecha_factura"]
        datos["fecha_fin"] = datos.get("fecha_fin") or datos["fecha_factura"]

    # ── Normalizar números ────────────────────────────────────────────────────
    # Los proveedores con puntos decorativos (ENVAC) también los intercalan en
    # el IVA, no solo en importe_total — mismo tratamiento para ambos campos.
    _campos_decorativos = ("importe_total", "iva") if tiene_decorativos else ()
    for campo_num in ["consumo_m3", "consumo_kwh", "termino_fijo", "termino_variable",
                      "impuestos", "iva", "importe_total"]:
        if campo_num in _campos_decorativos:
            patron_imp = config.get("regex", {}).get(campo_num, "")
            datos[campo_num] = _limpiar_importe_decorativo(texto, patron_imp) if patron_imp else 0.0
        elif datos.get(campo_num):
            datos[campo_num] = _limpiar_numero(
                datos[campo_num], campo_num in campos_con_puntos_millares
            )
        else:
            datos[campo_num] = 0.0

    # ── Días facturados como entero ───────────────────────────────────────────
    if datos.get("dias_facturados"):
        try:
            datos["dias_facturados"] = int(datos["dias_facturados"])
        except ValueError:
            datos["dias_facturados"] = None

    # ── Campos derivados ──────────────────────────────────────────────────────
    datos["consumo_total"] = datos.get("consumo_kwh") or datos.get("consumo_m3") or 0.0
    datos["unidad_consumo"] = config.get("unidad_consumo")
    datos["tipo_suministro"] = config.get("tipo_suministro")
    datos["proveedor"] = config.get("nombre_display")

    return datos


# ---------------------------------------------------------------------------
# EXTRACCIÓN ESPECIAL — RÍOS RENOVABLES
# ---------------------------------------------------------------------------

_MESES_NUM = {
    "enero": 1, "febrero": 2, "marzo": 3, "abril": 4, "mayo": 5, "junio": 6,
    "julio": 7, "agosto": 8, "septiembre": 9, "octubre": 10, "noviembre": 11, "diciembre": 12,
    "ene": 1, "feb": 2, "mar": 3, "abr": 4, "may": 5, "jun": 6,
    "jul": 7, "ago": 8, "sep": 9, "sept": 9, "oct": 10, "nov": 11, "dic": 12,
}


def _rango_mes_texto(texto: str, anio: str) -> tuple[str, str] | None:
    """
    Ríos Renovables no imprime fechas de inicio/fin, solo 'mes <nombre>'
    (ej. 'Lectura contadores mes febrero ACS'). Se infiere el rango como
    el mes completo, usando el año de fecha_factura.
    """
    m = re.search(r"(?i)mes\s+(\w+)", texto)
    if not m or not anio or not anio.isdigit():
        return None
    mes_num = _MESES_NUM.get(m.group(1).lower())
    if not mes_num:
        return None
    ultimo_dia = calendar.monthrange(int(anio), mes_num)[1]
    return f"{anio}-{mes_num:02d}-01", f"{anio}-{mes_num:02d}-{ultimo_dia:02d}"


def extraer_datos_rios(texto: str, config: dict) -> dict:
    """Extrae datos de factura de Ríos Renovables (lectura de contadores)."""
    datos = extraer_datos_factura(texto, config)

    # Extraer líneas de ACS y Calef individualmente
    for campo in ["linea_acs", "linea_calef"]:
        patron = config["regex"].get(campo, "")
        if patron:
            m = re.search(patron, texto, re.IGNORECASE)
            if m:
                datos[campo] = {
                    "unidades": _limpiar_numero(m.group("unidades")),
                    "precio":   _limpiar_numero(m.group("precio")),
                    "importe":  _limpiar_numero(m.group("importe")),
                }

    # El periodo no viene como fecha explícita, solo "mes <nombre>" — se
    # infiere el mes completo usando el año de fecha_factura.
    if not datos.get("fecha_inicio") or not datos.get("fecha_fin"):
        anio = None
        if datos.get("fecha_factura"):
            m_anio = re.match(r"(\d{4})", datos["fecha_factura"])
            if m_anio:
                anio = m_anio.group(1)
        rango = _rango_mes_texto(texto, anio) if anio else None
        if rango:
            datos["fecha_inicio"] = datos.get("fecha_inicio") or rango[0]
            datos["fecha_fin"] = datos.get("fecha_fin") or rango[1]

    return datos


# ---------------------------------------------------------------------------
# EXTRACCIÓN ESPECIAL — OFICINA MUNICIPAL DEL AGUA (ZARAGOZA)
# ---------------------------------------------------------------------------

# La factura del Ayuntamiento factura abastecimiento y saneamiento (Ecociudad)
# como dos bloques independientes, cada uno con su propia cuota fija, cuota
# variable e IVA. Un regex de un solo match solo capturaba el primer bloque
# (abastecimiento) y perdía saneamiento por completo. Aquí se suman ambos.
_PAT_AGUA_ZGZ_CUOTA_FIJA = re.compile(
    r"(?i)cuota\s+fija.*?Importe\s*\(.\).*?\n(?P<dias>\d+)\s+[\d,]+\s+(?P<valor>[\d,]+)\s*\n",
    re.DOTALL,
)
_PAT_AGUA_ZGZ_CUOTA_VARIABLE = re.compile(r"(?i)Total\s+cuota\s+variable\s+(?P<valor>[\d.,]+)")
_PAT_AGUA_ZGZ_IVA = re.compile(r"(?i)IVA\s*\(\d+%\)\s*de\s*[\d.,]+\s+(?P<valor>[\d.,]+)")


def extraer_datos_agua_zaragoza(texto: str, config: dict) -> dict:
    """Extrae datos de la factura de la Oficina Municipal del Agua, sumando
    los bloques de abastecimiento y saneamiento (Ecociudad Zaragoza)."""
    datos = extraer_datos_factura(texto, config)

    fijos = [_limpiar_numero(m.group("valor")) for m in _PAT_AGUA_ZGZ_CUOTA_FIJA.finditer(texto)]
    variables = [_limpiar_numero(m.group("valor")) for m in _PAT_AGUA_ZGZ_CUOTA_VARIABLE.finditer(texto)]
    ivas = [_limpiar_numero(m.group("valor")) for m in _PAT_AGUA_ZGZ_IVA.finditer(texto)]

    if fijos:
        datos["termino_fijo"] = sum(fijos)
    if variables:
        datos["termino_variable"] = sum(variables)
    if ivas:
        datos["iva"] = sum(ivas)

    return datos


# ---------------------------------------------------------------------------
# EXTRACCIÓN ESPECIAL — METRIGEST (lecturas por vecino)
# ---------------------------------------------------------------------------

def extraer_lecturas_metrigest(texto_paginas: list[str], config: dict) -> dict:
    """
    Extrae lecturas individuales de todos los vecinos de un resumen Metrigest.
    Formato real (dos líneas por vecino):
      Línea 1: 'SS22 BAJO IZDA - SILVIA MORER AGUARON 6,00 € 0,00 € 6,00 €'
      Línea 2: 'ACS 30/01/2026 1,00 02/03/2026 2,00 1,00 M3 6,00 € 6,00 €'
    """
    import re as _re
    texto_completo = "\n".join(texto_paginas)

    # Cabecera
    cabecera = {}
    for campo, patron in config.get("regex_cabecera", {}).items():
        m = _re.search(patron, texto_completo, _re.IGNORECASE)
        if m:
            if campo == "periodo":
                cabecera["fecha_inicio"] = _normalizar_fecha(m.group("inicio"))
                cabecera["fecha_fin"]    = _normalizar_fecha(m.group("fin"))
            else:
                cabecera[campo] = m.group("valor").strip()

    tipo = "ACS" if "ACS" in config.get("tipo_suministro", "") else "CALEFACCION"

    # Extraer vecinos línea a línea
    lineas = texto_completo.split("\n")
    vecinos = []
    i = 0
    while i < len(lineas) - 1:
        linea1 = lineas[i].strip()
        linea2 = lineas[i + 1].strip()

        # Línea 1: vivienda - nombre ... euros
        m1 = _re.match(
            r"^([A-Z]{2}\d{2}\s+[\w\s]+?)\s+-\s+([A-ZÁÉÍÓÚÑÜ][\w\s]+?)\s+[\d,\.]+\s+€",
            linea1, _re.IGNORECASE
        )
        if not m1:
            i += 1
            continue

        # Línea 2: ACS/CALEF fecha val fecha val consumo M3/kWh
        m2 = _re.match(
            r"^(ACS|CALEFACCION)\s+"
            r"(\d{2}/\d{2}/\d{4})\s+([\d,\.]+)\s+"
            r"(\d{2}/\d{2}/\d{4})\s+([\d,\.]+)\s+"
            r"([\d,\.]+)\s*(M3|kWh)",
            linea2, _re.IGNORECASE
        )
        if not m2:
            i += 1
            continue

        vecinos.append({
            "vivienda":  m1.group(1).strip(),
            "nombre":    m1.group(2).strip(),
            "tipo":      tipo,
            "fecha_ant": _normalizar_fecha(m2.group(2)),
            "val_ant":   _limpiar_numero(m2.group(3)),
            "fecha_act": _normalizar_fecha(m2.group(4)),
            "val_act":   _limpiar_numero(m2.group(5)),
            "consumo":   _limpiar_numero(m2.group(6)),
            "unidad":    m2.group(7).upper(),
        })
        i += 2  # saltar las dos líneas consumidas

    return {"cabecera": cabecera, "vecinos": vecinos, "tipo": tipo}


def procesar_archivo(ruta_archivo: str, codigo_comunidad: str = None,
                     con_bd=None, ruta_proveedores: str = None,
                     proveedores: dict | None = None) -> dict:
    """
    Procesa un archivo de factura o resumen Metrigest.

    Validación en tres niveles (se detiene en el primero que falla):
      Nivel 1 — Nombre del archivo: extrae código de comunidad si viene en el nombre
      Nivel 2 — CIF en el PDF: busca en la BD para confirmar la comunidad
      Nivel 3 — CUPS en el PDF: confirma que coincide con el CUPS registrado

    Args:
        ruta_archivo:       Ruta completa al archivo
        codigo_comunidad:   Código conocido (ej: '644'). Si es None, se autodetecta por CIF.
        con_bd:             Conexión SQLite activa (opcional). Si se pasa, habilita autodetección por CIF.
        ruta_proveedores:   Ruta a proveedores.json (opcional)

    Returns:
        dict con: ok, tipo, datos, proveedor_clave, hash_md5, nombre_archivo,
                  codigo_comunidad, cif_pdf, validacion (dict con niveles)
    """
    nombre = os.path.basename(ruta_archivo)
    hash_md5 = _md5_archivo(ruta_archivo)
    validacion = {"nivel_1_archivo": None, "nivel_2_cif": None, "nivel_3_cups": None}

    # 1. Los PDFs de imagen se detectan sin recorrerlos con pdfplumber: así el
    # OCR de portada evita bloqueos en facturas escaneadas de varias páginas.
    ruta = Path(ruta_archivo)
    texto = "" if _pdf_probablemente_escaneado(ruta) else extraer_texto(ruta_archivo)

    # Si no hay texto, intentar OCR automáticamente
    ocr_result = None
    if not texto.strip():
        ocr_result = extraer_texto_ocr_con_diagnostico(ruta_archivo)
        texto = ocr_result.text

    if not texto.strip():
        diagnostic = ocr_result.detail if ocr_result is not None else "No se obtuvo texto legible."
        return {"ok": False, "motivo": "SIN_TEXTO",
                "detalle": (f"No se pudo extraer texto de {nombre}. {diagnostic} "
                            "Abre el archivo y revisa la clasificación o los datos necesarios."),
                "nombre_archivo": nombre, "hash_md5": hash_md5,
                "requiere_ocr": True}

    # 2. Cargar proveedores e identificar proveedor
    proveedores = proveedores if proveedores is not None else cargar_proveedores(ruta_proveedores)
    clave_prov, config_prov = identificar_proveedor(texto, nombre, proveedores)

    if not clave_prov:
        return {"ok": False, "motivo": "PROVEEDOR_NO_IDENTIFICADO",
                "detalle": f"Ningún proveedor reconocido en {nombre}",
                "fragment": re.sub(r"\s+", " ", texto).strip()[:1000],
                "nombre_archivo": nombre, "hash_md5": hash_md5}

    # ----------------------------------------------------------------
    # NIVEL 1 — Comunidad por nombre de archivo
    # ----------------------------------------------------------------
    codigo_desde_archivo = codigo_comunidad  # puede venir del caller (que lo leyó del nombre)
    validacion["nivel_1_archivo"] = "ok" if codigo_desde_archivo else "no_disponible"

    # ----------------------------------------------------------------
    # NIVEL 2 — CIF en el PDF
    # ----------------------------------------------------------------
    cif_pdf = extraer_cif_pdf(texto)
    comunidad_desde_cif = None

    if cif_pdf and con_bd is not None:
        # Importar aquí para evitar dependencia circular si se usa el módulo solo
        try:
            from gestor_bd import buscar_comunidad_por_cif
            comunidad_desde_cif = buscar_comunidad_por_cif(con_bd, cif_pdf)
        except ImportError:
            pass

    if comunidad_desde_cif:
        codigo_por_cif = comunidad_desde_cif["codigo"]
        if codigo_desde_archivo and codigo_por_cif != codigo_desde_archivo:
            # Posible contradicción — pero primero comprobamos si el código extraído
            # del nombre del archivo es realmente un código de comunidad conocido,
            # o simplemente texto de nombre de archivo sin formato estándar.
            # Si el nombre del archivo NO empieza por "CODIGO_" (ej: "Endesa2v2-644.pdf"),
            # el código extraído es basura: confiamos en el CIF del PDF.
            nombre_sin_ext = os.path.splitext(nombre)[0]
            # Código de comunidad válido: solo dígitos (ej: "644"), o empieza por él
            codigo_parece_valido = (
                codigo_desde_archivo.isdigit() or
                re.match(r'^\d+[_\-]', nombre_sin_ext)
            )
            if codigo_parece_valido:
                # Contradicción real: el archivo dice una comunidad y el CIF dice otra
                validacion["nivel_2_cif"] = "contradiccion"
                return {
                    "ok": False,
                    "motivo": "COMUNIDAD_CONTRADICTORIA",
                    "detalle": (f"Nombre archivo sugiere comunidad '{codigo_desde_archivo}' "
                                f"pero CIF {cif_pdf} pertenece a '{codigo_por_cif}'"),
                    "nombre_archivo": nombre, "hash_md5": hash_md5,
                    "cif_pdf": cif_pdf, "validacion": validacion,
                }
            else:
                # El "código" del nombre del archivo no es un código real (ej: "Endesa2v2")
                # → confiamos en el CIF del PDF y continuamos
                validacion["nivel_2_cif"] = "cif_prevalece_sobre_nombre"
        codigo_comunidad = codigo_por_cif
        validacion["nivel_2_cif"] = validacion.get("nivel_2_cif") or "ok"

    elif cif_pdf:
        validacion["nivel_2_cif"] = "cif_no_en_bd"
    else:
        validacion["nivel_2_cif"] = "sin_cif_en_pdf"

    # Sin comunidad identificada por ningún medio → cuarentena
    if not codigo_comunidad:
        return {
            "ok": False,
            "motivo": "COMUNIDAD_NO_IDENTIFICADA",
            "detalle": "No se pudo determinar la comunidad ni por nombre de archivo ni por CIF",
            "cif_pdf": cif_pdf,
            "nombre_archivo": nombre, "hash_md5": hash_md5,
            "validacion": validacion,
        }

    # ----------------------------------------------------------------
    # NIVEL 3 — CUPS en el PDF vs. CUPS registrado para esta comunidad
    # ----------------------------------------------------------------
    cups_comunidades = config_prov.get("cups_comunidades", {})
    cups_esperado = cups_comunidades.get(codigo_comunidad)

    # Obtener lista de CUPS válidos para esta comunidad (puede haber más de uno)
    cups_multiples = config_prov.get("cups_multiples", {}).get(codigo_comunidad, [])
    if cups_esperado and cups_esperado not in cups_multiples:
        cups_multiples = [cups_esperado] + cups_multiples

    if cups_multiples:
        patron_cups = config_prov.get("regex", {}).get("cups", "")
        cups_en_pdf = None
        if patron_cups:
            m = re.search(patron_cups, texto, re.IGNORECASE)
            if m:
                cups_en_pdf = m.group("valor").strip().upper()

        if cups_en_pdf and cups_en_pdf not in [c.upper() for c in cups_multiples]:
            validacion["nivel_3_cups"] = "contradiccion"
            return {
                "ok": False,
                "motivo": "CUPS_NO_COINCIDE",
                "detalle": f"PDF: {cups_en_pdf} | Válidos: {cups_multiples}",
                "cups_pdf": cups_en_pdf, "cups_esperado": cups_multiples,
                "nombre_archivo": nombre, "hash_md5": hash_md5,
                "validacion": validacion,
            }
        validacion["nivel_3_cups"] = "ok" if cups_en_pdf else "sin_cups"
    elif cups_esperado:
        patron_cups = config_prov.get("regex", {}).get("cups", "")
        cups_en_pdf = None
        if patron_cups:
            m = re.search(patron_cups, texto, re.IGNORECASE)
            if m:
                cups_en_pdf = m.group("valor").strip().upper()
        if cups_en_pdf and cups_en_pdf != cups_esperado.upper():
            validacion["nivel_3_cups"] = "contradiccion"
            return {
                "ok": False,
                "motivo": "CUPS_NO_COINCIDE",
                "detalle": f"PDF: {cups_en_pdf} | Esperado: {cups_esperado}",
                "cups_pdf": cups_en_pdf, "cups_esperado": cups_esperado,
                "nombre_archivo": nombre, "hash_md5": hash_md5,
                "validacion": validacion,
            }
        validacion["nivel_3_cups"] = "ok" if cups_en_pdf else "sin_cups"
    else:
        validacion["nivel_3_cups"] = "no_aplica"  # Ríos, Metrigest no tienen CUPS

    # ----------------------------------------------------------------
    # Extracción de datos
    # ----------------------------------------------------------------
    tipo_suministro = config_prov.get("tipo_suministro", "")

    # Justificante de pago bancario — registrar como pago, no como factura
    if config_prov.get("es_justificante"):
        datos = extraer_datos_factura(texto, config_prov)
        datos["cups_o_referencia"] = None
        datos["archivo_origen"] = nombre
        return {
            "ok": True,
            "tipo": "JUSTIFICANTE_PAGO",
            "proveedor_clave": clave_prov,
            "datos": datos,
            "codigo_comunidad": codigo_comunidad,
            "cif_pdf": cif_pdf,
            "validacion": validacion,
            "nombre_archivo": nombre,
            "hash_md5": hash_md5,
        }

    if tipo_suministro in ("LECTURA_ACS", "LECTURA_CALEF"):
        paginas = []
        if zipfile.is_zipfile(ruta_archivo):
            with zipfile.ZipFile(ruta_archivo) as z:
                nombres_txt = sorted(
                    [n for n in z.namelist() if n.endswith(".txt")],
                    key=lambda x: int(x.replace(".txt", "")) if x.replace(".txt", "").isdigit() else 99
                )
                for n in nombres_txt:
                    with z.open(n) as f:
                        paginas.append(f.read().decode("utf-8", errors="replace"))
        else:
            paginas = [texto]

        datos = extraer_lecturas_metrigest(paginas, config_prov)
        return {
            "ok": True, "tipo": "LECTURA_METRIGEST",
            "proveedor_clave": clave_prov, "datos": datos,
            "codigo_comunidad": codigo_comunidad,
            "cif_pdf": cif_pdf, "validacion": validacion,
            "nombre_archivo": nombre, "hash_md5": hash_md5,
        }

    elif clave_prov == "RIOS_RENOVABLES":
        datos = extraer_datos_rios(texto, config_prov)
    elif clave_prov == "AGUA_ZARAGOZA":
        datos = extraer_datos_agua_zaragoza(texto, config_prov)
    else:
        datos = extraer_datos_factura(texto, config_prov)

    datos["cups_o_referencia"] = cups_esperado or datos.get("cups")
    datos["archivo_origen"] = nombre

    if datos.get("importe_total") is None or datos["importe_total"] == 0:
        return {
            "ok": False, "motivo": "IMPORTE_CERO",
            "detalle": f"No se extrajo importe_total en {nombre}",
            "nombre_archivo": nombre, "hash_md5": hash_md5,
            "datos_parciales": datos, "validacion": validacion,
        }

    return {
        "ok": True, "tipo": "FACTURA",
        "proveedor_clave": clave_prov, "datos": datos,
        "codigo_comunidad": codigo_comunidad,
        "cif_pdf": cif_pdf, "validacion": validacion,
        "nombre_archivo": nombre, "hash_md5": hash_md5,
    }


# ---------------------------------------------------------------------------
# EJECUCIÓN DIRECTA — test sobre los PDFs del proyecto
# ---------------------------------------------------------------------------

if __name__ == "__main__":
    import sys

    carpeta = sys.argv[1] if len(sys.argv) > 1 else "."
    comunidad = sys.argv[2] if len(sys.argv) > 2 else "644"

    archivos = [f for f in os.listdir(carpeta)
                if f.lower().endswith(".pdf") and not f.startswith("~")]

    if not archivos:
        print(f"No se encontraron archivos .pdf en {carpeta}")
        sys.exit(0)

    for nombre in sorted(archivos):
        ruta = os.path.join(carpeta, nombre)
        resultado = procesar_archivo(ruta, comunidad)
        print(f"\n{'='*60}")
        print(f"Archivo: {nombre}")
        if resultado["ok"]:
            tipo = resultado["tipo"]
            print(f"  Estado:    ✅ OK — {tipo}")
            print(f"  Proveedor: {resultado['proveedor_clave']}")
            if tipo == "FACTURA":
                d = resultado["datos"]
                print(f"  Factura:   {d.get('num_factura','?')}")
                print(f"  Periodo:   {d.get('fecha_inicio','?')} → {d.get('fecha_fin','?')}")
                print(f"  Consumo:   {d.get('consumo_total',0):,.0f} {d.get('unidad_consumo','')}")
                print(f"  T.Fijo:    {d.get('termino_fijo',0):,.2f} €")
                print(f"  T.Variable:{d.get('termino_variable',0):,.2f} €")
                print(f"  IVA:       {d.get('iva',0):,.2f} €")
                print(f"  TOTAL:     {d.get('importe_total',0):,.2f} €")
            elif tipo == "LECTURA_METRIGEST":
                d = resultado["datos"]
                cab = d.get("cabecera", {})
                vecs = d.get("vecinos", [])
                print(f"  Comunidad: {cab.get('comunidad','?')}")
                print(f"  Periodo:   {cab.get('fecha_inicio','?')} → {cab.get('fecha_fin','?')}")
                print(f"  Vecinos:   {len(vecs)} registros")
                for v in vecs[:3]:
                    print(f"    {v['vivienda']:<20} {v['val_ant']} → {v['val_act']} ({v['consumo']} {v['unidad']})")
                if len(vecs) > 3:
                    print(f"    ... y {len(vecs)-3} más")
        else:
            print(f"  Estado:  ❌ {resultado['motivo']}")
            print(f"  Detalle: {resultado.get('detalle','')}")
