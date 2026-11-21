"""Importador: lee un Word de planificaciones y lo convierte a datos (JSON).

Es lo contrario del generador: en vez de datos -> Word, hace Word -> datos.
Así no hay que volver a teclear lo que ya está transcrito.

Cómo reconoce las tablas (sin depender de la posición exacta):
  - Si contiene "Experiencia (Grupo Grande)"   -> hoja grupal
  - Si su primer encabezado es "Indicador(es)" -> hoja del indicador
Cada hoja grupal se empareja con la hoja de indicador que le sigue.
"""

from __future__ import annotations

import json
import re
from pathlib import Path

from docx import Document
from docx.text.paragraph import Paragraph

TITULO_PEQUENO = "Actividad Grupal (Expresión plástica)"
TITULO_GRUPAL = "Actividad Grupal (Expresión Plástica)."

CONTENIDOS = ("conceptuales", "procedimentales", "actitudinales")

# Etiqueta que aparece en el Word -> clave en los datos
ETIQUETAS_GENERALES = {
    "Centro Educativo": "centro",
    "Docente": "docente",
    "Regional": "regional",
    "Distrito": "distrito",
    "Grado": "grado",
    "Sección": "seccion",
    "Tanda": "tanda",
    "Nombre Lúdico": "nombre_ludico",
    "Fecha de inicio": "fecha_inicio",
    "Fecha de cierre": "fecha_cierre",
    "Tipo de planificación": "tipo",
    "Tema": "tema",
}
# Una etiqueta solo cuenta si termina en ":" (así "Docente de Ejemplo" no se
# confunde). Excepción: "Grado" al inicio de línea, que en tu Word va sin ":".
_NOMBRES = "|".join(map(re.escape, ETIQUETAS_GENERALES))
_PATRON_ETIQUETA = re.compile(rf"\b({_NOMBRES})\s*:|^(Grado)\b\s*:?")


# ---------------------------------------------------------------------------
# Utilidades para leer celdas
# ---------------------------------------------------------------------------
def _celdas_unicas(fila) -> list:
    """Las celdas combinadas aparecen repetidas; nos quedamos con una sola."""
    vistas, salida = set(), []
    for celda in fila.cells:
        if id(celda._tc) not in vistas:
            vistas.add(id(celda._tc))
            salida.append(celda)
    return salida


def _parrafos(celda) -> list[str]:
    """Párrafos no vacíos de una celda, sin espacios sobrantes."""
    textos = (" ".join(p.text.split()) for p in celda.paragraphs)
    return [t for t in textos if t]


def _recursos(celda) -> list[str]:
    """Una lista de recursos: quita viñetas como '- ', '• ' o '* '."""
    return [re.sub(r"^[-•*\s]+", "", t) for t in _parrafos(celda)
            if re.sub(r"^[-•*\s]+", "", t)]


def _texto(celda) -> str:
    return " ".join(celda.text.split())


# ---------------------------------------------------------------------------
# Lectura de cada tipo de hoja
# ---------------------------------------------------------------------------
def _leer_grupal(tabla) -> dict:
    filas = [_celdas_unicas(f) for f in tabla.rows]

    def titulo(celda, por_defecto):
        t = _texto(celda)
        return t if t.startswith("Actividad Grupal") else por_defecto

    return {
        "pequeno": {
            "titulo": titulo(filas[0][0], TITULO_PEQUENO),
            "actividad": _parrafos(filas[1][0]),
            "recursos": _recursos(filas[1][1]) if len(filas[1]) > 1 else [],
        },
        "grande": {
            "actividad": _parrafos(filas[4][0]),
            "recursos": _recursos(filas[4][1]) if len(filas[4]) > 1 else [],
        },
        "actividad_grupal": {
            "titulo": titulo(filas[6][0], TITULO_GRUPAL),
            "actividad": _parrafos(filas[7][0]),
            "recursos": _recursos(filas[7][1]) if len(filas[7]) > 1 else [],
        },
    }


def _leer_contenidos(celda) -> dict:
    """Separa el texto bajo 'Conceptuales:', 'Procedimentales:', 'Actitudinales:'."""
    resultado = {clave: [] for clave in CONTENIDOS}
    actual = None
    for linea in _parrafos(celda):
        encabezado, _, resto = linea.partition(":")
        clave = encabezado.strip().lower()
        if clave in CONTENIDOS:
            actual = clave
            if resto.strip():
                resultado[actual].append(resto.strip())
        elif actual:
            resultado[actual].append(linea)
    return resultado


def _leer_indicador(tabla) -> dict:
    # Fila 0 = encabezados. Fila 1 = el contenido (las filas siguientes, si
    # existen, son la misma celda combinada hacia abajo).
    celdas = _celdas_unicas(tabla.rows[1])
    hoja = {"indicador": re.sub(r"^[-•*\s]+", "", _texto(celdas[0]))}
    if len(celdas) == 4:
        hoja["contenidos"] = _leer_contenidos(celdas[1])
        hoja["actividad"] = _parrafos(celdas[2])
        hoja["recursos"] = _recursos(celdas[3])
    else:
        hoja["actividad"] = _parrafos(celdas[1])
        hoja["recursos"] = _recursos(celdas[2])
    return hoja


def _limpiar_valor(valor: str) -> str:
    """Quita espacios y las rayas '____' que usan como espacio en blanco."""
    return valor.strip().strip("_").strip()


def _leer_generales(lineas: list[str]) -> dict:
    """Convierte las líneas 'Etiqueta: valor   Etiqueta: valor' en un diccionario.

    Si una etiqueta aparece repetida, gana la última (en tu Word, 'Nombre
    Lúdico' sale en la línea 2 y otra vez junto al Tema; la buena es la última).
    """
    generales = {}
    for linea in lineas:
        linea = " ".join(linea.split())
        coincidencias = list(_PATRON_ETIQUETA.finditer(linea))
        for i, m in enumerate(coincidencias):
            fin = coincidencias[i + 1].start() if i + 1 < len(coincidencias) else len(linea)
            etiqueta = m.group(1) or m.group(2)
            generales[ETIQUETAS_GENERALES[etiqueta]] = _limpiar_valor(linea[m.end():fin])
    return generales


def _lineas_antes_de(tabla) -> list[str]:
    """Textos de los párrafos que están justo encima de una tabla."""
    lineas = []
    for hermano in tabla._tbl.itersiblings(preceding=True):
        if hermano.tag.endswith("}tbl"):
            break
        # Paragraph.text lee solo el texto visible (itertext trae contenido duplicado)
        texto = Paragraph(hermano, tabla._parent).text.strip() if hermano.tag.endswith("}p") else ""
        if texto:
            lineas.append(texto)
    return lineas[::-1]


def _leer_mensual(tabla_tema, tabla_kwl) -> dict:
    kwl = _celdas_unicas(tabla_kwl.rows[1])
    return {
        "generales": _leer_generales(_lineas_antes_de(tabla_tema)),
        "seleccion_tema": _texto(tabla_tema.rows[1].cells[0]),
        "que_sabemos": _recursos(kwl[0]),
        "que_queremos_saber": _recursos(kwl[1]),
        "como_lo_vamos_a_saber": _recursos(kwl[2]),
    }


# ---------------------------------------------------------------------------
# API pública
# ---------------------------------------------------------------------------
def _es_grupal(tabla) -> bool:
    return "Experiencia (Grupo Grande)" in tabla.rows[2].cells[0].text if len(tabla.rows) > 2 else False


def _es_indicador(tabla) -> bool:
    return _texto(tabla.rows[0].cells[0]).lower().startswith("indicador")


def importar_documento(ruta_docx: str | Path) -> dict:
    """Devuelve los datos de todo el Word en el formato que entiende el generador."""
    doc = Document(str(ruta_docx))
    dias, grupal_pendiente, mensual_tema, mensual_kwl = [], None, None, None

    for tabla in doc.tables:
        primer_texto = _texto(tabla.rows[0].cells[0])
        if _es_grupal(tabla):
            grupal_pendiente = _leer_grupal(tabla)
        elif _es_indicador(tabla) and grupal_pendiente is not None:
            dias.append({
                "numero": len(dias) + 1,
                "grupal": grupal_pendiente,
                "hoja2": _leer_indicador(tabla),
            })
            grupal_pendiente = None
        elif primer_texto.lower().startswith("planificación mensual"):
            mensual_tema = tabla
        elif primer_texto.startswith("¿Qué sabemos"):
            mensual_kwl = tabla

    datos = {"dias": dias}
    if mensual_tema is not None and mensual_kwl is not None:
        datos["mensual"] = _leer_mensual(mensual_tema, mensual_kwl)
    return datos


def importar_archivo(ruta_docx: str | Path, ruta_json: str | Path) -> Path:
    datos = importar_documento(ruta_docx)
    ruta_json = Path(ruta_json)
    ruta_json.parent.mkdir(parents=True, exist_ok=True)
    ruta_json.write_text(json.dumps(datos, ensure_ascii=False, indent=2),
                         encoding="utf-8")
    return ruta_json
