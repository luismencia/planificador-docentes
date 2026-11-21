"""Generador de planificaciones en Word (.docx).

Recibe datos (diccionarios de Python, normalmente cargados desde un JSON)
y produce un documento con el diseño que exige la escuela:
hoja apaisada, Times New Roman, encabezados azul claro y tablas con borde.

Cada día de clase genera DOS hojas:
  1) "Experiencia Grupal": tabla de actividades grupales con recursos.
  2) Hoja del indicador de logro (3 columnas, o 4 si lleva contenidos).
"""

from __future__ import annotations

import json
from pathlib import Path

from docx import Document
from docx.enum.section import WD_ORIENT
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Pt, Twips

# ---------------------------------------------------------------------------
# Constantes de diseño (medidas tomadas del Word original, en "twips":
# 1 twip = 1/1440 de pulgada)
# ---------------------------------------------------------------------------
FUENTE = "Times New Roman"
TAMANO = Pt(16)
COLOR_ENCABEZADO = "C5D6EE"

ANCHO_PAGINA = 15840  # carta apaisada
ALTO_PAGINA = 12240
MARGEN_LATERAL = 360
MARGEN_SUPERIOR = 880
MARGEN_INFERIOR = 280
SANGRIA_TABLA = 279

COLS_GRUPAL = (11482, 3260)
COLS_INDICADOR = (3224, 7519, 3332)
COLS_INDICADOR_CONTENIDOS = (2569, 3641, 5370, 2895)


# ---------------------------------------------------------------------------
# Utilidades de bajo nivel (python-docx no cubre todo, así que tocamos el XML)
# ---------------------------------------------------------------------------
def _sombrear(celda, color: str) -> None:
    tc_pr = celda._tc.get_or_add_tcPr()
    shd = OxmlElement("w:shd")
    shd.set(qn("w:val"), "clear")
    shd.set(qn("w:color"), "auto")
    shd.set(qn("w:fill"), color)
    tc_pr.append(shd)


def _bordes_tabla(tabla) -> None:
    tbl_pr = tabla._tbl.tblPr
    bordes = OxmlElement("w:tblBorders")
    for lado in ("top", "left", "bottom", "right", "insideH", "insideV"):
        el = OxmlElement(f"w:{lado}")
        el.set(qn("w:val"), "single")
        el.set(qn("w:sz"), "4")
        el.set(qn("w:space"), "0")
        el.set(qn("w:color"), "000000")
        bordes.append(el)
    tbl_pr.append(bordes)


def _layout_fijo(tabla, anchos: tuple[int, ...]) -> None:
    """Fija el ancho de cada columna (si no, Word las reparte a su gusto)."""
    tabla.autofit = False
    tbl_pr = tabla._tbl.tblPr

    ancho_total = OxmlElement("w:tblW")
    ancho_total.set(qn("w:w"), str(sum(anchos)))
    ancho_total.set(qn("w:type"), "dxa")
    for viejo in tbl_pr.findall(qn("w:tblW")):
        tbl_pr.remove(viejo)
    tbl_pr.append(ancho_total)

    sangria = OxmlElement("w:tblInd")
    sangria.set(qn("w:w"), str(SANGRIA_TABLA))
    sangria.set(qn("w:type"), "dxa")
    tbl_pr.append(sangria)

    layout = OxmlElement("w:tblLayout")
    layout.set(qn("w:type"), "fixed")
    tbl_pr.append(layout)

    for col, ancho in zip(tabla.columns, anchos):
        col.width = Twips(ancho)


def _ancho_celda(celda, ancho: int) -> None:
    celda.width = Twips(ancho)


def _formato_run(run, negrita=False, cursiva=False, subrayado=False, tamano=None) -> None:
    run.font.name = FUENTE
    run._element.rPr.rFonts.set(qn("w:cs"), FUENTE)
    run._element.rPr.rFonts.set(qn("w:hAnsi"), FUENTE)
    run.font.size = tamano or TAMANO
    run.bold = negrita
    run.italic = cursiva
    run.underline = subrayado


def _escribir(celda, textos, *, negrita=False, cursiva=False,
              centrado=False, prefijo="", espacio_despues=6,
              tamano=None, justificado=False) -> None:
    """Escribe uno o varios párrafos dentro de una celda."""
    if isinstance(textos, str):
        textos = [textos]
    if not textos:
        textos = [""]
    primero = True
    for texto in textos:
        parrafo = celda.paragraphs[0] if primero else celda.add_paragraph()
        primero = False
        parrafo.paragraph_format.space_after = Pt(espacio_despues)
        parrafo.paragraph_format.space_before = Pt(0)
        if centrado:
            parrafo.alignment = WD_ALIGN_PARAGRAPH.CENTER
        elif justificado:
            parrafo.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
        run = parrafo.add_run(f"{prefijo}{texto}" if texto else "")
        _formato_run(run, negrita, cursiva, tamano=tamano)


def _fila_encabezado(tabla, textos, anchos=None) -> None:
    fila = tabla.add_row()
    for celda, texto in zip(fila.cells, textos):
        _sombrear(celda, COLOR_ENCABEZADO)
        _escribir(celda, texto, negrita=True, cursiva=True, centrado=True,
                  espacio_despues=0)


def _fila_combinada(tabla, texto, *, sombreada=False, negrita=True,
                    cursiva=True, centrado=True) -> None:
    fila = tabla.add_row()
    celda = fila.cells[0].merge(fila.cells[-1])
    if sombreada:
        _sombrear(celda, COLOR_ENCABEZADO)
    _escribir(celda, texto, negrita=negrita, cursiva=cursiva,
              centrado=centrado, espacio_despues=0)


def _fila_contenido(tabla, actividad, recursos) -> None:
    fila = tabla.add_row()
    _escribir(fila.cells[0], actividad, cursiva=True)
    _escribir(fila.cells[1], recursos, cursiva=True, prefijo="- ",
              espacio_despues=3)


def _aplicar_anchos(tabla, anchos) -> None:
    for fila in tabla.rows:
        for celda, ancho in zip(fila.cells, anchos):
            _ancho_celda(celda, ancho)


# ---------------------------------------------------------------------------
# Hojas
# ---------------------------------------------------------------------------
def _titulo_semana(documento, numero, semana: dict | None, nueva_pagina=False) -> None:
    semana = semana or {}
    del_ = semana.get("del") or "____________"
    al = semana.get("al") or "______________"
    de = semana.get("de") or "_________________"
    p = documento.add_paragraph()
    p.paragraph_format.page_break_before = nueva_pagina
    p.paragraph_format.space_after = Pt(6)
    r1 = p.add_run(f"{numero}    ")
    _formato_run(r1, negrita=True, cursiva=True)
    r2 = p.add_run(f"Experiencia Grupal (grupo pequeño) semana {del_} al {al} de {de}")
    _formato_run(r2, negrita=True, cursiva=True)


def hoja_grupal(documento, dia: dict, nueva_pagina=False) -> None:
    g = dia["grupal"]
    _titulo_semana(documento, dia["numero"], dia.get("semana"), nueva_pagina)

    tabla = documento.add_table(rows=0, cols=2)
    tabla.alignment = WD_TABLE_ALIGNMENT.LEFT
    _bordes_tabla(tabla)
    _layout_fijo(tabla, COLS_GRUPAL)

    # Grupo pequeño
    pequeno = g["pequeno"]
    _fila_encabezado(tabla, [pequeno.get("titulo", "Actividad Grupal (Expresión plástica)"),
                             "Recursos"])
    _fila_contenido(tabla, pequeno["actividad"], pequeno.get("recursos", []))

    # Grupo grande
    _fila_combinada(tabla, "Experiencia (Grupo Grande)")
    _fila_encabezado(tabla, ["Actividad", "Recursos"])
    _fila_contenido(tabla, g["grande"]["actividad"], g["grande"].get("recursos", []))

    # Actividad grupal
    _fila_combinada(
        tabla,
        "Actividad Grupal: Literatura Infantil, expresión musical, expresión "
        "corporal, psicomotricidad, expresión plástica.",
        sombreada=False, centrado=False,
    )
    ag = g["actividad_grupal"]
    _fila_encabezado(tabla, [ag.get("titulo", "Actividad Grupal (Expresión Plástica)."),
                             "Recursos"])
    _fila_contenido(tabla, ag["actividad"], ag.get("recursos", []))

    _aplicar_anchos(tabla, COLS_GRUPAL)


def hoja_indicador(documento, dia: dict) -> None:
    h = dia["hoja2"]
    p = documento.add_paragraph()
    p.paragraph_format.page_break_before = True
    _formato_run(p.add_run(str(dia["numero"])), negrita=True, cursiva=True)

    contenidos = h.get("contenidos")
    if contenidos:
        anchos = COLS_INDICADOR_CONTENIDOS
        encabezados = ["Indicadores", "Contenidos:", "Actividad", "Recursos"]
    else:
        anchos = COLS_INDICADOR
        encabezados = ["Indicador de Logro", "Actividad", "Recursos"]

    tabla = documento.add_table(rows=0, cols=len(anchos))
    tabla.alignment = WD_TABLE_ALIGNMENT.LEFT
    _bordes_tabla(tabla)
    _layout_fijo(tabla, anchos)
    _fila_encabezado(tabla, encabezados)

    fila = tabla.add_row()
    _escribir(fila.cells[0], h["indicador"], cursiva=True)
    if contenidos:
        celda = fila.cells[1]
        etiquetas = (("conceptuales", "Conceptuales:"),
                     ("procedimentales", "Procedimentales:"),
                     ("actitudinales", "Actitudinales:"))
        primero = True
        for clave, etiqueta in etiquetas:
            parrafo = celda.paragraphs[0] if primero else celda.add_paragraph()
            primero = False
            _formato_run(parrafo.add_run(etiqueta), negrita=True, cursiva=True)
            for texto in contenidos.get(clave, []):
                parrafo = celda.add_paragraph()
                _formato_run(parrafo.add_run(texto), cursiva=True)
        _escribir(fila.cells[2], h["actividad"])
        _escribir(fila.cells[3], h.get("recursos", []), cursiva=True,
                  prefijo="- ", espacio_despues=3)
    else:
        _escribir(fila.cells[1], h["actividad"])
        _escribir(fila.cells[2], h.get("recursos", []), cursiva=True,
                  prefijo="- ", espacio_despues=3)

    # Altura mínima para que la tabla se vea completa, como en el original
    tr_pr = fila._tr.get_or_add_trPr()
    alto = OxmlElement("w:trHeight")
    alto.set(qn("w:val"), "6500")
    alto.set(qn("w:hRule"), "atLeast")
    tr_pr.append(alto)

    _aplicar_anchos(tabla, anchos)


# ---------------------------------------------------------------------------
# Esquema mensual (datos generales + selección del tema + cuadro ¿Qué sabemos?)
# ---------------------------------------------------------------------------
TAM_GENERALES = Pt(13)
TAM_ENCABEZADO_MENSUAL = Pt(14)
RAYA = "_________"  # lo que se muestra cuando un dato está en blanco

# Cada línea de los datos generales: lista de (etiqueta, clave del JSON)
LINEAS_GENERALES = (
    (("Centro Educativo", "centro"), ("Docente", "docente"),
     ("Regional", "regional"), ("Distrito", "distrito")),
    (("Grado", "grado"), ("Sección", "seccion"), ("Tanda", "tanda")),
    (("Fecha de inicio", "fecha_inicio"), ("Fecha de cierre", "fecha_cierre"),
     ("Tipo de planificación", "tipo")),
    (("Tema", "tema"), ("Nombre Lúdico", "nombre_ludico")),
)


def _altura_minima(fila, twips: int) -> None:
    tr_pr = fila._tr.get_or_add_trPr()
    alto = OxmlElement("w:trHeight")
    alto.set(qn("w:val"), str(twips))
    alto.set(qn("w:hRule"), "atLeast")
    tr_pr.append(alto)


def _sangria_tabla(tabla, anchos, sangria: int) -> None:
    """Como _layout_fijo, pero con sangría propia."""
    _layout_fijo(tabla, anchos)
    for el in tabla._tbl.tblPr.findall(qn("w:tblInd")):
        el.set(qn("w:w"), str(sangria))


def _linea_general(documento, campos: list[tuple[str, str]], primera: bool,
                   nueva_pagina: bool) -> None:
    p = documento.add_paragraph()
    p.paragraph_format.left_indent = Twips(680)
    p.paragraph_format.right_indent = Twips(680)
    p.paragraph_format.space_after = Pt(4)
    if primera:
        p.paragraph_format.page_break_before = nueva_pagina
    for i, (etiqueta, valor) in enumerate(campos):
        if i:
            _formato_run(p.add_run("    "), tamano=TAM_GENERALES)
        _formato_run(p.add_run(f"{etiqueta}: "), negrita=True, cursiva=True,
                     tamano=TAM_GENERALES)
        _formato_run(p.add_run(valor or RAYA), cursiva=True, subrayado=True,
                     tamano=TAM_GENERALES)


def hoja_esquema(documento, mensual: dict, nueva_pagina=False) -> None:
    """Datos generales y el recuadro 'Planificación mensual'."""
    generales = mensual.get("generales", {})
    for n, linea in enumerate(LINEAS_GENERALES):
        campos = [(etiqueta, generales.get(clave, "")) for etiqueta, clave in linea]
        _linea_general(documento, campos, primera=(n == 0), nueva_pagina=nueva_pagina)

    documento.add_paragraph()
    anchos = (14734,)
    tabla = documento.add_table(rows=0, cols=1)
    _bordes_tabla(tabla)
    _sangria_tabla(tabla, anchos, 137)

    encabezado = tabla.add_row()
    _sombrear(encabezado.cells[0], COLOR_ENCABEZADO)
    _escribir(encabezado.cells[0], "Planificación mensual", negrita=True,
              cursiva=True, tamano=TAM_ENCABEZADO_MENSUAL, espacio_despues=0)

    cuerpo = tabla.add_row()
    _escribir(cuerpo.cells[0], mensual.get("seleccion_tema", ""), negrita=True,
              cursiva=True, justificado=True)
    _altura_minima(cuerpo, 2369)
    _aplicar_anchos(tabla, anchos)


def hoja_kwl(documento, mensual: dict, nueva_pagina=True) -> None:
    """Cuadro ¿Qué sabemos? / ¿Qué queremos saber? / ¿Cómo lo vamos a saber?"""
    p = documento.add_paragraph()
    p.paragraph_format.page_break_before = nueva_pagina
    p.paragraph_format.space_after = Pt(0)

    anchos = (4111, 5103, 5528)
    tabla = documento.add_table(rows=0, cols=3)
    _bordes_tabla(tabla)
    _sangria_tabla(tabla, anchos, 137)

    fila = tabla.add_row()
    titulos = ("¿Qué sabemos sobre el tema?", "¿Qué queremos saber?",
               "¿Cómo lo vamos a saber?")
    for celda, titulo in zip(fila.cells, titulos):
        _sombrear(celda, COLOR_ENCABEZADO)
        _escribir(celda, titulo, negrita=True, cursiva=True, centrado=True,
                  tamano=TAM_ENCABEZADO_MENSUAL, espacio_despues=0)

    cuerpo = tabla.add_row()
    columnas = (("que_sabemos", False), ("que_queremos_saber", True),
                ("como_lo_vamos_a_saber", False))
    for celda, (clave, negrita) in zip(cuerpo.cells, columnas):
        _escribir(celda, mensual.get(clave, []), negrita=negrita, cursiva=True,
                  prefijo="- ", espacio_despues=14)
    _altura_minima(cuerpo, 5000)
    _aplicar_anchos(tabla, anchos)


# ---------------------------------------------------------------------------
# API pública
# ---------------------------------------------------------------------------
def _documento_base() -> Document:
    doc = Document()
    seccion = doc.sections[0]
    seccion.orientation = WD_ORIENT.LANDSCAPE
    seccion.page_width = Twips(ANCHO_PAGINA)
    seccion.page_height = Twips(ALTO_PAGINA)
    seccion.left_margin = seccion.right_margin = Twips(MARGEN_LATERAL)
    seccion.top_margin = Twips(MARGEN_SUPERIOR)
    seccion.bottom_margin = Twips(MARGEN_INFERIOR)

    normal = doc.styles["Normal"]
    normal.font.name = FUENTE
    normal.font.size = TAMANO
    normal.element.rPr.rFonts.set(qn("w:eastAsia"), FUENTE)
    normal.paragraph_format.space_after = Pt(0)
    return doc


def generar_documento(datos: dict) -> Document:
    """Construye el documento Word a partir de los datos."""
    doc = _documento_base()
    dias = datos["dias"]
    mensual = datos.get("mensual")
    if mensual:
        hoja_esquema(doc, mensual)
        hoja_kwl(doc, mensual)
    for i, dia in enumerate(dias):
        # la primera hoja grupal sale en página nueva solo si hay esquema antes
        hoja_grupal(doc, dia, nueva_pagina=(i > 0 or bool(mensual)))
        hoja_indicador(doc, dia)
    return doc


def generar_archivo(ruta_json: str | Path, ruta_docx: str | Path) -> Path:
    """Lee un JSON de planificación y escribe el .docx."""
    datos = json.loads(Path(ruta_json).read_text(encoding="utf-8"))
    ruta_docx = Path(ruta_docx)
    ruta_docx.parent.mkdir(parents=True, exist_ok=True)
    generar_documento(datos).save(ruta_docx)
    return ruta_docx
