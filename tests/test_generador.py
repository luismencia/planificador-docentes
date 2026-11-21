import json
from pathlib import Path

from docx import Document

from planificador.generador import generar_documento

EJEMPLO = Path(__file__).parent.parent / "datos" / "ejemplo.json"


def _datos():
    return json.loads(EJEMPLO.read_text(encoding="utf-8"))


def test_cada_dia_genera_dos_tablas():
    datos = _datos()
    doc = generar_documento(datos)
    # 2 tablas por día + 2 del esquema mensual (recuadro y cuadro ¿Qué sabemos?)
    assert len(doc.tables) == 2 * len(datos["dias"]) + 2


def test_hoja_con_contenidos_tiene_cuatro_columnas():
    doc = generar_documento(_datos())
    # el día 4 (último) lleva contenidos -> su 2.ª tabla tiene 4 columnas
    assert len(doc.tables[-1].columns) == 4


def test_el_texto_aparece_en_el_documento(tmp_path):
    doc = generar_documento(_datos())
    ruta = tmp_path / "x.docx"
    doc.save(ruta)
    texto = "\n".join(c.text for t in Document(ruta).tables for f in t.rows for c in f.cells)
    assert "Cumplen con sus deberes" in texto
    assert "Cinta pegante" in texto


def test_esquema_mensual_aparece_primero(tmp_path):
    doc = generar_documento(_datos())
    ruta = tmp_path / "m.docx"
    doc.save(ruta)
    texto = "\n".join(p.text for p in Document(ruta).paragraphs)
    assert "Centro Educativo: Centro Educativo Ejemplo" in texto
    assert "Tanda: _________" in texto  # dato en blanco -> raya
    primera_tabla = Document(ruta).tables[0]
    assert primera_tabla.rows[0].cells[0].text == "Planificación mensual"


def test_sin_mensual_no_agrega_hojas():
    datos = _datos()
    datos.pop("mensual")
    assert len(generar_documento(datos).tables) == 2 * len(datos["dias"])
