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
    assert len(doc.tables) == 2 * len(datos["dias"])


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
