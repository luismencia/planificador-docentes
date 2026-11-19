import json
from pathlib import Path

from planificador.generador import generar_documento
from planificador.importador import importar_documento

EJEMPLO = Path(__file__).parent.parent / "datos" / "ejemplo.json"


def test_ida_y_vuelta(tmp_path):
    """datos -> Word -> datos debe devolver lo mismo que entró."""
    original = json.loads(EJEMPLO.read_text(encoding="utf-8"))
    ruta = tmp_path / "ciclo.docx"
    generar_documento(original).save(ruta)

    recuperado = importar_documento(ruta)

    assert len(recuperado["dias"]) == len(original["dias"])
    for antes, despues in zip(original["dias"], recuperado["dias"]):
        assert despues["grupal"] == antes["grupal"]
        assert despues["hoja2"] == antes["hoja2"]


def test_importa_contenidos_por_categoria(tmp_path):
    original = json.loads(EJEMPLO.read_text(encoding="utf-8"))
    ruta = tmp_path / "c.docx"
    generar_documento(original).save(ruta)
    ultimo = importar_documento(ruta)["dias"][-1]["hoja2"]
    assert ultimo["contenidos"]["conceptuales"] == ["La Conversación."]
    assert "actitudinales" in ultimo["contenidos"]
