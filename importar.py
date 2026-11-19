"""Uso:  python importar.py mi_documento.docx datos/importado.json"""
import sys
from planificador.importador import importar_archivo

if __name__ == "__main__":
    if len(sys.argv) != 3:
        sys.exit("Uso: python importar.py <entrada.docx> <salida.json>")
    print("Creado:", importar_archivo(sys.argv[1], sys.argv[2]))
