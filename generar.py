"""Uso:  python generar.py datos/ejemplo.json salida/planificacion.docx"""
import sys
from planificador.generador import generar_archivo

if __name__ == "__main__":
    if len(sys.argv) != 3:
        sys.exit("Uso: python generar.py <entrada.json> <salida.docx>")
    print("Creado:", generar_archivo(sys.argv[1], sys.argv[2]))
