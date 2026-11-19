# Planificador para docentes

Aplicación para generar las planificaciones diarias de nivel inicial en Word,
con el diseño exacto que exige la escuela, a partir de datos estructurados.

## Estado del proyecto
- [x] Fase 1: generador de Word desde JSON
- [x] Fase 2: importar el Word actual a datos
- [ ] Fase 3: formulario web (celular)
- [ ] Fase 4: foto de la hoja -> formulario (IA con visión)
- [ ] Fase 5: plantillas para otros formatos de planificación
- [ ] Fase 6: despliegue

## Cómo usarlo
```bash
pip install -r requirements.txt
python generar.py datos/ejemplo.json salida/planificacion.docx   # datos -> Word
python importar.py mi_documento.docx datos/importado.json        # Word -> datos
pytest
```

## Estructura
- `planificador/generador.py`: construye el .docx (python-docx)
- `planificador/importador.py`: lee un .docx existente y lo convierte a datos
- `datos/ejemplo.json`: ejemplo de tres días de clase
- `tests/`: pruebas automáticas
