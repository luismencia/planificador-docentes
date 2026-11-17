# Git desde cero (lo mínimo para trabajar)

Git guarda "fotos" de tu proyecto (commits). Puedes volver a cualquiera,
ver qué cambiaste y trabajar sin miedo a romper algo.

## Una sola vez
```bash
git config --global user.name "Tu Nombre"
git config --global user.email "tu@correo.com"
```

## Empezar en este proyecto
```bash
cd planificador
git init
git add .
git commit -m "Fase 1: generador de Word desde JSON"
```

## El ciclo diario (el 90% de lo que usarás)
```bash
git status                 # ¿qué cambió?
git add .                  # preparo los cambios para guardar
git commit -m "mensaje"    # guardo la foto, con una descripción corta
git log --oneline          # historial de commits
```
**Regla de oro:** un commit = un cambio pequeño con un mensaje claro
("Agrega columna de contenidos", no "cambios").

## Deshacer
```bash
git diff                       # ver lo que cambié y aún no guardé
git restore archivo.py         # descartar cambios de un archivo
```

## Subirlo a GitHub (tu portafolio)
1. Crea un repositorio vacío en github.com (sin README).
2. Conéctalo y súbelo:
```bash
git remote add origin https://github.com/TU_USUARIO/planificador.git
git branch -M main
git push -u origin main
```
Después, cada vez que hagas commits: `git push`.

## Ramas (cuando quieras probar algo arriesgado)
```bash
git switch -c prueba-formulario   # crea y entra a una rama
# ...trabajas y haces commits...
git switch main                   # vuelves a la principal
git merge prueba-formulario       # unes lo bueno
```
