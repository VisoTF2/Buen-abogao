# Buen-abogao

Aplicación web de una sola página para organizar y anotar artículos legales en modo offline.

## ¿Está actualizado en GitHub?
No. Este repositorio local no publica automáticamente los cambios. Para que los commits aparezcan en GitHub debes ejecutar un `git push` al repositorio remoto y rama correspondientes.

## Cómo ver la app
1. Inicia un servidor estático en esta carpeta, por ejemplo: `python -m http.server 8000`.
2. Abre `http://localhost:8000/` en el navegador.
3. Usa **Ctrl+F** o **F3** para abrir el buscador interno, escribe el término y navega las coincidencias con **Enter**.

## Funcionalidades clave
- Agrupar artículos por normativa y materia, con colores y notas persistidas en `localStorage`.
- Panel de búsqueda flotante con resaltado de coincidencias en títulos, artículos y notas.
- Edición inline de contenido y notas, con reordenamiento automático de materias.
