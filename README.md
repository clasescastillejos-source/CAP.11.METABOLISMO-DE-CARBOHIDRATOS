# CAP.11.METABOLISMO-DE-CARBOHIDRATOS

## Documento de estudio

[Descargar el resumen detallado en Word](entregables/Resumen_detallado_Metabolismo_de_carbohidratos.docx)

El documento está elaborado a partir de `CAP.11.METABOLISMO DE CARBOHIDRATOS_compressed.pdf`, capítulo 11, páginas 243–270 del libro (28 páginas del PDF).

Incluye:

- Texto, títulos, tablas y pies de imagen en **14 puntos**.
- Desarrollo de las **diez reacciones de la glucólisis**, con cambios estructurales, enzimas, coenzimas, cofactores y balances.
- Las **16 figuras numeradas originales** y tres láminas adicionales del capítulo, con páginas de origen, ubicación en la guía y explicaciones.
- **Diez esquemas químicos propios** de alta resolución.
- Cuadros comparativos y sinópticos hechos como **tablas nativas editables de Word**.
- Digestión, otros monosacáridos, gluconeogénesis, Cori/alanina, piruvato, pentosas fosfato y glucógeno.
- Ejercicios resueltos, glosario y precisiones sobre posibles confusiones del capítulo.

Las figuras originales conservan las limitaciones del PDF comprimido; exportarlas a 300 ppp no recupera detalles perdidos. Los dibujos químicos nuevos se distinguen expresamente de las imágenes del libro.

## Regenerar el archivo

```bash
python -m venv .venv
.venv/bin/pip install -r requirements-documento.txt
.venv/bin/python scripts/generar_resumen.py
```

Se requiere la fuente DejaVu Sans de `/usr/share/fonts/truetype/dejavu/` para los esquemas químicos propios. El contenido editorial editable está en `scripts/resumen_contenido.md`. Los recortes y archivos temporales se producen en `.cache/`, que no se incluye en Git.

No se atribuyen autor, editorial o edición del libro sin verificarlos: el PDF facilitado contiene únicamente el capítulo.
