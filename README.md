# CAP.11.METABOLISMO-DE-CARBOHIDRATOS

## Presentación completa del resumen

- [Descargar presentación editable de PowerPoint](entregables/Presentacion_completa_Metabolismo_de_carbohidratos.pptx)
- [Abrir la presentación en PDF](entregables/Presentacion_completa_Metabolismo_de_carbohidratos.pdf)
- [Verificación de integridad y mapa de secciones](entregables/Verificacion_presentacion.json)

La presentación conserva **todos los párrafos y todas las celdas del Word**, sin resumirlos, abreviarlos ni simplificarlos. Es extensa porque la información se distribuye para mantener texto legible de **18 puntos como mínimo**: cuerpo principal de 22, cuadros de 20 y pies de 18 puntos. Los 15 temas del resumen tienen divisores e índice navegable.

Se conservan las **19 imágenes originales**, con **28 vistas ampliadas** de apoyo. Los **10 esquemas de la glucólisis** se redibujan como texto y enlaces vectoriales editables con letras de 18 puntos o mayores. Los originales rasterizados del libro conservan la nitidez de la fuente comprimida; no se afirma que pueda recuperarse información perdida por compresión.

El generador compara el texto visible del PowerPoint con las **1 200 unidades** de texto del Word y, adicionalmente, con todos sus nodos XML visibles. Verifica la cobertura completa de letras, cifras y símbolos, la conservación binaria de las imágenes originales y el tamaño de las fuentes. La copia PDF contiene las mismas diapositivas e incluye marcadores de sección y enlaces de navegación.

Para regenerar ambos formatos:

```bash
.venv/bin/pip install -r requirements-presentacion.txt
.venv/bin/python scripts/generar_presentacion.py
```

Las métricas y los recortes temporales se mantienen en `.cache/presentacion/`; no se incluyen en Git. La información de formato del **Word original** se conserva como tal, sin confundirse con el formato de esta presentación.

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
