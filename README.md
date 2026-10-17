# Procesador de tablas · Laboratorio Control de Calidad

Página web para arrastrar un archivo **CSV, HTML o Excel**, extraer sus tablas y descargar un **PDF listo para imprimir** con el formato del laboratorio (mismo diseño que el registro de análisis en Excel).

Todo se procesa en el navegador: los archivos **no se suben a ningún servidor**.

## Uso

1. Abre la página y arrastra uno o varios archivos a la zona verde (o haz clic para elegirlos).
2. Revisa la vista previa. Puedes editar el título, el folio y el nombre de la firma.
3. El tipo de reporte se elige solo, según el archivo:
   - **Varias muestras → reporte general**: la tabla completa. Con las casillas de la columna *Muestra* eliges qué muestras salen en el PDF.
   - **Una sola muestra → reporte por etiqueta (vertical)**: llenas la descripción (nombre del vino desde un menú, cosecha, variedad, fechas de embotellado y de análisis con calendario —la de análisis se prellena con la fecha del nombre del archivo, p. ej. `12,10,26` → 12/10/2026—, No. de muestra; el tipo siempre es *Vino*); el *Lote* se llena solo con el ID de la muestra y no se puede editar (sin la basura final, p. ej. `26113,,,` → `26113`). La tabla queda como PARÁMETRO · VALOR · UNIDAD · OBSERVACIONES; solo las observaciones se pueden escribir (la metodología va en el pie). Abajo va la firma de *Aprobado por*. El PDF sale en hoja vertical.
4. Pulsa **Descargar PDF**.

Las etiquetas del menú están en `PRODUCTOS` (`js/app.js`); las que se agregan con "Otro…" se recuerdan en ese navegador. Los métodos del BioSystems Y15 (unidad y nota del pie) están en `METODOS`: cada parámetro con datos lleva un número volado y el pie explica solo los métodos que aparecen en el reporte.

| Formato | Qué se extrae |
|---|---|
| `.html` / `.htm` | Cada `<table>`. El reporte "Tabla resumen de resultados" del analizador se reconstruye en una sola tabla (une las páginas y los encabezados escalonados). |
| `.xlsx` / `.xls` / `.ods` | Una tabla por hoja. |
| `.csv` / `.tsv` / `.txt` | Una tabla. |

Funciona mejor con tablas "limpias" (una fila de encabezados y después los datos). Las celdas vacías o `-x-` se muestran como "—".

## Estructura

```
index.html      Página (carga + vista previa)
css/            Estilos (colores y tipografías del Excel del laboratorio)
js/app.js       Lectura de archivos y vista previa
js/pdf.js       Generación del PDF (encabezado, tablas, firmas, pie)
img/            Logo
vendor/         Librerías incluidas en el repo: SheetJS 0.20.3, jsPDF 2.5.2, jsPDF-AutoTable 3.8.4
```

## Probar en local

```
python -m http.server 8000
```

y abre http://localhost:8000

## Despliegue

Es un sitio estático, sin paso de compilación. En Vercel: importa el repositorio, framework **Other**, sin comando de build, directorio de salida `.` (raíz). Cada push a la rama principal se despliega solo.
