# Procesador de tablas · Laboratorio Control de Calidad

Página web para arrastrar un archivo **CSV, HTML o Excel**, extraer sus tablas y descargar un **PDF listo para imprimir** con el formato del laboratorio (mismo diseño que el registro de análisis en Excel).

Todo se procesa en el navegador: los archivos **no se suben a ningún servidor**.

## Uso

1. Abre la página y arrastra uno o varios archivos a la zona verde (o haz clic para elegirlos).
2. Revisa la vista previa. Puedes editar el título, la fecha, la nota del pie y los nombres de firma.
3. Pulsa **Descargar PDF**.

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
