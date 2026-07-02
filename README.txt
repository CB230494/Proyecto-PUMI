ARCHIVOS PARA EL DASHBOARD NACIONAL PUMI 2026

Debe crear/subir estos archivos en el mismo nivel del repositorio:

1. index.html
2. style.css
3. app.js

Además, debe tener en el mismo nivel los Excel regionales:

DR1 C AVANCE JUNIO.xlsx
DR1 N AVANCE JUNIO.xlsx
DR1 S AVANCE JUNIO.xlsx
DR2 AVANCE JUNIO.xlsx
DR3 AVANCE JUNIO.xlsx
DR4 AVANCE JUNIO.xlsx
DR5 AVANCE JUNIO.xlsx
DR6 AVANCE JUNIO.xlsx
DR7 AVANCE JUNIO.xlsx
DR8 AVANCE JUNIO.xlsx
DR9 AVANCE JUNIO.xlsx
DR10 AVANCE JUNIO.xlsx
DR11 AVANCE JUNIO.xlsx
DR12 AVANCE JUNIO.xlsx

La app funciona como web estática en GitHub Pages.
No requiere Streamlit ni backend.

Librerías usadas por CDN:
- SheetJS XLSX para leer Excel.
- Chart.js para gráficos.
- jsPDF y jsPDF-AutoTable para informes PDF.

IMPORTANTE:
Los nombres de los Excel deben coincidir con la lista definida al inicio de app.js.
Si algún archivo tiene otro nombre, edite la constante ARCHIVOS_EXCEL.
