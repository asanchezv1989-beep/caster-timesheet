# Caster Timesheet (PWA)

Abre el `CASTER_OUT.xlsx` (hoja Timesheet) en el navegador, captura horas por día o por semana,
recalcula REG/OT/DT/PD y costos con las mismas reglas de las fórmulas del Excel, y descarga el
mismo archivo con solo las celdas cambiadas (formato y fórmulas intactos; Excel recalcula al abrir).
Todo se guarda en el dispositivo (IndexedDB). No hay servidor ni datos dentro del código.

- Probar local: `python -m http.server 8781` (no funciona con file://).
- Verificar el motor contra un Excel real: `node test/verify.mjs "ruta/CASTER_OUT.xlsx"` (debe dar 0 mismatches).
- Probar exportación: `node test/export.mjs "ruta/CASTER_OUT.xlsx"` → escribe test/out.xlsx (borrarlo después).
- Al cambiar archivos, subir la versión `CACHE` en `sw.js`.
