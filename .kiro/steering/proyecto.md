# E500 App — Estado del Proyecto

## Stack
- **Electron** (frame:false)
- **better-sqlite3 v9.4.3**
- **XLSX** para importación de datos
- Node.js v16+

## Arquitectura
- `main.js` — proceso principal, IPC handlers, SQLite
- `preload.js` — bridge contextBridge
- `index.html` — UI completa (HTML + CSS + JS vanilla)
- `db.js` — schema SQLite **compartido con monsam-app**
- `config.json` — configuración local

## Base de datos SQLite
- Ruta: `F:\DISEÑOS\Modelos 3D\control\datos.db` (COMPARTIDA con monsam-app)
- Campo en config.json: `rutaDB`
- Tablas principales: `filamentos`, `impresiones`
- `filamentos` es **propiedad de e500-app** — fuente de verdad de stock
- Colores degradados guardados como `hex1|hex2|hex3` en campo `color_hex`

## Funcionalidades principales
- Registro de impresiones con filamentos, tiempos, costos
- Gestión de stock de filamentos con alertas de nivel bajo
- Dashboard con totales y estadísticas
- Monitor de impresora via Moonraker

## Vista de Filamentos (Junio 2026)
- ✅ Grid de tarjetas en lugar de tabla
- ✅ Cada tarjeta muestra: swatch de color, nombre, tipo, costo/kg, barra de stock, porcentaje
- ✅ Soporte colores degradados bi/tricolor en swatch y modal
- ✅ Botón "🏷 Agrupar por marca" (toggle) — por defecto sin agrupar
- ✅ Sub-tabs por tipo: Todos / PLA / PLA+ / PETG / Silk

## Colores degradados
- Modal de filamento tiene checkbox "¿Es color degradado?"
- Si activo: campos Color 2 y Color 3 (tricolor opcional) con preview
- Se guarda como `hex1|hex2|hex3` en SQLite
- Al leer, se parsea automáticamente en `get-filamentos`
- Como la DB es compartida, los degradados se ven también en monsam-app

## config.json
```json
{
  "rutaExcel": "F:\\DISEÑOS\\Modelos 3D\\control\\e500_datos.xlsx",
  "rutaExcelMonsan": "F:\\DISEÑOS\\Modelos 3D\\aretes\\Catalogo_Aretes_3D.xlsx",
  "rutaDB": "F:\\DISEÑOS\\Modelos 3D\\control\\datos.db",
  "costoKwh": 2.8,
  "wattsPrinter": 350,
  "moonrakerUrl": "http://192.168.68.113"
}
```

## Ramas Git
- Repo: `fdtn91/e500-app`
- Rama local: puede variar (`main`, `fix-no-terminal`, etc.)
- Para actualizar: `git pull origin main`
