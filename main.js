const { app, BrowserWindow, ipcMain, dialog, shell } = require('electron')
const path = require('path')
const fs   = require('fs')
const XLSX = require('xlsx')

// ════════════════════════════════════════════════════════════
//  CONFIG
// ════════════════════════════════════════════════════════════
const CONFIG_PATH = path.join(__dirname, 'config.json')

function loadConfig () {
  try { return JSON.parse(fs.readFileSync(CONFIG_PATH, 'utf8')) }
  catch { return {} }
}
function saveConfigFile (cfg) {
  fs.writeFileSync(CONFIG_PATH, JSON.stringify(cfg, null, 2), 'utf8')
}

ipcMain.handle('get-config', () => loadConfig())
ipcMain.handle('save-config', (_, cfg) => { saveConfigFile(cfg); return true })

// ════════════════════════════════════════════════════════════
//  VENTANA
// ════════════════════════════════════════════════════════════
function createWindow () {
  const win = new BrowserWindow({
    width: 1200, height: 760,
    minWidth: 960, minHeight: 640,
    frame: false,
    backgroundColor: '#080B10',
    webPreferences: {
      preload: path.join(__dirname, 'preload.js'),
      contextIsolation: true,
      nodeIntegration: false
    }
  })
  win.loadFile('index.html')
}

app.whenReady().then(createWindow)
app.on('window-all-closed', () => { if (process.platform !== 'darwin') app.quit() })

ipcMain.on('win-minimize', e => BrowserWindow.fromWebContents(e.sender).minimize())
ipcMain.on('win-maximize', e => {
  const w = BrowserWindow.fromWebContents(e.sender)
  w.isMaximized() ? w.unmaximize() : w.maximize()
})
ipcMain.on('win-close', e => BrowserWindow.fromWebContents(e.sender).close())

// ════════════════════════════════════════════════════════════
//  DIÁLOGOS
// ════════════════════════════════════════════════════════════
ipcMain.handle('select-excel', async (_, def) => {
  const r = await dialog.showOpenDialog({
    filters: [{ name: 'Excel', extensions: ['xlsx'] }],
    defaultPath: def || ''
  })
  return r.canceled ? null : r.filePaths[0]
})
ipcMain.handle('open-excel', async (_, p) => {
  if (fs.existsSync(p)) { await shell.openPath(p); return true }
  return false
})

// ════════════════════════════════════════════════════════════
//  EXCEL — helpers
// ════════════════════════════════════════════════════════════
function readWB (filePath) {
  if (!fs.existsSync(filePath)) return null
  try { return XLSX.readFile(filePath) } catch { return null }
}
function saveWB (wb, filePath) {
  try { XLSX.writeFile(wb, filePath); return true } catch { return false }
}
function ensureSheet (wb, name, headers) {
  if (!wb.SheetNames.includes(name)) {
    const ws = XLSX.utils.aoa_to_sheet([headers])
    XLSX.utils.book_append_sheet(wb, ws, name)
  }
  return wb.Sheets[name]
}
function toRows (ws)    { return XLSX.utils.sheet_to_json(ws, { header: 1, defval: '' }) }
function toSheet (rows) { return XLSX.utils.aoa_to_sheet(rows) }

// ════════════════════════════════════════════════════════════
//  FILAMENTOS
//  Hoja "Filamentos": Nombre | Marca | Tipo | CostoKg | TotalGr | StockGr
// ════════════════════════════════════════════════════════════
const FIL_HDR = ['Nombre', 'Marca', 'Tipo', 'CostoKg', 'TotalGr', 'StockGr']

ipcMain.handle('get-filamentos', (_, filePath) => {
  const wb = readWB(filePath)
  if (!wb) return []
  ensureSheet(wb, 'Filamentos', FIL_HDR)
  const rows = toRows(wb.Sheets['Filamentos'])
  return rows.slice(1).filter(r => r[0]).map((r, i) => ({
    _idx:    i,
    nombre:  r[0],
    marca:   r[1],
    tipo:    r[2],
    costoKg: Number(r[3]) || 0,
    totalGr: Number(r[4]) || 0,
    stockGr: Number(r[5]) || 0
  }))
})

ipcMain.handle('save-filamento', (_, filePath, fil) => {
  let wb = readWB(filePath)
  if (!wb) { wb = XLSX.utils.book_new() }
  ensureSheet(wb, 'Filamentos', FIL_HDR)
  const ws   = wb.Sheets['Filamentos']
  const rows = toRows(ws)
  const row  = [fil.nombre, fil.marca || '', fil.tipo || 'PLA',
                fil.costoKg || 0, fil.totalGr || 1000, fil.stockGr || 0]
  const editIdx = fil._editIndex !== undefined ? fil._editIndex + 1 : -1
  const idx = editIdx > 0 ? editIdx : rows.findIndex((r, i) => i > 0 && r[0] === fil.nombre)
  if (idx > 0) rows[idx] = row
  else rows.push(row)
  wb.Sheets['Filamentos'] = toSheet(rows)
  return saveWB(wb, filePath)
})

ipcMain.handle('delete-filamento', (_, filePath, nombre) => {
  const wb = readWB(filePath)
  if (!wb) return false
  ensureSheet(wb, 'Filamentos', FIL_HDR)
  const rows = toRows(wb.Sheets['Filamentos']).filter((r, i) => i === 0 || r[0] !== nombre)
  wb.Sheets['Filamentos'] = toSheet(rows)
  return saveWB(wb, filePath)
})

ipcMain.handle('descontar-filamento', (_, filePath, nombre, gramos) => {
  const wb = readWB(filePath)
  if (!wb) return { ok: false, msg: 'Excel no encontrado' }
  ensureSheet(wb, 'Filamentos', FIL_HDR)
  const ws   = wb.Sheets['Filamentos']
  const rows = toRows(ws)
  const idx  = rows.findIndex((r, i) => i > 0 && r[0] === nombre)
  if (idx < 0) return { ok: false, msg: `Filamento "${nombre}" no encontrado` }
  const actual = Number(rows[idx][5]) || 0
  if (actual < gramos) return { ok: false, msg: `Stock insuficiente (${actual}gr disponibles)` }
  rows[idx][5] = +( actual - gramos).toFixed(2)
  wb.Sheets['Filamentos'] = toSheet(rows)
  const ok = saveWB(wb, filePath)
  return { ok, stockRestante: rows[idx][5], msg: ok ? `Stock actualizado: ${rows[idx][5]}gr` : 'Error al guardar' }
})

// ════════════════════════════════════════════════════════════
//  IMPRESIONES
//  Hoja "Impresiones": Fecha | Descripcion | Filamento |
//    GramosUsados | TiempoImpresion | Categoria | Resultado | CostoMaterial
// ════════════════════════════════════════════════════════════
const IMP_HDR = ['Fecha', 'Descripcion', 'Filamento', 'GramosUsados',
                 'TiempoImpresion', 'Categoria', 'Resultado', 'CostoMaterial']

const CATEGORIAS = ['Personal', 'Encargo externo', 'Mantenimiento / Repuesto',
                    'Prototipo / Prueba', 'MONSAN Aretes', 'General']

ipcMain.handle('get-impresiones', (_, filePath) => {
  const wb = readWB(filePath)
  if (!wb) return []
  ensureSheet(wb, 'Impresiones', IMP_HDR)
  const rows = toRows(wb.Sheets['Impresiones'])
  return rows.slice(1).filter(r => r[0] || r[1] || r[2]).map((r, i) => ({
    _idx:          i,
    fecha:         r[0],
    descripcion:   r[1],
    filamento:     r[2],
    gramosUsados:  Number(r[3]) || 0,
    tiempo:        r[4],
    categoria:     r[5],
    resultado:     r[6],
    costoMaterial: Number(r[7]) || 0
  }))
})

ipcMain.handle('save-impresion', async (_, filePath, imp) => {
  let wb = readWB(filePath)
  if (!wb) { wb = XLSX.utils.book_new() }
  ensureSheet(wb, 'Impresiones', IMP_HDR)
  ensureSheet(wb, 'Filamentos',  FIL_HDR)

  // Calcular costo material si no viene
  let costoMat = imp.costoMaterial || 0
  if (!costoMat && imp.gramosUsados) {
    const filRows = toRows(wb.Sheets['Filamentos'])
    const filRow  = filRows.find((r, i) => i > 0 && r[0] === imp.filamento)
    if (filRow) costoMat = +((imp.gramosUsados / 1000) * (Number(filRow[3]) || 0)).toFixed(2)
  }

  const row = [
    imp.fecha, imp.descripcion, imp.filamento,
    imp.gramosUsados, imp.tiempo || '',
    imp.categoria || 'General', imp.resultado || '', costoMat
  ]

  const ws   = wb.Sheets['Impresiones']
  const rows = toRows(ws)
  const editIdx = imp._editIndex !== undefined ? imp._editIndex + 1 : -1
  if (editIdx > 0 && editIdx < rows.length) rows[editIdx] = row
  else rows.push(row)
  wb.Sheets['Impresiones'] = toSheet(rows)

  // Descontar stock del filamento automáticamente (solo impresiones nuevas)
  if (editIdx < 0 && imp.filamento && imp.gramosUsados) {
    const filWs   = wb.Sheets['Filamentos']
    const filRows = toRows(filWs)
    const fi = filRows.findIndex((r, i) => i > 0 && r[0] === imp.filamento)
    if (fi > 0) {
      const actual = Number(filRows[fi][5]) || 0
      filRows[fi][5] = +Math.max(0, actual - imp.gramosUsados).toFixed(2)
      wb.Sheets['Filamentos'] = toSheet(filRows)
    }
  }

  return saveWB(wb, filePath)
})

ipcMain.handle('delete-impresion', (_, filePath, rowIndex) => {
  const wb = readWB(filePath)
  if (!wb) return false
  ensureSheet(wb, 'Impresiones', IMP_HDR)
  const rows = toRows(wb.Sheets['Impresiones'])
  rows.splice(rowIndex + 1, 1)
  wb.Sheets['Impresiones'] = toSheet(rows)
  return saveWB(wb, filePath)
})

// ════════════════════════════════════════════════════════════
//  SYNC STOCK CON MONSAN
//  Sincroniza la hoja "Colores" de MONSAN con "Filamentos" de E500
//  usando el nombre del filamento como clave de vinculación
// ════════════════════════════════════════════════════════════
ipcMain.handle('sync-stock-monsan', (_, rutaE500, rutaMonsan) => {
  const wbE   = readWB(rutaE500)
  const wbM   = readWB(rutaMonsan)
  if (!wbE || !wbM) return { ok: false, msg: 'No se encontraron los archivos Excel' }

  ensureSheet(wbE, 'Filamentos', FIL_HDR)
  // Hoja Colores de MONSAN: Nombre | CodigoColor | Hex | Descripcion | StockGr | CostoPorKg
  if (!wbM.SheetNames.includes('Colores')) return { ok: false, msg: 'MONSAN no tiene hoja Colores' }

  const filRows = toRows(wbE.Sheets['Filamentos'])
  const colRows = toRows(wbM.Sheets['Colores'])

  let actualizados = 0

  // Recorrer colores de MONSAN y buscar match en filamentos E500
  for (let mi = 1; mi < colRows.length; mi++) {
    const nombreColor = String(colRows[mi][0] || '').trim()
    if (!nombreColor) continue

    // Buscar filamento en E500 por nombre similar
    const fi = filRows.findIndex((r, i) => {
      if (i === 0) return false
      const nombreFil = String(r[0] || '').toLowerCase()
      return nombreFil.includes(nombreColor.toLowerCase()) ||
             nombreColor.toLowerCase().includes(nombreFil.toLowerCase())
    })

    if (fi > 0) {
      // Sincronizar stock: usar el menor de los dos como stock real
      const stockE500  = Number(filRows[fi][5]) || 0
      const stockMonsan = Number(colRows[mi][4]) || 0
      const stockReal  = Math.min(stockE500, stockMonsan)

      filRows[fi][5]  = stockReal  // actualizar E500
      colRows[mi][4]  = stockReal  // actualizar MONSAN
      actualizados++
    }
  }

  wbE.Sheets['Filamentos'] = toSheet(filRows)
  wbM.Sheets['Colores']    = toSheet(colRows)

  const ok1 = saveWB(wbE, rutaE500)
  const ok2 = saveWB(wbM, rutaMonsan)

  return {
    ok: ok1 && ok2,
    actualizados,
    msg: `${actualizados} filamentos sincronizados entre E500 y MONSAN`
  }
})
