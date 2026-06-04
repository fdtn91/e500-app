const { app, BrowserWindow, ipcMain, dialog, shell } = require('electron')
const path = require('path')
const fs   = require('fs')
const XLSX = require('xlsx')

// ════════════════════════════════════════════════════════════
//  CONFIG
// ════════════════════════════════════════════════════════════
const CONFIG_PATH  = path.join(__dirname, 'config.json')
const STOCK_MINIMO = 250  // gramos — umbral de advertencia

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
//  Hoja "Filamentos":
//  Nombre | Marca | Tipo | CostoKg | PesoBobina | StockGr |
//  CostoTotal | FechaCompra | Color | Notas
// ════════════════════════════════════════════════════════════
const FIL_HDR = ['Nombre','Marca','Tipo','CostoKg','PesoBobina',
                 'StockGr','CostoTotal','FechaCompra','Color','Notas']

ipcMain.handle('get-filamentos', (_, filePath) => {
  const wb = readWB(filePath)
  if (!wb) return []
  ensureSheet(wb, 'Filamentos', FIL_HDR)
  const rows = toRows(wb.Sheets['Filamentos'])
  return rows.slice(1).filter(r => r[0]).map((r, i) => ({
    _idx:        i,
    nombre:      r[0],
    marca:       r[1],
    tipo:        r[2],
    costoKg:     Number(r[3]) || 0,
    pesoBobina:  Number(r[4]) || 1000,
    stockGr:     Number(r[5]) || 0,
    costoTotal:  Number(r[6]) || 0,
    fechaCompra: r[7] || '',
    color:       r[8] || '#888888',
    notas:       r[9] || '',
    stockBajo:   (Number(r[5]) || 0) < STOCK_MINIMO
  }))
})

ipcMain.handle('save-filamento', (_, filePath, fil) => {
  let wb = readWB(filePath)
  if (!wb) { wb = XLSX.utils.book_new() }
  ensureSheet(wb, 'Filamentos', FIL_HDR)
  const ws   = wb.Sheets['Filamentos']
  const rows = toRows(ws)

  // Calcular costoKg automáticamente si viene costoTotal y pesoBobina
  let costoKg = fil.costoKg || 0
  if (!costoKg && fil.costoTotal && fil.pesoBobina) {
    costoKg = +((fil.costoTotal / fil.pesoBobina) * 1000).toFixed(2)
  }

  const row = [
    fil.nombre, fil.marca || '', fil.tipo || 'PLA',
    costoKg, fil.pesoBobina || 1000,
    fil.stockGr !== undefined ? fil.stockGr : (fil.pesoBobina || 1000),
    fil.costoTotal || 0, fil.fechaCompra || '',
    fil.color || '#888888', fil.notas || ''
  ]

  const editIdx = fil._editIndex !== undefined ? fil._editIndex + 1 : -1
  const idx = editIdx > 0 ? editIdx : rows.findIndex((r, i) => i > 0 && r[0] === fil.nombre)
  if (idx > 0) rows[idx] = row
  else rows.push(row)

  wb.Sheets['Filamentos'] = toSheet(rows)
  const ok = saveWB(wb, filePath)

  // Sincronizar automáticamente con MONSAN al guardar filamento
  const cfg = loadConfig()
  if (ok && cfg.rutaExcelMonsan) {
    sincronizarColorMonsan(cfg.rutaExcelMonsan, fil.nombre, row[5], costoKg)
  }

  return ok
})

ipcMain.handle('delete-filamento', (_, filePath, nombre) => {
  const wb = readWB(filePath)
  if (!wb) return false
  ensureSheet(wb, 'Filamentos', FIL_HDR)
  const rows = toRows(wb.Sheets['Filamentos']).filter((r, i) => i === 0 || r[0] !== nombre)
  wb.Sheets['Filamentos'] = toSheet(rows)
  return saveWB(wb, filePath)
})

ipcMain.handle('get-stock-minimo', () => STOCK_MINIMO)

// ════════════════════════════════════════════════════════════
//  SINCRONIZACIÓN AUTOMÁTICA CON MONSAN
//  Se ejecuta cada vez que se modifica el stock de un filamento
// ════════════════════════════════════════════════════════════
function sincronizarColorMonsan (rutaMonsan, nombreFilamento, nuevoStock, costoKg) {
  if (!rutaMonsan || !fs.existsSync(rutaMonsan)) return
  try {
    const wbM = readWB(rutaMonsan)
    if (!wbM || !wbM.SheetNames.includes('Colores')) return

    const colRows = toRows(wbM.Sheets['Colores'])
    const nombreBuscar = nombreFilamento.toLowerCase()

    for (let i = 1; i < colRows.length; i++) {
      const nombreCol = String(colRows[i][0] || '').toLowerCase()
      if (nombreCol.includes(nombreBuscar) || nombreBuscar.includes(nombreCol)) {
        colRows[i][4] = nuevoStock   // StockGr
        if (costoKg) colRows[i][5] = costoKg  // CostoPorKg
        break
      }
    }

    wbM.Sheets['Colores'] = toSheet(colRows)
    saveWB(wbM, rutaMonsan)
  } catch (e) {
    console.error('Error sync MONSAN:', e.message)
  }
}

// ════════════════════════════════════════════════════════════
//  IMPRESIONES
//  Hoja "Impresiones": Fecha | Descripcion | Filamento |
//    GramosUsados | TiempoImpresion | Categoria | Resultado | CostoMaterial
// ════════════════════════════════════════════════════════════
const IMP_HDR = ['Fecha','Descripcion','Filamento','GramosUsados',
                 'TiempoImpresion','Categoria','Resultado','CostoMaterial']

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

  // Calcular costo material
  let costoMat = imp.costoMaterial || 0
  let costoKgFil = 0
  if (imp.gramosUsados) {
    const filRows = toRows(wb.Sheets['Filamentos'])
    const filRow  = filRows.find((r, i) => i > 0 && r[0] === imp.filamento)
    if (filRow) {
      costoKgFil = Number(filRow[3]) || 0
      if (!costoMat) costoMat = +((imp.gramosUsados / 1000) * costoKgFil).toFixed(2)
    }
  }

  const row = [
    imp.fecha, imp.descripcion, imp.filamento,
    imp.gramosUsados, imp.tiempo || '',
    imp.categoria || 'General', imp.resultado || '', costoMat
  ]

  const ws   = wb.Sheets['Impresiones']
  const rows = toRows(ws)
  const editIdx = imp._editIndex !== undefined ? imp._editIndex + 1 : -1

  // Si es edición, recalcular diferencia de gramos para el stock
  let diferencia = imp.gramosUsados
  if (editIdx > 0 && editIdx < rows.length) {
    const gramosAnteriores = Number(rows[editIdx][3]) || 0
    diferencia = imp.gramosUsados - gramosAnteriores
    rows[editIdx] = row
  } else {
    rows.push(row)
  }

  wb.Sheets['Impresiones'] = toSheet(rows)

  // Descontar stock del filamento
  let nuevoStock = null
  if (imp.filamento && diferencia !== 0) {
    const filWs   = wb.Sheets['Filamentos']
    const filRows = toRows(filWs)
    const fi = filRows.findIndex((r, i) => i > 0 && r[0] === imp.filamento)
    if (fi > 0) {
      const actual = Number(filRows[fi][5]) || 0
      nuevoStock = +Math.max(0, actual - diferencia).toFixed(2)
      filRows[fi][5] = nuevoStock
      wb.Sheets['Filamentos'] = toSheet(filRows)
    }
  }

  const ok = saveWB(wb, filePath)

  // Sincronizar con MONSAN automáticamente
  if (ok && nuevoStock !== null) {
    const cfg = loadConfig()
    if (cfg.rutaExcelMonsan) {
      sincronizarColorMonsan(cfg.rutaExcelMonsan, imp.filamento, nuevoStock, costoKgFil)
    }
  }

  // Devolver advertencia si el stock quedó bajo
  const stockFinal = nuevoStock !== null ? nuevoStock : null
  return {
    ok,
    stockBajo: stockFinal !== null && stockFinal < STOCK_MINIMO,
    stockRestante: stockFinal,
    filamento: imp.filamento
  }
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
//  SYNC MANUAL COMPLETO CON MONSAN
// ════════════════════════════════════════════════════════════
ipcMain.handle('sync-stock-monsan', (_, rutaE500, rutaMonsan) => {
  const wbE = readWB(rutaE500)
  const wbM = readWB(rutaMonsan)
  if (!wbE || !wbM) return { ok: false, msg: 'No se encontraron los archivos Excel' }

  ensureSheet(wbE, 'Filamentos', FIL_HDR)
  if (!wbM.SheetNames.includes('Colores')) return { ok: false, msg: 'MONSAN no tiene hoja Colores' }

  const filRows = toRows(wbE.Sheets['Filamentos'])
  const colRows = toRows(wbM.Sheets['Colores'])
  let actualizados = 0

  for (let mi = 1; mi < colRows.length; mi++) {
    const nombreColor = String(colRows[mi][0] || '').toLowerCase().trim()
    if (!nombreColor) continue
    const fi = filRows.findIndex((r, i) => {
      if (i === 0) return false
      const n = String(r[0] || '').toLowerCase()
      return n.includes(nombreColor) || nombreColor.includes(n)
    })
    if (fi > 0) {
      const stockE500   = Number(filRows[fi][5]) || 0
      const stockMonsan = Number(colRows[mi][4]) || 0
      const stockReal   = Math.min(stockE500, stockMonsan)
      filRows[fi][5] = stockReal
      colRows[mi][4] = stockReal
      if (filRows[fi][3]) colRows[mi][5] = Number(filRows[fi][3])
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
    msg: `${actualizados} filamentos sincronizados`
  }
})

// ════════════════════════════════════════════════════════════
//  ALERTAS DE STOCK BAJO AL ARRANCAR
// ════════════════════════════════════════════════════════════
ipcMain.handle('get-alertas-stock', (_, filePath) => {
  const wb = readWB(filePath)
  if (!wb) return []
  ensureSheet(wb, 'Filamentos', FIL_HDR)
  const rows = toRows(wb.Sheets['Filamentos'])
  return rows.slice(1).filter(r => r[0] && (Number(r[5]) || 0) < STOCK_MINIMO).map(r => ({
    nombre:  r[0],
    stockGr: Number(r[5]) || 0,
    color:   r[8] || '#888888'
  }))
})
