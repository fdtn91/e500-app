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
    show: false,                          // oculta hasta que esté listo
    backgroundColor: '#080B10',
    icon: path.join(__dirname, 'icono.ico'),
    webPreferences: {
      preload: path.join(__dirname, 'preload.js'),
      contextIsolation: true,
      nodeIntegration: false,
      webSecurity: false   // permite cargar recursos de red local (cámara, Moonraker)
    }
  })
  win.loadFile('index.html')
  // Mostrar solo cuando el contenido esté pintado — sin flash de fondo negro
  win.once('ready-to-show', () => win.show())
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
//  Hoja "Filamentos" — estructura REAL del Excel:
//  [0]Nombre | [1]Marca | [2]Tipo | [3]CostoKg | [4]TotalGr |
//  [5]StockGr | [6](vacía) | [7](vacía) | [8]ColorHex
// ════════════════════════════════════════════════════════════
const FIL_HDR = ['Nombre','Marca','Tipo','CostoKg','TotalGr','StockGr','','','ColorHex']

// Índices de columna — fuente única de verdad
const FC = {
  NOMBRE:    0,
  MARCA:     1,
  TIPO:      2,
  COSTO_KG:  3,
  TOTAL_GR:  4,   // peso total de la bobina
  STOCK_GR:  5,   // stock actual
  COLOR_HEX: 8
}

ipcMain.handle('get-filamentos', (_, filePath) => {
  const wb = readWB(filePath)
  if (!wb) return []
  const rows = toRows(wb.Sheets['Filamentos'])
  // Detectar fila de cabecera (puede ser fila 0 o 1)
  const dataStart = rows[0] && String(rows[0][0]).toLowerCase().includes('nombre') ? 1 : 0
  return rows.slice(dataStart).filter(r => r[FC.NOMBRE]).map((r, i) => ({
    _idx:        i,
    nombre:      r[FC.NOMBRE],
    colorNombre: r[FC.NOMBRE] || '',   // en tu Excel el nombre ya incluye el color
    marca:       r[FC.MARCA]     || '',
    tipo:        r[FC.TIPO]      || 'PLA',
    costoKg:     Number(r[FC.COSTO_KG])  || 0,
    pesoBobina:  Number(r[FC.TOTAL_GR])  || 1000,
    stockGr:     Number(r[FC.STOCK_GR])  || 0,
    costoTotal:  0,
    fechaCompra: '',
    colorHex:    r[FC.COLOR_HEX] || '#888888',
    notas:       '',
    stockBajo:   (Number(r[FC.STOCK_GR]) || 0) < STOCK_MINIMO
  }))
})

ipcMain.handle('save-filamento', (_, filePath, fil) => {
  let wb = readWB(filePath)
  if (!wb) { wb = XLSX.utils.book_new() }
  if (!wb.Sheets['Filamentos']) {
    const ws = XLSX.utils.aoa_to_sheet([FIL_HDR])
    XLSX.utils.book_append_sheet(wb, ws, 'Filamentos')
  }
  const ws   = wb.Sheets['Filamentos']
  const rows = toRows(ws)
  const dataStart = rows[0] && String(rows[0][0]).toLowerCase().includes('nombre') ? 1 : 0

  const nombre  = fil.nombre || fil.colorNombre || ''
  let costoKg   = fil.costoKg || 0
  if (!costoKg && fil.costoTotal && fil.pesoBobina)
    costoKg = +((fil.costoTotal / fil.pesoBobina) * 1000).toFixed(2)

  // Construir fila según estructura real del Excel
  const buildRow = (existing) => {
    const r = existing ? [...existing] : new Array(9).fill('')
    r[FC.NOMBRE]    = nombre
    r[FC.MARCA]     = fil.marca       || r[FC.MARCA]     || ''
    r[FC.TIPO]      = fil.tipo        || r[FC.TIPO]      || 'PLA'
    r[FC.COSTO_KG]  = costoKg         || r[FC.COSTO_KG]  || 0
    r[FC.TOTAL_GR]  = fil.pesoBobina  || r[FC.TOTAL_GR]  || 1000
    r[FC.STOCK_GR]  = fil.stockGr !== undefined ? fil.stockGr : (r[FC.STOCK_GR] || fil.pesoBobina || 1000)
    r[FC.COLOR_HEX] = fil.colorHex   || r[FC.COLOR_HEX] || '#888888'
    return r
  }

  const editIdx = fil._editIndex !== undefined ? fil._editIndex + dataStart : -1
  const existingIdx = rows.findIndex((r, i) => i >= dataStart && r[FC.NOMBRE] === nombre)

  if (editIdx > 0 && editIdx < rows.length)
    rows[editIdx] = buildRow(rows[editIdx])
  else if (existingIdx >= dataStart)
    rows[existingIdx] = buildRow(rows[existingIdx])
  else
    rows.push(buildRow(null))

  wb.Sheets['Filamentos'] = toSheet(rows)
  const ok = saveWB(wb, filePath)

  // Sincronizar con MONSAN
  const cfg = loadConfig()
  if (ok && cfg.rutaExcelMonsan) {
    const stockGr = fil.stockGr !== undefined ? fil.stockGr : (fil.pesoBobina || 1000)
    sincronizarColorMonsan(cfg.rutaExcelMonsan, nombre, stockGr, costoKg,
                           nombre, fil.colorHex)
  }
  return ok
})

ipcMain.handle('delete-filamento', (_, filePath, nombre) => {
  const wb = readWB(filePath)
  if (!wb) return false
  const rows = toRows(wb.Sheets['Filamentos']).filter((r, i) => {
    if (i === 0 && String(r[0]).toLowerCase().includes('nombre')) return true
    return r[FC.NOMBRE] !== nombre
  })
  wb.Sheets['Filamentos'] = toSheet(rows)
  return saveWB(wb, filePath)
})

ipcMain.handle('get-stock-minimo', () => STOCK_MINIMO)

// ════════════════════════════════════════════════════════════
//  SINCRONIZACIÓN AUTOMÁTICA CON MONSAN
//  Hoja Colores de MONSAN: [0]Nombre [1]CodigoColor [2]Hex [3]Descripcion [4]StockGr [5]CostoPorKg
// ════════════════════════════════════════════════════════════
function sincronizarColorMonsan (rutaMonsan, nombreFilamento, nuevoStock, costoKg, colorNombre, colorHex) {
  if (!rutaMonsan || !fs.existsSync(rutaMonsan)) return
  try {
    const wbM = readWB(rutaMonsan)
    if (!wbM || !wbM.SheetNames.includes('Colores')) return

    const colRows    = toRows(wbM.Sheets['Colores'])
    const buscar     = String(nombreFilamento || '').toLowerCase().trim()
    let   encontrado = false

    for (let i = 1; i < colRows.length; i++) {
      const nombreCol = String(colRows[i][0] || '').toLowerCase().trim()
      if (nombreCol === buscar || nombreCol.includes(buscar) || buscar.includes(nombreCol)) {
        colRows[i][4] = nuevoStock                        // StockGr
        if (costoKg)  colRows[i][5] = costoKg            // CostoPorKg
        if (colorHex) colRows[i][2] = colorHex           // Hex
        encontrado = true
        break
      }
    }

    // Si no existe el color en MONSAN, lo agrega automáticamente
    if (!encontrado) {
      const nombre  = nombreFilamento
      const base    = nombre.replace(/[^a-zA-Z]/g, '').substring(0, 3).toUpperCase()
      const existentes = colRows.slice(1).map(r => String(r[1] || ''))
      let n = 1
      while (existentes.includes(`${base}${n}`)) n++
      const codigo = `${base}${n}`
      colRows.push([nombre, codigo, colorHex || '#888888', '', nuevoStock, costoKg || 0])
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
    const dataStart = filRows[0] && String(filRows[0][0]).toLowerCase().includes('nombre') ? 1 : 0
    const filRow  = filRows.find((r, i) => i >= dataStart && r[FC.NOMBRE] === imp.filamento)
    if (filRow) {
      costoKgFil = Number(filRow[FC.COSTO_KG]) || 0
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
  // Estructura real: [5] = StockGr
  let nuevoStock = null
  if (imp.filamento && diferencia !== 0) {
    const filWs   = wb.Sheets['Filamentos']
    const filRows = toRows(filWs)
    const dataStart = filRows[0] && String(filRows[0][0]).toLowerCase().includes('nombre') ? 1 : 0
    const fi = filRows.findIndex((r, i) => i >= dataStart && r[FC.NOMBRE] === imp.filamento)
    if (fi >= dataStart) {
      const actual = Number(filRows[fi][FC.STOCK_GR]) || 0
      nuevoStock = +Math.max(0, actual - diferencia).toFixed(2)
      filRows[fi][FC.STOCK_GR] = nuevoStock
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
//  IMPORTAR TODOS LOS FILAMENTOS DE E500 A MONSAN COMO COLORES
//  Borra los colores existentes y los reemplaza con los filamentos de E500
// ════════════════════════════════════════════════════════════
ipcMain.handle('importar-filamentos-a-monsan', (_, rutaE500, rutaMonsan) => {
  const wbE = readWB(rutaE500)
  const wbM = readWB(rutaMonsan)
  if (!wbE) return { ok: false, msg: 'Excel E500 no encontrado' }
  if (!wbM) return { ok: false, msg: 'Excel MONSAN no encontrado' }

  const filRows   = toRows(wbE.Sheets['Filamentos'])
  const dataStart = filRows[0] && String(filRows[0][0]).toLowerCase().includes('nombre') ? 1 : 0

  const genCodigo = (nombre, existentes) => {
    const base = nombre.replace(/[^a-zA-Z]/g, '').substring(0, 3).toUpperCase()
    let n = 1
    while (existentes.has(`${base}${n}`)) n++
    return `${base}${n}`
  }

  const COL_HDR_M  = ['Nombre', 'CodigoColor', 'Hex', 'Descripcion', 'StockGr', 'CostoPorKg']
  const nuevasFilas = [COL_HDR_M]
  const existentes  = new Set()

  for (let i = dataStart; i < filRows.length; i++) {
    const r = filRows[i]
    if (!r[FC.NOMBRE]) continue
    const nombre = String(r[FC.NOMBRE])
    const codigo = genCodigo(nombre, existentes)
    existentes.add(codigo)
    nuevasFilas.push([
      nombre,
      codigo,
      r[FC.COLOR_HEX] || '#888888',
      '',
      Number(r[FC.STOCK_GR])  || 0,
      Number(r[FC.COSTO_KG])  || 0
    ])
  }

  // Reemplazar hoja Colores en MONSAN
  if (wbM.SheetNames.includes('Colores')) {
    const idx = wbM.SheetNames.indexOf('Colores')
    wbM.SheetNames.splice(idx, 1)
    delete wbM.Sheets['Colores']
  }
  wbM.SheetNames.push('Colores')
  wbM.Sheets['Colores'] = toSheet(nuevasFilas)

  const ok = saveWB(wbM, rutaMonsan)
  return {
    ok,
    importados: nuevasFilas.length - 1,
    msg: ok
      ? `${nuevasFilas.length - 1} filamentos importados como colores en MONSAN`
      : 'Error al guardar el Excel de MONSAN'
  }
})

// ════════════════════════════════════════════════════════════
//  SYNC MANUAL COMPLETO CON MONSAN
// ════════════════════════════════════════════════════════════
ipcMain.handle('sync-stock-monsan', (_, rutaE500, rutaMonsan) => {
  const wbE = readWB(rutaE500)
  const wbM = readWB(rutaMonsan)
  if (!wbE || !wbM) return { ok: false, msg: 'No se encontraron los archivos Excel' }
  if (!wbM.SheetNames.includes('Colores')) return { ok: false, msg: 'MONSAN no tiene hoja Colores' }

  const filRows  = toRows(wbE.Sheets['Filamentos'])
  const colRows  = toRows(wbM.Sheets['Colores'])
  const dataStart = filRows[0] && String(filRows[0][0]).toLowerCase().includes('nombre') ? 1 : 0
  let actualizados = 0

  for (let mi = 1; mi < colRows.length; mi++) {
    const nombreColor = String(colRows[mi][0] || '').toLowerCase().trim()
    if (!nombreColor) continue
    const fi = filRows.findIndex((r, i) => {
      if (i < dataStart) return false
      const n = String(r[FC.NOMBRE] || '').toLowerCase().trim()
      return n === nombreColor || n.includes(nombreColor) || nombreColor.includes(n)
    })
    if (fi >= dataStart) {
      // E500 es la fuente de verdad — MONSAN siempre recibe el stock de E500
      colRows[mi][4] = Number(filRows[fi][FC.STOCK_GR]) || 0   // StockGr
      if (filRows[fi][FC.COSTO_KG]) colRows[mi][5] = Number(filRows[fi][FC.COSTO_KG])  // CostoPorKg
      if (filRows[fi][FC.COLOR_HEX]) colRows[mi][2] = filRows[fi][FC.COLOR_HEX]         // Hex
      actualizados++
    }
  }

  wbM.Sheets['Colores'] = toSheet(colRows)
  const ok = saveWB(wbM, rutaMonsan)

  return {
    ok,
    actualizados,
    msg: ok ? `${actualizados} filamentos sincronizados` : 'Error al guardar Excel de MONSAN'
  }
})

// ════════════════════════════════════════════════════════════
//  ALERTAS DE STOCK BAJO AL ARRANCAR
// ════════════════════════════════════════════════════════════
ipcMain.handle('get-alertas-stock', (_, filePath) => {
  const wb = readWB(filePath)
  if (!wb) return []
  const rows = toRows(wb.Sheets['Filamentos'])
  const dataStart = rows[0] && String(rows[0][0]).toLowerCase().includes('nombre') ? 1 : 0
  return rows.slice(dataStart).filter(r => r[FC.NOMBRE] && (Number(r[FC.STOCK_GR]) || 0) < STOCK_MINIMO).map(r => ({
    nombre:      r[FC.NOMBRE],
    colorNombre: r[FC.NOMBRE] || '',
    stockGr:     Number(r[FC.STOCK_GR]) || 0,
    colorHex:    r[FC.COLOR_HEX] || '#888888'
  }))
})


// ════════════════════════════════════════════════════════════
//  MOONRAKER — helpers HTTP
// ════════════════════════════════════════════════════════════
const https = require('https')
const http  = require('http')

function moonrakerGet (baseUrl, endpoint) {
  return new Promise((resolve, reject) => {
    const url     = `${baseUrl}${endpoint}`
    const lib     = url.startsWith('https') ? https : http
    const timeout = setTimeout(() => reject(new Error('timeout')), 5000)
    lib.get(url, res => {
      let data = ''
      res.on('data', chunk => { data += chunk })
      res.on('end', () => {
        clearTimeout(timeout)
        try { resolve(JSON.parse(data)) }
        catch (e) { reject(e) }
      })
    }).on('error', e => { clearTimeout(timeout); reject(e) })
  })
}

function moonrakerPost (baseUrl, endpoint, body) {
  return new Promise((resolve, reject) => {
    const postData = body ? JSON.stringify(body) : ''
    const url      = new URL(`${baseUrl}${endpoint}`)
    const lib      = url.protocol === 'https:' ? https : http
    const options  = {
      hostname: url.hostname,
      port:     url.port || (url.protocol === 'https:' ? 443 : 80),
      path:     url.pathname + url.search,
      method:   'POST',
      headers: {
        'Content-Type':   'application/json',
        'Content-Length': Buffer.byteLength(postData)
      }
    }
    const timeout = setTimeout(() => reject(new Error('timeout')), 5000)
    const req = lib.request(options, res => {
      let data = ''
      res.on('data', chunk => { data += chunk })
      res.on('end', () => {
        clearTimeout(timeout)
        try { resolve(JSON.parse(data)) }
        catch (e) { reject(e) }
      })
    })
    req.on('error', e => { clearTimeout(timeout); reject(e) })
    req.write(postData)
    req.end()
  })
}

// ── Estado general de la impresora ──────────────────────────
// Devuelve: { online, state, temps: { hotend, bed, hotendTarget, bedTarget } }
ipcMain.handle('moonraker-status', async (_, baseUrl) => {
  try {
    const [printerInfo, temps] = await Promise.all([
      moonrakerGet(baseUrl, '/printer/info'),
      moonrakerGet(baseUrl, '/printer/objects/query?extruder=temperature,target&heater_bed=temperature,target')
    ])

    const extruder = temps?.result?.status?.extruder   || {}
    const bed      = temps?.result?.status?.heater_bed || {}
    const state    = printerInfo?.result?.state || 'offline'

    // Determinar estado legible
    let estadoLabel = 'Desconocido'
    const hotendTemp = extruder.temperature || 0
    const bedTemp    = bed.temperature      || 0
    const hotendTgt  = extruder.target      || 0
    const bedTgt     = bed.target           || 0

    if (state === 'printing')    estadoLabel = 'Imprimiendo'
    else if (state === 'paused') estadoLabel = 'Pausada'
    else if (state === 'standby' || state === 'ready') {
      if (hotendTgt > 0 && bedTgt > 0)   estadoLabel = 'Calentando todo'
      else if (hotendTgt > 0)             estadoLabel = 'Calentando boquilla'
      else if (bedTgt > 0)                estadoLabel = 'Calentando cama'
      else                                estadoLabel = 'Encendida / Idle'
    } else if (state === 'error')         estadoLabel = 'Error'
    else if (state === 'shutdown')        estadoLabel = 'Apagada'
    else                                  estadoLabel = 'Apagada'

    return {
      online:        true,
      state,
      estadoLabel,
      hotendTemp:    Math.round(hotendTemp),
      hotendTarget:  Math.round(hotendTgt),
      bedTemp:       Math.round(bedTemp),
      bedTarget:     Math.round(bedTgt)
    }
  } catch (e) {
    return { online: false, state: 'offline', estadoLabel: 'Apagada / Sin conexión',
             hotendTemp: 0, hotendTarget: 0, bedTemp: 0, bedTarget: 0 }
  }
})

// ── Trabajo actual ───────────────────────────────────────────
// Devuelve info del print job en curso (tiempos, capas, archivo)
ipcMain.handle('moonraker-job', async (_, baseUrl) => {
  try {
    const [jobStatus, displayStatus, printStats] = await Promise.all([
      moonrakerGet(baseUrl, '/printer/objects/query?print_stats=filename,total_duration,print_duration,filament_used,state,message&display_status=progress,message'),
      moonrakerGet(baseUrl, '/printer/objects/query?display_status=progress,message'),
      moonrakerGet(baseUrl, '/printer/objects/query?print_stats=filename,total_duration,print_duration,filament_used,state,info')
    ])

    const ps  = printStats?.result?.status?.print_stats  || {}
    const ds  = displayStatus?.result?.status?.display_status || {}

    const totalSec    = ps.total_duration  || 0
    const printedSec  = ps.print_duration  || 0
    const progress    = ds.progress        || 0
    const filename    = ps.filename        || ''
    const state       = ps.state          || 'standby'

    // Estimar tiempo restante
    let remainSec = 0
    if (progress > 0 && progress < 1 && printedSec > 0) {
      remainSec = Math.max(0, (printedSec / progress) - printedSec)
    }

    // Capas — Moonraker las expone via display_status o toolhead
    let currentLayer = ps.info?.current_layer  || 0
    let totalLayers  = ps.info?.total_layer    || 0

    // Nombre del objeto (archivo sin extensión)
    const objectName = filename.replace(/\.[^.]+$/, '')

    return {
      online:       true,
      state,
      filename,
      objectName,
      progress:     Math.round(progress * 100),
      printedSec:   Math.round(printedSec),
      totalSec:     Math.round(totalSec),
      remainSec:    Math.round(remainSec),
      currentLayer,
      totalLayers
    }
  } catch (e) {
    return { online: false, state: 'standby', filename: '', objectName: '',
             progress: 0, printedSec: 0, totalSec: 0, remainSec: 0,
             currentLayer: 0, totalLayers: 0 }
  }
})

// ── Consola (últimas líneas del log de Klipper) ──────────────
ipcMain.handle('moonraker-console', async (_, baseUrl) => {
  try {
    const res = await moonrakerGet(baseUrl, '/server/gcode_store?count=50')
    const items = res?.result?.gcode_store || []
    return items.map(e => ({
      time: e.time,
      msg:  e.message,
      type: e.type || 'command'
    }))
  } catch (e) {
    return []
  }
})

// ── Enviar G-code ────────────────────────────────────────────
ipcMain.handle('moonraker-send-gcode', async (_, baseUrl, cmd) => {
  try {
    const encoded = encodeURIComponent(cmd)
    const res = await moonrakerPost(baseUrl, `/printer/gcode/script?script=${encoded}`, null)
    return { ok: true, result: res }
  } catch (e) {
    return { ok: false, error: e.message }
  }
})
