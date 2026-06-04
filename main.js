const { app, BrowserWindow, ipcMain, dialog, shell } = require('electron')
const path = require('path')
const fs   = require('fs')
const { getDB, migrateFromExcel } = require('./db')

// ════════════════════════════════════════════════════════════
//  CONFIG
// ════════════════════════════════════════════════════════════
const CONFIG_PATH  = path.join(__dirname, 'config.json')
const STOCK_MINIMO = 250

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
//  DB — inicializar al arrancar
// ════════════════════════════════════════════════════════════
let db = null

function initDB () {
  const cfg    = loadConfig()
  const dbPath = cfg.rutaDB || path.join(__dirname, 'datos.db')
  db = getDB(dbPath)
  const migrated = migrateFromExcel(db, cfg)
  const total = Object.values(migrated).reduce((a, b) => a + b, 0)
  if (total > 0) {
    console.log('Migración desde Excel completada:', migrated)
  }
  return db
}

// ════════════════════════════════════════════════════════════
//  VENTANA
// ════════════════════════════════════════════════════════════
function createWindow () {
  const win = new BrowserWindow({
    width: 1200, height: 760,
    minWidth: 960, minHeight: 640,
    frame: false,
    show: false,
    backgroundColor: '#080B10',
    icon: path.join(__dirname, 'icono.ico'),
    webPreferences: {
      preload: path.join(__dirname, 'preload.js'),
      contextIsolation: true,
      nodeIntegration: false,
      webSecurity: false
    }
  })
  win.loadFile('index.html')
  win.once('ready-to-show', () => win.show())
}

app.whenReady().then(() => {
  initDB()
  createWindow()
})
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
  if (p && fs.existsSync(p)) { await shell.openPath(p); return true }
  return false
})

// ════════════════════════════════════════════════════════════
//  FILAMENTOS
// ════════════════════════════════════════════════════════════
ipcMain.handle('get-filamentos', () => {
  return db.prepare(`
    SELECT id as _idx, nombre, nombre as colorNombre, marca, tipo,
           costo_kg as costoKg, total_gr as pesoBobina, stock_gr as stockGr,
           0 as costoTotal, '' as fechaCompra, color_hex as colorHex, notas,
           CASE WHEN stock_gr < ${STOCK_MINIMO} THEN 1 ELSE 0 END as stockBajo
    FROM filamentos ORDER BY nombre
  `).all()
})

ipcMain.handle('save-filamento', (_, __, fil) => {
  const nombre = fil.nombre || fil.colorNombre || ''
  if (!nombre) return false

  let costoKg = fil.costoKg || 0
  if (!costoKg && fil.costoTotal && fil.pesoBobina)
    costoKg = +((fil.costoTotal / fil.pesoBobina) * 1000).toFixed(2)

  const stockGr = fil.stockGr !== undefined ? fil.stockGr : (fil.pesoBobina || 1000)

  if (fil._editIndex !== undefined) {
    // Edición — buscar por nombre original si viene, si no por _editIndex
    const existing = db.prepare('SELECT id FROM filamentos ORDER BY nombre LIMIT 1 OFFSET ?')
                       .get(fil._editIndex)
    if (existing) {
      db.prepare(`
        UPDATE filamentos SET
          nombre=?, marca=?, tipo=?, costo_kg=?, total_gr=?, stock_gr=?,
          color_hex=?, notas=?, updated_at=datetime('now','localtime')
        WHERE id=?
      `).run(nombre, fil.marca||'', fil.tipo||'PLA', costoKg,
             fil.pesoBobina||1000, stockGr, fil.colorHex||'#888888', fil.notas||'',
             existing.id)
      sincronizarFilamentoAInventario(nombre, stockGr, costoKg, fil.colorHex)
      return true
    }
  }

  // Upsert por nombre
  db.prepare(`
    INSERT INTO filamentos (nombre, marca, tipo, costo_kg, total_gr, stock_gr, color_hex, notas)
    VALUES (?,?,?,?,?,?,?,?)
    ON CONFLICT(nombre) DO UPDATE SET
      marca=excluded.marca, tipo=excluded.tipo, costo_kg=excluded.costo_kg,
      total_gr=excluded.total_gr, stock_gr=excluded.stock_gr,
      color_hex=excluded.color_hex, notas=excluded.notas,
      updated_at=datetime('now','localtime')
  `).run(nombre, fil.marca||'', fil.tipo||'PLA', costoKg,
         fil.pesoBobina||1000, stockGr, fil.colorHex||'#888888', fil.notas||'')

  sincronizarFilamentoAInventario(nombre, stockGr, costoKg, fil.colorHex)
  return true
})

ipcMain.handle('delete-filamento', (_, __, nombre) => {
  db.prepare('DELETE FROM filamentos WHERE nombre=?').run(nombre)
  return true
})

ipcMain.handle('get-stock-minimo', () => STOCK_MINIMO)

ipcMain.handle('get-alertas-stock', () => {
  return db.prepare(`
    SELECT nombre, nombre as colorNombre, stock_gr as stockGr, color_hex as colorHex
    FROM filamentos WHERE stock_gr < ?
  `).all(STOCK_MINIMO)
})

// ════════════════════════════════════════════════════════════
//  SINCRONIZACIÓN AUTOMÁTICA FILAMENTO → MONSAM
//  Actualiza stock en la tabla filamentos (que monsam también lee)
//  No hay copia — la DB es compartida, monsam lee directamente
// ════════════════════════════════════════════════════════════
function sincronizarFilamentoAInventario (nombre, stockGr, costoKg, colorHex) {
  // La DB es compartida — monsam lee filamentos directamente
  // Solo actualizamos updated_at para que monsam pueda detectar cambios
  db.prepare(`
    UPDATE filamentos SET updated_at=datetime('now','localtime')
    WHERE nombre=?
  `).run(nombre)
}

// ════════════════════════════════════════════════════════════
//  IMPRESIONES
// ════════════════════════════════════════════════════════════
ipcMain.handle('get-impresiones', () => {
  return db.prepare(`
    SELECT id as _idx, fecha, descripcion, filamento,
           gramos_usados as gramosUsados, tiempo, categoria,
           resultado, costo_material as costoMaterial,
           tipo_impresion as tipoImpresion
    FROM impresiones ORDER BY rowid
  `).all()
})

ipcMain.handle('save-impresion', (_, __, imp) => {
  // Calcular costo de material
  let costoMat = imp.costoMaterial || 0
  let costoKgFil = 0
  if (!costoMat && imp.gramosUsados && imp.filamento) {
    const fil = db.prepare('SELECT costo_kg FROM filamentos WHERE nombre=?').get(imp.filamento)
    if (fil) {
      costoKgFil = fil.costo_kg || 0
      costoMat   = +((imp.gramosUsados / 1000) * costoKgFil).toFixed(2)
    }
  }

  let nuevoStock = null

  if (imp._editIndex !== undefined) {
    // Obtener impresión anterior para calcular diferencia de gramos
    const anterior = db.prepare('SELECT gramos_usados, filamento FROM impresiones ORDER BY rowid LIMIT 1 OFFSET ?')
                       .get(imp._editIndex)
    if (anterior) {
      const diferencia = (imp.gramosUsados || 0) - (anterior.gramos_usados || 0)
      if (diferencia !== 0 && imp.filamento) {
        const fil = db.prepare('SELECT stock_gr FROM filamentos WHERE nombre=?').get(imp.filamento)
        if (fil) {
          nuevoStock = Math.max(0, (fil.stock_gr || 0) - diferencia)
          db.prepare(`UPDATE filamentos SET stock_gr=?, updated_at=datetime('now','localtime') WHERE nombre=?`)
            .run(nuevoStock, imp.filamento)
        }
      }
      db.prepare(`
        UPDATE impresiones SET
          fecha=?, descripcion=?, filamento=?, gramos_usados=?, tiempo=?,
          categoria=?, resultado=?, costo_material=?, tipo_impresion=?
        WHERE rowid = (SELECT rowid FROM impresiones ORDER BY rowid LIMIT 1 OFFSET ?)
      `).run(imp.fecha, imp.descripcion, imp.filamento, imp.gramosUsados||0,
             imp.tiempo||'', imp.categoria||'General', imp.resultado||'',
             costoMat, imp.tipoImpresion||'', imp._editIndex)
    }
  } else {
    // Nueva impresión — descontar stock
    if (imp.filamento && imp.gramosUsados) {
      const fil = db.prepare('SELECT stock_gr FROM filamentos WHERE nombre=?').get(imp.filamento)
      if (fil) {
        nuevoStock = Math.max(0, (fil.stock_gr || 0) - (imp.gramosUsados || 0))
        db.prepare(`UPDATE filamentos SET stock_gr=?, updated_at=datetime('now','localtime') WHERE nombre=?`)
          .run(nuevoStock, imp.filamento)
      }
    }
    db.prepare(`
      INSERT INTO impresiones
        (fecha, descripcion, filamento, gramos_usados, tiempo,
         categoria, resultado, costo_material, tipo_impresion)
      VALUES (?,?,?,?,?,?,?,?,?)
    `).run(imp.fecha, imp.descripcion, imp.filamento, imp.gramosUsados||0,
           imp.tiempo||'', imp.categoria||'General', imp.resultado||'',
           costoMat, imp.tipoImpresion||'')
  }

  return {
    ok:           true,
    stockBajo:    nuevoStock !== null && nuevoStock < STOCK_MINIMO,
    stockRestante: nuevoStock,
    filamento:    imp.filamento
  }
})

ipcMain.handle('delete-impresion', (_, __, rowIndex) => {
  const row = db.prepare('SELECT rowid, filamento, gramos_usados FROM impresiones ORDER BY rowid LIMIT 1 OFFSET ?')
                .get(rowIndex)
  if (!row) return false
  // Restaurar stock al eliminar impresión
  if (row.filamento && row.gramos_usados) {
    db.prepare(`UPDATE filamentos SET stock_gr = stock_gr + ?, updated_at=datetime('now','localtime') WHERE nombre=?`)
      .run(row.gramos_usados, row.filamento)
  }
  db.prepare('DELETE FROM impresiones WHERE rowid=?').run(row.rowid)
  return true
})

// ════════════════════════════════════════════════════════════
//  SYNC MONSAM (ya no hace falta — DB compartida)
//  Se mantiene la API para no romper el frontend
// ════════════════════════════════════════════════════════════
ipcMain.handle('sync-stock-monsan', () => {
  // DB compartida — el stock ya está sincronizado automáticamente
  const count = db.prepare('SELECT COUNT(*) as n FROM filamentos').get().n
  return { ok: true, actualizados: count, msg: `DB compartida — ${count} filamentos disponibles en ambas apps` }
})

ipcMain.handle('importar-filamentos-a-monsan', () => {
  // DB compartida — monsam ya lee directamente de la tabla filamentos
  const count = db.prepare('SELECT COUNT(*) as n FROM filamentos').get().n
  return { ok: true, importados: count, msg: `DB compartida — monsam ya lee los ${count} filamentos directamente` }
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

ipcMain.handle('moonraker-status', async (_, baseUrl) => {
  try {
    const [printerInfo, temps] = await Promise.all([
      moonrakerGet(baseUrl, '/printer/info'),
      moonrakerGet(baseUrl, '/printer/objects/query?extruder=temperature,target&heater_bed=temperature,target')
    ])
    const extruder = temps?.result?.status?.extruder   || {}
    const bed      = temps?.result?.status?.heater_bed || {}
    const state    = printerInfo?.result?.state || 'offline'
    const hotendTemp = extruder.temperature || 0
    const bedTemp    = bed.temperature      || 0
    const hotendTgt  = extruder.target      || 0
    const bedTgt     = bed.target           || 0
    let estadoLabel = 'Desconocido'
    if      (state === 'printing')  estadoLabel = 'Imprimiendo'
    else if (state === 'paused')    estadoLabel = 'Pausada'
    else if (state === 'standby' || state === 'ready') {
      if (hotendTgt > 0 && bedTgt > 0)   estadoLabel = 'Calentando todo'
      else if (hotendTgt > 0)             estadoLabel = 'Calentando boquilla'
      else if (bedTgt > 0)                estadoLabel = 'Calentando cama'
      else                                estadoLabel = 'Encendida / Idle'
    } else if (state === 'error')   estadoLabel = 'Error'
    else                            estadoLabel = 'Apagada'
    return { online:true, state, estadoLabel,
             hotendTemp:Math.round(hotendTemp), hotendTarget:Math.round(hotendTgt),
             bedTemp:Math.round(bedTemp), bedTarget:Math.round(bedTgt) }
  } catch {
    return { online:false, state:'offline', estadoLabel:'Apagada / Sin conexión',
             hotendTemp:0, hotendTarget:0, bedTemp:0, bedTarget:0 }
  }
})

ipcMain.handle('moonraker-job', async (_, baseUrl) => {
  try {
    const [displayStatus, printStats] = await Promise.all([
      moonrakerGet(baseUrl, '/printer/objects/query?display_status=progress,message'),
      moonrakerGet(baseUrl, '/printer/objects/query?print_stats=filename,total_duration,print_duration,filament_used,state,info')
    ])
    const ps  = printStats?.result?.status?.print_stats      || {}
    const ds  = displayStatus?.result?.status?.display_status || {}
    const totalSec   = ps.total_duration  || 0
    const printedSec = ps.print_duration  || 0
    const progress   = ds.progress        || 0
    const filename   = ps.filename        || ''
    const state      = ps.state           || 'standby'
    let remainSec = 0
    if (progress > 0 && progress < 1 && printedSec > 0)
      remainSec = Math.max(0, (printedSec / progress) - printedSec)
    return {
      online: true, state, filename,
      objectName:   filename.replace(/\.[^.]+$/, ''),
      progress:     Math.round(progress * 100),
      printedSec:   Math.round(printedSec),
      totalSec:     Math.round(totalSec),
      remainSec:    Math.round(remainSec),
      currentLayer: ps.info?.current_layer || 0,
      totalLayers:  ps.info?.total_layer   || 0
    }
  } catch {
    return { online:false, state:'standby', filename:'', objectName:'',
             progress:0, printedSec:0, totalSec:0, remainSec:0,
             currentLayer:0, totalLayers:0 }
  }
})

ipcMain.handle('moonraker-console', async (_, baseUrl) => {
  try {
    const res   = await moonrakerGet(baseUrl, '/server/gcode_store?count=50')
    const items = res?.result?.gcode_store || []
    return items.map(e => ({ time: e.time, msg: e.message, type: e.type || 'command' }))
  } catch { return [] }
})

ipcMain.handle('moonraker-send-gcode', async (_, baseUrl, cmd) => {
  try {
    const res = await moonrakerPost(baseUrl, `/printer/gcode/script?script=${encodeURIComponent(cmd)}`, null)
    return { ok: true, result: res }
  } catch (e) {
    return { ok: false, error: e.message }
  }
})
