/**
 * db.js — Capa de acceso a datos SQLite compartida entre e500-app y monsam-app
 *
 * Base de datos: datos.db (ruta configurada en config.json → rutaDB)
 *
 * Tablas:
 *   filamentos   — stock de filamentos (fuente de verdad: e500-app)
 *   impresiones  — historial de impresiones (e500-app)
 *   colores      — alias/vista de filamentos para monsam-app
 *   inventario   — pares por SKU/modelo/color (monsam-app)
 *   clientes     — directos, distribuidores, mayoreo (monsam-app)
 *   costos       — historial de costos de producción (monsam-app)
 */

const Database = require('better-sqlite3')
const path     = require('path')
const fs       = require('fs')
const XLSX     = require('xlsx')

// ── Singleton ────────────────────────────────────────────────
let _db = null

function getDB (dbPath) {
  if (_db) return _db
  if (!dbPath) throw new Error('dbPath requerido')
  const dir = path.dirname(dbPath)
  if (!fs.existsSync(dir)) fs.mkdirSync(dir, { recursive: true })
  _db = new Database(dbPath)
  _db.pragma('journal_mode = WAL')   // escrituras concurrentes seguras
  _db.pragma('foreign_keys = ON')
  initSchema(_db)
  return _db
}

// ════════════════════════════════════════════════════════════
//  SCHEMA
// ════════════════════════════════════════════════════════════
function initSchema (db) {
  db.exec(`
    -- ── Filamentos (e500 es fuente de verdad) ──────────────
    CREATE TABLE IF NOT EXISTS filamentos (
      id          INTEGER PRIMARY KEY AUTOINCREMENT,
      nombre      TEXT    NOT NULL UNIQUE,
      marca       TEXT    DEFAULT '',
      tipo        TEXT    DEFAULT 'PLA',
      costo_kg    REAL    DEFAULT 0,
      total_gr    REAL    DEFAULT 1000,
      stock_gr    REAL    DEFAULT 0,
      color_hex   TEXT    DEFAULT '#888888',
      notas       TEXT    DEFAULT '',
      created_at  TEXT    DEFAULT (datetime('now','localtime')),
      updated_at  TEXT    DEFAULT (datetime('now','localtime'))
    );

    -- ── Impresiones (e500) ──────────────────────────────────
    CREATE TABLE IF NOT EXISTS impresiones (
      id              INTEGER PRIMARY KEY AUTOINCREMENT,
      fecha           TEXT    DEFAULT '',
      descripcion     TEXT    DEFAULT '',
      filamento       TEXT    DEFAULT '',
      gramos_usados   REAL    DEFAULT 0,
      tiempo          TEXT    DEFAULT '',
      categoria       TEXT    DEFAULT 'General',
      resultado       TEXT    DEFAULT '',
      costo_material  REAL    DEFAULT 0,
      tipo_impresion  TEXT    DEFAULT '',
      created_at      TEXT    DEFAULT (datetime('now','localtime'))
    );

    -- ── Inventario (monsam) ─────────────────────────────────
    CREATE TABLE IF NOT EXISTS inventario (
      id               INTEGER PRIMARY KEY AUTOINCREMENT,
      sku              TEXT    NOT NULL UNIQUE,
      modelo           TEXT    DEFAULT '',
      color            TEXT    DEFAULT '',
      codigo_color     TEXT    DEFAULT '',
      pares            INTEGER DEFAULT 0,
      costo_produccion REAL    DEFAULT 0,
      updated_at       TEXT    DEFAULT (datetime('now','localtime'))
    );

    -- ── Clientes (monsam) ───────────────────────────────────
    CREATE TABLE IF NOT EXISTS clientes (
      id                    INTEGER PRIMARY KEY AUTOINCREMENT,
      tipo                  TEXT    NOT NULL DEFAULT 'directo',
      nombre                TEXT    NOT NULL,
      contacto              TEXT    DEFAULT '',
      direccion             TEXT    DEFAULT '',
      rfc                   TEXT    DEFAULT '',
      notas                 TEXT    DEFAULT '',
      fecha_registro        TEXT    DEFAULT '',
      fecha_ultima_compra   TEXT    DEFAULT '',
      piezas_ultima_compra  INTEGER DEFAULT 0,
      acumulado_historico   REAL    DEFAULT 0,
      skus                  TEXT    DEFAULT '',
      created_at            TEXT    DEFAULT (datetime('now','localtime'))
    );

    -- ── Costos (monsam) ─────────────────────────────────────
    CREATE TABLE IF NOT EXISTS costos (
      id               INTEGER PRIMARY KEY AUTOINCREMENT,
      sku              TEXT    DEFAULT '',
      modelo           TEXT    DEFAULT '',
      color            TEXT    DEFAULT '',
      peso_gr          REAL    DEFAULT 0,
      horas_imp        REAL    DEFAULT 0,
      costo_filamento  REAL    DEFAULT 0,
      costo_elec       REAL    DEFAULT 0,
      costo_herrajes   REAL    DEFAULT 0,
      costo_empaque    REAL    DEFAULT 0,
      desperdicio      REAL    DEFAULT 0,
      costo_total      REAL    DEFAULT 0,
      fecha            TEXT    DEFAULT '',
      created_at       TEXT    DEFAULT (datetime('now','localtime'))
    );

    -- ── Índices ─────────────────────────────────────────────
    CREATE INDEX IF NOT EXISTS idx_imp_fecha     ON impresiones (fecha);
    CREATE INDEX IF NOT EXISTS idx_imp_filamento ON impresiones (filamento);
    CREATE INDEX IF NOT EXISTS idx_inv_sku       ON inventario  (sku);
    CREATE INDEX IF NOT EXISTS idx_cli_tipo      ON clientes    (tipo);
    CREATE INDEX IF NOT EXISTS idx_cos_sku       ON costos      (sku);
  `)
}

// ════════════════════════════════════════════════════════════
//  MIGRACIÓN DESDE EXCEL
//  Se ejecuta una sola vez si las tablas están vacías
// ════════════════════════════════════════════════════════════
function readWB (filePath) {
  if (!filePath || !fs.existsSync(filePath)) return null
  try { return XLSX.readFile(filePath) } catch { return null }
}
function toRows (ws) {
  return XLSX.utils.sheet_to_json(ws, { header: 1, defval: '' })
}

function migrateFromExcel (db, cfg) {
  const migrated = { filamentos: 0, impresiones: 0, inventario: 0, clientes: 0, costos: 0 }

  // ── Filamentos desde e500 ──────────────────────────────
  // Estructura real: [0]Nombre [1]Marca [2]Tipo [3]CostoKg [4]TotalGr [5]StockGr [8]ColorHex
  const countFil = db.prepare('SELECT COUNT(*) as n FROM filamentos').get().n
  if (countFil === 0 && cfg.rutaExcel) {
    const wbE = readWB(cfg.rutaExcel)
    if (wbE && wbE.Sheets['Filamentos']) {
      const rows      = toRows(wbE.Sheets['Filamentos'])
      const dataStart = rows[0] && String(rows[0][0]).toLowerCase().includes('nombre') ? 1 : 0
      const ins = db.prepare(`
        INSERT OR IGNORE INTO filamentos (nombre, marca, tipo, costo_kg, total_gr, stock_gr, color_hex)
        VALUES (?,?,?,?,?,?,?)
      `)
      const many = db.transaction((data) => {
        for (const r of data) ins.run(r)
      })
      const data = rows.slice(dataStart)
        .filter(r => String(r[0]).trim())
        .map(r => [
          String(r[0]).trim(),
          String(r[1] || '').trim(),
          String(r[2] || 'PLA').trim(),
          Number(r[3]) || 0,
          Number(r[4]) || 1000,
          Number(r[5]) || 0,
          String(r[8] || '#888888').trim()
        ])
      many(data)
      migrated.filamentos = data.length
    }
  }

  // ── Impresiones desde e500 ─────────────────────────────
  // Estructura: [0]Fecha [1]Descripcion [2]Filamento [3]GramosUsados
  //             [4]TiempoImpresion [5]Categoria [6]Resultado [7]CostoMaterial
  const countImp = db.prepare('SELECT COUNT(*) as n FROM impresiones').get().n
  if (countImp === 0 && cfg.rutaExcel) {
    const wbE = readWB(cfg.rutaExcel)
    if (wbE && wbE.Sheets['Impresiones']) {
      const rows      = toRows(wbE.Sheets['Impresiones'])
      const dataStart = rows[0] && String(rows[0][0]).toLowerCase().includes('fecha') ? 1 : 0
      const ins = db.prepare(`
        INSERT INTO impresiones
          (fecha, descripcion, filamento, gramos_usados, tiempo, categoria, resultado, costo_material)
        VALUES (?,?,?,?,?,?,?,?)
      `)
      const many = db.transaction((data) => {
        for (const r of data) ins.run(r)
      })
      const data = rows.slice(dataStart)
        .filter(r => r[0] || r[1] || r[2])
        .map(r => [
          String(r[0] || ''),
          String(r[1] || ''),
          String(r[2] || ''),
          Number(r[3]) || 0,
          String(r[4] || ''),
          String(r[5] || 'General'),
          String(r[6] || ''),
          Number(r[7]) || 0
        ])
      many(data)
      migrated.impresiones = data.length
    }
  }

  // ── Inventario desde monsam ────────────────────────────
  // Estructura: [0]SKU [1]Modelo [2]Color [3]CodigoColor [4]Pares [5]CostoProduccion
  const countInv = db.prepare('SELECT COUNT(*) as n FROM inventario').get().n
  if (countInv === 0 && cfg.rutaExcelMonsan) {
    const wbM = readWB(cfg.rutaExcelMonsan)
    if (wbM && wbM.Sheets['Inventario']) {
      const rows      = toRows(wbM.Sheets['Inventario'])
      const dataStart = rows[0] && String(rows[0][0]).toLowerCase().includes('sku') ? 1 : 0
      const ins = db.prepare(`
        INSERT OR IGNORE INTO inventario (sku, modelo, color, codigo_color, pares, costo_produccion)
        VALUES (?,?,?,?,?,?)
      `)
      const many = db.transaction((data) => {
        for (const r of data) ins.run(r)
      })
      const data = rows.slice(dataStart)
        .filter(r => String(r[0]).trim())
        .map(r => [
          String(r[0]).trim(),
          String(r[1] || ''),
          String(r[2] || ''),
          String(r[3] || ''),
          Number(r[4]) || 0,
          Number(r[5]) || 0
        ])
      many(data)
      migrated.inventario = data.length
    }
  }

  // ── Clientes desde monsam (3 hojas) ───────────────────
  const countCli = db.prepare('SELECT COUNT(*) as n FROM clientes').get().n
  if (countCli === 0 && cfg.rutaExcelMonsan) {
    const wbM = readWB(cfg.rutaExcelMonsan)
    if (wbM) {
      const TIPO_SHEET = {
        directo:      'Clientes_Directos',
        distribuidor: 'Distribuidores',
        mayoreo:      'Mayoreo'
      }
      const ins = db.prepare(`
        INSERT INTO clientes
          (tipo, nombre, contacto, direccion, rfc, notas,
           fecha_registro, fecha_ultima_compra, piezas_ultima_compra, acumulado_historico, skus)
        VALUES (?,?,?,?,?,?,?,?,?,?,?)
      `)
      const many = db.transaction((data) => {
        for (const r of data) ins.run(r)
      })
      for (const [tipo, sheet] of Object.entries(TIPO_SHEET)) {
        if (!wbM.Sheets[sheet]) continue
        const rows      = toRows(wbM.Sheets[sheet])
        const dataStart = rows[0] && String(rows[0][0]).toLowerCase().includes('nombre') ? 1 : 0
        const data = rows.slice(dataStart)
          .filter(r => String(r[0]).trim())
          .map(r => [
            tipo,
            String(r[0]).trim(),
            String(r[1] || ''),
            String(r[2] || ''),
            String(r[3] || ''),
            String(r[4] || ''),
            String(r[5] || ''),
            String(r[6] || ''),
            Number(r[7]) || 0,
            Number(r[8]) || 0,
            String(r[9] || '')
          ])
        many(data)
        migrated.clientes += data.length
      }
    }
  }

  // ── Costos desde monsam ───────────────────────────────
  const countCos = db.prepare('SELECT COUNT(*) as n FROM costos').get().n
  if (countCos === 0 && cfg.rutaExcelMonsan) {
    const wbM = readWB(cfg.rutaExcelMonsan)
    if (wbM && wbM.Sheets['Costos']) {
      const rows      = toRows(wbM.Sheets['Costos'])
      const dataStart = rows[0] && String(rows[0][0]).toLowerCase().includes('sku') ? 1 : 0
      const ins = db.prepare(`
        INSERT INTO costos
          (sku, modelo, color, peso_gr, horas_imp,
           costo_filamento, costo_elec, costo_herrajes,
           costo_empaque, desperdicio, costo_total, fecha)
        VALUES (?,?,?,?,?,?,?,?,?,?,?,?)
      `)
      const many = db.transaction((data) => {
        for (const r of data) ins.run(r)
      })
      const data = rows.slice(dataStart)
        .filter(r => r[0])
        .map(r => [
          String(r[0] || ''), String(r[1] || ''), String(r[2] || ''),
          Number(r[3]) || 0,  Number(r[4]) || 0,
          Number(r[5]) || 0,  Number(r[6]) || 0,  Number(r[7]) || 0,
          Number(r[8]) || 0,  Number(r[9]) || 0,  Number(r[10]) || 0,
          String(r[11] || '')
        ])
      many(data)
      migrated.costos = data.length
    }
  }

  return migrated
}

// ════════════════════════════════════════════════════════════
//  EXPORTS
// ════════════════════════════════════════════════════════════
module.exports = { getDB, migrateFromExcel }
