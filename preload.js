const { contextBridge, ipcRenderer } = require('electron')

contextBridge.exposeInMainWorld('api', {

  // ── Ventana ──────────────────────────────────────────────
  minimize:          ()          => ipcRenderer.send('win-minimize'),
  maximize:          ()          => ipcRenderer.send('win-maximize'),
  close:             ()          => ipcRenderer.send('win-close'),

  // ── Configuración ────────────────────────────────────────
  getConfig:         ()          => ipcRenderer.invoke('get-config'),
  saveConfig:        (cfg)       => ipcRenderer.invoke('save-config', cfg),
  selectExcel:       (def)       => ipcRenderer.invoke('select-excel', def),
  openExcel:         (p)         => ipcRenderer.invoke('open-excel', p),

  // ── Filamentos ───────────────────────────────────────────
  getFilamentos:     (p)         => ipcRenderer.invoke('get-filamentos', p),
  saveFilamento:     (p, f)      => ipcRenderer.invoke('save-filamento', p, f),
  deleteFilamento:   (p, n)      => ipcRenderer.invoke('delete-filamento', p, n),
  descontarFilamento:(p, n, gr)  => ipcRenderer.invoke('descontar-filamento', p, n, gr),

  // ── Impresiones ──────────────────────────────────────────
  getImpresiones:    (p)         => ipcRenderer.invoke('get-impresiones', p),
  saveImpresion:     (p, i)      => ipcRenderer.invoke('save-impresion', p, i),
  deleteImpresion:   (p, idx)    => ipcRenderer.invoke('delete-impresion', p, idx),

  // ── Sincronizar stock con MONSAN ─────────────────────────
  syncStockMonsan:   (p1, p2)    => ipcRenderer.invoke('sync-stock-monsan', p1, p2),
})
