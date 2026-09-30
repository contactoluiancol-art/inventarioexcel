/**
 * SISTEMA DE CONTROL DE AVANCE DE INVENTARIO
 * Archivo: app.js
 * Corrección: Blindaje de referencias del DOM para evitar errores de tipo null.
 */

'use strict';

document.addEventListener('DOMContentLoaded', () => {

  // ==========================================================================
  // 1. ESTADO DE LA APLICACIÓN (STATE)
  // ==========================================================================
  const state = {
    masterList: [],       // Registros individuales de la base teórica
    physicalList: [],     // Registros individuales del conteo físico
    reconciliation: [],   // Cruce fila a fila
    activeFilter: 'all',  // Filtro activo: 'all' | 'pending'
    searchQuery: ''       // Término de búsqueda de la tabla
  };

  // ==========================================================================
  // 2. REFERENCIAS A ELEMENTOS DEL DOM
  // ==========================================================================
  const elements = {
    // Inputs de archivo y zonas drag & drop
    fileMaster: document.getElementById('fileMaster'),
    filePhysical: document.getElementById('filePhysical'),
    zoneMaster: document.getElementById('zoneMaster'),
    zonePhysical: document.getElementById('zonePhysical'),
    nameMaster: document.getElementById('nameMaster'),
    namePhysical: document.getElementById('namePhysical'),
    systemStatus: document.getElementById('systemStatus'),

    // Tarjetas de Métricas (Dashboard)
    kpiMasterCount: document.getElementById('kpiMasterCount'),
    kpiMasterUnits: document.getElementById('kpiMasterUnits'),
    kpiCountedItems: document.getElementById('kpiCountedItems'),
    kpiCountedUnits: document.getElementById('kpiCountedUnits'),
    kpiPendingCount: document.getElementById('kpiPendingCount'),
    kpiPendingUnits: document.getElementById('kpiPendingUnits'),

    // Barra de Avance
    progressBar: document.getElementById('progressBar'),
    progressPercent: document.getElementById('progressPercent'),

    // Pestañas (solo Todos y Faltantes)
    tabButtons: document.querySelectorAll('.tab-btn'),
    tableSearch: document.getElementById('tableSearch'),
    btnExport: document.getElementById('btnExport'),
    countTabAll: document.getElementById('countTabAll'),
    countTabPending: document.getElementById('countTabPending'),
    tableBody: document.getElementById('tableBody'),
    tableSummary: document.getElementById('tableSummary'),

    // Overlay de Espera
    processingOverlay: document.getElementById('processingOverlay'),
    overlayTitle: document.getElementById('overlayTitle'),
    overlayMessage: document.getElementById('overlayMessage')
  };

  // ==========================================================================
  // 3. FUNCIONES DE UTILIDAD Y SANITIZACIÓN (DEFENSA XSS)
  // ==========================================================================

  function escapeHTML(val) {
    if (val === null || val === undefined) return '';
    return String(val)
      .replace(/&/g, '&amp;')
      .replace(/</g, '&lt;')
      .replace(/>/g, '&gt;')
      .replace(/"/g, '&quot;')
      .replace(/'/g, '&#039;');
  }

  function normalizeKey(str) {
    if (str === null || str === undefined) return '';
    return String(str).trim().toUpperCase();
  }

  function setOverlay(show, title = 'Procesando...', message = 'Por favor espera.') {
    if (!elements.processingOverlay) return;
    if (show) {
      if (elements.overlayTitle) elements.overlayTitle.textContent = title;
      if (elements.overlayMessage) elements.overlayMessage.textContent = message;
      elements.processingOverlay.classList.remove('hidden');
    } else {
      elements.processingOverlay.classList.add('hidden');
    }
  }

  function detectColumns(headers) {
    const normalized = headers.map(h => String(h || '').trim().toLowerCase());
    const patterns = {
      sku: ['sku', 'codigo', 'código', 'referencia', 'ref', 'item', 'artículo', 'articulo', 'cod'],
      desc: ['descripcion', 'descripción', 'detalle', 'nombre', 'producto', 'name'],
      qty: ['cantidad', 'cant', 'qty', 'stock', 'teorico', 'teórico', 'conteo', 'total', 'unidades', 'fisico', 'físico']
    };
    const findIndex = (keys) => normalized.findIndex(h => keys.some(k => h.includes(k)));

    return {
      skuIdx: findIndex(patterns.sku),
      descIdx: findIndex(patterns.desc),
      qtyIdx: findIndex(patterns.qty)
    };
  }

  function parseExcelFile(file) {
    return new Promise((resolve, reject) => {
      if (!window.XLSX) {
        reject(new Error('La librería SheetJS no está cargada.'));
        return;
      }
      const reader = new FileReader();
      reader.onload = (e) => {
        try {
          const data = new Uint8Array(e.target.result);
          const workbook = XLSX.read(data, { type: 'array', cellDates: true });
          const firstSheet = workbook.Sheets[workbook.SheetNames[0]];
          const rawRows = XLSX.utils.sheet_to_json(firstSheet, { header: 1, defval: '' });
          
          if (!rawRows || rawRows.length < 2) {
            reject(new Error('El archivo no contiene filas suficientes.'));
            return;
          }
          resolve(rawRows);
        } catch (err) {
          reject(new Error(`Error al leer el archivo: ${err.message}`));
        }
      };
      reader.onerror = () => reject(new Error('Error al leer el archivo físico.'));
      reader.readAsArrayBuffer(file);
    });
  }

  // ==========================================================================
  // 4. LECTURA DE ARCHIVOS
  // ==========================================================================

  async function handleMasterFile(file) {
    if (!file) return;
    setOverlay(true, 'Cargando Base Maestra', 'Leyendo total de filas teóricas...');

    try {
      const rows = await parseExcelFile(file);
      const cols = detectColumns(rows[0]);
      const skuIdx = cols.skuIdx !== -1 ? cols.skuIdx : 0;
      const descIdx = cols.descIdx !== -1 ? cols.descIdx : (rows[0].length > 1 ? 1 : -1);
      const qtyIdx = cols.qtyIdx !== -1 ? cols.qtyIdx : (rows[0].length > 2 ? 2 : -1);

      state.masterList = [];

      for (let i = 1; i < rows.length; i++) {
        const row = rows[i];
        const rawSku = row[skuIdx];
        if (rawSku === undefined || rawSku === null || String(rawSku).trim() === '') continue;

        const sku = normalizeKey(rawSku);
        const desc = descIdx !== -1 ? String(row[descIdx] || '').trim() : 'Sin descripción';
        
        let qtyExpected = 1;
        if (qtyIdx !== -1 && row[qtyIdx] !== undefined && row[qtyIdx] !== null && String(row[qtyIdx]).trim() !== '') {
          const sanitized = String(row[qtyIdx]).replace(/,/g, '.').trim();
          const parsed = parseFloat(sanitized);
          qtyExpected = isNaN(parsed) ? 1 : parsed;
        }

        state.masterList.push({
          rowId: `M-${i}`,
          sku: sku,
          desc: desc,
          qtyExpected: qtyExpected
        });
      }

      if (elements.nameMaster) {
        elements.nameMaster.textContent = `${file.name} (${state.masterList.length} registros)`;
      }
      runReconciliation();
    } catch (err) {
      alert(`Error en Base Maestra: ${err.message}`);
    } finally {
      setOverlay(false);
    }
  }

  async function handlePhysicalFile(file) {
    if (!file) return;
    setOverlay(true, 'Cargando Avance Físico', 'Calculando unidades contadas en campo...');

    try {
      const rows = await parseExcelFile(file);
      const cols = detectColumns(rows[0]);
      const skuIdx = cols.skuIdx !== -1 ? cols.skuIdx : 0;
      const qtyIdx = cols.qtyIdx !== -1 ? cols.qtyIdx : (rows[0].length > 1 ? 1 : -1);

      state.physicalList = [];

      for (let i = 1; i < rows.length; i++) {
        const row = rows[i];
        const rawSku = row[skuIdx];
        if (rawSku === undefined || rawSku === null || String(rawSku).trim() === '') continue;

        const sku = normalizeKey(rawSku);

        let countedAmount = 1;
        if (qtyIdx !== -1 && row[qtyIdx] !== undefined && row[qtyIdx] !== null && String(row[qtyIdx]).trim() !== '') {
          const sanitized = String(row[qtyIdx]).replace(/,/g, '.').trim();
          const parsed = parseFloat(sanitized);
          countedAmount = isNaN(parsed) ? 1 : parsed;
        }

        state.physicalList.push({
          rowId: `P-${i}`,
          sku: sku,
          qtyCounted: countedAmount
        });
      }

      if (elements.namePhysical) {
        elements.namePhysical.textContent = `${file.name} (${state.physicalList.length} registros contados)`;
      }
      runReconciliation();
    } catch (err) {
      alert(`Error en Conteo Físico: ${err.message}`);
    } finally {
      setOverlay(false);
    }
  }

  // ==========================================================================
  // 5. MOTOR DE CONCILIACIÓN
  // ==========================================================================
  function runReconciliation() {
    if (state.masterList.length === 0 && state.physicalList.length === 0) return;

    state.reconciliation = [];

    const physicalQueue = new Map();
    state.physicalList.forEach(item => {
      if (!physicalQueue.has(item.sku)) {
        physicalQueue.set(item.sku, []);
      }
      physicalQueue.get(item.sku).push({ ...item, matched: false });
    });

    // 1. Recorrer cada ítem de la base maestra
    state.masterList.forEach(masterItem => {
      const queue = physicalQueue.get(masterItem.sku);

      if (queue && queue.length > 0) {
        const match = queue.find(p => !p.matched);

        if (match) {
          match.matched = true;
          const diff = match.qtyCounted - masterItem.qtyExpected;
          state.reconciliation.push({
            id: masterItem.rowId,
            sku: masterItem.sku,
            desc: masterItem.desc,
            expected: masterItem.qtyExpected,
            counted: match.qtyCounted,
            difference: diff,
            status: diff === 0 ? 'EXACTO' : 'DISCREPANCIA',
            statusClass: diff === 0 ? 'exact' : 'discrepancy'
          });
        } else {
          state.reconciliation.push({
            id: masterItem.rowId,
            sku: masterItem.sku,
            desc: masterItem.desc,
            expected: masterItem.qtyExpected,
            counted: 0,
            difference: -masterItem.qtyExpected,
            status: 'PENDIENTE',
            statusClass: 'pending'
          });
        }
      } else {
        state.reconciliation.push({
          id: masterItem.rowId,
          sku: masterItem.sku,
          desc: masterItem.desc,
          expected: masterItem.qtyExpected,
          counted: 0,
          difference: -masterItem.qtyExpected,
          status: 'PENDIENTE',
          statusClass: 'pending'
        });
      }
    });

    // 2. Registros contados que no están en la base maestra
    physicalQueue.forEach((items, sku) => {
      items.filter(p => !p.matched).forEach(unmatched => {
        state.reconciliation.push({
          id: unmatched.rowId,
          sku: sku,
          desc: 'ÍTEM CONTADO NO REGISTRADO EN MAESTRO',
          expected: 0,
          counted: unmatched.qtyCounted,
          difference: unmatched.qtyCounted,
          status: 'NO CATALOGADO',
          statusClass: 'unknown'
        });
      });
    });

    updateDashboardMetrics();
    renderTable();
  }

  // ==========================================================================
  // 6. DASHBOARD DE MÉTRICAS (KPIS Y AVANCE)
  // ==========================================================================
  function updateDashboardMetrics() {
    // 1. Maestro Teórico
    const totalMasterRows = state.masterList.length;
    let totalMasterUnits = 0;
    state.masterList.forEach(item => totalMasterUnits += item.qtyExpected);

    // 2. Conteo Físico Realizado
    const totalPhysicalRows = state.physicalList.length;
    let totalUnitsCounted = 0;
    state.physicalList.forEach(item => {
      totalUnitsCounted += (typeof item.qtyCounted === 'number' ? item.qtyCounted : 0);
    });

    // 3. Faltantes
    let countedMasterRows = 0;
    let pendingRows = 0;
    let pendingUnits = 0;

    state.reconciliation.forEach(item => {
      if (item.status === 'EXACTO' || item.status === 'DISCREPANCIA') {
        countedMasterRows++;
      }
      if (item.status === 'PENDIENTE') {
        pendingRows++;
        pendingUnits += item.expected;
      }
    });

    // Actualización de contadores en pestañas (con verificación de existencia)
    if (elements.countTabAll) {
      elements.countTabAll.textContent = state.reconciliation.length;
    }
    if (elements.countTabPending) {
      elements.countTabPending.textContent = pendingRows;
    }

    // Porcentaje de avance global
    const progress = totalMasterRows > 0 ? (countedMasterRows / totalMasterRows) * 100 : 0;
    const cleanProgress = Math.min(100, Math.max(0, progress)).toFixed(1);

    // Render en Tarjetas de KPIs (con verificación segura)
    if (elements.kpiMasterCount) {
      elements.kpiMasterCount.textContent = totalMasterRows.toLocaleString();
    }
    if (elements.kpiMasterUnits) {
      elements.kpiMasterUnits.textContent = `${totalMasterUnits.toLocaleString()} unidades teóricas`;
    }

    if (elements.kpiCountedItems) {
      elements.kpiCountedItems.textContent = totalPhysicalRows.toLocaleString();
    }
    if (elements.kpiCountedUnits) {
      elements.kpiCountedUnits.textContent = `${totalUnitsCounted.toLocaleString()} unidades contadas`;
    }

    if (elements.kpiPendingCount) {
      elements.kpiPendingCount.textContent = pendingRows.toLocaleString();
    }
    if (elements.kpiPendingUnits) {
      elements.kpiPendingUnits.textContent = `${pendingUnits.toLocaleString()} unidades por contar`;
    }

    // Barra de Progreso
    if (elements.progressPercent) {
      elements.progressPercent.textContent = `${cleanProgress}%`;
    }
    if (elements.progressBar) {
      elements.progressBar.style.width = `${cleanProgress}%`;
      if (progress < 50) {
        elements.progressBar.style.backgroundColor = 'var(--color-warning)';
      } else {
        elements.progressBar.style.backgroundColor = 'var(--color-success)';
      }
    }

    if (elements.btnExport && (state.reconciliation.length > 0 || totalPhysicalRows > 0)) {
      elements.btnExport.removeAttribute('disabled');
    }
    if (elements.systemStatus) {
      elements.systemStatus.textContent = 'Avance Actualizado';
      elements.systemStatus.classList.add('active');
    }
  }

  // ==========================================================================
  // 7. RENDERIZADO DE LA TABLA
  // ==========================================================================
  function renderTable() {
    if (!elements.tableBody) return;

    const filter = state.activeFilter;
    const query = state.searchQuery.toLowerCase().trim();

    const filtered = state.reconciliation.filter(item => {
      let matchTab = true;
      if (filter === 'pending') {
        matchTab = item.status === 'PENDIENTE';
      }

      if (!matchTab) return false;
      if (query !== '') {
        return item.sku.toLowerCase().includes(query) || item.desc.toLowerCase().includes(query);
      }
      return true;
    });

    if (filtered.length === 0) {
      elements.tableBody.innerHTML = `
        <tr>
          <td colspan="6" class="table-empty">No hay registros con los filtros seleccionados.</td>
        </tr>
      `;
      if (elements.tableSummary) elements.tableSummary.textContent = 'Mostrando 0 registros';
      return;
    }

    const rows = filtered.slice(0, 500);

    elements.tableBody.innerHTML = rows.map(item => {
      const diffFormatted = item.difference > 0 ? `+${item.difference}` : `${item.difference}`;
      return `
        <tr>
          <td><strong>${escapeHTML(item.sku)}</strong></td>
          <td>${escapeHTML(item.desc)}</td>
          <td class="text-right">${escapeHTML(item.expected)}</td>
          <td class="text-right">${escapeHTML(item.counted)}</td>
          <td class="text-right ${item.difference !== 0 ? 'alert-text' : ''}"><strong>${escapeHTML(diffFormatted)}</strong></td>
          <td class="text-center"><span class="status-pill ${escapeHTML(item.statusClass)}">${escapeHTML(item.status)}</span></td>
        </tr>
      `;
    }).join('');

    if (elements.tableSummary) {
      elements.tableSummary.textContent = `Mostrando ${rows.length} de ${filtered.length} registros`;
    }
  }

  // ==========================================================================
  // 8. EXPORTACIÓN DEL REPORTE A EXCEL
  // ==========================================================================
  function exportReport() {
    if (state.reconciliation.length === 0) return;
    setOverlay(true, 'Exportando', 'Generando archivo Excel de avance...');

    setTimeout(() => {
      try {
        const wb = XLSX.utils.book_new();

        const dataSheet = state.reconciliation.map(item => ({
          'Código SKU': item.sku,
          'Descripción': item.desc,
          'Cantidad Teórica': item.expected,
          'Cantidad Contada': item.counted,
          'Diferencia': item.difference,
          'Estado': item.status
        }));

        const ws = XLSX.utils.json_to_sheet(dataSheet);
        XLSX.utils.book_append_sheet(wb, ws, 'Avance_Inventario');
        XLSX.writeFile(wb, `Avance_Inventario_${new Date().toISOString().slice(0, 10)}.xlsx`);
      } catch (err) {
        alert(`Error al exportar: ${err.message}`);
      } finally {
        setOverlay(false);
      }
    }, 150);
  }

  // ==========================================================================
  // 9. CONFIGURACIÓN DE EVENTOS Y DRAG & DROP
  // ==========================================================================
  function setupDragAndDrop(zone, input, handler) {
    if (!zone || !input) return;

    ['dragenter', 'dragover'].forEach(eName => {
      zone.addEventListener(eName, (e) => {
        e.preventDefault();
        zone.classList.add('drag-over');
      }, false);
    });

    ['dragleave', 'drop'].forEach(eName => {
      zone.addEventListener(eName, (e) => {
        e.preventDefault();
        zone.classList.remove('drag-over');
      }, false);
    });

    zone.addEventListener('drop', (e) => {
      if (e.dataTransfer.files.length > 0) handler(e.dataTransfer.files[0]);
    });

    input.addEventListener('change', (e) => {
      if (e.target.files.length > 0) handler(e.target.files[0]);
    });
  }

  setupDragAndDrop(elements.zoneMaster, elements.fileMaster, handleMasterFile);
  setupDragAndDrop(elements.zonePhysical, elements.filePhysical, handlePhysicalFile);

  elements.tabButtons.forEach(btn => {
    btn.addEventListener('click', () => {
      elements.tabButtons.forEach(b => b.classList.remove('active'));
      btn.classList.add('active');
      state.activeFilter = btn.getAttribute('data-filter');
      renderTable();
    });
  });

  if (elements.tableSearch) {
    elements.tableSearch.addEventListener('input', (e) => {
      state.searchQuery = e.target.value;
      renderTable();
    });
  }

  if (elements.btnExport) {
    elements.btnExport.addEventListener('click', exportReport);
  }

});