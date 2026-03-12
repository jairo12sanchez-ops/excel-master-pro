// ============================================================
//   MÓDULO DE CARTERA  –  ExcelMasterPro
// ============================================================

// ---- Estado global ----
let rawData = [];   // Filas crudas del Excel (objetos con claves originales)
let carteraData = []; // Datos enriquecidos después del análisis
let filteredData = [];
let workbook;
let chartEstado = null;
let chartTop10 = null;

// ---- Elementos del DOM ----
const carteraInput = document.getElementById('cartera-input');
const dropZone = document.getElementById('drop-zone');
const welcomeScreen = document.getElementById('welcome-screen');
const dashboard = document.getElementById('dashboard');
const exportBtn = document.getElementById('export-btn');
const sheetSelect = document.getElementById('sheet-select');
const sheetSelectorContainer = document.getElementById('sheet-selector-container');
const fileBadge = document.getElementById('file-badge');
const fileNameDisplay = document.getElementById('file-name-display');

// KPIs
const kpiClientes = document.getElementById('kpi-clientes');
const kpiSaldo = document.getElementById('kpi-saldo');
const kpiMora = document.getElementById('kpi-mora');
const kpiAldia = document.getElementById('kpi-aldia');

// Columnas detectadas automáticamente
let autoColCliente = '';
let autoColDocumento = '';
let autoColSaldo = '';
let autoColVencimiento = '';
let autoColDias = '';
let autoColCategoria = '';

// Filtros
const filterMora = document.getElementById('filter-mora');
const filterCliente = document.getElementById('filter-cliente');
const filterMin = document.getElementById('filter-min');
const filterMax = document.getElementById('filter-max');
const filterBtn = document.getElementById('filter-btn');
const clearFilterBtn = document.getElementById('clear-filter-btn');

// Tabla
const tableHead = document.getElementById('table-head');
const tableBody = document.getElementById('table-body');
const tableTitle = document.getElementById('table-title');
const tableCount = document.getElementById('table-count');
const agingContainer = document.getElementById('aging-container');


// ============================================================
//  CARGA DE ARCHIVO
// ============================================================
carteraInput.addEventListener('change', (e) => {
    const file = e.target.files[0];
    if (file) { processFile(file); e.target.value = ''; }
});

dropZone.addEventListener('dragover', (e) => { e.preventDefault(); dropZone.classList.add('active'); });
dropZone.addEventListener('dragleave', () => dropZone.classList.remove('active'));
dropZone.addEventListener('drop', (e) => {
    e.preventDefault();
    dropZone.classList.remove('active');
    const file = e.dataTransfer.files[0];
    if (file) processFile(file);
});

// Click en drop-zone también abre el selector
dropZone.addEventListener('click', () => carteraInput.click());

sheetSelect.addEventListener('change', () => loadSheet(sheetSelect.value));


function processFile(file) {
    const reader = new FileReader();
    reader.onload = (e) => {
        try {
            const dataArray = new Uint8Array(e.target.result);
            workbook = XLSX.read(dataArray, { type: 'array' });

            // Poblar el selector de hojas
            sheetSelect.innerHTML = workbook.SheetNames.map(name =>
                `<option value="${name}">${name}</option>`
            ).join('');

            if (workbook.SheetNames.length > 1) {
                sheetSelectorContainer.classList.remove('hidden');
            } else {
                sheetSelectorContainer.classList.add('hidden');
            }

            // Preferimos hoja que se llame "cartera" / "clientes"
            let sheetToLoad = workbook.SheetNames[0];
            const preferido = workbook.SheetNames.findIndex(n =>
                /cartera|cliente|deudor|cobro|mora/i.test(n)
            );
            if (preferido !== -1) {
                sheetToLoad = workbook.SheetNames[preferido];
                sheetSelect.value = sheetToLoad;
            }

            loadSheet(sheetToLoad);

            // Mostrar el badge con el nombre del archivo
            fileNameDisplay.textContent = file.name;
            fileBadge.classList.remove('hidden');

        } catch (err) {
            console.error('processFile error:', err);
            alert('❌ Error al leer el archivo.\n\nDetalle: ' + (err.message || err) + '\n\nAsegúrate de que sea un Excel (.xlsx / .xls) o CSV válido.');
        }
    };
    reader.onerror = () => alert('❌ Error al leer el archivo del disco.');
    reader.readAsArrayBuffer(file);
}


function loadSheet(sheetName) {
    const ws = workbook.Sheets[sheetName];
    const matrix = XLSX.utils.sheet_to_json(ws, { header: 1, defval: '' });
    if (!matrix || matrix.length === 0) { alert('La hoja está vacía.'); return; }

    // Buscamos la fila de encabezados (primera con múltiples celdas no vacías)
    let headIdx = 0;
    for (let i = 0; i < Math.min(matrix.length, 30); i++) {
        const row = matrix[i];
        const filled = row.filter(c => String(c).trim() !== '').length;
        if (filled >= 2) { headIdx = i; break; }
    }

    const headerRow = matrix[headIdx];
    const headers = headerRow.map((h, i) =>
        String(h).trim() !== '' ? String(h).trim() : `Columna_${i + 1}`
    );

    rawData = matrix.slice(headIdx + 1)
        .filter(row => row.some(c => String(c).trim() !== ''))
        .map(row => {
            const obj = {};
            headers.forEach((h, i) => { obj[h] = row[i] !== undefined ? row[i] : ''; });
            return obj;
        });

    autoDetectColumns(headers);

    // Siempre ocultar welcome y mostrar dashboard al cargar un archivo
    welcomeScreen.classList.add('hidden');
    dashboard.classList.remove('hidden');
    // Si el panel de cliente está abierto, cerrarlo para mostrar el dashboard actualizado
    const clientPanel = document.getElementById('client-result-panel');
    if (clientPanel && !clientPanel.classList.contains('hidden')) {
        clientPanel.classList.add('hidden');
        dashboard.classList.remove('hidden');
    }
    exportBtn.disabled = false;

    // Ejecutar análisis con detección automática
    runAnalysis();
}


// ============================================================
//  DETECCIÓN AUTOMÁTICA DE COLUMNAS
// ============================================================
function normalizeStr(s) {
    return String(s || '').toLowerCase().normalize('NFD').replace(/[\u0300-\u036f]/g, '').trim();
}

function autoDetectColumns(headers) {
    const patterns = {
        cliente: /cliente|deudor|nombre|razon\s?social|empresa|tercero/i,
        documento: /nit|cedula|documento|id|identificacion|rut/i,
        saldo: /saldo|deuda|valor|monto|capital|total/i,
        vencimiento: /vencimiento|vence|fecha\s?venc|due\s?date/i,
        dias: /dias|days|mora|atraso|ven?cido|antiguedad/i,
        categoria: /categoria|estado|clasificacion|bucket|rango|segmento/i
    };

    autoColCliente = '';
    autoColDocumento = '';
    autoColSaldo = '';
    autoColVencimiento = '';
    autoColDias = '';
    autoColCategoria = '';

    headers.forEach(h => {
        if (!autoColCliente && patterns.cliente.test(h)) autoColCliente = h;
        if (!autoColDocumento && patterns.documento.test(h)) autoColDocumento = h;
        if (!autoColSaldo && patterns.saldo.test(h)) autoColSaldo = h;
        if (!autoColVencimiento && patterns.vencimiento.test(h)) autoColVencimiento = h;
        if (!autoColDias && patterns.dias.test(h)) autoColDias = h;
        if (!autoColCategoria && patterns.categoria.test(h)) autoColCategoria = h;
    });
}


// ============================================================
//  ANÁLISIS PRINCIPAL
// ============================================================
function runAnalysis() {
    if (rawData.length === 0) return;

    const today = new Date();
    today.setHours(0, 0, 0, 0);

    const cCliente = autoColCliente;
    const cDocumento = autoColDocumento;
    const cSaldo = autoColSaldo;
    const cVencimiento = autoColVencimiento;
    const cDias = autoColDias;
    const cCategoria = autoColCategoria;

    carteraData = rawData.map((row, idx) => {
        const saldo = cleanNumber(cSaldo ? row[cSaldo] : 0);

        // Calcular días de mora
        let diasMora = 0;
        if (cDias && row[cDias] !== '') {
            diasMora = cleanNumber(row[cDias]);
        } else if (cVencimiento && row[cVencimiento] !== '') {
            const fechaVenc = parseDate(row[cVencimiento]);
            if (fechaVenc) {
                const diff = today - fechaVenc;
                diasMora = diff > 0 ? Math.floor(diff / 86400000) : 0;
            }
        }

        // Clasificación por días de mora
        let estado, estadoLabel;
        if (diasMora <= 0) {
            estado = 'aldia'; estadoLabel = 'Al Día';
        } else if (diasMora <= 30) {
            estado = 'reciente'; estadoLabel = '1-30 días';
        } else if (diasMora <= 60) {
            estado = 'mora30'; estadoLabel = '31-60 días';
        } else if (diasMora <= 90) {
            estado = 'mora60'; estadoLabel = '61-90 días';
        } else if (diasMora <= 180) {
            estado = 'mora90'; estadoLabel = '91-180 días';
        } else {
            estado = 'critica'; estadoLabel = '+180 días';
        }

        return {
            _idx: idx,
            _cliente: cCliente ? String(row[cCliente] || '').trim() : `Registro ${idx + 1}`,
            _documento: cDocumento ? String(row[cDocumento] || '').trim() : '',
            _saldo: saldo,
            _dias: diasMora,
            _estado: estado,
            _estadoLabel: estadoLabel,
            _categoria: cCategoria ? String(row[cCategoria] || '').trim() : '',
            ...row
        };
    });

    filteredData = [...carteraData];
    renderAll();
    // Poblar selector y sugerencias de búsqueda por cliente
    if (typeof populateColBusqueda === 'function') populateColBusqueda();
}


// ============================================================
//  RENDER COMPLETO
// ============================================================
function renderAll() {
    updateKPIs();
    renderAgingBuckets();
    renderCharts();
    renderTable();
}


// ============================================================
//  KPIs
// ============================================================
function updateKPIs() {
    const totalClientes = new Set(filteredData.map(r => r._cliente || r._idx)).size;
    const totalSaldo = filteredData.reduce((a, r) => a + r._saldo, 0);
    const enMora = filteredData.filter(r => r._estado !== 'aldia').length;
    const alDia = filteredData.filter(r => r._estado === 'aldia').length;

    kpiClientes.textContent = totalClientes.toLocaleString();
    kpiSaldo.textContent = '$ ' + formatNumber(totalSaldo);
    kpiMora.textContent = enMora.toLocaleString();
    kpiAldia.textContent = alDia.toLocaleString();
}


// ============================================================
//  AGING / DISTRIBUCIÓN POR MORA
// ============================================================
function renderAgingBuckets() {
    const buckets = [
        { key: 'aldia', label: 'Al Día', color: '#10b981' },
        { key: 'reciente', label: '1-30 días', color: '#fbbf24' },
        { key: 'mora30', label: '31-60 días', color: '#f97316' },
        { key: 'mora60', label: '61-90 días', color: '#ef4444' },
        { key: 'mora90', label: '91-180 días', color: '#dc2626' },
        { key: 'critica', label: '+180 días', color: '#7f1d1d' }
    ];

    const saldoTotal = filteredData.reduce((a, r) => a + r._saldo, 0) || 1;

    agingContainer.innerHTML = '';
    buckets.forEach(b => {
        const rows = filteredData.filter(r => r._estado === b.key);
        const total = rows.reduce((a, r) => a + r._saldo, 0);
        const pct = Math.round((total / saldoTotal) * 100);

        const item = document.createElement('div');
        item.className = 'aging-item';
        item.innerHTML = `
            <div class="aging-header">
                <span class="aging-label">${b.label} (${rows.length})</span>
                <span class="aging-amount">$ ${formatNumber(total)}</span>
            </div>
            <div class="aging-bar-bg">
                <div class="aging-bar-fill" style="width:${pct}%; background:${b.color}"></div>
            </div>
        `;
        agingContainer.appendChild(item);
    });
}


// ============================================================
//  GRÁFICOS
// ============================================================
function renderCharts() {
    renderChartEstado();
    renderChartTop10();
}

function renderChartEstado() {
    const ctx = document.getElementById('chart-estado').getContext('2d');
    if (chartEstado) chartEstado.destroy();

    const labels = ['Al Día', '1-30 días', '31-60 días', '61-90 días', '91-180 días', '+180 días'];
    const keys = ['aldia', 'reciente', 'mora30', 'mora60', 'mora90', 'critica'];
    const colors = ['#10b981', '#fbbf24', '#f97316', '#ef4444', '#dc2626', '#7f1d1d'];

    const dataVals = keys.map(k =>
        filteredData.filter(r => r._estado === k).reduce((a, r) => a + r._saldo, 0)
    );

    chartEstado = new Chart(ctx, {
        type: 'doughnut',
        data: {
            labels,
            datasets: [{
                data: dataVals,
                backgroundColor: colors,
                borderColor: '#1e293b',
                borderWidth: 3,
                hoverOffset: 8
            }]
        },
        options: {
            responsive: true,
            maintainAspectRatio: false,
            plugins: {
                legend: {
                    position: 'right',
                    labels: { color: '#94a3b8', font: { size: 11 }, boxWidth: 14, padding: 10 }
                },
                tooltip: {
                    callbacks: {
                        label: (ctx) => ` $ ${formatNumber(ctx.parsed)} (${ctx.label})`
                    }
                }
            }
        }
    });
}

function renderChartTop10() {
    const ctx = document.getElementById('chart-top10').getContext('2d');
    if (chartTop10) chartTop10.destroy();

    // Agrupa por cliente
    const clienteSaldos = {};
    filteredData.forEach(r => {
        const k = r._cliente || 'Sin nombre';
        clienteSaldos[k] = (clienteSaldos[k] || 0) + r._saldo;
    });

    const sorted = Object.entries(clienteSaldos)
        .sort((a, b) => b[1] - a[1])
        .slice(0, 10);

    const labels = sorted.map(([k]) => k.length > 20 ? k.slice(0, 19) + '…' : k);
    const dataVals = sorted.map(([, v]) => v);

    chartTop10 = new Chart(ctx, {
        type: 'bar',
        data: {
            labels,
            datasets: [{
                label: 'Saldo ($)',
                data: dataVals,
                backgroundColor: 'rgba(99, 102, 241, 0.7)',
                borderColor: '#6366f1',
                borderWidth: 1,
                borderRadius: 6
            }]
        },
        options: {
            indexAxis: 'y',
            responsive: true,
            maintainAspectRatio: false,
            plugins: {
                legend: { display: false },
                tooltip: {
                    callbacks: {
                        label: (ctx) => ` $ ${formatNumber(ctx.parsed.x)}`
                    }
                }
            },
            scales: {
                x: {
                    ticks: {
                        color: '#94a3b8', font: { size: 10 },
                        callback: v => '$ ' + (v >= 1e6 ? (v / 1e6).toFixed(1) + 'M' : formatNumber(v))
                    },
                    grid: { color: 'rgba(255,255,255,0.05)' }
                },
                y: {
                    ticks: { color: '#f8fafc', font: { size: 10 } },
                    grid: { display: false }
                }
            }
        }
    });
}


// ============================================================
//  TABLA
// ============================================================
function renderTable() {
    if (filteredData.length === 0) {
        tableHead.innerHTML = '';
        tableBody.innerHTML = '<tr><td style="text-align:center; padding:2rem; color:#94a3b8">No hay registros que coincidan con los filtros</td></tr>';
        tableCount.textContent = '0 registros';
        return;
    }

    tableCount.textContent = `${filteredData.length.toLocaleString()} registros`;

    // Determinar columnas a mostrar: primero las detectadas automáticamente, luego el resto
    const mappedCols = [
        autoColCliente, autoColDocumento, autoColSaldo,
        autoColDias, autoColVencimiento, autoColCategoria
    ].filter(c => c !== '');

    const allOriginalCols = rawData.length > 0 ? Object.keys(rawData[0]) : [];
    const otherCols = allOriginalCols.filter(c => !mappedCols.includes(c));
    const displayCols = [...new Set([...mappedCols, ...otherCols])];

    // Encabezado con columna extra de estado
    tableHead.innerHTML = `<tr>
        <th>Estado</th>
        ${displayCols.map(c => `<th>${c}</th>`).join('')}
    </tr>`;

    const maxRows = 500;
    const rows = filteredData.slice(0, maxRows);

    tableBody.innerHTML = rows.map(row => {
        const badgeClass = {
            aldia: 'badge-aldia',
            reciente: 'badge-reciente',
            mora30: 'badge-mora',
            mora60: 'badge-mora',
            mora90: 'badge-critica',
            critica: 'badge-critica'
        }[row._estado] || 'badge-mora';

        const cells = displayCols.map(col => {
            const val = row[col];
            if (col === autoColSaldo) {
                const num = cleanNumber(val);
                const cls = num > 0 ? 'amount-positive' : 'amount-zero';
                return `<td class="amount-cell ${cls}">$ ${formatNumber(num)}</td>`;
            }
            return `<td>${val !== undefined && val !== '' ? val : '—'}</td>`;
        }).join('');

        return `<tr>
            <td><span class="badge ${badgeClass}">${row._estadoLabel}</span></td>
            ${cells}
        </tr>`;
    }).join('');

    if (filteredData.length > maxRows) {
        tableBody.innerHTML += `<tr><td colspan="${displayCols.length + 1}" style="text-align:center; padding:1rem; color:#94a3b8; font-size:0.8rem">
            Mostrando ${maxRows.toLocaleString()} de ${filteredData.length.toLocaleString()} registros. Exporta para ver todos.
        </td></tr>`;
    }
}


// ============================================================
//  FILTROS
// ============================================================
filterBtn.addEventListener('click', applyFilters);
clearFilterBtn.addEventListener('click', clearFilters);
filterCliente.addEventListener('input', applyFilters);
filterMora.addEventListener('change', applyFilters);

function applyFilters() {
    const moraVal = filterMora.value;
    const clienteVal = filterCliente.value.trim().toLowerCase();
    const minVal = filterMin.value !== '' ? parseFloat(filterMin.value) : null;
    const maxVal = filterMax.value !== '' ? parseFloat(filterMax.value) : null;

    filteredData = carteraData.filter(row => {
        // Filtro mora
        if (moraVal === 'mora' && row._estado === 'aldia') return false;
        if (moraVal === 'aldia' && row._estado !== 'aldia') return false;

        // Filtro cliente
        if (clienteVal) {
            const clienteStr = (row._cliente + ' ' + row._documento).toLowerCase();
            if (!clienteStr.includes(clienteVal)) return false;
        }

        // Filtro saldo
        if (minVal !== null && row._saldo < minVal) return false;
        if (maxVal !== null && row._saldo > maxVal) return false;

        return true;
    });

    renderAll();
}

function clearFilters() {
    filterMora.value = 'todos';
    filterCliente.value = '';
    filterMin.value = '';
    filterMax.value = '';
    filteredData = [...carteraData];
    renderAll();
}


// ============================================================
//  EXPORTAR
// ============================================================
exportBtn.addEventListener('click', () => {
    if (filteredData.length === 0) return;

    try {
        const exportRows = filteredData.map(row => {
            const obj = { 'Estado Mora': row._estadoLabel };
            // Solo columnas originales
            Object.keys(rawData[0] || {}).forEach(k => {
                obj[k] = row[k];
            });
            return obj;
        });

        const ws = XLSX.utils.json_to_sheet(exportRows);
        const wb = XLSX.utils.book_new();
        XLSX.book_append_sheet(wb, ws, 'Cartera');
        XLSX.writeFile(wb, 'Reporte_Cartera.xlsx');
    } catch (err) {
        console.error(err);
        alert('Error al exportar. Intente de nuevo.');
    }
});


// ============================================================
//  UTILIDADES
// ============================================================
function cleanNumber(val) {
    if (typeof val === 'number') return val;
    if (!val && val !== 0) return 0;
    let s = String(val).trim().replace(/\s/g, '').replace(/[$,%]/g, '');
    if (s.includes('.') && s.includes(',')) s = s.replace(/\./g, '').replace(',', '.');
    else if (s.includes(',')) s = s.replace(',', '.');
    const n = parseFloat(s);
    return isNaN(n) ? 0 : n;
}

function formatNumber(n) {
    if (isNaN(n)) return '0';
    return Math.round(n).toLocaleString('es-CO');
}

function parseDate(val) {
    if (!val) return null;
    if (typeof val === 'number') {
        // Número serial de Excel
        const d = XLSX.SSF.parse_date_code(val);
        if (d) return new Date(d.y, d.m - 1, d.d);
    }
    const s = String(val).trim();
    // Formatos: dd/mm/yyyy, yyyy-mm-dd, mm/dd/yyyy, dd-mm-yyyy
    const patterns = [
        /^(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{4})$/,
        /^(\d{4})[\/\-](\d{1,2})[\/\-](\d{1,2})$/
    ];
    for (const p of patterns) {
        const m = s.match(p);
        if (m) {
            const d1 = new Date(`${m[3] || m[1]}-${String(m[2]).padStart(2, '0')}-${String(m[3] ? m[3] : m[1]).padStart(2, '0')}`);
            if (!isNaN(d1)) return d1;
        }
    }
    const d = new Date(s);
    return isNaN(d) ? null : d;
}


// ============================================================
//  BÚSQUEDA POR CÓDIGO DE CLIENTE (por columna seleccionada)
// ============================================================
const colBusqueda = document.getElementById('col-busqueda');
const clientCodeInput = document.getElementById('client-code-input');
const clientCodesDatalist = document.getElementById('client-codes-datalist');
const clientSearchBtn = document.getElementById('client-search-btn');
const clientClearBtn = document.getElementById('client-clear-btn');
const clientResultPanel = document.getElementById('client-result-panel');
const closeClientPanel = document.getElementById('close-client-panel');

// Referencias del panel de resultado
const clientAvatar = document.getElementById('client-avatar');
const clientResultName = document.getElementById('client-result-name');
const clientResultCode = document.getElementById('client-result-code');
const ckRegistros = document.getElementById('ck-registros');
const ckDeuda = document.getElementById('ck-deuda');
const ckDias = document.getElementById('ck-dias');
const ckEstado = document.getElementById('ck-estado');
const clientTableHead = document.getElementById('client-table-head');
const clientTableBody = document.getElementById('client-table-body');
const clientTableCount = document.getElementById('client-table-count');

// Buscar al presionar Enter o botón
clientCodeInput.addEventListener('keydown', (e) => {
    if (e.key === 'Enter') searchByClientCode();
});
clientSearchBtn.addEventListener('click', searchByClientCode);
clientClearBtn.addEventListener('click', closeClientResultPanel);
closeClientPanel.addEventListener('click', closeClientResultPanel);

// También buscar al seleccionar del datalist
clientCodeInput.addEventListener('change', () => {
    if (clientCodeInput.value.trim()) searchByClientCode();
});


/**
 * Llena el selector de columna de búsqueda con todas las columnas del Excel.
 * Por defecto deja seleccionada la primera columna (A).
 */
function populateColBusqueda() {
    if (rawData.length === 0) return;
    const cols = Object.keys(rawData[0]);
    const firstCol = cols[0] || '';

    colBusqueda.innerHTML = cols
        .map((c, i) => {
            const label = i === 0 ? `${c}  ← Columna A` : c;
            return `<option value="${c}">${label}</option>`;
        })
        .join('');

    // Seleccionar columna A por defecto
    colBusqueda.value = firstCol;

    // Cada vez que cambie la columna, actualizar el datalist
    colBusqueda.onchange = populateClientDatalist;

    populateClientDatalist();
}

/**
 * Rellena el datalist con los valores únicos de la columna seleccionada,
 * para que el usuario pueda autocompletar con los códigos reales del Excel.
 */
function populateClientDatalist() {
    if (carteraData.length === 0) return;

    const col = colBusqueda.value;
    if (!col) return;

    const values = new Set();
    carteraData.forEach(row => {
        const val = String(row[col] ?? '').trim();
        if (val !== '') values.add(val);
    });

    clientCodesDatalist.innerHTML = [...values]
        .sort((a, b) => String(a).localeCompare(String(b)))
        .map(v => `<option value="${v}">`)
        .join('');

    // Actualizar placeholder con el nombre de la columna
    const colNombre = col;
    clientCodeInput.placeholder = `Buscar en "${colNombre}"...`;
}


/**
 * Ejecuta la búsqueda usando la columna elegida en el selector.
 * Busca coincidencia exacta primero; si no hay, busca parcial (contiene).
 */
function searchByClientCode() {
    const query = clientCodeInput.value.trim();
    if (!query) return;

    if (carteraData.length === 0) {
        alert('Primero sube un archivo de cartera.');
        return;
    }

    const col = colBusqueda.value;
    if (!col) {
        alert('Selecciona primero la columna de búsqueda.');
        return;
    }

    const q = query.trim().toLowerCase();

    // 1) Búsqueda exacta (case-insensitive)
    let results = carteraData.filter(row =>
        String(row[col] ?? '').trim().toLowerCase() === q
    );

    // 2) Si no hay exactos, búsqueda parcial (contiene)
    if (results.length === 0) {
        results = carteraData.filter(row =>
            String(row[col] ?? '').trim().toLowerCase().includes(q)
        );
    }

    if (results.length === 0) {
        alert(`❌ No se encontró ningún registro con "${query}" en la columna "${col}"`);
        return;
    }

    // Ocultar dashboard y mostrar panel del cliente
    dashboard.classList.add('hidden');
    clientResultPanel.classList.remove('hidden');
    clientClearBtn.classList.remove('hidden');

    renderClientPanel(query, results, col);
}


/**
 * Rellena el panel con los datos del cliente encontrado.
 */
/**
 * Renderiza el panel de resultado para el cliente encontrado.
 * @param {string} query - Valor buscado
 * @param {Array}  results - Filas coincidentes
 * @param {string} col - Columna en la que se buscó
 */
function renderClientPanel(query, results, col) {
    // Nombre y código del cliente
    // Usamos el valor real de la columna buscada como identificador principal
    const primerRegistro = results[0];
    const colUsada = col || colBusqueda.value;
    const codigoCliente = String(primerRegistro[colUsada] ?? query).trim();
    // Nombre: la columna mapeada de cliente, o el propio código si no hay mapeo
    const nombreCliente = primerRegistro._cliente && primerRegistro._cliente !== codigoCliente
        ? primerRegistro._cliente
        : codigoCliente;

    clientResultName.textContent = nombreCliente;
    clientResultCode.textContent = codigoCliente;

    // Avatar: primera letra del nombre
    const letra = nombreCliente.charAt(0).toUpperCase();
    clientAvatar.textContent = /[A-Z0-9]/.test(letra) ? letra : '👤';

    // KPIs del cliente
    const totalRegistros = results.length;
    const totalDeuda = results.reduce((a, r) => a + r._saldo, 0);
    const maxDias = Math.max(...results.map(r => r._dias || 0));

    // Estado general: si tiene alguno en mora, el peor
    const prioridad = { critica: 5, mora90: 4, mora60: 3, mora30: 2, reciente: 1, aldia: 0 };
    const peorEstado = results.reduce((prev, r) =>
        (prioridad[r._estado] || 0) > (prioridad[prev._estado] || 0) ? r : prev
        , results[0]);

    const estadoBadges = {
        aldia: { label: 'AL DÍA', color: '#34d399' },
        reciente: { label: '1-30 DÍAS', color: '#fbbf24' },
        mora30: { label: '31-60 DÍAS', color: '#f97316' },
        mora60: { label: '61-90 DÍAS', color: '#ef4444' },
        mora90: { label: '91-180 DÍAS', color: '#fca5a5' },
        critica: { label: 'CRÍTICA', color: '#ef4444' }
    };
    const estadoInfo = estadoBadges[peorEstado._estado] || { label: 'DESCONOCIDO', color: '#94a3b8' };

    ckRegistros.textContent = totalRegistros.toLocaleString();
    ckDeuda.textContent = '$ ' + formatNumber(totalDeuda);
    ckDias.textContent = maxDias > 0 ? maxDias + ' días' : '0';
    ckEstado.style.color = estadoInfo.color;
    ckEstado.textContent = estadoInfo.label;

    // Tabla del cliente
    clientTableCount.textContent = `${totalRegistros} registro${totalRegistros !== 1 ? 's' : ''}`;

    const mappedCols = [
        autoColCliente, autoColDocumento, autoColSaldo,
        autoColDias, autoColVencimiento, autoColCategoria
    ].filter(c => c !== '');

    const allOriginalCols = rawData.length > 0 ? Object.keys(rawData[0]) : [];
    const otherCols = allOriginalCols.filter(c => !mappedCols.includes(c));
    const displayCols = [...new Set([...mappedCols, ...otherCols])];

    // Encabezado
    clientTableHead.innerHTML = `<tr>
        <th>Estado</th>
        <th>Días Mora</th>
        ${displayCols.map(c => `<th>${c}</th>`).join('')}
    </tr>`;

    // Filas — ordenadas por días de mora desc
    const sorted = [...results].sort((a, b) => b._dias - a._dias);

    clientTableBody.innerHTML = sorted.map(row => {
        const badgeClass = {
            aldia: 'badge-aldia',
            reciente: 'badge-reciente',
            mora30: 'badge-mora',
            mora60: 'badge-mora',
            mora90: 'badge-critica',
            critica: 'badge-critica'
        }[row._estado] || 'badge-mora';

        const cells = displayCols.map(col => {
            const val = row[col];
            if (col === autoColSaldo) {
                const num = cleanNumber(val);
                const cls = num > 0 ? 'amount-positive' : 'amount-zero';
                return `<td class="amount-cell ${cls}">$ ${formatNumber(num)}</td>`;
            }
            return `<td>${val !== undefined && val !== '' ? val : '—'}</td>`;
        }).join('');

        const diasCell = row._dias > 0
            ? `<td style="color:#f87171; font-weight:700;">${row._dias}</td>`
            : `<td style="color:#34d399; font-weight:700;">0</td>`;

        return `<tr>
            <td><span class="badge ${badgeClass}">${row._estadoLabel}</span></td>
            ${diasCell}
            ${cells}
        </tr>`;
    }).join('');
}


/**
 * Cierra el panel del cliente y vuelve al dashboard general.
 */
function closeClientResultPanel() {
    clientResultPanel.classList.add('hidden');
    dashboard.classList.remove('hidden');
    clientCodeInput.value = '';
    clientClearBtn.classList.add('hidden');
}


