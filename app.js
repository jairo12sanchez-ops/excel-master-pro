let workbook;
let worksheet;
let data = [];

const excelInput = document.getElementById('excel-input');
const welcomeScreen = document.getElementById('welcome-screen');
const tableContainer = document.getElementById('table-container');
const tableHead = document.getElementById('table-head');
const tableBody = document.getElementById('table-body');
const exportBtn = document.getElementById('export-btn');
const sumsContainer = document.getElementById('sums-container');
const grandTotalValue = document.getElementById('grand-total-value');
const selectColCodigo = document.getElementById('select-col-codigo');
const selectColTotal = document.getElementById('select-col-total');
const selectColCantidad = document.getElementById('select-col-cantidad');
const selectColCosto = document.getElementById('select-col-costo');
const sheetSelect = document.getElementById('sheet-select');
const sheetSelectorContainer = document.getElementById('sheet-selector-container');

// Stats elements
const statRows = document.getElementById('stat-rows');
const statCols = document.getElementById('stat-cols');
const statCosts = document.getElementById('stat-costs');

// Product Costs Logic
let productCosts = JSON.parse(localStorage.getItem('productCosts') || '{}');
let productCostsHistory = JSON.parse(localStorage.getItem('productCostsHistory') || '{}');

// Nuevo: Intentar cargar presupuesto maestro desde el servidor
async function loadDefaultCosts() {
    try {
        const response = await fetch('presupuesto_maestro.xlsx');
        if (response.ok) {
            const blob = await response.blob();
            const file = new File([blob], "presupuesto_maestro.xlsx", { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" });
            
            console.log("Cargando presupuesto maestro detectado en el servidor...");
            
            // 1. Cargarlo como base de datos de costos (Diccionario)
            processCostsFile(file);
            
            // 2. Cargarlo como archivo principal para mostrar en la tabla y habilitar el selector de hojas
            processFile(file);
        }
    } catch (error) {
        console.log("No se encontró presupuesto maestro inicial en el servidor.");
    }
}

loadDefaultCosts();
updateCostsStat();

const costsInput = document.getElementById('costs-input');
costsInput.addEventListener('change', (e) => {
    const file = e.target.files[0];
    if (file) {
        processCostsFile(file);
        e.target.value = ''; // Reset para permitir subir el mismo archivo
    }
});

const validateBtn = document.getElementById('validate-btn');
validateBtn.addEventListener('click', () => {
    if (data.length === 0) {
        alert("Primero sube un archivo de presupuesto.");
        return;
    }
    validateBudget(true);
});

// Handle File Selection (also reset)
excelInput.addEventListener('change', (e) => {
    const file = e.target.files[0];
    if (file) {
        processFile(file);
        e.target.value = '';
    }
});

// Drag & Drop
const dropZone = document.getElementById('drop-zone');
dropZone.addEventListener('dragover', (e) => {
    e.preventDefault();
    dropZone.classList.add('active');
});

dropZone.addEventListener('dragleave', () => {
    dropZone.classList.remove('active');
});

dropZone.addEventListener('drop', (e) => {
    e.preventDefault();
    const file = e.dataTransfer.files[0];
    if (file) {
        processFile(file);
        e.target.value = '';
    }
});

// Eventos de cambio en selectores
selectColCodigo.addEventListener('change', generateCodeSummary);
selectColTotal.addEventListener('change', generateCodeSummary);
sheetSelect.addEventListener('change', () => {
    loadSheet(sheetSelect.value);
});

const configBtn = document.getElementById('config-btn');
const columnConfigPanel = document.getElementById('column-config-panel');
configBtn.addEventListener('click', () => {
    columnConfigPanel.classList.toggle('hidden');
});

function processFile(file) {
    const reader = new FileReader();
    reader.onload = (e) => {
        const dataArray = new Uint8Array(e.target.result);
        workbook = XLSX.read(dataArray, { type: 'array' });

        sheetSelect.innerHTML = workbook.SheetNames.map(name => `<option value="${name}">${name}</option>`).join('');

        let sheetToLoad = workbook.SheetNames[0];
        const indexValidacion = workbook.SheetNames.findIndex(n => n.toUpperCase().includes('VALIDACION') || n.toUpperCase().includes('PRESUPUESTO'));

        if (indexValidacion !== -1) {
            sheetToLoad = workbook.SheetNames[indexValidacion];
            sheetSelect.value = sheetToLoad;
        }

        if (workbook.SheetNames.length > 1) {
            sheetSelectorContainer.classList.remove('hidden');
        } else {
            sheetSelectorContainer.classList.add('hidden');
        }

        loadSheet(sheetToLoad);

        welcomeScreen.classList.add('hidden');
        tableContainer.classList.remove('hidden');
        exportBtn.disabled = false;
    };
    reader.readAsArrayBuffer(file);
}

function processCostsFile(file) {
    // Feedback inmediato
    console.log("Iniciando lectura de:", file.name);

    const reader = new FileReader();

    reader.onloadstart = () => {
        // Podríamos poner un spinner, pero un log basta por ahora
        console.log("Lectura iniciada...");
    };

    reader.onload = (e) => {
        try {
            console.log("Archivo leído correctamente, procesando con SheetJS...");
            const dataArray = new Uint8Array(e.target.result);
            const tempWb = XLSX.read(dataArray, { type: 'array' });

            if (!tempWb || !tempWb.SheetNames || tempWb.SheetNames.length === 0) {
                alert("❌ El archivo no parece ser un Excel válido o está protegido.");
                return;
            }

            console.log("Hojas disponibles:", tempWb.SheetNames);
            
            // Prioridad estricta para encontrar la hoja de validación primero
            const prioritySheets = ['VALIDACION', 'PRESUPUESTO', 'INFORME', 'MAESTRO', 'BASE'];
            let firstSheetName = tempWb.SheetNames[0];
            for (const p of prioritySheets) {
                const found = tempWb.SheetNames.find(n => n.toUpperCase().includes(p));
                if (found) {
                    firstSheetName = found;
                    break;
                }
            }
            
            console.log("Cargando costos desde hoja:", firstSheetName);
            const firstSheet = tempWb.Sheets[firstSheetName];
            const matrix = XLSX.utils.sheet_to_json(firstSheet, { header: 1, defval: "" });

            if (!matrix || matrix.length === 0) {
                alert("⚠️ El archivo o la hoja seleccionada están vacíos.");
                return;
            }

            let headIdx = -1;
            let codeColIdx = -1;
            let costColIdx = -1;

            // Búsqueda exhaustiva de cabeceras (primeras 200 filas)
            for (let i = 0; i < Math.min(matrix.length, 200); i++) {
                const row = matrix[i];
                if (!row) continue;

                const cIdx = row.findIndex(cell => {
                    const s = String(cell || "").toLowerCase().normalize("NFD").replace(/[\u0300-\u036f]/g, "");
                    return s.includes('codigo') || s.includes('referencia') || s.includes('ref') || s.includes('prod') || s.includes('articulo') || s.includes('item');
                });

                const vIdx = row.findIndex(cell => {
                    const s = String(cell || "").toLowerCase().normalize("NFD").replace(/[\u0300-\u036f]/g, "");
                    return s.includes('costo') || s.includes('precio') || s.includes('valor') || s.includes('unitario') || s.includes('$') || s.includes('vlr');
                });

                if (cIdx !== -1 && vIdx !== -1) {
                    headIdx = i;
                    codeColIdx = cIdx;
                    costColIdx = vIdx;
                    break;
                }
            }

            // Fallback manual para Columnas M (12) y T (19) si la búsqueda automática falla pero hay datos
            if (headIdx === -1) {
                for (let i = 0; i < Math.min(matrix.length, 50); i++) {
                    const row = matrix[i];
                    if (row && row[12] && row[19]) {
                        headIdx = i;
                        codeColIdx = 12; // Columna M
                        costColIdx = 19; // Columna T
                        break;
                    }
                }
            }

            if (headIdx === -1) {
                alert(`❌ No logré identificar las columnas en la hoja: "${firstSheetName}"\n\nAsegúrate de tener títulos como 'Codigo' y 'Costo'.`);
                return;
            }

            let count = 0;
            let updatedProductCosts = { ...productCosts }; // Mantener los existentes por ahora

            for (let i = headIdx + 1; i < matrix.length; i++) {
                const row = matrix[i];
                if (!row) continue;

                const rawCode = String(row[codeColIdx] || "").trim();
                const code = rawCode.toUpperCase();
                const costValue = row[costColIdx];
                const cost = cleanNumber(costValue);

                if (code && !isNaN(cost) && costValue !== "") {
                    updatedProductCosts[code] = cost;

                    // Actualizar historial
                    if (!productCostsHistory[code]) productCostsHistory[code] = [];

                    // Solo agregar si el costo cambió o si es la primera vez en este proceso
                    const lastHistory = productCostsHistory[code][productCostsHistory[code].length - 1];
                    if (!lastHistory || lastHistory.cost !== cost) {
                        productCostsHistory[code].push({
                            date: new Date().toISOString(),
                            cost: cost,
                            file: file.name
                        });
                    }

                    count++;
                }
            }

            if (count > 0) {
                productCosts = updatedProductCosts;
                localStorage.setItem('productCosts', JSON.stringify(productCosts));
                localStorage.setItem('productCostsHistory', JSON.stringify(productCostsHistory));
                updateCostsStat();
                alert(`✅ ¡Éxito! Se han procesado ${count} productos correctamente.`);
                if (data.length > 0) generateCodeSummary();
            } else {
                alert("❌ Se encontró la cabecera pero no pude leer ningún dato válido debajo.");
            }
        } catch (error) {
            console.error("Error detallado:", error);
            alert("Error al abrir el archivo. Asegúrate de que sea un archivo Excel (.xlsx o .xls) válido.");
        }
    };
    reader.onerror = (err) => alert("Error al leer el archivo del disco.");
    reader.readAsArrayBuffer(file);
}

function updateCostsStat() {
    if (statCosts) {
        statCosts.innerText = Object.keys(productCosts).length;
    }
}

function loadSheet(sheetName) {
    worksheet = workbook.Sheets[sheetName];

    // 1. Leemos como matriz cruda
    const rawMatrix = XLSX.utils.sheet_to_json(worksheet, { header: 1, defval: "" });

    if (rawMatrix.length > 0) {
        // 2. Encontramos el ancho máximo real
        const maxColsFound = Math.max(...rawMatrix.map(row => row ? row.length : 0));

        // 3. Identificamos columnas que NO están vacías (que tienen algo en alguna fila)
        const activeColIndices = [];
        for (let colIdx = 0; colIdx < maxColsFound; colIdx++) {
            const hasData = rawMatrix.some(row => {
                const val = row[colIdx];
                return val !== undefined && val !== null && String(val).trim() !== "";
            });
            if (hasData) {
                activeColIndices.push(colIdx);
            }
        }

        // 4. Buscamos la fila de títulos (scaneamos hasta la 50)
        let headIdx = 0;
        for (let i = 0; i < Math.min(rawMatrix.length, 50); i++) {
            const rowArr = rawMatrix[i] || [];
            const rowStr = rowArr.join('|').toUpperCase();

            // Buscamos combinaciones que indiquen cabecera, no solo 'ATO' que puede estar en los datos
            const hasHeaderWords = rowStr.includes('CODIGO') || rowStr.includes('PRODUCTO') || rowStr.includes('ARTICULO') || rowStr.includes('COSTO') || (rowStr.includes('TOTAL') && rowStr.includes('ATO'));

            if (hasHeaderWords) {
                headIdx = i;
                break;
            }
        }

        // 5. Generamos encabezados solo para columnas activas
        const activeHeaders = [];
        const headerRow = rawMatrix[headIdx] || [];

        activeColIndices.forEach(idx => {
            const colLetter = getExcelColumnName(idx);
            const rawTitle = headerRow[idx];
            const cleanTitle = (rawTitle && String(rawTitle).trim()) ? String(rawTitle).trim() : `Columna ${colLetter}`;
            activeHeaders.push({
                fullName: `${colLetter} - ${cleanTitle}`,
                originalIndex: idx
            });
        });

        // 6. Mapeamos los datos filtrados
        data = rawMatrix.slice(headIdx + 1).map(rowArr => {
            const obj = {};
            activeHeaders.forEach(h => {
                obj[h.fullName] = (rowArr[h.originalIndex] !== undefined) ? rowArr[h.originalIndex] : "";
            });
            return obj;
        });
    }

    populateSelectors();
    displayData();
    calculateSums();
    generateCodeSummary();
    updateStats();

    // Validación de presupuesto automática si la hoja tiene un nombre clave
    const isBudgetSheet = sheetName.toUpperCase().includes('VALIDACION') || sheetName.toUpperCase().includes('PRESUPUESTO');
    if (isBudgetSheet) {
        // En modo automático mostramos éxito solo si el usuario acaba de subir el archivo
        // para dar feedback de que la validación se ejecutó.
        validateBudget(true);
    }
}

function validateBudget(showSuccess = false) {
    if (data.length === 0) return;

    const colCodigo = selectColCodigo.value;
    const colCostoArchivo = selectColCosto.value;

    if (!colCodigo) {
        if (showSuccess) alert("Seleccione la columna de Código primero.");
        return;
    }

    const missingCosts = new Set();
    const costDifferences = [];

    data.forEach(row => {
        const code = String(row[colCodigo]).trim().toUpperCase();
        if (!code || code === colCodigo.toUpperCase() || code.includes('TOTAL')) return;

        const costHistory = productCosts[code];

        if (costHistory === undefined) {
            missingCosts.add(code);
        } else if (colCostoArchivo) {
            const costInFile = cleanNumber(row[colCostoArchivo]);
            // Solo validamos si el archivo trae un costo mayor a 0
            if (costInFile > 0 && Math.abs(costInFile - costHistory) > 1) { // Tolerancia de 1 unidad
                costDifferences.push({
                    code,
                    history: costHistory,
                    file: costInFile
                });
            }
        }
    });

    let msg = "";
    if (missingCosts.size > 0) {
        const missingList = Array.from(missingCosts).sort();
        msg += `⚠️ FALTAN EN HISTORIAL (${missingCosts.size}):\n${missingList.slice(0, 10).join(', ')}${missingList.length > 10 ? '...' : ''}\n\n`;
    }

    if (costDifferences.length > 0) {
        // Eliminar duplicados de código en la lista de discrepancias
        const uniqueDiffs = {};
        costDifferences.forEach(d => { if (!uniqueDiffs[d.code]) uniqueDiffs[d.code] = d; });
        const diffList = Object.values(uniqueDiffs).sort((a, b) => a.code.localeCompare(b.code));

        msg += `❌ DISCREPANCIAS DE COSTO (${diffList.length}):\nEl costo del presupuesto NO coincide con el historial.\n\n`;
        diffList.slice(0, 10).forEach(d => {
            msg += `- ${d.code}: Hist $${d.history.toLocaleString()} | Arch $${d.file.toLocaleString()}\n`;
        });
        if (diffList.length > 10) msg += `...y ${diffList.length - 10} productos más.\n`;
    }

    if (msg) {
        alert("🔍 RESULTADOS DE VALIDACIÓN\n\n" + msg);
    } else if (showSuccess) {
        alert("✅ Validación exitosa:\n\nTodos los productos tienen costo y coinciden exactamente con el historial.");
    }
}

function displayData() {
    if (data.length === 0) return;
    const headers = Object.keys(data[0]);
    tableHead.innerHTML = `<tr>${headers.map(h => `<th>${h}</th>`).join('')}</tr>`;
    // Aumentamos el límite a 1000 filas para que el usuario vea sus datos (tiene ~404 según capturas)
    tableBody.innerHTML = data.slice(0, 1000).map(row => `
        <tr>
            ${headers.map(header => `<td contenteditable="true" onblur="updateCell(this, '${header}')">${row[header] || ''}</td>`).join('')}
        </tr>
    `).join('');

    if (data.length > 1000) {
        tableBody.innerHTML += `<tr><td colspan="${headers.length}" style="text-align:center; padding: 20px; color: var(--text-dim)">Mostrando las primeras 1000 filas. El cálculo y la exportación incluyen las ${data.length} filas totales.</td></tr>`;
    }
}

function updateCell(cell, header) {
    const rowIndex = cell.parentElement.rowIndex - 1;
    data[rowIndex][header] = cell.innerText;
    // calculateSums(); // Desactivado por solicitud del usuario
    generateCodeSummary();
}

function cleanNumber(val) {
    if (typeof val === 'number') return val;
    if (!val) return 0;
    let s = String(val).trim().replace(/\s/g, '');

    // Tratamiento de formatos regionales (puntos y comas)
    if (s.includes('.') && s.includes(',')) {
        // Estilo 1.234,56 -> 1234.56
        s = s.replace(/\./g, '').replace(',', '.');
    } else if (s.includes(',')) {
        // Estilo 1234,56 -> 1234.56
        s = s.replace(',', '.');
    } else if (s.includes('.')) {
        const parts = s.split('.');
        // Si hay varios puntos (1.234.567) o un punto seguido de 3 dígitos (1.234), es probable que sea separador de miles
        if (parts.length > 2 || (parts[parts.length - 1].length === 3 && parts[0].length <= 3)) {
            // Pero cuidado con los decimales de 3 dígitos. 
            // En este contexto de ATOS y CANTIDADES, es más probable que sean miles si el número es grande.
            // Para ser más seguro, si el valor total sin puntos es razonable, lo unimos.
            s = s.replace(/\./g, '');
        }
    }

    const parsed = parseFloat(s);
    return isNaN(parsed) ? 0 : parsed;
}

function calculateSums() {
    return; // Desactivado: El usuario prefiere ver solo el total de lo filtrado
    if (data.length === 0) return;
    const headers = Object.keys(data[0]);
    sumsContainer.innerHTML = '';
    let grandTotal = 0;
    headers.forEach(header => {
        const sum = data.reduce((acc, row) => acc + cleanNumber(row[header]), 0);
        const hasNumbers = data.some(row => String(row[header]).trim() !== "" && !isNaN(cleanNumber(row[header])) && cleanNumber(row[header]) !== 0);
        if (hasNumbers && sum !== 0) {
            grandTotal += sum;
            const item = document.createElement('div');
            item.className = 'sum-item';
            item.innerHTML = `<span class="sum-name">${header}</span><span class="sum-value">${sum.toLocaleString(undefined, { minimumFractionDigits: 0, maximumFractionDigits: 0 })}</span>`;
            sumsContainer.appendChild(item);
        }
    });
    grandTotalValue.innerText = `$ ${grandTotal.toLocaleString(undefined, { minimumFractionDigits: 0, maximumFractionDigits: 0 })}`;
}

function getExcelColumnName(columnNumber) {
    let columnName = "";
    while (columnNumber >= 0) {
        columnName = String.fromCharCode((columnNumber % 26) + 65) + columnName;
        columnNumber = Math.floor(columnNumber / 26) - 1;
    }
    return columnName;
}

function populateSelectors() {
    if (data.length === 0) return;
    const headers = Object.keys(data[0]);
    const options = headers.map(h => {
        return `<option value="${h}">${h}</option>`;
    }).join('');
    selectColCodigo.innerHTML = options;
    selectColTotal.innerHTML = options;
    selectColCantidad.innerHTML = options;
    selectColCosto.innerHTML = `<option value="">-- No validar --</option>` + options;

    // Estrategia de búsqueda mejorada
    const findHeaderByLetter = (letter) => headers.find(h => h.startsWith(letter + " -") || h.includes("Columna " + letter));
    const findHeaderByKeywords = (keywords) => headers.find(h => {
        const s = h.toUpperCase();
        return keywords.every(k => s.includes(k));
    });
    const findHeaderByName = (keyword) => headers.find(h => h.toUpperCase().includes(keyword.toUpperCase()));

    // 1. Código: Prioridad columna M o que diga "CODIGO"
    const colCodigoDefault = findHeaderByName("CODIGO") || findHeaderByLetter("M");

    // 2. Valores ($): Prioridad columna V o que diga "TOTAL ATOS"
    const colValoresDefault = findHeaderByKeywords(["TOTAL", "ATOS"]) || findHeaderByLetter("V");

    // 3. Cantidades (ATOs): Prioridad "ATOS NETOS", luego "NETO", luego "CANTIDAD", luego N
    const colCantidadesDefault = findHeaderByKeywords(["ATOS", "NETO"]) || findHeaderByName("NETO") || findHeaderByName("CANT") || findHeaderByLetter("N");

    // 4. Costo: Prioridad "UNITARIO", "COSTO", luego T
    const colCostoDefault = findHeaderByName("UNITARIO") || findHeaderByName("COSTO") || findHeaderByLetter("T");

    if (colCodigoDefault) selectColCodigo.value = colCodigoDefault;
    if (colValoresDefault) selectColTotal.value = colValoresDefault;
    if (colCantidadesDefault) selectColCantidad.value = colCantidadesDefault;
    if (colCostoDefault) selectColCosto.value = colCostoDefault;
}

const codeSearchInput = document.getElementById('code-search-input');
const codesDatalist = document.getElementById('codes-datalist');
const searchResult = document.getElementById('search-result');
const selectedCodesTags = document.getElementById('selected-codes-tags');
const selectedTotalsContainer = document.getElementById('selected-totals-container');

let selectedCodes = new Set();

// Manejar el pegado masivo desde Excel
codeSearchInput.addEventListener('paste', (e) => {
    e.preventDefault();
    const pasteData = (e.clipboardData || window.clipboardData).getData('text');
    if (!pasteData) return;

    // Separamos por saltos de línea, comas o pestañas (típico de Excel)
    const codes = pasteData.split(/[\n\r\t,]+/)
        .map(c => c.trim().toUpperCase())
        .filter(c => c !== "");

    codes.forEach(code => selectedCodes.add(code));
    renderTags();
    updateSpecificSearch();
});

// Manejar la entrada del buscador
codeSearchInput.addEventListener('keydown', (e) => {
    if (e.key === 'Enter') {
        const val = codeSearchInput.value.trim().toUpperCase();
        if (val) addCodeToSelection(val);
    }
});

codeSearchInput.addEventListener('change', () => {
    const val = codeSearchInput.value.trim().toUpperCase();
    if (val) addCodeToSelection(val);
});

function addCodeToSelection(code) {
    if (!code) return;
    selectedCodes.add(code);
    codeSearchInput.value = '';
    renderTags();
    updateSpecificSearch();
}

function removeCodeFromSelection(code) {
    selectedCodes.delete(code);
    renderTags();
    updateSpecificSearch();
}

function renderTags() {
    selectedCodesTags.innerHTML = '';
    selectedCodes.forEach(code => {
        const tag = document.createElement('span');
        tag.className = 'code-tag';
        tag.innerText = code;
        tag.onclick = () => removeCodeFromSelection(code);
        selectedCodesTags.appendChild(tag);
    });
}

function updateSpecificSearch() {
    if (selectedCodes.size === 0) {
        searchResult.innerHTML = `<p class="empty-msg">Seleccione códigos para sumar</p>`;
        return;
    }

    const colCodigo = selectColCodigo.value;
    const colTotalPesos = selectColTotal.value;
    const colTotalCantidades = selectColCantidad.value;

    let grandTotalPesos = 0;
    let grandTotalCantidades = 0;
    let foundAny = false;
    let individualResultsHtml = "";

    // Para cada código seleccionado, calcular su propia sumatoria
    selectedCodes.forEach(sel => {
        let codeTotalPesos = 0;
        let codeTotalCantidades = 0;
        let codeCount = 0;
        const selClean = sel.replace(/\s/g, '').toUpperCase();

        data.forEach(row => {
            const codigoOriginal = String(row[colCodigo]).trim();
            const codigoClean = codigoOriginal.toUpperCase().replace(/\s/g, '');

            if (codigoClean === selClean || codigoClean.includes(selClean)) {
                codeTotalPesos += cleanNumber(row[colTotalPesos]);
                codeTotalCantidades += cleanNumber(row[colTotalCantidades]);
                codeCount++;
                foundAny = true;
            }
        });

        if (codeCount > 0) {
            grandTotalPesos += codeTotalPesos;
            grandTotalCantidades += codeTotalCantidades;
            individualResultsHtml += `
                <div class="sum-item" style="margin-bottom: 8px; flex-direction: column; align-items: flex-start; gap: 4px;">
                    <div style="display: flex; justify-content: space-between; width: 100%; align-items: center;">
                        <span class="code-key">${sel}</span>
                        <span class="sum-value" style="color: var(--accent);">$ ${codeTotalPesos.toLocaleString(undefined, { minimumFractionDigits: 0, maximumFractionDigits: 0 })}</span>
                    </div>
                    <div style="display: flex; justify-content: space-between; width: 100%; font-size: 0.8rem; color: var(--text-dim);">
                        <span>Registros: ${codeCount}</span>
                        <span style="color: #fbbf24; font-weight: bold;">Cant: ${codeTotalCantidades.toLocaleString()} ATOs</span>
                    </div>
                </div>
            `;
        } else {
            individualResultsHtml += `
                <div class="sum-item" style="margin-bottom: 8px; border-left-color: #ef4444;">
                    <div class="code-info">
                        <span class="code-key">${sel}</span>
                        <div class="cost-info"><span style="color:#ef4444">No encontrado</span></div>
                    </div>
                </div>
            `;
        }
    });

    if (foundAny || selectedCodes.size > 0) {
        // Actualizar el cuadro de búsqueda con el GRAN TOTAL en PESOS y el total de ATOS
        searchResult.innerHTML = `
            <div class="result-highlight">
                <span class="result-code">Gran Total (${selectedCodes.size} códigos):</span>
                <span class="result-amount">$ ${grandTotalPesos.toLocaleString(undefined, { minimumFractionDigits: 0, maximumFractionDigits: 0 })}</span>
                <span style="display: block; font-size: 0.9rem; color: #fbbf24; font-weight: bold; margin-top: 4px;">Total ATOS: ${grandTotalCantidades.toLocaleString()}</span>
            </div>
        `;

        // Actualizar el panel lateral con el desglose INDIVIDUAL
        selectedTotalsContainer.innerHTML = individualResultsHtml;
    } else {
        const emptyHtml = `<p class="empty-msg" style="color:#ef4444">No se encontraron datos</p>`;
        searchResult.innerHTML = emptyHtml;
        selectedTotalsContainer.innerHTML = emptyHtml;
    }
}

function updateDatalist() {
    const colCodigo = selectColCodigo.value;
    if (!colCodigo) return;
    const uniqueCodes = [...new Set(data.map(row => String(row[colCodigo]).trim()))]
        .filter(c => c && c.toUpperCase() !== String(colCodigo).toUpperCase())
        .sort();
    codesDatalist.innerHTML = uniqueCodes.map(c => `<option value="${c}">`).join('');
}

function generateCodeSummary() {
    if (data.length === 0) return;
    updateDatalist();
    updateSpecificSearch();
}

// Modal History Logic
const historyModal = document.getElementById('history-modal');
const closeModal = document.getElementById('close-modal');
const historyTableBody = document.getElementById('history-table-body');
const historyProductCode = document.getElementById('history-product-code');

closeModal.onclick = () => historyModal.classList.add('hidden');
window.onclick = (event) => { if (event.target == historyModal) historyModal.classList.add('hidden'); };

function showHistory(code) {
    const codeUpper = code.toUpperCase();
    historyProductCode.innerText = codeUpper;
    const history = productCostsHistory[codeUpper] || [];

    if (history.length === 0 && productCosts[codeUpper]) {
        // Migración simple si hay costo pero no historia
        history.push({
            date: new Date().toISOString(),
            cost: productCosts[codeUpper],
            file: 'Memoria local'
        });
    }

    historyTableBody.innerHTML = history.slice().reverse().map(h => `
        <tr>
            <td>${new Date(h.date).toLocaleDateString()} ${new Date(h.date).toLocaleTimeString([], { hour: '2-digit', minute: '2-digit' })}</td>
            <td style="color:var(--accent); font-weight:bold">$ ${h.cost.toLocaleString()}</td>
            <td style="font-size:0.7rem; color:var(--text-dim)">${h.file || 'N/A'}</td>
        </tr>
    `).join('') || '<tr><td colspan="3" style="text-align:center">No hay historial registrado</td></tr>';

    historyModal.classList.remove('hidden');
}

function updateStats() {
    statRows.innerText = data.length;
    statCols.innerText = data.length > 0 ? Object.keys(data[0]).length : 0;
}

exportBtn.addEventListener('click', () => {
    if (data.length === 0) return;

    try {
        // 1. Limpiamos los encabezados para el Excel de salida (quitamos el "M - ")
        const exportData = data.map(row => {
            const cleanRow = {};
            Object.keys(row).forEach(key => {
                // Si el nombre tiene el formato "A - Titulo", tomamos solo "Titulo"
                const cleanKey = key.includes(' - ') ? key.split(' - ').slice(1).join(' - ') : key;
                cleanRow[cleanKey] = row[key];
            });
            return cleanRow;
        });

        // 2. Creamos la hoja de cálculo
        const newWs = XLSX.utils.json_to_sheet(exportData);

        // 3. Creamos el libro y añadimos la hoja
        const newWb = XLSX.utils.book_new();
        XLSX.book_append_sheet(newWb, newWs, "Resultados");

        // 4. Generamos la descarga
        XLSX.writeFile(newWb, "ExcelMasterPro_Exportado.xlsx");

    } catch (e) {
        console.error(e);
        alert("Error al exportar el archivo. Por favor intente de nuevo.");
    }
});
