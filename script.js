// Establecer celda activa
function setActiveCell(row, col) {
    // Limpiar celda activa anterior
    document.querySelectorAll('.cell-active').forEach(cell => 
        cell.classList.remove('cell-active'));

    activeCell = { row, col };
    
    // Marcar nueva celda activa
    const cell = document.querySelector(`.sheet-table td[data-row="${row}"][data-col="${col}"]`);
    if (cell) {
        cell.classList.add('cell-active');
        
        // Hacer scroll para asegurar que la celda sea visible
        cell.scrollIntoView({ 
            behavior: 'smooth', 
            block: 'nearest', 
            inline: 'nearest' 
        });
    }

    // Actualizar selección
    selectedCells.clear();
    selectedCells.add(`${row},${col}`);
    highlightSelectedCells();
}

// Iniciar edición de celda activa
function startEditingActiveCell() {
    if (!activeCell) return;

    const cell = document.querySelector(`.sheet-table td[data-row="${activeCell.row}"][data-col="${activeCell.col}"] input`);
    if (cell) {
        cell.focus();
        cell.select();
    }
}

// Limpiar selección
function clearSelection() {
    selectedCells.clear();
    activeCell = null;
    document.querySelectorAll('.cell-active, .cell-selected').forEach(cell => {
        cell.classList.remove('cell-active', 'cell-selected');
    });
}

// Configurar eventos de selección de celdas mejorados
function setupSelectionEvents() {
    const sheetContainer = document.getElementById('sheetContainer');
    
    sheetContainer.addEventListener('mousedown', function(e) {
        if (e.target.classList.contains('column-resizer') || e.target.tagName === 'TH' || e.target.classList.contains('row-number-cell')) {
            return;
        }
        
        if (e.target.tagName === 'INPUT') {
            // Si hacemos clic en un input, establecer como celda activa
            const cell = e.target.closest('td');
            if (cell) {
                const row = parseInt(cell.dataset.row);
                const col = parseInt(cell.dataset.col);
                setActiveCell(row, col);
            }
            return;
        }
        
        const cell = e.target.closest('td');
        if (!cell) return;
        
        const row = parseInt(cell.dataset.row);
        const col = parseInt(cell.dataset.col);
        
        // Establecer como celda activa
        setActiveCell(row, col);
        
        // Iniciar selección múltiple
        isSelecting = true;
        selectionStartCell = { row, col };
        selectionEndCell = { row, col };
        
        if (!e.ctrlKey && !e.shiftKey) {
            selectedCells.clear();
        }
        
        updateSelection(e.ctrlKey, e.shiftKey);
        e.preventDefault();
    });
    
    sheetContainer.addEventListener('mousemove', function(e) {
        if (!isSelecting) return;
        
        const cell = e.target.closest('td');
        if (!cell) return;
        
        const row = parseInt(cell.dataset.row);
        const col = parseInt(cell.dataset.col);
        selectionEndCell = { row, col };
        
        updateSelection(e.ctrlKey, e.shiftKey);
    });
    
    document.addEventListener('mouseup', function(e) {
        if (!isSelecting) return;
        isSelecting = false;
        selectionStartCell = null;
        selectionEndCell = null;
    });

    // Doble clic para editar
    sheetContainer.addEventListener('dblclick', function(e) {
        const cell = e.target.closest('td');
        if (cell) {
            const input = cell.querySelector('input');
            if (input) {
                input.focus();
                input.select();
            }
        }
    });
}

// Actualizar el conjunto de celdas seleccionadas
function updateSelection(ctrlKey, shiftKey) {
    if (!selectionStartCell || !selectionEndCell) return;
    
    const newSelection = new Set();
    const minRow = Math.min(selectionStartCell.row, selectionEndCell.row);
    const maxRow = Math.max(selectionStartCell.row, selectionEndCell.row);
    const minCol = Math.min(selectionStartCell.col, selectionEndCell.col);
    const maxCol = Math.max(selectionStartCell.col, selectionEndCell.col);
    
    for (let r = minRow; r <= maxRow; r++) {
        for (let c = minCol; c <= maxCol; c++) {
            newSelection.add(`${r},${c}`);
        }
    }
    
    if (ctrlKey) {
        newSelection.forEach(cell => selectedCells.add(cell));
    } else {
        selectedCells = newSelection;
    }
    
    highlightSelectedCells();
}

// Resaltar las celdas seleccionadas
function highlightSelectedCells() {
    document.querySelectorAll('.sheet-table td.cell-selected').forEach(cell => 
        cell.classList.remove('cell-selected'));
    
    selectedCells.forEach(cellKey => {
        const [row, col] = cellKey.split(',').map(Number);
        const cell = document.querySelector(`.sheet-table td[data-row="${row}"][data-col="${col}"]`);
        if (cell) {
            cell.classList.add('cell-selected');
        }
    });
}

// Función para procesar los datos GeoJSON
function processGeoData() {
    saveToHistory();
    
    const propertyKeys = new Set();
    geoData.features.forEach(feature => {
        if (feature.properties) {
            Object.keys(feature.properties).forEach(key => {
                propertyKeys.add(key);
            });
        }
    });
    
    currentColumns = ['geometry', ...Array.from(propertyKeys).sort()];
    selectedCells.clear();
    activeCell = null;
    
    rebuildSheet();
    updateUndoRedoButtons();
    updateWrapButton();
}

// Función para convertir índice numérico a nombre de columna estilo Excel
function getExcelColumnName(index) {
    let columnName = '';
    
    while (index >= 0) {
        columnName = String.fromCharCode(65 + (index % 26)) + columnName;
        index = Math.floor(index / 26) - 1;
    }
    
    return columnName;
}

// Función para reconstruir la tabla completa
function rebuildSheet() {
    const table = document.getElementById('sheetTable');
    const headerRow = document.getElementById('columnHeadersRow');
    const tableBody = document.getElementById('sheetTableBody');
    
    headerRow.innerHTML = '<th class="corner-header-cell"></th>';
    tableBody.innerHTML = '';
    
    // Crear encabezados de columna
    currentColumns.forEach((col, colIndex) => {
        const th = document.createElement('th');
        th.dataset.index = colIndex;
        
        const excelColName = getExcelColumnName(colIndex);
        th.textContent = excelColName;
        th.title = col;
        
        if (!columnNameMap[excelColName]) {
            columnNameMap[excelColName] = col;
        }
        
        const colWidth = columnWidths[colIndex] || 120;
        th.style.width = colWidth + 'px';
        th.style.minWidth = colWidth + 'px';
        th.style.maxWidth = colWidth + 'px';
        
        th.addEventListener('click', function() {
            selectColumn(colIndex);
        });
        
        const resizer = document.createElement('div');
        resizer.className = 'column-resizer';
        resizer.dataset.colIndex = colIndex;
        resizer.addEventListener('mousedown', startColumnResize);
        th.appendChild(resizer);
        
        headerRow.appendChild(th);
    });
    
    // Crear fila para nombres de propiedades
    const propertyNameRow = document.createElement('tr');
    propertyNameRow.className = 'property-name-row';
    
    const propertyRowNumTd = document.createElement('td');
    propertyRowNumTd.className = 'row-number-cell';
    propertyRowNumTd.textContent = '1';
    propertyNameRow.appendChild(propertyRowNumTd);
    
    currentColumns.forEach((col, colIndex) => {
        const td = document.createElement('td');
        td.dataset.row = -1;
        td.dataset.col = colIndex;
        
        const colWidth = columnWidths[colIndex] || 120;
        td.style.width = colWidth + 'px';
        td.style.minWidth = colWidth + 'px';
        td.style.maxWidth = colWidth + 'px';
        
        if (isWrapMode) {
            td.classList.add('wrap-mode');
        }
        
        const input = document.createElement('input');
        input.type = 'text';
        input.dataset.row = -1;
        input.dataset.col = colIndex;
        input.value = col;
        
        if (col.length > 15) {
            td.setAttribute('data-tooltip', col);
        }
        
        input.addEventListener('change', function(e) {
            saveToHistory();
            
            const oldName = currentColumns[colIndex];
            const newName = e.target.value.trim();
            
            if (!newName) {
                updateStatus("El nombre de la propiedad no puede estar vacío", true);
                e.target.value = oldName;
                return;
            }
            
            if (newName !== oldName && currentColumns.includes(newName)) {
                updateStatus("Ya existe una propiedad con ese nombre", true);
                e.target.value = oldName;
                return;
            }
            
            if (oldName === 'geometry') {
                updateStatus("No se puede cambiar el nombre de la columna geometry", true);
                e.target.value = 'geometry';
                return;
            }
            
            currentColumns[colIndex] = newName;
            
            const excelColName = getExcelColumnName(colIndex);
            columnNameMap[excelColName] = newName;
            
            geoData.features.forEach(feature => {
                if (feature.properties && feature.properties[oldName] !== undefined) {
                    feature.properties[newName] = feature.properties[oldName];
                    if (oldName !== newName) {
                        delete feature.properties[oldName];
                    }
                }
            });
            
            if (newName.length > 15) {
                td.setAttribute('data-tooltip', newName);
            } else {
                td.removeAttribute('data-tooltip');
            }
            
            updateStatus(`Propiedad renombrada de "${oldName}" a "${newName}"`);
        });
        
        td.appendChild(input);
        propertyNameRow.appendChild(td);
    });
    
    tableBody.appendChild(propertyNameRow);
    
    // Crear filas de datos
    geoData.features.forEach((feature, rowIndex) => {
        const tr = document.createElement('tr');
        
        const rowNumTd = document.createElement('td');
        rowNumTd.className = 'row-number-cell';
        rowNumTd.textContent = rowIndex + 2;
        rowNumTd.dataset.index = rowIndex;
        rowNumTd.addEventListener('click', function() {
            selectRow(rowIndex);
        });
        tr.appendChild(rowNumTd);
        
        currentColumns.forEach((col, colIndex) => {
            const td = document.createElement('td');
            td.dataset.row = rowIndex;
            td.dataset.col = colIndex;
            
            const colWidth = columnWidths[colIndex] || 120;
            td.style.width = colWidth + 'px';
            td.style.minWidth = colWidth + 'px';
            td.style.maxWidth = colWidth + 'px';
            
            if (col === 'geometry') td.classList.add('coords-cell');
            
            if (isWrapMode) {
                td.classList.add('wrap-mode');
            }
            
            const input = document.createElement('input');
            input.type = 'text';
            input.dataset.row = rowIndex;
            input.dataset.col = colIndex;
            
            if (col === 'geometry') {
                if (feature.geometry && feature.geometry.coordinates) {
                    input.value = feature.geometry.coordinates.join(', ');
                } else {
                    input.value = '0, 0';
                }
                
                input.addEventListener('change', function(e) {
                    try {
                        saveToHistory();
                        
                        const coords = e.target.value.split(',').map(Number);
                        if (coords.length === 2 && !isNaN(coords[0]) && !isNaN(coords[1])) {
                            if (!feature.geometry) feature.geometry = { type: "Point" };
                            feature.geometry.coordinates = coords;
                            updateStatus(`Coordenadas actualizadas en fila ${rowIndex + 2}`);
                        } else {
                            throw new Error("Formato de coordenadas inválido");
                        }
                    } catch (error) {
                    console.error("Error al guardar coordenadas:", error);
                }
            } else {
                if (!feature.properties) feature.properties = {};
                feature.properties[colName] = input.value;
            }
        }
    });
    
    updateStatus("Datos guardados en memoria");
}

// Función para descargar el archivo
function downloadFile() {
    // Primero guardar los datos actuales
    saveData();
    
    // Crear una copia profunda de los datos para exportar
    const dataToExport = JSON.parse(JSON.stringify(geoData));
    
    // Limpiar y reorganizar las propiedades según el orden actual de las columnas
    dataToExport.features.forEach(feature => {
        const newProperties = {};
        currentColumns.forEach(colName => {
            if (colName !== 'geometry') {
                // Asegurarse de que la propiedad existe, aunque sea vacía
                newProperties[colName] = feature.properties && feature.properties[colName] !== undefined 
                    ? feature.properties[colName] 
                    : "";
            }
        });
        feature.properties = newProperties;
    });
    
    // Crear el blob y descargar
    const blob = new Blob([JSON.stringify(dataToExport, null, 2)], {type: 'application/json'});
    const a = document.createElement('a');
    a.href = URL.createObjectURL(blob);
    a.download = 'geojson_export.geojson';
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    URL.revokeObjectURL(a.href);
    
    updateStatus("Archivo descargado como geojson_export.geojson");
}

// Función para actualizar el estado
function updateStatus(message, isError = false) {
    const statusBar = document.getElementById('statusBar');
    statusBar.textContent = message;
    
    if (isError) {
        statusBar.classList.add('error');
    } else {
        statusBar.classList.remove('error');
    }
}error) {
                        updateStatus(`Error: ${error.message}`, true);
                        e.target.value = feature.geometry?.coordinates?.join(', ') || '0, 0';
                    }
                });
            } else {
                const value = feature.properties?.[col] || '';
                input.value = value;
                
                if (String(value).length > 15) {
                    td.setAttribute('data-tooltip', value);
                }
                
                input.addEventListener('change', function(e) {
                    saveToHistory();
                    
                    if (!feature.properties) feature.properties = {};
                    feature.properties[col] = e.target.value;
                    
                    if (String(e.target.value).length > 15) {
                        td.setAttribute('data-tooltip', e.target.value);
                    } else {
                        td.removeAttribute('data-tooltip');
                    }
                    
                    updateStatus(`Propiedad ${col} actualizada en fila ${rowIndex + 2}`);
                });
            }
            
            td.appendChild(input);
            tr.appendChild(td);
        });
        
        tableBody.appendChild(tr);
    });
    
    highlightSelectedCells();
}

// Función para seleccionar una fila
function selectRow(rowIndex) {
    saveToHistory();
    selectedCells.clear();
    
    for (let colIndex = 0; colIndex < currentColumns.length; colIndex++) {
        selectedCells.add(`${rowIndex},${colIndex}`);
    }
    
    setActiveCell(rowIndex, 0);
    highlightSelectedCells();
    updateStatus(`Fila ${rowIndex + 2} seleccionada`);
}

// Función para seleccionar una columna
function selectColumn(colIndex) {
    saveToHistory();
    selectedCells.clear();
    
    for (let rowIndex = -1; rowIndex < geoData.features.length; rowIndex++) {
        selectedCells.add(`${rowIndex},${colIndex}`);
    }
    
    setActiveCell(0, colIndex);
    highlightSelectedCells();
    
    const excelColName = getExcelColumnName(colIndex);
    const propertyName = columnNameMap[excelColName];
    updateStatus(`Columna ${excelColName} (${propertyName}) seleccionada`);
}

// Funciones de redimensionamiento de columnas
function startColumnResize(e) {
    e.preventDefault();
    e.stopPropagation();
    
    isResizing = true;
    currentColIndex = parseInt(this.dataset.colIndex);
    startX = e.pageX;
    startWidth = columnWidths[currentColIndex] || 120;
    
    document.addEventListener('mousemove', doColumnResize);
    document.addEventListener('mouseup', stopColumnResize);
    document.body.style.cursor = 'col-resize';
}

function doColumnResize(e) {
    if (!isResizing) return;
    const delta = e.pageX - startX;
    const newWidth = Math.max(60, startWidth + delta);
}

function stopColumnResize(e) {
    if (!isResizing) return;
    
    saveToHistory();
    const delta = e.pageX - startX;
    const newWidth = Math.max(60, startWidth + delta);
    columnWidths[currentColIndex] = newWidth;
    
    document.removeEventListener('mousemove', doColumnResize);
    document.removeEventListener('mouseup', stopColumnResize);
    document.body.style.cursor = '';
    
    rebuildSheet();
    isResizing = false;
    
    const excelColName = getExcelColumnName(currentColIndex);
    const propertyName = columnNameMap[excelColName];
    updateStatus(`Columna ${excelColName} (${propertyName}) redimensionada a ${newWidth}px`);
}

// Función para agregar una nueva columna
function addColumn() {
    saveToHistory();
    
    let newColName = "Column1";
    let counter = 1;
    while (currentColumns.includes(newColName)) {
        counter++;
        newColName = `Column${counter}`;
    }
    
    currentColumns.push(newColName);
    
    const newColIndex = currentColumns.length - 1;
    const excelColName = getExcelColumnName(newColIndex);
    columnNameMap[excelColName] = newColName;
    
    // CORRECCIÓN: Agregar la propiedad a todas las features existentes
    geoData.features.forEach(feature => {
        if (!feature.properties) feature.properties = {};
        feature.properties[newColName] = "";
    });
    
    rebuildSheet();
    updateStatus(`Nueva columna "${newColName}" (${excelColName}) agregada`);
}

// Función para agregar una nueva fila
function addRow() {
    saveToHistory();
    
    const newFeature = {
        type: "Feature",
        geometry: {
            type: "Point",
            coordinates: [0, 0]
        },
        properties: {}
    };
    
    // Agregar todas las propiedades existentes con valores vacíos
    currentColumns.forEach(col => {
        if (col !== 'geometry') {
            newFeature.properties[col] = "";
        }
    });
    
    geoData.features.push(newFeature);
    rebuildSheet();
    updateStatus(`Nueva fila agregada`);
}

// Función para eliminar la selección actual
function deleteSelected() {
    if (selectedCells.size === 0) {
        updateStatus("No hay nada seleccionado para eliminar", true);
        return;
    }
    
    saveToHistory();
    
    const rowsToDelete = new Set();
    const colsToDelete = new Set();
    const cellsToClear = new Set();
    
    selectedCells.forEach(cellKey => {
        const [row, col] = cellKey.split(',').map(Number);
        if (row === -1) {
            colsToDelete.add(col);
        } else {
            cellsToClear.add(cellKey);
        }
    });
    
    if (colsToDelete.size > 0) {
        const colsToDeleteArray = Array.from(colsToDelete).sort((a, b) => b - a);
        
        colsToDeleteArray.forEach(colIndex => {
            const colName = currentColumns[colIndex];
            
            if (colName === 'geometry') {
                updateStatus("No se puede eliminar la columna geometry", true);
                return;
            }
            
            geoData.features.forEach(feature => {
                if (feature.properties && feature.properties[colName] !== undefined) {
                    delete feature.properties[colName];
                }
            });
            
            currentColumns.splice(colIndex, 1);
            delete columnWidths[colIndex];
            
            const newColumnWidths = {};
            Object.keys(columnWidths).forEach(key => {
                const keyNum = parseInt(key);
                if (keyNum < colIndex) {
                    newColumnWidths[keyNum] = columnWidths[keyNum];
                } else if (keyNum > colIndex) {
                    newColumnWidths[keyNum - 1] = columnWidths[keyNum];
                }
            });
            columnWidths = newColumnWidths;
            
            const newColumnNameMap = {};
            currentColumns.forEach((name, index) => {
                const excelName = getExcelColumnName(index);
                newColumnNameMap[excelName] = name;
            });
            columnNameMap = newColumnNameMap;
        });
        
        updateStatus(`${colsToDeleteArray.length} columna(s) eliminada(s)`);
    }
    
    if (colsToDelete.size === 0 && cellsToClear.size > 0) {
        cellsToClear.forEach(cellKey => {
            const [row, col] = cellKey.split(',').map(Number);
            const colName = currentColumns[col];
            
            if (row >= 0 && row < geoData.features.length) {
                if (colName !== 'geometry') {
                    if (geoData.features[row].properties) {
                        geoData.features[row].properties[colName] = "";
                    }
                }
            }
        });
        updateStatus(`${cellsToClear.size} celda(s) limpiada(s)`);
    }
    
    selectedCells.clear();
    activeCell = null;
    rebuildSheet();
}

// Función para copiar la selección actual
function copySelection() {
    if (selectedCells.size === 0) {
        updateStatus("No hay nada seleccionado para copiar", true);
        return;
    }
    
    clipboardData = {
        type: 'cells',
        data: []
    };
    
    const rows = new Set();
    const cols = new Set();
    selectedCells.forEach(cellKey => {
        const [row, col] = cellKey.split(',').map(Number);
        rows.add(row);
        cols.add(col);
    });
    
    const minRow = Math.min(...Array.from(rows));
    const maxRow = Math.max(...Array.from(rows));
    const minCol = Math.min(...Array.from(cols));
    const maxCol = Math.max(...Array.from(cols));
    
    for (let r = minRow; r <= maxRow; r++) {
        const rowData = [];
        for (let c = minCol; c <= maxCol; c++) {
            if (selectedCells.has(`${r},${c}`)) {
                const cell = document.querySelector(`.sheet-table td[data-row="${r}"][data-col="${c}"] input`);
                rowData.push(cell ? cell.value : "");
            } else {
                rowData.push(null);
            }
        }
        clipboardData.data.push(rowData);
    }
    
    updateStatus(`Selección copiada (${clipboardData.data.length}x${clipboardData.data[0].length})`);
}

// Función para pegar la selección copiada
function pasteSelection() {
    if (!clipboardData || clipboardData.type !== 'cells') {
        updateStatus("No hay datos de celdas para pegar", true);
        return;
    }
    
    if (selectedCells.size === 0) {
        updateStatus("Selecciona una celda para empezar a pegar", true);
        return;
    }
    
    const startCellKey = Array.from(selectedCells)[0];
    const [startRow, startCol] = startCellKey.split(',').map(Number);
    
    saveToHistory();
    
    clipboardData.data.forEach((rowData, rOffset) => {
        rowData.forEach((cellValue, cOffset) => {
            if (cellValue !== null) {
                const targetRow = startRow + rOffset;
                const targetCol = startCol + cOffset;
                
                if (targetRow >= -1 && targetRow < geoData.features.length && targetCol >= 0 && targetCol < currentColumns.length) {
                    const colName = currentColumns[targetCol];
                    
                    if (targetRow === -1) {
                        // No permitir pegar en la fila de nombres por ahora
                    } else {
                        if (colName === 'geometry') {
                            try {
                                const coords = cellValue.split(',').map(Number);
                                if (coords.length === 2 && !isNaN(coords[0]) && !isNaN(coords[1])) {
                                    if (!geoData.features[targetRow].geometry) geoData.features[targetRow].geometry = { type: "Point" };
                                    geoData.features[targetRow].geometry.coordinates = coords;
                                }
                            } catch (error) {
                                console.error("Error al pegar coordenadas:", error);
                            }
                        } else {
                            if (!geoData.features[targetRow].properties) geoData.features[targetRow].properties = {};
                            geoData.features[targetRow].properties[colName] = cellValue;
                        }
                    }
                }
            }
        });
    });
    
    rebuildSheet();
    updateStatus(`Datos pegados`);
}

// Función para alternar el modo wrap
function toggleWrapMode() {
    isWrapMode = !isWrapMode;
    updateWrapButton();
    
    const cells = document.querySelectorAll('.sheet-table td');
    cells.forEach(cell => {
        if (isWrapMode) {
            cell.classList.add('wrap-mode');
        } else {
            cell.classList.remove('wrap-mode');
        }
    });
    
    updateStatus(`Modo wrap ${isWrapMode ? 'activado' : 'desactivado'}`);
}

// Función para actualizar el botón de wrap
function updateWrapButton() {
    const wrapButton = document.getElementById('wrapButton');
    if (isWrapMode) {
        wrapButton.classList.add('active');
    } else {
        wrapButton.classList.remove('active');
    }
}

// Función para guardar el estado actual en el historial de deshacer
function saveToHistory() {
    undoStack.push({
        geoData: JSON.parse(JSON.stringify(geoData)),
        currentColumns: [...currentColumns],
        columnWidths: {...columnWidths},
        selectedCells: new Set(selectedCells),
        columnNameMap: {...columnNameMap}
    });
    
    if (undoStack.length > maxUndoSteps) {
        undoStack.shift();
    }
    
    redoStack = [];
    updateUndoRedoButtons();
}

// Función para deshacer la última acción
function undoAction() {
    if (undoStack.length === 0) {
        updateStatus("No hay acciones para deshacer", true);
        return;
    }
    
    redoStack.push({
        geoData: JSON.parse(JSON.stringify(geoData)),
        currentColumns: [...currentColumns],
        columnWidths: {...columnWidths},
        selectedCells: new Set(selectedCells),
        columnNameMap: {...columnNameMap}
    });
    
    const prevState = undoStack.pop();
    geoData = prevState.geoData;
    currentColumns = prevState.currentColumns;
    columnWidths = prevState.columnWidths;
    selectedCells = prevState.selectedCells;
    columnNameMap = prevState.columnNameMap;
    activeCell = null;
    
    rebuildSheet();
    updateUndoRedoButtons();
    updateStatus("Acción deshecha");
}

// Función para rehacer la última acción deshecha
function redoAction() {
    if (redoStack.length === 0) {
        updateStatus("No hay acciones para rehacer", true);
        return;
    }
    
    undoStack.push({
        geoData: JSON.parse(JSON.stringify(geoData)),
        currentColumns: [...currentColumns],
        columnWidths: {...columnWidths},
        selectedCells: new Set(selectedCells),
        columnNameMap: {...columnNameMap}
    });
    
    const nextState = redoStack.pop();
    geoData = nextState.geoData;
    currentColumns = nextState.currentColumns;
    columnWidths = nextState.columnWidths;
    selectedCells = nextState.selectedCells;
    columnNameMap = nextState.columnNameMap;
    activeCell = null;
    
    rebuildSheet();
    updateUndoRedoButtons();
    updateStatus("Acción rehecha");
}

// Función para actualizar los botones de deshacer/rehacer
function updateUndoRedoButtons() {
    const undoButton = document.getElementById('undoButton');
    const redoButton = document.getElementById('redoButton');
    const undoOption = document.getElementById('undoOption');
    const redoOption = document.getElementById('redoOption');
    
    if (undoStack.length === 0) {
        undoButton.classList.add('disabled');
        undoOption.classList.add('disabled');
    } else {
        undoButton.classList.remove('disabled');
        undoOption.classList.remove('disabled');
    }
    
    if (redoStack.length === 0) {
        redoButton.classList.add('disabled');
        redoOption.classList.add('disabled');
    } else {
        redoButton.classList.remove('disabled');
        redoOption.classList.remove('disabled');
    }
}

// NUEVAS FUNCIONES: Separar Save y Download

// Función para guardar los datos (actualizar los datos en memoria)
function saveData() {
    // Asegurarse de que todos los valores actuales estén guardados
    const inputs = document.querySelectorAll('.sheet-table input');
    inputs.forEach(input => {
        const row = parseInt(input.dataset.row);
        const col = parseInt(input.dataset.col);
        const colName = currentColumns[col];
        
        if (row === -1) {
            // Actualizar nombre de columna si es necesario
            if (colName !== input.value.trim() && input.value.trim() !== '') {
                currentColumns[col] = input.value.trim();
                const excelColName = getExcelColumnName(col);
                columnNameMap[excelColName] = input.value.trim();
            }
        } else if (row >= 0 && row < geoData.features.length) {
            const feature = geoData.features[row];
            
            if (colName === 'geometry') {
                try {
                    const coords = input.value.split(',').map(Number);
                    if (coords.length === 2 && !isNaN(coords[0]) && !isNaN(coords[1])) {
                        if (!feature.geometry) feature.geometry = { type: "Point" };
                        feature.geometry.coordinates = coords;
                    }
                } catch (// Variables globales
let geoData = { 
    type: "FeatureCollection", 
    features: [
        {
            type: "Feature",
            geometry: {
                type: "Point",
                coordinates: [-97.360245, 21.02]
            },
            properties: {
                Nombre_corto: "TLACOLULA",
                campo: "EXPLORATORIO",
                clasificac: "DISPONIBLE",
                entidad: "VERACRUZ DE IGNACIO DE LA LLAVE",
                estado_act: "INACTIVO",
                fecha_fin: "23-MAY-1984",
                fecha_inic: "21-FEB-1979",
                pozo: "GALO-1"
            }
        },
        {
            type: "Feature",
            geometry: {
                type: "Point",
                coordinates: [-97.361245, 21.03]
            },
            properties: {
                Nombre_corto: "TLACOLULA",
                campo: "EXPLORATORIO",
                clasificac: "DISPONIBLE",
                entidad: "VERACRUZ DE IGNACIO DE LA LLAVE",
                estado_act: "INACTIVO",
                fecha_fin: "23-MAY-1984",
                fecha_inic: "21-FEB-1979",
                pozo: "GALO-2"
            }
        },
        {
            type: "Feature",
            geometry: {
                type: "Point",
                coordinates: [-97.362245, 21.04]
            },
            properties: {
                Nombre_corto: "TLACOLULA",
                campo: "EXPLORATORIO",
                clasificac: "DISPONIBLE",
                entidad: "VERACRUZ DE IGNACIO DE LA LLAVE",
                estado_act: "INACTIVO",
                fecha_fin: "23-MAY-1984",
                fecha_inic: "21-FEB-1979",
                pozo: "GALO-3"
            }
        }
    ]
};

let currentColumns = [];
let columnWidths = {};
let selectedCells = new Set();
let activeCell = null; // Celda que tiene el foco para navegación
let isWrapMode = true;
let undoStack = [];
let redoStack = [];
let maxUndoSteps = 50;
let clipboardData = null;
let isResizing = false;
let startX, startWidth, currentColIndex;
let columnNameMap = {};
let isSelecting = false;
let selectionStartCell = null;
let selectionEndCell = null;
let selectionBox = null;

// Inicializar la aplicación
document.addEventListener('DOMContentLoaded', function() {
    processGeoData();
    setupMenuEvents();
    setupKeyboardEvents();
    setupSelectionEvents();
});

// Configurar eventos de menú
function setupMenuEvents() {
    const menuItems = document.querySelectorAll('.menu-item');
    
    menuItems.forEach(item => {
        item.addEventListener('click', function(e) {
            menuItems.forEach(mi => {
                if (mi !== item) mi.classList.remove('active');
            });
            item.classList.toggle('active');
            e.stopPropagation();
        });
    });
    
    document.addEventListener('click', function() {
        menuItems.forEach(mi => mi.classList.remove('active'));
    });
    
    document.getElementById('openFileOption').addEventListener('click', function() {
        document.getElementById('fileInput').click();
    });
    
    document.getElementById('fileInput').addEventListener('change', function(e) {
        const file = e.target.files[0];
        if (file) {
            const reader = new FileReader();
            reader.onload = function(e) {
                try {
                    geoData = JSON.parse(e.target.result);
                    processGeoData();
                    updateStatus(`Archivo ${file.name} cargado correctamente`);
                    document.getElementById('titleStatus').textContent = `Archivo ${file.name} cargado`;
                    document.getElementById('fileInput').value = '';
                } catch (error) {
                    updateStatus(`Error al procesar el archivo: ${error.message}`, true);
                }
            };
            reader.readAsText(file);
        }
    });
}

// Configurar eventos de teclado mejorados
function setupKeyboardEvents() {
    document.addEventListener('keydown', function(e) {
        // Solo procesar eventos de teclado si no estamos editando un input
        if (document.activeElement.tagName === 'INPUT') {
            return;
        }

        // Undo: Ctrl+Z
        if (e.ctrlKey && e.key === 'z') {
            e.preventDefault();
            undoAction();
        }
        // Redo: Ctrl+Y
        else if (e.ctrlKey && e.key === 'y') {
            e.preventDefault();
            redoAction();
        }
        // Delete: Delete key
        else if (e.key === 'Delete') {
            e.preventDefault();
            deleteSelected();
        }
        // Copy: Ctrl+C
        else if (e.ctrlKey && e.key === 'c') {
            e.preventDefault();
            copySelection();
        }
        // Paste: Ctrl+V
        else if (e.ctrlKey && e.key === 'v') {
            e.preventDefault();
            pasteSelection();
        }
        // Navegación con flechas
        else if (['ArrowUp', 'ArrowDown', 'ArrowLeft', 'ArrowRight'].includes(e.key)) {
            e.preventDefault();
            navigateWithArrows(e.key);
        }
        // Enter para editar celda
        else if (e.key === 'Enter' && activeCell) {
            e.preventDefault();
            startEditingActiveCell();
        }
        // Escape para cancelar selección
        else if (e.key === 'Escape') {
            clearSelection();
        }
    });

    // Manejo especial para inputs
    document.addEventListener('keydown', function(e) {
        if (document.activeElement.tagName === 'INPUT') {
            if (e.key === 'Enter') {
                e.preventDefault();
                document.activeElement.blur();
                if (activeCell) {
                    moveActiveCell('down');
                }
            } else if (e.key === 'Tab') {
                e.preventDefault();
                document.activeElement.blur();
                if (activeCell) {
                    moveActiveCell(e.shiftKey ? 'left' : 'right');
                }
            }
        }
    });
}

// Navegación con flechas
function navigateWithArrows(direction) {
    if (!activeCell) {
        // Si no hay celda activa, activar la primera celda
        setActiveCell(0, 0);
        return;
    }

    const [row, col] = [activeCell.row, activeCell.col];
    
    switch (direction) {
        case 'ArrowUp':
            moveActiveCell('up');
            break;
        case 'ArrowDown':
            moveActiveCell('down');
            break;
        case 'ArrowLeft':
            moveActiveCell('left');
            break;
        case 'ArrowRight':
            moveActiveCell('right');
            break;
    }
}

// Mover celda activa
function moveActiveCell(direction) {
    if (!activeCell) return;

    let newRow = activeCell.row;
    let newCol = activeCell.col;

    switch (direction) {
        case 'up':
            newRow = Math.max(-1, newRow - 1);
            break;
        case 'down':
            newRow = Math.min(geoData.features.length - 1, newRow + 1);
            break;
        case 'left':
            newCol = Math.max(0, newCol - 1);
            break;
        case 'right':
            newCol = Math.min(currentColumns.length - 1, newCol + 1);
            break;
    }

    setActiveCell(newRow, newCol);
}

// Establecer celda activa
function setActiveCell(row, col) {
    // Limpiar celda activa anterior
    document.querySelectorAll('.cell-active').// Variables globales
let geoData = { 
    type: "FeatureCollection", 
    features: [
        {
            type: "Feature",
            geometry: {
                type: "Point",
                coordinates: [-97.360245, 21.02]
            },
            properties: {
                Nombre_corto: "TLACOLULA",
                campo: "EXPLORATORIO",
                clasificac: "DISPONIBLE",
                entidad: "VERACRUZ DE IGNACIO DE LA LLAVE",
                estado_act: "INACTIVO",
                fecha_fin: "23-MAY-1984",
                fecha_inic: "21-FEB-1979",
                pozo: "GALO-1"
            }
        },
        {
            type: "Feature",
            geometry: {
                type: "Point",
                coordinates: [-97.361245, 21.03]
            },
            properties: {
                Nombre_corto: "TLACOLULA",
                campo: "EXPLORATORIO",
                clasificac: "DISPONIBLE",
                entidad: "VERACRUZ DE IGNACIO DE LA LLAVE",
                estado_act: "INACTIVO",
                fecha_fin: "23-MAY-1984",
                fecha_inic: "21-FEB-1979",
                pozo: "GALO-2"
            }
        },
        {
            type: "Feature",
            geometry: {
                type: "Point",
                coordinates: [-97.362245, 21.04]
            },
            properties: {
                Nombre_corto: "TLACOLULA",
                campo: "EXPLORATORIO",
                clasificac: "DISPONIBLE",
                entidad: "VERACRUZ DE IGNACIO DE LA LLAVE",
                estado_act: "INACTIVO",
                fecha_fin: "23-MAY-1984",
                fecha_inic: "21-FEB-1979",
                pozo: "GALO-3"
            }
        }
    ]
};

let currentColumns = [];
let columnWidths = {};
let selectedCells = new Set();
let activeCell = null;
let isWrapMode = true;
let undoStack = [];
let redoStack = [];
let maxUndoSteps = 50;
let clipboardData = null;
let isResizing = false;
let startX, startWidth, currentColIndex;
let columnNameMap = {};
let isSelecting = false;
let selectionStartCell = null;
let selectionEndCell = null;
let selectionBox = null;

// Inicializar la aplicación
document.addEventListener('DOMContentLoaded', function() {
    processGeoData();
    setupMenuEvents();
    setupKeyboardEvents();
    setupSelectionEvents();
});

// Configurar eventos de menú
function setupMenuEvents() {
    const menuItems = document.querySelectorAll('.menu-item');
    
    menuItems.forEach(item => {
        item.addEventListener('click', function(e) {
            menuItems.forEach(mi => {
                if (mi !== item) mi.classList.remove('active');
            });
            item.classList.toggle('active');
            e.stopPropagation();
        });
    });
    
    document.addEventListener('click', function() {
        menuItems.forEach(mi => mi.classList.remove('active'));
    });
    
    document.getElementById('openFileOption').addEventListener('click', function() {
        document.getElementById('fileInput').click();
    });
    
    document.getElementById('fileInput').addEventListener('change', function(e) {
        const file = e.target.files[0];
        if (file) {
            const reader = new FileReader();
            reader.onload = function(e) {
                try {
                    geoData = JSON.parse(e.target.result);
                    processGeoData();
                    updateStatus(`Archivo ${file.name} cargado correctamente`);
                    document.getElementById('titleStatus').textContent = `Archivo ${file.name} cargado`;
                    document.getElementById('fileInput').value = '';
                } catch (error) {
                    updateStatus(`Error al procesar el archivo: ${error.message}`, true);
                }
            };
            reader.readAsText(file);
        }
    });
}

// Configurar eventos de teclado
function setupKeyboardEvents() {
    document.addEventListener('keydown', function(e) {
        if (document.activeElement.tagName === 'INPUT') {
            return;
        }

        if (e.ctrlKey && e.key === 'z') {
            e.preventDefault();
            undoAction();
        } else if (e.ctrlKey && e.key === 'y') {
            e.preventDefault();
            redoAction();
        } else if (e.key === 'Delete') {
            e.preventDefault();
            deleteSelected();
        } else if (e.ctrlKey && e.key === 'c') {
            e.preventDefault();
            copySelection();
        } else if (e.ctrlKey && e.key === 'v') {
            e.preventDefault();
            pasteSelection();
        } else if (['ArrowUp', 'ArrowDown', 'ArrowLeft', 'ArrowRight'].includes(e.key)) {
            e.preventDefault();
            navigateWithArrows(e.key);
        } else if (e.key === 'Enter' && activeCell) {
            e.preventDefault();
            startEditingActiveCell();
        } else if (e.key === 'Escape') {
            clearSelection();
        }
    });

    document.addEventListener('keydown', function(e) {
        if (document.activeElement.tagName === 'INPUT') {
            if (e.key === 'Enter') {
                e.preventDefault();
                document.activeElement.blur();
                if (activeCell) {
                    moveActiveCell('down');
                }
            } else if (e.key === 'Tab') {
                e.preventDefault();
                document.activeElement.blur();
                if (activeCell) {
                    moveActiveCell(e.shiftKey ? 'left' : 'right');
                }
            }
        }
    });
}

// Navegación con flechas
function navigateWithArrows(direction) {
    if (!activeCell) {
        setActiveCell(0, 0);
        return;
    }

    switch (direction) {
        case 'ArrowUp':
            moveActiveCell('up');
            break;
        case 'ArrowDown':
            moveActiveCell('down');
            break;
        case 'ArrowLeft':
            moveActiveCell('left');
            break;
        case 'ArrowRight':
            moveActiveCell('right');
            break;
    }
}

// Mover celda activa
function moveActiveCell(direction) {
    if (!activeCell) return;

    let newRow = activeCell.row;
    let newCol = activeCell.col;

    switch (direction) {
        case 'up':
            newRow = Math.max(-1, newRow - 1);
            break;
        case 'down':
            newRow = Math.min(geoData.features.length - 1, newRow + 1);
            break;
        case 'left':
            newCol = Math.max(0, newCol - 1);
            break;
        case 'right':
            newCol = Math.min(currentColumns.length - 1, newCol + 1);
            break;
    }

    setActiveCell(newRow, newCol);
}

// Establecer celda activa
function setActiveCell(row, col) {
    document.querySelectorAll('.cell-active').forEach(cell => 
        cell.classList.remove('cell-active'));

    activeCell = { row, col };
    
    const cell = document.querySelector(`.sheet-table td[data-row="${row}"][data-col="${col}"]`);
    if (cell) {
        cell.classList.add('cell-active');
        cell.scrollIntoView({ 
            behavior: 'smooth', 
            block: 'nearest', 
            inline: 'nearest' 
        });
    }

    selectedCells.clear();
    selectedCells.add(`${row},${col}`);
    highlightSelectedCells();
}

// Iniciar edición de celda activa
function startEditingActiveCell() {
    if (!activeCell) return;

    const cell = document.querySelector(`.sheet-table td[data-row="${activeCell.row}"][data-col="${activeCell.col}"] input`);
    if (cell) {
        cell.focus();
        cell.select();
    }
}

// Limpiar selección
function clearSelection() {
    selectedCells.clear();
    activeCell = null;
    document.querySelectorAll('.cell-active, .cell-selected, .column-selected, .row-selected').forEach(cell => {
        cell.classList.remove('cell-active', 'cell-selected', 'column-selected', 'row-selected');
    });
}

// Configurar eventos de selección CORREGIDO
function setupSelectionEvents() {
    const sheetContainer = document.getElementById('sheetContainer');
    
    sheetContainer.addEventListener('mousedown', function(e) {
        // CORRECCIÓN: Manejar clicks en encabezados de columna
        if (e.target.tagName === 'TH' && !e.target.classList.contains('corner-header-cell')) {
            const colIndex = parseInt(e.target.dataset.index);
            if (!isNaN(colIndex)) {
                selectColumn(colIndex);
                return;
            }
        }
        
        // CORRECCIÓN: Manejar clicks en números de fila
        if (e.target.classList.contains('row-number-cell')) {
            const rowIndex = parseInt(e.target.dataset.index);
            if (!isNaN(rowIndex)) {
                selectRow(rowIndex);
                return;
            }
        }
        
        if (e.target.classList.contains('column-resizer')) {
            return;
        }
        
        if (e.target.tagName === 'INPUT') {
            const cell = e.target.closest('td');
            if (cell) {
                const row = parseInt(cell.dataset.row);
                const col = parseInt(cell.dataset.col);
                setActiveCell(row, col);
            }
            return;
        }
        
        const cell = e.target.closest('td');
        if (!cell) return;
        
        const row = parseInt(cell.dataset.row);
        const col = parseInt(cell.dataset.col);
        
        setActiveCell(row, col);
        
        isSelecting = true;
        selectionStartCell = { row, col };
        selectionEndCell = { row, col };
        
        if (!e.ctrlKey && !e.shiftKey) {
            selectedCells.clear();
        }
        
        updateSelection(e.ctrlKey, e.shiftKey);
        e.preventDefault();
    });
    
    sheetContainer.addEventListener('mousemove', function(e) {
        if (!isSelecting) return;
        
        const cell = e.target.closest('td');
        if (!cell) return;
        
        const row = parseInt(cell.dataset.row);
        const col = parseInt(cell.dataset.col);
        selectionEndCell = { row, col };
        
        updateSelection(e.ctrlKey, e.shiftKey);
    });
    
    document.addEventListener('mouseup', function(e) {
        if (!isSelecting) return;
        isSelecting = false;
        selectionStartCell = null;
        selectionEndCell = null;
    });

    sheetContainer.addEventListener('dblclick', function(e) {
        const cell = e.target.closest('td');
        if (cell) {
            const input = cell.querySelector('input');
            if (input) {
                input.focus();
                input.select();
            }
        }
    });
}

// Actualizar selección
function updateSelection(ctrlKey, shiftKey) {
    if (!selectionStartCell || !selectionEndCell) return;
    
    const newSelection = new Set();
    const minRow = Math.min(selectionStartCell.row, selectionEndCell.row);
    const maxRow = Math.max(selectionStartCell.row, selectionEndCell.row);
    const minCol = Math.min(selectionStartCell.col, selectionEndCell.col);
    const maxCol = Math.max(selectionStartCell.col, selectionEndCell.col);
    
    for (let r = minRow; r <= maxRow; r++) {
        for (let c = minCol; c <= maxCol; c++) {
            newSelection.add(`${r},${c}`);
        }
    }
    
    if (ctrlKey) {
        newSelection.forEach(cell => selectedCells.add(cell));
    } else {
        selectedCells = newSelection;
    }
    
    highlightSelectedCells();
}

// Resaltar celdas seleccionadas
function highlightSelectedCells() {
    // Limpiar selecciones anteriores
    document.querySelectorAll('.sheet-table td.cell-selected, .sheet-table th.column-selected, .sheet-table .row-number-cell.row-selected').forEach(cell => 
        cell.classList.remove('cell-selected', 'column-selected', 'row-selected'));
    
    selectedCells.forEach(cellKey => {
        const [row, col] = cellKey.split(',').map(Number);
        const cell = document.querySelector(`.sheet-table td[data-row="${row}"][data-col="${col}"]`);
        if (cell) {
            cell.classList.add('cell-selected');
        }
    });
}

// Procesar datos GeoJSON
function processGeoData() {
    saveToHistory();
    
    const propertyKeys = new Set();
    geoData.features.forEach(feature => {
        if (feature.properties) {
            Object.keys(feature.properties).forEach(key => {
                propertyKeys.add(key);
            });
        }
    });
    
    currentColumns = ['geometry', ...Array.from(propertyKeys).sort()];
    selectedCells.clear();
    activeCell = null;
    
    rebuildSheet();
    updateUndoRedoButtons();
    updateWrapButton();
}

// Convertir índice a nombre de columna Excel
function getExcelColumnName(index) {
    let columnName = '';
    
    while (index >= 0) {
        columnName = String.fromCharCode(65 + (index % 26)) + columnName;
        index = Math.floor(index / 26) - 1;
    }
    
    return columnName;
}

// Reconstruir tabla CORREGIDO
function rebuildSheet() {
    const table = document.getElementById('sheetTable');
    const headerRow = document.getElementById('columnHeadersRow');
    const tableBody = document.getElementById('sheetTableBody');
    
    headerRow.innerHTML = '<th class="corner-header-cell"></th>';
    tableBody.innerHTML = '';
    
    // Crear encabezados de columna CORREGIDO
    currentColumns.forEach((col, colIndex) => {
        const th = document.createElement('th');
        th.dataset.index = colIndex; // IMPORTANTE: Agregar dataset.index
        
        const excelColName = getExcelColumnName(colIndex);
        th.textContent = excelColName;
        th.title = col;
        
        if (!columnNameMap[excelColName]) {
            columnNameMap[excelColName] = col;
        }
        
        const colWidth = columnWidths[colIndex] || 120;
        th.style.width = colWidth + 'px';
        th.style.minWidth = colWidth + 'px';
        th.style.maxWidth = colWidth + 'px';
        
        // CORRECCIÓN: Event listener para selección de columna
        th.addEventListener('click', function(e) {
            if (!e.target.classList.contains('column-resizer')) {
                selectColumn(colIndex);
            }
        });
        
        const resizer = document.createElement('div');
        resizer.className = 'column-resizer';
        resizer.dataset.colIndex = colIndex;
        resizer.addEventListener('mousedown', startColumnResize);
        th.appendChild(resizer);
        
        headerRow.appendChild(th);
    });
    
    // Crear fila para nombres de propiedades
    const propertyNameRow = document.createElement('tr');
    propertyNameRow.className = 'property-name-row';
    
    const propertyRowNumTd = document.createElement('td');
    propertyRowNumTd.className = 'row-number-cell';
    propertyRowNumTd.textContent = '1';
    propertyNameRow.appendChild(propertyRowNumTd);
    
    currentColumns.forEach((col, colIndex) => {
        const td = document.createElement('td');
        td.dataset.row = -1;
        td.dataset.col = colIndex;
        
        const colWidth = columnWidths[colIndex] || 120;
        td.style.width = colWidth + 'px';
        td.style.minWidth = colWidth + 'px';
        td.style.maxWidth = colWidth + 'px';
        
        if (isWrapMode) {
            td.classList.add('wrap-mode');
        }
        
        const input = document.createElement('input');
        input.type = 'text';
        input.dataset.row = -1;
        input.dataset.col = colIndex;
        input.value = col;
        
        if (col.length > 15) {
            td.setAttribute('data-tooltip', col);
        }
        
        input.addEventListener('change', function(e) {
            saveToHistory();
            
            const oldName = currentColumns[colIndex];
            const newName = e.target.value.trim();
            
            if (!newName) {
                updateStatus("El nombre de la propiedad no puede estar vacío", true);
                e.target.value = oldName;
                return;
            }
            
            if (newName !== oldName && currentColumns.includes(newName)) {
                updateStatus("Ya existe una propiedad con ese nombre", true);
                e.target.value = oldName;
                return;
            }
            
            if (oldName === 'geometry') {
                updateStatus("No se puede cambiar el nombre de la columna geometry", true);
                e.target.value = 'geometry';
                return;
            }
            
            currentColumns[colIndex] = newName;
            
            const excelColName = getExcelColumnName(colIndex);
            columnNameMap[excelColName] = newName;
            
            geoData.features.forEach(feature => {
                if (feature.properties && feature.properties[oldName] !== undefined) {
                    feature.properties[newName] = feature.properties[oldName];
                    if (oldName !== newName) {
                        delete feature.properties[oldName];
                    }
                }
            });
            
            if (newName.length > 15) {
                td.setAttribute('data-tooltip', newName);
            } else {
                td.removeAttribute('data-tooltip');
            }
            
            updateStatus(`Propiedad renombrada de "${oldName}" a "${newName}"`);
        });
        
        td.appendChild(input);
        propertyNameRow.appendChild(td);
    });
    
    tableBody.appendChild(propertyNameRow);
    
    // Crear filas de datos CORREGIDO
    geoData.features.forEach((feature, rowIndex) => {
        const tr = document.createElement('tr');
        
        const rowNumTd = document.createElement('td');
        rowNumTd.className = 'row-number-cell';
        rowNumTd.textContent = rowIndex + 2;
        rowNumTd.dataset.index = rowIndex; // IMPORTANTE: Agregar dataset.index
        // CORRECCIÓN: Event listener para selección de fila
        rowNumTd.addEventListener('click', function(e) {
            selectRow(rowIndex);
        });
        tr.appendChild(rowNumTd);
        
        currentColumns.forEach((col, colIndex) => {
            const td = document.createElement('td');
            td.dataset.row = rowIndex;
            td.dataset.col = colIndex;
            
            const colWidth = columnWidths[colIndex] || 120;
            td.style.width = colWidth + 'px';
            td.style.minWidth = colWidth + 'px';
            td.style.maxWidth = colWidth + 'px';
            
            if (col === 'geometry') td.classList.add('coords-cell');
            
            if (isWrapMode) {
                td.classList.add('wrap-mode');
            }
            
            const input = document.createElement('input');
            input.type = 'text';
            input.dataset.row = rowIndex;
            input.dataset.col = colIndex;
            
            if (col === 'geometry') {
                if (feature.geometry && feature.geometry.coordinates) {
                    input.value = feature.geometry.coordinates.join(', ');
                } else {
                    input.value = '0, 0';
                }
                
                input.addEventListener('change', function(e) {
                    try {
                        saveToHistory();
                        
                        const coords = e.target.value.split(',').map(Number);
                        if (coords.length === 2 && !isNaN(coords[0]) && !isNaN(coords[1])) {
                            if (!feature.geometry) feature.geometry = { type: "Point" };
                            feature.geometry.coordinates = coords;
                            updateStatus(`Coordenadas actualizadas en fila ${rowIndex + 2}`);
                        } else {
                            throw new Error("Formato de coordenadas inválido");
                        }
                    } catch (error) {
                        updateStatus(`Error: ${error.message}`, true);
                        e.target.value = feature.geometry?.coordinates?.join(', ') || '0, 0';
                    }
                });
            } else {
                // CORRECCIÓN CRÍTICA: Asegurar que la propiedad existe
                const value = feature.properties && feature.properties.hasOwnProperty(col) 
                    ? feature.properties[col] 
                    : '';
                input.value = value;
                
                if (String(value).length > 15) {
                    td.setAttribute('data-tooltip', value);
                }
                
                input.addEventListener('change', function(e) {
                    saveToHistory();
                    
                    if (!feature.properties) feature.properties = {};
                    feature.properties[col] = e.target.value;
                    
                    if (String(e.target.value).length > 15) {
                        td.setAttribute('data-tooltip', e.target.value);
                    } else {
                        td.removeAttribute('data-tooltip');
                    }
                    
                    updateStatus(`Propiedad ${col} actualizada en fila ${rowIndex + 2}`);
                });
            }
            
            td.appendChild(input);
            tr.appendChild(td);
        });
        
        tableBody.appendChild(tr);
    });
    
    highlightSelectedCells();
}

// Seleccionar fila CORREGIDO
function selectRow(rowIndex) {
    clearSelection();
    
    // Marcar todas las celdas de la fila como seleccionadas
    for (let colIndex = 0; colIndex < currentColumns.length; colIndex++) {
        selectedCells.add(`${rowIndex},${colIndex}`);
    }
    
    // Marcar el número de fila como seleccionado
    const rowNumCell = document.querySelector(`.row-number-cell[data-index="${rowIndex}"]`);
    if (rowNumCell) {
        rowNumCell.classList.add('row-selected');
    }
    
    setActiveCell(rowIndex, 0);
    highlightSelectedCells();
    updateStatus(`Fila ${rowIndex + 2} seleccionada`);
}

// Seleccionar columna CORREGIDO
function selectColumn(colIndex) {
    clearSelection();
    
    // Seleccionar todas las celdas de la columna (incluyendo la fila de nombres)
    for (let rowIndex = -1; rowIndex < geoData.features.length; rowIndex++) {
        selectedCells.add(`${rowIndex},${colIndex}`);
    }
    
    // Marcar el encabezado de columna como seleccionado
    const headerCell = document.querySelector(`th[data-index="${colIndex}"]`);
    if (headerCell) {
        headerCell.classList.add('column-selected');
    }
    
    setActiveCell(0, colIndex);
    highlightSelectedCells();
    
    const excelColName = getExcelColumnName(colIndex);
    const propertyName = columnNameMap[excelColName];
    updateStatus(`Columna ${excelColName} (${propertyName}) seleccionada`);
}

// Funciones de redimensionamiento
function startColumnResize(e) {
    e.preventDefault();
    e.stopPropagation();
    
    isResizing = true;
    currentColIndex = parseInt(this.dataset.colIndex);
    startX = e.pageX;
    startWidth = columnWidths[currentColIndex] || 120;
    
    document.addEventListener('mousemove', doColumnResize);
    document.addEventListener('mouseup', stopColumnResize);
    document.body.style.cursor = 'col-resize';
}

function doColumnResize(e) {
    if (!isResizing) return;
    const delta = e.pageX - startX;
    const newWidth = Math.max(60, startWidth + delta);
}

function stopColumnResize(e) {
    if (!isResizing) return;
    
    saveToHistory();
    const delta = e.pageX - startX;
    const newWidth = Math.max(60, startWidth + delta);
    columnWidths[currentColIndex] = newWidth;
    
    document.removeEventListener('mousemove', doColumnResize);
    document.removeEventListener('mouseup', stopColumnResize);
    document.body.style.cursor = '';
    
    rebuildSheet();
    isResizing = false;
    
    const excelColName = getExcelColumnName(currentColIndex);
    const propertyName = columnNameMap[excelColName];
    updateStatus(`Columna ${excelColName} (${propertyName}) redimensionada a ${newWidth}px`);
}

// Agregar columna CORREGIDO
function addColumn() {
    saveToHistory();
    
    let newColName = "Column1";
    let counter = 1;
    while (currentColumns.includes(newColName)) {
        counter++;
        newColName = `Column${counter}`;
    }
    
    currentColumns.push(newColName);
    
    const newColIndex = currentColumns.length - 1;
    const excelColName = getExcelColumnName(newColIndex);
    columnNameMap[excelColName] = newColName;
    
    // CORRECCIÓN CRÍTICA: Agregar la propiedad a TODAS las features
    geoData.features.forEach(feature => {
        if (!feature.properties) feature.properties = {};
        feature.properties[newColName] = "";
    });
    
    rebuildSheet();
    updateStatus(`Nueva columna "${newColName}" (${excelColName}) agregada`);
}

// Agregar fila
function addRow() {
    saveToHistory();
    
    const newFeature = {
        type: "Feature",
        geometry: {
            type: "Point",
            coordinates: [0, 0]
        },
        properties: {}
    };
    
    // Agregar todas las propiedades existentes con valores vacíos
    currentColumns.forEach(col => {
        if (col !== 'geometry') {
            newFeature.properties[col] = "";
        }
    });
    
    geoData.features.push(newFeature);
    rebuildSheet();
    updateStatus(`Nueva fila agregada`);
}

// Eliminar selección
function deleteSelected() {
    if (selectedCells.size === 0) {
        updateStatus("No hay nada seleccionado para eliminar", true);
        return;
    }
    
    saveToHistory();
    
    const rowsToDelete = new Set();
    const colsToDelete = new Set();
    const cellsToClear = new Set();
    
    selectedCells.forEach(cellKey => {
        const [row, col] = cellKey.split(',').map(Number);
        if (row === -1) {
            colsToDelete.add(col);
        } else {
            cellsToClear.add(cellKey);
        }
    });
    
    if (colsToDelete.size > 0) {
        const colsToDeleteArray = Array.from(colsToDelete).sort((a, b) => b - a);
        
        colsToDeleteArray.forEach(colIndex => {
            const colName = currentColumns[colIndex];
            
            if (colName === 'geometry') {
                updateStatus("No se puede eliminar la columna geometry", true);
                return;
            }
            
            geoData.features.forEach(feature => {
                if (feature.properties && feature.properties[colName] !== undefined) {
                    delete feature.properties[colName];
                }
            });
            
            currentColumns.splice(colIndex, 1);
            delete columnWidths[colIndex];
            
            const newColumnWidths = {};
            Object.keys(columnWidths).forEach(key => {
                const keyNum = parseInt(key);
                if (keyNum < colIndex) {
                    newColumnWidths[keyNum] = columnWidths[keyNum];
                } else if (keyNum > colIndex) {
                    newColumnWidths[keyNum - 1] = columnWidths[keyNum];
                }
            });
            columnWidths = newColumnWidths;
            
            const newColumnNameMap = {};
            currentColumns.forEach((name, index) => {
                const excelName = getExcelColumnName(index);
                newColumnNameMap[excelName] = name;
            });
            columnNameMap = newColumnNameMap;
        });
        
        updateStatus(`${colsToDeleteArray.length} columna(s) eliminada(s)`);
    }
    
    if (colsToDelete.size === 0 && cellsToClear.size > 0) {
        cellsToClear.forEach(cellKey => {
            const [row, col] = cellKey.split(',').map(Number);
            const colName = currentColumns[col];
            
            if (row >= 0 && row < geoData.features.length) {
                if (colName !== 'geometry') {
                    if (geoData.features[row].properties) {
                        geoData.features[row].properties[colName] = "";
                    }
                }
            }
        });
        updateStatus(`${cellsToClear.size} celda(s) limpiada(s)`);
    }
    
    selectedCells.clear();
    activeCell = null;
    rebuildSheet();
}

// Copiar selección
function copySelection() {
    if (selectedCells.size === 0) {
        updateStatus("No hay nada seleccionado para copiar", true);
        return;
    }
    
    clipboardData = {
        type: 'cells',
        data: []
    };
    
    const rows = new Set();
    const cols = new Set();
    selectedCells.forEach(cellKey => {
        const [row, col] = cellKey.split(',').map(Number);
        rows.add(row);
        cols.add(col);
    });
    
    const minRow = Math.min(...Array.from(rows));
    const maxRow = Math.max(...Array.from(rows));
    const minCol = Math.min(...Array.from(cols));
    const maxCol = Math.max(...Array.from(cols));
    
    for (let r = minRow; r <= maxRow; r++) {
        const rowData = [];
        for (let c = minCol; c <= maxCol; c++) {
            if (selectedCells.has(`${r},${c}`)) {
                const cell = document.querySelector(`.sheet-table td[data-row="${r}"][data-col="${c}"] input`);
                rowData.push(cell ? cell.value : "");
            } else {
                rowData.push(null);
            }
        }
        clipboardData.data.push(rowData);
    }
    
    updateStatus(`Selección copiada (${clipboardData.data.length}x${clipboardData.data[0].length})`);
}

// Pegar selección
function pasteSelection() {
    if (!clipboardData || clipboardData.type !== 'cells') {
        updateStatus("No hay datos de celdas para pegar", true);
        return;
    }
    
    if (selectedCells.size === 0) {
        updateStatus("Selecciona una celda para empezar a pegar", true);
        return;
    }
    
    const startCellKey = Array.from(selectedCells)[0];
    const [startRow, startCol] = startCellKey.split(',').map(Number);
    
    saveToHistory();
    
    clipboardData.data.forEach((rowData, rOffset) => {
        rowData.forEach((cellValue, cOffset) => {
            if (cellValue !== null) {
                const targetRow = startRow + rOffset;
                const targetCol = startCol + cOffset;
                
                if (targetRow >= -1 && targetRow < geoData.features.length && targetCol >= 0 && targetCol < currentColumns.length) {
                    const colName = currentColumns[targetCol];
                    
                    if (targetRow === -1) {
                        // No permitir pegar en la fila de nombres por ahora
                    } else {
                        if (colName === 'geometry') {
                            try {
                                const coords = cellValue.split(',').map(Number);
                                if (coords.length === 2 && !isNaN(coords[0]) && !isNaN(coords[1])) {
                                    if (!geoData.features[targetRow].geometry) geoData.features[targetRow].geometry = { type: "Point" };
                                    geoData.features[targetRow].geometry.coordinates = coords;
                                }
                            } catch (error) {
                                console.error("Error al pegar coordenadas:", error);
                            }
                        } else {
                            if (!geoData.features[targetRow].properties) geoData.features[targetRow].properties = {};
                            geoData.features[targetRow].properties[colName] = cellValue;
                        }
                    }
                }
            }
        });
    });
    
    rebuildSheet();
    updateStatus(`Datos pegados`);
}

// Toggle wrap mode
function toggleWrapMode() {
    isWrapMode = !isWrapMode;
    updateWrapButton();
    
    const cells = document.querySelectorAll('.sheet-table td');
    cells.forEach(cell => {
        if (isWrapMode) {
            cell.classList.add('wrap-mode');
        } else {
            cell.classList.remove('wrap-mode');
        }
    });
    
    updateStatus(`Modo wrap ${isWrapMode ? 'activado' : 'desactivado'}`);
}

function updateWrapButton() {
    const wrapButton = document.getElementById('wrapButton');
    if (isWrapMode) {
        wrapButton.classList.add('active');
    } else {
        wrapButton.classList.remove('active');
    }
}

// Historial Undo/Redo
function saveToHistory() {
    undoStack.push({
        geoData: JSON.parse(JSON.stringify(geoData)),
        currentColumns: [...currentColumns],
        columnWidths: {...columnWidths},
        selectedCells: new Set(selectedCells),
        columnNameMap: {...columnNameMap}
    });
    
    if (undoStack.length > maxUndoSteps) {
        undoStack.shift();
    }
    
    redoStack = [];
    updateUndoRedoButtons();
}

function undoAction() {
    if (undoStack.length === 0) {
        updateStatus("No hay acciones para deshacer", true);
        return;
    }
    
    redoStack.push({
        geoData: JSON.parse(JSON.stringify(geoData)),
        currentColumns: [...currentColumns],
        columnWidths: {...columnWidths},
        selectedCells: new Set(selectedCells),
        columnNameMap: {...columnNameMap}
    });
    
    const prevState = undoStack.pop();
    geoData = prevState.geoData;
    currentColumns = prevState.currentColumns;
    columnWidths = prevState.columnWidths;
    selectedCells = prevState.selectedCells;
    columnNameMap = prevState.columnNameMap;
    activeCell = null;
    
    rebuildSheet();
    updateUndoRedoButtons();
    updateStatus("Acción deshecha");
}

function redoAction() {
    if (redoStack.length === 0) {
        updateStatus("No hay acciones para rehacer", true);
        return;
    }
    
    undoStack.push({
        geoData: JSON.parse(JSON.stringify(geoData)),
        currentColumns: [...currentColumns],
        columnWidths: {...columnWidths},
        selectedCells: new Set(selectedCells),
        columnNameMap: {...columnNameMap}
    });
    
    const nextState = redoStack.pop();
    geoData = nextState.geoData;
    currentColumns = nextState.currentColumns;
    columnWidths = nextState.columnWidths;
    selectedCells = nextState.selectedCells;
    columnNameMap = nextState.columnNameMap;
    activeCell = null;
    
    rebuildSheet();
    updateUndoRedoButtons();
    updateStatus("Acción rehecha");
}

function updateUndoRedoButtons() {
    const undoButton = document.getElementById('undoButton');
    const redoButton = document.getElementById('redoButton');
    const undoOption = document.getElementById('undoOption');
    const redoOption = document.getElementById('redoOption');
    
    if (undoStack.length === 0) {
        undoButton.classList.add('disabled');
        undoOption.classList.add('disabled');
    } else {
        undoButton.classList.remove('disabled');
        undoOption.classList.remove('disabled');
    }
    
    if (redoStack.length === 0) {
        redoButton.classList.add('disabled');
        redoOption.classList.add('disabled');
    } else {
        redoButton.classList.remove('disabled');
        redoOption.classList.remove('disabled');
    }
}

// Guardar y descargar separados
function saveData() {
    const inputs = document.querySelectorAll('.sheet-table input');
    inputs.forEach(input => {
        const row = parseInt(input.dataset.row);
        const col = parseInt(input.dataset.col);
        const colName = currentColumns[col];
        
        if (row === -1) {
            if (colName !== input.value.trim() && input.value.trim() !== '') {
                currentColumns[col] = input.value.trim();
                const excelColName = getExcelColumnName(col);
                columnNameMap[excelColName] = input.value.trim();
            }
        } else if (row >= 0 && row < geoData.features.length) {
            const feature = geoData.features[row];
            
            if (colName === 'geometry') {
                try {
                    const coords = input.value.split(',').map(Number);
                    if (coords.length === 2 && !isNaN(coords[0]) && !isNaN(coords[1])) {
                        if (!feature.geometry) feature.geometry = { type: "Point" };
                        feature.geometry.coordinates = coords;
                    }
                } catch (error) {
                    console.error("Error al guardar coordenadas:", error);
                }
            } else {
                if (!feature.properties) feature.properties = {};
                feature.properties[colName] = input.value;
            }
        }
    });
    
    updateStatus("Datos guardados en memoria");
}

function downloadFile() {
    saveData();
    
    const dataToExport = JSON.parse(JSON.stringify(geoData));
    
    dataToExport.features.forEach(feature => {
        const newProperties = {};
        currentColumns.forEach(colName => {
            if (colName !== 'geometry') {
                newProperties[colName] = feature.properties && feature.properties[colName] !== undefined 
                    ? feature.properties[colName] 
                    : "";
            }
        });
        feature.properties = newProperties;
    });
    
    const blob = new Blob([JSON.stringify(dataToExport, null, 2)], {type: 'application/json'});
    const a = document.createElement('a');
    a.href = URL.createObjectURL(blob);
    a.download = 'geojson_export.geojson';
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    URL.revokeObjectURL(a.href);
    
    updateStatus("Archivo descargado como geojson_export.geojson");
}

function updateStatus(message, isError = false) {
    const statusBar = document.getElementById('statusBar');
    statusBar.textContent = message;
    
    if (isError) {
        statusBar.classList.add('error');
    } else {
        statusBar.classList.remove('error');
    }
}