let collarData = [];
let assayData = [];

const fileInput = document.getElementById('fileInput');
const holeSelect = document.getElementById('holeSelect');
const welcomeMessage = document.getElementById('welcomeMessage');
const tableContainer = document.getElementById('tableContainer');
const assayBody = document.getElementById('assayBody');
const appContainer = document.getElementById('app');
const toggleBtn = document.getElementById('toggleAbove');
const summaryBtn = document.getElementById('showSummary');
const showOreBtn = document.getElementById('showOreCalc');
const showDiagramBtn = document.getElementById('showDiagram');
const modal = document.getElementById('summaryModal');
const oreModal = document.getElementById('oreModal');
const diagramModal = document.getElementById('diagramModal');
const closeModalBtns = document.querySelectorAll('.close-modal');
const summaryStats = document.getElementById('summaryStats');
const availableOnlyCheckbox = document.getElementById('availableOnly');
const oreAvailableOnlyCheckbox = document.getElementById('oreAvailableOnly');
const oreResults = document.getElementById('oreResults');
const chartWidthSlider = document.getElementById('chartWidth');
const chartHeightSlider = document.getElementById('chartHeight');
const chartWrapper = document.getElementById('chartWrapper');
const widthValue = document.getElementById('widthValue');
const heightValue = document.getElementById('heightValue');
const niMaxInput = document.getElementById('niMax');
const othersMaxInput = document.getElementById('othersMax');
const fullscreenBtn = document.getElementById('toggleFullscreen');
const exportBtn = document.getElementById('exportJPG');

let activeFilteredData = [];
let assayChart = null;

fileInput.addEventListener('change', handleFile);
holeSelect.addEventListener('change', updateTable);
toggleBtn.addEventListener('click', toggleMined);
summaryBtn.addEventListener('click', showSummary);
showOreBtn.addEventListener('click', calculateOre);
showDiagramBtn.addEventListener('click', showDiagram);
availableOnlyCheckbox.addEventListener('change', showSummary);
oreAvailableOnlyCheckbox.addEventListener('change', calculateOre);

fullscreenBtn.addEventListener('click', () => {
    if (!document.fullscreenElement) {
        diagramModal.requestFullscreen().catch(err => {
            alert(`Error attempting to enable full-screen mode: ${err.message}`);
        });
    } else {
        document.exitFullscreen();
    }
});

exportBtn.addEventListener('click', () => {
    const canvas = document.getElementById('assayChart');
    const selectedHoleId = holeSelect.value;
    
    // Create a temporary canvas to add a white background for JPG
    const tempCanvas = document.createElement('canvas');
    tempCanvas.width = canvas.width;
    tempCanvas.height = canvas.height;
    const tctx = tempCanvas.getContext('2d');
    
    tctx.fillStyle = "#ffffff";
    tctx.fillRect(0, 0, tempCanvas.width, tempCanvas.height);
    tctx.drawImage(canvas, 0, 0);
    
    const link = document.createElement('a');
    link.download = `Assay_Diagram_${selectedHoleId}.jpg`;
    link.href = tempCanvas.toDataURL('image/jpeg', 1.0);
    link.click();
});

chartWidthSlider.addEventListener('input', (e) => {
    const w = e.target.value;
    chartWrapper.style.width = `${w}px`;
    widthValue.textContent = `${w}px`;
    if (assayChart) assayChart.resize();
});

chartHeightSlider.addEventListener('input', (e) => {
    const h = e.target.value;
    chartWrapper.style.height = `${h}px`;
    heightValue.textContent = `${h}px`;
    if (assayChart) assayChart.resize();
});

niMaxInput.addEventListener('input', () => {
    if (assayChart) {
        assayChart.options.scales.xNi.max = parseFloat(niMaxInput.value) || 3.5;
        assayChart.update();
    }
});

othersMaxInput.addEventListener('input', () => {
    if (assayChart) {
        assayChart.options.scales.xOthers.max = parseFloat(othersMaxInput.value) || 70;
        assayChart.update();
    }
});

document.addEventListener('change', (e) => {
    if (e.target.classList.contains('dataset-toggle')) {
        const index = e.target.dataset.index;
        if (assayChart) {
            assayChart.setDatasetVisibility(index, e.target.checked);
            assayChart.update();
        }
    }
});

closeModalBtns.forEach(btn => {
    btn.addEventListener('click', () => {
        modal.classList.add('hidden');
        oreModal.classList.add('hidden');
        diagramModal.classList.add('hidden');
    });
});

window.addEventListener('click', (e) => {
    if (e.target === modal) modal.classList.add('hidden');
    if (e.target === oreModal) oreModal.classList.add('hidden');
    if (e.target === diagramModal) diagramModal.classList.add('hidden');
});

function toggleMined() {
    appContainer.classList.toggle('dim-mined');
    toggleBtn.classList.toggle('active');
    toggleBtn.textContent = appContainer.classList.contains('dim-mined') ? '🌕 Show Mined' : '🌑 Dim Mined';
}

function handleFile(e) {
    const file = e.target.files[0];
    if (!file) return;

    const reader = new FileReader();
    reader.onload = function (e) {
        const data = new Uint8Array(e.target.result);
        const workbook = XLSX.read(data, { type: 'array' });

        processWorkbook(workbook);
    };
    reader.readAsArrayBuffer(file);
}

function processWorkbook(workbook) {
    const sheetNames = workbook.SheetNames;

    // Find sheets by pattern or specific names discovered earlier
    const collarSheetName = sheetNames.find(n => n.toLowerCase().includes('collar'));
    const assaySheetName = sheetNames.find(n => n.toLowerCase().includes('assay'));

    if (!collarSheetName || !assaySheetName) {
        alert('Could not find both "collar" and "assay" sheets in this Excel file.');
        return;
    }

    const collarSheet = workbook.Sheets[collarSheetName];
    const assaySheet = workbook.Sheets[assaySheetName];

    collarData = XLSX.utils.sheet_to_json(collarSheet);
    assayData = XLSX.utils.sheet_to_json(assaySheet);

    populateHoleSelector();
}

function populateHoleSelector() {
    // Collect unique Hole IDs from collar data (prioritize collar as per request)
    // The explorer found columns 'Hole ID' in collar and 'HoleID' in assay
    const holeIds = [...new Set(collarData.map(row => row['Hole ID'] || row['HoleID']).filter(Boolean))];
    holeIds.sort();

    holeSelect.innerHTML = '<option value="">Select a Hole ID...</option>';
    holeIds.forEach(id => {
        const option = document.createElement('option');
        option.value = id;
        option.textContent = id;
        holeSelect.appendChild(option);
    });

    holeSelect.disabled = false;
    document.querySelector('.btn-file').textContent = '✅ Data Loaded';
}

function updateTable() {
    const selectedHoleId = holeSelect.value;
    if (!selectedHoleId) {
        welcomeMessage.classList.remove('hidden');
        tableContainer.classList.add('hidden');
        return;
    }

    welcomeMessage.classList.add('hidden');
    tableContainer.classList.remove('hidden');

    // Get z value for elevation calculation
    const collarRow = collarData.find(row => (row['Hole ID'] || row['HoleID']) === selectedHoleId);
    const collarZ = collarRow ? (parseFloat(collarRow['z']) || 0) : 0;

    // Filter and sort assay data
    activeFilteredData = assayData.filter(row => (row['HoleID'] || row['Hole ID']) === selectedHoleId);

    // Sort by FROM ascending
    activeFilteredData.sort((a, b) => (parseFloat(a['From']) || 0) - (parseFloat(b['From']) || 0));

    summaryBtn.disabled = activeFilteredData.length === 0;
    showOreBtn.disabled = activeFilteredData.length === 0;
    showDiagramBtn.disabled = activeFilteredData.length === 0;
    renderRows(activeFilteredData, collarZ, selectedHoleId);
}

function showSummary() {
    if (activeFilteredData.length === 0) return;

    const selectedHoleId = holeSelect.value;
    const modalTitle = modal.querySelector('h2');
    modalTitle.textContent = `Assay Summary: ${selectedHoleId}`;

    let targetData = activeFilteredData;
    
    // Filter by Available Materials (Below Topo) if enabled
    if (availableOnlyCheckbox.checked) {
        targetData = activeFilteredData.filter(row => {
            const topoKey = row['Topo Position'] ? 'Topo Position' : 'Topo';
            const topoVal = (row[topoKey] || '').toLowerCase();
            return topoVal.includes('below');
        });
    }

    if (targetData.length === 0) {
        summaryStats.innerHTML = '<p class="empty-state">No data matches the selected criteria.</p>';
        modal.classList.remove('hidden');
        return;
    }

    const chemCols = ['Ni', 'Co', 'Fe', 'SiO2', 'CaO', 'MgO', 'Al2O3', 'SiMa'];
    const stats = {};

    chemCols.forEach(col => {
        const vals = targetData.map(d => parseFloat(d[col])).filter(v => !isNaN(v));
        if (vals.length > 0) {
            const min = Math.min(...vals);
            const max = Math.max(...vals);
            const mean = vals.reduce((a, b) => a + b, 0) / vals.length;
            
            // Variance: sum((x - mean)^2) / n
            const variance = vals.reduce((a, b) => a + Math.pow(b - mean, 2), 0) / vals.length;
            const stdDev = Math.sqrt(variance);

            stats[col] = { min, max, mean, variance, stdDev };
        } else {
            stats[col] = null;
        }
    });

    summaryStats.innerHTML = `
        <div class="summary-table-container">
            <table class="summary-table">
                <thead>
                    <tr>
                        <th style="text-align: left;">Element</th>
                        <th>Min</th>
                        <th>Max</th>
                        <th>Average</th>
                        <th>Variance</th>
                        <th>Std Dev</th>
                    </tr>
                </thead>
                <tbody>
                    ${chemCols.map(col => {
                        const s = stats[col];
                        if (!s) return `<tr><td style="text-align: left;"><strong>${col}</strong></td><td colspan="5" style="text-align: center;">No Data</td></tr>`;
                        return `
                            <tr>
                                <td style="text-align: left;"><strong>${col}</strong></td>
                                <td>${s.min.toFixed(2)}</td>
                                <td>${s.max.toFixed(2)}</td>
                                <td>${s.mean.toFixed(2)}</td>
                                <td>${s.variance.toFixed(2)}</td>
                                <td>${s.stdDev.toFixed(2)}</td>
                            </tr>
                        `;
                    }).join('')}
                </tbody>
            </table>
        </div>
        <p style="font-size: 0.75rem; color: #777; margin-top: 1rem;">* Calculations based on arithmetic mean.</p>
    `;

    modal.classList.remove('hidden');
}

function calculateOre() {
    if (activeFilteredData.length === 0) return;

    const selectedHoleId = holeSelect.value;
    const modalTitle = oreModal.querySelector('h2');
    modalTitle.textContent = `Ore Calculation: ${selectedHoleId}`;

    let targetData = activeFilteredData;
    if (oreAvailableOnlyCheckbox.checked) {
        targetData = activeFilteredData.filter(row => {
            const topoKey = row['Topo Position'] ? 'Topo Position' : 'Topo';
            const topoVal = (row[topoKey] || '').toLowerCase();
            return topoVal.includes('below');
        });
    }

    if (targetData.length === 0) {
        oreResults.innerHTML = '<p class="empty-state">No data matches the selected criteria.</p>';
        oreModal.classList.remove('hidden');
        return;
    }

    let totalThick = 0;
    let oreThick = 0;
    let nonOreThick = 0;
    let obThick = 0;
    let oreNiSum = 0;
    let oreCount = 0;

    targetData.forEach(row => {
        const from = parseFloat(row['From']) || 0;
        const to = parseFloat(row['To']) || 0;
        const thick = Math.max(0, to - from);
        const ni = parseFloat(row['Ni']) || 0;
        const litho = (row['Zonasi'] || '').toUpperCase();

        totalThick += thick;

        if (ni >= 1) {
            oreThick += thick;
            oreNiSum += ni;
            oreCount++;
        } else {
            nonOreThick += thick;
            // OB: Ni < 1 AND Litho is not BRK
            if (litho !== 'BRK') {
                obThick += thick;
            }
        }
    });

    const oreAvgNi = oreCount > 0 ? (oreNiSum / oreCount) : 0;
    const sr = oreThick > 0 ? (obThick / oreThick) : 0;

    oreResults.innerHTML = `
        <div class="results-grid">
            <div class="metric-card thick">
                <div class="metric-info">
                    <h3>Total Thickness</h3>
                    <div class="metric-value">${totalThick.toFixed(2)}<span class="metric-unit">m</span></div>
                </div>
            </div>
            <div class="metric-card ore">
                <div class="metric-info">
                    <h3>Total Ore (Ni ≥ 1.0)</h3>
                    <div class="metric-value">${oreThick.toFixed(2)}<span class="metric-unit">m</span></div>
                    <p style="font-size: 0.8rem; margin-top: 0.5rem; color: #2e7d32; font-weight: 600;">
                        Avg Ni: ${oreAvgNi.toFixed(2)}%
                    </p>
                </div>
            </div>
            <div class="metric-card">
                <div class="metric-info">
                    <h3>Non-Ore (Ni < 1.0)</h3>
                    <div class="metric-value">${nonOreThick.toFixed(2)}<span class="metric-unit">m</span></div>
                </div>
            </div>
            <div class="metric-card ob">
                <div class="metric-info">
                    <h3>Overburden (OB)</h3>
                    <div class="metric-value">${obThick.toFixed(2)}<span class="metric-unit">m</span></div>
                </div>
            </div>
            <div class="metric-card sr">
                <div class="metric-info">
                    <h3>Stripping Ratio (SR)</h3>
                    <div class="metric-value">${sr.toFixed(2)}<span class="metric-unit">OB:Ore</span></div>
                </div>
            </div>
        </div>
        <p style="font-size: 0.75rem; color: #777; margin-top: 1rem;">* OB is defined as Ni < 1.0 and Litho (Zonasi) is not "BRK".</p>
    `;

    oreModal.classList.remove('hidden');
}

function showDiagram() {
    if (activeFilteredData.length === 0) return;

    const selectedHoleId = holeSelect.value;
    document.getElementById('diagramTitle').textContent = `Assay Diagram: ${selectedHoleId}`;
    
    // Set initial size from sliders
    chartWrapper.style.width = `${chartWidthSlider.value}px`;
    chartWrapper.style.height = `${chartHeightSlider.value}px`;
    widthValue.textContent = `${chartWidthSlider.value}px`;
    heightValue.textContent = `${chartHeightSlider.value}px`;

    diagramModal.classList.remove('hidden');

    const ctx = document.getElementById('assayChart').getContext('2d');
    
    if (assayChart) {
        assayChart.destroy();
    }

    // Prepare data as {x, y} where y is Depth (using "To" depth as requested)
    const niData = activeFilteredData.map(d => ({ x: parseFloat(d['Ni']) || 0, y: parseFloat(d['To']) || 0 }));
    const feData = activeFilteredData.map(d => ({ x: parseFloat(d['Fe']) || 0, y: parseFloat(d['To']) || 0 }));
    const mgoData = activeFilteredData.map(d => ({ x: parseFloat(d['MgO']) || 0, y: parseFloat(d['To']) || 0 }));
    const sio2Data = activeFilteredData.map(d => ({ x: parseFloat(d['SiO2']) || 0, y: parseFloat(d['To']) || 0 }));
    const coData = activeFilteredData.map(d => ({ x: parseFloat(d['Co']) || 0, y: parseFloat(d['To']) || 0 }));
    const caoData = activeFilteredData.map(d => ({ x: parseFloat(d['CaO']) || 0, y: parseFloat(d['To']) || 0 }));
    const al2o3Data = activeFilteredData.map(d => ({ x: parseFloat(d['Al2O3']) || 0, y: parseFloat(d['To']) || 0 }));
    const simaData = activeFilteredData.map(d => ({ x: parseFloat(d['SiMa']) || 0, y: parseFloat(d['To']) || 0 }));

    // Find max depth for scale padding
    const maxDepth = Math.max(...activeFilteredData.map(d => parseFloat(d['To']) || 0));

    assayChart = new Chart(ctx, {
        type: 'line',
        data: {
            datasets: [
                {
                    label: 'Ni (%)',
                    data: niData,
                    borderColor: '#1e88e5',
                    backgroundColor: '#1e88e5',
                    xAxisID: 'xNi',
                    yAxisID: 'y',
                    pointStyle: 'rectRot',
                    pointRadius: 4,
                    borderWidth: 2,
                    tension: 0.1,
                    hidden: !document.querySelector('.dataset-toggle[data-index="0"]').checked
                },
                {
                    label: 'Fe (%)',
                    data: feData,
                    borderColor: '#b71c1c',
                    backgroundColor: '#b71c1c',
                    xAxisID: 'xOthers',
                    yAxisID: 'y',
                    pointStyle: 'rect',
                    pointRadius: 4,
                    borderWidth: 2,
                    tension: 0.1,
                    hidden: !document.querySelector('.dataset-toggle[data-index="1"]').checked
                },
                {
                    label: 'MgO (%)',
                    data: mgoData,
                    borderColor: '#ffca28',
                    backgroundColor: '#ffca28',
                    xAxisID: 'xOthers',
                    yAxisID: 'y',
                    pointStyle: 'circle',
                    pointRadius: 4,
                    borderWidth: 2,
                    tension: 0.1,
                    hidden: !document.querySelector('.dataset-toggle[data-index="2"]').checked
                },
                {
                    label: 'SiO2 (%)',
                    data: sio2Data,
                    borderColor: '#7cb342',
                    backgroundColor: '#7cb342',
                    xAxisID: 'xOthers',
                    yAxisID: 'y',
                    pointStyle: 'triangle',
                    pointRadius: 4,
                    borderWidth: 2,
                    tension: 0.1,
                    hidden: !document.querySelector('.dataset-toggle[data-index="3"]').checked
                },
                {
                    label: 'Co (%)',
                    data: coData,
                    borderColor: '#f06292',
                    backgroundColor: '#f06292',
                    xAxisID: 'xOthers',
                    yAxisID: 'y',
                    pointStyle: 'star',
                    pointRadius: 5,
                    borderWidth: 1,
                    tension: 0.1,
                    hidden: !document.querySelector('.dataset-toggle[data-index="4"]').checked
                },
                {
                    label: 'CaO (%)',
                    data: caoData,
                    borderColor: '#9c27b0',
                    backgroundColor: '#9c27b0',
                    xAxisID: 'xOthers',
                    yAxisID: 'y',
                    pointStyle: 'crossRot',
                    pointRadius: 4,
                    borderWidth: 1,
                    tension: 0.1,
                    hidden: !document.querySelector('.dataset-toggle[data-index="5"]').checked
                },
                {
                    label: 'Al2O3 (%)',
                    data: al2o3Data,
                    borderColor: '#8d6e63',
                    backgroundColor: '#8d6e63',
                    xAxisID: 'xOthers',
                    yAxisID: 'y',
                    pointStyle: 'rectRounded',
                    pointRadius: 4,
                    borderWidth: 1,
                    tension: 0.1,
                    hidden: !document.querySelector('.dataset-toggle[data-index="6"]').checked
                },
                {
                    label: 'SiMa',
                    data: simaData,
                    borderColor: '#4db6ac',
                    backgroundColor: '#4db6ac',
                    xAxisID: 'xOthers',
                    yAxisID: 'y',
                    pointStyle: 'dash',
                    pointRadius: 4,
                    borderWidth: 1,
                    tension: 0.1,
                    hidden: !document.querySelector('.dataset-toggle[data-index="7"]').checked
                }
            ]
        },
        options: {
            responsive: true,
            maintainAspectRatio: false,
            indexAxis: 'y', 
            scales: {
                y: {
                    type: 'linear',
                    position: 'left',
                    title: { display: true, text: 'Depth (meter)' },
                    reverse: true,
                    min: 0,
                    max: Math.ceil(maxDepth),
                    ticks: {
                        stepSize: 1,
                        callback: (value) => value % 1 === 0 ? value : ''
                    },
                    grid: { color: '#f0f0f0' }
                },
                yRight: {
                    type: 'linear',
                    position: 'right',
                    reverse: true,
                    min: 0,
                    max: Math.ceil(maxDepth),
                    ticks: {
                        stepSize: 1,
                        callback: (value) => value % 1 === 0 ? value : ''
                    },
                    grid: { display: false }
                },
                xNi: {
                    type: 'linear',
                    position: 'bottom',
                    min: 0,
                    max: parseFloat(niMaxInput.value) || 3.5,
                    title: { display: true, text: 'Ni %', color: '#1e88e5', font: { weight: 'bold' } },
                    ticks: { color: '#1e88e5' },
                    grid: { display: false }
                },
                xOthers: {
                    type: 'linear',
                    position: 'top',
                    min: 0,
                    max: parseFloat(othersMaxInput.value) || 70,
                    title: { display: true, text: 'Fe, MgO, SiO2 %', color: '#b71c1c', font: { weight: 'bold' } },
                    ticks: { color: '#b71c1c' },
                    grid: { color: '#f0f0f0' }
                }
            },
            plugins: {
                tooltip: {
                    mode: 'index',
                    intersect: false,
                    callbacks: {
                        title: (tooltipItems) => `Depth: ${tooltipItems[0].parsed.y.toFixed(2)}m`
                    }
                },
                legend: {
                    position: 'bottom',
                    labels: { usePointStyle: true, padding: 20 }
                }
            }
        }
    });
}

function renderRows(data, collarZ, selectedHoleId) {
    assayBody.innerHTML = '';

    // Chemical columns to apply data bars to
    const chemCols = ['Ni', 'Co', 'Fe', 'SiO2', 'CaO', 'MgO', 'Al2O3', 'SiMa'];

    // Find max values for each column to scale data bars
    const maxVals = {};
    chemCols.forEach(col => {
        maxVals[col] = Math.max(...data.map(row => parseFloat(row[col]) || 0), 0.1); // Avoid div by zero
    });

    data.forEach((row, index) => {
        const fromVal = parseFloat(row['From']) || 0;
        const elev = (collarZ - fromVal).toFixed(2);
        const ni = parseFloat(row['Ni']) || 0;

        let niClass = '';
        if (ni < 1) niClass = 'ni-low';
        else if (ni <= 1.3) niClass = 'ni-med';
        else if (ni > 1.3) niClass = 'ni-high';

        const tr = document.createElement('tr');
        if (niClass) tr.classList.add(niClass);

        // Apply Topo-aware styling
        const topoKey = row['Topo Position'] ? 'Topo Position' : 'Topo';
        const topoVal = (row[topoKey] || '').toLowerCase();
        if (topoVal.includes('above')) {
            tr.classList.add('row-above');
        } else if (topoVal.includes('below')) {
            tr.classList.add('row-below');
        }

        // Helper to create a cell with data bar and 2-decimal formatting
        const createChemCell = (val, colName) => {
            const num = parseFloat(val);
            if (isNaN(num)) return `<td></td>`;

            const percentage = (num / maxVals[colName]) * 100;
            const formatted = num.toFixed(2);

            return `
                <td class="data-bar-cell">
                    <div class="data-bar-bg bar-${colName.toLowerCase()}" style="width: ${percentage}%"></div>
                    <span class="data-bar-text">${formatted}</span>
                </td>
            `;
        };

        const formatNum = (val) => {
            const num = parseFloat(val);
            return isNaN(num) ? (val ?? '') : num.toFixed(2);
        };

        tr.innerHTML = `
            <td>${index + 1}</td>
            <td style="font-weight: 600;">${selectedHoleId}</td>
            <td style="font-style: italic; color: #555;">${row['Zonasi'] || ''}</td>
            <td>${formatNum(row['From'])}</td>
            <td>${formatNum(row['To'])}</td>
            <td style="font-weight: 700; color: var(--primary-dark);">${elev}</td>
            <td>${row[topoKey] || ''}</td>
            ${createChemCell(row['Ni'], 'Ni')}
            ${createChemCell(row['Co'], 'Co')}
            ${createChemCell(row['Fe'], 'Fe')}
            ${createChemCell(row['SiO2'], 'SiO2')}
            ${createChemCell(row['CaO'], 'CaO')}
            ${createChemCell(row['MgO'], 'MgO')}
            ${createChemCell(row['Al2O3'], 'Al2O3')}
            ${createChemCell(row['SiMa'], 'SiMa')}
        `;

        assayBody.appendChild(tr);
    });
}

// Service Worker Registration
if ('serviceWorker' in navigator) {
    window.addEventListener('load', () => {
        navigator.serviceWorker.register('service-worker.js')
            .then(reg => console.log('SW registered', reg))
            .catch(err => console.error('SW error', err));
    });
}
