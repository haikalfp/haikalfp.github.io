let collarData = [];
let assayData = [];

// Map variables
let map = null;
let markersLayer = null;
let imageOverlay = null;
let nwMarker = null;
let seMarker = null;

const fileInput = document.getElementById('fileInput');
const mapInput = document.getElementById('mapInput');
const mapContainer = document.getElementById('map');
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
mapInput.addEventListener('change', handleMapFile);
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

    plotCollarPoints();
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

// Georeference modal elements
const georefModal = document.getElementById('georefModal');
const georefPreview = document.getElementById('georefPreview');
const georefImageWrapper = document.getElementById('georefImageWrapper');
const georefImageTransform = document.getElementById('georefImageTransform');
const gcpCrosshair = document.getElementById('gcpCrosshair');
const gcpStatus = document.getElementById('gcpStatus');
const applyGeorefBtn = document.getElementById('applyGeoref');
const resetGCPsBtn = document.getElementById('resetGCPs');
const coordDisplay = document.getElementById('coordDisplay');
const zoomLevelDisplay = document.getElementById('zoomLevel');

// GCP DOM references for all 4 points
const gcpElements = {};
for (let i = 1; i <= 4; i++) {
    gcpElements[i] = {
        marker: document.getElementById(`gcpMarker${i}`),
        pixel: document.getElementById(`gcp${i}Pixel`),
        easting: document.getElementById(`gcp${i}E`),
        northing: document.getElementById(`gcp${i}N`),
        panel: document.getElementById(`gcp${i}Panel`)
    };
}

let pendingImageUrl = null;
let mapGridLayer = null;

// GCP state
const MAX_GCP = 4;
let gcpPoints = {}; // { 1: {px, py, pctX, pctY}, 2: {...}, ... }
let currentGCP = 1;
let imgNaturalWidth = 0;
let imgNaturalHeight = 0;

// Zoom/Pan state
let zoomScale = 1;
let panX = 0, panY = 0;
let isPanning = false;
let panStartX = 0, panStartY = 0;
let panStartPanX = 0, panStartPanY = 0;
let didDrag = false; // to distinguish click from drag

// Close georef modal
georefModal.querySelector('.close-modal').addEventListener('click', () => {
    georefModal.classList.add('hidden');
});
window.addEventListener('click', (e) => {
    if (e.target === georefModal) georefModal.classList.add('hidden');
});

// ---- Zoom/Pan controls ----
function updateTransform() {
    georefImageTransform.style.transform = `translate(${panX}px, ${panY}px) scale(${zoomScale})`;
    zoomLevelDisplay.textContent = Math.round(zoomScale * 100) + '%';
}

function zoomTo(newScale, centerX, centerY) {
    // Zoom relative to a center point
    const wrapperRect = georefImageWrapper.getBoundingClientRect();
    const cx = centerX !== undefined ? centerX : wrapperRect.width / 2;
    const cy = centerY !== undefined ? centerY : wrapperRect.height / 2;
    
    const ratio = newScale / zoomScale;
    panX = cx - ratio * (cx - panX);
    panY = cy - ratio * (cy - panY);
    zoomScale = newScale;
    updateTransform();
}

document.getElementById('zoomInBtn').addEventListener('click', (e) => {
    e.stopPropagation();
    zoomTo(Math.min(zoomScale * 1.3, 20));
});
document.getElementById('zoomOutBtn').addEventListener('click', (e) => {
    e.stopPropagation();
    zoomTo(Math.max(zoomScale / 1.3, 0.1));
});
document.getElementById('zoomFitBtn').addEventListener('click', (e) => {
    e.stopPropagation();
    zoomScale = 1;
    panX = 0;
    panY = 0;
    updateTransform();
});

// Mouse wheel zoom
georefImageWrapper.addEventListener('wheel', function(e) {
    e.preventDefault();
    const rect = georefImageWrapper.getBoundingClientRect();
    const mouseX = e.clientX - rect.left;
    const mouseY = e.clientY - rect.top;
    
    const factor = e.deltaY < 0 ? 1.15 : 1 / 1.15;
    const newScale = Math.min(Math.max(zoomScale * factor, 0.1), 20);
    zoomTo(newScale, mouseX, mouseY);
}, { passive: false });

// Pan with mouse drag
georefImageWrapper.addEventListener('mousedown', function(e) {
    if (e.button !== 0) return; // only left button
    isPanning = true;
    didDrag = false;
    panStartX = e.clientX;
    panStartY = e.clientY;
    panStartPanX = panX;
    panStartPanY = panY;
    georefImageWrapper.style.cursor = 'grabbing';
    e.preventDefault();
});

window.addEventListener('mousemove', function(e) {
    if (!isPanning) return;
    const dx = e.clientX - panStartX;
    const dy = e.clientY - panStartY;
    if (Math.abs(dx) > 3 || Math.abs(dy) > 3) didDrag = true;
    panX = panStartPanX + dx;
    panY = panStartPanY + dy;
    updateTransform();
});

window.addEventListener('mouseup', function(e) {
    if (!isPanning) return;
    isPanning = false;
    georefImageWrapper.style.cursor = '';
});

// Click handler for placing GCPs (only if no drag occurred)
georefImageWrapper.addEventListener('click', function(e) {
    if (didDrag) return; // was a pan drag, not a click
    if (currentGCP === 0) return; // all points placed
    
    // Get click position relative to the wrapper
    const wrapperRect = georefImageWrapper.getBoundingClientRect();
    const clickInWrapperX = e.clientX - wrapperRect.left;
    const clickInWrapperY = e.clientY - wrapperRect.top;
    
    // Convert to image pixel coordinates (accounting for zoom/pan)
    const imgRect = georefPreview.getBoundingClientRect();
    const clickOnImgX = e.clientX - imgRect.left;
    const clickOnImgY = e.clientY - imgRect.top;
    
    // Check if click is within the image
    if (clickOnImgX < 0 || clickOnImgY < 0 || clickOnImgX > imgRect.width || clickOnImgY > imgRect.height) return;
    
    // Calculate pixel position on the original image
    const scaleDisplayX = imgNaturalWidth / imgRect.width;
    const scaleDisplayY = imgNaturalHeight / imgRect.height;
    const pixelX = Math.round(clickOnImgX * scaleDisplayX);
    const pixelY = Math.round(clickOnImgY * scaleDisplayY);
    
    // Percentage position on the image (for marker placement inside transform container)
    const pctX = (clickOnImgX / imgRect.width) * 100;
    const pctY = (clickOnImgY / imgRect.height) * 100;
    
    placeGCP(currentGCP, pixelX, pixelY, pctX, pctY);
});

function placeGCP(pointNum, pixelX, pixelY, pctX, pctY) {
    gcpPoints[pointNum] = { px: pixelX, py: pixelY, pctX, pctY };
    
    const el = gcpElements[pointNum];
    
    // Show marker
    el.marker.classList.remove('hidden');
    el.marker.style.left = pctX + '%';
    el.marker.style.top = pctY + '%';
    
    // Update pixel info
    el.pixel.textContent = `px(${pixelX}, ${pixelY})`;
    
    // Enable coordinate inputs
    el.easting.disabled = false;
    el.northing.disabled = false;
    el.easting.focus();
    
    // Highlight panel
    el.panel.classList.add('active');
    
    // Advance to next point
    const nextPoint = pointNum + 1;
    if (nextPoint <= MAX_GCP) {
        currentGCP = nextPoint;
        gcpStatus.innerHTML = `Click on the image to place <strong>Point ${nextPoint}</strong> <span class="gcp-optional">${nextPoint > 2 ? '(optional)' : ''}</span>`;
    } else {
        currentGCP = 0;
        gcpStatus.innerHTML = 'All 4 points placed. Enter coordinates and click <strong>Apply</strong>.';
    }
    
    // Enable apply if >= 2 points placed
    const placedCount = Object.keys(gcpPoints).length;
    if (placedCount >= 2) {
        applyGeorefBtn.disabled = false;
        if (currentGCP > 0) {
            gcpStatus.innerHTML += ' — or click <strong>Apply</strong> now with current points.';
        }
    }
}

// Show crosshair on wrapper hover
georefImageWrapper.addEventListener('mousemove', function(e) {
    if (currentGCP === 0 || isPanning) {
        gcpCrosshair.classList.add('hidden');
        return;
    }
    const rect = georefImageWrapper.getBoundingClientRect();
    const x = e.clientX - rect.left;
    const y = e.clientY - rect.top;
    
    gcpCrosshair.classList.remove('hidden');
    gcpCrosshair.style.left = x + 'px';
    gcpCrosshair.style.top = y + 'px';
});

georefImageWrapper.addEventListener('mouseleave', function() {
    gcpCrosshair.classList.add('hidden');
});

// Reset GCPs
resetGCPsBtn.addEventListener('click', resetGCPState);

function resetGCPState() {
    gcpPoints = {};
    currentGCP = 1;
    
    for (let i = 1; i <= MAX_GCP; i++) {
        const el = gcpElements[i];
        el.marker.classList.add('hidden');
        el.pixel.textContent = '—';
        el.easting.value = '';
        el.northing.value = '';
        el.easting.disabled = true;
        el.northing.disabled = true;
        el.panel.classList.remove('active');
    }
    
    applyGeorefBtn.disabled = true;
    gcpStatus.innerHTML = 'Click on the image to place <strong>Point 1</strong>';
    
    // Reset zoom/pan
    zoomScale = 1;
    panX = 0;
    panY = 0;
    updateTransform();
}

applyGeorefBtn.addEventListener('click', applyGeoreference);

// Map Functions
function initMap() {
    if (map) return;
    mapContainer.classList.remove('hidden');
    
    map = L.map('map', {
        crs: L.CRS.Simple,
        zoomSnap: 0.25,
        zoomDelta: 0.5,
        minZoom: -10,
        maxZoom: 10
    }).setView([0, 0], 0);

    markersLayer = L.featureGroup().addTo(map);

    map.on('mousemove', function(e) {
        const easting = e.latlng.lng.toFixed(1);
        const northing = e.latlng.lat.toFixed(1);
        coordDisplay.textContent = `E: ${easting}  |  N: ${northing}  (UTM 51S)`;
    });

    map.on('mouseout', function() {
        coordDisplay.textContent = '';
    });
}

function handleMapFile(e) {
    const file = e.target.files[0];
    if (!file) return;

    if (!map) initMap();

    const name = file.name.toLowerCase();
    const isTiff = name.endsWith('.tif') || name.endsWith('.tiff');
    const isPdf  = name.endsWith('.pdf');

    if (isTiff) {
        handleGeoTiff(file);
    } else if (isPdf) {
        handlePdf(file);
    } else {
        handleRasterImage(file);
    }
}

// ---- GeoTIFF: auto-georeference from embedded bounds ----
function handleGeoTiff(file) {
    document.getElementById('mapUploadBtn').textContent = '⏳ Parsing GeoTIFF…';

    const reader = new FileReader();
    reader.onload = function(event) {
        const arrayBuffer = event.target.result;

        parseGeoraster(arrayBuffer).then(georaster => {
            // georaster gives us xmin/xmax/ymin/ymax in the raster's native SRS.
            // For a UTM-projected tiff these are already in metres (Easting/Northing).
            const minEasting  = georaster.xmin;
            const maxEasting  = georaster.xmax;
            const minNorthing = georaster.ymin;
            const maxNorthing = georaster.ymax;

            // Sanity check: UTM eastings are ~100k–900k, northings are 0–10M
            const looksLikeUTM = Math.abs(minEasting) > 1000 && Math.abs(minNorthing) > 1000;

            if (!looksLikeUTM) {
                alert(
                    'This GeoTIFF does not appear to be in a projected (UTM) coordinate system.\n' +
                    `Bounds: E ${minEasting.toFixed(1)}–${maxEasting.toFixed(1)}, ` +
                    `N ${minNorthing.toFixed(1)}–${maxNorthing.toFixed(1)}\n\n` +
                    'Please use a GeoTIFF exported in UTM Zone 51S, or use a PNG/JPG with manual GCP georeferencing.'
                );
                document.getElementById('mapUploadBtn').textContent = '🗺️ Upload Map Image';
                return;
            }

            // Render raster to a canvas → get data URL for the overlay
            renderGeoRasterToDataUrl(georaster).then(dataUrl => {
                placeImageOnMap(dataUrl, minEasting, maxEasting, minNorthing, maxNorthing, 'auto');
                document.getElementById('mapUploadBtn').textContent = '✅ GeoTIFF Auto-Georeferenced';
            }).catch(err => {
                console.error('GeoTIFF render error:', err);
                alert('GeoTIFF loaded but could not render to image: ' + err.message);
                document.getElementById('mapUploadBtn').textContent = '🗺️ Upload Map Image';
            });

        }).catch(err => {
            console.error('GeoTIFF parse error:', err);
            alert('Failed to parse GeoTIFF: ' + err.message);
            document.getElementById('mapUploadBtn').textContent = '🗺️ Upload Map Image';
        });
    };
    reader.readAsArrayBuffer(file);
}

// Render a georaster to a data URL via a hidden canvas
function renderGeoRasterToDataUrl(georaster) {
    return new Promise((resolve, reject) => {
        try {
            const width  = georaster.width;
            const height = georaster.height;
            const canvas  = document.createElement('canvas');
            canvas.width  = width;
            canvas.height = height;
            const ctx = canvas.getContext('2d');
            const imageData = ctx.createImageData(width, height);
            const data = imageData.data;

            const bands = georaster.values; // [band0, band1, band2, ...]
            const noData = georaster.noDataValue;

            for (let y = 0; y < height; y++) {
                for (let x = 0; x < width; x++) {
                    const idx = (y * width + x) * 4;
                    let r, g, b, a = 255;

                    if (bands.length >= 3) {
                        // RGB or RGBA
                        r = bands[0][y][x];
                        g = bands[1][y][x];
                        b = bands[2][y][x];
                        if (bands.length >= 4) a = bands[3][y][x];
                    } else {
                        // Grayscale
                        r = g = b = bands[0][y][x];
                    }

                    // Handle nodata
                    if (noData !== null && noData !== undefined && r === noData) a = 0;

                    data[idx]     = r;
                    data[idx + 1] = g;
                    data[idx + 2] = b;
                    data[idx + 3] = a;
                }
            }

            ctx.putImageData(imageData, 0, 0);
            resolve(canvas.toDataURL('image/png'));
        } catch (err) {
            reject(err);
        }
    });
}

// ---- PDF: try auto-georeference from GeoPDF metadata, fallback to GCP modal ----
function handlePdf(file) {
    document.getElementById('mapUploadBtn').textContent = '⏳ Rendering PDF…';

    if (typeof pdfjsLib !== 'undefined') {
        pdfjsLib.GlobalWorkerOptions.workerSrc =
            'https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.worker.min.js';
    }

    const reader = new FileReader();
    reader.onload = function(event) {
        const arrayBuffer = event.target.result;
        const typedArray  = new Uint8Array(arrayBuffer);

        // 1. Try to extract GeoPDF info (geo bounds + LPTS data-frame crop region)
        const geoInfo = tryParseGeoPDF(arrayBuffer);
        if (geoInfo) {
            console.log('[GeoPDF] Bounds:', geoInfo.bounds, '| LPTS bbox:', geoInfo.lptsBbox);
        } else {
            console.log('[GeoPDF] No georef metadata found — using manual GCP modal.');
        }

        // 2. Render page 1 → canvas (return canvas, not dataUrl, so we can crop)
        pdfjsLib.getDocument({ data: typedArray }).promise.then(pdf => {
            return pdf.getPage(1);
        }).then(page => {
            const scale       = 2;
            const viewport    = page.getViewport({ scale });
            const pageHPts    = viewport.height / scale; // page height in PDF pts (y=0 at bottom)
            const canvas      = document.createElement('canvas');
            canvas.width      = viewport.width;
            canvas.height     = viewport.height;
            const ctx         = canvas.getContext('2d');
            return page.render({ canvasContext: ctx, viewport }).promise.then(() =>
                ({ canvas, scale, pageHPts })
            );
        }).then(({ canvas, scale, pageHPts }) => {
            // 3. Crop canvas to the map data-frame region.
            //
            //    Per OGC GeoPDF spec, LPTS [0-1] values are normalised WITHIN the viewport's
            //    /BBox (in PDF pts). If bboxPts is available:
            //      data_left   = bx0 + lx1*(bx1-bx0)
            //      data_top    = by0 + ly2*(by1-by0)   (PDF y↑, so y2 = max)
            //      canvas_y    = (pageHPts - data_top) * scale   (flip because canvas y↓)
            //    Fallback when BBox is unknown: treat LPTS as proportion of full canvas.
            let dataUrl;
            if (geoInfo) {
                const { x1, y1, x2, y2 } = geoInfo.lptsBbox;
                let cx, cy, cw, ch;

                if (geoInfo.bboxPts) {
                    const [bx0, by0, bx1, by1] = geoInfo.bboxPts;
                    const dataLeft   = bx0 + x1 * (bx1 - bx0);
                    const dataRight  = bx0 + x2 * (bx1 - bx0);
                    const dataBottom = by0 + y1 * (by1 - by0);
                    const dataTop    = by0 + y2 * (by1 - by0);
                    cx = Math.round(dataLeft  * scale);
                    cy = Math.round((pageHPts - dataTop) * scale);   // PDF→canvas Y flip
                    cw = Math.max(1, Math.round((dataRight  - dataLeft)  * scale));
                    ch = Math.max(1, Math.round((dataTop    - dataBottom) * scale));
                    console.log(`[GeoPDF] BBox crop: PDF (${dataLeft.toFixed(1)},${dataBottom.toFixed(1)})–(${dataRight.toFixed(1)},${dataTop.toFixed(1)}) → canvas (${cx},${cy}) ${cw}×${ch}`);
                } else {
                    // Fallback: assume LPTS normalised to full page
                    cx = Math.round(x1 * canvas.width);
                    cy = Math.round((1 - y2) * canvas.height);
                    cw = Math.max(1, Math.round((x2 - x1) * canvas.width));
                    ch = Math.max(1, Math.round((y2 - y1) * canvas.height));
                    console.log(`[GeoPDF] Full-page crop: (${cx},${cy}) ${cw}×${ch} of ${canvas.width}×${canvas.height}`);
                }

                const needsCrop = (cx > 0 || cy > 0 || cw < canvas.width || ch < canvas.height);
                if (needsCrop) {
                    const crop = document.createElement('canvas');
                    crop.width  = cw;
                    crop.height = ch;
                    crop.getContext('2d').drawImage(canvas, cx, cy, cw, ch, 0, 0, cw, ch);
                    dataUrl = crop.toDataURL('image/png');
                } else {
                    dataUrl = canvas.toDataURL('image/png');
                }

                const b = geoInfo.bounds;
                console.log(`[GeoPDF] Placing: E ${b.minEasting.toFixed(0)}–${b.maxEasting.toFixed(0)}, N ${b.minNorthing.toFixed(0)}–${b.maxNorthing.toFixed(0)}`);
                placeImageOnMap(dataUrl, b.minEasting, b.maxEasting, b.minNorthing, b.maxNorthing, 'auto');
                document.getElementById('mapUploadBtn').textContent = '✅ GeoPDF Auto-Georeferenced';
            } else {
                // No georef — open GCP modal with full page image
                dataUrl = canvas.toDataURL('image/png');
                pendingImageUrl = dataUrl;
                const tempImg   = new Image();
                tempImg.onload  = function() {
                    imgNaturalWidth  = tempImg.naturalWidth;
                    imgNaturalHeight = tempImg.naturalHeight;
                    georefPreview.src = dataUrl;
                    resetGCPState();
                    georefModal.classList.remove('hidden');
                    document.getElementById('mapUploadBtn').textContent = '🗺️ Upload Map (PNG/JPG/TIF/PDF)';
                };
                tempImg.src = dataUrl;
            }
        }).catch(err => {
            console.error('PDF render error:', err);
            alert('Failed to render PDF: ' + err.message);
            document.getElementById('mapUploadBtn').textContent = '🗺️ Upload Map (PNG/JPG/TIF/PDF)';
        });
    };
    reader.readAsArrayBuffer(file);
}

// ---- GeoPDF parsing --------------------------------------------------------
// Supports (all formats, both raw text and compressed ObjStm streams):
//   1. OGC ISO 32000  — /GPTS [lat lon ...] + /LPTS [lx ly ...] (QGIS, GDAL, ArcGIS)
//   2. OGC LGIDict 2.x — /Registration [[px py lat lon] ...] (NGA, older tools)
//   3. OGC LGIDict CTM — /CTM [a b c d e f] + /Neatline (ArcGIS Pro variant)
//
// searchGeoPDFInfo  → {bounds, lptsBbox} or null
//   bounds   : {minEasting, maxEasting, minNorthing, maxNorthing} in UTM 51S
//   lptsBbox : {x1, y1, x2, y2} in normalised LPTS page space [0,1]
//              Defines which portion of the rendered page contains the map data frame.
//              Default (0,0,1,1) = full page (no cropping needed).
//
// tryParseGeoPDF    → same, but also searches pako-decompressed FlateDecode streams.
// ---------------------------------------------------------------------------

function bytesToLatin1(arr, len) {
    const parts = [];
    const CHUNK = 32768;
    for (let i = 0; i < len; i += CHUNK) {
        parts.push(String.fromCharCode.apply(null, arr.subarray(i, Math.min(i + CHUNK, len))));
    }
    return parts.join('');
}

function searchGeoPDFInfo(text) {
    // --- Method 1: /GPTS [lat0 lon0 ...] + optional /LPTS [lx0 ly0 ...] + /BBox ---
    //
    // OGC spec: LPTS is normalised 0-1 WITHIN the viewport's /BBox (in PDF pts).
    // So we must also parse /BBox to convert LPTS into absolute PDF pt coordinates,
    // which we then convert to canvas pixels for the crop step.
    const gptsRe   = /\/GPTS\s*\[([^\]]+)\]/;
    const gptsExec = gptsRe.exec(text);
    const lptsMatch = text.match(/\/LPTS\s*\[([^\]]+)\]/);

    if (gptsExec) {
        const gptsVals = gptsExec[1].trim().split(/\s+/).map(Number).filter(v => !isNaN(v));
        if (gptsVals.length >= 8) {
            const lats = [], lons = [];
            for (let i = 0; i < gptsVals.length; i += 2) {
                lats.push(gptsVals[i]);
                lons.push(gptsVals[i + 1]);
            }
            const minLat = Math.min(...lats), maxLat = Math.max(...lats);
            const minLon = Math.min(...lons), maxLon = Math.max(...lons);

            // Parse LPTS bounding box (normalised 0-1 within BBox)
            let lptsBbox = { x1: 0, y1: 0, x2: 1, y2: 1 }; // default: full BBox
            if (lptsMatch) {
                const lptsVals = lptsMatch[1].trim().split(/\s+/).map(Number).filter(v => !isNaN(v));
                if (lptsVals.length >= 8) {
                    const lxs = [], lys = [];
                    for (let i = 0; i < lptsVals.length; i += 2) {
                        lxs.push(lptsVals[i]);
                        lys.push(lptsVals[i + 1]);
                    }
                    lptsBbox = {
                        x1: Math.min(...lxs), y1: Math.min(...lys),
                        x2: Math.max(...lxs), y2: Math.max(...lys)
                    };
                }
            }

            // Parse /BBox [x0 y0 x1 y1] of the viewport that owns this GPTS.
            // Strategy: look up to 3000 chars BEFORE the GPTS occurrence; take
            // the LAST /BBox found (closest one = the enclosing viewport's BBox).
            let bboxPts = null;
            const preGPTS      = text.substring(Math.max(0, gptsExec.index - 3000), gptsExec.index);
            const allBBoxMatches = [...preGPTS.matchAll(/\/BBox\s*\[([^\]]+)\]/g)];
            const closestBBox  = allBBoxMatches.length ? allBBoxMatches[allBBoxMatches.length - 1] : null;
            if (closestBBox) {
                const bv = closestBBox[1].trim().split(/\s+/).map(Number);
                if (bv.length >= 4 && !bv.some(isNaN)) bboxPts = bv.slice(0, 4);
            }

            console.log('[GeoPDF] /GPTS:', minLat.toFixed(5), maxLat.toFixed(5), minLon.toFixed(5), maxLon.toFixed(5));
            console.log('[GeoPDF] /LPTS bbox:', JSON.stringify(lptsBbox));
            console.log('[GeoPDF] /BBox (PDF pts):', bboxPts ? bboxPts.join(' ') : 'not found');

            let bounds;
            if (Math.abs(minLat) <= 90 && Math.abs(maxLat) <= 90 &&
                Math.abs(minLon) <= 180 && Math.abs(maxLon) <= 180) {
                bounds = latLonBoundsToUTM51S(minLat, maxLat, minLon, maxLon);
            } else {
                bounds = { minEasting: Math.min(...lons), maxEasting: Math.max(...lons),
                           minNorthing: Math.min(...lats), maxNorthing: Math.max(...lats) };
            }
            if (bounds) return { bounds, lptsBbox, bboxPts };
        }
    }

    // --- Method 2: /Registration [[px py lat lon] ...] ---
    const regMatch = text.match(/\/Registration\s*\[(\s*\[[^\]]*\]\s*)+\]/);
    if (regMatch) {
        const blocks = [...regMatch[0].matchAll(/\[\s*([\d.\-eE+\s]+)\s*\]/g)];
        const lats = [], lons = [];
        for (const b of blocks) {
            const v = b[1].trim().split(/\s+/).map(Number);
            if (v.length >= 4) { lats.push(v[2]); lons.push(v[3]); }
        }
        if (lats.length >= 2) {
            const minLat = Math.min(...lats), maxLat = Math.max(...lats);
            const minLon = Math.min(...lons), maxLon = Math.max(...lons);
            console.log('[GeoPDF] /Registration:', minLat, maxLat, minLon, maxLon);
            const bounds = latLonBoundsToUTM51S(minLat, maxLat, minLon, maxLon);
            if (bounds) return { bounds, lptsBbox: { x1: 0, y1: 0, x2: 1, y2: 1 } };
        }
    }

    // --- Method 3: /CTM [a b c d e f] + /Neatline [x y ...] ---
    const ctmMatch = text.match(/\/CTM\s*\[([^\]]+)\]/);
    const nlMatch  = text.match(/\/Neatline\s*\[([^\]]+)\]/);
    if (ctmMatch && nlMatch) {
        const ctm = ctmMatch[1].trim().split(/\s+/).map(Number);
        const nl  = nlMatch[1].trim().split(/\s+/).map(Number);
        if (ctm.length >= 6 && nl.length >= 8) {
            const [a, b, c, d, e, f] = ctm;
            const geoXs = [], geoYs = [];
            for (let i = 0; i < nl.length; i += 2) {
                geoXs.push(a * nl[i] + c * nl[i+1] + e);
                geoYs.push(b * nl[i] + d * nl[i+1] + f);
            }
            const minX = Math.min(...geoXs), maxX = Math.max(...geoXs);
            const minY = Math.min(...geoYs), maxY = Math.max(...geoYs);
            console.log('[GeoPDF] /CTM+/Neatline geo:', minX, maxX, minY, maxY);
            let bounds;
            if (Math.abs(minX) <= 180 && Math.abs(maxX) <= 180 &&
                Math.abs(minY) <= 90  && Math.abs(maxY) <= 90) {
                bounds = latLonBoundsToUTM51S(minY, maxY, minX, maxX);
            } else {
                bounds = { minEasting: minX, maxEasting: maxX, minNorthing: minY, maxNorthing: maxY };
            }
            if (bounds) return { bounds, lptsBbox: { x1: 0, y1: 0, x2: 1, y2: 1 } };
        }
    }

    return null;
}

// Searches both raw PDF text and pako-decompressed FlateDecode streams.
// Returns {bounds, lptsBbox} or null (same shape as searchGeoPDFInfo).
function tryParseGeoPDF(arrayBuffer) {
    try {
        const bytes    = new Uint8Array(arrayBuffer);
        const maxBytes = Math.min(bytes.length, 10 * 1024 * 1024);

        // Pass 1: raw text (fast — works unless page dict is in ObjStm)
        const rawText = bytesToLatin1(bytes, maxBytes);
        const direct  = searchGeoPDFInfo(rawText);
        if (direct) return direct;

        // Pass 2: decompress every FlateDecode stream (handles ArcGIS Pro ObjStm)
        if (typeof pako === 'undefined') {
            console.warn('[GeoPDF] pako not loaded — cannot search compressed streams');
            return null;
        }

        let pos = 0, streamsSearched = 0;
        while (pos < rawText.length && streamsSearched < 200) {
            const streamKeyIdx = rawText.indexOf('\nstream', pos);
            if (streamKeyIdx === -1) break;

            const lookback   = Math.max(0, streamKeyIdx - 600);
            const dictRegion = rawText.substring(lookback, streamKeyIdx);
            const hasFD      = /\/FlateDecode|\/Fl[\s>\/]/.test(dictRegion);

            pos = streamKeyIdx + 1;
            if (!hasFD) continue;

            let dataStart = streamKeyIdx + 7;
            if (rawText[dataStart] === '\r') dataStart++;

            const endIdx = rawText.indexOf('endstream', dataStart);
            if (endIdx === -1) break;

            let dataEnd = endIdx;
            while (dataEnd > dataStart &&
                   (rawText[dataEnd - 1] === '\n' || rawText[dataEnd - 1] === '\r')) dataEnd--;

            if (dataEnd <= dataStart || dataEnd - dataStart > 2 * 1024 * 1024) {
                pos = endIdx + 9; continue;
            }

            try {
                const decompText = bytesToLatin1(
                    pako.inflate(bytes.subarray(dataStart, dataEnd)),
                    Infinity
                );
                const result = searchGeoPDFInfo(decompText);
                if (result) return result;
                streamsSearched++;
            } catch (_) { /* not a deflate stream — skip */ }

            pos = endIdx + 9;
        }

        console.log(`[GeoPDF] Checked ${streamsSearched} compressed streams — no georef found.`);
        return null;
    } catch (err) {
        console.warn('[GeoPDF] Parse error:', err);
        return null;
    }
}

// Backward-compat wrapper used by nothing new — kept in case of future calls.
function tryParseGeoPDFBounds(arrayBuffer) {
    const info = tryParseGeoPDF(arrayBuffer);
    return info ? info.bounds : null;
}

// Convert WGS84 lat/lon bounding box → UTM Zone 51S (EPSG:32751) using proj4.
function latLonBoundsToUTM51S(minLat, maxLat, minLon, maxLon) {
    if (typeof proj4 === 'undefined') {
        console.warn('proj4 not loaded — cannot convert lat/lon to UTM');
        return null;
    }
    if (!proj4.defs('EPSG:32751')) {
        proj4.defs('EPSG:32751', '+proj=utm +zone=51 +south +datum=WGS84 +units=m +no_defs');
    }
    const sw = proj4('EPSG:4326', 'EPSG:32751', [minLon, minLat]);
    const ne = proj4('EPSG:4326', 'EPSG:32751', [maxLon, maxLat]);
    return { minEasting: sw[0], maxEasting: ne[0], minNorthing: sw[1], maxNorthing: ne[1] };
}

// ---- PNG / JPG: open GCP modal ----
function handleRasterImage(file) {
    const reader = new FileReader();
    reader.onload = function(event) {
        pendingImageUrl = event.target.result;

        const tempImg = new Image();
        tempImg.onload = function() {
            imgNaturalWidth  = tempImg.naturalWidth;
            imgNaturalHeight = tempImg.naturalHeight;
            georefPreview.src = pendingImageUrl;
            resetGCPState();
            georefModal.classList.remove('hidden');
        };
        tempImg.src = pendingImageUrl;
    };
    reader.readAsDataURL(file);
}

// ---- Shared: place image on the Leaflet map ----
function placeImageOnMap(dataUrl, minEasting, maxEasting, minNorthing, maxNorthing, source) {
    // Remove previous overlay and handles
    if (imageOverlay) map.removeLayer(imageOverlay);
    if (nwMarker)     map.removeLayer(nwMarker);
    if (seMarker)     map.removeLayer(seMarker);

    // L.CRS.Simple: [lat, lng] = [Northing, Easting]
    const sw = L.latLng(minNorthing, minEasting);
    const ne = L.latLng(maxNorthing, maxEasting);
    const bounds = L.latLngBounds(sw, ne);

    imageOverlay = L.imageOverlay(dataUrl, bounds, { opacity: 0.9, interactive: false }).addTo(map);
    imageOverlay.bringToBack();

    // Draggable corner handles for fine-tuning
    if (source !== 'auto') {
        const handleIcon = L.divIcon({
            className: 'georef-handle',
            html: '<div style="background:#ff6f00;width:14px;height:14px;border:2px solid white;border-radius:50%;box-shadow:0 0 6px rgba(0,0,0,0.6);cursor:grab;"></div>',
            iconSize: [18, 18],
            iconAnchor: [9, 9]
        });

        nwMarker = L.marker(bounds.getNorthWest(), { draggable: true, icon: handleIcon, zIndexOffset: 1000 }).addTo(map);
        seMarker = L.marker(bounds.getSouthEast(), { draggable: true, icon: handleIcon, zIndexOffset: 1000 }).addTo(map);
        nwMarker.bindTooltip('⬉ Drag: Top-Left (NW)');
        seMarker.bindTooltip('⬊ Drag: Bottom-Right (SE)');

        function updateOverlayBounds() {
            const nb = L.latLngBounds(
                L.latLng(Math.min(nwMarker.getLatLng().lat, seMarker.getLatLng().lat),
                         Math.min(nwMarker.getLatLng().lng, seMarker.getLatLng().lng)),
                L.latLng(Math.max(nwMarker.getLatLng().lat, seMarker.getLatLng().lat),
                         Math.max(nwMarker.getLatLng().lng, seMarker.getLatLng().lng))
            );
            imageOverlay.setBounds(nb);
        }
        nwMarker.on('drag', updateOverlayBounds);
        seMarker.on('drag', updateOverlayBounds);
    }

    map.fitBounds(bounds, { padding: [30, 30] });
    addCoordinateGrid(minEasting, maxEasting, minNorthing, maxNorthing);
    pendingImageUrl = null;
}


function getCollarCoordinates() {
    const coords = [];
    collarData.forEach(row => {
        const holeId = row['Hole ID'] || row['HoleID'];
        if (!holeId) return;
        
        let xKey = Object.keys(row).find(k => k.toLowerCase() === 'x' || k.toLowerCase() === 'easting' || k.toLowerCase() === 'lon' || k.toLowerCase() === 'longitude');
        let yKey = Object.keys(row).find(k => k.toLowerCase() === 'y' || k.toLowerCase() === 'northing' || k.toLowerCase() === 'lat' || k.toLowerCase() === 'latitude');
        
        if (!xKey && row['East']) xKey = 'East';
        if (!yKey && row['North']) yKey = 'North';
        
        const x = xKey ? parseFloat(row[xKey]) : null;
        const y = yKey ? parseFloat(row[yKey]) : null;
        
        if (x !== null && !isNaN(x) && y !== null && !isNaN(y)) {
            coords.push({ holeId, x, y });
        }
    });
    return coords;
}

function applyGeoreference() {
    // Collect all valid GCPs (placed + coordinates filled)
    const validGCPs = [];
    for (let i = 1; i <= MAX_GCP; i++) {
        if (!gcpPoints[i]) continue;
        const e = parseFloat(gcpElements[i].easting.value);
        const n = parseFloat(gcpElements[i].northing.value);
        if (isNaN(e) || isNaN(n)) continue;
        validGCPs.push({ px: gcpPoints[i].px, py: gcpPoints[i].py, e, n });
    }

    if (validGCPs.length < 2) {
        alert('Please place and fill coordinates for at least 2 points.');
        return;
    }
    if (!pendingImageUrl) {
        alert('No image loaded. Please upload a map image first.');
        return;
    }

    // Least-squares regression:  E = offsetE + scaleX·px
    //                            N = offsetN + scaleY·py  (scaleY < 0)
    const cnt    = validGCPs.length;
    const meanPx = validGCPs.reduce((s, g) => s + g.px, 0) / cnt;
    const meanE  = validGCPs.reduce((s, g) => s + g.e,  0) / cnt;
    const meanPy = validGCPs.reduce((s, g) => s + g.py, 0) / cnt;
    const meanN  = validGCPs.reduce((s, g) => s + g.n,  0) / cnt;

    let numE = 0, denPx = 0, numN = 0, denPy = 0;
    for (const g of validGCPs) {
        numE  += (g.px - meanPx) * (g.e - meanE);
        denPx += (g.px - meanPx) ** 2;
        numN  += (g.py - meanPy) * (g.n - meanN);
        denPy += (g.py - meanPy) ** 2;
    }

    if (denPx === 0) { alert('All points share the same X pixel — use different horizontal positions.'); return; }
    if (denPy === 0) { alert('All points share the same Y pixel — use different vertical positions.'); return; }

    const scaleX  = numE  / denPx;
    const offsetE = meanE - scaleX  * meanPx;
    const scaleY  = numN  / denPy;
    const offsetN = meanN - scaleY  * meanPy;

    const minEasting  = offsetE;
    const maxEasting  = offsetE + imgNaturalWidth  * scaleX;
    const maxNorthing = offsetN;
    const minNorthing = offsetN + imgNaturalHeight * scaleY;

    const dataUrl = pendingImageUrl;
    placeImageOnMap(dataUrl, minEasting, maxEasting, minNorthing, maxNorthing, 'gcp');

    georefModal.classList.add('hidden');
    document.getElementById('mapUploadBtn').textContent = '✅ Map Georeferenced';
}

function addCoordinateGrid(minX, maxX, minY, maxY) {
    // Remove previous grid
    if (mapGridLayer) {
        map.removeLayer(mapGridLayer);
    }
    
    mapGridLayer = L.layerGroup().addTo(map);
    
    // Calculate nice grid interval
    const rangeX = maxX - minX;
    const rangeY = maxY - minY;
    const maxRange = Math.max(rangeX, rangeY);
    
    let gridInterval;
    if (maxRange > 10000) gridInterval = 2000;
    else if (maxRange > 5000) gridInterval = 1000;
    else if (maxRange > 2000) gridInterval = 500;
    else if (maxRange > 1000) gridInterval = 200;
    else if (maxRange > 500) gridInterval = 100;
    else gridInterval = 50;
    
    const padding = gridInterval;
    const startX = Math.floor((minX - padding) / gridInterval) * gridInterval;
    const endX = Math.ceil((maxX + padding) / gridInterval) * gridInterval;
    const startY = Math.floor((minY - padding) / gridInterval) * gridInterval;
    const endY = Math.ceil((maxY + padding) / gridInterval) * gridInterval;
    
    // Vertical lines (Easting)
    for (let x = startX; x <= endX; x += gridInterval) {
        const line = L.polyline([
            [startY, x],
            [endY, x]
        ], {
            color: '#999',
            weight: 0.5,
            opacity: 0.4,
            dashArray: '4 4',
            interactive: false
        });
        mapGridLayer.addLayer(line);
        
        // Label
        const label = L.marker([startY, x], {
            icon: L.divIcon({
                className: 'grid-label',
                html: `<span>${x.toLocaleString()}</span>`,
                iconSize: [80, 16],
                iconAnchor: [40, -4]
            }),
            interactive: false
        });
        mapGridLayer.addLayer(label);
    }
    
    // Horizontal lines (Northing)
    for (let y = startY; y <= endY; y += gridInterval) {
        const line = L.polyline([
            [y, startX],
            [y, endX]
        ], {
            color: '#999',
            weight: 0.5,
            opacity: 0.4,
            dashArray: '4 4',
            interactive: false
        });
        mapGridLayer.addLayer(line);
        
        // Label
        const label = L.marker([y, startX], {
            icon: L.divIcon({
                className: 'grid-label grid-label-y',
                html: `<span>${y.toLocaleString()}</span>`,
                iconSize: [80, 16],
                iconAnchor: [84, 8]
            }),
            interactive: false
        });
        mapGridLayer.addLayer(label);
    }
}

function plotCollarPoints() {
    if (!collarData || collarData.length === 0) return;
    if (!map) initMap();
    
    markersLayer.clearLayers();
    
    const coords = getCollarCoordinates();
    
    if (coords.length === 0) return;
    
    coords.forEach(({ holeId, x, y }) => {
        // In L.CRS.Simple: [lat, lng] = [Northing, Easting] = [Y, X]
        const marker = L.circleMarker([y, x], {
            radius: 7,
            fillColor: '#1a237e',
            color: '#ffffff',
            weight: 2,
            opacity: 1,
            fillOpacity: 0.9
        });
        
        // Rich tooltip with coordinates
        marker.bindTooltip(
            `<strong>${holeId}</strong><br>E: ${x.toFixed(1)}<br>N: ${y.toFixed(1)}`,
            { className: 'collar-tooltip' }
        );
        
        marker.on('click', () => {
            holeSelect.value = holeId;
            updateTable();
            
            // Highlight selected marker
            markersLayer.eachLayer(m => {
                if (m.setStyle) {
                    m.setStyle({ fillColor: '#1a237e', color: '#ffffff', weight: 2, radius: 7 });
                }
            });
            marker.setStyle({ fillColor: '#ff6f00', color: '#ffffff', weight: 3, radius: 9 });
        });
        
        markersLayer.addLayer(marker);
    });
    
    // Fit map to collar points
    map.fitBounds(markersLayer.getBounds(), { padding: [50, 50] });
    
    // Add a basic grid around the collar points
    const eastings = coords.map(c => c.x);
    const northings = coords.map(c => c.y);
    addCoordinateGrid(
        Math.min(...eastings),
        Math.max(...eastings),
        Math.min(...northings),
        Math.max(...northings)
    );
}
