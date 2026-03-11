let collarData = [];
let assayData = [];

const fileInput = document.getElementById('fileInput');
const holeSelect = document.getElementById('holeSelect');
const welcomeMessage = document.getElementById('welcomeMessage');
const tableContainer = document.getElementById('tableContainer');
const assayBody = document.getElementById('assayBody');

fileInput.addEventListener('change', handleFile);
holeSelect.addEventListener('change', updateTable);

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
    const filteredAssay = assayData.filter(row => (row['HoleID'] || row['Hole ID']) === selectedHoleId);

    // Sort by FROM ascending
    filteredAssay.sort((a, b) => (parseFloat(a['From']) || 0) - (parseFloat(b['From']) || 0));

    renderRows(filteredAssay, collarZ, selectedHoleId);
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
        const topoVal = (row['Topo'] || '').toLowerCase();
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
            <td>${row['Topo'] || ''}</td>
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
