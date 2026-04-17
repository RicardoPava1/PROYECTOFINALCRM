// UI Elements
const btnDashboard = document.getElementById('btn-dashboard');
const btnNewClient = document.getElementById('btn-new-client');
const viewDashboard = document.getElementById('view-dashboard');
const viewNewClient = document.getElementById('view-new-client');
const uploadContainer = document.getElementById('upload-container');
const fileInput = document.getElementById('excel-upload');
const formNewClient = document.getElementById('new-client-form');
const toast = document.getElementById('toast');

// Metrics
const metricTotalClients = document.getElementById('metric-total-clients');
const metricTotalEquipments = document.getElementById('metric-total-equipments');
const metricAvgEquipments = document.getElementById('metric-avg-equipments');

// Table and Chart reference
const tbody = document.getElementById('top-clients-tbody');
let segmentChartInstance = null;

// State Data
let appData = []; // Array of client objects

// Initialize App
document.addEventListener('DOMContentLoaded', () => {
    initNavigation();
    initForm();
    loadLocalStoreData();
    attemptFetchExcel();
});

// Navigation Handling
function initNavigation() {
    btnDashboard.addEventListener('click', () => switchView('dashboard'));
    btnNewClient.addEventListener('click', () => switchView('new-client'));
}

function switchView(view) {
    if (view === 'dashboard') {
        btnDashboard.classList.add('active');
        btnNewClient.classList.remove('active');
        viewDashboard.classList.add('active');
        viewDashboard.classList.remove('hidden');
        viewNewClient.classList.remove('active');
        viewNewClient.classList.add('hidden');
        document.getElementById('page-title').innerText = 'Resumen de Ventas y Análisis';
        document.getElementById('page-subtitle').innerText = 'Información del Pipeline Comercial Corporativo';
        uploadContainer.style.display = 'block';
    } else {
        btnNewClient.classList.add('active');
        btnDashboard.classList.remove('active');
        viewNewClient.classList.add('active');
        viewNewClient.classList.remove('hidden');
        viewDashboard.classList.remove('active');
        viewDashboard.classList.add('hidden');
        document.getElementById('page-title').innerText = 'Registro CRM';
        document.getElementById('page-subtitle').innerText = 'Añadir nuevo registro al pipeline comercial';
        uploadContainer.style.display = 'none';
    }
}

// Data Processing Flow
function attemptFetchExcel() {
    // Attempt to read the .xlsx from the same directory in case user uses Local Server
    const filePath = 'PIPELINE COMERCIAL CORPORATIVO  PROYECTO CRM.xlsx';
    
    fetch(filePath)
    .then(response => {
        if (!response.ok) throw new Error("Fetch not supported for local files without server");
        return response.arrayBuffer();
    })
    .then(data => {
        // Automatically hides the upload container, because it fetched successfully
        uploadContainer.style.display = 'none';
        parseExcelData(data);
    })
    .catch(err => {
        console.warn("Could not auto-fetch XLSX. Proceeding with upload fallback.", err);
        // Show upload container
        initExcelUpload();
    });
}

function initExcelUpload() {
    fileInput.addEventListener('change', (e) => {
        const file = e.target.files[0];
        if (!file) return;
        const reader = new FileReader();
        reader.onload = function(evt) {
            const data = evt.target.result;
            parseExcelData(data);
        };
        reader.readAsArrayBuffer(file);
    });
}

// SheetJS Parsing
function parseExcelData(dataBuffer) {
    try {
        const workbook = XLSX.read(dataBuffer, { type: 'array' });
        const firstSheetName = workbook.SheetNames[0];
        const worksheet = workbook.Sheets[firstSheetName];
        // Convert to JSON
        let jsonData = XLSX.utils.sheet_to_json(worksheet);
        
        // Transform keys to remove trailing spaces (the excel has "Nombre ", "Cargo ", etc.)
        const cleanedData = jsonData.map(row => {
            let processedRow = {};
            for(let key in row) {
                processedRow[key.trim()] = row[key];
            }
            return processedRow;
        });

        // Integrate with localStorage state
        mergeDataAndRefresh(cleanedData);
    } catch(err) {
        console.error("Error reading Excel", err);
        alert("Ocurrió un error leyendo el archivo Excel. Verifique el formato.");
    }
}

// Local Storage for New Clients
function loadLocalStoreData() {
    const local = localStorage.getItem('nexcrm_data');
    if (local) {
        appData = JSON.parse(local);
        // If we have just local data (and haven't fetched excel yet), refresh UI with this minimal data.
        updateDashboard();
    }
}

function mergeDataAndRefresh(excelData) {
    const local = localStorage.getItem('nexcrm_data');
    const localData = local ? JSON.parse(local) : [];

    // Combine both sets: Excel Data + Local Form Data
    appData = [...excelData, ...localData];
    updateDashboard();
}

// Visualizations and Metrics Update
function updateDashboard() {
    if (!appData || appData.length === 0) return;

    let totalClients = appData.length;
    let totalEquipments = 0;
    
    // Segment mapping setup
    const segmentMap = {};
    const topClients = [];

    appData.forEach(row => {
        // Standardize properties that might be different due to form VS excel
        const clientName = row['Cliente'] || row['Nombre de Cliente / Empresa'] || 'Desconocido';
        const segment = row['Segmento'] || 'Otro';
        let equipmentsStr = row['Cantidad equipos cotizados'];
        let equipments = parseFloat(equipmentsStr);
        if (isNaN(equipments)) equipments = 0;

        totalEquipments += equipments;

        // Process segment count
        segmentMap[segment] = (segmentMap[segment] || 0) + 1;

        // Collect for ranking
        topClients.push({
            name: clientName,
            segment: segment,
            equipments: equipments
        });
    });

    const avgEquipments = totalClients > 0 ? (totalEquipments / totalClients).toFixed(1) : 0;

    // Update Widget DOM
    metricTotalClients.innerText = totalClients;
    metricTotalEquipments.innerText = totalEquipments;
    metricAvgEquipments.innerText = avgEquipments;

    // Build Chart
    updateChart(segmentMap);

    // Build Top 10 Ranking Table
    topClients.sort((a, b) => b.equipments - a.equipments);
    const top10 = topClients.slice(0, 10);
    renderTable(top10);
}

function renderTable(top10Array) {
    tbody.innerHTML = '';
    top10Array.forEach((c, index) => {
        const tr = document.createElement('tr');
        tr.innerHTML = `
            <td>#${index + 1}</td>
            <td>${c.name}</td>
            <td><span style="background: rgba(255,255,255,0.1); padding: 4px 8px; border-radius: 6px; font-size: 0.8rem;">${c.segment}</span></td>
            <td style="font-weight: 700;">${c.equipments}</td>
        `;
        tbody.appendChild(tr);
    });
}

function updateChart(segmentMap) {
    const ctx = document.getElementById('segmentChart').getContext('2d');
    const labels = Object.keys(segmentMap);
    const dataVals = Object.values(segmentMap);

    if (segmentChartInstance) {
        segmentChartInstance.destroy();
    }

    segmentChartInstance = new Chart(ctx, {
        type: 'doughnut',
        data: {
            labels: labels,
            datasets: [{
                data: dataVals,
                backgroundColor: [
                    '#4F46E5', '#10B981', '#F43F5E', '#F59E0B', '#8B5CF6', '#EC4899', '#06B6D4'
                ],
                borderWidth: 0,
                hoverOffset: 10
            }]
        },
        options: {
            responsive: true,
            maintainAspectRatio: false,
            plugins: {
                legend: {
                    position: 'right',
                    labels: { color: '#f8fafc', font: { family: 'Plus Jakarta Sans' } }
                }
            },
            cutout: '75%'
        }
    });
}

// Form Handling
function initForm() {
    formNewClient.addEventListener('submit', (e) => {
        e.preventDefault();
        
        const newRecord = {
            'Nombre': document.getElementById('frm-nombre').value,
            'Cargo': document.getElementById('frm-cargo').value,
            'Cliente': document.getElementById('frm-cliente').value,
            'Número de teléfono': document.getElementById('frm-telefono').value,
            'Correo': document.getElementById('frm-correo').value,
            'Origen': document.getElementById('frm-origen').value,
            'Segmento': document.getElementById('frm-segmento').value,
            'Cantidad equipos cotizados': document.getElementById('frm-equipos').value
        };

        // --- Integración con HubSpot ---
        const portalId = "51270682";
        const formId = "68355e33-40e1-4697-bbd8-326e2f0ca9de";
        const hsUrl = `https://api.hsforms.com/submissions/v3/integration/submit/${portalId}/${formId}`;
        
        const hsData = {
            fields: [
                { name: "firstname", value: newRecord["Nombre"] },
                { name: "jobtitle", value: newRecord["Cargo"] },
                { name: "company", value: newRecord["Cliente"] },
                { name: "phone", value: newRecord["Número de teléfono"] },
                { name: "email", value: newRecord["Correo"] },
                { name: "origen", value: newRecord["Origen"] },
                { name: "segmento", value: newRecord["Segmento"] },
                { name: "cantidad_equipos_cotizados", value: newRecord["Cantidad equipos cotizados"] }
            ],
            context: {
                pageUri: window.location.href,
                pageName: "Registro CRM Interno"
            }
        };

        fetch(hsUrl, {
            method: 'POST',
            headers: {
                'Content-Type': 'application/json'
            },
            body: JSON.stringify(hsData)
        }).then(response => {
            console.log("Enviado a Hubspot con estado:", response.status);
        }).catch(err => {
            console.error("Error conectando a Hubspot:", err);
        });
        // ---------------------------------

        // Save to local storage state
        const local = localStorage.getItem('nexcrm_data');
        const localData = local ? JSON.parse(local) : [];
        localData.push(newRecord);
        localStorage.setItem('nexcrm_data', JSON.stringify(localData));

        // Add to global state so it renders if dashboard is refreshed this session
        appData.push(newRecord);
        updateDashboard();

        // Notification and form clearing
        showToast();
        formNewClient.reset();

        // Optional: Switch back to dashboard after 1.5 seconds
        setTimeout(() => {
            switchView('dashboard');
        }, 1500);
    });
}

function showToast() {
    toast.classList.remove('hidden');
    toast.classList.add('show');
    setTimeout(() => {
        toast.classList.remove('show');
        setTimeout(() => toast.classList.add('hidden'), 400); // Wait for transition
    }, 3000);
}
