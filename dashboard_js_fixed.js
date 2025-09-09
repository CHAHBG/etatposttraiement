/**
 * @typedef {Object} CommuneData
 * @property {string} commune
 * @property {string} region
 * @property {number} totalParcelles
 * @property {number} percentTotal
 * @property {number} nicad
 * @property {number} percentNicad
 * @property {number} ctasf
 * @property {number} percentCtasf
 * @property {number} deliberated
 * @property {number} percentDeliberated
 * @property {number} parcellesBrutes
 * @property {number} collected
 * @property {number} surveyed
 * @property {string} rejectionReasons
 * @property {number} retained
 * @property {number} validated
 * @property {number} rejected
 * @property {string} urmRejectionReasons
 * @property {number} corrected
 * @property {string} geomaticien
 * @property {number} individualJoined
 * @property {number} collectiveJoined
 * @property {number} unjoined
 * @property {number} duplicatesRemoved
 * @property {number} duplicateRemovalRate
 * @property {number} parcelsInConflict
 * @property {number} significantDuplicates
 * @property {number} postProcessedLot1_46
 * @property {string} jointureStatus
 * @property {string} jointureErrorMessage
 */

/** @type {CommuneData[]} */
let communesData = [];
const GOOGLE_SHEET_CSV_URL = 'https://docs.google.com/spreadsheets/d/e/2PACX-1vTcSN5w8XmgV1BO_alL_yFDXiSnSaad9sTxj2z16GoS3IglT9FKu6fT9R41YHo21Q/pub?output=csv';
const AUTO_REFRESH_INTERVAL = 300000; // 5 minutes in ms

// Color scheme based on the provided branding
const CHART_COLORS = {
    primary: '#4CAF50',        // Green
    secondary: '#0072BC',      // Blue
    accent: '#A52A2A',         // Brown/Red
    earth: '#6B4226',          // Earth brown
    highlight: '#FFD700',      // Yellow
    darkBlue: '#1C3D72',       // Dark blue
    success: '#10b981',        // Success green
    warning: '#f59e0b',        // Warning orange
    danger: '#ef4444',         // Danger red
    info: '#3b82f6',           // Info blue
    gray: '#6B7280'            // Gray
};

// Chart.js global configuration
Chart.defaults.font.family = "'Segoe UI', 'Roboto', 'Helvetica Neue', Arial, sans-serif";
Chart.defaults.font.size = 12;
Chart.defaults.color = '#333333';
Chart.defaults.plugins.tooltip.backgroundColor = 'rgba(28, 61, 114, 0.8)';
Chart.defaults.plugins.tooltip.titleColor = '#ffffff';
Chart.defaults.plugins.tooltip.bodyColor = '#ffffff';
Chart.defaults.plugins.tooltip.borderColor = '#ffffff';
Chart.defaults.plugins.tooltip.borderWidth = 1;
Chart.defaults.plugins.tooltip.padding = 10;
Chart.defaults.plugins.tooltip.cornerRadius = 8;

/** @type {{ [key: string]: Chart }} */
const charts = {
    status: null,
    communes: null,
    efficiency: null,
    rejection: null,
    workload: null,
    jointure: null,
    conflict: null,
    validationTrend: null
};

/** @type {Object.<string, boolean>} */
const columnVisibility = {
    commune: true,
    region: true,
    totalParcelles: true,
    percentTotal: true,
    nicad: true,
    percentNicad: true,
    ctasf: true,
    percentCtasf: true,
    deliberated: true,
    percentDeliberated: true,
    parcellesBrutes: true,
    collected: true,
    surveyed: true,
    rejectionReasons: true,
    retained: true,
    validated: true,
    rejected: true,
    urmRejectionReasons: true,
    corrected: true,
    geomaticien: true,
    individualJoined: true,
    collectiveJoined: true,
    unjoined: true,
    duplicatesRemoved: true,
    duplicateRemovalRate: true,
    parcelsInConflict: true,
    significantDuplicates: true,
    postProcessedLot1_46: true,
    jointureStatus: true,
    jointureErrorMessage: true
};

/**
 * The default order of columns (used for reset)
 * @type {string[]}
 */
const defaultColumnOrder = [
    'commune', 'region', 'totalParcelles', 'percentTotal', 'nicad', 
    'percentNicad', 'ctasf', 'percentCtasf', 'deliberated', 'percentDeliberated', 
    'parcellesBrutes', 'collected', 'surveyed', 'rejectionReasons', 'retained', 
    'validated', 'rejected', 'urmRejectionReasons', 'corrected', 'geomaticien', 
    'individualJoined', 'collectiveJoined', 'unjoined', 'duplicatesRemoved', 
    'duplicateRemovalRate', 'parcelsInConflict', 'significantDuplicates', 
    'postProcessedLot1_46', 'jointureStatus', 'jointureErrorMessage'
];

/**
 * The current order of columns (may be modified by user)
 * @type {string[]}
 */
let currentColumnOrder = [...defaultColumnOrder];

/**
 * Determines progress bar color based on percentage
 * @param {number} percentage
 * @returns {string}
 */
function getProgressBarColor(percentage) {
    if (percentage < 33) return '#ef4444'; // Red
    if (percentage < 66) return '#f59e0b'; // Orange
    return '#10b981'; // Green
}

/**
 * Fetches and parses commune data from Google Sheets
 * @param {boolean} showLoading
 * @returns {Promise<void>}
 */
async function fetchCommunesData(showLoading = true) {
    if (showLoading) showLoadingIndicator();
    try {
        const response = await fetch(GOOGLE_SHEET_CSV_URL, {
            cache: 'no-cache',
            headers: { 'Accept': 'text/csv' }
        });
        if (!response.ok) throw new Error(`HTTP ${response.status}: ${response.statusText}`);
        const csvText = await response.text();
        communesData = parseCSV(csvText);
        refreshDashboard();
        if (showLoading) {
            hideLoadingIndicator();
            showNotification('Data refreshed successfully', 'success');
        }
        updateLastRefreshTime();
    } catch (error) {
        console.error('Error fetching data:', error);
        if (showLoading) {
            hideLoadingIndicator();
            showNotification(`Failed to refresh data: ${error.message}`, 'error');
        }
    }
}

/**
 * Parses CSV data into structured objects
 * @param {string} csv
 * @returns {CommuneData[]}
 */
function parseCSV(csv) {
    if (!csv.trim()) return [];
    
    const sep = csv.includes(';') && csv.split(';').length > csv.split(',').length ? ';' : ',';
    const lines = csv.split(/\r?\n/).filter(line => line.trim());
    if (lines.length < 2) return [];

    const splitLine = (line) => {
        const result = [];
        let current = '';
        let inQuotes = false;
        
        for (let i = 0; i < line.length; i++) {
            const char = line[i];
            if (char === '"') {
                inQuotes = !inQuotes;
            } else if (char === sep && !inQuotes) {
                result.push(current.trim());
                current = '';
            } else {
                current += char;
            }
        }
        result.push(current.trim());
        return result;
    };

    const headers = splitLine(lines[0]);
    const headerMap = {
    'Commune': 'commune',
    'Région': 'region',
        'Total Parcelles': 'totalParcelles',
        '% du total': 'percentTotal',
        'NICAD': 'nicad',
        '% NICAD': 'percentNicad',
        'CTASF': 'ctasf',
        '% CTASF': 'percentCtasf',
    'Délibérées': 'deliberated',
    '% Délibérée': 'percentDeliberated',
        'Parcelles brutes': 'parcellesBrutes',
    'Parcelles collectées (sans doublon géométrique)': 'collected',
    'Parcelles enquêtées': 'surveyed',
        'Motifs de rejet post-traitement': 'rejectionReasons',
        'Parcelles retenues après post-traitement': 'retained',
    'Parcelles validées par l\'URM': 'validated',
    'Parcelles rejetées par l\'URM': 'rejected',
        'Motifs de rejet URM': 'urmRejectionReasons',
    'Parcelles corrigées': 'corrected',
        'Geomaticien': 'geomaticien',
        'Parcelles individuelles jointes': 'individualJoined',
        'Parcelles collectives jointes': 'collectiveJoined',
        'Parcelles non jointes': 'unjoined',
    'Doublons supprimés': 'duplicatesRemoved',
        'Taux suppression doublons (%)': 'duplicateRemovalRate',
        'Parcelles en conflit': 'parcelsInConflict',
        'Significant Duplicates': 'significantDuplicates',
    'Parc. post-traitées lot 1-46': 'postProcessedLot1_46',
        'Statut jointure': 'jointureStatus',
    'Message d’erreur jointure': 'jointureErrorMessage'
    };

    const headerIndexes = headers.map(h => headerMap[h] || h);

    const data = lines.slice(1).map(line => {
        const values = splitLine(line);
        if (values.length < 2) return null;
        const obj = {};
        headerIndexes.forEach((key, i) => {
            let value = values[i]?.trim() ?? '';
            if ([
                'totalParcelles', 'nicad', 'ctasf', 'deliberated', 'parcellesBrutes', 'collected', 'surveyed', 'retained', 'validated', 'rejected', 'corrected', 'individualJoined', 'collectiveJoined', 'unjoined', 'duplicatesRemoved', 'parcelsInConflict', 'significantDuplicates', 'postProcessedLot1_46'
            ].includes(key)) {
                value = Number(value.replace(/[^0-9.-]/g, '').replace(/\s/g, '')) || 0;
            } else if ([
                'percentTotal', 'percentNicad', 'percentCtasf', 'percentDeliberated', 'duplicateRemovalRate'
            ].includes(key)) {
                value = Number(value.replace(/%/g, '').replace(/,/g, '.')) || 0;
            }
            obj[key] = value;
        });
        return obj;
    }).filter(Boolean);

    if (!data.length) {
        const err = document.createElement('div');
        err.id = 'csvError';
        err.style.cssText = 'color: red; font-size: 1.2rem; margin: 20px;';
        err.textContent = 'Aucune donnée chargée depuis Google Sheet. Vérifiez le format de la feuille.';
        document.body.prepend(err);
    }

    return data;
}

/**
 * Updates all dashboard components
 */
function refreshDashboard() {
    updateMetrics();
    populateCommunesTable();
    generateProgressCards();
    generateActivityList();
    Object.values(charts).forEach(chart => chart?.destroy());
    initializeCharts();
    initializeColumnFilter();
    updateTableColumns();
}

/**
 * Initializes the dashboard
 * @returns {Promise<void>}
 */
async function init() {
    try {
        await fetchCommunesData();
        initializeDashboard();
        setupEventListeners();
        loadOverviewData();
        setInterval(() => fetchCommunesData(false), AUTO_REFRESH_INTERVAL);
    } catch (error) {
        console.error('Initialization error:', error);
        showNotification('Failed to initialize dashboard', 'error');
    }
}

document.addEventListener('DOMContentLoaded', init);

/**
 * Sets up initial dashboard state
 */
function initializeDashboard() {
    showTab('overview');
}

/**
 * Sets up event listeners for UI interactions
 */
function setupEventListeners() {
    document.querySelectorAll('.nav-item').forEach(link => {
        link.addEventListener('click', e => {
            e.preventDefault();
            const tabName = link.dataset.tab;
            showTab(tabName);
            document.querySelectorAll('.nav-item').forEach(l => l.classList.remove('active'));
            link.classList.add('active');
            document.getElementById('page-title').textContent = {
                overview: 'Vue d\'ensemble',
                communes: 'Détails des communes',
                progress: 'Avancement du traitement',
                analytics: 'Analytique & Rapports',
                reports: 'Générer des rapports'
            }[tabName];
            if (tabName === 'analytics') {
                initializeAnalyticsCharts();
            }
            if (tabName === 'communes') {
                initializeColumnFilter();
                updateTableColumns();
                setupDraggableColumns();
            }
        });
    });
    
    // Toggle insights panel
    const toggleInsightsBtn = document.getElementById('toggleInsights');
    const insightsContent = document.getElementById('insightsContent');
    if (toggleInsightsBtn && insightsContent) {
        toggleInsightsBtn.addEventListener('click', () => {
            const isVisible = insightsContent.style.display !== 'none';
            insightsContent.style.display = isVisible ? 'none' : 'grid';
                toggleInsightsBtn.innerHTML = isVisible ? 
                '<i class="fas fa-chevron-down"></i> Voir détails' : 
                '<i class="fas fa-chevron-up"></i> Masquer détails';
        });
    }
    
    // Add chart download functionality
    document.querySelectorAll('.chart-download').forEach(button => {
        button.addEventListener('click', function() {
            const chartId = this.dataset.chart;
            downloadChart(chartId);
        });
    });
    
    // Add chart resize listeners to improve rendering
    window.addEventListener('resize', debounce(() => {
        if (document.getElementById('analytics').classList.contains('active')) {
            initializeAnalyticsCharts();
        }
        if (document.getElementById('overview').classList.contains('active')) {
            initializeCharts();
        }
    }, 250));

    const sidebarToggle = document.querySelector('.sidebar-toggle');
    sidebarToggle?.addEventListener('click', () => {
        document.querySelector('.sidebar').classList.toggle('mobile-open');
    });

    const searchInput = document.querySelector('.search-box input');
    searchInput?.addEventListener('input', debounce(function() {
        filterCommunesTable(this.value.toLowerCase());
    }, 300));

    document.getElementById('exportBtn')?.addEventListener('click', () => exportWithColumnOrder());
    document.querySelector('.refresh-button')?.addEventListener('click', manualRefresh);

    document.querySelectorAll('.report-card .btn').forEach(btn => {
        btn.addEventListener('click', () => generateReport(btn.closest('.report-card').querySelector('h4').textContent));
    });
}

/**
 * Initializes column filter dropdown
 */
function initializeColumnFilter() {
    const tableContainer = document.querySelector('#communes .table-container');
    if (!tableContainer) {
        console.warn('Table container not found for column filter');
        return;
    }

    // Remove existing filter to prevent duplicates
    const existingFilter = document.getElementById('columnFilter');
    if (existingFilter) existingFilter.remove();

    const headers = [
        'Commune', 'Région', 'Total Parcelles', '% du Total', 'NICAD', '% NICAD', 'CTASF', '% CTASF',
        'Délibérées', '% Délibérée', 'Parcelles brutes', 'Parcelles collectées', 'Parcelles enquêtées',
        'Motifs de rejet post-traitement', 'Parcelles retenues', 'Parcelles validées', 'Parcelles rejetées',
        'Motifs de rejet URM', 'Parcelles corrigées', 'Géomaticien', 'Parcelles individuelles jointes',
        'Parcelles collectives jointes', 'Parcelles non jointes', 'Doublons supprimés', 'Taux suppression doublons (%)',
        'Parcelles en conflit', 'Significant Duplicates', 'Parc. post-traitées lot 1-46', 'Statut jointure',
        'Message d\'erreur jointure'
    ];

    const headerKeys = [
        'commune', 'region', 'totalParcelles', 'percentTotal', 'nicad', 'percentNicad', 'ctasf', 'percentCtasf',
        'deliberated', 'percentDeliberated', 'parcellesBrutes', 'collected', 'surveyed', 'rejectionReasons',
        'retained', 'validated', 'rejected', 'urmRejectionReasons', 'corrected', 'geomaticien', 'individualJoined',
        'collectiveJoined', 'unjoined', 'duplicatesRemoved', 'duplicateRemovalRate', 'parcelsInConflict',
        'significantDuplicates', 'postProcessedLot1_46', 'jointureStatus', 'jointureErrorMessage'
    ];

    const filterContainer = document.createElement('div');
    filterContainer.id = 'columnFilter';
    filterContainer.style.marginBottom = '1rem';
    filterContainer.innerHTML = `
        <div class="dropdown">
                <button class="btn btn-secondary dropdown-toggle" type="button" id="columnFilterBtn" data-bs-toggle="dropdown" aria-expanded="false">
                Sélectionner les colonnes
            </button>
            <div class="dropdown-menu" style="max-height: 400px; overflow-y: auto;">
                ${headers.map((header, index) => `
                    <div class="dropdown-item">
                        <input type="checkbox" id="col-${headerKeys[index]}" ${columnVisibility[headerKeys[index]] ? 'checked' : ''} data-column="${headerKeys[index]}">
                        <label for="col-${headerKeys[index]}">${header}</label>
                    </div>
                `).join('')}
            </div>
        </div>
    `;
    tableContainer.insertBefore(filterContainer, tableContainer.querySelector('table'));

    // Use event delegation for checkbox events
    filterContainer.addEventListener('change', (event) => {
        if (event.target.type === 'checkbox') {
            const columnKey = event.target.dataset.column;
            columnVisibility[columnKey] = event.target.checked;
            updateTableColumns();
        }
    });
}

/**
 * Updates table column visibility
 */
function updateTableColumns() {
    const table = document.querySelector('#communesTable');
    if (!table) {
        console.warn('Communes table not found');
        return;
    }

    // Use currentColumnOrder instead of fixed headerKeys array
    const headers = table.querySelectorAll('th');
    const rows = table.querySelectorAll('#communesTableBody tr');

    headers.forEach((th) => {
        const key = th.dataset.column;
        if (key && columnVisibility.hasOwnProperty(key)) {
            th.classList.toggle('hidden-column', !columnVisibility[key]);
        }
    });

    rows.forEach(row => {
        const cells = row.querySelectorAll('td');
        cells.forEach((cell) => {
            const key = cell.dataset.column;
            if (key && columnVisibility.hasOwnProperty(key)) {
                cell.classList.toggle('hidden-column', !columnVisibility[key]);
            }
        });
    });
    
    // Apply column order
    applyColumnOrder();
}

/**
 * Sets up drag and drop functionality for table columns
 */
function setupDraggableColumns() {
    const headerCells = document.querySelectorAll('#communesTable th[draggable="true"]');
    let draggedCell = null;
    
    // Reset button functionality
    document.getElementById('resetColumnOrder')?.addEventListener('click', () => {
        currentColumnOrder = [...defaultColumnOrder];
        applyColumnOrder();
    showNotification('Ordre des colonnes réinitialisé', 'success');
        // Save column order to local storage
        localStorage.setItem('columnOrder', JSON.stringify(currentColumnOrder));
    });
    
    // Try to load column order from localStorage
    try {
        const savedOrder = localStorage.getItem('columnOrder');
        if (savedOrder) {
            const parsedOrder = JSON.parse(savedOrder);
            // Verify the saved order contains all columns
            const hasAllColumns = defaultColumnOrder.every(col => parsedOrder.includes(col));
            if (hasAllColumns && parsedOrder.length === defaultColumnOrder.length) {
                currentColumnOrder = parsedOrder;
                applyColumnOrder();
            }
        }
    } catch (err) {
        console.error('Error loading saved column order:', err);
    }

    headerCells.forEach(cell => {
        // Dragstart event - when user starts dragging
        cell.addEventListener('dragstart', function(e) {
            draggedCell = this;
            this.classList.add('dragging');
            e.dataTransfer.effectAllowed = 'move';
            e.dataTransfer.setData('text/plain', this.dataset.column);
        });
        
        // Dragend event - when user stops dragging
        cell.addEventListener('dragend', function() {
            draggedCell = null;
            headerCells.forEach(c => c.classList.remove('dragging', 'drop-target'));
        });
        
        // Dragover event - when dragged element is over a drop target
        cell.addEventListener('dragover', function(e) {
            if (draggedCell && draggedCell !== this) {
                e.preventDefault();
            }
        });
        
        // Dragenter event - when dragged element enters a drop target
        cell.addEventListener('dragenter', function(e) {
            if (draggedCell && draggedCell !== this) {
                e.preventDefault();
                this.classList.add('drop-target');
            }
        });
        
        // Dragleave event - when dragged element leaves a drop target
        cell.addEventListener('dragleave', function() {
            this.classList.remove('drop-target');
        });
        
        // Drop event - when user releases the dragged element over a drop target
        cell.addEventListener('drop', function(e) {
            e.preventDefault();
            if (draggedCell && draggedCell !== this) {
                const fromColumn = draggedCell.dataset.column;
                const toColumn = this.dataset.column;
                
                // Update column order
                reorderColumns(fromColumn, toColumn);
                
                // Apply new column order
                applyColumnOrder();
                
                // Save column order to local storage
                localStorage.setItem('columnOrder', JSON.stringify(currentColumnOrder));
                
                // Show feedback to user
                showNotification('Ordre des colonnes mis à jour', 'success');
            }
            this.classList.remove('drop-target');
        });
    });
}

/**
 * Reorders columns by moving fromColumn to before toColumn
 * @param {string} fromColumn The column being moved
 * @param {string} toColumn The target column position
 */
function reorderColumns(fromColumn, toColumn) {
    // Get current indexes
    const fromIndex = currentColumnOrder.indexOf(fromColumn);
    const toIndex = currentColumnOrder.indexOf(toColumn);
    
    if (fromIndex !== -1 && toIndex !== -1) {
        // Remove the column from its current position
        currentColumnOrder.splice(fromIndex, 1);
        
        // Calculate the new position (if we're moving right, we need to adjust for the removed item)
        const newToIndex = fromIndex < toIndex ? toIndex - 1 : toIndex;
        
        // Insert the column at the new position
        currentColumnOrder.splice(newToIndex, 0, fromColumn);
    }
}

/**
 * Applies the current column order to the table
 */
function applyColumnOrder() {
    const thead = document.querySelector('#communesTable thead tr');
    const tbody = document.getElementById('communesTableBody');
    
    if (!thead || !tbody) return;
    
    // Get all header cells
    const headerCells = Array.from(thead.querySelectorAll('th'));
    
    // Sort header cells according to currentColumnOrder
    headerCells.sort((a, b) => {
        const aIndex = currentColumnOrder.indexOf(a.dataset.column);
        const bIndex = currentColumnOrder.indexOf(b.dataset.column);
        return aIndex - bIndex;
    });
    
    // Append header cells in new order
    headerCells.forEach(cell => thead.appendChild(cell));
    
    // Update body rows to match header order
    Array.from(tbody.querySelectorAll('tr')).forEach(row => {
        const cells = Array.from(row.querySelectorAll('td'));
        
        // Sort cells according to currentColumnOrder
        cells.sort((a, b) => {
            const aIndex = currentColumnOrder.indexOf(a.dataset.column);
            const bIndex = currentColumnOrder.indexOf(b.dataset.column);
            return aIndex - bIndex;
        });
        
        // Append cells in new order
        cells.forEach(cell => row.appendChild(cell));
    });
}

/**
 * Debounces a function
 * @param {Function} func
 * @param {number} wait
 * @returns {Function}
 */
function debounce(func, wait) {
    let timeout;
    return function executedFunction(...args) {
        const later = () => {
            clearTimeout(timeout);
            func.apply(this, args);
        };
        clearTimeout(timeout);
        timeout = setTimeout(later, wait);
    };
}

/**
 * Shows specified tab
 * @param {string} tabName
 */
function showTab(tabName) {
    document.querySelectorAll('.tab-content').forEach(tab => tab.classList.remove('active'));
    const tab = document.getElementById(tabName);
    if (tab) {
        tab.classList.add('active');
    }
}

/**
 * Updates dashboard metrics
 */
function updateMetrics() {
    const totals = calculateTotals();
    const stats = calculateAdditionalStats();

    const elements = {
        totalCollected: totals.collected,
        totalSurveyed: totals.surveyed,
        totalValidated: totals.validated,
        totalRejected: totals.rejected,
        avgValidationRate: stats.avgValidation,
        fastestCommune: stats.fastestCommune,
        slowestCommune: stats.slowestCommune
    };

    Object.entries(elements).forEach(([id, value]) => {
        const element = document.getElementById(id);
        if (element) element.textContent = typeof value === 'number' ? value.toLocaleString() : value;
    });

    const goal = 70000;
    const validationVsGoal = goal > 0 ? ((totals.validated / goal) * 100).toFixed(1) : '0.0';
    const deliberatedVsCollected = totals.collected > 0 ? ((totals.deliberated / totals.collected) * 100).toFixed(1) : '0.0';
    const ctasfVsCollected = totals.collected > 0 ? ((totals.ctasf / totals.collected) * 100).toFixed(1) : '0.0';
    const ctasfVsValidated = totals.validated > 0 ? ((totals.ctasf / totals.validated) * 100).toFixed(1) : '0.0';
    const deliberatedVsValidated = totals.validated > 0 ? ((totals.deliberated / totals.validated) * 100).toFixed(1) : '0.0';

    const updateProgressBar = (id, fillId, value) => {
        const element = document.getElementById(id);
        const fillElement = document.getElementById(fillId);
        if (element) element.textContent = `${value}%`;
        if (fillElement) {
            fillElement.style.width = `${Math.min(value, 100)}%`;
            fillElement.style.backgroundColor = getProgressBarColor(value);
        }
    };

    const progressStats = document.querySelector('.progress-stats');
    if (progressStats && !document.getElementById('validationVsGoal')) {
        const extra = document.createElement('div');
        extra.className = 'progress-extra-metrics';
        extra.innerHTML = `
            <div class="progress-item">
                <span class="progress-label">Taux de validation par rapport à l'objectif (70 000)</span>
                <div class="progress-bar">
                    <div class="progress-fill" id="validationVsGoalFill"></div>
                </div>
                <span class="progress-value" id="validationVsGoal">${validationVsGoal}%</span>
            </div>
            <div class="progress-item">
                <span class="progress-label">Délibérées / Collectées</span>
                <div class="progress-bar">
                    <div class="progress-fill" id="deliberatedVsCollectedFill"></div>
                </div>
                <span class="progress-value" id="deliberatedVsCollected">${deliberatedVsCollected}%</span>
            </div>
            <div class="progress-item">
                <span class="progress-label">CTASF / Collectées</span>
                <div class="progress-bar">
                    <div class="progress-fill" id="ctasfVsCollectedFill"></div>
                </div>
                <span class="progress-value" id="ctasfVsCollected">${ctasfVsCollected}%</span>
            </div>
            <div class="progress-item">
                <span class="progress-label">CTASF / Validées</span>
                <div class="progress-bar">
                    <div class="progress-fill" id="ctasfVsValidatedFill"></div>
                </div>
                <span class="progress-value" id="ctasfVsValidated">${ctasfVsValidated}%</span>
            </div>
            <div class="progress-item">
                <span class="progress-label">Délibérées / Validées</span>
                <div class="progress-bar">
                    <div class="progress-fill" id="deliberatedVsValidatedFill"></div>
                </div>
                <span class="progress-value" id="deliberatedVsValidated">${deliberatedVsValidated}%</span>
            </div>
        `;
        progressStats.appendChild(extra);
    } else if (progressStats) {
        updateProgressBar('validationVsGoal', 'validationVsGoalFill', validationVsGoal);
        updateProgressBar('deliberatedVsCollected', 'deliberatedVsCollectedFill', deliberatedVsCollected);
        updateProgressBar('ctasfVsCollected', 'ctasfVsCollectedFill', ctasfVsCollected);
        updateProgressBar('ctasfVsValidated', 'ctasfVsValidatedFill', ctasfVsValidated);
        updateProgressBar('deliberatedVsValidated', 'deliberatedVsValidatedFill', deliberatedVsValidated);
    }
}

/**
 * Calculates total metrics
 * @returns {Object}
 */
function calculateTotals() {
    return communesData.reduce((totals, commune) => {
        if (!commune.commune || commune.commune.trim().toLowerCase() === 'total') return totals;
        totals.collected += Number(commune.collected) || 0;
        totals.surveyed += Number(commune.surveyed) || 0;
        totals.validated += Number(commune.validated) || 0;
        totals.rejected += Number(commune.rejected) || 0;
        totals.parcellesBrutes += Number(commune.parcellesBrutes) || 0;
        totals.retained += Number(commune.retained) || 0;
        totals.individualJoined += Number(commune.individualJoined) || 0;
        totals.collectiveJoined += Number(commune.collectiveJoined) || 0;
        totals.unjoined += Number(commune.unjoined) || 0;
        totals.parcelsInConflict += Number(commune.parcelsInConflict) || 0;
        totals.postProcessedLot1_46 += Number(commune.postProcessedLot1_46) || 0;
        totals.deliberated += Number(commune.deliberated) || 0;
        totals.ctasf += Number(commune.ctasf) || 0;
        return totals;
    }, {
        collected: 0, surveyed: 0, validated: 0, rejected: 0,
        parcellesBrutes: 0, retained: 0, individualJoined: 0,
        collectiveJoined: 0, unjoined: 0, parcelsInConflict: 0,
        postProcessedLot1_46: 0, deliberated: 0, ctasf: 0
    });
}

/**
 * Calculates additional statistics
 * @returns {Object}
 */
function calculateAdditionalStats() {
    const totals = calculateTotals();
    let avgValidation = totals.collected > 0 ? ((totals.validated / totals.collected) * 100).toFixed(1) + '%' : '--%';
    let maxRate = -1, minRate = Infinity;
    let fastest = '-', slowest = '-';
    
    communesData.forEach(commune => {
        if (commune.collected > 0) {
            const rate = commune.validated / commune.collected;
            if (rate > maxRate) {
                maxRate = rate;
                fastest = commune.commune;
            }
            if (rate < minRate) {
                minRate = rate;
                slowest = commune.commune;
            }
        }
    });
    
    return { avgValidation, fastestCommune: fastest, slowestCommune: slowest };
}

/**
 * Loads overview data
 */
function loadOverviewData() {
    console.log('Overview data loaded');
}

/**
 * Initializes all charts
 */
function initializeCharts() {
    initializeStatusChart();
    initializeCommunesChart();
    
    // Add download capability to overview charts
    setupChartDownloadButtons();
}

/**
 * Downloads chart as image
 * @param {string} chartId Chart canvas ID
 */
function downloadChart(chartId) {
    const canvas = document.getElementById(chartId);
    if (!canvas) return;
    
    const chartName = {
        'efficiencyChart': 'Analyse_Efficacité',
        'rejectionChart': 'Motifs_Rejet',
        'workloadChart': 'Charge_Géomaticiens',
        'jointureChart': 'Statuts_Jointure',
        'conflictHeatmap': 'Conflits_Parcelles',
        'validationTrendChart': 'Tendance_Validation',
        'statusChart': 'Statuts_Traitement',
        'communesChart': 'Top_Communes'
    }[chartId] || 'chart';
    
    // Create a temporary link for download
    const link = document.createElement('a');
    link.download = `${chartName}_${new Date().toISOString().split('T')[0]}.png`;
    link.href = canvas.toDataURL('image/png');
    link.click();
    link.remove();
    
    showNotification(`Le graphique a été téléchargé avec succès.`, 'success');
}

/**
 * Sets up chart download buttons
 */
function setupChartDownloadButtons() {
    document.querySelectorAll('.chart-download').forEach(button => {
        button.addEventListener('click', function() {
            const chartId = this.dataset.chart;
            downloadChart(chartId);
        });
    });
    if (document.getElementById('analytics')?.classList.contains('active')) {
        initializeAnalyticsCharts();
    }
}

/**
 * Initializes analytics charts
 */
function initializeAnalyticsCharts() {
    initializeEfficiencyChart();
    initializeRejectionChart();
    initializeWorkloadChart();
    initializeJointureChart();
    initializeConflictHeatmap();
    initializeValidationTrendChart();
}

/**
 * Initializes status distribution chart
 */
function initializeStatusChart() {
    const ctx = document.getElementById('statusChart')?.getContext('2d');
    if (!ctx) return;
    
    const statusCounts = {
        completed: 0,
        processing: 0,
        pending: 0,
        rejected: 0
    };

    communesData.forEach(commune => {
        if (!commune.commune || commune.commune.toLowerCase() === 'total') return;
        const status = getStatus(commune);
        switch (status.text) {
            case 'Terminé': statusCounts.completed++; break;
            case 'En cours': statusCounts.processing++; break;
            case 'En attente': statusCounts.pending++; break;
            case 'Problèmes': statusCounts.rejected++; break;
        }
    });

    if (charts.status) charts.status.destroy();
    charts.status = new Chart(ctx, {
        type: 'pie',
        data: {
            labels: ['Terminé', 'En cours', 'En attente', 'Problèmes'],
            datasets: [{
                data: [statusCounts.completed, statusCounts.processing, statusCounts.pending, statusCounts.rejected],
                backgroundColor: ['#10b981', '#3b82f6', '#f59e0b', '#ef4444']
            }]
        },
        options: {
            responsive: true,
            maintainAspectRatio: false,
            plugins: {
                legend: { position: 'bottom' },
                tooltip: {
                    callbacks: {
                        label: context => `${context.label}: ${context.raw}`
                    }
                }
            }
        }
    });
}

/**
 * Initializes communes volume chart
 */
function initializeCommunesChart() {
    const ctx = document.getElementById('communesChart')?.getContext('2d');
    if (!ctx) return;
    
    const topCommunes = communesData
        .filter(c => c.commune && c.commune.toLowerCase() !== 'total')
        .sort((a, b) => b.collected - a.collected)
        .slice(0, 5);
    
    if (charts.communes) charts.communes.destroy();
    charts.communes = new Chart(ctx, {
        type: 'bar',
        data: {
            labels: topCommunes.map(c => c.commune),
            datasets: [{
                label: 'Parcelles Collectées',
                data: topCommunes.map(c => c.collected),
                backgroundColor: '#3498db'
            }]
        },
        options: {
            responsive: true,
            maintainAspectRatio: false,
            scales: {
                y: { beginAtZero: true },
                x: { ticks: { maxRotation: 45, minRotation: 0 } }
            },
            plugins: {
                tooltip: {
                    callbacks: {
                        label: context => `${context.label}: ${context.raw.toLocaleString()}`
                    }
                }
            }
        }
    });
}

/**
 * Initializes processing efficiency chart
 */
function initializeEfficiencyChart() {
    const ctx = document.getElementById('efficiencyChart')?.getContext('2d');
    if (!ctx) return;
    
    const filteredData = communesData.filter(c => 
        c.commune && 
        c.commune.toLowerCase() !== 'total' && 
        c.collected > 0
    ).sort((a, b) => b.collected - a.collected).slice(0, 12); // Take top 12 for better visibility
    
    console.debug('Efficiency chart data:', filteredData.map(c => ({
        commune: c.commune,
        collected: c.collected,
        validated: c.validated,
        validationRate: c.validated && c.collected ? (c.validated / c.collected) * 100 : 0
    })));
    
    const communeNames = filteredData.map(c => c.commune);
    const collectedData = filteredData.map(c => c.collected);
    const validatedData = filteredData.map(c => c.validated || 0);
    const validationRates = filteredData.map(c => 
        c.validated && c.collected ? (c.validated / c.collected) * 100 : 0
    );
    
    if (charts.efficiency) charts.efficiency.destroy();
    charts.efficiency = new Chart(ctx, {
        type: 'bar',
        data: {
            labels: communeNames,
            datasets: [
                {
                    label: 'Parcelles collectées',
                    data: collectedData,
                    backgroundColor: '#0072BC99', // With opacity
                    borderColor: '#0072BC',
                    borderWidth: 1,
                    order: 2,
                    barPercentage: 0.7,
                    categoryPercentage: 0.8
                },
                {
                    label: 'Parcelles validées',
                    data: validatedData,
                    backgroundColor: '#4CAF5099', // With opacity
                    borderColor: '#4CAF50',
                    borderWidth: 1,
                    order: 3,
                    barPercentage: 0.7,
                    categoryPercentage: 0.8
                },
                {
                    label: 'Taux de validation (%)',
                    data: validationRates,
                    type: 'line',
                    borderColor: '#A52A2A',
                    backgroundColor: '#A52A2A22', // Very light opacity
                    borderWidth: 2,
                    fill: false,
                    tension: 0.2,
                    yAxisID: 'y1',
                    order: 1,
                    pointRadius: 5,
                    pointHoverRadius: 7,
                    pointBackgroundColor: '#A52A2A'
                }
            ]
        },
        options: {
            responsive: true,
            maintainAspectRatio: false,
            scales: {
                x: {
                    ticks: { 
                        maxRotation: 45, 
                        minRotation: 30,
                        font: {
                            weight: 'bold',
                            size: 11
                        },
                        color: '#1C3D72'
                    },
                    grid: {
                        display: false
                    }
                },
                y: {
                    beginAtZero: true,
                    title: {
                        display: true,
                        text: 'Nombre de parcelles',
                        font: {
                            weight: 'bold'
                        }
                    },
                    grid: {
                        color: '#e2e8f0'
                    },
                    ticks: {
                        callback: value => value.toLocaleString()
                    }
                },
                y1: {
                    beginAtZero: true,
                    title: {
                        display: true,
                        text: 'Taux de validation (%)',
                        font: {
                            weight: 'bold'
                        },
                        color: '#A52A2A'
                    },
                    position: 'right',
                    max: 100,
                    grid: {
                        drawOnChartArea: false
                    },
                    ticks: {
                        color: '#A52A2A'
                    }
                }
            },
            plugins: {
                legend: {
                    position: 'top',
                    labels: {
                        usePointStyle: true,
                        boxWidth: 10,
                        font: {
                            size: 12
                        }
                    }
                },
                tooltip: {
                    backgroundColor: 'rgba(28, 61, 114, 0.9)',
                    titleFont: {
                        size: 14,
                        weight: 'bold'
                    },
                    bodyFont: {
                        size: 13
                    },
                    padding: 12,
                    cornerRadius: 8,
                    usePointStyle: true,
                    callbacks: {
                        label: context => {
                            const label = context.dataset.label;
                            const value = context.raw;
                            if (label === 'Taux de validation (%)') {
                                return `${label}: ${value.toFixed(1)}%`;
                            }
                            return `${label}: ${value.toLocaleString()}`;
                        }
                    }
                }
            }
        }
    });
}

/**
 * Initializes rejection reasons chart
 */
function initializeRejectionChart() {
    const ctx = document.getElementById('rejectionChart')?.getContext('2d');
    if (!ctx) return;
    
    const reasons = {};
    communesData.forEach(commune => {
        if (commune.rejectionReasons) {
            commune.rejectionReasons.split(',').forEach(reason => {
                reason = reason.trim();
                if (reason) reasons[reason] = (reasons[reason] || 0) + 1;
            });
        }
        if (commune.urmRejectionReasons) {
            commune.urmRejectionReasons.split(',').forEach(reason => {
                reason = reason.trim();
                if (reason) reasons[reason] = (reasons[reason] || 0) + 1;
            });
        }
    });
    
    if (charts.rejection) charts.rejection.destroy();
    charts.rejection = new Chart(ctx, {
        type: 'pie',
        data: {
            labels: Object.keys(reasons),
            datasets: [{
                data: Object.values(reasons),
                backgroundColor: ['#ef4444', '#f59e0b', '#d97706', '#b91c1c', '#7c2d12']
            }]
        },
        options: {
            responsive: true,
            maintainAspectRatio: false,
            plugins: {
                legend: { position: 'bottom' },
                tooltip: {
                    callbacks: {
                        label: context => `${context.label}: ${context.raw}`
                    }
                }
            }
        }
    });
}

/**
 * Initializes geomaticien workload chart
 */
function initializeWorkloadChart() {
    const ctx = document.getElementById('workloadChart')?.getContext('2d');
    if (!ctx) return;
    
    const workload = {};
    communesData.forEach(commune => {
        if (commune.geomaticien && !isNaN(Number(commune.validated))) {
            workload[commune.geomaticien] = (workload[commune.geomaticien] || 0) + commune.validated;
        }
    });
    
    if (charts.workload) charts.workload.destroy();
    charts.workload = new Chart(ctx, {
        type: 'bar',
        data: {
            labels: Object.keys(workload),
            datasets: [{
                label: 'Parcelles Validées',
                data: Object.values(workload),
                backgroundColor: '#3b82f6'
            }]
        },
        options: {
            responsive: true,
            maintainAspectRatio: false,
            scales: {
                y: { beginAtZero: true },
                x: { ticks: { maxRotation: 45, minRotation: 0 } }
            },
            plugins: {
                tooltip: {
                    callbacks: {
                        label: context => `${context.label}: ${context.raw.toLocaleString()}`
                    }
                }
            }
        }
    });
}

/**
 * Initializes jointure status chart
 */
function initializeJointureChart() {
    const ctx = document.getElementById('jointureChart')?.getContext('2d');
    if (!ctx) return;
    
    const statusCounts = {
        completed: 0,
        pending: 0,
        error: 0
    };
    
    communesData.forEach(commune => {
        if (commune.jointureStatus) {
            const status = commune.jointureStatus.toLowerCase();
            if (status.includes('complet')) statusCounts.completed++;
            else if (status.includes('en attente')) statusCounts.pending++;
            else if (status.includes('erreur')) statusCounts.error++;
        }
    });
    
    if (charts.jointure) charts.jointure.destroy();
    charts.jointure = new Chart(ctx, {
        type: 'pie',
        data: {
            labels: ['Complété', 'En attente', 'Erreur'],
            datasets: [{
                data: [statusCounts.completed, statusCounts.pending, statusCounts.error],
                backgroundColor: ['#10b981', '#f59e0b', '#ef4444']
            }]
        },
        options: {
            responsive: true,
            maintainAspectRatio: false,
            plugins: {
                legend: { position: 'bottom' },
                tooltip: {
                    callbacks: {
                        label: context => `${context.label}: ${context.raw}`
                    }
                }
            }
        }
    });
}

/**
 * Initializes conflict heatmap
 */
function initializeConflictHeatmap() {
    const ctx = document.getElementById('conflictHeatmap')?.getContext('2d');
    if (!ctx) return;
    
    const heatmapData = communesData
        .filter(c => c.commune && c.commune.toLowerCase() !== 'total')
        .map(commune => ({
            x: commune.commune,
            y: commune.parcelsInConflict || 0,
            v: commune.parcelsInConflict || 0
        }));
    
    if (charts.conflict) charts.conflict.destroy();
    charts.conflict = new Chart(ctx, {
        type: 'scatter',
        data: {
            datasets: [{
                label: 'Conflits',
                data: heatmapData,
                backgroundColor: 'rgba(239, 68, 68, 0.5)',
                pointRadius: 10
            }]
        },
        options: {
            responsive: true,
            maintainAspectRatio: false,
            scales: {
                x: {
                    title: { display: true, text: 'Commune' },
                    ticks: { maxRotation: 45, minRotation: 0 }
                },
                y: {
                    title: { display: true, text: 'Conflits' },
                    beginAtZero: true
                }
            },
            plugins: {
                tooltip: {
                    callbacks: {
                        label: context => `${context.raw.x}: ${context.raw.y.toLocaleString()} conflits`
                    }
                }
            }
        }
    });
}

/**
 * Initializes validation trend chart
 */
function initializeValidationTrendChart() {
    const ctx = document.getElementById('validationTrendChart')?.getContext('2d');
    if (!ctx) return;
    
    const trendData = communesData
        .filter(c => c.commune && c.commune.toLowerCase() !== 'total')
        .map(commune => ({
            commune: commune.commune,
            validated: commune.validated || 0
        }))
        .sort((a, b) => a.validated - b.validated);
    
    if (charts.validationTrend) charts.validationTrend.destroy();
    charts.validationTrend = new Chart(ctx, {
        type: 'line',
        data: {
            labels: trendData.map(d => d.commune),
            datasets: [{
                label: 'Parcelles Validées',
                data: trendData.map(d => d.validated),
                borderColor: '#10b981',
                fill: false
            }]
        },
        options: {
            responsive: true,
            maintainAspectRatio: false,
            scales: {
                y: { beginAtZero: true },
                x: { ticks: { maxRotation: 45, minRotation: 0 } }
            },
            plugins: {
                tooltip: {
                    callbacks: {
                        label: context => `${context.label}: ${context.raw.toLocaleString()}`
                    }
                }
            }
        }
    });
}

/**
 * Populates communes table
 */
function populateCommunesTable() {
    const tableBody = document.getElementById('communesTableBody');
    if (!tableBody) {
        console.warn('Communes table body not found');
        return;
    }
    
    tableBody.innerHTML = communesData.map(commune => `
        <tr>
            <td data-column="commune">${commune.commune || '-'}</td>
            <td data-column="region">${commune.region || '-'}</td>
            <td data-column="totalParcelles">${commune.totalParcelles?.toLocaleString() || '0'}</td>
            <td data-column="percentTotal">${commune.percentTotal?.toFixed(1) || '0'}%</td>
            <td data-column="nicad">${commune.nicad?.toLocaleString() || '0'}</td>
            <td data-column="percentNicad">${commune.percentNicad?.toFixed(1) || '0'}%</td>
            <td data-column="ctasf">${commune.ctasf?.toLocaleString() || '0'}</td>
            <td data-column="percentCtasf">${commune.percentCtasf?.toFixed(1) || '0'}%</td>
            <td data-column="deliberated">${commune.deliberated?.toLocaleString() || '0'}</td>
            <td data-column="percentDeliberated">${commune.percentDeliberated?.toFixed(1) || '0'}%</td>
            <td data-column="parcellesBrutes">${commune.parcellesBrutes?.toLocaleString() || '0'}</td>
            <td data-column="collected">${commune.collected?.toLocaleString() || '0'}</td>
            <td data-column="surveyed">${commune.surveyed?.toLocaleString() || '0'}</td>
            <td data-column="rejectionReasons">${commune.rejectionReasons || '-'}</td>
            <td data-column="retained">${commune.retained?.toLocaleString() || '0'}</td>
            <td data-column="validated">${commune.validated?.toLocaleString() || '0'}</td>
            <td data-column="rejected">${commune.rejected?.toLocaleString() || '0'}</td>
            <td data-column="urmRejectionReasons">${commune.urmRejectionReasons || '-'}</td>
            <td data-column="corrected">${commune.corrected?.toLocaleString() || '0'}</td>
            <td data-column="geomaticien">${commune.geomaticien || '-'}</td>
            <td data-column="individualJoined">${commune.individualJoined?.toLocaleString() || '0'}</td>
            <td data-column="collectiveJoined">${commune.collectiveJoined?.toLocaleString() || '0'}</td>
            <td data-column="unjoined">${commune.unjoined?.toLocaleString() || '0'}</td>
            <td data-column="duplicatesRemoved">${commune.duplicatesRemoved?.toLocaleString() || '0'}</td>
            <td data-column="duplicateRemovalRate">${commune.duplicateRemovalRate?.toFixed(1) || '0'}%</td>
            <td data-column="parcelsInConflict">${commune.parcelsInConflict?.toLocaleString() || '0'}</td>
            <td data-column="significantDuplicates">${commune.significantDuplicates?.toLocaleString() || '0'}</td>
            <td data-column="postProcessedLot1_46">${commune.postProcessedLot1_46?.toLocaleString() || '0'}</td>
            <td data-column="jointureStatus">${commune.jointureStatus || '-'}</td>
            <td data-column="jointureErrorMessage">${commune.jointureErrorMessage || '-'}</td>
        </tr>
    `).join('');

    updateTableColumns();
}

/**
 * Gets status for a commune
 * @param {CommuneData} commune
 * @returns {{ class: string, text: string }}
 */
function getStatus(commune) {
    if (!commune || commune.collected === 0) return { class: 'status-pending', text: 'En attente' };
    if (commune.validated && commune.collected) {
        const validationRate = (commune.validated / commune.collected) * 100;
    if (validationRate >= 80) return { class: 'status-completed', text: 'Terminé' };
        if (validationRate >= 50) return { class: 'status-processing', text: 'En cours' };
    }
    return { class: 'status-rejected', text: 'Problèmes' };
}

/**
 * Generates progress cards
 */
function generateProgressCards() {
    const progressCards = document.getElementById('progressCards');
    if (!progressCards) return;
    
    progressCards.innerHTML = communesData
        .filter(c => c.collected > 0 && c.commune && c.commune.toLowerCase() !== 'total')
        .map(commune => {
            const validationRate = commune.validated && commune.collected ? ((commune.validated / commune.collected) * 100).toFixed(1) : 0;
            const status = getStatus(commune);
            const geomaticienName = commune.geomaticien && !isNaN(Number(commune.geomaticien)) ? '-' : (commune.geomaticien || '-');
            
            // Calculate additional metrics for display
            const retainedRate = commune.retained && commune.collected ? ((commune.retained / commune.collected) * 100).toFixed(1) : 0;
            const duplicateRateDisplay = commune.duplicateRemovalRate?.toFixed(1) || '0';
            const conflictCount = commune.parcelsInConflict || 0;
            
            // Determine alert status for special issues
            const hasDuplicateIssue = commune.duplicateRemovalRate > 40;
            const hasConflictIssue = conflictCount > 20;
            const hasJointureIssue = commune.unjoined > 1000;
            
            return `
                <div class="progress-card">
                    <div class="progress-card-header">
                        <h4 class="commune-name">${commune.commune}</h4>
                        <span class="status-badge ${status.class}">${status.text}</span>
                    </div>
                    <div class="progress-details">
                        <div class="detail-row">
                            <div>
                                <span class="detail-label">Collectées</span>
                                <span class="detail-value">${commune.collected.toLocaleString()}</span>
                            </div>
                            <div>
                                <span class="detail-label">Validées</span>
                                <span class="detail-value">${commune.validated?.toLocaleString() || '0'}</span>
                            </div>
                            <div>
                                <span class="detail-label">Non jointes</span>
                                <span class="detail-value ${hasJointureIssue ? 'text-danger' : ''}">${commune.unjoined?.toLocaleString() || '0'}</span>
                            </div>
                        </div>
                        
                        <div class="detail-row mt-2">
                            <div>
                                <span class="detail-label">Géomaticien</span>
                                <span class="detail-value geomaticien-badge">${geomaticienName}</span>
                            </div>
                            <div>
                                <span class="detail-label">Doublons</span>
                                <span class="detail-value ${hasDuplicateIssue ? 'text-danger' : ''}">${duplicateRateDisplay}%</span>
                            </div>
                            <div>
                                <span class="detail-label">Conflits</span>
                                <span class="detail-value ${hasConflictIssue ? 'text-danger' : ''}">${conflictCount}</span>
                            </div>
                        </div>
                        
                        <div class="progress-section mt-3">
                            <div class="progress-header">
                                <span>Taux de validation</span>
                                <span class="progress-percentage" style="color: ${getProgressBarColor(validationRate)};">${validationRate}%</span>
                            </div>
                            <div class="progress-bar">
                                <div class="progress-fill" style="width: ${Math.min(validationRate, 100)}%; background-color: ${getProgressBarColor(validationRate)};"></div>
                            </div>
                        </div>
                        
                        ${hasDuplicateIssue || hasConflictIssue || hasJointureIssue ? `
                            <div class="alert-section mt-2">
                                ${hasDuplicateIssue ? `<div class="alert-item"><i class="fas fa-exclamation-triangle"></i> Taux élevé de doublons (${duplicateRateDisplay}%)</div>` : ''}
                                ${hasConflictIssue ? `<div class="alert-item"><i class="fas fa-exclamation-triangle"></i> Nombreux conflits (${conflictCount})</div>` : ''}
                                ${hasJointureIssue ? `<div class="alert-item"><i class="fas fa-exclamation-triangle"></i> Problèmes de jointure (${commune.unjoined})</div>` : ''}
                            </div>
                        ` : ''}
                    </div>
                </div>
            `;
        }).join('');
}

/**
 * Generates activity list
 */
function generateActivityList() {
    const activityList = document.getElementById('communesList');
    if (!activityList) return;
    
    const sortedCommunes = communesData
        .filter(c => c.collected > 0 && c.commune && c.commune.toLowerCase() !== 'total')
        .sort((a, b) => (b.validated / b.collected || 0) - (a.validated / a.collected || 0))
        .slice(0, 8);
    
    activityList.innerHTML = sortedCommunes.map(commune => {
        const validationRate = commune.validated && commune.collected ? ((commune.validated / commune.collected) * 100).toFixed(1) : 0;
        return `
            <div class="activity-item">
                <div>
                    <strong>${commune.commune}</strong>
                    <p class="text-gray-600 text-sm mt-1">${commune.validated?.toLocaleString() || '0'} of ${commune.collected.toLocaleString()} validated</p>
                </div>
                <span class="progress-percentage">${validationRate}%</span>
            </div>
        `;
    }).join('');
}

/**
 * Filters communes table
 * @param {string} searchTerm
 */
function filterCommunesTable(searchTerm) {
    const tableBody = document.getElementById('communesTableBody');
    if (!tableBody) return;

    const rows = tableBody.querySelectorAll('tr');
    rows.forEach(row => {
        const communeName = row.querySelector('td[data-column="commune"]')?.textContent.toLowerCase() || '';
        const geomaticien = row.querySelector('td[data-column="geomaticien"]')?.textContent.toLowerCase() || '';
        row.style.display = (communeName.includes(searchTerm) || geomaticien.includes(searchTerm)) && columnVisibility.commune ? '' : 'none';
    });

    updateTableColumns();
}

/**
 * Triggers manual refresh
 */
function manualRefresh() {
    fetchCommunesData(true);
}

/**
 * Shows loading indicator
 */
function showLoadingIndicator() {
    if (!document.getElementById('loadingIndicator')) {
        const indicator = document.createElement('div');
        indicator.id = 'loadingIndicator';
        indicator.className = 'notification';
        indicator.innerHTML = `
            <div class="notification-content">
                <div class="loading-spinner"></div>
                <span>Refreshing data...</span>
            </div>
        `;
        document.body.appendChild(indicator);
    }
}

/**
 * Hides loading indicator
 */
function hideLoadingIndicator() {
    document.getElementById('loadingIndicator')?.remove();
}

/**
 * Shows notification
 * @param {string} message
 * @param {'success' | 'error'} type
 */
function showNotification(message, type = 'success') {
    const notification = document.createElement('div');
    notification.className = `notification notification-${type}`;
    notification.innerHTML = `
        <div class="notification-content">
            <i class="fas fa-${type === 'success' ? 'check-circle' : 'exclamation-circle'}"></i>
            <span>${message}</span>
        </div>
        <button class="notification-close"><i class="fas fa-times"></i></button>
    `;
    document.body.appendChild(notification);
    notification.querySelector('.notification-close').addEventListener('click', () => notification.remove());
    setTimeout(() => notification.classList.add('opacity-0'), 4000);
    setTimeout(() => notification.remove(), 4500);
}

/**
 * Updates last refresh time
 */
function updateLastRefreshTime() {
    const status = document.getElementById('autoRefreshStatus');
    if (status) {
        const time = new Date().toLocaleTimeString('en-US', { hour: '2-digit', minute: '2-digit', timeZone: 'GMT' });
        status.querySelector('span').textContent = `Last refresh: ${time} | Auto-refresh: ON`;
        status.classList.add('active');
    }
}

/**
 * Exports data to specified format
 * @param {'csv'} format
 */
function exportData(format = 'csv') {
    const data = communesData.map(commune => ({
        Commune: commune.commune,
    Région: commune.region,
        'Total Parcelles': commune.totalParcelles,
        '% du Total': commune.percentTotal,
        NICAD: commune.nicad,
        '% NICAD': commune.percentNicad,
        CTASF: commune.ctasf,
        '% CTASF': commune.percentCtasf,
    Délibérées: commune.deliberated,
    '% Délibérée': commune.percentDeliberated,
        'Parcelles brutes': commune.parcellesBrutes,
    'Parcelles collectées (sans doublon géométrique)': commune.collected,
    'Parcelles enquêtées': commune.surveyed,
        'Motifs de rejet post-traitement': commune.rejectionReasons,
    'Parcelles retenues après post-traitement': commune.retained,
    'Parcelles validées par l\'URM': commune.validated,
    'Parcelles rejetées par l\'URM': commune.rejected,
        'Motifs de rejet URM': commune.urmRejectionReasons,
    'Parcelles corrigées': commune.corrected,
        Geomaticien: commune.geomaticien,
        'Parcelles individuelles jointes': commune.individualJoined,
        'Parcelles collectives jointes': commune.collectiveJoined,
        'Parcelles non jointes': commune.unjoined,
    'Doublons supprimés': commune.duplicatesRemoved,
        'Taux suppression doublons (%)': commune.duplicateRemovalRate,
        'Parcelles en conflit': commune.parcelsInConflict,
        'Significant Duplicates': commune.significantDuplicates,
    'Parc. post-traitées lot 1-46': commune.postProcessedLot1_46,
    'Statut jointure': commune.jointureStatus,
    'Message d\'erreur jointure': commune.jointureErrorMessage
    }));

    if (format === 'csv') {
        exportToCSV(data);
    }
}

/**
 * Exports data to CSV
 * @param {Object[]} data
 */
function exportToCSV(data) {
    if (!data.length) {
        showNotification('No data to export', 'error');
        return;
    }
    
    const headers = Object.keys(data[0]).map(header => `"${header.replace(/"/g, '""')}"`).join(',');
    const rows = data.map(row => Object.values(row).map(value => `"${String(value).replace(/"/g, '""')}"`).join(','));
    const csv = [headers, ...rows].join('\n');
    const blob = new Blob([csv], { type: 'text/csv;charset=utf-8' });
    const url = window.URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = `communes_data_${new Date().toISOString().split('T')[0]}.csv`;
    a.click();
    window.URL.revokeObjectURL(url);
}

/**
 * Ensures columns are exported in the correct order based on current visible columns and their order
 */
function exportWithColumnOrder() {
    // Column name mapping for display
    const columnDisplayNames = {
        'commune': 'Commune',
    'region': 'Région',
        'totalParcelles': 'Total Parcelles',
        'percentTotal': '% du Total',
        'nicad': 'NICAD',
        'percentNicad': '% NICAD',
        'ctasf': 'CTASF',
        'percentCtasf': '% CTASF',
    'deliberated': 'Délibérées',
    'percentDeliberated': '% Délibérée',
        'parcellesBrutes': 'Parcelles brutes',
    'collected': 'Parcelles collectées (sans doublon géométrique)',
    'surveyed': 'Parcelles enquêtées',
        'rejectionReasons': 'Motifs de rejet post-traitement',
    'retained': 'Parcelles retenues après post-traitement',
    'validated': 'Parcelles validées par l\'URM',
    'rejected': 'Parcelles rejetées par l\'URM',
        'urmRejectionReasons': 'Motifs de rejet URM',
    'corrected': 'Parcelles corrigées',
        'geomaticien': 'Geomaticien',
        'individualJoined': 'Parcelles individuelles jointes',
        'collectiveJoined': 'Parcelles collectives jointes',
        'unjoined': 'Parcelles non jointes',
    'duplicatesRemoved': 'Doublons supprimés',
        'duplicateRemovalRate': 'Taux suppression doublons (%)',
        'parcelsInConflict': 'Parcelles en conflit',
        'significantDuplicates': 'Significant Duplicates',
    'postProcessedLot1_46': 'Parc. post-traitées lot 1-46',
        'jointureStatus': 'Statut jointure',
        'jointureErrorMessage': 'Message d\'erreur jointure'
    };
    
    // Get only visible columns in their current order
    const visibleColumns = currentColumnOrder.filter(col => columnVisibility[col]);
    
    // Map data according to visible columns and their order
    const data = communesData.map(commune => {
        const row = {};
        visibleColumns.forEach(col => {
            row[columnDisplayNames[col] || col] = commune[col];
        });
        return row;
    });

    exportToCSV(data);
}

/**
 * Generates report
 * @param {string} reportType
 */
function generateReport(reportType) {
    showNotification(`Generating ${reportType}...`, 'success');
}
