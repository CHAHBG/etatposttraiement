/**
 * PROCASEF Boundou Operations Dashboard
 * Comprehensive JavaScript implementation for multi-sheet data integration
 * @author PROCASEF Development Team
 * @version 1.0.0
 * @dependencies Requires Chart.js for charts and SheetJS (XLSX) for Excel export
 */

// ============================
// GLOBAL CONFIGURATION
// ============================

const CONFIG = {
    // Google Sheets Configuration
    SHEETS: {
        BASE_URL: 'https://docs.google.com/spreadsheets/d/1CbDBJtoWWPRjEH4DOSlv4jjnJ2G-1MvW/edit',
        EXPORT_BASE: 'https://docs.google.com/spreadsheets/d/1CbDBJtoWWPRjEH4DOSlv4jjnJ2G-1MvW/export',
        KPI_DASHBOARD: { gid: '589154853', name: 'KPI Dashboard' },
        DATA_COLLECTION: { gid: '778550489', name: 'Data Collection' },
        PROJECTION_COLLECTION: { gid: '1687610293', name: 'Projection Collection' },
        YIELDS_PROJECTION: { gid: '1397416280', name: 'Yields Projection' },
        TEAM_DEPLOYMENT: { gid: '660642704', name: 'Team Deployment' },
        NICAD_PROJECTION: { gid: '1151168155', name: 'NICAD Projection' },
        PUBLIC_DISPLAY: { gid: '2035205606', name: 'Public Display' },
        CTASF_PROJECTION: { gid: '1574818070', name: 'CTASF projection' },
        COMMUNE_STATUS: { gid: '1421590976', name: 'Commune Status' },
        METHODOLOGY_NOTES: { gid: '1203072335', name: 'Methodology Notes' }
    },
    
    // UI Configuration
    UI: {
        REFRESH_INTERVAL: 300000, // 5 minutes
        CHART_ANIMATION_DURATION: 1000,
        NOTIFICATION_DURATION: 5000,
        LOADING_MIN_DURATION: 1500,
    FONT_FAMILY: 'Arial, sans-serif', // Added to resolve undefined FONT_FAMILY
    debugFieldResolution: false // Set to true to enable field-resolution logging
    },
    
    // Chart Colors (PROCASEF Branding)
    COLORS: {
        primary: '#1e40af',
        secondary: '#0072BC',
        success: '#059669',
        warning: '#d97706',
        error: '#dc2626',
        info: '#0ea5e9',
        agricultural: '#4CAF50',
        earth: '#6B4226',
        highlight: '#FFD700',
        gradients: {
            primary: ['#1e40af', '#3b82f6'],
            success: ['#059669', '#10b981'],
            warning: ['#d97706', '#f59e0b'],
            earth: ['#6B4226', '#8B5A3C']
        }
    }
};

// ============================
// UTILITY FUNCTIONS
// ============================

/**
 * Sanitizes HTML input to prevent XSS
 * @param {string} input - Input string
 * @returns {string} Sanitized string
 */
function sanitizeHTML(input) {
    const div = document.createElement('div');
    div.textContent = input;
    return div.innerHTML;
}

/**
 * Shows a loading indicator
 * @param {string} message - Loading message
 */
function showLoadingIndicator(message) {
    const loadingElement = document.getElementById('loadingIndicator');
    if (loadingElement) {
        loadingElement.textContent = sanitizeHTML(message);
        loadingElement.classList.remove('hidden');
    }
}

/**
 * Hides the loading indicator
 */
function hideLoadingIndicator() {
    const loadingElement = document.getElementById('loadingIndicator');
    if (loadingElement) {
        loadingElement.classList.add('hidden');
    }
}

/**
 * Shows a notification
 * @param {string} message - Notification message
 * @param {string} type - Notification type (success, error, warning, info)
 */
function showNotification(message, type) {
    if (!dashboardState.settings.alerts) return;
    
    const notificationContainer = document.getElementById('notificationContainer');
    if (!notificationContainer) return;
    
    const notification = document.createElement('div');
    notification.className = `notification ${type}`;
    notification.innerHTML = sanitizeHTML(message);
    
    notificationContainer.appendChild(notification);
    
    setTimeout(() => {
        notification.classList.add('fade-out');
        setTimeout(() => notification.remove(), 500);
    }, CONFIG.UI.NOTIFICATION_DURATION);
    
    if (dashboardState.settings.sound) {
        // Play notification sound (assuming an audio element exists)
        const sound = document.getElementById('notificationSound');
        if (sound) sound.play();
    }
}

/**
 * Creates a modal
 * @param {string} title - Modal title
 * @param {string} content - Modal content
 * @returns {HTMLElement} Modal element
 */
function createModal(title, content) {
    const modal = document.createElement('div');
    modal.className = 'modal';
    modal.innerHTML = `
        <div class="modal-content">
            <div class="modal-header">
                <h3>${sanitizeHTML(title)}</h3>
                <button class="modal-close">&times;</button>
            </div>
            <div class="modal-body">
                ${content}
            </div>
        </div>
    `;
    return modal;
}

/**
 * Shows a modal
 * @param {HTMLElement} modal - Modal element
 */
function showModal(modal) {
    document.body.appendChild(modal);
    modal.querySelector('.modal-close').addEventListener('click', () => {
        modal.classList.add('fade-out');
        setTimeout(() => modal.remove(), 500);
    });
}

// ============================
// GLOBAL STATE MANAGEMENT
// ============================

class DashboardState {
    constructor() {
        this.data = new Map();
        this.charts = new Map();
        this.activeTab = 'executive';
        this.settings = this.loadSettings();
        this.listeners = new Map();
        this.lastUpdate = null;
        this.refreshTimer = null;
    }
    
    loadSettings() {
        const defaultSettings = {
            autoRefresh: true,
            refreshInterval: 5,
            darkMode: false,
            animations: true,
            alerts: true,
            sound: false
        };
        
        try {
            const saved = localStorage.getItem('procasef_dashboard_settings');
            return saved ? { ...defaultSettings, ...JSON.parse(saved) } : defaultSettings;
        } catch (error) {
            console.warn('Failed to load settings:', error);
            return defaultSettings;
        }
    }
    
    saveSettings() {
        try {
            localStorage.setItem('procasef_dashboard_settings', JSON.stringify(this.settings));
        } catch (error) {
            console.error('Failed to save settings:', error);
        }
    }
    
    updateSetting(key, value) {
        this.settings[key] = value;
        this.saveSettings();
        this.emit('settingsChanged', { key, value });
    }
    
    on(event, callback) {
        if (!this.listeners.has(event)) {
            this.listeners.set(event, []);
        }
        this.listeners.get(event).push(callback);
    }
    
    emit(event, data) {
        const callbacks = this.listeners.get(event) || [];
        callbacks.forEach(callback => callback(data));
    }
    
    setData(key, value) {
        this.data.set(key, value);
        this.emit('dataChanged', { key, value });
    }
    
    getData(key) {
        return this.data.get(key);
    }
    
    setChart(key, chart) {
        // Destroy existing chart
        const existing = this.charts.get(key);
        if (existing && typeof existing.destroy === 'function') {
            existing.destroy();
        }
        this.charts.set(key, chart);
    }
    
    getChart(key) {
        return this.charts.get(key);
    }
    
    destroyAllCharts() {
        this.charts.forEach(chart => {
            if (chart && typeof chart.destroy === 'function') {
                chart.destroy();
            }
        });
        this.charts.clear();
    }

    /**
     * Ensures a canvas is free by destroying any chart using it (by canvas id) or a specific key.
     * @param {HTMLCanvasElement|string} canvasOrId - Canvas element or canvas id
     * @param {string} [key] - Optional chart key to also clear
     */
    ensureCanvasFree(canvasOrId, key) {
        try {
            let canvasId = typeof canvasOrId === 'string' ? canvasOrId : (canvasOrId && canvasOrId.id ? canvasOrId.id : null);
            // Destroy chart by provided key first
            if (key) {
                const existing = this.charts.get(key);
                if (existing && typeof existing.destroy === 'function') {
                    try { existing.destroy(); } catch (e) { /* ignore */ }
                    try { this.charts.delete(key); } catch (e) { /* ignore */ }
                }
            }

            // Then iterate charts and destroy any with matching canvas id
            if (canvasId) {
                this.charts.forEach((ch, k) => {
                    try {
                        if (ch && ch.canvas && ch.canvas.id === canvasId) {
                            if (typeof ch.destroy === 'function') {
                                try { ch.destroy(); } catch (e) { /* ignore */ }
                            }
                            try { this.charts.delete(k); } catch (e) { /* ignore */ }
                        }
                    } catch (e) { /* ignore per-chart errors */ }
                });
            }
        } catch (e) {
            // Silently ignore to avoid breaking caller
        }
    }
}

// Global state instance
const dashboardState = new DashboardState();

// ============================
// DATA FETCHING & PROCESSING
// ============================

class DataManager {
    constructor() {
        this.cache = new Map();
        this.cacheTimeout = 5 * 60 * 1000; // 5 minutes
    }
    
    /**
     * Constructs CSV URL for a specific Google Sheet
     * @param {string} gid - Sheet GID
     * @returns {string} CSV URL
     */
    getSheetURL(gid) {
        return `${CONFIG.SHEETS.EXPORT_BASE}?format=csv&gid=${gid}`;
    }
    
    /**
     * Fetches data from Google Sheets with caching
     * @param {string} sheetKey - Sheet configuration key
     * @param {boolean} forceRefresh - Force refresh cache
     * @returns {Promise<Array>} Parsed data
     */
    async fetchSheetData(sheetKey, forceRefresh = false) {
        const sheet = CONFIG.SHEETS[sheetKey];
        if (!sheet) {
            throw new Error(`Unknown sheet: ${sheetKey}`);
        }
        
        const cacheKey = `sheet_${sheetKey}`;
        const cached = this.cache.get(cacheKey);
        
        // Return cached data if valid and not forcing refresh
        if (!forceRefresh && cached && (Date.now() - cached.timestamp) < this.cacheTimeout) {
            return cached.data;
        }
        
        try {
            showLoadingIndicator(`Loading ${sheet.name}...`);
            
            const url = this.getSheetURL(sheet.gid);
            const response = await fetch(url, {
                headers: {
                    'Accept': 'text/csv',
                    'Cache-Control': forceRefresh ? 'no-cache' : 'default'
                }
            });
            
            if (!response.ok) {
                throw new Error(`HTTP ${response.status}: Failed to fetch ${sheet.name}`);
            }
            
            const csvText = await response.text();
            const data = this.parseCSV(csvText);
            // Optional debug: log parsed rows for troubleshooting
            try {
                if (typeof dashboardState !== 'undefined' && dashboardState.settings && dashboardState.settings.debug) {
                    console.debug(`Fetched sheet ${sheetKey} (${sheet.name}), first rows:`, data.slice(0, 5));
                }
            } catch (e) {
                // ignore logging errors
            }
            
            // Cache the data
            this.cache.set(cacheKey, {
                data,
                timestamp: Date.now()
            });
            
            hideLoadingIndicator();
            return data;
            
        } catch (error) {
            hideLoadingIndicator();
            console.error(`Error fetching ${sheet.name}:`, error);
            showNotification(`Failed to load ${sheet.name}: ${error.message}`, 'error');
            
            // Return cached data if available, otherwise empty array
            return cached ? cached.data : [];
        }
    }
    
    /**
     * Parses CSV text into structured data
     * @param {string} csvText - Raw CSV text
     * @returns {Array<Object>} Parsed data objects
     */
    parseCSV(csvText) {
        if (!csvText || !csvText.trim()) {
            return [];
        }
        // Determine delimiter robustly: try a set of candidate delimiters and pick the one
        // that yields the most rows with the same number of columns as the header.
        const lines = csvText.split(/\r?\n/).filter(line => line.trim());
        const candidates = [';', '\t', ','];
        const scoreFor = (d) => {
            try {
                const hdr = this.parseLine(lines[0], d);
                if (!hdr || hdr.length <= 1) return 0;
                let ok = 0;
                const sampleN = Math.min(50, Math.max(5, lines.length - 1));
                for (let i = 1; i <= sampleN && i < lines.length; i++) {
                    const vals = this.parseLine(lines[i], d);
                    if (vals.length === hdr.length) ok++;
                }
                return ok;
            } catch (e) { return 0; }
        };
        let best = {d: ',', score: -1};
        for (const d of candidates) {
            const s = scoreFor(d);
            if (s > best.score) best = {d, score: s};
        }
        // Fallback to original simple heuristic if no candidate stands out
        let delimiter = best.d;
        if (best.score <= 0) {
            const sampleLines = lines.slice(0, 5);
            let commaCount = 0, semicolonCount = 0;
            sampleLines.forEach(l => { commaCount += (l.match(/,/g) || []).length; semicolonCount += (l.match(/;/g) || []).length; });
            delimiter = semicolonCount > commaCount ? ';' : ',';
        }
        if (lines.length < 2) {
            return [];
        }
        
        // Parse header row
        const headers = this.parseLine(lines[0], delimiter);
        const data = [];
        
        // Parse data rows
        for (let i = 1; i < lines.length; i++) {
            let values = this.parseLine(lines[i], delimiter);
            // If row has more fields than headers (e.g., unescaped delimiter inside free-text),
            // merge extras into the last header so columns remain aligned.
            if (values.length > headers.length) {
                const merged = values.slice(0, headers.length - 1);
                merged.push(values.slice(headers.length - 1).join(delimiter));
                values = merged;
            }
            // If row has fewer fields than headers, pad with empty strings
            if (values.length < headers.length) {
                values = values.concat(Array(headers.length - values.length).fill(''));
            }

            const row = {};
            headers.forEach((header, index) => {
                let value = values[index] !== undefined ? values[index] : '';

                // Remove BOM from header if present and normalize
                const originalHeader = (header || '').replace(/^\uFEFF/, '');
                let cleanHeader = originalHeader.trim();
                try {
                    cleanHeader = cleanHeader.normalize('NFD').replace(/[\u0300-\u036f]/g, '');
                } catch (e) {
                    // normalize may not be available in some environments; ignore if so
                }
                cleanHeader = cleanHeader.replace(/[^\w\s]/g, '').replace(/\s+/g, '_').toLowerCase();

                // Map cleaned header variants to canonical keys used across the dashboard
                const headerAliasMap = {
                    // common French/English variants -> canonical keys
                    'commune': 'commune',
                    'region': 'region',
                    'total_parcelles': 'total_parcels',
                    'total_parcelles_collections': 'total_parcels',
                    'parcelles_totales': 'total_parcels',
                    'parcelles_collectees_sans_doublon_geometrique': 'collected',
                    'parcelles_collectees': 'collected',
                    'collected': 'collected',
                    'parcelles_enquetees': 'surveyed',
                    'parcelles_enquetees': 'surveyed',
                    'parcelles_brutes': 'parcelles_brutes',
                    // New French header variants
                    'région': 'region',
                    'region': 'region',
                    'total_parcelles': 'total_parcels',
                    'total_parcelles_': 'total_parcels',
                    'pourcent_du_total': 'percentTotal',
                    'pourcent_total': 'percentTotal',
                    'percent_total': 'percentTotal',
                    'nicad': 'nicad',
                    'pourcent_nicad': 'percentNicad',
                    'percent_nicad': 'percentNicad',
                    'ctasf': 'ctasf',
                    'pourcent_ctasf': 'percentCtasf',
                    'percent_ctasf': 'percentCtasf',
                    'deliberees': 'deliberated',
                    'pourcent_deliberee': 'percentDeliberated',
                    'parcelles_collectees_sans_doublon_geometrique': 'collected',
                    'parcelles_collectees': 'collected',
                    'parcelles_enquetees': 'surveyed',
                    'motifs_de_rejet_post_traitement': 'rejectionReasons',
                    'parcelles_retenues_apres_post_traitement': 'retained',
                    'parcelles_retenues_apres_posttraitement': 'retained',
                    'parcelles_validees_par_lurm': 'validated',
                    'parcelles_rejetees_par_lurm': 'rejected',
                    'motifs_de_rejet_urm': 'urmRejectionReasons',
                    'parcelles_corrigees': 'corrected',
                    'geomaticien': 'geomaticien',
                    'parcelles_individuelles_jointes': 'individualJoined',
                    'parcelles_collectives_jointes': 'collectiveJoined',
                    'parcelles_non_jointes': 'unjoined',
                    'doublons_supprimes': 'duplicatesRemoved',
                    'taux_suppression_doublons': 'duplicateRemovalRate',
                    'parcelles_en_conflit': 'parcelsInConflict',
                    'significant_duplicates': 'significantDuplicates',
                    'parcelles_post_traitees_lot_1_46': 'post_processed_parcels',
                    'statut_jointure': 'joinStatus',
                    'message_derreur_jointure': 'joinErrorMessage',
                    'nicad': 'nicad',
                    'ctasf': 'ctasf',
                    'deliberees': 'deliberated',
                    'deliberees': 'deliberated',
                    'parcelles_retenues_apres_posttraitement': 'retained',
                    'parcelles_retenues': 'retained',
                    'parcelles_validees_par_lurm': 'validated',
                    'parcelles_validees': 'validated',
                    'parcelles_rejetees_par_lurm': 'rejected',
                    'parcelles_rejetees': 'rejected',
                    'doublons_supprimes': 'duplicates_removed',
                    'taux_suppression_doublons': 'duplicateRemovalRate',
                    'taux_suppression_doublons_': 'duplicateRemovalRate',
                    'parcelles_en_conflit': 'parcelsInConflict',
                    'parcelles_corrigees': 'corrected',
                    'geomaticien': 'geomaticien',
                    'geomatician': 'geomaticien',
                    'champs_equipe_jour': 'champs_equipe_jour',
                    'champs_equipe_jour_': 'champs_equipe_jour',
                    'percent_total': 'percentTotal',
                    'percent_nicad': 'percentNicad',
                    'percent_ctasf': 'percentCtasf',
                    'percent_deliberated': 'percentDeliberated',
                    'duplicate_removal_rate': 'duplicateRemovalRate',
                    'urm_validated_parcels': 'validated',
                    'urm_rejected_parcels': 'rejected',
                    'collected_parcels_no_duplicates': 'collected'
                };

                // More explicit mappings for common English header names provided by the user
                const extraAliases = {
                    'total_parcels': 'total_parcels',
                    'total_': 'percentTotal',
                    'nicad_': 'percentNicad',
                    'ctasf_': 'percentCtasf',
                    'deliberated_': 'percentDeliberated',
                    'raw_parcels': 'parcelles_brutes',
                    'collected_parcels_no_duplicates': 'collected',
                    'surveyed_parcels': 'surveyed',
                    'postprocessing_rejection_reasons': 'rejectionReasons',
                    'post_processing_rejection_reasons': 'rejectionReasons',
                    'post_processed_parcels': 'retained',
                    'urm_validated_parcels': 'validated',
                    'urm_rejected_parcels': 'rejected',
                    'urm_rejection_reasons': 'urmRejectionReasons',
                    'corrected_parcels': 'corrected',
                    'individual_parcels_joined': 'individualJoined',
                    'collective_parcels_joined': 'collectiveJoined',
                    'non_joined_parcels': 'unjoined',
                    'duplicates_removed': 'duplicatesRemoved',
                    'duplicate_removal_rate_': 'duplicateRemovalRate',
                    'parcels_in_conflict': 'parcelsInConflict',
                    'significant_duplicates': 'significantDuplicates'
                };

                // Merge extraAliases into headerAliasMap without overwriting existing
                Object.keys(extraAliases).forEach(k => { if (!headerAliasMap[k]) headerAliasMap[k] = extraAliases[k]; });

                const mappedKey = headerAliasMap[cleanHeader] || cleanHeader;

                // Try to convert numeric values (handle French decimals and percentages)
                if (typeof value === 'string') {
                    value = value.trim();

                    // Percentage with dot or comma decimals e.g. 12.3% or 12,3%
                    const pct = value.match(/^(-?[\d\s\.,]+)\s*%$/);
                    if (pct) {
                        const num = parseFloat(pct[1].replace(/\s/g, '').replace(/,/, '.'));
                        if (!isNaN(num)) value = num;
                    } else {
                        // French style decimal: 1 234,56 or 1234,56
                        if (/^-?[\d\s]+,\d+$/.test(value)) {
                            const num = parseFloat(value.replace(/\s/g, '').replace(',', '.'));
                            if (!isNaN(num)) value = num;
                        }
                        // Standard dot decimal
                        else if (/^-?\d+(\.\d+)?$/.test(value)) {
                            value = parseFloat(value);
                        }
                        // Thousands with commas (only if delimiter isn't comma, but attempt to sanitize)
                        else if (/^-?\d{1,3}(,\d{3})*(\.\d+)?$/.test(value)) {
                            const num = parseFloat(value.replace(/,/g, ''));
                            if (!isNaN(num)) value = num;
                        }
                    }
                }

                // Store only the canonical mapped key to avoid duplicate/english header columns
                row[mappedKey] = value;
            });

            // Ensure we populate a canonical 'geomaticien' key when a name exists elsewhere
            try {
                const extracted = this.getGeomaticienName(row);
                const existing = row.geomaticien || row.geomaticien || row.Geomaticien || '';
                // If extracted name is found and existing value is missing or clearly not a name, set canonical key
                if (extracted && (!existing || !/[A-Za-zÀ-ÖØ-öø-ÿ]/.test(String(existing)))) {
                    row.geomaticien = extracted;
                }
            } catch (e) { /* ignore */ }

            data.push(row);
        }
        
        return data;
    }
    
    /**
     * Parses a single CSV line handling quoted values
     * @param {string} line - CSV line
     * @param {string} delimiter - Field delimiter
     * @returns {Array<string>} Parsed fields
     */
    parseLine(line, delimiter) {
        const result = [];
        let current = '';
        let inQuotes = false;
        for (let i = 0; i < line.length; i++) {
            const char = line[i];

            if (char === '"') {
                // Handle escaped quotes
                if (inQuotes && line[i + 1] === '"') {
                    current += '"';
                    i++; // Skip next quote
                } else {
                    inQuotes = !inQuotes;
                }
            } else if (char === delimiter && !inQuotes) {
                result.push(current.trim());
                current = '';
            } else {
                current += char;
            }
        }

        // Push last field (trim only trailing/leading whitespace)
        result.push(current.trim());
        return result;
    }
    
    /**
     * Fetches all dashboard data
     * @param {boolean} forceRefresh - Force refresh all data
     * @returns {Promise<Object>} All dashboard data
     */
    async fetchAllData(forceRefresh = false) {
        const startTime = Date.now();
        
        try {
            // Fetch all sheets in parallel
            const promises = Object.keys(CONFIG.SHEETS)
                .filter(key => key !== 'BASE_URL' && key !== 'EXPORT_BASE')
                .map(async (key) => {
                    const data = await this.fetchSheetData(key, forceRefresh);
                    return { key: key.toLowerCase(), data };
                });
            
            const results = await Promise.all(promises);
            
            // Store data in global state
            const allData = {};
            results.forEach(({ key, data }) => {
                allData[key] = data;
                dashboardState.setData(key, data);
            });
            
            dashboardState.lastUpdate = new Date();
            
            // Ensure minimum loading time for better UX
            const elapsedTime = Date.now() - startTime;
            if (elapsedTime < CONFIG.UI.LOADING_MIN_DURATION) {
                await new Promise(resolve => 
                    setTimeout(resolve, CONFIG.UI.LOADING_MIN_DURATION - elapsedTime)
                );
            }
            
            return allData;
            
        } catch (error) {
            console.error('Error fetching all data:', error);
            showNotification('Failed to load dashboard data', 'error');
            throw error;
        }
    }
    
    /**
     * Clears all cached data
     */
    clearCache() {
        this.cache.clear();
        showNotification('Cache cleared successfully', 'success');
    }
}

// Global data manager instance
const dataManager = new DataManager();

// ============================
// CHART MANAGEMENT
// ============================

class ChartManager {
    constructor() {
        this.defaultOptions = this.getDefaultChartOptions();
    }
    
    /**
     * Gets default Chart.js configuration
     * @returns {Object} Default chart options
     */
    getDefaultChartOptions() {
        return {
            responsive: true,
            maintainAspectRatio: false,
            interaction: {
                intersect: false,
                mode: 'index'
            },
            animation: {
                duration: dashboardState.settings.animations ? CONFIG.UI.CHART_ANIMATION_DURATION : 0,
                easing: 'easeOutQuart'
            },
            elements: {
                point: {
                    radius: 4,
                    hoverRadius: 6
                },
                line: {
                    tension: 0.2
                }
            },
            plugins: {
                legend: {
                    position: 'bottom',
                    labels: {
                        usePointStyle: true,
                        boxWidth: 10,
                        font: {
                            size: 12,
                            family: CONFIG.UI.FONT_FAMILY
                        }
                    }
                },
                tooltip: {
                    backgroundColor: 'rgba(30, 64, 175, 0.9)',
                    titleColor: '#ffffff',
                    bodyColor: '#ffffff',
                    borderColor: '#ffffff',
                    borderWidth: 1,
                    cornerRadius: 8,
                    padding: 12,
                    titleFont: {
                        size: 14,
                        weight: 'bold'
                    },
                    bodyFont: {
                        size: 13
                    },
                    callbacks: {
                        label: (context) => {
                            const label = context.dataset.label || '';
                            const value = typeof context.parsed.y !== 'undefined' ? 
                                context.parsed.y : context.parsed;
                            
                            if (typeof value === 'number') {
                                return `${label}: ${value.toLocaleString()}`;
                            }
                            return `${label}: ${value}`;
                        }
                    }
                }
            }
        };
    }
    
    /**
     * Creates a KPI dashboard chart showing completion rates
     * @param {string} canvasId - Canvas element ID
     * @param {Object} data - KPI data
     */
    createKPIChart(canvasId, data) {
        const canvas = document.getElementById(canvasId);
        if (!canvas) {
            console.warn(`Canvas not found: ${canvasId}`);
            return null;
        }
        
        const ctx = canvas.getContext('2d');
        const chartData = this.processKPIData(data);
        
    try { dashboardState.ensureCanvasFree(canvas, canvas && canvas.id ? canvas.id : null); } catch (e) { /* ignore */ }
    const chart = new Chart(ctx, {
            type: 'doughnut',
            data: {
                labels: chartData.labels,
                datasets: [{
                    data: chartData.values,
                    backgroundColor: [
                        CONFIG.COLORS.success,
                        CONFIG.COLORS.warning,
                        CONFIG.COLORS.info,
                        CONFIG.COLORS.error
                    ],
                    borderWidth: 2,
                    borderColor: '#ffffff',
                    hoverBorderWidth: 3
                }]
            },
            options: {
                ...this.defaultOptions,
                cutout: '60%',
                plugins: {
                    ...this.defaultOptions.plugins,
                    legend: {
                        position: 'right'
                    }
                }
            }
        });
        
        dashboardState.setChart(canvasId, chart);
        return chart;
    }
    
    /**
     * Creates collection phase breakdown chart
     * @param {string} canvasId - Canvas element ID
     * @param {Array} data - Collection data
     */
    createCollectionPhaseChart(canvasId, data) {
    // Destroy existing chart on this canvas if present
    try { if (dashboardState.getChart(canvasId)) dashboardState.getChart(canvasId).destroy(); } catch (e) { /* ignore */ }
    const canvas = document.getElementById(canvasId);
        if (!canvas) {
            console.warn(`Canvas not found: ${canvasId}`);
            return null;
        }
        
        const ctx = canvas.getContext('2d');
        const chartData = this.processCollectionData(data);
        
    try { dashboardState.ensureCanvasFree(canvas, canvas && canvas.id ? canvas.id : null); } catch (e) { /* ignore */ }
    const chart = new Chart(ctx, {
            type: 'pie',
            data: {
                labels: chartData.phases,
                datasets: [{
                    data: chartData.totals,
                    backgroundColor: [
                        CONFIG.COLORS.primary,
                        CONFIG.COLORS.secondary,
                        CONFIG.COLORS.agricultural,
                        CONFIG.COLORS.earth,
                        CONFIG.COLORS.highlight
                    ],
                    borderWidth: 2,
                    borderColor: '#ffffff'
                }]
            },
            options: {
                ...this.defaultOptions,
                plugins: {
                    ...this.defaultOptions.plugins,
                    tooltip: {
                        ...this.defaultOptions.plugins.tooltip,
                        callbacks: {
                            label: (context) => {
                                const total = context.dataset.data.reduce((a, b) => a + b, 0);
                                const percentage = ((context.parsed / total) * 100).toFixed(1);
                                return `${context.label}: ${context.parsed.toLocaleString()} (${percentage}%)`;
                            }
                        }
                    }
                }
            }
        });
        
        dashboardState.setChart(canvasId, chart);
        return chart;
    }
    
    /**
     * Creates team performance comparison chart
     * @param {string} canvasId - Canvas element ID
     * @param {Array} data - Team yields data
     */
    createTeamPerformanceChart(canvasId, data) {
    try { if (dashboardState.getChart(canvasId)) dashboardState.getChart(canvasId).destroy(); } catch (e) { /* ignore */ }
    const canvas = document.getElementById(canvasId);
        if (!canvas) {
            console.warn(`Canvas not found: ${canvasId}`);
            return null;
        }
        
        const ctx = canvas.getContext('2d');
        const chartData = this.processYieldsData(data);
        
    try { dashboardState.ensureCanvasFree(canvas, canvas && canvas.id ? canvas.id : null); } catch (e) { /* ignore */ }
    const chart = new Chart(ctx, {
            type: 'bar',
            data: {
                labels: chartData.teams,
                datasets: [{
                    label: 'Parcelles/Équipe/Jour',
                    data: chartData.yields,
                    backgroundColor: CONFIG.COLORS.gradients.primary[0],
                    borderColor: CONFIG.COLORS.gradients.primary[1],
                    borderWidth: 1,
                    borderRadius: 6,
                    borderSkipped: false
                }]
            },
            options: {
                ...this.defaultOptions,
                scales: {
                    y: {
                        beginAtZero: true,
                        title: {
                            display: true,
                            text: 'Parcelles par jour'
                        }
                    },
                    x: {
                        title: {
                            display: true,
                            text: 'Équipes'
                        }
                    }
                }
            }
        });
        
        dashboardState.setChart(canvasId, chart);
        return chart;
    }
    
    /**
     * Creates tool utilization chart
     * @param {string} canvasId - Canvas element ID
     * @param {Array} data - Collection data with tools
     */
    createToolUtilizationChart(canvasId, data) {
    try { if (dashboardState.getChart(canvasId)) dashboardState.getChart(canvasId).destroy(); } catch (e) { /* ignore */ }
    const canvas = document.getElementById(canvasId);
        if (!canvas) {
            console.warn(`Canvas not found: ${canvasId}`);
            return null;
        }
        
        const ctx = canvas.getContext('2d');
        const chartData = this.processToolData(data);
        
    try { dashboardState.ensureCanvasFree(canvas, canvas && canvas.id ? canvas.id : null); } catch (e) { /* ignore */ }
    const chart = new Chart(ctx, {
            type: 'bar', // Changed to supported type (horizontalBar is not supported by Chart.js v3+)
            data: {
                labels: chartData.tools,
                datasets: [{
                    label: 'Utilisation',
                    data: chartData.usage,
                    backgroundColor: CONFIG.COLORS.secondary,
                    borderColor: CONFIG.COLORS.primary,
                    borderWidth: 1
                }]
            },
            options: {
                ...this.defaultOptions,
                indexAxis: 'y',
                scales: {
                    x: {
                        beginAtZero: true,
                        title: {
                            display: true,
                            text: 'Nombre d\'utilisations'
                        }
                    }
                }
            }
        });
        
        dashboardState.setChart(canvasId, chart);
        return chart;
    }
    
    /**
     * Creates multi-month projections chart
     * @param {string} canvasId - Canvas element ID
     * @param {Object} projectionData - NICAD/CTASF projection data
     */
    createProjectionsChart(canvasId, projectionData) {
    try { if (dashboardState.getChart(canvasId)) dashboardState.getChart(canvasId).destroy(); } catch (e) { /* ignore */ }
    const canvas = document.getElementById(canvasId);
        if (!canvas) {
            console.warn(`Canvas not found: ${canvasId}`);
            return null;
        }
        
        const ctx = canvas.getContext('2d');
        const chartData = this.processProjectionData(projectionData);
        
    try { dashboardState.ensureCanvasFree(canvas, canvas && canvas.id ? canvas.id : null); } catch (e) { /* ignore */ }
    const chart = new Chart(ctx, {
            type: 'line',
            data: {
                labels: chartData.months,
                datasets: [
                    {
                        label: 'NICAD Prévisions',
                        data: chartData.nicad,
                        borderColor: CONFIG.COLORS.success,
                        backgroundColor: CONFIG.COLORS.success + '20',
                        fill: false,
                        tension: 0.3
                    },
                    {
                        label: 'CTASF Prévisions',
                        data: chartData.ctasf,
                        borderColor: CONFIG.COLORS.warning,
                        backgroundColor: CONFIG.COLORS.warning + '20',
                        fill: false,
                        tension: 0.3
                    },
                    {
                        label: 'Affichage Public',
                        data: chartData.public,
                        borderColor: CONFIG.COLORS.info,
                        backgroundColor: CONFIG.COLORS.info + '20',
                        fill: false,
                        tension: 0.3
                    }
                ]
            },
            options: {
                ...this.defaultOptions,
                scales: {
                    y: {
                        beginAtZero: true,
                        title: {
                            display: true,
                            text: 'Nombre de parcelles'
                        }
                    },
                    x: {
                        title: {
                            display: true,
                            text: 'Mois'
                        }
                    }
                },
                plugins: {
                    ...this.defaultOptions.plugins,
                    legend: {
                        position: 'top'
                    }
                }
            }
        });
        
        dashboardState.setChart(canvasId, chart);
        return chart;
    }
    
    /**
     * Creates quality control gauge chart
     * @param {string} canvasId - Canvas element ID
     * @param {number} value - Percentage value (0-100)
     * @param {string} label - Chart label
     */
    createGaugeChart(canvasId, value, label) {
    try { if (dashboardState.getChart(canvasId)) dashboardState.getChart(canvasId).destroy(); } catch (e) { /* ignore */ }
    const canvas = document.getElementById(canvasId);
        if (!canvas) {
            console.warn(`Canvas not found: ${canvasId}`);
            return null;
        }
        
        const ctx = canvas.getContext('2d');
        
    // Ensure numeric and clamp between 0 and 100
    let numericValue = Number(value) || 0;
    if (!isFinite(numericValue)) numericValue = 0;
    numericValue = Math.max(0, Math.min(100, numericValue));

    // Determine color based on numericValue
    let color = CONFIG.COLORS.error;
    if (numericValue >= 80) color = CONFIG.COLORS.success;
    else if (numericValue >= 60) color = CONFIG.COLORS.warning;
        
    try { dashboardState.ensureCanvasFree(canvas, canvas && canvas.id ? canvas.id : null); } catch (e) { /* ignore */ }
    const chart = new Chart(ctx, {
            type: 'doughnut',
            data: {
                datasets: [{
                    data: [numericValue, 100 - numericValue],
                    backgroundColor: [color, '#e5e7eb'],
                    borderWidth: 0,
                    cutout: '75%'
                }]
            },
            options: {
                ...this.defaultOptions,
                rotation: -90,
                circumference: 180,
                plugins: {
                    legend: {
                        display: false
                    },
                    tooltip: {
                        enabled: false
                    }
                }
            },
            plugins: [{
                id: 'gaugeText',
                beforeDraw: (chart) => {
                    const { ctx, chartArea } = chart;
                    const centerX = (chartArea.left + chartArea.right) / 2;
                    const centerY = chartArea.bottom - 20;

                    const displayText = `${numericValue.toFixed(1)}%`;

                    ctx.save();
                    ctx.font = 'bold 24px Arial';
                    ctx.fillStyle = color;
                    ctx.textAlign = 'center';
                    ctx.fillText(displayText, centerX, centerY);

                    ctx.font = '14px Arial';
                    ctx.fillStyle = '#6b7280';
                    ctx.fillText(label, centerX, centerY + 20);
                    ctx.restore();
                }
            }]
        });
        
        dashboardState.setChart(canvasId, chart);
        // Also update the textual percentage element next to the canvas if present
        try {
            const textId = canvasId.replace(/Gauge$/, 'Value');
            const textEl = document.getElementById(textId);
            if (textEl) textEl.textContent = `${numericValue.toFixed(1)}%`;
        } catch (e) {
            // ignore
        }

        return chart;
    }
    
    // ============================
    // DATA PROCESSING METHODS
    // ============================
    
    /**
     * Processes KPI data for visualization
     * @param {Array} data - Raw KPI data
     * @returns {Object} Processed chart data
     */
    processKPIData(data) {
        if (!data || !data.length) {
            return { labels: [], values: [] };
        }
        
        // Find completion metrics
        const metrics = {};
        data.forEach(row => {
            const metric = row.metric || row.Metric || '';
            const value = parseFloat(row.value || row.Value || 0);
            
            if (metric.toLowerCase().includes('completion') || metric.toLowerCase().includes('taux')) {
                metrics[metric] = value;
            }
        });
        
        return {
            labels: Object.keys(metrics),
            values: Object.values(metrics)
        };
    }
    
    /**
     * Processes collection data for phase breakdown
     * @param {Array} data - Raw collection data
     * @returns {Object} Processed chart data
     */
    processCollectionData(data) {
        if (!data || !data.length) {
            return { phases: [], totals: [] };
        }
        
        const phaseMap = new Map();
        
        data.forEach(row => {
            const phase = row.phase || row.Phase || 'Unknown';
            const total = parseFloat(row.total || row.Total || 0);
            
            if (phaseMap.has(phase)) {
                phaseMap.set(phase, phaseMap.get(phase) + total);
            } else {
                phaseMap.set(phase, total);
            }
        });
        
        return {
            phases: Array.from(phaseMap.keys()),
            totals: Array.from(phaseMap.values())
        };
    }
    
    /**
     * Processes yields data for team performance
     * @param {Array} data - Raw yields data
     * @returns {Object} Processed chart data
     */
    processYieldsData(data) {
        if (!data || !data.length) {
            return { teams: [], yields: [] };
        }
        
        const teams = [];
        const yields = [];
        
        data.forEach(row => {
            const team = row.team || row.Team || 'Team';
            const teamYield = parseFloat(row.champs_equipe_jour || row['Champs/Equipe/Jour'] || 0);
            
            teams.push(team);
            yields.push(teamYield);
        });
        
        return { teams, yields };
    }
    
    /**
     * Processes tool utilization data
     * @param {Array} data - Raw collection data
     * @returns {Object} Processed chart data
     */
    processToolData(data) {
        if (!data || !data.length) {
            return { tools: [], usage: [] };
        }
        
        const toolMap = new Map();
        
        data.forEach(row => {
            const tool = row.tool || row.Tool || 'Unknown';
            const total = parseFloat(row.total || row.Total || 0);
            
            if (toolMap.has(tool)) {
                toolMap.set(tool, toolMap.get(tool) + total);
            } else {
                toolMap.set(tool, total);
            }
        });
        
        return {
            tools: Array.from(toolMap.keys()),
            usage: Array.from(toolMap.values())
        };
    }
    
    /**
     * Processes projection data for multi-month charts
     * @param {Object} data - Raw projection data
     * @returns {Object} Processed chart data
     */
    processProjectionData(data) {
        if (!data || !data.nicadData || !data.ctasfData || !data.publicData) {
            return { months: [], nicad: [], ctasf: [], public: [] };
        }
        
        const months = ['Sept 2025', 'Oct 2025', 'Nov 2025', 'Dec 2025'];
        const nicad = [];
        const ctasf = [];
        const publicData = [];
        
        // Process NICAD projections
        months.forEach(month => {
            const monthKey = month.toLowerCase().replace(' ', '_');
            const value = data.nicadData.find(row => 
                row[monthKey] !== undefined
            )?.[monthKey] || 0;
            nicad.push(parseFloat(value) || 0);
        });
        
        // Process CTASF projections
        months.forEach(month => {
            const monthKey = month.toLowerCase().replace(' ', '_');
            const value = data.ctasfData.find(row => 
                row[monthKey] !== undefined
            )?.[monthKey] || 0;
            ctasf.push(parseFloat(value) || 0);
        });
        
        // Process Public Display projections
        months.forEach(month => {
            const monthKey = month.toLowerCase().replace(' ', '_');
            const value = data.publicData.find(row => 
                row[monthKey] !== undefined
            )?.[monthKey] || 0;
            publicData.push(parseFloat(value) || 0);
        });
        
        return {
            months,
            nicad,
            ctasf,
            public: publicData
        };
    }
}

// Global chart manager instance
const chartManager = new ChartManager();

// ============================
// UI MANAGEMENT
// ============================

class UIManager {
    constructor() {
        this.activeTab = 'executive';
        this.loadingCount = 0;
        this.notifications = [];
    }
    
    /**
     * Initializes the dashboard UI
     */
    async initialize() {
        try {
            // Show loading screen
            this.showLoadingScreen();
            
            // Load all data
            await dataManager.fetchAllData();
            
            // Initialize UI components
            this.setupTabNavigation();
            this.setupRefreshControls();
            this.setupSettings();
            this.setupExportControls();
            this.setupKeyboardShortcuts();
            
            // Initialize dashboard content
            await this.initializeAllTabs();
            
            // Setup auto-refresh
            if (dashboardState.settings.autoRefresh) {
                this.startAutoRefresh();
            }
            
            // Hide loading screen and show dashboard
            await this.hideLoadingScreen();
            
            showNotification('Dashboard loaded successfully', 'success');
            
        } catch (error) {
            console.error('Dashboard initialization failed:', error);
            showNotification('Failed to initialize dashboard', 'error');
            this.hideLoadingScreen();
        }
    }
    
    /**
     * Shows the loading screen
     */
    showLoadingScreen() {
        const loadingScreen = document.getElementById('loadingScreen');
        const dashboardContainer = document.getElementById('dashboardContainer');
        
        if (loadingScreen) loadingScreen.classList.remove('hidden');
        if (dashboardContainer) dashboardContainer.classList.add('hidden');
    }

    /**
     * Determine commune status (method on UIManager)
     * @param {Object} commune
     * @returns {{class: string, text: string}}
     */
    getStatus(commune) {
    if (!commune) return { class: 'status-pending', text: 'En attente' };
    // Prefer French sheet headers when available. Fall back to legacy/english keys.
    const collected = Number(commune.collected_parcels_no_duplicates || commune.collected || 0);
    if (collected === 0) return { class: 'status-pending', text: 'En attente' };
    const validated = Number(commune.parcelles_validees_par_lurm || commune.parcelles_validees || commune.validated || commune.urm_validated_parcels || 0);
        if (validated && collected) {
            const validationRate = (validated / collected) * 100;
            if (validationRate >= 80) return { class: 'status-completed', text: 'Terminé' };
            if (validationRate >= 50) return { class: 'status-processing', text: 'En cours' };
        }
        return { class: 'status-rejected', text: 'Problèmes' };
    }
    
    /**
     * Hides the loading screen and shows dashboard
     */
    async hideLoadingScreen() {
        return new Promise((resolve) => {
            setTimeout(() => {
                const loadingScreen = document.getElementById('loadingScreen');
                const dashboardContainer = document.getElementById('dashboardContainer');
                
                if (loadingScreen) {
                    loadingScreen.classList.add('fade-out');
                    setTimeout(() => {
                        loadingScreen.classList.add('hidden');
                        loadingScreen.classList.remove('fade-out');
                    }, 500);
                }
                
                if (dashboardContainer) {
                    dashboardContainer.classList.remove('hidden');
                    dashboardContainer.classList.add('fade-in');
                }
                
                resolve();
            }, CONFIG.UI.LOADING_MIN_DURATION);
        });
    }
    
    /**
     * Sets up tab navigation
     */
    setupTabNavigation() {
        const navTabs = document.querySelectorAll('.nav-tab');
        const tabContents = document.querySelectorAll('.tab-content');
        
        navTabs.forEach(tab => {
            tab.addEventListener('click', (e) => {
                e.preventDefault();
                
                const targetTab = tab.dataset.tab;
                if (!targetTab || targetTab === this.activeTab) return;
                
                // Update active states
                navTabs.forEach(t => t.classList.remove('active'));
                tab.classList.add('active');
                
                tabContents.forEach(content => {
                    content.classList.remove('active');
                    if (content.id === targetTab) {
                        content.classList.add('active');
                    }
                });
                
                this.activeTab = targetTab;
                
                // Initialize tab-specific content
                this.initializeTab(targetTab);
            });
        });
    }
    
    /**
     * Sets up refresh controls
     */
    setupRefreshControls() {
        const refreshBtn = document.getElementById('refreshBtn');
        if (refreshBtn) {
            refreshBtn.addEventListener('click', async () => {
                await this.refreshDashboard(true);
            });
        }
        
        // Update last refresh time
        this.updateLastRefreshTime();
    }
    
    /**
     * Sets up export controls
     */
    setupExportControls() {
        const exportBtn = document.getElementById('exportBtn');
        if (exportBtn) {
            exportBtn.addEventListener('click', () => {
                this.exportData();
            });
        }
        
        // Setup chart download buttons
        document.addEventListener('click', (e) => {
            if (e.target.classList.contains('chart-download') || 
                e.target.closest('.chart-download')) {
                const button = e.target.classList.contains('chart-download') ? 
                    e.target : e.target.closest('.chart-download');
                const chartId = button.dataset.chart;
                if (chartId) {
                    this.downloadChart(chartId);
                }
            }
        });
    }
    
    /**
     * Sets up settings panel
     */
    setupSettings() {
        const settingsBtn = document.getElementById('settingsBtn');
        const settingsPanel = document.getElementById('settingsPanel');
        const closeSettings = document.getElementById('closeSettings');
        
        if (settingsBtn && settingsPanel) {
            settingsBtn.addEventListener('click', () => {
                settingsPanel.classList.toggle('hidden');
            });
        }
        
        if (closeSettings && settingsPanel) {
            closeSettings.addEventListener('click', () => {
                settingsPanel.classList.add('hidden');
            });
        }
        
        // Setup setting controls
        this.setupSettingControls();
    }
    
    /**
     * Sets up individual setting controls
     */
    setupSettingControls() {
        // Auto-refresh toggle
        const autoRefreshToggle = document.getElementById('autoRefreshToggle');
        if (autoRefreshToggle) {
            autoRefreshToggle.checked = dashboardState.settings.autoRefresh;
            autoRefreshToggle.addEventListener('change', (e) => {
                dashboardState.updateSetting('autoRefresh', e.target.checked);
                if (e.target.checked) {
                    this.startAutoRefresh();
                } else {
                    this.stopAutoRefresh();
                }
            });
        }
        
        // Refresh interval
        const refreshInterval = document.getElementById('refreshInterval');
        if (refreshInterval) {
            refreshInterval.value = dashboardState.settings.refreshInterval;
            refreshInterval.addEventListener('change', (e) => {
                const interval = parseInt(e.target.value);
                if (!isNaN(interval) && interval > 0) {
                    dashboardState.updateSetting('refreshInterval', interval);
                    if (dashboardState.settings.autoRefresh) {
                        this.startAutoRefresh(); // Restart with new interval
                    }
                }
            });
        }
        
        // Dark mode toggle
        const darkModeToggle = document.getElementById('darkModeToggle');
        if (darkModeToggle) {
            darkModeToggle.checked = dashboardState.settings.darkMode;
            darkModeToggle.addEventListener('change', (e) => {
                dashboardState.updateSetting('darkMode', e.target.checked);
                this.toggleDarkMode(e.target.checked);
            });
        }
        
        // Animations toggle
        const animationsToggle = document.getElementById('animationsToggle');
        if (animationsToggle) {
            animationsToggle.checked = dashboardState.settings.animations;
            animationsToggle.addEventListener('change', (e) => {
                dashboardState.updateSetting('animations', e.target.checked);
            });
        }
        
        // Alerts toggle
        const alertsToggle = document.getElementById('alertsToggle');
        if (alertsToggle) {
            alertsToggle.checked = dashboardState.settings.alerts;
            alertsToggle.addEventListener('change', (e) => {
                dashboardState.updateSetting('alerts', e.target.checked);
            });
        }
        
        // Sound toggle
        const soundToggle = document.getElementById('soundToggle');
        if (soundToggle) {
            soundToggle.checked = dashboardState.settings.sound;
            soundToggle.addEventListener('change', (e) => {
                dashboardState.updateSetting('sound', e.target.checked);
            });
        }
    }
    
    /**
     * Sets up keyboard shortcuts
     */
    setupKeyboardShortcuts() {
        document.addEventListener('keydown', (e) => {
            // Only process shortcuts when not in input fields
            if (e.target.tagName === 'INPUT' || e.target.tagName === 'TEXTAREA') {
                return;
            }
            
            switch (e.key) {
                case '1':
                case '2':
                case '3':
                case '4':
                case '5':
                case '6':
                    e.preventDefault();
                    this.switchToTab(parseInt(e.key) - 1);
                    break;
                case 'r':
                case 'R':
                    if (e.ctrlKey || e.metaKey) {
                        e.preventDefault();
                        this.refreshDashboard(true);
                    }
                    break;
                case 's':
                case 'S':
                    if (e.ctrlKey || e.metaKey) {
                        e.preventDefault();
                        document.getElementById('settingsPanel')?.classList.toggle('hidden');
                    }
                    break;
            }
        });
    }
    
    /**
     * Switches to tab by index
     * @param {number} tabIndex - Tab index (0-based)
     */
    switchToTab(tabIndex) {
        const navTabs = document.querySelectorAll('.nav-tab');
        if (navTabs[tabIndex]) {
            navTabs[tabIndex].click();
        }
    }
    
    /**
     * Toggles dark mode
     * @param {boolean} enabled - Whether dark mode is enabled
     */
    toggleDarkMode(enabled) {
        document.documentElement.setAttribute('data-theme', enabled ? 'dark' : 'light');
    }
    
    /**
     * Starts auto-refresh timer
     */
    startAutoRefresh() {
        this.stopAutoRefresh(); // Clear existing timer
        
        const intervalMs = dashboardState.settings.refreshInterval * 60 * 1000;
        dashboardState.refreshTimer = setInterval(() => {
            this.refreshDashboard(false);
        }, intervalMs);
    }
    
    /**
     * Stops auto-refresh timer
     */
    stopAutoRefresh() {
        if (dashboardState.refreshTimer) {
            clearInterval(dashboardState.refreshTimer);
            dashboardState.refreshTimer = null;
        }
    }
    
    /**
     * Refreshes dashboard data and UI
     * @param {boolean} showLoader - Whether to show loading indicator
     */
    async refreshDashboard(showLoader = false) {
        try {
            if (showLoader) {
                showLoadingIndicator('Refreshing data...');
            }
            
            // Fetch latest data
            await dataManager.fetchAllData(true);
            
            // Update current tab
            await this.initializeTab(this.activeTab);
            
            // Update last refresh time
            this.updateLastRefreshTime();
            
            if (showLoader) {
                hideLoadingIndicator();
                showNotification('Data refreshed successfully', 'success');
            }
            
        } catch (error) {
            console.error('Refresh failed:', error);
            if (showLoader) {
                hideLoadingIndicator();
            }
            showNotification('Failed to refresh data', 'error');
        }
    }
    
    /**
     * Updates last refresh time display
     */
    updateLastRefreshTime() {
        const lastUpdateElement = document.getElementById('lastUpdateTime');
        if (lastUpdateElement && dashboardState.lastUpdate) {
            const timeStr = dashboardState.lastUpdate.toLocaleTimeString('fr-FR', {
                hour: '2-digit',
                minute: '2-digit'
            });
            lastUpdateElement.textContent = `Dernière mise à jour: ${timeStr}`;
        }
    }
    
    /**
     * Initializes all tabs
     */
    async initializeAllTabs() {
        // Initialize executive summary (default active)
        await this.initializeTab('executive');
    }
    
    /**
     * Initializes specific tab content
     * @param {string} tabName - Tab name to initialize
     */
    async initializeTab(tabName) {
        try {
            switch (tabName) {
                case 'executive':
                    await this.initializeExecutiveTab();
                    break;
                case 'collection':
                    await this.initializeCollectionTab();
                    break;
                case 'geographic':
                    await this.initializeGeographicTab();
                    break;
                case 'projections':
                    await this.initializeProjectionsTab();
                    break;
                case 'quality':
                    await this.initializeQualityTab();
                    break;
                case 'workflow':
                    await this.initializeWorkflowTab();
                    break;
                case 'commune-status':
                    await this.initializeCommuneStatusTab();
                    break;
                default:
                    console.warn(`Unknown tab: ${tabName}`);
            }
        } catch (error) {
            console.error(`Failed to initialize tab ${tabName}:`, error);
            showNotification(`Failed to initialize ${tabName} tab`, 'error');
        }
    }
    
    /**
     * Initializes executive summary tab
     */
    async initializeExecutiveTab() {
        try {
            // Update KPI cards
            this.updateKPICards();
            
            // Create charts
            const kpiData = dashboardState.getData('kpi_dashboard') || [];
            const collectionData = dashboardState.getData('data_collection') || [];
            const yieldsData = dashboardState.getData('yields_projection') || [];
            
            chartManager.createCollectionPhaseChart('collectionPhaseChart', collectionData);
            // Replace Team Performance chart with Progression globale metric block
            this.renderProgressionGlobale(collectionData, yieldsData);
            
            // Update regional status
            this.updateRegionalStatus();
            
            // Setup quick actions
            this.setupQuickActions();
        } catch (error) {
            console.error('Failed to initialize executive tab:', error);
            showNotification('Failed to load executive summary', 'error');
        }
    }

    /**
     * Renders the Progression globale metrics block
     * @param {Array} collectionData
     * @param {Array} yieldsData
     */
    renderProgressionGlobale(collectionData = [], yieldsData = []) {
        try {
            const container = document.getElementById('progressionGlobale');
            if (!container) return;

            // Clear existing
            container.innerHTML = '';
            // Prefer totals from 'commune_status' sheet when available
            const communeData = dashboardState.getData('commune_status') || [];

            // Strict exact-key summing helpers: only use exact French headers from commune_status
            const exactSum = (rows, keyName) => {
                if (!rows || !rows.length) return 0;
                return rows.reduce((s, row) => s + (Number(row && Object.prototype.hasOwnProperty.call(row, keyName) ? (row[keyName] || 0) : 0) || 0), 0);
            };

            let collected = 0, validated = 0, deliberated = 0, ctasfCollected = 0, ctasfValidated = 0;

            // Normalize commune_status rows using a header map (French headers -> canonical keys)
            const headerMapLocal = {
                'Commune': 'commune',
                'Région': 'region',
                'Total Parcelles': 'totalParcelles',
                '% du total': 'percentTotal',
                'NICAD': 'nicad',
                '% NICAD': 'percentNicad',
                'CTASF': 'ctasf',
                // possible header variants for CTASF validated counts
                'CTASF validées': 'ctasf_validated',
                'CTASF validées par l\u2019URM': 'ctasf_validated',
                "CTASF validées par l'URM": 'ctasf_validated',
                'Parcelles CTASF validées': 'ctasf_validated',
                '% CTASF': 'percentCtasf',
                'Délibérées': 'deliberated',
                '% Délibérée': 'percentDeliberated',
                'Parcelles brutes': 'parcellesBrutes',
                'Parcelles collectées (sans doublon géométrique)': 'collected',
                'Parcelles enquêtées': 'surveyed',
                'Motifs de rejet post-traitement': 'rejectionReasons',
                'Parcelles retenues après post-traitement': 'retained',
                'Parcelles validées par l\u2019URM': 'validated',
                'Parcelles rejetées par l\u2019URM': 'rejected',
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
                'Message d\u2019erreur jointure': 'jointureErrorMessage'
            };

            const normalizeRows = (rows) => {
                if (!rows || !rows.length) return [];
                return rows.map(r => {
                    const nr = {};
                    // copy existing canonical-like keys first
                    Object.keys(r || {}).forEach(k => { nr[k] = r[k]; });
                    // map known French headers to canonical keys
                    Object.entries(headerMapLocal).forEach(([src, dest]) => {
                        if (Object.prototype.hasOwnProperty.call(r, src) && (r[src] !== undefined)) {
                            // only set canonical key if not already present to prefer existing canonical fields
                            if (!Object.prototype.hasOwnProperty.call(nr, dest) || nr[dest] === undefined) nr[dest] = r[src];
                        }
                    });
                    return nr;
                });
            };

            const normalized = normalizeRows(communeData || []);

            if (normalized && normalized.length) {
                // Exclude any aggregate/total rows (e.g., 'TOTAL', 'Total') to avoid double-counting
                const isTotalName = (n) => {
                    if (!n) return false;
                    const s = String(n).trim().toLowerCase();
                    return s === 'total' || s === 'totale' || s === 'totaux' || s.startsWith('total ') || s === 'totales';
                };

                const rowsForSums = normalized.filter(r => !isTotalName(r.commune || r.Commune || ''));

                // Debug: how many rows were excluded
                try { console.info('[progression-debug] excluded total/summary rows', normalized.length - rowsForSums.length); } catch (e) {}

                // Debug: report mapped CTASF-validated source headers and quick check for values
                try {
                    const mappedCtasfValidatedSources = Object.keys(headerMapLocal).filter(k => headerMapLocal[k] === 'ctasf_validated');
                    if (mappedCtasfValidatedSources.length) {
                        console.info('[progression-debug] mappedCtasfValidatedSources', mappedCtasfValidatedSources);
                        const sumFromOrig = rowsForSums.reduce((s, row) => {
                            return s + mappedCtasfValidatedSources.reduce((rs, src) => rs + (Number(row[src] || 0) || 0), 0);
                        }, 0);
                        console.info('[progression-debug] ctasfValidated found in original columns sum=', sumFromOrig);
                    }
                } catch (e) { /* ignore */ }

                collected = rowsForSums.reduce((s, row) => s + (Number(row.collected || 0) || 0), 0);
                validated = rowsForSums.reduce((s, row) => s + (Number(row.validated || 0) || 0), 0);
                deliberated = rowsForSums.reduce((s, row) => s + (Number(row.deliberated || 0) || 0), 0);
                ctasfCollected = rowsForSums.reduce((s, row) => s + (Number(row.ctasf || 0) || 0), 0);
                // CTASF validated may not be present in headerMap; try common canonical key
                ctasfValidated = rowsForSums.reduce((s, row) => s + (Number(row.ctasf_validated || row.ctasf_validees || 0) || 0), 0);

                // Fallback: if canonical CTASF-validated is zero, try to detect columns whose header contains 'ctasf' and 'valid' (handles header variants)
                try {
                    if (!ctasfValidated) {
                        const candidateKeys = new Set();
                        rowsForSums.forEach(row => {
                            Object.keys(row || {}).forEach(k => {
                                const lk = String(k || '').toLowerCase();
                                if (lk.includes('ctasf') && (lk.includes('valid') || lk.includes('valide') || lk.includes('validees') || lk.includes('validées') || lk.includes('validated'))) {
                                    candidateKeys.add(k);
                                }
                            });
                        });

                        if (candidateKeys.size) {
                            console.info('[progression-debug] ctasf candidate keys', Array.from(candidateKeys));
                            const fallbackSum = rowsForSums.reduce((sum, row) => {
                                return sum + Array.from(candidateKeys).reduce((rs, k) => {
                                    const raw = row[k];
                                    const n = Number(String(raw || '').replace(/[^0-9.-]/g, ''));
                                    return rs + (Number.isNaN(n) ? 0 : n);
                                }, 0);
                            }, 0);
                            console.info('[progression-debug] ctasfValidated from candidate keys sum=', fallbackSum);
                            if (fallbackSum) ctasfValidated = fallbackSum;
                        }
                    }
                } catch (e) { /* ignore fallback errors */ }

                // Extra inspection: for rows that have CTASF counts, print their CTASF-related keys/values to help mapping
                try {
                    rowsForSums.forEach((row, ridx) => {
                        const ctasfKeysPresent = Object.keys(row || {}).filter(k => String(k||'').toLowerCase().includes('ctasf'));
                        const anyCtasfCount = Number(row.ctasf || row['CTASF'] || row.ctasf_collected || 0) || 0;
                        if ((ctasfKeysPresent.length || anyCtasfCount) && anyCtasfCount > 0) {
                            const keyValues = {};
                            ctasfKeysPresent.forEach(k => { keyValues[k] = row[k]; });
                            console.info('[progression-debug] ctasf-row-inspect', { idx: ridx, commune: row.commune || row.Commune || `(row ${ridx})`, ctasfKeys: ctasfKeysPresent, keyValues, rawRow: row });
                        }
                    });
                } catch (e) { /* ignore */ }

                // Heuristic fallback: if ctasfValidated still zero, try computing it from a per-row CTASF percentage column (e.g. '_ctasf' = 19.9 means 19.9%)
                try {
                    if (!ctasfValidated) {
                        // find candidate percent keys (contain 'ctasf' and look like percentages)
                        const percentKeys = new Set();
                        rowsForSums.forEach(row => {
                            Object.keys(row || {}).forEach(k => {
                                const lk = String(k||'').toLowerCase();
                                if (lk.includes('ctasf') && !lk.includes('valid') && !lk.includes('valide') ) {
                                    const v = String(row[k] || '').replace(/[,%\s]/g, '');
                                    const n = Number(v);
                                    // Accept percent-like values: numeric between 0 and 100 (allow integers, decimals, or percent strings)
                                    if (!Number.isNaN(n) && n > 0 && n <= 100) {
                                        percentKeys.add(k);
                                    }
                                }
                            });
                        });

                        if (percentKeys.size) {
                            console.info('[progression-debug] ctasf percent candidate keys', Array.from(percentKeys));
                            const percentSum = rowsForSums.reduce((sum, row) => {
                                // read CTASF count and percent (if present)
                                const c = Number(row.ctasf || row['CTASF'] || 0) || 0;
                                if (c <= 0) return sum;
                                let p = 0;
                                for (const pk of percentKeys) {
                                    const raw = String(row[pk] || '').replace(/[^0-9.-]/g, '');
                                    const num = Number(raw);
                                    if (!Number.isNaN(num) && num > 0 && num <= 100) { p = num; break; }
                                }
                                return sum + Math.round((c * p) / 100);
                            }, 0);
                            console.info('[progression-debug] ctasfValidated computed from percent keys sum=', percentSum);
                            if (percentSum) ctasfValidated = percentSum;
                        }
                    }
                } catch (e) { /* ignore */ }
            } else {
                collected = 0; validated = 0; deliberated = 0; ctasfCollected = 0; ctasfValidated = 0;
            }

            // Immediate debug: log totals and search for suspicious raw values that users reported
            try {
                console.info('[progression-debug] totals', { collected, validated, deliberated, ctasfCollected, ctasfValidated });
                const suspects = [58332, 88914];
                normalized.forEach((row, idx) => {
                    Object.keys(row || {}).forEach(k => {
                        const v = row[k];
                        if (v !== null && v !== undefined && v !== '' ) {
                            const n = Number(String(v).replace(/[^0-9.-]/g,''));
                            if (suspects.includes(n)) {
                                console.warn(`[progression-debug] Found suspicious value ${n} in row ${idx}`, { key: k, value: v, row });
                            }
                        }
                    });
                });
                // Per-row validated source tracing
                try {
                    const mappedValidatedSources = Object.keys(headerMapLocal).filter(k => headerMapLocal[k] === 'validated');
                    const rowsWithValidated = normalized.map((nr, i) => ({
                        idx: i,
                        commune: (normalized[i] && (normalized[i].commune || normalized[i].Commune)) || `(row ${i})`,
                        validated: Number(nr.validated || 0),
                        collected: Number(nr.collected || 0),
                        origRow: communeData[i] || {}
                    })).filter(r => r.validated > 0);

                    if (rowsWithValidated.length) {
                        console.info('[progression-debug] rows with validated (count)', rowsWithValidated.length);
                        // log each row's matching original keys and any mapped source
                        rowsWithValidated.forEach(r => {
                            const matches = [];
                            Object.keys(r.origRow || {}).forEach(k => {
                                const cell = r.origRow[k];
                                const num = Number(String(cell).replace(/[^0-9.-]/g,''));
                                if (!Number.isNaN(num) && num === r.validated) matches.push(k);
                            });
                            const mappedSourcesPresent = mappedValidatedSources.filter(src => Object.prototype.hasOwnProperty.call(r.origRow || {}, src));
                            console.info('[progression-debug-row]', { idx: r.idx, commune: r.commune, validated: r.validated, collected: r.collected, matchedKeys: matches, mappedSourcesPresent });
                        });

                        // Top contributors by validated
                        const top = rowsWithValidated.slice().sort((a,b) => b.validated - a.validated).slice(0,10);
                        console.info('[progression-debug] top validated contributors', top.map(t => ({ commune: t.commune, validated: t.validated, collected: t.collected })));
                    } else {
                        console.info('[progression-debug] no rows with validated>0');
                    }
                } catch (e) { /* ignore per-row tracing errors */ }
            } catch (e) { /* ignore debug errors */ }

            // Target (try to read from KPI sheet if present)
            let target = 70000;
            try {
                const kpi = dashboardState.getData('kpi_dashboard') || [];
                if (kpi && kpi.length) {
                    const trow = kpi.find(r => (String(r.metric || r.Metric || '').toLowerCase()).includes('objectif') || (String(r.metric || r.Metric || '').toLowerCase()).includes('target'));
                    if (trow) target = Number(trow.value || trow.Value || target) || target;
                }
            } catch (e) { /* ignore */ }

            const metrics = [
                // Taux de validation using the exact URM-validated column
                {
                    label: `Taux de validation (URM) (${target.toLocaleString()})`,
                    value: target ? (Number(validated) / Number(target) * 100) : 0,
                    raw: { num: Number(validated), denom: Number(target) }
                },
                // Taux de collecte: Parcelles collectées (sans doublon géométrique) / objectif
                {
                    label: `Taux de collecte (Sans doublons) (${target.toLocaleString()})`,
                    value: target ? (Number(collected) / Number(target) * 100) : 0,
                    raw: { num: Number(collected), denom: Number(target) }
                },
                // Délibérées / Collectées using exact French headers
                {
                    label: `Délibérées / Collectées `,
                    value: collected ? (Number(deliberated) / Number(collected) * 100) : 0,
                    raw: { num: Number(deliberated), denom: Number(collected) }
                },
                // CTASF / Collectées
                {
                    label: `CTASF / Collectées `,
                    value: collected ? (Number(ctasfCollected) / Number(collected) * 100) : 0,
                    raw: { num: Number(ctasfCollected), denom: Number(collected) }
                },
                // CTASF / Validées: CTASF collected over Parcelles validées par l'URM (numerator = CTASF collected, denom = URM-validated)
                {
                    label: `CTASF / Validées (URM)`,
                    value: Number(validated) ? (Number(ctasfCollected) / Number(validated) * 100) : 0,
                    raw: { num: Number(ctasfCollected), denom: Number(validated) }
                },
                // Délibérées / Validées
                {
                    label: `Délibérées / Validées (URM)`,
                    value: Number(validated) ? (Number(deliberated) / Number(validated) * 100) : 0,
                    raw: { num: Number(deliberated), denom: Number(validated) }
                }
            ];

            // Optional debug log for metric raw values
            if (CONFIG.UI && CONFIG.UI.debugMetrics) {
                console.debug('[metrics-debug] collected=', collected, 'validated=', validated, 'deliberated=', deliberated, 'ctasfCollected=', ctasfCollected, 'ctasfValidated=', ctasfValidated);
            }

            // Render rows
            metrics.forEach(m => {
                const row = document.createElement('div');
                row.className = 'progress-row';

                const label = document.createElement('div');
                label.className = 'progress-label';
                label.textContent = m.label;

                const barWrap = document.createElement('div');
                barWrap.className = 'progress-bar-wrap';

                const bar = document.createElement('div');
                bar.className = 'progress-bar-bg';

                const fill = document.createElement('div');
                fill.className = 'progress-bar-fill';
                const pct = Math.max(0, Math.min(100, (isFinite(m.value) ? m.value : 0)));
                fill.style.width = `${pct.toFixed(1)}%`;
                // color coding similar to KPIs: red for low, orange for mid, green for high
                if (pct >= 60) fill.style.backgroundColor = CONFIG.COLORS.warning;
                if (pct >= 80) fill.style.backgroundColor = CONFIG.COLORS.success;
                if (pct < 40) fill.style.backgroundColor = CONFIG.COLORS.error;

                const percentText = document.createElement('div');
                percentText.className = 'progress-percent';
                percentText.textContent = `${pct.toFixed(1)}%`;
                // Add tooltip with raw numbers
                if (m.raw) {
                    row.title = `${m.raw.num.toLocaleString()} / ${m.raw.denom.toLocaleString()}`;
                }

                bar.appendChild(fill);
                barWrap.appendChild(bar);

                row.appendChild(label);
                row.appendChild(barWrap);
                row.appendChild(percentText);

                container.appendChild(row);
            });

        } catch (error) {
            console.error('Failed to render Progression globale:', error);
        }
    }
    
    /**
     * Updates KPI cards with latest data
     */
    updateKPICards() {
        try {
            const communeData = dashboardState.getData('commune_status') || [];
            
            if (!communeData.length) {
                showNotification('No commune data available', 'warning');
                return;
            }
            
            // Calculate totals
            let totalParcels = 0;
            let totalNicad = 0;
            let totalCtasf = 0;
            let totalDeliberated = 0;
            
            communeData.forEach(row => {
                if (row.commune && row.commune.toLowerCase() !== 'total') {
                    totalParcels += parseFloat(row.total_parcels || 0);
                    totalNicad += parseFloat(row.nicad || 0);
                    totalCtasf += parseFloat(row.ctasf || 0);
                    totalDeliberated += parseFloat(row.deliberated || 0);
                }
            });
            
            // Update KPI displays
            const totalParcelsKPI = document.getElementById('totalParcelsKPI');
            if (totalParcelsKPI) {
                this.animateNumber(totalParcelsKPI, totalParcels);
            }
            
            const nicadCompletion = totalParcels > 0 ? (totalNicad / totalParcels) * 100 : 0;
            const ctasfCompletion = totalParcels > 0 ? (totalCtasf / totalParcels) * 100 : 0;
            const deliberationCompletion = totalParcels > 0 ? (totalDeliberated / totalParcels) * 100 : 0;
            
            this.updateKPICard('nicadCompletionKPI', 'nicadProgressFill', nicadCompletion);
            this.updateKPICard('ctasfCompletionKPI', 'ctasfProgressFill', ctasfCompletion);
            this.updateKPICard('deliberationKPI', 'deliberationProgressFill', deliberationCompletion); // Completed the method
        } catch (error) {
            console.error('Failed to update KPI cards:', error);
            showNotification('Failed to update KPI cards', 'error');
        }
    }
    
    /**
     * Updates a KPI card with animated progress
     * @param {string} kpiId - KPI element ID
     * @param {string} progressId - Progress fill element ID
     * @param {number} percentage - Percentage value
     */
    updateKPICard(kpiId, progressId, percentage) {
        try {
            const kpiElement = document.getElementById(kpiId);
            const progressElement = document.getElementById(progressId);
            
            if (kpiElement) {
                this.animatePercentage(kpiElement, percentage);
            }
            
            if (progressElement) {
                setTimeout(() => {
                    progressElement.style.width = `${Math.min(percentage, 100)}%`;
                    progressElement.style.backgroundColor = this.getProgressColor(percentage);
                }, 500);
            }
        } catch (error) {
            console.error(`Failed to update KPI card ${kpiId}:`, error);
            showNotification(`Failed to update KPI ${kpiId}`, 'error');
        }
    }
    
    /**
     * Animates a number counter
     * @param {HTMLElement} element - Target element
     * @param {number} targetValue - Target number value
     */
    animateNumber(element, targetValue) {
        if (!dashboardState.settings.animations) {
            element.textContent = targetValue.toLocaleString();
            return;
        }
        
        const duration = 2000;
        const startValue = 0;
        const startTime = performance.now();
        
        const animate = (currentTime) => {
            const elapsedTime = currentTime - startTime;
            const progress = Math.min(elapsedTime / duration, 1);
            
            // Easing function (easeOutQuart)
            const easeProgress = 1 - Math.pow(1 - progress, 4);
            const currentValue = Math.floor(startValue + (targetValue - startValue) * easeProgress);
            
            element.textContent = currentValue.toLocaleString();
            
            if (progress < 1) {
                requestAnimationFrame(animate);
            }
        };
        
        requestAnimationFrame(animate);
    }
    
    /**
     * Animates a percentage value
     * @param {HTMLElement} element - Target element
     * @param {number} targetPercentage - Target percentage
     */
    animatePercentage(element, targetPercentage) {
        if (!dashboardState.settings.animations) {
            element.textContent = `${targetPercentage.toFixed(1)}%`;
            return;
        }
        
        const duration = 2000;
        const startPercentage = 0;
        const startTime = performance.now();
        
        const animate = (currentTime) => {
            const elapsedTime = currentTime - startTime;
            const progress = Math.min(elapsedTime / duration, 1);
            
            const easeProgress = 1 - Math.pow(1 - progress, 4);
            const currentPercentage = startPercentage + (targetPercentage - startPercentage) * easeProgress;
            
            element.textContent = `${currentPercentage.toFixed(1)}%`;
            
            if (progress < 1) {
                requestAnimationFrame(animate);
            }
        };
        
        requestAnimationFrame(animate);
    }
    
    /**
     * Gets progress color based on percentage
     * @param {number} percentage - Percentage value
     * @returns {string} Color code
     */
    getProgressColor(percentage) {
        if (percentage >= 75) return CONFIG.COLORS.success;
        if (percentage >= 50) return CONFIG.COLORS.warning;
        return CONFIG.COLORS.error;
    }
    
    /**
     * Updates regional status information
     */
    updateRegionalStatus() {
        try {
            const communeData = dashboardState.getData('commune_status') || [];

            // Build a dynamic map of regions from the commune rows
            const regionsMap = {};
            (communeData || []).forEach(row => {
                if (!row) return;
                const rawRegion = (row.region || row.Région || '').trim();
                if (!rawRegion) return;
                // skip TOTAL rows
                if (String(row.commune || '').trim().toLowerCase() === 'total') return;

                const key = rawRegion.toLowerCase();
                if (!regionsMap[key]) regionsMap[key] = { display: rawRegion, communes: 0, completed: 0 };
                regionsMap[key].communes++;
                if (this.isCompleted(row)) regionsMap[key].completed++;
            });

            // Prefer rendering dynamic cards into a container if present
            const container = document.getElementById('regionalStatusCards') || document.getElementById('regionalStatus') || document.querySelector('.regional-status-cards');

            if (container) {
                const cards = Object.values(regionsMap).map(r => {
                    const completion = r.communes > 0 ? (r.completed / r.communes) * 100 : 0;
                    const statusClass = completion >= 80 ? 'status-completed' : (completion >= 50 ? 'status-processing' : 'status-pending');
                    const statusText = completion >= 80 ? 'TERMINÉ' : (completion >= 50 ? 'EN COURS' : 'EN ATTENTE');
                    return `\n                        <div class="region-card" role="region" aria-label="${sanitizeHTML(r.display)}">\n                            <div class="region-card-inner">\n                                <div class="region-title">\n                                    <h4>${sanitizeHTML(r.display)}</h4>\n                                    <span class="status-badge ${statusClass}">${statusText}</span>\n                                </div>\n                                <div class="region-meta">\n                                    <div class="meta-item"><div class="meta-label">COMMUNES</div><div class="meta-value">${r.communes}</div></div>\n                                    <div class="meta-item"><div class="meta-label">COMPLETION</div><div class="meta-value">${completion.toFixed(1)}%</div></div>\n                                </div>\n                            </div>\n                        </div>`;
                }).join('') || '<p>Aucune région disponible</p>';
                container.innerHTML = `<div class="region-cards-grid">${cards}</div>`;
            } else {
                // Fallback: update known hard-coded elements (backwards compatible)
                Object.keys(regionsMap).forEach(k => {
                    const r = regionsMap[k];
                    const regionKey = k.replace(/[^a-z0-9]/g, '');
                    const communesElement = document.getElementById(`${regionKey}Communes`);
                    const completionElement = document.getElementById(`${regionKey}Completion`);
                    const completion = r.communes > 0 ? (r.completed / r.communes) * 100 : 0;
                    if (communesElement) communesElement.textContent = r.communes;
                    if (completionElement) completionElement.textContent = `${completion.toFixed(1)}%`;
                });
            }
        } catch (error) {
            console.error('Failed to update regional status:', error);
            showNotification('Failed to update regional status', 'error');
        }
    }
    
    /**
     * Checks if a commune is completed
     * @param {Object} row - Commune data row
     * @returns {boolean} Whether commune is completed
     */
    isCompleted(row) {
    const validated = parseFloat(row.parcelles_validees_par_lurm || row.parcelles_validees || row.validated || row.urm_validated_parcels || 0);
        const collected = parseFloat(row.collected_parcels_no_duplicates || row.collected || 0);
        
        if (collected === 0) return false;
        return (validated / collected) >= 0.8; // 80% completion threshold
    }
    
    /**
     * Sets up quick action buttons
     */
    setupQuickActions() {
        try {
            const generateReportBtn = document.getElementById('generateReportBtn');
            const exportDataBtn = document.getElementById('exportDataBtn');
            const syncDataBtn = document.getElementById('syncDataBtn');
            const viewDetailBtn = document.getElementById('viewDetailBtn');
            
            if (generateReportBtn) {
                generateReportBtn.addEventListener('click', () => {
                    this.generateReport();
                });
            }
            
            if (exportDataBtn) {
                exportDataBtn.addEventListener('click', () => {
                    this.exportData();
                });
            }
            
            if (syncDataBtn) {
                syncDataBtn.addEventListener('click', () => {
                    this.refreshDashboard(true);
                });
            }
            
            if (viewDetailBtn) {
                viewDetailBtn.addEventListener('click', () => {
                    // Switch to commune details tab
                    const communesTab = document.querySelector('[data-tab="geographic"]');
                    if (communesTab) communesTab.click();
                });
            }
        } catch (error) {
            console.error('Failed to setup quick actions:', error);
            showNotification('Failed to setup quick actions', 'error');
        }
    }
    
    /**
     * Initializes collection analysis tab
     */
    async initializeCollectionTab() {
        try {
            const collectionData = dashboardState.getData('data_collection') || [];
            const projectionData = dashboardState.getData('projection_collection') || [];
            const yieldsData = dashboardState.getData('yields_projection') || [];
            
            // Create collection breakdown chart
            chartManager.createCollectionPhaseChart('collectionBreakdown', collectionData);
            
            // Create tool utilization chart
            chartManager.createToolUtilizationChart('toolUtilization', collectionData);
            
            // Create collection trends chart
            this.createCollectionTrendsChart(collectionData, projectionData);
            
            // Update productivity metrics
            this.updateProductivityMetrics(yieldsData);
            
            // Update dominant phase insight
            this.updateDominantPhase(collectionData);
        } catch (error) {
            console.error('Failed to initialize collection tab:', error);
            showNotification('Failed to load collection analysis', 'error');
        }
    }

    /**
     * Initializes commune-status tab
     */
    async initializeCommuneStatusTab() {
        try {
            const communeData = dashboardState.getData('commune_status') || [];

            // Create a simple status distribution chart for communes
            const ctx = document.getElementById('communeStatusChart')?.getContext('2d');
            if (ctx) {
                // compute counts
                const counts = { completed: 0, processing: 0, pending: 0, issues: 0 };
                communeData.forEach(row => {
                    if (!row.commune || row.commune.toLowerCase() === 'total') return;
                    const status = this.getStatus(row).text;
                    if (status === 'Terminé') counts.completed++;
                    else if (status === 'En cours') counts.processing++;
                    else if (status === 'En attente') counts.pending++;
                    else counts.issues++;
                });

                try { dashboardState.ensureCanvasFree(ctx && ctx.canvas ? ctx.canvas : ctx, 'communeStatusChart'); } catch (e) { /* ignore */ }
                const chart = new Chart(ctx, {
                    type: 'doughnut',
                    data: {
                        labels: ['Terminé', 'En cours', 'En attente', 'Problèmes'],
                        datasets: [{ data: [counts.completed, counts.processing, counts.pending, counts.issues],
                            backgroundColor: ['#10b981', '#3b82f6', '#f59e0b', '#ef4444'] }]
                    },
                    options: { responsive: true, maintainAspectRatio: false }
                });
                dashboardState.setChart('communeStatusChart', chart);
                // Render legend if container exists
                const legendContainer = document.getElementById('communeStatusLegend');
                if (legendContainer) {
                    try {
                        const labels = chart.data.labels || [];
                        const colors = chart.data.datasets?.[0]?.backgroundColor || [];
                        legendContainer.innerHTML = labels.map((lab, idx) => `
                            <div class="legend-item" style="display:inline-flex;align-items:center;margin-right:12px;">
                                <span style="display:inline-block;width:12px;height:12px;background:${colors[idx]||'#ccc'};margin-right:6px;border-radius:2px;"></span>
                                <small>${lab}</small>
                            </div>
                        `).join('');
                    } catch (e) {
                        // ignore legend errors
                    }
                }
            }

            // Create a detailed top communes bar chart
            const ctx2 = document.getElementById('communesChartDetailed')?.getContext('2d');
            if (ctx2) {
                const top = communeData.filter(c => c.commune && c.commune.toLowerCase() !== 'total')
                    .sort((a,b) => (b.collected || 0) - (a.collected || 0)).slice(0,10);

                if (dashboardState.getChart('communesChartDetailed')) dashboardState.getChart('communesChartDetailed').destroy();
                const chart2 = new Chart(ctx2, {
                    type: 'bar',
                    data: { labels: top.map(t => t.commune), datasets: [{ label: 'Parcelles collectées', data: top.map(t => t.collected || 0), backgroundColor: '#0072BC' }] },
                    options: { responsive: true, maintainAspectRatio: false }
                });
                dashboardState.setChart('communesChartDetailed', chart2);
            }

            // Populate detailed list
            // Removed rendering of the left-hand detailed communes list per user request.
            // this.populateCommunesListDetailed(communeData);

            // Render geomatician selector and KPIs
            this.renderGeomaticianControls(communeData);

            // Toggle column visibility panel
            const toggleBtn = document.getElementById('toggleColumnsBtn');
            if (toggleBtn) toggleBtn.addEventListener('click', () => {
                const panel = document.getElementById('columnVisibility');
                if (panel) panel.style.display = panel.style.display === 'none' ? 'block' : 'none';
            });

            // Export button
            const exportBtn = document.getElementById('exportExcelBtn');
            if (exportBtn) exportBtn.addEventListener('click', () => this.exportCommuneTableToXLSX());
        } catch (error) {
            console.error('Failed to initialize commune-status tab:', error);
            showNotification('Failed to load Statut Communes', 'error');
        }
    }

    /**
     * Populates detailed commune list for commune-status tab
     * @param {Array} communeData
     */
    populateCommunesListDetailed(communeData) {
        try {
            const container = document.getElementById('communesListDetailed');
            if (!container) return;
            const items = (communeData || []).filter(c => c.commune && c.commune.toLowerCase() !== 'total')
                .map(c => {
                    const status = this.getStatus(c);
                    const validationRate = c.collected ? ((c.validated || 0) / c.collected * 100).toFixed(1) : '0.0';
                    return `
                        <div class="activity-item">
                            <div>
                                <strong>${sanitizeHTML(c.commune)}</strong>
                                <p class="text-sm">Collectées: ${Number(c.collected||0).toLocaleString()} • Validées: ${Number(c.validated||0).toLocaleString()}</p>
                            </div>
                            <div class="status-block">
                                <span class="status-badge ${status.class}">${status.text}</span>
                                <span class="progress-percentage">${validationRate}%</span>
                            </div>
                        </div>
                    `;
                });
            container.innerHTML = items.join('') || '<p>Aucune commune disponible</p>';
        } catch (error) {
            console.error('Failed to populate communesListDetailed:', error);
        }
    }
    
    /* === Commune table, KPIs and export === */
    renderGeomaticianControls(communeData) {
        try {
            const select = document.getElementById('geomaticianFilter');
            if (!select) return;
            const geomaticians = Array.from(new Set((communeData||[]).map(c => this.getGeomaticienName(c)).filter(Boolean)));
            // Clear and populate
            select.innerHTML = '<option value="all">Tous</option>' + geomaticians.map(g => `<option value="${sanitizeHTML(g)}">${sanitizeHTML(g)}</option>`).join('');

            select.addEventListener('change', () => this.updateGeomaticianKPIsAndTable(select.value));
            // initial render
            this.updateGeomaticianKPIsAndTable('all');
        } catch (e) { console.error('renderGeomaticianControls error', e); }
    }

    updateGeomaticianKPIsAndTable(geomatician) {
        // prefer dashboardState.getData; fall back to any global stored sheet or empty
        let data = [];
        if (typeof dashboardState.getData === 'function') data = dashboardState.getData('commune_status') || [];
        else if (typeof dashboardState.getSheet === 'function') data = dashboardState.getSheet('commune_status') || [];
        else data = window.__commune_status_backup || [];
        const filtered = (data || []).filter(r => {
            if (!r.commune || r.commune.toLowerCase() === 'total') return false;
            if (geomatician === 'all') return true;
            return String(this.getGeomaticienName(r) || '').toLowerCase() === String(geomatician || '').toLowerCase();
        });

        // KPIs: total communes assigned, total collected, validated %, avg validation rate
        const totalCommunes = filtered.length;
    // Sum using French-first fallbacks for robustness across sheet header variants
    const totalCollected = filtered.reduce((s, r) => s + (Number(r.collected_parcels_no_duplicates || r.collected || 0) || 0), 0);
    const totalValidated = filtered.reduce((s, r) => s + (Number(r.parcelles_validees_par_lurm || r.parcelles_validees || r.validated || r.urm_validated_parcels || 0) || 0), 0);
        const avgValidationRate = totalCollected ? ((totalValidated / totalCollected) * 100) : 0;

        const kpiContainer = document.getElementById('geomaticianKPIs');
        if (kpiContainer) {
            kpiContainer.innerHTML = `
                <div class="kpi-card small">
                    <div class="kpi-content"><h4>${totalCommunes}</h4><p>Communes assignées</p></div>
                </div>
                <div class="kpi-card small">
                    <div class="kpi-content"><h4>${Number(totalCollected).toLocaleString()}</h4><p>Parcelles collectées</p></div>
                </div>
                <div class="kpi-card small">
                    <div class="kpi-content"><h4>${avgValidationRate.toFixed(1)}%</h4><p>Taux validation moyen</p></div>
                </div>
            `;
        }

        // Render table
        this.renderCommuneTable(filtered);
        // Render commune KPI cards
        this.renderCommuneCards(filtered);
    }

    renderCommuneCards(communes) {
        try {
            const container = document.getElementById('communeCards');
            if (!container) return;
            // Build cards for top N communes
            const top = (communes || []).slice(0, 12);
            const cards = top.map(c => {
                const status = this.getStatus(c);
                const collected = Number(c.collected || 0);
                const validated = Number(c.validated || 0);
                const nonJoined = Number(c.unjoined || c.non_joined_parcels || c.non_joined || 0);
                const duplicates = Number(c.duplicatesRemoved || c.doublons || c.doublons_removed || 0);
                const conflict = Number(c.parcelsInConflict || c.parcelles_en_conflit || 0);
                const validationPct = collected ? Math.min(100, (validated / collected) * 100) : 0;
                return `
                    <div class="commune-card ${status.class}">
                        <div class="card-top">
                            <h4 class="commune-name">${sanitizeHTML(c.commune || '—')}</h4>
                            <span class="status-badge ${status.class}">${sanitizeHTML(status.text)}</span>
                        </div>
                        <div class="card-body">
                            <div class="metrics-row">
                                <div class="metric"><small>Collectées</small><strong>${Number(collected).toLocaleString()}</strong></div>
                                <div class="metric"><small>Validées</small><strong>${Number(validated).toLocaleString()}</strong></div>
                                <div class="metric"><small>Non jointes</small><strong>${Number(nonJoined).toLocaleString()}</strong></div>
                            </div>
                            <div class="metrics-row">
                                <div class="metric"><small>Géomaticien</small><div class="geomaticien-compact">${sanitizeHTML(this.getGeomaticienName(c) || '')}</div></div>
                                <div class="metric"><small>Doublons</small><strong>${(Number(duplicates) || 0) ? (Number(duplicates).toLocaleString()) : '—'}</strong></div>
                                <div class="metric"><small>Conflits</small><strong>${Number(conflict).toLocaleString()}</strong></div>
                            </div>
                            <div class="validation-row">
                                <div class="validation-bar"><div class="validation-fill" style="width:${validationPct}% ; background:${validationPct>80? 'var(--success-color)': (validationPct>50? 'var(--warning-color)':'var(--error-color)')};"></div></div>
                                <div class="validation-label">Taux de validation ${validationPct.toFixed(1)}%</div>
                            </div>
                        </div>
                    </div>
                `;
            }).join('');
            container.innerHTML = cards || '<p>Aucune commune affichée</p>';
        } catch (e) { console.error('renderCommuneCards error', e); }
    }

    renderCommuneTable(dataRows) {
        try {
            const table = document.getElementById('communeStatusTable');
            if (!table) return;
            // Determine all keys present in the data
            const rows = Array.isArray(dataRows) ? dataRows : [];
            const keySet = new Set();
            rows.forEach(r => Object.keys(r || {}).forEach(k => keySet.add(k)));

            // Normalize keys to collapse duplicates (english/french variants, cleaned versions)
            const normalize = k => (String(k||'')).toLowerCase().replace(/[^a-z0-9]/g, '');
            const groups = {}; // norm -> [origKeys]
            Array.from(keySet).forEach(k => {
                const n = normalize(k);
                groups[n] = groups[n] || [];
                groups[n].push(k);
            });

            // Preferred canonical mapping for some common fields
            const canonicalPreferred = ['commune','region','geomaticien','collected','validated','retained','rejected','duplicatesremoved','duplicateremovalrate','parcelsinconflict','parcellesbrutes'];

            // Create chosenKeys in preferred order then others
            const chosenKeys = [];
            // add preferred if present
            canonicalPreferred.forEach(pref => {
                if (groups[pref]) {
                    // choose best original key: prefer exact match, else shortest
                    const options = groups[pref];
                    const pick = options.find(o => o.toLowerCase() === pref) || options.sort((a,b)=>a.length-b.length)[0];
                    chosenKeys.push({ norm: pref, key: pick, alternates: options });
                    delete groups[pref];
                }
            });
            // Remaining groups: pick a representative key
            Object.keys(groups).sort().forEach(norm => {
                const options = groups[norm];
                const pick = options.sort((a,b)=>a.length-b.length)[0];
                chosenKeys.push({ norm, key: pick, alternates: options });
            });

            // Load saved order and reconcile
            const saved = localStorage.getItem('commune_table_order');
            let ordered = chosenKeys.map(c => c.key);
            if (saved) {
                try {
                    const order = JSON.parse(saved);
                    // Keep saved order only for keys that exist now, append missing ones preserving ordered
                    const fromSaved = order.filter(k => ordered.includes(k));
                    const missing = ordered.filter(k => !fromSaved.includes(k));
                    ordered = fromSaved.concat(missing);
                } catch (e) { /* ignore parse errors */ }
            }

            // Build columns array preserving alternates for value lookup
            const humanize = (s) => (s || '').replace(/_/g,' ').replace(/\b\w/g, c => c.toUpperCase());
            const columnLabelOverrides = {
                // French labels for canonical keys (snake_case and camelCase variants)
                'geomaticien': 'Géomaticien', 'Geomaticien': 'Géomaticien',
                'collected': 'Parcelles collectées', 'collected_parcels_no_duplicates': 'Parcelles collectées',
                'validated': 'Parcelles validées', 'urm_validated_parcels': 'Parcelles validées',
                'parcelles_validees_par_lurm': 'Parcelles validées par l’URM',
                'rejected': 'Parcelles rejetées', 'urm_rejected_parcels': 'Parcelles rejetées',
                'rejectionReasons': 'Motifs de rejet post-traitement', 'motifs_de_rejet_post_traitement': 'Motifs de rejet post-traitement',
                'urmRejectionReasons': 'Motifs de rejet URM', 'motifs_de_rejet_urm': 'Motifs de rejet URM',
                'duplicateRemovalRate': "Taux suppression doublons (%)", 'taux_suppression_doublons': "Taux suppression doublons (%)",
                'duplicatesRemoved': 'Doublons supprimés', 'doublons_supprimes': 'Doublons supprimés',
                'parcelsInConflict': 'Parcelles en conflit', 'parcelles_en_conflit': 'Parcelles en conflit',
                'parcelles_brutes': 'Parcelles brutes', 'raw_parcels': 'Parcelles brutes',
                'nicad': 'NICAD', 'percentNicad': '% NICAD', 'percent_nicad': '% NICAD',
                'ctasf': 'CTASF', 'percentCtasf': '% CTASF', 'percent_ctasf': '% CTASF',
                'total_parcels': 'Total Parcelles', 'total_parcelles': 'Total Parcelles', 'percentTotal': '% du total', 'percent_total': '% du total',
                'deliberated': 'Parcelles délibérées', 'deliberees': 'Parcelles délibérées', 'percentDeliberated': '% Délibérées',
                'retained': 'Parcelles retenues après post-traitement', 'parcelles_retenues': 'Parcelles retenues',
                'corrected': 'Parcelles corrigées', 'parcelles_corrigees': 'Parcelles corrigées',
                'individualJoined': 'Parcelles individuelles jointes', 'collectiveJoined': 'Parcelles collectives jointes',
                'unjoined': 'Parcelles non jointes', 'non_joined_parcels': 'Parcelles non jointes',
                'significantDuplicates': 'Doublons significatifs', 'SignificantDuplicates': 'Doublons significatifs',
                'post_processed_parcels': 'Parcelles post-traitées lot 1-46', 'parcelles_post_traitees_lot_1_46': 'Parcelles post-traitées lot 1-46',
                'joinStatus': 'Statut jointure', 'joinErrorMessage': 'Message d’erreur jointure',
                'parcelles_post_traitement': 'Parcelles post-traitées'
            };
            const columns = ordered.map(k => {
                // find the group for this key
                const norm = normalize(k);
                const group = (Array.from(keySet).filter(orig => normalize(orig) === norm));
                // Prefer normalized override, then exact key override, then humanized fallback
                const label = (columnLabelOverrides[norm] || columnLabelOverrides[k] || humanize(k));
                return { key: k, alternates: group, label };
            });

            // Build header with drag handles and visibility checkboxes
            const thead = document.createElement('thead');
            const tr = document.createElement('tr');
            tr.classList.add('draggable-row');
            columns.forEach(col => {
                const th = document.createElement('th');
                th.setAttribute('draggable', 'true');
                th.dataset.key = col.key;
                th.innerHTML = `<span class="drag-handle" title="Glisser pour réordonner">☰</span> ${sanitizeHTML(col.label)}`;
                tr.appendChild(th);
            });
            thead.appendChild(tr);

            // Build body
            const tbody = document.createElement('tbody');
            const getValueForColumn = (row, col) => {
                // Try primary key first, then alternates
                const primary = row[col.key];
                if (primary !== undefined && primary !== null && String(primary).trim() !== '') return primary;
                for (const alt of (col.alternates || [])) {
                    const v = row[alt]; if (v !== undefined && v !== null && String(v).trim() !== '') return v;
                }
                return '';
            };

            (rows||[]).forEach(row => {
                const r = document.createElement('tr');
                columns.forEach(col => {
                    const td = document.createElement('td');
                    const raw = getValueForColumn(row, col);
                    // Special rendering for commune
                    if (col.key === 'commune') {
                        const communeName = sanitizeHTML(String(raw || ''));
                        const region = sanitizeHTML(String(getValueForColumn(row, { key: 'region', alternates: [ 'region' ] }) || row.region || ''));
                        td.classList.add('commune-cell');
                        td.innerHTML = `<strong>${communeName}</strong>${region ? `<div class="muted">${region}</div>` : ''}`;
                    }
                    // Special rendering for geomaticien (detect by normalized key)
                    else if (normalize(col.key) === 'geomaticien') {
                        const name = String(this.getGeomaticienName(row) || '').trim();
                        const display = sanitizeHTML(name || '—');
                        const color = this.generateColorFromString(name || '');
                        td.innerHTML = `<span class="geomaticien-badge" style="background:${color};">${sanitizeHTML((name||'').split(' ').map(p=>p[0]||'').join('').toUpperCase() || '')}</span> <span class="geomaticien-name">${display}</span>`;
                    }
                    else {
                        if (typeof raw === 'number') td.textContent = Number(raw).toLocaleString();
                        else td.textContent = sanitizeHTML(String(raw || ''));
                    }
                    r.appendChild(td);
                });
                tbody.appendChild(r);
            });

            // Replace table content
            table.innerHTML = '';
            table.appendChild(thead);
            table.appendChild(tbody);

            // Column visibility list
            this.buildColumnVisibility(columns);

            // Attach drag handlers for header reordering
            this.enableHeaderDrag(table, columns.map(c=>c.key));
        } catch (e) { console.error('renderCommuneTable error', e); }
    }

    buildColumnVisibility(columns) {
        try {
            const list = document.getElementById('columnList');
            if (!list) return;
            // Load saved visibility
            const savedVis = localStorage.getItem('commune_table_visibility');
            let visMap = {};
            if (savedVis) {
                try { visMap = JSON.parse(savedVis); } catch(e) { visMap = {}; }
            }

            list.innerHTML = columns.map(c => {
                const isChecked = (visMap[c.key] === undefined) ? true : !!visMap[c.key];
                return `
                <label style="display:block;margin:4px 0;">
                    <input type="checkbox" data-col="${sanitizeHTML(c.key)}" ${isChecked ? 'checked' : ''} /> ${sanitizeHTML(c.label)}
                </label>
            `}).join('');

            // toggle visible columns and persist
            list.querySelectorAll('input[type=checkbox]').forEach(cb => cb.addEventListener('change', (e) => {
                const key = cb.dataset.col;
                const table = document.getElementById('communeStatusTable');
                if (!table) return;
                const idx = Array.from(table.querySelectorAll('thead th')).findIndex(th => th.dataset.key === key);
                if (idx === -1) return;
                const show = cb.checked;
                // Toggle header
                const th = table.querySelectorAll('thead th')[idx];
                if (th) th.style.display = show ? '' : 'none';
                // Toggle cells
                table.querySelectorAll('tbody tr').forEach(tr => {
                    const td = tr.children[idx]; if (td) td.style.display = show ? '' : 'none';
                });
                // persist
                const current = JSON.parse(localStorage.getItem('commune_table_visibility') || '{}');
                current[key] = show;
                localStorage.setItem('commune_table_visibility', JSON.stringify(current));
            }));
        } catch (e) { console.error('buildColumnVisibility error', e); }
    }

    enableHeaderDrag(table, columnKeys) {
        try {
            const headers = Array.from(table.querySelectorAll('thead th'));
            let dragSrcIndex = null;
            headers.forEach((th, idx) => {
                th.addEventListener('dragstart', (e) => { dragSrcIndex = idx; th.classList.add('dragging'); e.dataTransfer.effectAllowed = 'move'; });
                th.addEventListener('dragend', () => { headers.forEach(h=>h.classList.remove('dragging')); });
                th.addEventListener('dragover', (e) => { e.preventDefault(); e.dataTransfer.dropEffect = 'move'; });
                th.addEventListener('drop', (e) => {
                    e.preventDefault();
                    const target = th;
                    const targetIndex = Array.from(headers).indexOf(target);
                    if (dragSrcIndex === null || dragSrcIndex === targetIndex) return;
                    // reorder header cells
                    const thead = table.querySelector('thead tr');
                    const ths = Array.from(thead.children);
                    const dragged = ths[dragSrcIndex];
                    thead.removeChild(dragged);
                    thead.insertBefore(dragged, ths[targetIndex]);
                    // reorder body cells for each row
                    table.querySelectorAll('tbody tr').forEach(row => {
                        const cells = Array.from(row.children);
                        const cellDragged = cells[dragSrcIndex];
                        row.removeChild(cellDragged);
                        row.insertBefore(cellDragged, cells[targetIndex]);
                    });
                    // persist order
                    const newOrder = Array.from(table.querySelectorAll('thead th')).map(h => h.dataset.key);
                    localStorage.setItem('commune_table_order', JSON.stringify(newOrder));
                });
            });
        } catch (e) { console.error('enableHeaderDrag error', e); }
    }

    exportCommuneTableToXLSX() {
        try {
            const table = document.getElementById('communeStatusTable');
            if (!table) return;
            // Build array of arrays for XLSX
            const rows = [];
            // Only export visible columns, in the order of the current headers
            const headerThs = Array.from(table.querySelectorAll('thead th'));
            const visibleIndexes = headerThs.map((th, idx) => ({th, idx})).filter(x => (x.th.style.display || '') !== 'none');
            const headers = visibleIndexes.map(x => x.th.textContent.trim());
            rows.push(headers);
            table.querySelectorAll('tbody tr').forEach(tr => {
                const row = visibleIndexes.map(x => {
                    const td = tr.children[x.idx]; return td ? td.textContent.trim() : '';
                });
                rows.push(row);
            });

            const ws = XLSX.utils.aoa_to_sheet(rows);
            const wb = XLSX.utils.book_new();
            XLSX.utils.book_append_sheet(wb, ws, 'Communes');
            const wbout = XLSX.write(wb, {bookType:'xlsx', type:'array'});
            const blob = new Blob([wbout], {type:'application/octet-stream'});
            const url = URL.createObjectURL(blob);
            const a = document.createElement('a');
            a.href = url; a.download = 'commune_status.xlsx'; document.body.appendChild(a); a.click(); a.remove(); URL.revokeObjectURL(url);
        } catch (e) { console.error('exportCommuneTableToXLSX', e); }
    }

    // Generate a deterministic pastel color from a string (for badge backgrounds)
    generateColorFromString(str) {
        let hash = 0;
        for (let i = 0; i < str.length; i++) hash = str.charCodeAt(i) + ((hash << 5) - hash);
        const h = Math.abs(hash) % 360;
        return `hsl(${h} 65% 70%)`;
    }

    // Try to resolve a numeric value from a row using multiple candidate keys
    getNumericValue(row, candidates) {
        if (!row || typeof row !== 'object') return 0;
        const tryParse = (v) => {
            if (v === null || v === undefined || v === '') return null;
            if (typeof v === 'number') return v;
            const s = String(v).trim().replace(/\s/g, '');
            // replace comma decimal
            if (/^-?\d+[\.,]\d+$/.test(s)) return parseFloat(s.replace(',', '.'));
            // thousands with commas
            if (/^-?\d{1,3}(,\d{3})*(\.\d+)?$/.test(s)) return parseFloat(s.replace(/,/g, ''));
            // plain integer/float
            if (/^-?\d+(\.\d+)?$/.test(s)) return parseFloat(s);
            return null;
        };

        for (const k of candidates) {
            if (row[k] !== undefined) {
                const p = tryParse(row[k]);
                if (p !== null && !isNaN(p)) {
                    if (CONFIG.UI.debugFieldResolution) console.debug(`[field-resolve] numeric from key='${k}' value='${row[k]}' parsed=${p}`);
                    return p;
                }
            }
            // try lower/upper variants
            const lk = k.toLowerCase();
            if (row[lk] !== undefined) {
                const p = tryParse(row[lk]);
                if (p !== null && !isNaN(p)) {
                    if (CONFIG.UI.debugFieldResolution) console.debug(`[field-resolve] numeric from key='${lk}' value='${row[lk]}' parsed=${p}`);
                    return p;
                }
            }
        }
        // If specific candidates were provided we do NOT perform a loose scan across all fields
        // (loose scanning previously caused incorrect matches, e.g. picking 'collected' when 'validated' was intended).
        // Only perform a broad scan when no candidates were given.
        if (!candidates || !candidates.length) {
            for (const k of Object.keys(row)) {
                const p = tryParse(row[k]);
                if (p !== null && !isNaN(p)) return p;
            }
        }
        return 0;
    }

    // Heuristic: pick the geomaticien/person name from a parsed row
    getGeomaticienName(row) {
        if (!row || typeof row !== 'object') return '';
        // canonical candidate keys (lowercased)
        const candidateKeys = ['geomaticien','geomatician','geomaticien','geomaticien_name','geomaticien_nom','geomaticienprenom','geomaticien_prenom','geomatician_name','geomatician_nom','geomaticienprenom'];

        const isName = (s) => {
            if (!s) return false;
            const t = String(s).trim();
            if (t.length < 2) return false;
            // must contain at least one letter, and not be purely numeric (allow commas/dots)
            if (!/[A-Za-zÀ-ÖØ-öø-ÿ]/.test(t)) return false;
            if (/^[-+]?\d+[\d\s,\.]*$/.test(t)) return false;
            return true;
        };

        // Prefer explicit candidate keys (try exact and lowercased variants)
        for (const k of Object.keys(row)) {
            const ln = String(k || '').toLowerCase();
            if (candidateKeys.includes(ln)) {
                const val = row[k];
                if (isName(val)) return String(val).trim();
            }
        }

        // Try direct canonical access as well
        for (const ck of candidateKeys) {
            if (row[ck] !== undefined && isName(row[ck])) {
                if (CONFIG.UI.debugFieldResolution) console.debug(`[field-resolve] geomaticien from key='${ck}' value='${row[ck]}'`);
                return String(row[ck]).trim();
            }
        }

        // Avoid returning commune/region values; scan remaining fields but skip noisy ones
        const noisyKeyPattern = /commune|region|total|parc|collec|nicad|ctasf|percent|duplicate|joined|validated|rejected|raw|corrected|deliberat|rejection|duplicates_removed|duplicate_removal_rate|individual_joined/i;
        for (const k of Object.keys(row)) {
            const ln = String(k || '').toLowerCase();
            if (noisyKeyPattern.test(ln)) continue;
            const val = String(row[k] || '').trim();
            if (isName(val)) {
                // avoid returning commune name
                const commune = String(row.commune || row.Commune || '').trim();
                if (commune && val.toLowerCase() === commune.toLowerCase()) continue;
                if (CONFIG.UI.debugFieldResolution) console.debug(`[field-resolve] geomaticien fallback key='${k}' value='${val}'`);
                return val;
            }
        }

        return '';
    }

    
    
    /**
     * Creates collection trends chart
     * @param {Array} collectionData - Collection data
     * @param {Array} projectionData - Projection data
     */
    createCollectionTrendsChart(collectionData, projectionData) {
        const canvas = document.getElementById('collectionTrends');
        if (!canvas) {
            console.warn('Canvas not found: collectionTrends');
            return;
        }
        
        const ctx = canvas.getContext('2d');
        
        // TODO: Replace mock data with real data processing
        const days = Array.from({ length: 30 }, (_, i) => {
            const date = new Date();
            date.setDate(date.getDate() - (29 - i));
            return date.toLocaleDateString('fr-FR', { month: 'short', day: 'numeric' });
        });
        
        const trendData = collectionData.map(row => parseFloat(row.total || 0)) // Replace with actual data processing
            .slice(0, 30)
            .concat(Array(30).fill(0).slice(collectionData.length)); // Pad with zeros if needed
        
        const chart = new Chart(ctx, {
            type: 'line',
            data: {
                labels: days,
                datasets: [{
                    label: 'Parcelles collectées/jour',
                    data: trendData,
                    borderColor: CONFIG.COLORS.primary,
                    backgroundColor: CONFIG.COLORS.primary + '20',
                    fill: true,
                    tension: 0.3
                }]
            },
            options: {
                ...chartManager.defaultOptions,
                scales: {
                    y: {
                        beginAtZero: true,
                        title: {
                            display: true,
                            text: 'Parcelles'
                        }
                    },
                    x: {
                        title: {
                            display: true,
                            text: 'Date'
                        }
                    }
                }
            }
        });
        
        dashboardState.setChart('collectionTrends', chart);
    }
    
    /**
     * Updates productivity metrics
     * @param {Array} yieldsData - Yields data
     */
    updateProductivityMetrics(yieldsData) {
        try {
            if (!yieldsData.length) {
                showNotification('No yields data available', 'warning');
                return;
            }
            
            // Calculate average productivity
            const totalYield = yieldsData.reduce((sum, row) => {
                return sum + (parseFloat(row.champs_equipe_jour || row['Champs/Equipe/Jour'] || 0));
            }, 0);
            
            const avgProductivity = (totalYield / yieldsData.length).toFixed(1);
            
            // Find best performing team
            let bestTeam = '';
            let bestYield = 0;
            
            yieldsData.forEach(row => {
                const teamYield = parseFloat(row.champs_equipe_jour || row['Champs/Equipe/Jour'] || 0);
                if (teamYield > bestYield) {
                    bestYield = teamYield;
                    bestTeam = row.team || row.Team || 'Unknown';
                }
            });
            
            // Update UI
            const avgProductivityElement = document.getElementById('avgTeamProductivity');
            const avgProcessingTimeElement = document.getElementById('avgProcessingTime');
            const bestPerformingTeamElement = document.getElementById('bestPerformingTeam');
            
            if (avgProductivityElement) {
                avgProductivityElement.textContent = avgProductivity;
            }
            
            if (avgProcessingTimeElement) {
                // TODO: Replace mock data with real processing time calculation
                avgProcessingTimeElement.textContent = '2.3h';
            }
            
            if (bestPerformingTeamElement) {
                bestPerformingTeamElement.textContent = sanitizeHTML(bestTeam);
            }
        } catch (error) {
            console.error('Failed to update productivity metrics:', error);
            showNotification('Failed to update productivity metrics', 'error');
        }
    }
    
    /**
     * Updates dominant phase insight
     * @param {Array} collectionData - Collection data
     */
    updateDominantPhase(collectionData) {
        try {
            if (!collectionData.length) {
                showNotification('No collection data available', 'warning');
                return;
            }
            
            const phaseMap = new Map();
            collectionData.forEach(row => {
                const phase = row.phase || row.Phase || 'Unknown';
                const total = parseFloat(row.total || row.Total || 0);
                phaseMap.set(phase, (phaseMap.get(phase) || 0) + total);
            });
            
            let dominantPhase = '';
            let maxTotal = 0;
            
            phaseMap.forEach((total, phase) => {
                if (total > maxTotal) {
                    maxTotal = total;
                    dominantPhase = phase;
                }
            });
            
            const dominantPhaseElement = document.getElementById('dominantPhase');
            if (dominantPhaseElement) {
                dominantPhaseElement.textContent = sanitizeHTML(dominantPhase || 'Aucune donnée');
            }
        } catch (error) {
            console.error('Failed to update dominant phase:', error);
            showNotification('Failed to update dominant phase', 'error');
        }
    }
    
    /**
     * Initializes geographic operations tab
     */
    async initializeGeographicTab() {
        try {
            const communeData = dashboardState.getData('commune_status') || [];
            const deploymentData = dashboardState.getData('team_deployment') || [];
            
            // Create commune heatmap
            this.createCommuneHeatmap(communeData);
            
            // Create deployment timeline
            this.createDeploymentTimeline(deploymentData);
            
            // Update status indicators
            this.updateStatusIndicators(communeData);
            
            // Generate operational alerts
            this.generateOperationalAlerts(communeData);
        } catch (error) {
            console.error('Failed to initialize geographic tab:', error);
            showNotification('Failed to load geographic operations', 'error');
        }
    }
    
    /**
     * Creates commune progress heatmap
     * @param {Array} communeData - Commune data
     */
    createCommuneHeatmap(communeData) {
        try {
            const heatmapContainer = document.getElementById('communeHeatmap');
            if (!heatmapContainer) {
                console.warn('Heatmap container not found');
                return;
            }
            
            // Clear existing heatmap
            heatmapContainer.innerHTML = '';
            
            communeData.forEach(row => {
                if (!row.commune || row.commune.toLowerCase() === 'total') return;
                
                const validated = parseFloat(row.parcelles_validees_par_lurm || row.parcelles_validees || row.validated || row.urm_validated_parcels || 0);
                const collected = parseFloat(row.collected_parcels_no_duplicates || row.collected || 0);
                const completion = collected > 0 ? (validated / collected) * 100 : 0;
                
                const cell = document.createElement('div');
                cell.className = 'heatmap-cell';
                cell.textContent = sanitizeHTML(row.commune);
                cell.style.backgroundColor = this.getHeatmapColor(completion);
                cell.setAttribute('data-tooltip', `${sanitizeHTML(row.commune)}: ${completion.toFixed(1)}% terminé`);
                
                // Add click handler for details
                cell.addEventListener('click', () => {
                    this.showCommuneDetails(row);
                });
                
                heatmapContainer.appendChild(cell);
            });
        } catch (error) {
            console.error('Failed to create commune heatmap:', error);
            showNotification('Failed to create commune heatmap', 'error');
        }
    }
    
    /**
     * Gets heatmap color based on completion percentage
     * @param {number} completion - Completion percentage
     * @returns {string} Color code
     */
    getHeatmapColor(completion) {
        if (completion >= 75) return CONFIG.COLORS.success;
        if (completion >= 50) return CONFIG.COLORS.info;
        if (completion >= 25) return CONFIG.COLORS.warning;
        return CONFIG.COLORS.error;
    }
    
    /**
     * Shows commune details modal
     * @param {Object} communeData - Commune data
     */
    showCommuneDetails(communeData) {
        try {
            const modal = createModal('Détails de la Commune', `
                <div class="commune-details">
                    <h4>${sanitizeHTML(communeData.commune)}</h4>
                    <div class="detail-grid">
                        <div class="detail-item">
                            <span class="label">Région:</span>
                            <span class="value">${sanitizeHTML(communeData.region || 'N/A')}</span>
                        </div>
                        <div class="detail-item">
                            <span class="label">Parcelles totales:</span>
                            <span class="value">${(parseFloat(communeData.total_parcels || 0)).toLocaleString()}</span>
                        </div>
                        <div class="detail-item">
                            <span class="label">Parcelles collectées:</span>
                            <span class="value">${(parseFloat(communeData.collected_parcels_no_duplicates || communeData.collected || 0)).toLocaleString()}</span>
                        </div>
                        <div class="detail-item">
                            <span class="label">Parcelles validées:</span>
                            <span class="value">${(parseFloat(communeData.urm_validated_parcels || communeData.validated || 0)).toLocaleString()}</span>
                        </div>
                        <div class="detail-item">
                            <span class="label">Géomaticien:</span>
                            <span class="value">${sanitizeHTML(this.getGeomaticienName(communeData) || communeData.geomaticien || 'N/A')}</span>
                        </div>
                    </div>
                </div>
            `);
            
            showModal(modal);
        } catch (error) {
            console.error('Failed to show commune details:', error);
            showNotification('Failed to show commune details', 'error');
        }
    }
    
    /**
     * Creates deployment timeline
     * @param {Array} deploymentData - Deployment data
     */
    createDeploymentTimeline(deploymentData) {
        try {
            const timelineContainer = document.getElementById('deploymentTimeline');
            if (!timelineContainer) {
                console.warn('Timeline container not found');
                return;
            }
            
            // Clear existing timeline
            timelineContainer.innerHTML = '';
            
            deploymentData.forEach((row, index) => {
                const timelineItem = document.createElement('div');
                timelineItem.className = 'timeline-item';
                
                const status = this.getDeploymentStatus(row);
                
                timelineItem.innerHTML = `
                    <div class="timeline-content">
                        <div class="timeline-header">
                            <span class="timeline-title">${sanitizeHTML(row.tambacounda || `Déploiement ${index + 1}`)}</span>
                            <span class="timeline-date">${sanitizeHTML(row.tambacounda_start_date || 'TBD')}</span>
                        </div>
                        <div class="timeline-description">
                            Statut: <strong>${sanitizeHTML(status)}</strong><br>
                            Région: Tambacounda<br>
                            Fin prévue: ${sanitizeHTML(row.tambacounda_end_date || 'TBD')}
                        </div>
                    </div>
                `;
                
                timelineContainer.appendChild(timelineItem);
            });
        } catch (error) {
            console.error('Failed to create deployment timeline:', error);
            showNotification('Failed to create deployment timeline', 'error');
        }
    }
    
    /**
     * Gets deployment status
     * @param {Object} deploymentRow - Deployment data row
     * @returns {string} Status
     */
    getDeploymentStatus(deploymentRow) {
        const status = deploymentRow.tambacounda_status || deploymentRow.kedougou_status || '';
        return status || 'En attente';
    }
    
    /**
     * Updates status indicators
     * @param {Array} communeData - Commune data
     */
    updateStatusIndicators(communeData) {
        try {
            const regions = {
                tambacounda: { completed: 0, processing: 0, pending: 0 },
                kedougou: { completed: 0, processing: 0, pending: 0 }
            };
            
            communeData.forEach(row => {
                if (!row.commune || row.commune.toLowerCase() === 'total') return;
                
                const region = (row.region || '').toLowerCase();
                const completion = this.getCompletionStatus(row);
                
                if (region.includes('tambacounda')) {
                    regions.tambacounda[completion]++;
                } else if (region.includes('kedougou') || region.includes('kédougou')) {
                    regions.kedougou[completion]++;
                }
            });
            
            // Update UI for each region
            Object.keys(regions).forEach(regionKey => {
                const region = regions[regionKey];
                
                ['completed', 'processing', 'pending'].forEach(status => {
                    const element = document.getElementById(`${regionKey}${status.charAt(0).toUpperCase() + status.slice(1)}`);
                    if (element) {
                        element.textContent = region[status];
                    }
                });
            });
        } catch (error) {
            console.error('Failed to update status indicators:', error);
            showNotification('Failed to update status indicators', 'error');
        }
    }

    /**
     * Determines a simple completion category for a commune row
     * @param {Object} row
     * @returns {'completed'|'processing'|'pending'}
     */
    getCompletionStatus(row) {
        try {
            if (!row) return 'pending';
            const collected = Number(row.collected_parcels_no_duplicates || row.collected || 0) || 0;
            const validated = Number(row.parcelles_validees_par_lurm || row.parcelles_validees || row.validated || row.urm_validated_parcels || 0) || 0;
            if (collected === 0) return 'pending';
            const rate = (validated / collected) * 100;
            if (rate >= 80) return 'completed';
            if (rate >= 50) return 'processing';
            return 'pending';
        } catch (e) {
            return 'pending';
        }
    }

    /**
     * Generates operational alerts (separate from quality alerts)
     * @param {Array} communeData
     */
    generateOperationalAlerts(communeData) {
        try {
            const alertsContainer = document.getElementById('operationalAlerts') || document.getElementById('geoAlerts') || document.getElementById('qualityAlerts');
            if (!alertsContainer) {
                console.warn('Operational alerts container not found (operationalAlerts/geoAlerts/qualityAlerts)');
                return;
            }

            const alerts = [];
            (communeData || []).forEach(row => {
                if (!row || !row.commune) return;
                if (String(row.commune).trim().toLowerCase() === 'total') return;

                const rejected = Number(row.urm_rejected_parcels || row.rejected || 0) || 0;
                const validated = Number(row.parcelles_validees_par_lurm || row.parcelles_validees || row.validated || row.urm_validated_parcels || 0) || 0;
                const collected = Number(row.collected_parcels_no_duplicates || row.collected || 0) || 0;

                const total = rejected + validated;
                if (total > 0 && (rejected / total) > 0.3) {
                    alerts.push({ type: 'warning', message: `${sanitizeHTML(row.commune)}: Taux de rejet élevé (${((rejected / total) * 100).toFixed(1)}%)` });
                }

                if (collected > 0 && validated > collected) {
                    alerts.push({ type: 'info', message: `${sanitizeHTML(row.commune)}: Incohérence détectée (validées > collectées)` });
                }

                // Low completion warning
                if (collected > 0) {
                    const rate = (validated / collected) * 100;
                    if (rate < 30) {
                        alerts.push({ type: 'danger', message: `${sanitizeHTML(row.commune)}: Taux de validation faible (${rate.toFixed(1)}%)` });
                    }
                }
            });

            alertsContainer.innerHTML = '';
            if (!alerts.length) {
                alertsContainer.innerHTML = '<div class="alert-item info"><i class="fas fa-check"></i><span>Aucune alerte opérationnelle</span></div>';
                return;
            }

            alerts.slice(0, 10).forEach(alert => {
                const el = document.createElement('div');
                el.className = `alert-item ${alert.type}`;
                el.innerHTML = `
                    <i class="fas fa-${alert.type === 'danger' ? 'exclamation-triangle' : alert.type === 'warning' ? 'exclamation-circle' : 'info-circle'}"></i>
                    <span>${sanitizeHTML(alert.message)}</span>
                `;
                alertsContainer.appendChild(el);
            });
        } catch (error) {
            console.error('Failed to generate operational alerts:', error);
        }
    }
    

    
    /**
     * Initializes projections tab
     */
    async initializeProjectionsTab() {
        try {
            const nicadData = dashboardState.getData('nicad_projection') || [];
            const ctasfData = dashboardState.getData('ctasf_projection') || [];
            const publicData = dashboardState.getData('public_display') || [];
            
            // Create multi-month projections chart
            chartManager.createProjectionsChart('multiMonthProjections', { nicadData, ctasfData, publicData });
            
            // Create NICAD timeline chart
            this.createNicadTimelineChart(nicadData);
            
            // Create CTASF forecast chart
            this.createCtasfForecastChart(ctasfData);
            
            // Update public display progress
            this.updatePublicDisplayProgress(publicData);
            // Also render a small monthly chart for public display
            this.createPublicDisplayChart(publicData);
            
            // Update projection metrics
            this.updateProjectionMetrics(nicadData, ctasfData);
        } catch (error) {
            console.error('Failed to initialize projections tab:', error);
            showNotification('Failed to load projections', 'error');
        }
    }
    
    /**
     * Creates NICAD assignment timeline chart
     * @param {Array} nicadData - NICAD projection data
     */
    createNicadTimelineChart(nicadData) {
        try {
            // prefer visible ID 'nicadTimeline', fallback to legacy 'nicadTimelineChart'
            var canvas = document.getElementById('nicadTimeline') || document.getElementById('nicadTimelineChart');
            if (!canvas) {
                console.warn('Canvas not found: nicadTimeline (or nicadTimelineChart)');
                return;
            }

            var ctx = canvas.getContext('2d');
            const chartData = chartManager.processProjectionData({ nicadData, ctasfData: [], publicData: [] });

            try { dashboardState.ensureCanvasFree(canvas, 'nicadTimeline'); } catch (e) { /* ignore */ }
            try { dashboardState.ensureCanvasFree(canvas, canvas && canvas.id ? canvas.id : null); } catch (e) { /* ignore */ }
            const chart = new Chart(ctx, {
                type: 'line',
                data: {
                    labels: chartData.months,
                    datasets: [{
                        label: 'NICAD Prévisions',
                        data: chartData.nicad,
                        borderColor: CONFIG.COLORS.success,
                        backgroundColor: CONFIG.COLORS.success + '20',
                        fill: true,
                        tension: 0.3
                    }]
                },
                options: {
                    ...chartManager.defaultOptions,
                    scales: {
                        y: {
                            beginAtZero: true,
                            title: {
                                display: true,
                                text: 'Nombre de parcelles'
                            }
                        },
                        x: {
                            title: {
                                display: true,
                                text: 'Mois'
                            }
                        }
                    }
                }
            });

            dashboardState.setChart('nicadTimeline', chart);
        } catch (error) {
            console.error('Failed to create NICAD timeline chart:', error);
            showNotification('Failed to create NICAD timeline chart', 'error');
        }
    }

    /**
     * Creates CTASF forecast chart
     * @param {Array} ctasfData - CTASF projection data
     */
    createCtasfForecastChart(ctasfData) {
        try {
            // prefer visible ID 'ctasfForecast', fallback to legacy 'ctasfForecastChart'
            var canvas = document.getElementById('ctasfForecast') || document.getElementById('ctasfForecastChart');
            if (!canvas) {
                console.warn('Canvas not found: ctasfForecast (or ctasfForecastChart)');
                return;
            }

            var ctx = canvas.getContext('2d');
            const chartData = chartManager.processProjectionData({ nicadData: [], ctasfData, publicData: [] });

            try { dashboardState.ensureCanvasFree(canvas, 'ctasfForecast'); } catch (e) { /* ignore */ }
            try { dashboardState.ensureCanvasFree(canvas, canvas && canvas.id ? canvas.id : null); } catch (e) { /* ignore */ }
            const chart = new Chart(ctx, {
                type: 'line',
                data: {
                    labels: chartData.months,
                    datasets: [{
                        label: 'CTASF Prévisions',
                        data: chartData.ctasf,
                        borderColor: CONFIG.COLORS.warning,
                        backgroundColor: CONFIG.COLORS.warning + '20',
                        fill: true,
                        tension: 0.3
                    }]
                },
                options: {
                    ...chartManager.defaultOptions,
                    scales: {
                        y: {
                            beginAtZero: true,
                            title: {
                                display: true,
                                text: 'Nombre de parcelles'
                            }
                        },
                        x: {
                            title: {
                                display: true,
                                text: 'Mois'
                            }
                        }
                    }
                }
            });

            dashboardState.setChart('ctasfForecast', chart);
        } catch (error) {
            console.error('Failed to create CTASF forecast chart:', error);
            showNotification('Failed to create CTASF forecast chart', 'error');
        }
    }

    /**
     * Creates a monthly chart for public display inside the Progression Affichage Public panel
     * @param {Array} publicData - Public display projection data
     */
    createPublicDisplayChart(publicData) {
        try {
            const container = document.getElementById('publicDisplayProgress');
            if (!container) return;

            // Ensure a canvas exists in the container
            let canvas = container.querySelector('#publicDisplayChart');
            if (!canvas) {
                canvas = document.createElement('canvas');
                canvas.id = 'publicDisplayChart';
                canvas.style.width = '100%';
                canvas.style.height = '160px';
                container.insertAdjacentElement('afterbegin', canvas);
            }

            const ctx = canvas.getContext('2d');
            if (!ctx) return;

            try { dashboardState.ensureCanvasFree(canvas, 'publicDisplayChart'); } catch (e) { /* ignore */ }
            
            const chartData = chartManager.processProjectionData({ nicadData: [], ctasfData: [], publicData: publicData || [] });

            try { dashboardState.ensureCanvasFree(canvas, canvas && canvas.id ? canvas.id : null); } catch (e) { /* ignore */ }
            const chart = new Chart(ctx, {
                type: 'line',
                data: { labels: chartData.months, datasets: [{ label: 'Affichage Public', data: chartData.public, borderColor: CONFIG.COLORS.info, backgroundColor: CONFIG.COLORS.info + '20', fill: true, tension: 0.3 }] },
                options: { responsive: true, maintainAspectRatio: false, scales: { y: { beginAtZero: true, title: { display: true, text: 'Nombre de parcelles' } }, x: { title: { display: true, text: 'Mois' } } }, plugins: { legend: { display: true } } }
            });

            dashboardState.setChart('publicDisplayChart', chart);
        } catch (error) {
            console.error('Failed to render public display chart:', error);
        }
    }

    /**
     * Updates public display progress
     * @param {Array} publicData - Public display data
     */
    updatePublicDisplayProgress(publicData) {
        try {
            // If explicit public display data is present, use it. Otherwise attempt a sensible fallback.
            let totalDisplayed = 0;
            let totalTarget = 0;

            if (publicData && publicData.length) {
                totalDisplayed = publicData.reduce((sum, row) => sum + (parseFloat(row.displayed_parcels || 0) || 0), 0);
                totalTarget = publicData.reduce((sum, row) => sum + (parseFloat(row.target_parcels || 0) || 0), 0);
            } else {
                // Fallback: derive a target from projection metrics if available
                const nicadProjection = dashboardState.getData('nicad_projection') || [];
                const ctasfProjection = dashboardState.getData('ctasf_projection') || [];
                const nicadTotal = nicadProjection.reduce((s, r) => s + Object.values(r).reduce((ss, v) => ss + (parseFloat(v) || 0), 0), 0);
                const ctasfTotal = ctasfProjection.reduce((s, r) => s + Object.values(r).reduce((ss, v) => ss + (parseFloat(v) || 0), 0), 0);
                // Assume displayed ~ min(nicad, ctasf) and target = max(nicad, ctasf) as a pragmatic proxy
                totalDisplayed = Math.min(nicadTotal, ctasfTotal) || 0;
                totalTarget = Math.max(nicadTotal, ctasfTotal) || 0;
            }

            const progressPercentage = totalTarget > 0 ? (totalDisplayed / totalTarget) * 100 : 0;

            const container = document.getElementById('publicDisplayProgress');
            if (container) {
                // Preserve any existing canvas (chart) and create/update a progress block after it
                let progressBlock = container.querySelector('.public-progress-block');
                const progressHTML = `
                    <div class="public-progress-bar">
                        <div class="public-progress-fill" style="width:${Math.min(progressPercentage,100)}%;background:${this.getProgressColor(progressPercentage)}"></div>
                    </div>
                    <div class="public-progress-meta"><strong id="publicDisplayPercentage">${progressPercentage.toFixed(1)}%</strong> affiché</div>
                `;

                if (progressBlock) {
                    progressBlock.innerHTML = progressHTML;
                } else {
                    progressBlock = document.createElement('div');
                    progressBlock.className = 'public-progress-block';
                    progressBlock.innerHTML = progressHTML;
                    container.appendChild(progressBlock);
                }

                // Animate the percentage text if animations are enabled
                const percentEl = progressBlock.querySelector('#publicDisplayPercentage');
                if (percentEl) this.animatePercentage(percentEl, progressPercentage);
            }
        } catch (error) {
            console.error('Failed to update public display progress:', error);
            showNotification('Failed to update public display progress', 'error');
        }
    }

    /**
     * Updates projection metrics
     * @param {Array} nicadData - NICAD projection data
     * @param {Array} ctasfData - CTASF projection data
     */
    updateProjectionMetrics(nicadData, ctasfData) {
        try {
            if (!nicadData.length || !ctasfData.length) {
                showNotification('Insufficient projection data available', 'warning');
                return;
            }

            // Calculate NICAD and CTASF totals
            const nicadTotal = nicadData.reduce((sum, row) => {
                return sum + Object.values(row).reduce((s, v) => s + (parseFloat(v) || 0), 0);
            }, 0);

            const ctasfTotal = ctasfData.reduce((sum, row) => {
                return sum + Object.values(row).reduce((s, v) => s + (parseFloat(v) || 0), 0);
            }, 0);

            // Update UI
            const nicadTotalElement = document.getElementById('nicadProjectionTotal');
            const ctasfTotalElement = document.getElementById('ctasfProjectionTotal');
            const varianceElement = document.getElementById('projectionVariance');

            if (nicadTotalElement) {
                this.animateNumber(nicadTotalElement, nicadTotal);
            }

            if (ctasfTotalElement) {
                this.animateNumber(ctasfTotalElement, ctasfTotal);
            }

            if (varianceElement) {
                const variance = Math.abs(nicadTotal - ctasfTotal) / ((nicadTotal + ctasfTotal) / 2) * 100;
                this.animatePercentage(varianceElement, variance);
            }
        } catch (error) {
            console.error('Failed to update projection metrics:', error);
            showNotification('Failed to update projection metrics', 'error');
        }
    }

    /**
     * Initializes quality control tab
     */
    async initializeQualityTab() {
        try {
            const communeData = dashboardState.getData('commune_status') || [];
            const collectionData = dashboardState.getData('data_collection') || [];

            // Create quality metrics charts
            this.createQualityMetricsCharts(communeData);

            // Create validation trends chart
            this.createValidationTrendsChart(collectionData);

            // Update quality KPIs
            this.updateQualityKPIs(communeData);

            // Generate quality alerts
            this.generateQualityAlerts(communeData);
        } catch (error) {
            console.error('Failed to initialize quality tab:', error);
            showNotification('Failed to load quality control', 'error');
        }
    }

    /**
     * Creates quality metrics charts
     * @param {Array} communeData - Commune data
     */
    createQualityMetricsCharts(communeData) {
        try {
            if (!communeData.length) {
                showNotification('No commune data for quality metrics', 'warning');
                return;
            }

            // Use aggregated totals so gauges match the KPI totals
            const totalValidated = communeData.reduce((sum, row) => sum + (parseFloat(row.parcelles_validees_par_lurm || row.parcelles_validees || row.validated || row.urm_validated_parcels || 0) || 0), 0);
            const totalRejected = communeData.reduce((sum, row) => sum + (parseFloat(row.urm_rejected_parcels || row.rejected || 0) || 0), 0);
            const totalCollected = communeData.reduce((sum, row) => sum + (parseFloat(row.collected_parcels_no_duplicates || row.collected || 0) || 0), 0);

            const validationRate = totalCollected > 0 ? (totalValidated / totalCollected) * 100 : 0;
            const rejectionRate = (totalValidated + totalRejected) > 0 ? (totalRejected / (totalValidated + totalRejected)) * 100 : 0;

            const totalDuplicatesRemoved = communeData.reduce((sum, row) => sum + (Number(row.duplicatesRemoved || row.doublons || row.doublons_removed || 0) || 0), 0);
            const duplicateRemovalRate = totalCollected > 0 ? (totalDuplicatesRemoved / totalCollected) * 100 : 0;

            chartManager.createGaugeChart('errorRateGauge', rejectionRate, 'Taux de rejet');
            chartManager.createGaugeChart('validationGauge', validationRate, 'Taux de validation');
            chartManager.createGaugeChart('duplicateGauge', duplicateRemovalRate, 'Taux suppression doublons');
        } catch (error) {
            console.error('Failed to create quality metrics charts:', error);
            showNotification('Failed to create quality metrics charts', 'error');
        }
    }

    /**
     * Creates validation trends chart
     * @param {Array} collectionData - Collection data
     */
    createValidationTrendsChart(collectionData) {
        try {
            const days = Array.from({ length: 30 }, (_, i) => {
                const date = new Date();
                date.setDate(date.getDate() - (29 - i));
                return date.toLocaleDateString('fr-FR', { month: 'short', day: 'numeric' });
            });

            const validatedSeries = (collectionData || []).map(r => Number(r.parcelles_validees_par_lurm || r.parcelles_validees || r.validated || r.urm_validated_parcels || 0));
            const padded = validatedSeries.slice(0, 30).concat(Array(Math.max(0, 30 - validatedSeries.length)).fill(0));

            const renderLine = (canvasId, label, series) => {
                const el = document.getElementById(canvasId);
                if (!el) return;
                const ctx = el.getContext('2d');
                if (!ctx) return;
                if (dashboardState.getChart(canvasId)) dashboardState.getChart(canvasId).destroy();
                try { dashboardState.ensureCanvasFree(canvas, canvas && canvas.id ? canvas.id : null); } catch (e) { /* ignore */ }
                const chart = new Chart(ctx, {
                    type: 'line',
                    data: { labels: days, datasets: [{ label: label, data: series, borderColor: CONFIG.COLORS.success, backgroundColor: CONFIG.COLORS.success + '20', fill: true, tension: 0.3 }] },
                    options: { responsive: true, maintainAspectRatio: false, scales: { y: { beginAtZero: true, title: { display: true, text: 'Parcelles validées' } }, x: { title: { display: true, text: 'Date' } } }, plugins: { legend: { display: true } } }
                });
                dashboardState.setChart(canvasId, chart);
            };

            // Render both panels so the right-side 'Statistiques Validation URM' is not empty
            renderLine('errorAnalysis', 'Parcelles validées/jour', padded);
            renderLine('urmValidationStats', 'Parcelles validées/jour (URM)', padded);
        } catch (error) {
            console.error('Failed to create validation trends chart:', error);
            showNotification('Failed to create validation trends chart', 'error');
        }
    }

    /**
     * Updates quality KPIs
     * @param {Array} communeData - Commune data
     */
    updateQualityKPIs(communeData) {
        try {
            if (!communeData.length) {
                showNotification('No commune data for quality KPIs', 'warning');
                return;
            }

            const totalValidated = communeData.reduce((sum, row) => {
                return sum + (parseFloat(row.parcelles_validees_par_lurm || row.parcelles_validees || row.validated || row.urm_validated_parcels || 0));
            }, 0);

            const totalRejected = communeData.reduce((sum, row) => {
                return sum + (parseFloat(row.urm_rejected_parcels || row.rejected || 0));
            }, 0);

            const totalCollected = communeData.reduce((sum, row) => {
                return sum + (parseFloat(row.collected_parcels_no_duplicates || row.collected || 0));
            }, 0);

            const validationRate = totalCollected > 0 ? (totalValidated / totalCollected) * 100 : 0;
            const rejectionRate = (totalValidated + totalRejected) > 0 ? (totalRejected / (totalValidated + totalRejected)) * 100 : 0;

            // Update visible KPI <p> elements in the quality top cards
            const validationRateTextEl = document.getElementById('validationRateValue');
            const errorRateTextEl = document.getElementById('errorRateValue');
            const duplicateRateTextEl = document.getElementById('duplicateRateValue');

            // Compute duplicate removal rate if a column exists
            const totalDuplicatesRemoved = communeData.reduce((s, r) => s + (Number(r.duplicatesRemoved || r.doublons || r.doublons_removed || 0)), 0);
            const duplicateRemovalRate = totalCollected > 0 ? (totalDuplicatesRemoved / totalCollected) * 100 : 0;

            if (validationRateTextEl) {
                this.animatePercentage(validationRateTextEl, validationRate);
            }

            if (errorRateTextEl) {
                this.animatePercentage(errorRateTextEl, rejectionRate);
            }

            if (duplicateRateTextEl) {
                this.animatePercentage(duplicateRateTextEl, duplicateRemovalRate);
            }

            // Also preserve previous KPI elements if present (backwards compatibility)
            const validationRateElement = document.getElementById('validationRateKPI');
            const rejectionRateElement = document.getElementById('rejectionRateKPI');
            const totalValidatedElement = document.getElementById('totalValidatedKPI');
            if (validationRateElement) this.animatePercentage(validationRateElement, validationRate);
            if (rejectionRateElement) this.animatePercentage(rejectionRateElement, rejectionRate);
            if (totalValidatedElement) this.animateNumber(totalValidatedElement, totalValidated);
        } catch (error) {
            console.error('Failed to update quality KPIs:', error);
            showNotification('Failed to update quality KPIs', 'error');
        }
    }

    /**
     * Generates quality alerts
     * @param {Array} communeData - Commune data
     */
    generateQualityAlerts(communeData) {
        try {
            const alertsContainer = document.getElementById('qualityAlerts');
            if (!alertsContainer) {
                console.warn('Quality alerts container not found');
                return;
            }

            const alerts = [];

            communeData.forEach(row => {
                if (!row.commune || row.commune.toLowerCase() === 'total') return;

                const rejected = parseFloat(row.urm_rejected_parcels || row.rejected || 0);
                const validated = parseFloat(row.parcelles_validees_par_lurm || row.parcelles_validees || row.validated || row.urm_validated_parcels || 0);
                const total = rejected + validated;

                if (total > 0 && (rejected / total) > 0.3) {
                    alerts.push({
                        type: 'warning',
                        message: `${sanitizeHTML(row.commune)}: High rejection rate (${((rejected / total) * 100).toFixed(1)}%)`
                    });
                }

                const collected = parseFloat(row.collected_parcels_no_duplicates || row.collected || 0);
                if (collected > 0 && (validated / collected) < 0.5) {
                    alerts.push({
                        type: 'danger',
                        message: `${sanitizeHTML(row.commune)}: Low validation rate (${((validated / collected) * 100).toFixed(1)}%)`
                    });
                }
            });

            alertsContainer.innerHTML = '';
            alerts.slice(0, 10).forEach(alert => {
                const alertElement = document.createElement('div');
                alertElement.className = `alert-item ${alert.type}`;
                alertElement.innerHTML = `
                    <i class="fas fa-${alert.type === 'danger' ? 'exclamation-triangle' : 'exclamation-circle'}"></i>
                    <span>${sanitizeHTML(alert.message)}</span>
                `;
                alertsContainer.appendChild(alertElement);
            });

            if (alerts.length === 0) {
                alertsContainer.innerHTML = '<div class="alert-item info"><i class="fas fa-check"></i><span>No quality alerts</span></div>';
            }
        } catch (error) {
            console.error('Failed to generate quality alerts:', error);
            showNotification('Failed to generate quality alerts', 'error');
        }
    }

    /**
     * Initializes workflow tab
     */
    async initializeWorkflowTab() {
        try {
            const collectionData = dashboardState.getData('data_collection') || [];
            const deploymentData = dashboardState.getData('team_deployment') || [];
            const methodologyData = dashboardState.getData('methodology_notes') || [];

            // Create process flow visualization
            this.createProcessFlowVisualization(collectionData);

            // Update workflow metrics
            this.updateWorkflowMetrics(collectionData);

            // Update team assignments
            this.updateTeamAssignments(deploymentData);

            // Populate bottleneck details and methodology panels
            this.updateBottleneckDetails(collectionData);
            this.updateMethodologyChecklist(collectionData, methodologyData);
            this.updateMethodologyNotes(methodologyData);
        } catch (error) {
            console.error('Failed to initialize workflow tab:', error);
            showNotification('Failed to load workflow', 'error');
        }
    }

    /**
     * Creates process flow visualization
     * @param {Array} collectionData - Collection data
     */
    createProcessFlowVisualization(collectionData) {
        try {
            // prefer visible container 'processFlow' and a visible canvas inside it; fallback to 'processFlowChart'
            var container = document.getElementById('processFlow') || document.getElementById('processFlowChart');
            var canvas = container && container.querySelector && container.querySelector('canvas') ? container.querySelector('canvas') : (document.getElementById('processFlowChart') || null);
            if (!canvas) {
                console.warn('Canvas not found in processFlow or processFlowChart');
                return;
            }

            var ctx = canvas.getContext('2d');

            // TODO: Replace mock data with real data processing
            // Map dashboard process stages to canonical field keys that may exist in collectionData
            const stageMap = [
                { label: 'Collection', keys: ['collected', 'parcelles_collectees', 'parcelles_collectees_sans_doublon_geometrique', 'collected_parcels_no_duplicates'] },
                { label: 'Validation', keys: ['validated', 'parcelles_validees_par_lurm', 'urm_validated_parcels'] },
                { label: 'Deliberation', keys: ['deliberated', 'deliberees'] },
                { label: 'Registration', keys: ['registered', 'registration', 'enregistrement'] }
            ];

            const stages = stageMap.map(s => s.label);
            const stageData = stageMap.map(s => {
                return collectionData.reduce((sum, row) => {
                    const v = this.getNumericValue(row, s.keys);
                    return sum + (isNaN(v) ? 0 : v);
                }, 0);
            });

            try { dashboardState.ensureCanvasFree(canvas, canvas && canvas.id ? canvas.id : null); } catch (e) { /* ignore */ }
            const chart = new Chart(ctx, {
                type: 'bar',
                data: {
                    labels: stages,
                    datasets: [{
                        label: 'Parcelles',
                        data: stageData,
                        backgroundColor: CONFIG.COLORS.gradients.primary,
                        borderColor: CONFIG.COLORS.primary,
                        borderWidth: 1
                    }]
                },
                options: {
                    ...chartManager.defaultOptions,
                    scales: {
                        y: {
                            beginAtZero: true,
                            title: {
                                display: true,
                                text: 'Nombre de parcelles'
                            }
                        },
                        x: {
                            title: {
                                display: true,
                                text: 'Étapes du processus'
                            }
                        }
                    }
                }
            });

            dashboardState.setChart('processFlow', chart);
        } catch (error) {
            console.error('Failed to create process flow visualization:', error);
            showNotification('Failed to create process flow visualization', 'error');
        }
    }

    /**
     * Updates workflow metrics
     * @param {Array} collectionData - Collection data
     */
    updateWorkflowMetrics(collectionData) {
        try {
            if (!collectionData.length) {
                showNotification('No collection data for workflow metrics', 'warning');
                return;
            }

            const totalProcessed = collectionData.reduce((sum, row) => {
                return sum + this.getNumericValue(row, ['total', 'total_parcels', 'total_parcelles', 'parcelles_post_traitees_lot_1_46']);
            }, 0);

            const avgProcessingTime = collectionData.length ? (collectionData.reduce((sum, row) => {
                return sum + this.getNumericValue(row, ['processing_time', 'avg_processing_time', 'temps_traitement']);
            }, 0) / collectionData.length) : 0;

            const totalProcessedElement = document.getElementById('totalProcessedParcels');
            const avgProcessingTimeElement = document.getElementById('avgProcessingTime');

            if (totalProcessedElement) {
                this.animateNumber(totalProcessedElement, totalProcessed || 0);
            }

            if (avgProcessingTimeElement) {
                this.animateNumber(avgProcessingTimeElement, avgProcessingTime || 0);
            }
        } catch (error) {
            console.error('Failed to update workflow metrics:', error);
            showNotification('Failed to update workflow metrics', 'error');
        }
    }

    /**
     * Updates team assignments
     * @param {Array} deploymentData - Deployment data
     */
    updateTeamAssignments(deploymentData) {
        try {
            // prefer visible team leaderboard container
            const assignmentsContainer = document.getElementById('teamLeaderboard') || document.getElementById('teamAssignments');
            if (!assignmentsContainer) {
                console.warn('Team assignments container not found (teamLeaderboard/teamAssignments)');
                return;
            }

            assignmentsContainer.innerHTML = '';

            deploymentData.forEach(row => {
                const assignment = document.createElement('div');
                assignment.className = 'assignment-item';
                assignment.innerHTML = `
                    <div class="assignment-header">
                        <span class="team-name">${sanitizeHTML(row.team || 'Unknown')}</span>
                        <span class="status ${this.getDeploymentStatus(row).toLowerCase()}">${sanitizeHTML(this.getDeploymentStatus(row))}</span>
                    </div>
                    <div class="assignment-details">
                        <span>Région: ${sanitizeHTML(row.region || 'N/A')}</span>
                        <span>Commune: ${sanitizeHTML(row.commune || 'N/A')}</span>
                        <span>Début: ${sanitizeHTML(row.tambacounda_start_date || row.kedougou_start_date || 'TBD')}</span>
                    </div>
                `;
                assignmentsContainer.appendChild(assignment);
            });
        } catch (error) {
            console.error('Failed to update team assignments:', error);
            showNotification('Failed to update team assignments', 'error');
        }
    }

    /**
     * Populates bottleneck details panel based on collection data
     * @param {Array} collectionData
     */
    updateBottleneckDetails(collectionData) {
        try {
            const container = document.getElementById('bottleneckDetails');
            if (!container) return;

            container.innerHTML = '';

            // Heuristic: find communes with high rejection rate, long processing time or many unjoined parcels
            const candidates = [];
            collectionData.forEach(row => {
                if (!row.commune || row.commune.toLowerCase() === 'total') return;

                const rejected = this.getNumericValue(row, ['urm_rejected_parcels', 'urm_rejected_parcels', 'rejected', 'parcelles_rejetees_par_lurm']);
                const validated = this.getNumericValue(row, ['urm_validated_parcels', 'validated', 'parcelles_validees_par_lurm']);
                const collected = this.getNumericValue(row, ['collected_parcels_no_duplicates', 'collected', 'parcelles_collectees']);
                const unjoined = this.getNumericValue(row, ['unjoined', 'parcelles_non_jointes']);
                const processingTime = this.getNumericValue(row, ['processing_time', 'temps_traitement']);

                const total = (isNaN(rejected) ? 0 : rejected) + (isNaN(validated) ? 0 : validated);
                const rejectionRate = total > 0 ? ((isNaN(rejected) ? 0 : rejected) / total) : 0;

                // score for rank
                const score = (rejectionRate * 2) + (unjoined / (collected || 1)) + (processingTime / 100 || 0);

                candidates.push({ row, score, rejectionRate, unjoined, processingTime });
            });

            candidates.sort((a, b) => b.score - a.score);

            // Show top 5 bottlenecks
            candidates.slice(0, 5).forEach(item => {
                const r = item.row;
                const el = document.createElement('div');
                el.className = 'bottleneck-item';
                el.innerHTML = `
                    <div class="bottleneck-header">
                        <strong>${sanitizeHTML(r.commune || 'Unknown')}</strong>
                        <span class="bottleneck-severity ${item.score > 1 ? 'high' : (item.score > 0.5 ? 'medium' : 'low')}">${(item.rejectionRate * 100).toFixed(1)}%</span>
                    </div>
                    <div class="bottleneck-body">
                        <div>Parcelles collectées: ${((this.getNumericValue(r, ['collected', 'collected_parcels_no_duplicates', 'parcelles_collectees']) || 0)).toLocaleString()}</div>
                        <div>Parcelles validées: ${((this.getNumericValue(r, ['validated', 'urm_validated_parcels', 'parcelles_validees_par_lurm']) || 0)).toLocaleString()}</div>
                        <div>Parcelles rejetées: ${((this.getNumericValue(r, ['rejected', 'urm_rejected_parcels', 'parcelles_rejetees_par_lurm']) || 0)).toLocaleString()}</div>
                        <div>Parcelles non jointes: ${((this.getNumericValue(r, ['unjoined', 'parcelles_non_jointes']) || 0)).toLocaleString()}</div>
                    </div>
                `;
                container.appendChild(el);
            });

            if (!candidates.length) {
                container.innerHTML = '<div class="bottleneck-empty">Aucun goulot identifié</div>';
            }
        } catch (error) {
            console.error('Failed to update bottleneck details:', error);
        }
    }

    /**
     * Populates methodology checklist (Adhérence Méthodologique)
     * @param {Array} collectionData
     * @param {Array} methodologyData
     */
    updateMethodologyChecklist(collectionData, methodologyData) {
        try {
            const container = document.getElementById('methodologyChecklist');
            if (!container) return;

            container.innerHTML = '';

            // Basic checks: presence of geomaticien, processing_time, join status
            const checks = [
                { key: 'geomaticien', label: 'Géomaticien renseigné', pass: collectionData.some(r => !!this.getGeomaticienName(r)) },
                { key: 'processing_time', label: 'Temps de traitement renseigné', pass: collectionData.some(r => this.getNumericValue(r, ['processing_time', 'temps_traitement']) > 0) },
                { key: 'joinStatus', label: 'Statut de jointure présent', pass: collectionData.some(r => !!(r.joinStatus || r.statut_jointure || r.join_status)) },
                { key: 'urmRejectionReasons', label: 'Motifs de rejet URM documentés', pass: collectionData.some(r => !!(r.urmRejectionReasons || r.motifs_de_rejet_urm)) }
            ];

            checks.forEach(c => {
                const item = document.createElement('div');
                item.className = 'check-item';
                item.innerHTML = `<span class="check ${c.pass ? 'pass' : 'fail'}">${c.pass ? '✓' : '✗'}</span> <span class="check-label">${sanitizeHTML(c.label)}</span>`;
                container.appendChild(item);
            });

            // If methodology notes exist, show how many
            if (methodologyData && methodologyData.length) {
                const noteSummary = document.createElement('div');
                noteSummary.className = 'methodology-summary';
                noteSummary.textContent = `${methodologyData.length} note(s) méthodologiques disponibles`;
                container.appendChild(noteSummary);
            }
        } catch (error) {
            console.error('Failed to update methodology checklist:', error);
        }
    }

    /**
     * Renders methodology notes
     * @param {Array} methodologyData
     */
    updateMethodologyNotes(methodologyData) {
        try {
            const container = document.getElementById('methodologyNotes');
            if (!container) return;

            container.innerHTML = '';

            if (!methodologyData || !methodologyData.length) {
                container.innerHTML = '<div class="note-empty">Aucune note méthodologique</div>';
                return;
            }

            methodologyData.forEach((row, idx) => {
                const note = document.createElement('div');
                note.className = 'note-item';
                const title = row.title || row.titre || `Note ${idx + 1}`;
                const desc = row.note || row.description || row.note_methodologique || row.comment || '';
                note.innerHTML = `
                    <div class="note-title">${sanitizeHTML(title)}</div>
                    <div class="note-description">${sanitizeHTML(desc)}</div>
                `;
                container.appendChild(note);
            });
        } catch (error) {
            console.error('Failed to update methodology notes:', error);
        }
    }

    /**
     * Exports dashboard data to Excel
     * @requires SheetJS (XLSX) library
     */
    exportData() {
        try {
            if (typeof XLSX === 'undefined') {
                console.warn('SheetJS library not loaded');
                showNotification('Export functionality requires SheetJS library', 'error');
                return;
            }

            showLoadingIndicator('Exporting data...');

            const workbook = XLSX.utils.book_new();
            const allData = dashboardState.data;

            allData.forEach((data, sheetName) => {
                if (!data || !data.length) return;

                // Convert data to worksheet
                const worksheetData = data.map(row => {
                    const cleanedRow = {};
                    Object.keys(row).forEach(key => {
                        cleanedRow[sanitizeHTML(key)] = row[key];
                    });
                    return cleanedRow;
                });

                const worksheet = XLSX.utils.json_to_sheet(worksheetData);
                XLSX.utils.book_append_sheet(workbook, worksheet, sheetName.slice(0, 31)); // Excel sheet name limit
            });

            // Generate and download file
                // Use SheetJS correct browser API
                if (typeof XLSX.writeFile === 'function') {
                    XLSX.writeFile(workbook, 'procasef_dashboard_export.xlsx');
                } else if (typeof XLSX.write === 'function') {
                    const wbout = XLSX.write(workbook, {bookType:'xlsx', type:'array'});
                    const blob = new Blob([wbout], {type:'application/octet-stream'});
                    const url = URL.createObjectURL(blob);
                    const a = document.createElement('a'); a.href = url; a.download = 'procasef_dashboard_export.xlsx'; document.body.appendChild(a); a.click(); a.remove(); URL.revokeObjectURL(url);
                } else {
                    console.error('SheetJS API not available (writeFile/write)');
                }

            hideLoadingIndicator();
            showNotification('Data exported successfully', 'success');
        } catch (error) {
            console.error('Failed to export data:', error);
            hideLoadingIndicator();
            showNotification('Failed to export data', 'error');
        }
    }

    /**
     * Downloads a chart as an image
     * @param {string} chartId - Chart canvas ID
     */
    downloadChart(chartId) {
        try {
            const chart = dashboardState.getChart(chartId);
            if (!chart) {
                showNotification('Chart not found', 'error');
                return;
            }

            const canvas = chart.canvas;
            const link = document.createElement('a');
            link.download = `${chartId}_${new Date().toISOString().slice(0, 10)}.png`;
            link.href = canvas.toDataURL('image/png');
            link.click();
            showNotification('Chart downloaded successfully', 'success');
        } catch (error) {
            console.error(`Failed to download chart ${chartId}:`, error);
            showNotification('Failed to download chart', 'error');
        }
    }

    /**
     * Generates a PDF report
     * @requires jsPDF library (not implemented)
     */
    generateReport() {
        try {
            // Note: Requires jsPDF library to be included
            showNotification('PDF report generation not implemented (requires jsPDF)', 'warning');
            // TODO: Implement PDF report generation with actual data
        } catch (error) {
            console.error('Failed to generate report:', error);
            showNotification('Failed to generate report', 'error');
        }
    }
}

// ============================
// INITIALIZATION
// ============================

/**
 * Initializes the dashboard on page load
 */
async function initDashboard() {
    try {
        const uiManager = new UIManager();
        await uiManager.initialize();

        // Setup global event listeners
        dashboardState.on('settingsChanged', ({ key, value }) => {
            console.log(`Setting ${key} changed to ${value}`);
            if (key === 'darkMode') {
                uiManager.toggleDarkMode(value);
            }
        });

        // Apply initial dark mode setting
        uiManager.toggleDarkMode(dashboardState.settings.darkMode);
    } catch (error) {
        console.error('Dashboard initialization failed:', error);
        showNotification('Failed to initialize dashboard', 'error');
    }
}

// Run initialization when DOM is loaded
document.addEventListener('DOMContentLoaded', initDashboard);