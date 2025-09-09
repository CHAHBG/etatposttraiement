/* PROCASEF Dashboard Loader - script.js */
if (!window.__PROCASEF_LOADER_LOADED) {
  class PROCASEFDashboard {
    constructor(opts = {}) {
        this.baseUrl = opts.baseUrl || 'https://docs.google.com/spreadsheets/d/e/2PACX-1vR3G2rPcnFBbmpVVeE5k227lcMA9b-mrkTPiUxeNoIBdM8tPry6uEcxyaCUCns6gw/pub';
        this.sheets = opts.sheets || [
        'KPI Dashboard', 'Data Collection', 'Projection Collection', 'Yields Projection', 'Team Deployment',
        'NICAD Projection', 'Public Display', 'CTASF Projection', 'Commune Status', 'Methodology Notes'
      ];
      this.cacheKey = opts.cacheKey || 'procasef_sheets_cache_v1';
      this.cacheTTL = opts.cacheTTL || 1000 * 60 * 5; // 5 minutes
  this.timeout = opts.timeout || 15000; // 15s per request
    }

  // ...existing code...

    async fetchAll() {
      const cached = this._readCache();
      if (cached && (Date.now() - cached.ts) < this.cacheTTL) {
        this._emit('cache', cached.data);
        return cached.data;
      }

      const promises = this.sheets.map(sheet => this._fetchSheetCSV(sheet).catch(err => ({ error: String(err), sheet })));
      const results = await Promise.all(promises);
      const data = {};
      results.forEach((result, i) => {
        const sheet = this.sheets[i];
        data[sheet] = result.error ? { error: result.error } : this._parseCSV(result);
      });

      this._writeCache(data);
      this._emit('network', data);
      return data;
    }

    // Fetch multiple sheets directly from Google Sheets API using a client API key
    async fetchAllWithKey(spreadsheetId, apiKey) {
  const ranges = this.sheets.map(s => `ranges=${encodeURIComponent("'" + s + "'")}`).join('&');
      const url = `https://sheets.googleapis.com/v4/spreadsheets/${encodeURIComponent(spreadsheetId)}/values:batchGet?${ranges}&majorDimension=ROWS&key=${encodeURIComponent(apiKey)}`;
      const resp = await fetch(url, { cache: 'no-store' });
      if (!resp.ok) {
        const txt = await resp.text();
        throw new Error('Google Sheets API error: ' + resp.status + ' ' + txt);
      }
      const json = await resp.json();
      const out = {};
      (json.valueRanges || []).forEach(vr => {
        const rangeName = vr.range ? vr.range.split('!')[0].replace(/^'/, '').replace(/'$/, '') : '';
        const values = vr.values || [];
        if (!values.length) { out[rangeName] = []; return; }
        const headers = values[0].map(h => String(h || '').trim());
        const rows = values.slice(1).map(r => {
          const obj = {};
          headers.forEach((h, i) => { obj[h || `col_${i}`] = r[i] !== undefined ? String(r[i]) : ''; });
          return obj;
        });
        out[rangeName] = rows;
      });
      return out;
    }

    // Fetch all published CSV sheets by gid using the publish URL base
    async fetchAllPublished(publishBase, gidMap) {
      const data = {};
      for (const sheetName of this.sheets) {
        try {
          const gid = gidMap[sheetName];
          if (!gid) { data[sheetName] = []; continue; }
          const url = `${publishBase}?gid=${encodeURIComponent(gid)}&single=true&output=csv`;
          const resp = await fetch(url, { cache: 'no-store' });
          if (!resp.ok) {
            data[sheetName] = { error: `HTTP ${resp.status}` };
            continue;
          }
          const txt = await resp.text();
          data[sheetName] = this._parseCSV(txt);
        } catch (err) {
          data[sheetName] = { error: String(err) };
        }
      }
      this._writeCache(data);
      this._emit('published', data);
      return data;
    }

    async _fetchSheetCSV(sheetName) {
      const url = `${this.baseUrl}?output=csv&sheet=${encodeURIComponent(sheetName)}`;
      const controller = new AbortController();
      const timer = setTimeout(() => controller.abort(), this.timeout);
      try {
        const resp = await fetch(url, { signal: controller.signal });
        clearTimeout(timer);
        if (!resp.ok) throw new Error(`HTTP ${resp.status} fetching ${sheetName}`);
        return await resp.text();
      } catch (err) {
        clearTimeout(timer);
        throw err;
      }
    }

    _parseCSV(text) {
      const rows = [];
      let i = 0, row = [], field = '', inQuotes = false;
      while (i < text.length) {
        const ch = text[i];
        if (inQuotes) {
          if (ch === '"' && text[i + 1] !== '"') { inQuotes = false; i++; continue; }
          if (ch === '"' && text[i + 1] === '"') { field += '"'; i += 2; continue; }
          field += ch; i++; continue;
        }
        if (ch === '"') { inQuotes = true; i++; continue; }
        if (ch === ',') { row.push(field); field = ''; i++; continue; }
        if (ch === '\n') { row.push(field); rows.push(row); row = []; field = ''; i++; continue; }
        field += ch; i++;
      }
      if (field || row.length) row.push(field);
      if (row.length) rows.push(row);

      if (!rows.length) return [];
      const headers = rows[0].map(h => (h || '').trim());
      return rows.slice(1).map(r => {
        const obj = {};
        headers.forEach((h, j) => { obj[h || `col_${j}`] = (r[j] || '').trim(); });
        return obj;
      });
    }

    _readCache() {
      try {
        const raw = localStorage.getItem(this.cacheKey);
        return raw ? JSON.parse(raw) : null;
      } catch (err) {
        console.warn('Failed reading cache', err);
        return null;
      }
    }

    _writeCache(data) {
      try {
        localStorage.setItem(this.cacheKey, JSON.stringify({ ts: Date.now(), data }));
      } catch (err) {
        console.warn('Failed writing cache', err);
      }
    }

    _emit(source, data) {
      window.dispatchEvent(new CustomEvent('procasef:data:loaded', { detail: { source, data, ts: Date.now() } }));
    }
  }

  const PROCASEF = new PROCASEFDashboard();
  // If you want to use the server proxy, set this to your spreadsheet ID and
  // enable proxy by setting PROCASEF.useProxy = true and PROCASEF.proxyBase = 'http://localhost:3000'
  const SPREADSHEET_ID = '1CbDBJtoWWPRjEH4DOSlv4jjnJ2G-1MvW';
  // Client-side API key mode (not recommended for production because the key is exposed)
  // You said it's OK to expose the key — set it here to enable direct client requests.
  const CLIENT_API_KEY = 'AIzaSyAvky6j6mvCDPNg0dBYfBepOvNL3JG_rtc';
  // Enable direct client-side calls to Google Sheets API (set to true to use the key above)
  // Default to false to avoid accidental batchGet calls; published CSV mode is preferred.
  PROCASEF.useClientKey = false;
  // Published CSV mode (no API key required). Provide the publish base URL and mapping of sheet names -> gid.
  // Example publish base (the one you gave):
  const PUBLISHED_BASE = 'https://docs.google.com/spreadsheets/d/e/2PACX-1vSu_1nF47cOxavQnnbiyo2XbTnV-6XLypzrsHnHmIjHVhhtYMKYVQHBgurb7Mh8fg/pub';
  // Map sheet titles used in this loader to their publish gids (from google_sheets_index.md)
  const PUBLISHED_GIDS = {
    'KPI Dashboard': '589154853',
    'Data Collection': '778550489',
    'Projection Collection': '1687610293',
    'Yields Projection': '1397416280',
    'Team Deployment': '660642704',
    'NICAD Projection': '1151168155',
    'Public Display': '2035205606',
    'CTASF Projection': '1574818070',
    'Commune Status': '1421590976',
    'Methodology Notes': '1203072335'
  };
  // Enable published CSV mode (set to true to fetch per-gid published CSVs)
  PROCASEF.usePublished = true;

  function toNum(v) {
    if (v === null || v === undefined) return 0;
    if (typeof v === 'number') return v;
    const n = Number(String(v).replace(/[^0-9.-]+/g, ''));
    return Number.isFinite(n) ? n : 0;
  }

  async function initialLoad() {
    const overlay = document.getElementById('loadingOverlay');
    if (overlay) overlay.style.display = 'flex';
    try {
      let sheets;
      if (PROCASEF.useProxy) {
        sheets = await PROCASEF.fetchAllProxy(SPREADSHEET_ID);
      } else if (PROCASEF.usePublished) {
        // Published CSV per-gid is the preferred no-key option when enabled.
        sheets = await PROCASEF.fetchAllPublished(PUBLISHED_BASE, PUBLISHED_GIDS);
      } else if (PROCASEF.useClientKey) {
        // Client API key mode only runs when explicitly enabled (useClientKey = true).
        sheets = await PROCASEF.fetchAllWithKey ? await PROCASEF.fetchAllWithKey(SPREADSHEET_ID, CLIENT_API_KEY) : await PROCASEF.fetchAll();
      } else {
        sheets = await PROCASEF.fetchAll();
      }
      
      // Process Commune Status
      if (Array.isArray(sheets['Commune Status'])) {
        window.communesData = sheets['Commune Status'].map(row => ({
          commune: row['Commune'] || '',
          region: row['Region'] || '',
          totalParcels: toNum(row['Total Parcels'] || '0'),
          totalPct: toNum(row['Total %'] || '0'),
          nicad: toNum(row['NICAD'] || '0'),
          nicadPct: toNum(row['NICAD %'] || '0'),
          ctasf: toNum(row['CTASF'] || '0'),
          ctasfPct: toNum(row['CTASF %'] || '0'),
          deliberated: toNum(row['Deliberated'] || '0'),
          deliberatedPct: toNum(row['Deliberated %'] || '0'),
          collected: toNum(row['Collected Parcels (No Duplicates)'] || '0'),
          surveyed: toNum(row['Surveyed Parcels'] || '0'),
          validated: toNum(row['URM Validated Parcels'] || '0'),
          rejected: toNum(row['URM Rejected Parcels'] || '0'),
          geomatician: row['Geomatician'] || '',
          duplicateRemovalRatePct: toNum(row['Duplicate Removal Rate (%)'] || '0'),
          retained: toNum(row['Retained Parcels'] || '0')
        }));
        window.dispatchEvent(new CustomEvent('procasef:communes:updated', { detail: { data: window.communesData } }));
      }

      // Process KPI Dashboard
      if (Array.isArray(sheets['KPI Dashboard'])) {
        window.kpiData = sheets['KPI Dashboard'].map(row => ({
          totalCollected: toNum(row['Total Parcels Collected'] || '0'),
          totalSurveyed: toNum(row['Surveyed Parcels'] || '0'),
          totalValidated: toNum(row['Parcels Validated'] || '0'),
          totalRejected: toNum(row['Parcels Rejected'] || '0'),
          nicadPct: toNum(row['NICAD %'] || '0'),
          ctasfPct: toNum(row['CTASF %'] || '0'),
          deliberatedPct: toNum(row['Deliberated %'] || '0')
        }));
        window.dispatchEvent(new CustomEvent('procasef:kpis:loaded', { detail: { kpis: window.kpiData } }));
      }

      // Process Data Collection
      if (Array.isArray(sheets['Data Collection'])) {
        window.collectionData = sheets['Data Collection'].map(row => ({
          phase: row['Phase'] || '',
          collected: toNum(row['Collected'] || '0'),
          date: row['Date'] || ''
        }));
      }

      // Process Projection Collection
      if (Array.isArray(sheets['Projection Collection'])) {
        window.projectionData = sheets['Projection Collection'].map(row => ({
          month: row['Month'] || '',
          projectedParcels: toNum(row['Projected Parcels'] || '0')
        }));
      }

      // Process Yields Projection
      if (Array.isArray(sheets['Yields Projection'])) {
        window.yieldsData = sheets['Yields Projection'].map(row => ({
          commune: row['Commune'] || '',
          yieldEstimate: toNum(row['Yield Estimate'] || '0')
        }));
      }

      // Process Team Deployment
      if (Array.isArray(sheets['Team Deployment'])) {
        window.teamData = sheets['Team Deployment'].map(row => ({
          geomatician: row['Geomatician'] || '',
          assignedParcels: toNum(row['Assigned Parcels'] || '0'),
          startDate: row['Start Date'] || '',
          endDate: row['End Date'] || ''
        }));
      }

      // Process NICAD Projection
      if (Array.isArray(sheets['NICAD Projection'])) {
        window.nicadProjectionData = sheets['NICAD Projection'].map(row => ({
          week: row['Week'] || '',
          assignments: toNum(row['Assignments'] || '0')
        }));
      }

      // Process Public Display
      if (Array.isArray(sheets['Public Display'])) {
        window.publicDisplayData = sheets['Public Display'].map(row => ({
          week: row['Week'] || '',
          progress: toNum(row['Progress %'] || '0')
        }));
      }

      // Process CTASF Projection
      if (Array.isArray(sheets['CTASF Projection'])) {
        window.ctasfProjectionData = sheets['CTASF Projection'].map(row => ({
          week: row['Week'] || '',
          completion: toNum(row['Completion %'] || '0')
        }));
      }

      // Process Methodology Notes
      if (Array.isArray(sheets['Methodology Notes'])) {
        window._procasef_methodology_notes = sheets['Methodology Notes'].map(row => ({
          section: row['Section'] || '',
          completed: row['Description']?.includes('Complete') || false
        }));
      }

      if (overlay) overlay.style.display = 'none';
    } catch (err) {
      if (overlay) overlay.style.display = 'none';
      const toast = document.getElementById('notificationToast');
      if (toast) {
        document.getElementById('notificationMessage').textContent = 'Erreur lors du chargement des données';
        toast.classList.add('show');
        setTimeout(() => toast.classList.remove('show'), 5000);
      }
      console.error('Failed to load PROCASEF sheets', err);
    }
  }

  document.addEventListener('DOMContentLoaded', () => {
    initialLoad();
    const refreshBtn = document.getElementById('refreshDataBtn');
    if (refreshBtn) {
      refreshBtn.addEventListener('click', async () => {
        localStorage.removeItem(PROCASEF.cacheKey);
        await initialLoad();
        document.getElementById('lastRefreshTime').textContent = new Date().toLocaleString();
      });
    }
    const clearCacheBtn = document.getElementById('clearCacheBtn');
    if (clearCacheBtn) {
      clearCacheBtn.addEventListener('click', () => {
        localStorage.removeItem(PROCASEF.cacheKey);
        const toast = document.getElementById('notificationToast');
        if (toast) {
          document.getElementById('notificationMessage').textContent = 'Cache cleared successfully';
          toast.classList.add('show');
          setTimeout(() => toast.classList.remove('show'), 3000);
        }
      });
    }
  });

  window.__PROCASEF_LOADER_LOADED = true;
}
    