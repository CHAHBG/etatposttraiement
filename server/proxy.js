const express = require('express');
const fetch = require('node-fetch');
const app = express();
const PORT = process.env.PORT || 3000;
const API_KEY = process.env.GOOGLE_API_KEY;

if (!API_KEY) {
  console.warn('WARNING: GOOGLE_API_KEY is not set. The proxy will fail until you set process.env.GOOGLE_API_KEY');
}

// Development helper: set permissive CSP/headers on proxy responses so dev pages can load fonts/scripts
app.use((req, res, next) => {
  res.setHeader('Content-Security-Policy', "default-src 'self' data: https:; script-src 'self' https://cdn.jsdelivr.net https://cdnjs.cloudflare.com 'unsafe-inline'; style-src 'self' 'unsafe-inline' https://fonts.googleapis.com https://cdnjs.cloudflare.com; font-src 'self' https://fonts.gstatic.com data:; connect-src 'self' http://localhost:3000 https://sheets.googleapis.com https://docs.google.com; img-src 'self' data: https:");
  res.setHeader('Access-Control-Allow-Origin', '*');
  next();
});

// Simple request logger for debugging
app.use((req, res, next) => {
  try {
    console.log(new Date().toISOString(), req.method, req.url);
  } catch (e) {}
  next();
});

// Simple endpoint: /api/sheets?spreadsheetId=<id>&sheets=Sheet1,Sheet2
app.get('/api/sheets', async (req, res) => {
  const spreadsheetId = req.query.spreadsheetId;
  const sheetsParam = req.query.sheets || '';
  if (!spreadsheetId) return res.status(400).json({ error: 'spreadsheetId query param required' });
  const sheets = sheetsParam.split(',').map(s => s.trim()).filter(Boolean);

  // Build ranges query params
  const ranges = sheets.map(s => `ranges=${encodeURIComponent("'" + s + "'")}`).join('&');
  const url = `https://sheets.googleapis.com/v4/spreadsheets/${encodeURIComponent(spreadsheetId)}/values:batchGet?${ranges}&majorDimension=ROWS&key=${API_KEY}`;

  try {
    const resp = await fetch(url, { method: 'GET' });
    if (!resp.ok) {
      const text = await resp.text();
      return res.status(resp.status).json({ error: text });
    }
    const payload = await resp.json();

    // Convert valueRanges to a friendly object: { sheetName: [ {col:val,...}, ... ] }
    const out = {};
    (payload.valueRanges || []).forEach(vr => {
      const rangeName = vr.range ? vr.range.split('!')[0].replace(/^'/, '').replace(/'$/, '') : '';
      const values = vr.values || [];
      if (!values.length) {
        out[rangeName] = [];
        return;
      }
      const headers = values[0].map(h => String(h || '').trim());
      const rows = values.slice(1).map(r => {
        const obj = {};
        headers.forEach((h, i) => { obj[h || `col_${i}`] = r[i] !== undefined ? String(r[i]) : ''; });
        return obj;
      });
      out[rangeName] = rows;
    });

    res.json({ data: out, ts: Date.now() });
  } catch (err) {
    console.error('Proxy fetch error', err);
    res.status(500).json({ error: String(err) });
  }
});

// Root/help page to avoid 404 on / and provide basic instructions
app.get('/', (req, res) => {
  res.setHeader('Content-Type', 'text/html; charset=utf-8');
  res.send(`
    <html><head><title>Google Sheets Proxy</title></head><body>
      <h2>Google Sheets Proxy</h2>
      <p>Endpoints:</p>
      <ul>
        <li><code>GET /api/sheets?spreadsheetId=&lt;id&gt;&sheets=Sheet1,Sheet2</code> - returns JSON data</li>
      </ul>
      <p>Example (curl):</p>
      <pre>curl "http://localhost:${PORT}/api/sheets?spreadsheetId=1CbDBJtoWWPRjEH4DOSlv4jjnJ2G-1MvW&sheets=Commune Status,Data Collection"</pre>
      <p>Make sure <code>GOOGLE_API_KEY</code> is set in the environment before starting this server.</p>
    </body></html>
  `);
});

app.listen(PORT, () => {
  console.log(`Google Sheets proxy listening on http://localhost:${PORT}`);
});
