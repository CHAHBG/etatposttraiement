# Google Sheets Proxy

This folder contains a minimal Node/Express proxy to fetch Google Sheets data server-side using a Google API key. It prevents exposing the API key to client-side code.

Requirements
- Node >= 14
- Set the environment variable `GOOGLE_API_KEY` with your API key (do NOT commit it).

Quick start

1. Install dependencies:

```powershell
npm init -y; npm install express node-fetch@2
```

2. Run the proxy with your API key:

```powershell
$env:GOOGLE_API_KEY = 'YOUR_KEY_HERE'; node server/proxy.js
```

3. Example request (fetch two sheets):

```powershell
curl "http://localhost:3000/api/sheets?spreadsheetId=1CbDBJtoWWPRjEH4DOSlv4jjnJ2G-1MvW&sheets=\"Commune Status\",\"Data Collection\""
```

Response format: { data: { 'Commune Status': [ ...rows... ], 'Data Collection': [ ... ] }, ts: 123456 }

Notes
- This server is intentionally minimal. For production use, add rate-limiting, authentication, and error handling.
