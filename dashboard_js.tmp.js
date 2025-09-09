/* temp dashboard_js - final clean implementation */
(function(){
  if (window.__PROCASEF_DASH_LOADED) return; window.__PROCASEF_DASH_LOADED = true;
  var SHEET_CSV_URL = 'https://docs.google.com/spreadsheets/d/e/2PACX-1vSu_1nF47cOxavQnnbiyo2XbTnV-6XLypzrsHnHmIjHVhhtYMKYVQHBgurb7Mh8fg/pub?output=csv';
  var CACHE_KEY = 'procasef_sheets_cache_v1'; var CACHE_TTL = 5*60*1000;
  function splitCSVLine(line){ var res=[], cur='', inQ=false; for(var i=0;i<line.length;i++){ var c=line[i]; if(c==='"'){ if(inQ && line[i+1]==='"'){ cur+='"'; i++; } else inQ=!inQ; } else if(c===',' && !inQ){ res.push(cur); cur=''; } else cur+=c; } res.push(cur); return res; }
  function parseCSV(text){ var lines=String(text||'').replace(/\r\n/g,'\n').split('\n').filter(function(l){return l.trim()!==''}); if(!lines.length) return []; var hdr=splitCSVLine(lines[0]).map(function(h){return h.trim()}); var out=[]; for(var i=1;i<lines.length;i++){ var cols=splitCSVLine(lines[i]); var obj={}; for(var j=0;j<hdr.length;j++) obj[hdr[j]||('col_'+j)]=(cols[j]||'').trim(); out.push(obj);} return out; }
  function saveCache(v){ try{ localStorage.setItem(CACHE_KEY, JSON.stringify({ts:Date.now(),payload:v})); }catch(e){} }
  function loadCache(){ try{ var r=localStorage.getItem(CACHE_KEY); if(!r) return null; var p=JSON.parse(r); if(!p.ts||Date.now()-p.ts>CACHE_TTL) return null; return p.payload;}catch(e){return null;} }
  async function fetchSheetCSV(url){ var r=await fetch(url,{cache:'no-store'}); if(!r.ok) throw new Error('fetch failed'); var t=await r.text(); return parseCSV(t); }
  function norm(k){ return String(k||'').trim().toLowerCase().replace(/\s+/g,'_').replace(/[^a-z0-9_]/g,''); }
  function handleParsedRows(rows){ var normalized=rows.map(function(r){ var o={}; for(var k in r) o[norm(k)]=r[k]; o.commune=o.commune||''; o.region=o.region||''; o.collected=Number(o.collected)||0; o.surveyed=Number(o.surveyed)||0; o.validated=Number(o.validated)||0; o.rejected=Number(o.rejected)||0; return o; }); window.dispatchEvent(new CustomEvent('procasef:communes:updated',{detail:{data:normalized}})); var sum=function(k){ return normalized.reduce(function(s,r){ return s + (Number(r[k])||0); },0); }; var kpis={ total_parcels:sum('collected'), surveyed:sum('surveyed'), validated:sum('validated'), rejected:sum('rejected') }; window.dispatchEvent(new CustomEvent('procasef:kpis:loaded',{detail:{kpis:kpis}})); }
  async function initialLoad(){ var cached=loadCache(); if(cached&&Array.isArray(cached)&&cached.length) handleParsedRows(cached); try{ var rows=await fetchSheetCSV(SHEET_CSV_URL); saveCache(rows); handleParsedRows(rows); }catch(err){ console.error('sheet load error',err); if(!cached) window.dispatchEvent(new CustomEvent('procasef:communes:updated',{detail:{data:[]}})); } }
  document.addEventListener('DOMContentLoaded', function(){ initialLoad(); });
})();
