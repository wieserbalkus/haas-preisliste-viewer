(function(global){
  const HaasApp = global.HaasApp || (global.HaasApp = {});

  if(typeof global.requestAnimationFrame !== 'function'){
    global.requestAnimationFrame = (cb)=>global.setTimeout(()=>{
      try{ cb(Date.now()); }
      catch(err){
        if(global.console && typeof global.console.error === 'function'){
          global.console.error('requestAnimationFrame callback error', err);
        }
      }
    }, 16);
  }
  if(typeof global.cancelAnimationFrame !== 'function'){
    global.cancelAnimationFrame = (handle)=>global.clearTimeout(handle);
  }

  /* ================= Basis-Einstellungen ================ */
  const META_DEFAULTS = {
    version: 'dev',
    buildDate: '2025-11-10',
    defaultFile: 'Artikelpreisliste.xlsx',
    defaultFilePath: './data/Artikelpreisliste.xlsx',
    buildSource: './data/Artikelpreisliste.xlsx'
  };

  const APP_META = Object.assign({}, META_DEFAULTS, global.APP_META || {});
  if(!APP_META.defaultFilePath){
    APP_META.defaultFilePath = `./data/${APP_META.defaultFile || META_DEFAULTS.defaultFile}`;
  }
  if(!APP_META.buildSource){
    APP_META.buildSource = APP_META.defaultFilePath;
  }

  function versionedAsset(path){
    if(!path) return path;
    const versionValue = APP_META.version || META_DEFAULTS.version;
    const encoded = encodeURIComponent(versionValue);
    return path.includes('?') ? `${path}&v=${encoded}` : `${path}?v=${encoded}`;
  }

  const APP_VERSION = APP_META.version || META_DEFAULTS.version;
  const APP_BUILD_DATE = APP_META.buildDate || META_DEFAULTS.buildDate;
  const DEFAULT_FILE = APP_META.defaultFile || META_DEFAULTS.defaultFile;
  const DEFAULT_FILE_PATH = APP_META.defaultFilePath;
  const APP_BUILD_SOURCE = APP_META.buildSource;
  const DEFAULT_FILE_URL = versionedAsset(DEFAULT_FILE_PATH);
  const VERSION_CHECK_INTERVAL = 60_000;

  /* ================= Hilfsfunktionen ==================== */
  function $(q){return document.querySelector(q);}
  function escapeHtml(s){return String(s).replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;').replace(/\"/g,'&quot;').replace(/'/g,'&#39;');}
  function cssEscape(value){
    const string = String(value ?? '');
    if(typeof CSS !== 'undefined' && CSS && typeof CSS.escape === 'function'){
      return CSS.escape(string);
    }
    const length = string.length;
    let index = -1;
    let codeUnit;
    let result = '';
    const firstCodeUnit = length > 0 ? string.charCodeAt(0) : 0;
    while(++index < length){
      codeUnit = string.charCodeAt(index);
      if(codeUnit === 0){
        result += '\uFFFD';
        continue;
      }
      if(
        (codeUnit >= 0x0001 && codeUnit <= 0x001F) ||
        codeUnit === 0x007F ||
        (index === 0 && codeUnit >= 0x0030 && codeUnit <= 0x0039) ||
        (index === 1 && codeUnit >= 0x0030 && codeUnit <= 0x0039 && firstCodeUnit === 0x002D)
      ){
        result += '\\' + codeUnit.toString(16) + ' ';
        continue;
      }
      if(index === 0 && codeUnit === 0x002D && length === 1){
        result += '\\-';
        continue;
      }
      if(
        codeUnit >= 0x0080 ||
        codeUnit === 0x002D ||
        codeUnit === 0x005F ||
        (codeUnit >= 0x0030 && codeUnit <= 0x0039) ||
        (codeUnit >= 0x0041 && codeUnit <= 0x005A) ||
        (codeUnit >= 0x0061 && codeUnit <= 0x007A)
      ){
        result += string.charAt(index);
        continue;
      }
      result += '\\' + string.charAt(index);
    }
    return result;
  }
  function isGroupId(id){ const s=String(id||"").replace(/\D/g,""); if(!s) return false; const n=parseInt(s,10); return Number.isFinite(n) && n%100===0; }
  function parseEuro(str){ if(str==null) return NaN; let s=String(str).trim(); if(!s) return NaN; if(s.includes(',')) s=s.replace(/\./g,'').replace(',', '.'); return Number(s); }
  function fmtPrice(v){
    const n=parseEuro(v);
    return Number.isFinite(n)
      ? n.toLocaleString('de-AT',{style:'currency',currency:'EUR',minimumFractionDigits:2,maximumFractionDigits:2})
      : (v??'');
  }
  function debounced(fn,ms){ let t; return (...a)=>{ clearTimeout(t); t=setTimeout(()=>fn(...a),ms); }; }
  let toastTimer=null;
  let statusTimer=null;
  function setStatus(kind, html, ttl=3000){
    const el = $('#hint');
    if(!kind){ el.textContent=''; el.removeAttribute('data-status'); if(statusTimer){clearTimeout(statusTimer);} return; }
    el.dataset.status = kind; el.innerHTML = html || '';
    if(statusTimer){ clearTimeout(statusTimer); }
    statusTimer = setTimeout(()=>{ el.textContent=''; el.removeAttribute('data-status'); }, ttl);
  }

  function triggerUpdatedBadge(){
    const toast = $('#toast');
    if(!toast) return;
    toast.classList.add('show');
    if(toastTimer){ clearTimeout(toastTimer); }
    toastTimer = setTimeout(()=>toast.classList.remove('show'), 1600);
  }
  function isSonderEditable(id){ const s=String(id||'').replace(/\D/g,''); return /(?:98|99)$/.test(s); }

  function startVersionWatcher(){
    const metaUrl = global.APP_META_URL || './data/app-meta.json';
    if(!metaUrl) return;
    let currentVersion = APP_VERSION || META_DEFAULTS.version;

    async function checkForUpdates(){
      try{
        const res = await fetch(metaUrl + '?ts=' + Date.now(), {cache:'no-store'});
        if(!res.ok) return;
        const data = await res.json();
        if(data && data.version && data.version !== currentVersion){
          currentVersion = data.version;
          const toast = document.getElementById('toast');
          if(toast){
            toast.textContent = 'Neue Version verfügbar – Seite wird neu geladen…';
            toast.classList.add('show');
          }
          setTimeout(()=>global.location.reload(), 1200);
        }
      }catch(err){
        console.warn('Versionsprüfung fehlgeschlagen', err);
      }
    }

    global.setInterval(checkForUpdates, VERSION_CHECK_INTERVAL);
  }

  const DIACRITICS_RE = /[\u0300-\u036f]/g;
  function normalizeText(value){
    if(value == null) return '';
    try{
      return String(value).normalize('NFD').replace(DIACRITICS_RE,'').toLowerCase();
    }catch{
      return String(value).toLowerCase();
    }
  }

  const QTY_RE=/^[-+]?(?:\d+(?:[.]\d*)?|[.]\d+)$/;
  function parseQty(str){
    if(str==null) return 0;
    let s = String(str).trim();
    if(!s) return 0;
    s = s.replace(',', '.');
    if(!QTY_RE.test(s)) return NaN;
    const n = Number(s);
    return Number.isFinite(n) ? n : NaN;
  }
  function fmtQty(n){ return Number.isFinite(n)? n.toLocaleString('de-AT',{minimumFractionDigits:0, maximumFractionDigits:2}) : ''; }

  function escRe(s){ return s.replace(/[.*+?^${}()|[\]\\]/g,'\\$&'); }
  function splitTerms(s){ if(!s) return []; return String(s).trim().split(/\s+/).filter(Boolean); }
  const REGEX_CACHE_LIMIT = 200;
  const regexCache = new Map();
  function makeRegex(terms){
    const normalized = Array.isArray(terms) ? terms.filter(Boolean) : [];
    const key = normalized.slice().sort().join('\u0000');
    if(!key){
      regexCache.delete('');
      return null;
    }
    if(regexCache.has(key)){
      return regexCache.get(key);
    }
    const orderedTerms = [...new Set(normalized)].sort((a,b)=>b.length-a.length);
    if(!orderedTerms.length){
      regexCache.delete(key);
      return null;
    }
    const rx = new RegExp('(' + orderedTerms.map(escRe).join('|') + ')','gi');
    regexCache.set(key, rx);
    while(regexCache.size > REGEX_CACHE_LIMIT){
      const oldestKey = regexCache.keys().next().value;
      if(oldestKey === undefined) break;
      regexCache.delete(oldestKey);
    }
    return rx;
  }
  function hi(text, terms){
    const s = String(text ?? '');
    if(!s) return '';
    const rx = makeRegex(terms);
    if(!rx) return escapeHtml(s);
    const esc = escapeHtml(s);
    return esc.replace(rx, '<mark class="hl">$1</mark>');
  }
  const BASE_LOCATION = (global.location && typeof global.location.href === 'string')
    ? global.location.href
    : 'https://localhost/';

  const EMAIL_RE = /^[\w.+-]+@[A-Za-z0-9.-]+\.[A-Za-z]{2,}$/;

  function toSafeHref(raw){
    if(!raw) return '';
    let s = String(raw).trim();
    if(!s) return '';
    s = s.replace(/[\u0000-\u001F\u007F]/g, '');
    if(!s) return '';

    if(/^mailto:/i.test(s)){
      const addrPart = s.slice(7);
      const [addrRaw, queryRaw=''] = addrPart.split('?');
      if(!EMAIL_RE.test(addrRaw)) return '';
      const query = queryRaw ? '?' + encodeURI(queryRaw) : '';
      return 'mailto:' + addrRaw + query;
    }

    if(EMAIL_RE.test(s)){
      return 'mailto:' + s;
    }

    const ensureHttpUrl = (value)=>{
      try{
        if(typeof URL === 'function'){
          const url = new URL(value, BASE_LOCATION);
          if(url.protocol === 'http:' || url.protocol === 'https:'){
            return url.href;
          }
          return '';
        }
      }catch(err){
        try{
          const encoded = encodeURI(value);
          if(typeof URL === 'function'){
            const url = new URL(encoded, BASE_LOCATION);
            if(url.protocol === 'http:' || url.protocol === 'https:'){
              return url.href;
            }
          }
          return /^(https?:)?\/\//i.test(value) ? encoded : '';
        }catch{}
        return '';
      }
      return /^(https?:)?\/\//i.test(value) ? encodeURI(value) : '';
    };

    if(/^https?:\/\//i.test(s)){
      return ensureHttpUrl(s);
    }

    if(/^\/\//.test(s)){
      return ensureHttpUrl('https:' + s);
    }

    if(/^www\./i.test(s)){
      return ensureHttpUrl('https://' + s);
    }

    return '';
  }
  function linkify(escapedHtml){
    if(!escapedHtml) return '';
    const URL_RE = /\b((https?:\/\/|www\.)[^\s<]+[^\s<\.)])/gi;
    const MAIL_RE = /\b[A-Z0-9._%+-]+@[A-Z0-9.-]+\.[A-Z]{2,}\b/gi;
    return escapedHtml.split(/(<[^>]+>)/g).map(part=>{
      if(part.startsWith('<')) return part;
      return part
        .replace(URL_RE, (m)=>{
          const href = toSafeHref(m);
          if(!href) return m;
          return `<a href="${href}" target="_blank" rel="noopener noreferrer">${m}</a>`;
        })
        .replace(MAIL_RE, (m)=>{
          const href = toSafeHref(m);
          return `<a href="${href}" target="_blank" rel="noopener noreferrer">${m}</a>`;
        });
    }).join('');
  }
  function hlAndLink(text, terms){
    const highlighted = hi(text, terms);
    return linkify(highlighted);
  }

  function applyVersionInfo(){
    const badge=document.getElementById('versionBadge');
    if(badge) badge.textContent=APP_VERSION;
    const chip=document.getElementById('versionChip');
    if(chip) chip.title=`Build: ${APP_BUILD_DATE} · Quelle: ${APP_BUILD_SOURCE}`;
  }

  HaasApp.meta = {
    META_DEFAULTS,
    APP_META,
    APP_VERSION,
    APP_BUILD_DATE,
    DEFAULT_FILE,
    DEFAULT_FILE_PATH,
    DEFAULT_FILE_URL,
    APP_BUILD_SOURCE,
    VERSION_CHECK_INTERVAL,
    versionedAsset,
    startVersionWatcher,
    applyVersionInfo,
  };

  HaasApp.utils = Object.assign({}, HaasApp.utils || {}, {
    $, escapeHtml, cssEscape, isGroupId, parseEuro, fmtPrice, debounced,
    setStatus, triggerUpdatedBadge, isSonderEditable, normalizeText,
    parseQty, fmtQty, hi, linkify, hlAndLink, splitTerms, makeRegex,
    toSafeHref
  });

  Object.assign(global, {
    META_DEFAULTS,
    APP_META,
    APP_VERSION,
    APP_BUILD_DATE,
    DEFAULT_FILE,
    DEFAULT_FILE_PATH,
    DEFAULT_FILE_URL,
    APP_BUILD_SOURCE,
    VERSION_CHECK_INTERVAL,
    versionedAsset,
    $, escapeHtml, cssEscape, isGroupId, parseEuro, fmtPrice, debounced,
    setStatus, triggerUpdatedBadge, isSonderEditable, normalizeText,
    parseQty, fmtQty, hi, linkify, hlAndLink, splitTerms, makeRegex,
    toSafeHref, startVersionWatcher, applyVersionInfo
  });

})(window);
