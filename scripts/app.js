/* ================= Basis-Einstellungen ================ */
const APP_VERSION = 'V1.1.0';
const APP_BUILD_DATE = '2024-06-05';
const DEFAULT_FILE = 'Artikelpreisliste.xlsx';
const DEFAULT_FILE_PATH = `./data/${DEFAULT_FILE}`;
const APP_BUILD_SOURCE = DEFAULT_FILE_PATH;
const DEFAULT_FILE_URL = `${DEFAULT_FILE_PATH}?v=${APP_VERSION}`;

/* ================= Hilfsfunktionen ==================== */
function $(q){return document.querySelector(q)}
function escapeHtml(s){return String(s).replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;').replace(/\"/g,'&quot;').replace(/'/g,'&#39;')}
function isGroupId(id){ const s=String(id||"").replace(/\D/g,""); if(!s) return false; const n=parseInt(s,10); return Number.isFinite(n) && n%100===0; }
function parseEuro(str){ if(str==null) return NaN; let s=String(str).trim(); if(!s) return NaN; if(s.includes(',')) s=s.replace(/\./g,'').replace(',', '.'); return Number(s); }
function fmtPrice(v){
  const n=parseEuro(v);
  return Number.isFinite(n)
    ? n.toLocaleString('de-AT',{style:'currency',currency:'EUR',minimumFractionDigits:2,maximumFractionDigits:2})
    : (v??'');
}
function debounced(fn,ms){ let t; return (...a)=>{ clearTimeout(t); t=setTimeout(()=>fn(...a),ms); } }
let toastTimer=null;
let statusTimer=null;
function setStatus(kind, html, ttl=3000){
  const el = $('#hint');
  if(!kind){ el.textContent=''; el.removeAttribute('data-status'); if(statusTimer){clearTimeout(statusTimer)}; return; }
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

const DIACRITICS_RE = /[\u0300-\u036f]/g;
function normalizeText(value){
  if(value == null) return '';
  try{
    return String(value).normalize('NFD').replace(DIACRITICS_RE,'').toLowerCase();
  }catch{
    return String(value).toLowerCase();
  }
}

/* Menge parsing (beliebiger Betrag) */
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

/* ====== Hervorhebung (PDF-Style) + Linkify ====== */
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
function toSafeHref(raw){
  if(!raw) return '';
  let s = String(raw).trim();
  if(/^mailto:/i.test(s)) return s;
  if(/^[\w.+-]+@[\w.-]+\.[A-Za-z]{2,}$/.test(s)) return 'mailto:'+s;
  if(!/^(https?:)?\/\//i.test(s)){
    if(/^www\./i.test(s)) s = 'http://' + s;
    else if (/^[A-Za-z][A-Za-z0-9+.-]*:/.test(s)) return '';
  }
  if(!/^https?:\/\//i.test(s)) s = s;
  return s;
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

/* ================== State ============================= */
let GROUPS=[], LAST_WB=null;
let DRAWER_OPEN = false;
let CURRENT_SHEET = '';
const BASKET_STATE = {
  items: [],
  mergeSameItems: false,
};
const MANUAL_DEFAULT_ARTNR = '0000';
const MANUAL_DEFAULT_TITLE = 'Basispreis Haus';
const MANUAL_DEFAULT_UNIT = 'pau';
const MANUAL_DESCRIPTION_PLACEHOLDER = 'Ausbaustufe, Fensterelemente, etc.';
let GROUP_SWITCH_ANIM = Promise.resolve();
const ROW_STATES = new Map();
let BASKET_LINE_COUNTER = 1;

function nextLineId(){
  return `ln_${BASKET_LINE_COUNTER++}`;
}

function cloneBasketItem(payload, options={}){
  if(!payload || !payload.id){ return null; }
  const qtyRaw = Number.isFinite(payload.qtyNum) ? payload.qtyNum : parseQty(payload.qtyNum);
  if(!Number.isFinite(qtyRaw) || qtyRaw === 0){ return null; }
  const preisRaw = Number.isFinite(payload.preisNum) ? payload.preisNum : parseEuro(payload.preisNum);
  const item = {
    lineId: options.forceNewLineId ? nextLineId() : (payload.lineId || nextLineId()),
    id: payload.id,
    kurz: payload.kurz ?? '',
    beschreibung: payload.beschreibung ?? '',
    eh: payload.eh ?? '',
    preisNum: Number.isFinite(preisRaw) ? preisRaw : 0,
    qtyNum: qtyRaw,
    isAlternative: !!payload.isAlternative,
  };
  item.totalNum = item.preisNum * item.qtyNum;
  return item;
}

function addToBasket(payload){
  const item = cloneBasketItem(payload, {forceNewLineId:true});
  if(!item){ return null; }
  BASKET_STATE.items.push(item);
  return item;
}

function createManualBasketItem(){
  return {
    lineId: nextLineId(),
    id: MANUAL_DEFAULT_ARTNR,
    kurz: MANUAL_DEFAULT_TITLE,
    beschreibung: '',
    eh: MANUAL_DEFAULT_UNIT,
    preisNum: 0,
    qtyNum: 0,
    totalNum: 0,
    isManual: true,
    isAlternative: false,
  };
}

function displayArtnr(item){
  if(!item) return '';
  const id = item.id ?? '';
  return item.isAlternative ? `A-${id}` : id;
}

function addManualTopItem(){
  const manualItem = createManualBasketItem();
  BASKET_STATE.items.unshift(manualItem);
  setDrawer(true);
  renderSummary(false, {focusLineId: manualItem.lineId, focusField: 'title'});
  triggerUpdatedBadge();
  setStatus('ok','Manuelle Position hinzugefügt.',2000);
  return manualItem;
}

function setLineQty(lineId, qtyValue, options){
  const idx = BASKET_STATE.items.findIndex(item=>item.lineId===lineId);
  if(idx<0){
    return {status:'missing', item:null};
  }
  const existing = BASKET_STATE.items[idx];
  const raw = Number.isFinite(qtyValue) ? qtyValue : (qtyValue ?? '');
  const rawStr = typeof raw === 'number' ? String(raw) : String(raw).trim();

  if(rawStr === ''){
    if(options?.commit === false){
      return {status:'empty', item:{...existing, qtyNum:NaN, totalNum:NaN}};
    }
    return {status:'empty', item:existing};
  }

  const qtyNum = Number.isFinite(qtyValue) ? qtyValue : parseQty(rawStr);
  if(!Number.isFinite(qtyNum)){
    return {status:'invalid', item:existing};
  }

  const nextItem = {...existing, qtyNum};
  nextItem.totalNum = nextItem.preisNum * nextItem.qtyNum;

  if(options?.commit === false){
    return {status:'preview', item:nextItem};
  }

  if(qtyNum === 0){
    BASKET_STATE.items.splice(idx,1);
    return {status:'removed', item:existing};
  }

  BASKET_STATE.items[idx] = nextItem;
  return {status:'updated', item:nextItem};
}

const updateLineQty = setLineQty;

function removeLine(lineId){
  const idx = BASKET_STATE.items.findIndex(item=>item.lineId===lineId);
  if(idx<0) return false;
  BASKET_STATE.items.splice(idx,1);
  return true;
}

function findBasketItem(lineId){
  return BASKET_STATE.items.find(item=>item.lineId===lineId);
}

function setAltFlag(lineId, checked){
  const item = findBasketItem(lineId);
  if(!item) return;
  const next = !!checked;
  if(item.isAlternative === next) return;
  item.isAlternative = next;
  const wrap = $('#summaryTableWrap');
  if(wrap){
    const safeId = (typeof CSS!=='undefined' && CSS && typeof CSS.escape==='function') ? CSS.escape(lineId) : lineId.replace(/"/g,'\\"');
    const row = wrap.querySelector(`[data-line-id="${safeId}"]`);
    if(row){
      row.dataset.alt = next ? 'true' : 'false';
      row.classList.toggle('alt-item', next);
      const artCell = row.querySelector('[data-role="artnr"]');
      if(artCell){ artCell.textContent = displayArtnr(item); }
      const totalCell = row.querySelector('[data-role="line-total"]');
      if(totalCell){ updateLineTotalCell(totalCell, getLineTotalValue(item), next); }
    }
    const toggle = wrap.querySelector(`input.alt-toggle-basket[data-line-id="${safeId}"]`);
    if(toggle){ toggle.checked = next; }
  }
  const sums = computeBasketSums();
  updateSummaryDisplay(sums, BASKET_STATE.items.length, true);
  triggerUpdatedBadge();
}

function clearBasket(){
  BASKET_STATE.items.length = 0;
  BASKET_LINE_COUNTER = 1;
}

function getBasketItems(){
  return BASKET_STATE.items.slice();
}

function getDisplayOrderedItems(){
  const items = getBasketItems();
  if(items.length <= 1){
    return items;
  }
  const positions = new Map();
  BASKET_STATE.items.forEach((it, idx)=>positions.set(it.lineId, idx));
  return items.sort((a,b)=>{
    const aManual = !!a.isManual;
    const bManual = !!b.isManual;
    if(aManual !== bManual){
      return aManual ? -1 : 1;
    }
    if(aManual && bManual){
      const posA = positions.get(a.lineId) ?? 0;
      const posB = positions.get(b.lineId) ?? 0;
      return posA - posB;
    }
    const byId = String(a.id).localeCompare(String(b.id),'de',{numeric:true,sensitivity:'base'});
    if(byId !== 0) return byId;
    return String(a.lineId).localeCompare(String(b.lineId),'de',{numeric:true,sensitivity:'base'});
  });
}

function getBasketSize(){
  return BASKET_STATE.items.length;
}

const VIRTUAL = {
  items:[],
  container:null,
  viewport:null,
  topSpacer:null,
  bottomSpacer:null,
  nodes:new Map(),
  heightCache:[],
  averageHeight:56,
  renderStart:0,
  renderEnd:0,
  pending:false,
  lastFilter:null,
  boundScroll:false,
  active:false,
};
const VIRTUAL_THRESHOLD = 1800;

function ensureRowState(row){
  const key=row.id;
  if(!ROW_STATES.has(key)){
    ROW_STATES.set(key,{
      qty:'',
      preis: row.preis!=null?String(row.preis):'',
      kurz: row.kurz_raw ?? '',
      beschreibung: row.beschreibung_raw ?? '',
      einheit: row.einheit ?? '',
      einheitInfo: row.einheitInfo ?? ''
    });
  }
  return ROW_STATES.get(key);
}

/* ================= Excel lesen ======================== */
function extractRows(ws){
  try{
    const ref = ws['!ref']; if(!ref) throw new Error('Kein Zellbereich (!ref).');
    const range = XLSX.utils.decode_range(ref);
    const out = [];
    for(let r=range.s.r;r<=range.e.r;r++){
      const C = (c)=>ws[XLSX.utils.encode_cell({r,c})];
      const A=C(0),B=C(1),C2=C(2),D=C(3),E=C(4),F=C(5),G=C(6);
      const id  = A&&A.v!=null?String(A.v).trim():'';
      const kurz= B&&B.v!=null?String(B.v).trim():'';
      const beschr = C2&&C2.v!=null?String(C2.v).trim():'';
      if(!(id||kurz||beschr)) continue;

      const linkB = (B && B.l && B.l.Target) ? String(B.l.Target) : '';
      const linkC = (C2 && C2.l && C2.l.Target) ? String(C2.l.Target) : '';
      const linkG = (G && G.l && G.l.Target) ? String(G.l.Target) : '';

      const row = {
        id,
        kurz_raw:kurz,
        beschreibung_raw:beschr,
        styleB:{ bold:!!(B&&B.s&&B.s.font&&B.s.font.bold), underline:!!(B&&B.s&&B.s.font&&(B.s.font.underline||B.s.font.u)) },
        styleC:{ bold:!!(C2&&C2.s&&C2.s.font&&C2.s.font.bold), underline:!!(C2&&C2.s&&C2.s.font&&(C2.s.font.underline||C2.s.font.u)) },
        einheit:D&&D.v!=null?String(D.v).trim():'',
        einheitInfo:E&&E.v!=null?String(E.v).trim():'',
        preis:F?F.v:'',
        hinweis_raw:G&&G.v!=null?String(G.v).trim():'',
        linkB,
        linkC,
        linkG
      };
      row.norm = {
        id: normalizeText(row.id),
        kurz: normalizeText(row.kurz_raw),
        beschreibung: normalizeText(row.beschreibung_raw),
        einheit: normalizeText(row.einheit),
        einheitInfo: normalizeText(row.einheitInfo),
        preis: normalizeText(row.preis),
        hinweis: normalizeText(row.hinweis_raw),
        haystack: normalizeText([row.id,row.kurz_raw,row.beschreibung_raw,row.einheit,row.einheitInfo,row.preis,row.hinweis_raw].join(' \u2022 '))
      };
      out.push(row);
    }
    return out;
  }catch{
    const json=XLSX.utils.sheet_to_json(ws,{defval:""}); if(!json.length) return [];
    const k=Object.keys(json[0]);
    return json.map(r=>{
      const row={ id:String(r[k[0]]||'').trim(), kurz_raw:String(r[k[1]]||'').trim(), beschreibung_raw:String(r[k[2]]||'').trim(),
        einheit:String(r[k[3]]||'').trim(), einheitInfo:String(r[k[4]]||'').trim(), preis:r[k[5]], hinweis_raw:String(r[k[6]]||'').trim(),
        styleB:null, styleC:null, linkB:'', linkC:'', linkG:'' };
      row.norm = {
        id: normalizeText(row.id),
        kurz: normalizeText(row.kurz_raw),
        beschreibung: normalizeText(row.beschreibung_raw),
        einheit: normalizeText(row.einheit),
        einheitInfo: normalizeText(row.einheitInfo),
        preis: normalizeText(row.preis),
        hinweis: normalizeText(row.hinweis_raw),
        haystack: normalizeText([row.id,row.kurz_raw,row.beschreibung_raw,row.einheit,row.einheitInfo,row.preis,row.hinweis_raw].join(' \u2022 '))
      };
      return row;
    });
  }
}
function buildGroups(rows){
  const groups=[]; let cur=null;
  for(const r of rows){
    if(isGroupId(r.id)){
      cur={groupId:r.id,title:(r.kurz_raw||r.beschreibung_raw),normTitle:normalizeText(r.kurz_raw||r.beschreibung_raw),children:[]};
      groups.push(cur);
    }
    else if(cur){ cur.children.push(r); }
  }
  return groups;
}

/* ================= Render ============================= */
function autoGrow(el){ el.style.height='auto'; el.style.height = (el.scrollHeight)+'px'; }

function trGroup(g, f){
  const tr=document.createElement('tr');
  tr.className='group';
  tr.id = 'grp_'+g.groupId;
  const title = f && f.rawQ ? hlAndLink(g.title||'', f.terms) : escapeHtml(g.title||'');
  tr.innerHTML=`<td colspan="9"><strong>${escapeHtml(g.groupId)} – ${title}</strong></td>`;
  return tr;
}

function trChild(c, f){
  const tr=document.createElement('tr');
  const editable=isSonderEditable(c.id);
  const state=ensureRowState(c);
  const preisSource = editable ? state.preis : c.preis;
  let preisNum=parseEuro(preisSource);
  if(!Number.isFinite(preisNum)) preisNum=parseEuro(c.preis);

  const qTerms = f ? f.terms : [];
  const tID = [...qTerms, f?.rawA].filter(Boolean);
  const tK  = [...qTerms, f?.rawB].filter(Boolean);
  const tB  = [...qTerms, f?.rawC].filter(Boolean);
  const tG  = qTerms;

  const kurzValue = editable ? state.kurz : c.kurz_raw;
  const beschrValue = editable ? state.beschreibung : c.beschreibung_raw;
  const ehValue = editable ? state.einheit : c.einheit;
  const ehInfoValue = editable ? state.einheitInfo : c.einheitInfo;

  const kurzStatic = c.linkB
    ? `<div class="desc"><a href="${escapeHtml(toSafeHref(c.linkB))}" target="_blank" rel="noopener noreferrer">${hi(kurzValue, tK)}</a></div>`
    : `<div class="desc">${hlAndLink(kurzValue, tK)}</div>`;

  const beschrStatic = c.linkC
    ? `<div class="desc"><a href="${escapeHtml(toSafeHref(c.linkC))}" target="_blank" rel="noopener noreferrer">${hi(beschrValue, tB)}</a></div>`
    : `<div class="desc">${hlAndLink(beschrValue, tB)}</div>`;

  const ehHTML = editable
    ? `<input class="cell-edit" data-field="einheit" value="${escapeHtml(ehValue||'')}" />`
    : hlAndLink(ehValue||'', tG);

  const ehInfoHTML = editable
    ? `<input class="cell-edit" data-field="einheitInfo" value="${escapeHtml(ehInfoValue||'')}" />`
    : hlAndLink(ehInfoValue||'', tG);

  const preisHTML = editable
    ? `<input class="cell-edit price" data-field="preis" inputmode="decimal" placeholder="0" value="${escapeHtml((state.preis||'').replace('.',','))}" />`
    : hlAndLink(fmtPrice(preisSource), tG);

  const hinweisHTML = c.linkG
    ? `<div class="desc"><a href="${escapeHtml(toSafeHref(c.linkG))}" target="_blank" rel="noopener noreferrer">${hi(c.hinweis_raw||'', tG)}</a></div>`
    : `<div class="desc">${hlAndLink(c.hinweis_raw||'', tG)}</div>`;

  const kurzHTML = editable
    ? `<textarea class="cell-edit ta" data-field="kurz">${escapeHtml(state.kurz||'')}</textarea>`
    : kurzStatic;

  const beschrHTML = editable
    ? `<textarea class="cell-edit ta" data-field="beschreibung">${escapeHtml(state.beschreibung||'')}</textarea>`
    : beschrStatic;

  const qtyId=`q_${c.id}`, addBtnId=`add_${c.id}`;
  const qtyValue = state.qty || '';
  const controlsHTML = `<div class="qty-action">
      <input id="${qtyId}" class="qty" type="text" inputmode="decimal" placeholder="0" title="Menge" value="${escapeHtml(qtyValue)}" />
      <label class="alt-flag" title="Als Alternative markieren">
        <input type="checkbox" class="alt-toggle" aria-label="Als Alternative markieren" />
      </label>
      <button id="${addBtnId}" type="button" class="btn-plus" title="Zur Zusammenfassung hinzufügen">➕</button>
    </div>`;

  tr.innerHTML = `
    <td>${hlAndLink(c.id, tID)}</td>
    <td>${kurzHTML}</td>
    <td>${beschrHTML}</td>
    <td>${ehHTML}</td>
    <td>${ehInfoHTML}</td>
    <td class="right" data-sort="${preisNum}">${preisHTML}</td>
    <td class="control-cell">${controlsHTML}</td>
    <td class="right" data-total="0">–</td>
    <td class="desc">${hinweisHTML}</td>`;

  tr.querySelectorAll('textarea.cell-edit.ta').forEach(ta=>{
    autoGrow(ta);
    ta.addEventListener('input',()=>{
      autoGrow(ta);
      const field = ta.dataset.field;
      if(field==='kurz'){ state.kurz=ta.value; }
      if(field==='beschreibung'){ state.beschreibung=ta.value; }
    });
  });

  const qtyInp=tr.querySelector('input.qty');
  const preisInp=tr.querySelector('input[data-field="preis"]');
  const totalCell=tr.querySelector('[data-total]');
  const addBtn=tr.querySelector('.btn-plus');
  const altToggle=tr.querySelector('.alt-toggle');
  let addFeedbackTimer=null;
  let pendingAddReset=null;

  function resetAddButton(){
    addBtn.classList.remove('added','removed');
    addBtn.textContent='➕';
    addFeedbackTimer=null;
  }

  function showAddButtonFeedback(kind){
    if(addFeedbackTimer){
      clearTimeout(addFeedbackTimer);
      addFeedbackTimer=null;
    }
    if(kind==='removed'){
      addBtn.classList.remove('added');
      addBtn.classList.add('removed');
      addBtn.textContent='entfernt ✖';
      pendingAddReset=null;
      addFeedbackTimer=setTimeout(()=>{
        resetAddButton();
        addFeedbackTimer=null;
      },2400);
      return;
    }
    addBtn.classList.remove('removed');
    addBtn.classList.add('added');
    addBtn.textContent='✓';
    const resetToken=Symbol('added-feedback');
    pendingAddReset={token:resetToken, qtyChanged:false, altChanged:false};
    addFeedbackTimer=setTimeout(()=>{
      if(pendingAddReset && pendingAddReset.token===resetToken){
        if(!pendingAddReset.qtyChanged && qtyInp){
          qtyInp.value='';
          qtyInp.placeholder='0';
          state.qty='';
          recalcRowTotal();
        }
        if(!pendingAddReset.altChanged && altToggle){
          altToggle.checked=false;
          altToggle.indeterminate=false;
        }
        pendingAddReset=null;
      }
      resetAddButton();
      addFeedbackTimer=null;
    }, 2500);
  }

  const einheitInp=tr.querySelector('input[data-field="einheit"]');
  const einheitInfoInp=tr.querySelector('input[data-field="einheitInfo"]');

  function currentPreis(){
    if(preisInp){
      const p=parseEuro(preisInp.value);
      return Number.isFinite(p)?p:0;
    }
    return Number.isFinite(preisNum)?preisNum:0;
  }
  function currentKurz(){ return editable ? (tr.querySelector('[data-field="kurz"]')?.value.trim()||'') : (c.kurz_raw||''); }
  function currentBeschr(){ return editable ? (tr.querySelector('[data-field="beschreibung"]')?.value.trim()||'') : (c.beschreibung_raw||''); }
  function currentEH(){ return editable ? (einheitInp?.value.trim()||'') : (c.einheit||''); }

  function updateTotalCell(q,p){
    if(Number.isNaN(q)||q===0){
      totalCell.textContent='–';
      totalCell.dataset.total='0';
      totalCell.classList.toggle('neg',false);
    }else{
      const total=p*q;
      totalCell.textContent=fmtPrice(total);
      totalCell.dataset.total=String(total);
      totalCell.classList.toggle('neg',total<0);
    }
  }

  function buildBasketPayload(){
    const q=parseQty(qtyInp?.value??'');
    if(Number.isNaN(q)||q===0){ return null; }
    const p=currentPreis();
    return {
      id:c.id,
      kurz:currentKurz(),
      beschreibung:currentBeschr(),
      eh:currentEH(),
      preisNum:p,
      qtyNum:q,
      isAlternative: !!altToggle?.checked,
    };
  }

  function recalcRowTotal(){
    const qRaw = qtyInp?.value ?? '';
    state.qty = qRaw;
    const q=parseQty(qRaw);
    const p=currentPreis();
    updateTotalCell(q,p);
  }

  qtyInp?.addEventListener('input',()=>{
    qtyInp.value=qtyInp.value.replace(/[^\d.,-]/g,'').replace(/(?!^)-/g,'');
    state.qty=qtyInp.value;
    if(pendingAddReset){
      pendingAddReset.qtyChanged=true;
    }
    recalcRowTotal();
  });

  qtyInp?.addEventListener('change',()=>{
    if(pendingAddReset){
      pendingAddReset.qtyChanged=true;
    }
  });

  altToggle?.addEventListener('change',()=>{
    if(pendingAddReset){
      pendingAddReset.altChanged=true;
    }
  });

  if(preisInp){
    preisInp.addEventListener('input',()=>{
      if(!/^[\d.,-]*$/.test(preisInp.value)){
        preisInp.value=preisInp.value.replace(/[^\d.,-]/g,'');
      }
      state.preis=preisInp.value;
      preisInp.closest('td').dataset.sort=String(currentPreis());
      recalcRowTotal();
    });
  }

  einheitInp?.addEventListener('input',()=>{
    state.einheit=einheitInp.value;
  });
  einheitInfoInp?.addEventListener('input',()=>{
    state.einheitInfo=einheitInfoInp.value;
  });

  addBtn.addEventListener('click',()=>{
    const payload=buildBasketPayload();
    if(!payload){
      setStatus('warn','Bitte zuerst eine gültige Menge (≠ 0) eingeben.',3500);
      if(addBtn.animate){
        addBtn.animate([{transform:'scale(1)'},{transform:'scale(1.08)'},{transform:'scale(1)'}],{duration:160});
      }
      return;
    }
    const added=addToBasket(payload);
    if(!added){
      setStatus('warn','Element konnte nicht hinzugefügt werden.',3000);
      return;
    }
    if(added.removed){
      showAddButtonFeedback('removed');
      renderSummary(true);
      triggerUpdatedBadge();
      setStatus('ok','Bereit.',1500);
      return;
    }
    showAddButtonFeedback('added');
    renderSummary(true);
    triggerUpdatedBadge();
    setStatus('ok','Bereit.',1500);
  });

  resetAddButton();

  recalcRowTotal();
  return tr;
}

function currentFilters(){
  const rawQ=$('#search').value||'';
  const rawA=$('#fA').value||'';
  const rawB=$('#fB').value||'';
  const rawC=$('#fC').value||'';
  return {
    rawQ,
    rawA,
    rawB,
    rawC,
    normQ:normalizeText(rawQ),
    normA:normalizeText(rawA),
    normB:normalizeText(rawB),
    normC:normalizeText(rawC),
    group:$('#groupFilter').value||'',
    terms:splitTerms(rawQ)
  };
}
function matchesFilters(row,f){
  const norm=row?.norm||{};
  const global = !f.normQ || (norm.haystack||'').includes(f.normQ);
  const spec = (!f.normA || (norm.id||'').includes(f.normA))
    && (!f.normB || (norm.kurz||'').includes(f.normB))
    && (!f.normC || (norm.beschreibung||'').includes(f.normC));
  return global && spec;
}

function createSpacer(){
  const tr=document.createElement('tr');
  tr.className='virtual-spacer';
  tr.setAttribute('aria-hidden','true');
  tr.innerHTML='<td colspan="9" style="padding:0;border:none;height:0;border:none"></td>';
  return tr;
}
function setSpacerHeight(spacer,height){
  if(!spacer) return;
  const cell=spacer.firstElementChild;
  if(cell){ cell.style.height = Math.max(height,0)+'px'; }
}
const virtualViewportObserver = typeof ResizeObserver==='function' ? new ResizeObserver(()=>scheduleVirtualUpdate(false)) : null;
function ensureVirtualSetup(){
  if(!VIRTUAL.container){ VIRTUAL.container=document.getElementById('rows'); }
  if(!VIRTUAL.viewport){ VIRTUAL.viewport=document.getElementById('tableWrap'); }
  if(VIRTUAL.container && !VIRTUAL.topSpacer){ VIRTUAL.topSpacer=createSpacer(); }
  if(VIRTUAL.container && !VIRTUAL.bottomSpacer){ VIRTUAL.bottomSpacer=createSpacer(); }
  if(VIRTUAL.container){
    if(!VIRTUAL.container.contains(VIRTUAL.topSpacer)){ VIRTUAL.container.appendChild(VIRTUAL.topSpacer); }
    if(!VIRTUAL.container.contains(VIRTUAL.bottomSpacer)){ VIRTUAL.container.appendChild(VIRTUAL.bottomSpacer); }
  }
  if(VIRTUAL.viewport && !VIRTUAL.boundScroll){
    VIRTUAL.viewport.addEventListener('scroll', ()=>scheduleVirtualUpdate(false));
    if(virtualViewportObserver){ virtualViewportObserver.observe(VIRTUAL.viewport); }
    VIRTUAL.boundScroll=true;
  }
}
function clearVirtual(){
  if(VIRTUAL.container){
    for(const node of VIRTUAL.nodes.values()){
      if(node.parentNode===VIRTUAL.container) VIRTUAL.container.removeChild(node);
    }
    if(VIRTUAL.topSpacer && VIRTUAL.topSpacer.parentNode===VIRTUAL.container) VIRTUAL.container.removeChild(VIRTUAL.topSpacer);
    if(VIRTUAL.bottomSpacer && VIRTUAL.bottomSpacer.parentNode===VIRTUAL.container) VIRTUAL.container.removeChild(VIRTUAL.bottomSpacer);
  }
  VIRTUAL.nodes.clear();
  VIRTUAL.heightCache=[];
  VIRTUAL.renderStart=0;
  VIRTUAL.renderEnd=0;
  VIRTUAL.active=false;
}
function updateAverageHeight(index,height){
  if(!Number.isFinite(height) || height<=0) return;
  VIRTUAL.heightCache[index]=height;
  const known=VIRTUAL.heightCache.filter(v=>Number.isFinite(v)&&v>0);
  if(known.length){
    const sum=known.reduce((a,b)=>a+b,0);
    VIRTUAL.averageHeight=sum/known.length;
  }
}
function computePadding(start,end){
  let total=0;
  for(let i=start;i<end;i++){
    const h=VIRTUAL.heightCache[i];
    total+=Number.isFinite(h)&&h>0?h:(VIRTUAL.averageHeight||56);
  }
  return total;
}
function materializeRow(index){
  const item=VIRTUAL.items[index];
  if(!item) return null;
  let node=null;
  if(item.type==='group'){ node=trGroup(item.group,VIRTUAL.lastFilter); }
  else if(item.type==='row'){ node=trChild(item.row,VIRTUAL.lastFilter); }
  if(node){ node.dataset.virtualIndex=String(index); }
  return node;
}
function updateVirtualRange(force){
  if(!VIRTUAL.items.length || !VIRTUAL.active){ return; }
  ensureVirtualSetup();
  const viewport=VIRTUAL.viewport;
  const container=VIRTUAL.container;
  if(!viewport || !container) return;

  let scrollTop=viewport.scrollTop;
  const viewportHeight=viewport.clientHeight||1;
  const avg=Math.max(24, VIRTUAL.averageHeight||56);
  const nearTop = scrollTop <= avg;
  if(nearTop && scrollTop!==0){
    viewport.scrollTop = 0;
    scrollTop = 0;
  }
  let start=Math.max(0, Math.floor(scrollTop/avg)-10);
  if(nearTop){
    start = 0;
  }
  let end=Math.min(VIRTUAL.items.length, Math.ceil((scrollTop+viewportHeight)/avg)+10);
  if(!force && start===VIRTUAL.renderStart && end===VIRTUAL.renderEnd) return;

  for(const [idx,node] of [...VIRTUAL.nodes.entries()]){
    if(idx<start || idx>=end){
      if(node.parentNode===container) container.removeChild(node);
      VIRTUAL.nodes.delete(idx);
    }
  }

  const frag=document.createDocumentFragment();
  for(let i=start;i<end;i++){
    if(!VIRTUAL.nodes.has(i)){
      const node=materializeRow(i);
      if(node){
        VIRTUAL.nodes.set(i,node);
        frag.appendChild(node);
        requestAnimationFrame(()=>{
          const rect=node.getBoundingClientRect();
          if(rect && rect.height){ updateAverageHeight(i, rect.height); scheduleVirtualUpdate(false); }
        });
      }
    }
  }
  if(frag.childNodes.length){ container.insertBefore(frag, VIRTUAL.bottomSpacer); }

  setSpacerHeight(VIRTUAL.topSpacer, computePadding(0,start));
  setSpacerHeight(VIRTUAL.bottomSpacer, computePadding(end,VIRTUAL.items.length));

  VIRTUAL.renderStart=start;
  VIRTUAL.renderEnd=end;
}
function scheduleVirtualUpdate(force){
  if(force){ updateVirtualRange(true); return; }
  if(VIRTUAL.pending) return;
  VIRTUAL.pending=true;
  requestAnimationFrame(()=>{ VIRTUAL.pending=false; updateVirtualRange(false); });
}
function fillGroupFilter(){
  const sel=$('#groupFilter'); const val=sel.value;
  sel.innerHTML='<option value="">Obergruppe (alle)</option>'+GROUPS.map(g=>`<option value="${g.groupId}">${escapeHtml(g.title||'')}</option>`).join('');
  if([...sel.options].some(o=>o.value===val)) sel.value=val;
}
function render(){
  const body=$('#rows');
  const f=currentFilters();
  VIRTUAL.lastFilter=f;

  const groups=GROUPS.filter(g=>!f.group||g.groupId===f.group);
  const items=[];
  let groupCount=0;
  let posCount=0;

  for(const g of groups){
    const groupProxy={ norm:{
      haystack: normalizeText(`${g.groupId||''} ${g.title||''}`),
      id: normalizeText(g.groupId||''),
      kurz: g.normTitle||'',
      beschreibung: g.normTitle||''
    }};
    const groupMatch=matchesFilters(groupProxy,f);
    const filteredChildren=g.children.filter(child=>matchesFilters(child,f));
    if(!groupMatch && !filteredChildren.length) continue;

    items.push({type:'group', group:g});
    groupCount++;
    const useChildren = groupMatch ? g.children : filteredChildren;
    for(const child of useChildren){
      ensureRowState(child);
      items.push({type:'row', row:child});
      posCount++;
    }
  }

  $('#count').textContent=`${groupCount} Gruppen · ${posCount} Positionen sichtbar`;

  if(!items.length){
    clearVirtual();
    if(body){
      body.innerHTML='<tr class="empty"><td colspan="9">Keine Positionen gefunden.</td></tr>';
    }
    VIRTUAL.items=[];
    return;
  }

  const shouldVirtualize = items.length > VIRTUAL_THRESHOLD;

  if(!shouldVirtualize){
    clearVirtual();
    if(body){
      body.innerHTML='';
      const frag=document.createDocumentFragment();
      for(const item of items){
        const node = item.type==='group' ? trGroup(item.group,f) : trChild(item.row,f);
        if(node){ frag.appendChild(node); }
      }
      body.appendChild(frag);
    }
    VIRTUAL.items=[];
    return;
  }

  if(body){ body.innerHTML=''; }
  VIRTUAL.container=body;
  ensureVirtualSetup();
  if(body){
    if(!body.contains(VIRTUAL.topSpacer)) body.appendChild(VIRTUAL.topSpacer);
    if(!body.contains(VIRTUAL.bottomSpacer)) body.appendChild(VIRTUAL.bottomSpacer);
  }

  for(const node of VIRTUAL.nodes.values()){
    if(body && node.parentNode===body){ body.removeChild(node); }
  }
  VIRTUAL.nodes.clear();

  VIRTUAL.items=items;
  VIRTUAL.active=true;
  VIRTUAL.heightCache=new Array(items.length).fill(null);
  VIRTUAL.averageHeight=56;
  VIRTUAL.renderStart=0;
  VIRTUAL.renderEnd=0;
  setSpacerHeight(VIRTUAL.topSpacer,0);
  setSpacerHeight(VIRTUAL.bottomSpacer,0);
  scheduleVirtualUpdate(true);
}

/* >>> Immer zum Listenanfang springen */
function jumpToTop(){
  const wrap=$('#tableWrap');
  if(wrap){ wrap.scrollTop = 0; }
  scheduleVirtualUpdate(true);
}

function getCurrentPriceListLabel(){
  const sheet = (CURRENT_SHEET || '').trim();
  const currentFileEl = $('#currentFile');
  const fileText = currentFileEl && currentFileEl.textContent ? currentFileEl.textContent.trim() : '';
  return sheet || fileText || '–';
}

/* ===== Dynamische Höhe ===== */
function recomputeChromeOffset(){
  const controls = document.getElementById('controls');
  const filters = document.getElementById('filters');
  const thead = document.getElementById('thead');
  const h = (controls?.offsetHeight||0) + (filters?.offsetHeight||0) + (thead?.offsetHeight||0);
  document.documentElement.style.setProperty('--chrome-offset', h + 'px');
}
const ro = new ResizeObserver(()=>recomputeChromeOffset());
['#controls','#filters','#thead'].forEach(sel=>{
  const el = document.querySelector(sel);
  if(el) ro.observe(el);
});
window.addEventListener('resize', recomputeChromeOffset);

/* ================= Workbook =========================== */
async function loadWorkbook(fileOrBuffer){
  const buf=(fileOrBuffer instanceof ArrayBuffer)?fileOrBuffer:await fileOrBuffer.arrayBuffer();
  const wb=XLSX.read(buf,{type:'array'}); LAST_WB=wb;
  $('#sheetSel').innerHTML=wb.SheetNames.map((n,i)=>`<option value="${n}" ${i===0?'selected':''}>${n}</option>`).join('');
  return wb;
}
async function loadFromSelectedSheet(wb){
  wb=wb||LAST_WB; if(!wb) return;
  const name=$('#sheetSel').value||wb.SheetNames[0];
  CURRENT_SHEET = name;
  setStatus('info',`Gelesen aus Blatt: <b>${escapeHtml(name)}</b>.`,2500);
  const ws=wb.Sheets[name];
  const rows=extractRows(ws);
  GROUPS=buildGroups(rows);
  ROW_STATES.clear();
  fillGroupFilter();
  render();
  clearBasket();
  renderSummary(false);
  requestAnimationFrame(()=>{ jumpToTop(); recomputeChromeOffset(); });
}

/* ================= Events ============================= */
$('#manualBtn').addEventListener('click', () => $('#file').click());

$('#file').addEventListener('change', async e=>{
  const f=e.target.files&&e.target.files[0]; if(!f) return;
  if(!/\.(xlsx|xlsm)$/i.test(f.name)){ setStatus('warn','Nur <b>.xlsx</b> oder <b>.xlsm</b> erlaubt.',3500); e.target.value=''; return; }
  setStatus('info','Lade Datei…',2500);
  $('#currentFile').textContent=f.name;
  const wb=await loadWorkbook(f); await loadFromSelectedSheet(wb);
  setDrawer(false); setStatus('ok','Bereit.',1500);
  $('#manualBtn').style.display='none';
});

['#search','#fA','#fB','#fC'].forEach(sel=>{
  document.querySelector(sel).addEventListener('input',debounced(()=>{ render(); },200));
});

document.querySelector('#groupFilter').addEventListener('change', ()=>{
  const wrap = document.getElementById('tableWrap');
  const duration = 180;
  const timing = {duration, easing:'ease', fill:'forwards'};
  const run = async ()=>{
    if(wrap && typeof wrap.animate === 'function'){
      try{ await wrap.animate([{opacity:1},{opacity:0.08}], timing).finished; }
      catch{}
    }else if(wrap){
      wrap.style.transition='opacity 180ms ease';
      wrap.style.opacity='0';
      await new Promise(res=>setTimeout(res, duration));
    }

    render();
    jumpToTop();
    recomputeChromeOffset();

    if(wrap && typeof wrap.animate === 'function'){
      try{ await wrap.animate([{opacity:0.08},{opacity:1}], timing).finished; }
      catch{}
      wrap.style.opacity='';
    }else if(wrap){
      wrap.style.opacity='';
      wrap.style.transition='';
    }
  };
  GROUP_SWITCH_ANIM = GROUP_SWITCH_ANIM.then(()=>run()).catch(()=>{});
});

const sheetSel=$('#sheetSel'); let lastSheetValue=null;
sheetSel.addEventListener('focus',()=>{ lastSheetValue=sheetSel.value; });
sheetSel.addEventListener('mousedown',()=>{ lastSheetValue=sheetSel.value; });
sheetSel.addEventListener('change', async ()=>{
  const hasFilters = ($('#search').value||$('#fA').value||$('#fB').value||$('#fC').value||$('#groupFilter').value);
  const hasSelection = getBasketSize()>0;
  if(hasFilters||hasSelection){
    const ok=confirm('Blatt wechseln? Alle Filter und markierten Positionen werden zurückgesetzt.');
    if(!ok){ sheetSel.value=lastSheetValue??sheetSel.value; return; }
  }
  await loadFromSelectedSheet(); setDrawer(false); setStatus('ok','Blatt gewechselt.',1500);
});

$('#reset').addEventListener('click', ()=>{
  if(!confirm('Alle Filter, Mengen und markierten Positionen werden zurückgesetzt. Fortfahren?')) return;
  $('#search').value=$('#fA').value=$('#fB').value=$('#fC').value=''; $('#groupFilter').value='';
  ROW_STATES.clear();
  render(); clearBasket(); renderSummary(false); setDrawer(false); setStatus('info','Zurückgesetzt.',1500);
  requestAnimationFrame(()=>{ jumpToTop(); recomputeChromeOffset(); });
});

function setDrawer(open){
  DRAWER_OPEN=!!open;
  $('#summary').classList.toggle('open',DRAWER_OPEN);
  document.body.classList.toggle('drawer-open',DRAWER_OPEN);
  requestAnimationFrame(recomputeChromeOffset);
  scheduleVirtualUpdate(true);
}
document.addEventListener('click',(e)=>{ if(e.target && (e.target.id==='toggleDrawer' || e.target.closest('#toggleDrawer'))) setDrawer(!DRAWER_OPEN); });
$('#drawerHead').addEventListener('dblclick',()=>setDrawer(!DRAWER_OPEN));
document.addEventListener('keydown',(e)=>{ if(e.altKey && (e.key==='o' || e.key==='O')){ e.preventDefault(); setDrawer(!DRAWER_OPEN); } });
document.addEventListener('keydown',(e)=>{
  if(!e.altKey || !e.shiftKey) return;
  const key = String(e.key||'').toLowerCase();
  if(key==='h'){ e.preventDefault(); triggerListPrint('all'); return; }
  if(key==='o'){ e.preventDefault(); triggerListPrint('current'); return; }
  if(key==='n'){
    const target = e.target || document.activeElement;
    if(target){
      const tag = target.tagName;
      const ignore = tag==='INPUT' || tag==='TEXTAREA' || tag==='SELECT' || target.isContentEditable;
      if(ignore){ return; }
    }
    e.preventDefault();
    addManualTopItem();
  }
});

/* Zusammenfassung + Feedback */
let lastSum = 0;
function pulseHead(){ const h=$('#drawerHead'); h.classList.add('pulse'); setTimeout(()=>h.classList.remove('pulse'), 800); }
function updateDelta(sum){
  const el=$('#selSum'); el.classList.remove('sum-up','sum-down');
  if(sum>lastSum){ el.classList.add('sum-up'); setTimeout(()=>el.classList.remove('sum-up'),1000); }
  else if(sum<lastSum){ el.classList.add('sum-down'); setTimeout(()=>el.classList.remove('sum-down'),1000); }
  lastSum = sum;
}

function getLineTotalValue(item){
  if(!item) return 0;
  if(Number.isFinite(item.totalNum)) return item.totalNum;
  if(Number.isFinite(item.preisNum) && Number.isFinite(item.qtyNum)){
    return item.preisNum * item.qtyNum;
  }
  return 0;
}

function computeBasketSums(){
  let sumMain = 0;
  let sumAlt = 0;
  for(const item of BASKET_STATE.items){
    const total = getLineTotalValue(item);
    if(!Number.isFinite(total)){ continue; }
    if(item?.isAlternative){ sumAlt += total; }
    else{ sumMain += total; }
  }
  return { main: sumMain, alt: sumAlt, total: sumMain + sumAlt };
}

function computeBasketSum(){
  const sums = computeBasketSums();
  return sums.total;
}

function formatQtyInputValue(qty){
  return Number.isFinite(qty) ? String(qty) : '';
}

function formatPriceInputValue(price){
  return Number.isFinite(price) ? String(price) : '';
}

function formatLineTotalDisplay(totalNum, isAlternative){
  let text = '–';
  let dataset = '0';
  let isNegative = false;
  if(Number.isFinite(totalNum)){
    text = fmtPrice(totalNum);
    dataset = String(totalNum);
    isNegative = totalNum < 0;
    if(totalNum === 0){
      text = fmtPrice(0);
    }
  }
  if(isAlternative){
    text = `(${text})`;
  }
  return { text, dataset, isNegative, isAlternative: !!isAlternative };
}

function updateLineTotalCell(cell, totalNum, isAlternative){
  if(!cell) return;
  const { text, dataset, isNegative, isAlternative: alt } = formatLineTotalDisplay(totalNum, isAlternative);
  cell.textContent = text;
  cell.dataset.total = dataset;
  cell.classList.toggle('neg', isNegative);
  cell.classList.toggle('alt-total', alt);
}

function updateSummaryDisplay(sums, count, commitDelta){
  const total = sums?.total ?? 0;
  const sumMain = sums?.main ?? 0;
  const sumAlt = sums?.alt ?? 0;
  $('#selCount').textContent = String(count);
  const selSumEl = $('#selSum');
  if(selSumEl){
    selSumEl.textContent = total.toLocaleString('de-AT',{style:'currency',currency:'EUR'});
  }
  const sumMainEl = document.getElementById('sum-main');
  if(sumMainEl){
    sumMainEl.textContent = fmtPrice(sumMain);
    sumMainEl.dataset.sum = String(sumMain);
    sumMainEl.classList.toggle('neg', sumMain<0);
  }
  const sumAltEl = document.getElementById('sum-alt');
  if(sumAltEl){
    const altText = fmtPrice(sumAlt);
    sumAltEl.textContent = `(${altText})`;
    sumAltEl.dataset.sum = String(sumAlt);
    sumAltEl.classList.toggle('neg', sumAlt<0);
  }
  if(commitDelta){
    updateDelta(total);
  }
}

function focusSummaryEditor(lineId, field){
  if(!lineId) return;
  const wrap = $('#summaryTableWrap');
  if(!wrap) return;
  const safeId = (typeof CSS!=='undefined' && CSS && typeof CSS.escape==='function') ? CSS.escape(lineId) : lineId.replace(/"/g,'\\"');
  const selectors = [];
  if(field){
    selectors.push(`[data-line-id="${safeId}"][data-field="${field}"]`);
  }
  selectors.push(`[data-line-id="${safeId}"] input`, `[data-line-id="${safeId}"] textarea`);
  let target=null;
  for(const sel of selectors){
    const el = wrap.querySelector(sel);
    if(el){ target=el; break; }
  }
  if(target){
    target.focus();
    if(typeof target.select==='function'){ try{ target.select(); }catch{} }
  }
}

function renderSummary(feedback, options={}){
  const wrap=$('#summaryTableWrap');
  if(!wrap) return;
  const items=getDisplayOrderedItems();

  let sumMain=0;
  let sumAlt=0;
  const rows=items.map(it=>{
    const totalNum = getLineTotalValue(it);
    const isManual = !!it.isManual;
    const isAlternative = !!it.isAlternative;
    if(Number.isFinite(totalNum)){
      if(isAlternative){ sumAlt += totalNum; }
      else{ sumMain += totalNum; }
    }
    const kurzHTML = linkify(escapeHtml(it.kurz||''));
    const beschrHTML = linkify(escapeHtml(it.beschreibung||''));
    const qtyVal = formatQtyInputValue(it.qtyNum);
    const totalInfo = formatLineTotalDisplay(totalNum, isAlternative);
    const totalClasses = ['right'];
    if(totalInfo.isNegative){ totalClasses.push('neg'); }
    if(isAlternative){ totalClasses.push('alt-total'); }
    const manualEditor = isManual
      ? `<div class="manual-editor">
          <input type="text" class="manual-title" data-line-id="${escapeHtml(it.lineId)}" data-field="title" value="${escapeHtml(it.kurz||'')}" placeholder="${escapeHtml(MANUAL_DEFAULT_TITLE)}" />
          <textarea class="manual-desc" data-line-id="${escapeHtml(it.lineId)}" data-field="description" placeholder="${escapeHtml(MANUAL_DESCRIPTION_PLACEHOLDER)}">${escapeHtml(it.beschreibung||'')}</textarea>
        </div>`
      : `<div><b>${kurzHTML}</b></div><div class="desc">${beschrHTML}</div>`;
    const priceContent = isManual
      ? `<div class="price-editor">
          <input type="number" step="0.01" inputmode="decimal" class="price-input" data-line-id="${escapeHtml(it.lineId)}" data-field="price" placeholder="0,00" value="${escapeHtml(formatPriceInputValue(it.preisNum))}" />
        </div>`
      : fmtPrice(it.preisNum);
    let unitValue = it.eh || '';
    if(isManual){
      unitValue = unitValue || MANUAL_DEFAULT_UNIT;
      if(it.eh !== unitValue){
        it.eh = unitValue;
      }
    }
    const unitEditor = isManual
      ? `<span class="unit-label" data-role="unit-label">${escapeHtml(unitValue)}</span>`
      : `<span class="qty-unit" data-role="qty-unit">${escapeHtml(unitValue)}</span>`;
    const rowClasses = ['basket-row'];
    if(isManual){ rowClasses.push('manual-row'); }
    if(isAlternative){ rowClasses.push('alt-item'); }
    const altToggle = `<label class="alt-flag alt-flag-basket" title="Als Alternative markieren">
        <input type="checkbox" class="alt-toggle-basket" data-line-id="${escapeHtml(it.lineId)}" ${isAlternative?'checked':''} aria-label="Alternativposition" />
      </label>`;
    return `<tr class="${rowClasses.join(' ')}" data-line-id="${escapeHtml(it.lineId)}" data-manual="${isManual?'true':'false'}" data-alt="${isAlternative?'true':'false'}" data-price="${Number.isFinite(it.preisNum)?String(it.preisNum):'0'}">
      <td class="artnr-cell" data-role="artnr">${escapeHtml(displayArtnr(it))}</td>
      <td class="desc-cell">${manualEditor}</td>
      <td class="right price-cell" data-role="price-cell">${priceContent}</td>
      <td class="right qty-cell">
        <input type="number" step="0.01" inputmode="decimal" class="qty-input" data-line-id="${escapeHtml(it.lineId)}" data-field="qty" value="${escapeHtml(qtyVal)}" />
      </td>
      <td class="unit-cell">${unitEditor}</td>
      <td class="center alt-cell">${altToggle}</td>
      <td class="${totalClasses.join(' ')}" data-role="line-total" data-line-id="${escapeHtml(it.lineId)}" data-total="${Number.isFinite(totalNum)?String(totalNum):'0'}">${totalInfo.text}</td>
      <td class="action" style="width:48px">
        <button type="button" class="remove-line" data-line-id="${escapeHtml(it.lineId)}" title="Position entfernen" aria-label="Position entfernen">✖</button>
      </td>
    </tr>`;
  }).join('');

  const bodyHTML = rows || '<tr class="empty"><td colspan="8" class="muted">Keine Positionen markiert.</td></tr>';
  const sums = { main: sumMain, alt: sumAlt, total: sumMain + sumAlt };
  wrap.innerHTML = `
    <table>
      <thead><tr><th>Art.Nr.</th><th>Bezeichnung (Kurztext + Beschreibung)</th><th class="right">EH-Preis</th><th class="right">Menge</th><th>EH</th><th class="center">Alt.</th><th class="right">Gesamtpreis</th><th class="center">&nbsp;</th></tr></thead>
      <tbody>${bodyHTML}</tbody>
      <tfoot>
        <tr class="tot total-main"><td colspan="6" class="right">Gesamtsumme</td><td class="right total-cell" id="sum-main" data-sum="${String(sumMain)}">${fmtPrice(sumMain)}</td><td></td></tr>
        <tr class="tot total-alt"><td colspan="6" class="right"><em>Gesamtsumme Alternativpositionen</em></td><td class="right total-cell alt-total-cell" id="sum-alt" data-sum="${String(sumAlt)}">(${fmtPrice(sumAlt)})</td><td></td></tr>
      </tfoot>
    </table>`;

  updateSummaryDisplay(sums, items.length, !!feedback);

  attachSummaryInteractions(wrap);

  if(options?.focusLineId){
    focusSummaryEditor(options.focusLineId, options.focusField);
  }

  if(feedback){ pulseHead(); setStatus('ok','Bereit.',1500); }
}

function attachSummaryInteractions(wrap){
  const qtyInputs = wrap.querySelectorAll('input.qty-input');
  const priceInputs = wrap.querySelectorAll('input.price-input');
  const titleInputs = wrap.querySelectorAll('input.manual-title');
  const descInputs = wrap.querySelectorAll('textarea.manual-desc');

  const recomputeAndDisplay = (commit=false)=>{
    const sums = computeBasketSums();
    updateSummaryDisplay(sums, BASKET_STATE.items.length, commit);
  };

  const setRowInvalid = (row, key, invalid)=>{
    if(!row) return;
    if(invalid){ row.dataset[key] = '1'; }
    else{ delete row.dataset[key]; }
    if(row.dataset.qtyInvalid || row.dataset.priceInvalid){
      row.classList.add('invalid');
    }else{
      row.classList.remove('invalid');
    }
  };

  const setInputInvalid = (input, invalid)=>{
    if(!input) return;
    if(invalid){ input.classList.add('invalid'); }
    else{ input.classList.remove('invalid'); }
  };

  qtyInputs.forEach(inp=>{
    const lineId = inp.dataset.lineId;
    const row = inp.closest('tr');
    const totalCell = row?.querySelector('[data-role="line-total"]');

    inp.addEventListener('input',()=>{
      if(!lineId) return;
      const existing = findBasketItem(lineId);
      if(!existing){ return; }
      const baseTotal = getLineTotalValue(existing);
      const preview = setLineQty(lineId, inp.value, {commit:false});
      if(preview.status === 'missing'){
        renderSummary(false);
        return;
      }
      if(preview.status === 'invalid'){
        setRowInvalid(row,'qtyInvalid',true);
        setInputInvalid(inp,true);
        return;
      }

      setRowInvalid(row,'qtyInvalid',false);
      setInputInvalid(inp,false);

      if(preview.status === 'preview' && preview.item){
        updateLineTotalCell(totalCell, preview.item.totalNum, preview.item.isAlternative);
        const previewSums = computeBasketSums();
        const previewTotal = getLineTotalValue(preview.item);
        const delta = previewTotal - baseTotal;
        if(existing?.isAlternative){ previewSums.alt += delta; }
        else{ previewSums.main += delta; }
        previewSums.total = previewSums.main + previewSums.alt;
        updateSummaryDisplay(previewSums, BASKET_STATE.items.length, false);
      }
      else if(preview.status === 'empty'){
        updateLineTotalCell(totalCell, NaN, existing?.isAlternative);
        const previewSums = computeBasketSums();
        if(existing?.isAlternative){ previewSums.alt -= baseTotal; }
        else{ previewSums.main -= baseTotal; }
        previewSums.total = previewSums.main + previewSums.alt;
        updateSummaryDisplay(previewSums, BASKET_STATE.items.length, false);
      }
      else{
        updateLineTotalCell(totalCell, baseTotal, existing?.isAlternative);
        recomputeAndDisplay(false);
      }
    });

    inp.addEventListener('change',()=>{
      if(!lineId) return;
      const before = findBasketItem(lineId);
      const result = setLineQty(lineId, inp.value);

      if(result.status === 'removed'){
        recomputeAndDisplay(true);
        renderSummary(false);
        triggerUpdatedBadge();
        return;
      }

      if(result.status === 'missing'){
        renderSummary(false);
        return;
      }

      if(result.status === 'invalid'){
        setRowInvalid(row,'qtyInvalid',true);
        setInputInvalid(inp,true);
        if(before){
          inp.value = formatQtyInputValue(before.qtyNum);
          updateLineTotalCell(totalCell, getLineTotalValue(before), before.isAlternative);
        }
        setStatus('warn','Ungültige Menge.',2500);
        recomputeAndDisplay(false);
        return;
      }

      setRowInvalid(row,'qtyInvalid',false);
      setInputInvalid(inp,false);

      if(result.status === 'updated' && result.item){
        inp.value = formatQtyInputValue(result.item.qtyNum);
        updateLineTotalCell(totalCell, result.item.totalNum, result.item.isAlternative);
        recomputeAndDisplay(true);
        triggerUpdatedBadge();
      }
      else if(result.status === 'empty' && before){
        inp.value = formatQtyInputValue(before.qtyNum);
        updateLineTotalCell(totalCell, getLineTotalValue(before), before.isAlternative);
        recomputeAndDisplay(false);
      }
      else{
        const fresh = findBasketItem(lineId);
        updateLineTotalCell(totalCell, getLineTotalValue(fresh), fresh?.isAlternative);
        recomputeAndDisplay(false);
      }
    });
  });

  priceInputs.forEach(inp=>{
    const lineId = inp.dataset.lineId;
    const row = inp.closest('tr');
    const totalCell = row?.querySelector('[data-role="line-total"]');

    inp.addEventListener('input',()=>{
      if(!lineId) return;
      const existing = findBasketItem(lineId);
      if(!existing){ renderSummary(false); return; }
      const baseTotal = getLineTotalValue(existing);
      const raw = inp.value;
      const trimmed = typeof raw === 'string' ? raw.trim() : '';
      let previewPrice = 0;
      if(trimmed){
        const parsed = parseEuro(raw);
        if(!Number.isFinite(parsed)){
          setRowInvalid(row,'priceInvalid',true);
          setInputInvalid(inp,true);
          return;
        }
        previewPrice = parsed;
      }
      setRowInvalid(row,'priceInvalid',false);
      setInputInvalid(inp,false);
      const qty = Number.isFinite(existing.qtyNum) ? existing.qtyNum : 0;
      const previewTotal = qty * previewPrice;
      updateLineTotalCell(totalCell, previewTotal, existing?.isAlternative);
      const previewSums = computeBasketSums();
      const delta = previewTotal - baseTotal;
      if(existing?.isAlternative){ previewSums.alt += delta; }
      else{ previewSums.main += delta; }
      previewSums.total = previewSums.main + previewSums.alt;
      updateSummaryDisplay(previewSums, BASKET_STATE.items.length, false);
    });

    inp.addEventListener('change',()=>{
      if(!lineId) return;
      const existing = findBasketItem(lineId);
      if(!existing){ renderSummary(false); return; }
      const raw = inp.value;
      const trimmed = typeof raw === 'string' ? raw.trim() : '';
      let newPrice = 0;
      if(trimmed){
        const parsed = parseEuro(raw);
        if(!Number.isFinite(parsed)){
          setRowInvalid(row,'priceInvalid',true);
          setInputInvalid(inp,true);
          inp.value = formatPriceInputValue(existing.preisNum);
          updateLineTotalCell(totalCell, getLineTotalValue(existing), existing.isAlternative);
          setStatus('warn','Ungültiger Preis.',2500);
          recomputeAndDisplay(false);
          return;
        }
        newPrice = parsed;
      }
      existing.preisNum = newPrice;
      const qty = Number.isFinite(existing.qtyNum) ? existing.qtyNum : 0;
      existing.totalNum = qty * newPrice;
      inp.value = formatPriceInputValue(existing.preisNum);
      updateLineTotalCell(totalCell, existing.totalNum, existing.isAlternative);
      setRowInvalid(row,'priceInvalid',false);
      setInputInvalid(inp,false);
      recomputeAndDisplay(true);
      triggerUpdatedBadge();
    });
  });

  titleInputs.forEach(inp=>{
    const lineId = inp.dataset.lineId;
    if(!lineId) return;
    inp.addEventListener('input',()=>{
      const existing = findBasketItem(lineId);
      if(existing){ existing.kurz = inp.value; }
    });
  });

  descInputs.forEach(inp=>{
    const lineId = inp.dataset.lineId;
    if(!lineId) return;
    requestAnimationFrame(()=>autoGrow(inp));
    inp.addEventListener('input',()=>{
      autoGrow(inp);
      const existing = findBasketItem(lineId);
      if(existing){ existing.beschreibung = inp.value; }
    });
  });


  const altToggles = wrap.querySelectorAll('input.alt-toggle-basket');
  altToggles.forEach(toggle=>{
    toggle.addEventListener('change',()=>{
      const lineId = toggle.dataset.lineId;
      if(!lineId) return;
      setAltFlag(lineId, toggle.checked);
    });
  });

  wrap.querySelectorAll('button.remove-line').forEach(btn=>{
    btn.addEventListener('click',()=>{
      const lineId = btn.dataset.lineId;
      if(!lineId) return;
      if(removeLine(lineId)){
        recomputeAndDisplay(true);
        renderSummary(false);
        triggerUpdatedBadge();
      }
    });
  });
}

function applyVersionInfo(){
  const badge=document.getElementById('versionBadge');
  if(badge) badge.textContent=APP_VERSION;
  const chip=document.getElementById('versionChip');
  if(chip) chip.title=`Build: ${APP_BUILD_DATE} · Quelle: ${APP_BUILD_SOURCE}`;
}

function printableText(text){
  if(!text) return '';
  return linkify(escapeHtml(text)).replace(/\n/g,'<br>');
}

function collectPrintableGroups(mode){
  if(mode==='all'){
    return GROUPS.map(group=>({group, rows:group.children.slice()}));
  }
  const f=currentFilters();
  const groups=GROUPS.filter(g=>!f.group||g.groupId===f.group);
  const result=[];
  for(const g of groups){
    const groupProxy={ norm:{
      haystack: normalizeText(`${g.groupId||''} ${g.title||''}`),
      id: normalizeText(g.groupId||''),
      kurz: g.normTitle||'',
      beschreibung: g.normTitle||''
    }};
    const groupMatch=matchesFilters(groupProxy,f);
    const filtered=g.children.filter(child=>matchesFilters(child,f));
    if(!groupMatch && !filtered.length) continue;
    result.push({group:g, rows: groupMatch? g.children : filtered});
  }
  return result;
}

function buildListPrintMarkup(mode){
  const groups=collectPrintableGroups(mode);
  if(!groups.length) return '';
  const parts=['<table class="print-table">',
    '<thead><tr><th>Art.Nr.</th><th>Kurztext</th><th>Beschreibung</th><th>EH</th><th>EH-Info</th><th class="right">EH-Preis</th><th class="right">Menge</th><th class="right">Gesamtpreis</th><th>Hinweis</th></tr></thead>',
    '<tbody>'];
  for(const {group,rows} of groups){
    parts.push(`<tr class="group-row"><td colspan="9"><strong>${escapeHtml(group.groupId)} – ${escapeHtml(group.title||'')}</strong></td></tr>`);
    for(const row of rows){
      const state=ensureRowState(row);
      const editable=isSonderEditable(row.id);
      const preisSource=editable?state.preis:row.preis;
      const preisText=fmtPrice(preisSource)||'–';
      const qtyNum=parseQty(state.qty);
      const qtyText=(!Number.isNaN(qtyNum) && qtyNum!==0)?fmtQty(qtyNum):'–';
      const preisNum=parseEuro(preisSource);
      const totalVal=(!Number.isNaN(qtyNum) && qtyNum!==0 && Number.isFinite(preisNum))?preisNum*qtyNum:NaN;
      const totalText=Number.isFinite(totalVal)?fmtPrice(totalVal):'–';
      const totalClass=Number.isFinite(totalVal)&&totalVal<0?' class="neg"':'';
      const kurzText=printableText(editable?state.kurz:row.kurz_raw);
      const beschrText=printableText(editable?state.beschreibung:row.beschreibung_raw);
      const ehText=escapeHtml((editable?state.einheit:row.einheit)||'');
      const ehInfoText=escapeHtml((editable?state.einheitInfo:row.einheitInfo)||'');
      const hinweisText=printableText(row.hinweis_raw||'');
      parts.push(`<tr>
        <td>${escapeHtml(row.id||'')}</td>
        <td>${kurzText}</td>
        <td>${beschrText}</td>
        <td>${ehText}</td>
        <td>${ehInfoText}</td>
        <td class="right">${preisText}</td>
        <td class="right">${qtyText}</td>
        <td class="right"${totalClass}>${totalText}</td>
        <td class="desc">${hinweisText}</td>
      </tr>`);
    }
  }
  parts.push('</tbody></table>');
  return parts.join('');
}

function mountPrintTable(mode){
  const portal=document.getElementById('printPortal');
  if(!portal) return false;
  const markup=buildListPrintMarkup(mode);
  if(!markup){ portal.innerHTML=''; return false; }
  portal.innerHTML=markup;
  return true;
}

function clearPrintState(){
  const portal=document.getElementById('printPortal');
  if(portal) portal.innerHTML='';
  if(document.body.dataset.printMode) delete document.body.dataset.printMode;
  if(document.body.hasAttribute('data-print-mode')) document.body.removeAttribute('data-print-mode');
}

function triggerListPrint(mode){
  if(!mountPrintTable(mode)){
    setStatus('warn','Keine Daten zum Drucken.',3000);
    return;
  }
  document.body.dataset.printMode=mode;
  setStatus('info','Druckansicht (A4 quer) geöffnet…',2000);
  setTimeout(()=>window.print(),60);
}

window.addEventListener('afterprint', clearPrintState);

/* ============== Drucken (ein Fenster) ============== */
function buildPrintDoc(){
  const items=getDisplayOrderedItems();
  if(items.length===0) return null;

  const BVH=($('#bvhInput')?.value||'').trim()||'–';
  const AUF=($('#auftragInput')?.value||'').trim()||'–';
  const ERS=($('#erstellerInput')?.value||'').trim()||'–';
  const BETR=($('#betreffInput')?.value||'').trim()||'–';
  const DAT=new Date().toLocaleDateString('de-AT');
  const SHEET = CURRENT_SHEET || '–';

  let tbody=''; let sumMain=0; let sumAlt=0;
  items.forEach(it=>{
    const lineTotal=Number.isFinite(it.totalNum)?it.totalNum:0;
    if(it.isAlternative){ sumAlt+=lineTotal; } else { sumMain+=lineTotal; }
    const kurzHTML = linkify(escapeHtml(it.kurz||''));
    const beschrHTML = linkify(escapeHtml(it.beschreibung||''));
    const rowClasses=['item-row'];
    if(it.isAlternative){ rowClasses.push('alt-item'); }
    const totalInfo = formatLineTotalDisplay(lineTotal, !!it.isAlternative);
    const unitText = escapeHtml(it.eh||'');
    const qtyText = fmtQty(it.qtyNum);
    tbody+=`<tr class="${rowClasses.join(' ')}">
      <td>${escapeHtml(displayArtnr(it))}</td>
      <td><div><b>${kurzHTML}</b></div><div class="desc">${beschrHTML}</div></td>
      <td class="right">${fmtPrice(it.preisNum)}</td>
      <td class="right">${qtyText} ${unitText}</td>
      <td class="right${totalInfo.isNegative?' neg':''}">${totalInfo.text}</td>
    </tr>`;
  });

  return `<!DOCTYPE html><html lang="de"><head><meta charset="utf-8"><title>Kostenvoranschlag</title>
  <style>
    @page {
      size: A4 portrait;
      margin: 10mm 14mm 12mm 14mm;
      counter-reset: page 0;
      counter-increment: page;
      @bottom-center { content: element(summaryFooter); }
    }
    body{ font:12px/1.35 -apple-system,system-ui,Segoe UI,Roboto,Helvetica,Arial; color:#111; margin:0; }
    main{ padding-bottom:0; }
    main::after{ content:""; display:block; height:12mm; }

    .p-head{ display:grid; grid-template-columns: 1fr auto; column-gap: 16px; align-items:flex-start; margin-bottom:10px; }
    .title{ font-size:18px; font-weight:800; margin:0 0 6px 0; }
    .meta{ font-size:12px; line-height:1.45; }
    .meta b{ display:inline-block; min-width:120px; }
    .rightcol{ text-align:right; }
    .logo{ height:60px; margin-bottom:4px; display:block; margin-left:auto; }
    .created{ font-size:12px; white-space:nowrap; }

    table{ width:100%; border-collapse:collapse; }
    thead th{ text-align:left; padding:6px 6px; background:#fff6e6; border-bottom:1px solid #d8dee6; -webkit-print-color-adjust:exact; print-color-adjust:exact; }
    td{ padding:6px 6px; vertical-align:top; }
    tr.item-row{ page-break-inside: avoid; }
    td, th{ page-break-inside: avoid; }
    .right{ text-align:right; }
    .neg{ color:#b91c1c; font-weight:600; }
    tr.alt-item td{ font-style:italic; }
    tr.alt-item td .desc{ font-style:normal; }
    tr.alt-item td .desc b{ font-style:italic; }
    .desc{ white-space:pre-wrap; }
    .desc a{ text-decoration:underline; word-break:break-all; }

    .grand-totals{ page-break-inside: avoid; margin-top:10px; border-top:2px solid #bbb; padding-top:8px; display:grid; gap:6px; }
    .grand-total{ display:grid; grid-template-columns:1fr 140px; gap:8px; align-items:center; font-weight:700; }
    .grand-total .label{ text-align:right; padding-right:8px; }
    .grand-total .value{ text-align:right; }
    .grand-total.alt{ font-style:italic; font-weight:600; opacity:.9; }
    .grand-total.alt .value{ font-style:italic; }

    .note{ margin-top:12px; font-size:11px; color:#374151; }

  </style></head><body>

    <main>
    <div class="p-head">
      <div>
        <h1 class="title">Kostenvoranschlag</h1>
        <div class="meta">
          <div><b>BVH:</b> ${escapeHtml(BVH)}</div>
          <div><b>Auftragsnummer:</b> ${escapeHtml(AUF)}</div>
          <div><b>Ersteller:</b> ${escapeHtml(ERS)}</div>
          <div style="margin-top:4px"><b>Betreff:</b> ${escapeHtml(BETR)}</div>
        </div>
      </div>
      <div class="rightcol">
        <img src="./assets/img/Logo_Haas.jpg" class="logo" alt="Logo" />
        <div class="created"><b>Erstellt am:</b> ${escapeHtml(DAT)}</div>
      </div>
    </div>

    <table>
      <thead><tr>
        <th style="width:30px">Art.Nr.</th>
        <th>Kurztext und Beschreibung</th>
        <th class="right" style="width:50px">EH-Preis</th>
        <th class="right" style="width:80px">Menge</th>
        <th class="right" style="width:60px">Gesamt</th>
      </tr></thead>
      <tbody>${tbody}</tbody>
    </table>

    <div class="grand-totals">
      <div class="grand-total"><div class="label">Gesamtsumme</div><div class="value${sumMain<0?' neg':''}">${sumMain.toLocaleString('de-AT',{style:'currency',currency:'EUR'})}</div></div>
      <div class="grand-total alt"><div class="label"><em>Gesamtsumme Alternativpositionen</em></div><div class="value${sumAlt<0?' neg':''}">(${fmtPrice(sumAlt)})</div></div>
    </div>

    <div class="note">
      Der Preis versteht sich inkl. 20% MwSt.
      Änderungen, Irrtümer und Preisänderungen vorbehalten.
      <br>Hinweis: Grundlage der Preise laut <b>${escapeHtml(SHEET)}</b>.
    </div>
    </main>

  </body></html>`;
}

document.getElementById('printSummary').addEventListener('click', async ()=>{
  if(getBasketSize()===0){ setStatus('warn','Keine markierten Positionen zum Drucken.',3000); return; }
  const html = buildPrintDoc(); if(!html){ setStatus('warn','Nichts zu drucken.',3000); return; }

  const iframe = document.createElement('iframe');
  Object.assign(iframe.style,{position:'fixed',right:'0',bottom:'0',width:'0',height:'0',border:'0'});
  iframe.setAttribute('aria-hidden','true');
  document.body.appendChild(iframe);

  const win = iframe.contentWindow, doc = win.document;
  doc.open(); doc.write(html); doc.close();

  const done = () => { try{ iframe.remove(); }catch{} };
  const safePrint = () => { try{ win.focus(); win.onafterprint = done; win.print(); } catch { done(); } }

  if (doc.readyState === 'complete') { setTimeout(safePrint, 250); }
  else { win.addEventListener('load', () => setTimeout(safePrint, 250), {once:true}); }
});

/* ================= Init =============================== */
window.addEventListener('DOMContentLoaded', async ()=>{
  try{ localStorage.clear(); sessionStorage.clear(); }catch{}
  applyVersionInfo();
  ['#search','#fA','#fB','#fC','#groupFilter','#bvhInput','#auftragInput','#erstellerInput','#betreffInput'].forEach(sel=>{
    const el=document.querySelector(sel);
    if(el) el.value='';
  });
  setDrawer(false);
  recomputeChromeOffset();

  setStatus('info',`Automatisches Laden der <b>${escapeHtml(DEFAULT_FILE)}</b> (falls vorhanden).`,2500);
  try{
    const res=await fetch(DEFAULT_FILE_URL,{cache:'no-store'});
    if(res.ok){
      const buf=await res.arrayBuffer();
      $('#currentFile').textContent=DEFAULT_FILE;
      const wb=await loadWorkbook(buf);
      await loadFromSelectedSheet(wb);
      setDrawer(false); setStatus('ok','Bereit.',1500);
      requestAnimationFrame(()=>{ jumpToTop(); recomputeChromeOffset(); });
      $('#manualBtn').style.display='none';
    }else{
      setStatus('warn',`Standarddatei <b>${escapeHtml(DEFAULT_FILE)}</b> konnte nicht geladen werden.`,4000);
      $('#manualBtn').style.display='inline-flex';
    }
  }catch{
    setStatus('warn',`Standarddatei <b>${escapeHtml(DEFAULT_FILE)}</b> konnte nicht geladen werden.`,4000);
    $('#manualBtn').style.display='inline-flex';
  }
});
