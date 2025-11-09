/* ================== State ============================= */
let GROUPS = [], LAST_WB = null;
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
  return Number.isFinite(qty) ? qty.toLocaleString('de-AT',{minimumFractionDigits:0, maximumFractionDigits:2}) : '';
}

function formatPriceInputValue(price){
  return Number.isFinite(price) ? price.toLocaleString('de-AT',{minimumFractionDigits:2, maximumFractionDigits:2}) : '';
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
    sumMainEl.classList.toggle('neg', sumMain < 0);
  }
  const sumAltEl = document.getElementById('sum-alt');
  if(sumAltEl){
    const altText = fmtPrice(sumAlt);
    sumAltEl.textContent = `(${altText})`;
    sumAltEl.dataset.sum = String(sumAlt);
    sumAltEl.classList.toggle('neg', sumAlt < 0);
  }
  if(commitDelta){
    updateDelta(total);
  }
}
