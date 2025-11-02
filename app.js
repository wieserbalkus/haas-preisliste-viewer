/* ================= Basis-Einstellungen ================ */
const DEFAULT_FILE = 'Artikelpreisliste.xlsx';
const STORAGE_KEY  = 'plv22_state';
const QTY_MAX      = 999.99;

/* ================= Hilfsfunktionen ==================== */
function $(q){return document.querySelector(q)}
function escapeHtml(s){return String(s).replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;').replace(/\"/g,'&quot;').replace(/'/g,'&#39;')}
function isGroupId(id){ const s=String(id||"").replace(/\D/g,""); if(!s) return false; const n=parseInt(s,10); return Number.isFinite(n) && n%100===0; }
function parseEuro(str){ if(str==null) return NaN; let s=String(str).trim(); if(!s) return NaN; if(s.includes(',')) s=s.replace(/\./g,'').replace(',', '.'); return Number(s); }
function fmtPrice(v){ const n=parseEuro(v); return Number.isFinite(n)? n.toLocaleString('de-AT',{style:'currency',currency:'EUR'}) : (v??''); }
function debounced(fn,ms){ let t; return (...a)=>{ clearTimeout(t); t=setTimeout(()=>fn(...a),ms); } }
let statusTimer=null;
function setStatus(kind, html, ttl=3000){
  const el = $('#hint');
  if(!kind){ el.textContent=''; el.removeAttribute('data-status'); if(statusTimer){clearTimeout(statusTimer)}; return; }
  el.dataset.status = kind; el.innerHTML = html || '';
  if(statusTimer){ clearTimeout(statusTimer); }
  statusTimer = setTimeout(()=>{ el.textContent=''; el.removeAttribute('data-status'); }, ttl);
}
function isSonderEditable(id){ const s=String(id||'').replace(/\D/g,''); return /(?:98|99)$/.test(s); }

/* Menge parsing (±0–999,99) */
const QTY_RE=/^-?\d{1,3}([.]\d{1,2})?$/;
function parseQty(str){ if(str==null)return 0; let s=String(str).trim().replace(',', '.'); if(!s)return 0; if(!QTY_RE.test(s))return NaN; const n=Number(s); if(!Number.isFinite(n)||Math.abs(n)>QTY_MAX)return NaN; return n; }
function fmtQty(n){ return Number.isFinite(n)? n.toLocaleString('de-AT',{minimumFractionDigits:0, maximumFractionDigits:2}) : ''; }

/* ====== Hervorhebung (PDF-Style) + Linkify ====== */
function escRe(s){ return s.replace(/[.*+?^${}()|[\]\\]/g,'\\$&'); }
function splitTerms(s){ if(!s) return []; return String(s).trim().split(/\s+/).filter(Boolean); }
function makeRegex(terms){
  const t = [...new Set(terms.filter(Boolean))].sort((a,b)=>b.length-a.length);
  if(!t.length) return null;
  return new RegExp('(' + t.map(escRe).join('|') + ')','gi');
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

/* ========= Zustandsverwaltung (Filter/Meta) =========== */
function loadState(){ try{ return JSON.parse(localStorage.getItem(STORAGE_KEY)||'{}'); }catch{ return {}; } }
function saveState(patch){
  const st = loadState();
  const next = Object.assign({}, st, patch);
  localStorage.setItem(STORAGE_KEY, JSON.stringify(next));
}
function saveFilters(){
  const filters={ q:$('#search').value||'', A:$('#fA').value||'', B:$('#fB').value||'', C:$('#fC').value||'', group:$('#groupFilter').value||'' };
  saveState({filters});
}
function saveMeta(){
  const meta={ bvh:$('#bvhInput').value||'', auftrag:$('#auftragInput').value||'', ersteller:$('#erstellerInput').value||'', betreff:$('#betreffInput').value||'' };
  saveState({meta});
}

/* ================== State ============================= */
let GROUPS=[], LAST_WB=null, SORT={col:null,dir:1};
let DRAWER_OPEN = false;
let CURRENT_SHEET = '';
const SELECTED = new Map();
let GROUP_SWITCH_ANIM = Promise.resolve();

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

      out.push({
        id,
        kurz_raw:kurz, beschreibung_raw:beschr,
        styleB:{ bold:!!(B&&B.s&&B.s.font&&B.s.font.bold), underline:!!(B&&B.s&&B.s.font&&(B.s.font.underline||B.s.font.u)) },
        styleC:{ bold:!!(C2&&C2.s&&C2.s.font&&C2.s.font.bold), underline:!!(C2&&C2.s&&C2.s.font&&(C2.s.font.underline||C2.s.font.u)) },
        einheit:D&&D.v!=null?String(D.v).trim():'', einheitInfo:E&&E.v!=null?String(E.v).trim():'', preis:F?F.v:'', hinweis_raw:G&&G.v!=null?String(G.v).trim():'',
        linkB, linkC, linkG
      });
    }
    return out;
  }catch{
    const json=XLSX.utils.sheet_to_json(ws,{defval:""}); if(!json.length) return [];
    const k=Object.keys(json[0]);
    return json.map(r=>({ id:String(r[k[0]]||'').trim(), kurz_raw:String(r[k[1]]||'').trim(), beschreibung_raw:String(r[k[2]]||'').trim(),
      einheit:String(r[k[3]]||'').trim(), einheitInfo:String(r[k[4]]||'').trim(), preis:r[k[5]], hinweis_raw:String(r[k[6]]||'').trim(),
      styleB:null, styleC:null, linkB:'', linkC:'', linkG:'' }));
  }
}
function buildGroups(rows){
  const groups=[]; let cur=null;
  for(const r of rows){
    if(isGroupId(r.id)){ cur={groupId:r.id,title:(r.kurz_raw||r.beschreibung_raw),children:[]}; groups.push(cur); }
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
  const title = f && f.q ? hlAndLink(g.title||'', splitTerms(f.q)) : escapeHtml(g.title||'');
  tr.innerHTML=`<td colspan="9"><strong>${escapeHtml(g.groupId)} – ${title}</strong></td>`;
  return tr;
}

function trChild(c, f){
  const tr=document.createElement('tr');
  const editable=isSonderEditable(c.id);
  let preisNum=parseEuro(c.preis);

  const qTerms = f ? splitTerms(f.q) : [];
  const tID = [...qTerms, f?.A].filter(Boolean);
  const tK  = [...qTerms, f?.B].filter(Boolean);
  const tB  = [...qTerms, f?.C].filter(Boolean);
  const tG  = qTerms;

  const kurzStatic = c.linkB
    ? `<div class="desc"><a href="${escapeHtml(toSafeHref(c.linkB))}" target="_blank" rel="noopener noreferrer">${hi(c.kurz_raw, tK)}</a></div>`
    : `<div class="desc">${hlAndLink(c.kurz_raw, tK)}</div>`;

  const beschrStatic = c.linkC
    ? `<div class="desc"><a href="${escapeHtml(toSafeHref(c.linkC))}" target="_blank" rel="noopener noreferrer">${hi(c.beschreibung_raw, tB)}</a></div>`
    : `<div class="desc">${hlAndLink(c.beschreibung_raw, tB)}</div>`;

  const ehHTML = editable
    ? `<input class="cell-edit" data-field="einheit" value="${escapeHtml(c.einheit||'')}" />`
    : hlAndLink(c.einheit||'', tG);

  const ehInfoHTML = editable
    ? `<input class="cell-edit" data-field="einheitInfo" value="${escapeHtml(c.einheitInfo||'')}" />`
    : hlAndLink(c.einheitInfo||'', tG);

  const preisHTML = editable
    ? `<input class="cell-edit price" data-field="preis" inputmode="decimal" placeholder="0" value="${Number.isFinite(preisNum)? String(preisNum).replace('.',',') : ''}" />`
    : hlAndLink(fmtPrice(c.preis), tG);

  const hinweisHTML = c.linkG
    ? `<div class="desc"><a href="${escapeHtml(toSafeHref(c.linkG))}" target="_blank" rel="noopener noreferrer">${hi(c.hinweis_raw||'', tG)}</a></div>`
    : `<div class="desc">${hlAndLink(c.hinweis_raw||'', tG)}</div>`;

  const kurzHTML = editable
    ? `<textarea class="cell-edit ta" data-field="kurz">${escapeHtml(c.kurz_raw)}</textarea>`
    : kurzStatic;

  const beschrHTML = editable
    ? `<textarea class="cell-edit ta" data-field="beschreibung">${escapeHtml(c.beschreibung_raw)}</textarea>`
    : beschrStatic;

  const qtyId=`q_${c.id}`, addBtnId=`add_${c.id}`;
  const qtyHTML = `<div class="qty-wrap">
      <input id="${qtyId}" class="qty" type="text" inputmode="decimal" placeholder="0" maxlength="7" title="Menge (±0–999,99)" />
      <button id="${addBtnId}" type="button" class="addbtn" title="Zur Zusammenfassung hinzufügen">➕</button>
    </div>`;

  tr.innerHTML = `
    <td>${hlAndLink(c.id, tID)}</td>
    <td>${kurzHTML}</td>
    <td>${beschrHTML}</td>
    <td>${ehHTML}</td>
    <td>${ehInfoHTML}</td>
    <td class="right" data-sort="${preisNum}">${preisHTML}</td>
    <td>${qtyHTML}</td>
    <td class="right" data-total="0">–</td>
    <td class="desc">${hinweisHTML}</td>`;

  setTimeout(()=>{
    tr.querySelectorAll('textarea.cell-edit.ta').forEach(ta=>{
      autoGrow(ta); ta.addEventListener('input',()=>{ autoGrow(ta); if(SELECTED.has(c.id)) addOrUpdateSelectedFromRow(tr,c.id); });
    });
    const qtyInp=tr.querySelector('#'+CSS.escape(qtyId));
    const preisInp=tr.querySelector('input[data-field="preis"]');
    const totalCell=tr.querySelector('[data-total]');
    const addBtn=tr.querySelector('#'+CSS.escape(addBtnId));

    function currentPreis(){ if(preisInp){ const p=parseEuro(preisInp.value); return Number.isFinite(p)?p:0; } return Number.isFinite(preisNum)?preisNum:0; }
    function currentKurz(){ const k=tr.querySelector('[data-field="kurz"]'); return k?k.value.trim():tr.querySelector('td:nth-child(2)').innerText.trim(); }
    function currentBeschr(){ const b=tr.querySelector('[data-field="beschreibung"]'); return b?b.value.trim():tr.querySelector('td:nth-child(3)').innerText.trim(); }
    function currentEH(){ const e=tr.querySelector('[data-field="einheit"]'); return e?e.value.trim():(c.einheit||''); }

    function recalcRowTotal(){
      const q=parseQty(qtyInp?.value??''); const p=currentPreis();
      if(Number.isNaN(q)||q===0){
        totalCell.textContent='–'; totalCell.dataset.total='0'; totalCell.classList.toggle('neg',false);
        if(SELECTED.has(c.id)){ SELECTED.delete(c.id); addBtn.classList.remove('added'); addBtn.textContent='➕'; renderSummary(true); }
      }else{
        const total=p*q; totalCell.textContent=fmtPrice(total); totalCell.dataset.total=String(total); totalCell.classList.toggle('neg',total<0);
        if(SELECTED.has(c.id)){
          SELECTED.set(c.id,{id:c.id,kurz:currentKurz(),beschreibung:currentBeschr(),eh:currentEH(),preisNum:p,qtyNum:q,totalNum:total});
          renderSummary(true);
        }
      }
    }
    function addOrUpdateSelectedFromRow(trEl,id){
      const q=parseQty(qtyInp?.value??''); if(Number.isNaN(q)||q===0){ return false; }
      const p=currentPreis(); const total=p*q;
      SELECTED.set(id,{id,kurz:currentKurz(),beschreibung:currentBeschr(),eh:currentEH(),preisNum:p,qtyNum:q,totalNum:total});
      return true;
    }
    window.addOrUpdateSelectedFromRow=addOrUpdateSelectedFromRow;

    qtyInp?.addEventListener('input',()=>{ qtyInp.value=qtyInp.value.replace(/[^\d.,-]/g,'').replace(/(?!^)-/g,''); recalcRowTotal(); });
    preisInp?.addEventListener('input',()=>{ if(!/^[\d.,-]*$/.test(preisInp.value)){ preisInp.value=preisInp.value.replace(/[^\d.,-]/g,''); } preisInp.closest('td').dataset.sort=String(currentPreis()); recalcRowTotal(); });

    addBtn.addEventListener('click',()=>{
      if(SELECTED.has(c.id)){ SELECTED.delete(c.id); addBtn.classList.remove('added'); addBtn.textContent='➕'; renderSummary(true); return; }
      const ok=addOrUpdateSelectedFromRow(tr,c.id);
      if(ok){ addBtn.classList.add('added'); addBtn.textContent='✔︎'; renderSummary(true); setStatus('ok','Bereit.',1500); }
      else { setStatus('warn','Bitte zuerst eine gültige Menge (≠ 0) eingeben.',3500); addBtn.animate([{transform:'scale(1)'},{transform:'scale(1.08)'},{transform:'scale(1)'}],{duration:160}); }
    });

    if(SELECTED.has(c.id)){
      const sel=SELECTED.get(c.id);
      if(qtyInp) qtyInp.value=String(sel.qtyNum).replace('.',',');
      const preisInp2=tr.querySelector('input[data-field="preis"]'); if(preisInp2&&Number.isFinite(sel.preisNum)){ preisInp2.value=String(sel.preisNum).replace('.',','); }
      addBtn.classList.add('added'); addBtn.textContent='✔︎'; recalcRowTotal();
    }
    recalcRowTotal();
  },0);

  return tr;
}

function currentFilters(){ return { q:($('#search').value||'').toLowerCase(), A:($('#fA').value||'').toLowerCase(), B:($('#fB').value||'').toLowerCase(), C:($('#fC').value||'').toLowerCase(), group:$('#groupFilter').value } }
function matchesFilters(row,f){
  const fields=[row.id,row.kurz_raw,row.beschreibung_raw,row.einheit,row.einheitInfo,String(row.preis),row.hinweis_raw].map(x=>String(x||'').toLowerCase());
  const [a,b,c]=[fields[0],fields[1],fields[2]];
  const global=!f.q||fields.some(v=>v.includes(f.q));
  const spec=(!f.A||a.includes(f.A))&&(!f.B||b.includes(f.B))&&(!f.C||c.includes(f.C));
  return global&&spec;
}
function fillGroupFilter(){
  const sel=$('#groupFilter'); const val=sel.value;
  sel.innerHTML='<option value="">Obergruppe (alle)</option>'+GROUPS.map(g=>`<option value="${g.groupId}">${escapeHtml(g.title||'')}</option>`).join('');
  if([...sel.options].some(o=>o.value===val)) sel.value=val;
}
function render(){
  const body=$('#rows'); body.innerHTML='';
  const f=currentFilters(); let rendered=0, groupCount=0, posCount=0;
  const sortPreis = SORT.col==='preis' ? SORT.dir : 0;
  const groups=GROUPS.filter(g=>!f.group||g.groupId===f.group);
  for(const g of groups){
    const groupMatch=matchesFilters({id:g.groupId,kurz_raw:g.title,beschreibung_raw:g.title,einheit:'',einheitInfo:'',preis:'',hinweis_raw:''},f);
    let children=g.children.filter(c=>matchesFilters(c,f));
    if(sortPreis) children=children.slice().sort((a,b)=>(parseEuro(a.preis)-parseEuro(b.preis))*sortPreis);
    if(!groupMatch && !children.length) continue;
    body.appendChild(trGroup(g, f)); rendered++; groupCount++;
    const list=groupMatch?g.children:children;
    for(const c of list){ body.appendChild(trChild(c, f)); rendered++; posCount++; }
  }
  if(rendered===0){ const tr=document.createElement('tr'); tr.className='empty'; tr.innerHTML='<td colspan="9">Keine Positionen gefunden.</td>'; body.appendChild(tr); }
  $('#count').textContent=`${groupCount} Gruppen · ${posCount} Positionen sichtbar`;
  saveFilters();
}

/* >>> Immer zum Listenanfang springen */
function jumpToTop(){ $('#tableWrap').scrollTop = 0; }

function getCurrentPriceListLabel(){
  const sheet = (CURRENT_SHEET || '').trim();
  const currentFileEl = $('#currentFile');
  const fileText = currentFileEl && currentFileEl.textContent ? currentFileEl.textContent.trim() : '';
  return sheet || fileText || '–';
}

function updatePrintFooterMeta(){
  const el = document.getElementById('printPriceListName');
  if(!el) return;
  el.textContent = getCurrentPriceListLabel();
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
  GROUPS=buildGroups(rows); fillGroupFilter(); render();
  SELECTED.clear(); renderSummary(false);
  requestAnimationFrame(()=>{ jumpToTop(); recomputeChromeOffset(); });
  updatePrintFooterMeta();
  saveState({sheet:name});
}

/* ================= Events ============================= */
$('#manualBtn').addEventListener('click', () => $('#file').click());

$('#file').addEventListener('change', async e=>{
  const f=e.target.files&&e.target.files[0]; if(!f) return;
  if(!/\.(xlsx|xlsm)$/i.test(f.name)){ setStatus('warn','Nur <b>.xlsx</b> oder <b>.xlsm</b> erlaubt.',3500); e.target.value=''; return; }
  setStatus('info','Lade Datei…',2500);
  $('#currentFile').textContent=f.name;
  updatePrintFooterMeta();
  const wb=await loadWorkbook(f); await loadFromSelectedSheet(wb);
  setDrawer(false); setStatus('ok','Bereit.',1500);
  $('#manualBtn').style.display='none';
});

['#search','#fA','#fB','#fC'].forEach(sel=>{
  document.querySelector(sel).addEventListener('input',debounced(()=>{ render(); },150));
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
  const hasSelection = SELECTED.size>0;
  if(hasFilters||hasSelection){
    const ok=confirm('Blatt wechseln? Alle Filter und markierten Positionen werden zurückgesetzt.');
    if(!ok){ sheetSel.value=lastSheetValue??sheetSel.value; return; }
  }
  await loadFromSelectedSheet(); setDrawer(false); setStatus('ok','Blatt gewechselt.',1500);
});

$('#reset').addEventListener('click', ()=>{
  if(!confirm('Alle Filter, Mengen und markierten Positionen werden zurückgesetzt. Fortfahren?')) return;
  $('#search').value=$('#fA').value=$('#fB').value=$('#fC').value=''; $('#groupFilter').value=''; SORT={col:null,dir:1};
  render(); SELECTED.clear(); renderSummary(false); setDrawer(false); setStatus('info','Zurückgesetzt.',1500);
  requestAnimationFrame(()=>{ jumpToTop(); recomputeChromeOffset(); });
});

function setDrawer(open){
  DRAWER_OPEN=!!open;
  $('#summary').classList.toggle('open',DRAWER_OPEN);
  document.body.classList.toggle('drawer-open',DRAWER_OPEN);
  saveState(Object.assign(loadState(),{drawerOpen:DRAWER_OPEN}));
  requestAnimationFrame(recomputeChromeOffset);
}
document.addEventListener('click',(e)=>{ if(e.target && (e.target.id==='toggleDrawer' || e.target.closest('#toggleDrawer'))) setDrawer(!DRAWER_OPEN); });
$('#drawerHead').addEventListener('dblclick',()=>setDrawer(!DRAWER_OPEN));
document.addEventListener('keydown',(e)=>{ if(e.altKey && (e.key==='o' || e.key==='O')){ e.preventDefault(); setDrawer(!DRAWER_OPEN); } });

['#bvhInput','#auftragInput','#erstellerInput','#betreffInput'].forEach(sel=>{
  document.querySelector(sel).addEventListener('input',debounced(saveMeta,200));
});

/* Zusammenfassung + Feedback */
let lastSum = 0;
function showToast(){ const t=$('#toast'); t.classList.add('show'); setTimeout(()=>t.classList.remove('show'), 1100); }
function pulseHead(){ const h=$('#drawerHead'); h.classList.add('pulse'); setTimeout(()=>h.classList.remove('pulse'), 800); }
function updateDelta(sum){
  const el=$('#selSum'); el.classList.remove('sum-up','sum-down');
  if(sum>lastSum){ el.classList.add('sum-up'); setTimeout(()=>el.classList.remove('sum-up'),1000); }
  else if(sum<lastSum){ el.classList.add('sum-down'); setTimeout(()=>el.classList.remove('sum-down'),1000); }
  lastSum = sum;
}

function renderSummary(feedback){
  const wrap=$('#summaryTableWrap');
  const items=[...SELECTED.values()].sort((a,b)=> String(a.id).localeCompare(String(b.id),'de',{numeric:true,sensitivity:'base'}));
  let sum=0;
  const rows=items.map(it=>{ sum+=it.totalNum;
    const kurzHTML = linkify(escapeHtml(it.kurz||''));
    const beschrHTML = linkify(escapeHtml(it.beschreibung||''));
    return `<tr>
      <td style="width:120px">${escapeHtml(it.id)}</td>
      <td style="min-width:220px"><div><b>${kurzHTML}</b></div><div class="desc">${beschrHTML}</div></td>
      <td class="right" style="width:120px">${fmtPrice(it.preisNum)}</td>
      <td class="right" style="width:90px">${fmtQty(it.qtyNum)} ${escapeHtml(it.eh)}</td>
      <td class="right${it.totalNum<0?' neg':''}" style="width:140px">${fmtPrice(it.totalNum)}</td>
    </tr>`; }).join('');
  wrap.innerHTML = `
    <table>
      <thead><tr><th>Art.Nr.</th><th>Bezeichnung (Kurztext + Beschreibung)</th><th class="right">EH-Preis</th><th class="right">Menge</th><th class="right">Gesamt</th></tr></thead>
      <tbody>${rows||''}</tbody>
      <tfoot><tr class="tot"><td colspan="4" class="right">Summe</td><td class="right${sum<0?' neg':''}">${fmtPrice(sum)}</td></tr></tfoot>
    </table>`;
  $('#selCount').textContent=String(items.length);
  $('#selSum').textContent=sum.toLocaleString('de-AT',{style:'currency',currency:'EUR'});
  if(feedback){ pulseHead(); showToast(); updateDelta(sum); setStatus('ok','Bereit.',1500); }
}

/* ============== Drucken (ein Fenster) ============== */
function buildPrintDoc(){
  const items=[...SELECTED.values()].sort((a,b)=> String(a.id).localeCompare(String(b.id),'de',{numeric:true,sensitivity:'base'}));
  if(items.length===0) return null;

  const BVH=($('#bvhInput')?.value||'').trim()||'–';
  const AUF=($('#auftragInput')?.value||'').trim()||'–';
  const ERS=($('#erstellerInput')?.value||'').trim()||'–';
  const BETR=($('#betreffInput')?.value||'').trim()||'–';
  const DAT=new Date().toLocaleDateString('de-AT');
  const SHEET = CURRENT_SHEET || '–';
  const PRICE_LIST = getCurrentPriceListLabel();

  let tbody=''; let total=0;
  items.forEach(it=>{
    total+=it.totalNum;
    const kurzHTML = linkify(escapeHtml(it.kurz||''));
    const beschrHTML = linkify(escapeHtml(it.beschreibung||''));
    tbody+=`<tr class="item-row">
      <td>${escapeHtml(it.id)}</td>
      <td><div><b>${kurzHTML}</b></div><div class="desc">${beschrHTML}</div></td>
      <td class="right">${fmtPrice(it.preisNum)}</td>
      <td class="right">${fmtQty(it.qtyNum)} ${escapeHtml(it.eh)}</td>
      <td class="right${it.totalNum<0?' neg':''}">${fmtPrice(it.totalNum)}</td>
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
    .desc{ white-space:pre-wrap; }
    .desc a{ text-decoration:underline; word-break:break-all; }

    .grand-total{ page-break-inside: avoid; margin-top:10px; border-top:2px solid #bbb; display:grid; grid-template-columns:1fr 140px; gap:8px; align-items:center; font-weight:700; }
    .grand-total .label{ text-align:right; padding-right:8px; }
    .grand-total .value{ text-align:right; }

    .note{ margin-top:12px; font-size:11px; color:#374151; }

    .page-footer{ position:running(summaryFooter); font-size:11px; color:#111; display:flex; justify-content:space-between; align-items:flex-end; padding:1.2mm 0 0.8mm; margin:0; width:100%; box-sizing:border-box; gap:16px; flex-wrap:wrap; }
    .page-footer-left{ flex:1 1 auto; }
    .page-footer-right{ white-space:nowrap; font-variant-numeric:tabular-nums; }
    .page-footer-right .page-num,
    .page-footer-right .page-total{ display:inline-block; min-width:1.6em; text-align:right; }
    .page-footer-right .page-num::after{ content:counter(page, decimal); }
    @media screen { .page-footer{ display:none; } }
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
        <img src="Logo_Haas.jpg" class="logo" alt="Logo" />
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

    <div class="grand-total"><div class="label">Gesamtsumme</div><div class="value${total<0?' neg':''}">${(total).toLocaleString('de-AT',{style:'currency',currency:'EUR'})}</div></div>

    <div class="note">
      Der Preis versteht sich inkl. 20% MwSt.
      Änderungen, Irrtümer und Preisänderungen vorbehalten.
      <br>Hinweis: Grundlage der Preise laut <b>${escapeHtml(SHEET)}</b>.
    </div>
    </main>

    <footer class="page-footer" aria-hidden="true">
      <div class="page-footer-left">www.haas-fertigbau.at | Preise gemäß ${escapeHtml(PRICE_LIST)}</div>
      <div class="page-footer-right">Seite <span class="page-num"></span> / <span class="page-total" data-total="–">–</span></div>
    </footer>

    <script>
      (function(){
        const TOP_MARGIN_MM = 10;
        const BOTTOM_MARGIN_MM = 12;
        const CONTENT_BOTTOM_BUFFER_MM = 12;
        let pxPerMm = 0;

        function ensurePxPerMm(){
          if(pxPerMm) return pxPerMm;
          const probe = document.createElement('div');
          probe.style.cssText = 'position:absolute;visibility:hidden;height:1mm;width:0;padding:0;margin:0;border:0;';
          document.body.appendChild(probe);
          pxPerMm = probe.getBoundingClientRect().height || 0;
          probe.remove();
          return pxPerMm;
        }

        function computePageTotal(){
          const main = document.querySelector('main');
          const totalNode = document.querySelector('.page-total');
          if(!main || !totalNode) return;

          const scale = ensurePxPerMm();
          if(!scale) return;

          const printableHeight = Math.max(((297 - (TOP_MARGIN_MM + BOTTOM_MARGIN_MM)) * scale) - (CONTENT_BOTTOM_BUFFER_MM * scale), 1);
          const contentHeight = main.scrollHeight;
          if(!contentHeight || !Number.isFinite(contentHeight)) return;

          const epsilon = scale * 0.5;
          const pages = Math.max(1, Math.ceil((contentHeight + epsilon) / printableHeight));
          const text = String(pages);
          totalNode.textContent = text;
          totalNode.setAttribute('data-total', text);
        }

        function scheduleCompute(){
          computePageTotal();
          setTimeout(computePageTotal, 120);
        }

        if(document.readyState === 'complete') scheduleCompute();
        else window.addEventListener('load', scheduleCompute, {once:true});

        if(document.fonts && document.fonts.ready){
          document.fonts.ready.then(scheduleCompute).catch(()=>{});
        }

        window.addEventListener('beforeprint', scheduleCompute);

        const mq = window.matchMedia ? window.matchMedia('print') : null;
        if(mq){
          const handler = (ev)=>{ if(ev.matches) scheduleCompute(); };
          if(typeof mq.addEventListener === 'function') mq.addEventListener('change', handler);
          else if(typeof mq.addListener === 'function') mq.addListener(handler);
        }
      })();
    </script>

  </body></html>`;
}

document.getElementById('printSummary').addEventListener('click', async ()=>{
  if(SELECTED.size===0){ setStatus('warn','Keine markierten Positionen zum Drucken.',3000); return; }
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
  const state = loadState();
  if(state.filters){ $('#search').value=state.filters.q||''; $('#fA').value=state.filters.A||''; $('#fB').value=state.filters.B||''; $('#fC').value=state.filters.C||''; $('#groupFilter').value=state.filters.group||''; }
  if(state.meta){ $('#bvhInput').value=state.meta.bvh||''; $('#auftragInput').value=state.meta.auftrag||''; $('#erstellerInput').value=state.meta.ersteller||''; $('#betreffInput').value=state.meta.betreff||''; }
  if(state.drawerOpen===true){ setDrawer(true); } else { setDrawer(false); }

  /* erste Messung vor Autoload */
  recomputeChromeOffset();
  updatePrintFooterMeta();

  setStatus('info','Automatisches Laden der <b>Artikelpreisliste.xlsx</b> (falls vorhanden).',2500);
  try{
    const res=await fetch(DEFAULT_FILE,{cache:'no-store'});
    if(res.ok){
      const buf=await res.arrayBuffer();
      $('#currentFile').textContent=DEFAULT_FILE;
      const wb=await loadWorkbook(buf);
      if(state.sheet && wb.SheetNames.includes(state.sheet)) $('#sheetSel').value=state.sheet;
      await loadFromSelectedSheet(wb);
      if(state.filters && state.filters.group) $('#groupFilter').value=state.filters.group;
      render(); setDrawer(false); setStatus('ok','Bereit.',1500);
      requestAnimationFrame(()=>{ jumpToTop(); recomputeChromeOffset(); });
      $('#manualBtn').style.display='none';
    }else{
      setStatus('warn',`Standarddatei <b>${DEFAULT_FILE}</b> konnte nicht geladen werden.`,4000);
      $('#manualBtn').style.display='inline-flex';
    }
  }catch{
    setStatus('warn',`Standarddatei <b>${DEFAULT_FILE}</b> konnte nicht geladen werden.`,4000);
    $('#manualBtn').style.display='inline-flex';
  }
});
