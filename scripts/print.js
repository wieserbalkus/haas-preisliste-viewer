
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

// Drucktabelle: Spaltenklassen für gezielte Layout-Anpassungen in print.css
// col-artnr (Artikelnummer) · col-kurz (Kurztext) · col-beschr (Beschreibung)
// col-eh (Einheit) · col-ehinfo (Einheitsinfo) · col-preis (Einheitspreis)
// col-hinweis (Hinweise)
function buildListPrintMarkup(mode){
  const groups=collectPrintableGroups(mode);
  if(!groups.length) return '';
  const parts=['<table class="print-table">',
    '<thead><tr><th class="col-artnr">Art.Nr.</th><th class="col-kurz">Kurztext</th><th class="col-beschr">Beschreibung</th><th class="col-eh">EH</th><th class="col-ehinfo">EHI</th><th class="col-preis right">Preis</th><th class="col-hinweis">Hinweis</th></tr></thead>',
    '<tbody>'];
  for(const {group,rows} of groups){
    parts.push(`<tr class="group-row"><td colspan="7"><strong>${escapeHtml(group.groupId)} – ${escapeHtml(group.title||'')}</strong></td></tr>`);
    for(const row of rows){
      const state=ensureRowState(row);
      const editable=isSonderEditable(row.id);
      const preisSource=editable?state.preis:row.preis;
      const preisText=fmtPrice(preisSource)||'–';
      const kurzText=printableText(editable?state.kurz:row.kurz_raw);
      const beschrText=printableText(editable?state.beschreibung:row.beschreibung_raw);
      const ehText=escapeHtml((editable?state.einheit:row.einheit)||'');
      const ehInfoText=escapeHtml((editable?state.einheitInfo:row.einheitInfo)||'');
      const hinweisText=printableText(row.hinweis_raw||'');
      parts.push(`<tr>
        <td class="col-artnr">${escapeHtml(row.id||'')}</td>
        <td class="col-kurz">${kurzText}</td>
        <td class="col-beschr">${beschrText}</td>
        <td class="col-eh">${ehText}</td>
        <td class="col-ehinfo">${ehInfoText}</td>
        <td class="col-preis right">${preisText}</td>
        <td class="col-hinweis desc">${hinweisText}</td>
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

let listPrintPageStyle=null;

function removeListPrintPageRule(){
  if(listPrintPageStyle){
    try{ listPrintPageStyle.remove(); }
    catch{}
    listPrintPageStyle=null;
  }
}

function applyListPrintPageRule(){
  if(listPrintPageStyle) return;
  const style=document.createElement('style');
  style.setAttribute('data-print-page','list');
  style.textContent='@page { size: A4 landscape; margin: 10mm 3mm; }';
  document.head.appendChild(style);
  listPrintPageStyle=style;
}

function clearPrintState(){
  const portal=document.getElementById('printPortal');
  if(portal) portal.innerHTML='';
  delete document.body.dataset.printMode;
  delete document.body.dataset.printListScope;
  if(document.body.hasAttribute('data-print-mode')) document.body.removeAttribute('data-print-mode');
  if(document.body.hasAttribute('data-print-list-scope')) document.body.removeAttribute('data-print-list-scope');
  removeListPrintPageRule();
}

function triggerListPrint(mode){
  if(!mountPrintTable(mode)){
    setStatus('warn','Keine Daten zum Drucken.',3000);
    return;
  }
  document.body.dataset.printMode='list';
  document.body.dataset.printListScope=mode;
  applyListPrintPageRule();
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
  const logoUrl = versionedAsset('./assets/img/Logo_Haas.jpg');

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
        <img src="'+logoUrl+'" class="logo" alt="Logo" />
        <div class="created"><b>Erstellt am:</b> ${escapeHtml(DAT)}</div>
      </div>
    </div>

    <table>
      <thead><tr>
        <!-- ANPASSUNG: Breite Art.Nr. (Sollwert ca. 5 %) hier ändern -->
        <th style="width:30px">Art.Nr.</th>
        <!-- ANPASSUNG: Breite Kurztext (20 %) + Beschreibung (40 %) bei Bedarf hier anpassen -->
        <th>Kurztext und Beschreibung</th>
        <!-- ANPASSUNG: Breite EH-Preis (Sollwert ca. 5 %) hier ändern -->
        <th class="right" style="width:50px">EH-Preis</th>
        <!-- ANPASSUNG: Breite Menge (entspricht ehemaliger 5 %-Spalte) hier ändern -->
        <th class="right" style="width:80px">Menge</th>
        <!-- ANPASSUNG: Breite Gesamtpreis (Restanteile, z. B. 20 %) hier ändern -->
        <th class="right" style="width:65px">Gesamt</th>
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

function doPrintSummary(){
  if(getBasketSize()===0){ setStatus('warn','Keine markierten Positionen zum Drucken.',3000); return false; }
  const html = buildPrintDoc(); if(!html){ setStatus('warn','Nichts zu drucken.',3000); return false; }

  const iframe = document.createElement('iframe');
  Object.assign(iframe.style,{position:'fixed',right:'0',bottom:'0',width:'0',height:'0',border:'0'});
  iframe.setAttribute('aria-hidden','true');
  document.body.appendChild(iframe);

  const win = iframe.contentWindow, doc = win.document;
  doc.open(); doc.write(html); doc.close();

  const done = () => { try{ iframe.remove(); }catch{} };
  const safePrint = () => { try{ win.focus(); win.onafterprint = done; win.print(); } catch { done(); } };

  if (doc.readyState === 'complete') { setTimeout(safePrint, 250); }
  else { win.addEventListener('load', () => setTimeout(safePrint, 250), {once:true}); }
  return true;
}

document.getElementById('printSummary').addEventListener('click', ()=>{
  doPrintSummary();
});

document.addEventListener('keydown',(e)=>{
  if(String(e.key||'').toLowerCase()!=='p') return;
  if(!(e.ctrlKey || e.metaKey)) return;
  if(e.altKey) return;
  if(getBasketSize()<=0){ return; }
  e.preventDefault();
  doPrintSummary();
});

