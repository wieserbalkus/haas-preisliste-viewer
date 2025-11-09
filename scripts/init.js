/* ================= Init =============================== */
let appInitStarted = false;
function clearAppStorage(prefix){
  if(!prefix) return;
  try{
    const clearKeys = (store)=>{
      const keys = [];
      for(let i=0;i<store.length;i++){
        const key = store.key(i);
        if(key && key.startsWith(prefix)){ keys.push(key); }
      }
      keys.forEach(k=>{ try{ store.removeItem(k); }catch{} });
    };
    if(globalThis.localStorage){ clearKeys(globalThis.localStorage); }
    if(globalThis.sessionStorage){ clearKeys(globalThis.sessionStorage); }
  }catch{}
}
async function initApp(){
  if(appInitStarted) return;
  appInitStarted = true;
  startVersionWatcher();
  clearAppStorage('haas_');
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
}

if(document.readyState === 'loading'){
  window.addEventListener('DOMContentLoaded', initApp, {once:true});
}else{
  initApp();
}
