(() => {
  'use strict';

  const REPO_OWNER = 'MrMartinPet';
  const REPO_NAME = 'Frekvensanalys';
  const REPO_BRANCH = 'main';
  const API_ROOT = `https://api.github.com/repos/${REPO_OWNER}/${REPO_NAME}/contents/`;

  const CATEGORIES = [
    { key:'Processarbete', color:'#fb9846' },
    { key:'Materialförsörjning', color:'#c5534d' },
    { key:'Problem', color:'#285487' },
    { key:'Administration', color:'#4eacc4' },
    { key:'Övrigt', color:'#9cbc5b' }
  ];

  const LEGACY_CATEGORY_RULES = [
    { re:/(material|logistik|truck|lager|försörj|forsorj|transport|hämta|hamta)/i, target:'Materialförsörjning' },
    { re:/(problem|stopp|fel|avvik|vänt|vant|saknas|brist)/i, target:'Problem' },
    { re:/(admin|plan|rapport|dator|system|dokument|mail|mejl)/i, target:'Administration' },
    { re:/(möte|mote|städ|stad|övr|ovr|gång|gang|rast)/i, target:'Övrigt' }
  ];

  const $ = id => document.getElementById(id);
  const els = {
    setupView:$('setupView'), analysisView:$('analysisView'), summaryView:$('summaryView'),
    sourceSelect:$('sourceSelect'), refreshSourcesBtn:$('refreshSourcesBtn'), sourceStatus:$('sourceStatus'),
    leaderInput:$('leaderInput'), flowInput:$('flowInput'), dateInput:$('dateInput'), startBtn:$('startBtn'),
    categoryPreview:$('categoryPreview'), themeBtn:$('themeBtn'), exportBtnTop:$('exportBtnTop'), resetBtnTop:$('resetBtnTop'),
    liveTitle:$('liveTitle'), sessionMeta:$('sessionMeta'), activeActivity:$('activeActivity'), activeTimer:$('activeTimer'), sessionTimer:$('sessionTimer'),
    addActivitiesBtn:$('addActivitiesBtn'), finishBtn:$('finishBtn'), categoryTabs:$('categoryTabs'), activityPanel:$('activityPanel'),
    liveCategoryBars:$('liveCategoryBars'), trackedShare:$('trackedShare'), recentHistory:$('recentHistory'),
    summaryMeta:$('summaryMeta'), summaryExportBtn:$('summaryExportBtn'), newAnalysisBtn:$('newAnalysisBtn'),
    sumTotal:$('sumTotal'), sumTracked:$('sumTracked'), sumSwitches:$('sumSwitches'), sumUsed:$('sumUsed'), activityRanking:$('activityRanking'), categoryChart:$('categoryChart'),
    modalBackdrop:$('modalBackdrop'), addCategorySelect:$('addCategorySelect'), newActivitiesInput:$('newActivitiesInput'),
    closeModalBtn:$('closeModalBtn'), cancelModalBtn:$('cancelModalBtn'), saveActivitiesBtn:$('saveActivitiesBtn'), toast:$('toast')
  };

  const state = {
    sources:[],
    selectedSource:null,
    activities:emptyActivityMap(),
    selectedCategory:'Processarbete',
    header:{leader:'', flow:'', date:'', source:''},
    totals:{},
    counts:{},
    active:null,
    activeStart:0,
    sessionStart:0,
    sessionEnd:0,
    history:[],
    raf:0,
    chart:null,
    mode:'setup',
    lastLivePaint:0
  };

  function emptyActivityMap(){
    return Object.fromEntries(CATEGORIES.map(c => [c.key, []]));
  }

  function normalize(s){
    return String(s ?? '')
      .normalize('NFD').replace(/[\u0300-\u036f]/g,'')
      .toLowerCase().replace(/[^a-z0-9]/g,'');
  }

  function categoryFromHeader(value){
    const n = normalize(value);
    return CATEGORIES.find(c => normalize(c.key) === n)?.key || null;
  }

  function legacyCategoryToStandard(value){
    const direct = categoryFromHeader(value);
    if(direct) return direct;
    const s = String(value || '').trim();
    for(const rule of LEGACY_CATEGORY_RULES){ if(rule.re.test(s)) return rule.target; }
    return 'Processarbete';
  }

  function cleanActivity(v){
    return String(v ?? '').replace(/\s+/g,' ').trim();
  }

  function dedupe(list){
    const seen = new Set();
    return list.filter(v => {
      const k = normalize(v);
      if(!k || seen.has(k)) return false;
      seen.add(k); return true;
    });
  }

  function parseWorkbook(arrayBuffer){
    const wb = XLSX.read(arrayBuffer,{type:'array'});
    const sheet = wb.Sheets[wb.SheetNames[0]];
    const rows = XLSX.utils.sheet_to_json(sheet,{header:1,defval:'',raw:false});
    if(!rows.length) throw new Error('Excel-filen är tom.');

    const result = emptyActivityMap();
    let mode = 'wide';
    let headerRow = -1;
    let headerCols = {};

    for(let r=0;r<Math.min(rows.length,30);r++){
      const cols = {};
      (rows[r] || []).forEach((cell,idx) => {
        const cat = categoryFromHeader(cell);
        if(cat) cols[cat] = idx;
      });
      if(Object.keys(cols).length >= 2){ headerRow = r; headerCols = cols; break; }
    }

    if(headerRow >= 0){
      for(const cat of CATEGORIES){
        const col = headerCols[cat.key];
        if(col === undefined) continue;
        for(let r=headerRow+1;r<rows.length;r++){
          const activity = cleanActivity(rows[r]?.[col]);
       if(activity) result[cat.key].push(activity);
        }
        result[cat.key] = dedupe(result[cat.key]);
      }
      return {activities:result, format:'wide', warning:Object.keys(headerCols).length < CATEGORIES.length ? 'Filen saknar någon standardkategori. Saknade kategorier visas tomma.' : ''};
    }

    mode = 'legacy';
    let legacyHeader = -1, catIdx = -1, actIdx = -1;
    for(let r=0;r<Math.min(rows.length,30);r++){
      const normalized = (rows[r] || []).map(normalize);
      const ci = normalized.indexOf('kategori');
      const ai = normalized.indexOf('aktivitet');
      if(ci >= 0 && ai >= 0){ legacyHeader = r; catIdx = ci; actIdx = ai; break; }
    }

    if(legacyHeader >= 0){
      for(let r=legacyHeader+1;r<rows.length;r++){
        const activity = cleanActivity(rows[r]?.[actIdx]);
        if(!activity) continue;
        const target = legacyCategoryToStandard(rows[r]?.[catIdx]);
        result[target].push(activity);
      }
      CATEGORIES.forEach(c => result[c.key] = dedupe(result[c.key]));
      return {activities:result, format:mode, warning:'Gammalt Excel-format upptäckt. Aktiviteterna har automatiskt sorterats in i de fem standardkategorierna. Nytt format rekommenderas.'};
    }

    throw new Error('Okänt Excel-format. Använd fem kolumner med standardkategorierna som rubriker.');
  }

  async function loadSources(){
    setSourceStatus('Läser in Excel-filer…','');
    els.sourceSelect.disabled = true;
    els.startBtn.disabled = true;
    try{
      const res = await fetch(`${API_ROOT}?ref=${encodeURIComponent(REPO_BRANCH)}`,{cache:'no-store',headers:{'Accept':'application/vnd.github+json'}});
      if(!res.ok) throw new Error(`GitHub svarade ${res.status}`);
      const data = await res.json();
      state.sources = data
        .filter(x => x.type === 'file' && /\.xlsx$/i.test(x.name) && !/^frekvensanalys_export/i.test(x.name))
        .sort((a,b) => a.name.localeCompare(b.name,'sv'));
      renderSourceOptions();
      if(state.sources.length){
        setSourceStatus(`${state.sources.length} Excel-filer hittades. Välj källa.`,'ok');
      }else{
        setSourceStatus('Inga .xlsx-filer hittades i repots rot.','warn');
      }
    }catch(err){
      console.error(err);
      setSourceStatus('Kunde inte läsa källistan från GitHub. Försök igen.','error');
      els.sourceSelect.innerHTML = '<option value="">Kunde inte läsa källor</option>';
    }finally{
      els.sourceSelect.disabled = false;
    }
  }

  function renderSourceOptions(){
    const previous = state.selectedSource?.name || '';
    els.sourceSelect.innerHTML = '<option value="">Välj Excel-fil…</option>' + state.sources.map((s,i) => `<option value="${i}">${escapeHtml(s.name)}</option>`).join('');
    if(previous){
      const idx = state.sources.findIndex(s => s.name === previous);
      if(idx >= 0){ els.sourceSelect.value = String(idx); }
    }
  }

  async function loadSelectedSource(){
    const idx = Number(els.sourceSelect.value);
    if(!Number.isInteger(idx) || !state.sources[idx]){
      state.selectedSource = null;
      state.activities = emptyActivityMap();
      renderCategoryPreview();
      els.startBtn.disabled = true;
      setSourceStatus('Välj en källa.','');
      return;
    }
    const source = state.sources[idx];
    state.selectedSource = source;
    els.startBtn.disabled = true;
    setSourceStatus(`Läser ${source.name}…`,'');
    try{
      const res = await fetch(source.download_url,{cache:'no-store'});
      if(!res.ok) throw new Error(`Kunde inte hämta filen (${res.status}).`);
      const parsed = parseWorkbook(await res.arrayBuffer());
      state.activities = parsed.activities;
      state.selectedCategory = firstNonEmptyCategory();
      renderCategoryPreview();
      const total = totalConfiguredActivities();
      setSourceStatus(`${source.name}: ${total} aktiviteter inlästa.${parsed.warning ? ' ' + parsed.warning : ''}`, parsed.warning ? 'warn' : 'ok');
      els.startBtn.disabled = total === 0;
    }catch(err){
      console.error(err);
      state.activities = emptyActivityMap();
      renderCategoryPreview();
      setSourceStatus(err.message || 'Kunde inte läsa Excel-filen.','error');
      els.startBtn.disabled = true;
    }
  }

  function totalConfiguredActivities(){
    return CATEGORIES.reduce((sum,c) => sum + state.activities[c.key].length,0);
  }

  function firstNonEmptyCategory(){
    return CATEGORIES.find(c => state.activities[c.key].length)?.key || CATEGORIES[0].key;
  }

  function setSourceStatus(text,kind){
    els.sourceStatus.textContent = text;
    els.sourceStatus.className = `status-box${kind ? ' ' + kind : ''}`;
  }

  function renderCategoryPreview(){
    els.categoryPreview.innerHTML = CATEGORIES.map(c => `
      <div class="preview-row">
        <span class="preview-dot" style="background:${c.color}"></span>
        <span class="preview-name">${c.key}</span>
        <span class="preview-count">${state.activities[c.key].length} aktiviteter</span>
      </div>`).join('');
  }

  function startSession(){
    if(!state.selectedSource || totalConfiguredActivities() === 0) return;
    state.header = {
      leader:els.leaderInput.value.trim(),
      flow:els.flowInput.value.trim(),
      date:els.dateInput.value || todayISO(),
      source:state.selectedSource.name
    };
    state.totals = {};
    state.counts = {};
    CATEGORIES.forEach(c => state.activities[c.key].forEach(a => {
      const id = activityId(c.key,a); state.totals[id] = 0; state.counts[id] = 0;
    }));
    state.active = null;
    state.activeStart = 0;
    state.sessionStart = performance.now();
    state.sessionEnd = 0;
    state.history = [];
    state.mode = 'analysis';
    showView('analysis');
    els.exportBtnTop.disabled = false;
    renderAnalysis();
    tick();
    persistDraft();
  }

  function showView(name){
    els.setupView.hidden = name !== 'setup';
    els.analysisView.hidden = name !== 'analysis';
    els.summaryView.hidden = name !== 'summary';
  }

  function renderAnalysis(){
    els.liveTitle.textContent = state.header.flow || 'Frekvensanalys';
    els.sessionMeta.textContent = [state.header.leader || 'Ej angiven observatör', state.header.date, state.header.source].filter(Boolean).join(' · ');
    renderCategoryTabs();
    renderActivityPanel();
    renderLiveBars();
    renderRecentHistory();
    updateActiveHeader();
  }

  function renderCategoryTabs(){
    els.categoryTabs.innerHTML = CATEGORIES.map(c => `
      <button type="button" class="category-tab${state.selectedCategory===c.key?' active':''}" data-category="${escapeAttr(c.key)}" style="--cat-color:${c.color}">
        ${escapeHtml(c.key)}
        <span class="tab-count">${state.activities[c.key].length} aktiviteter</span>
      </button>`).join('');
    els.categoryTabs.querySelectorAll('.category-tab').forEach(btn => btn.addEventListener('click',() => {
      state.selectedCategory = btn.dataset.category;
      renderCategoryTabs();
      renderActivityPanel();
    }));
  }

  function renderActivityPanel(){
    const cat = getCategory(state.selectedCategory);
    const list = state.activities[state.selectedCategory] || [];
    els.activityPanel.style.setProperty('--cat-color',cat.color);
    const head = `<div class="activity-head"><div class="activity-head-title"><span class="category-swatch" style="background:${cat.color}"></span><h2>${escapeHtml(cat.key)}</h2></div><span class="pill">${list.length}</span></div>`;
    if(!list.length){
      els.activityPanel.innerHTML = head + '<div class="empty-state">Inga aktiviteter i denna kategori. Använd “Lägg till aktiviteter” under pågående analys.</div>';
      return;
    }
    els.activityPanel.innerHTML = head + `<div class="activity-grid">${list.map(a => {
      const id = activityId(cat.key,a);
      const isActive = state.active?.id === id;
      return `<button type="button" class="activity-btn${isActive?' active':''}" data-id="${escapeAttr(id)}" style="--cat-color:${cat.color}">
        <span class="activity-name">${escapeHtml(a)}</span>
        <span class="activity-time" id="activityTime_${safeDomId(id)}">${fmt(state.totals[id]||0)}</span>
      </button>`;
    }).join('')}</div>`;
    els.activityPanel.querySelectorAll('.activity-btn').forEach(btn => btn.addEventListener('click',() => selectActivityById(btn.dataset.id)));
  }

  function selectActivityById(id){
    const info = activityInfoFromId(id);
    if(!info) return;
    const now = performance.now();
    closeActiveSegment(now);
    state.active = {id, category:info.category, name:info.name};
    state.activeStart = now;
    state.counts[id] = (state.counts[id] || 0) + 1;
    updateActiveHeader();
    renderActivityPanel();
    persistDraft();
  }

  function closeActiveSegment(now){
    if(!state.active) return;
    const ms = Math.max(0,now - state.activeStart);
    if(ms > 0){
      state.totals[state.active.id] = (state.totals[state.active.id] || 0) + ms;
      state.history.push({
        id:state.active.id, category:state.active.category, activity:state.active.name,
        start:state.activeStart, end:now, duration:ms
      });
    }
  }

  function tick(){
    if(state.mode !== 'analysis') return;
    const now = performance.now();
    const sessionMs = now - state.sessionStart;
    els.sessionTimer.textContent = fmt(sessionMs);
    els.activeTimer.textContent = state.active ? fmt(now - state.activeStart) : '00:00:00';

    const preview = previewTotals(now);
    Object.entries(preview).forEach(([id,ms]) => {
      const el = $(`activityTime_${safeDomId(id)}`);
      if(el) el.textContent = fmt(ms);
    });
    if(now - state.lastLivePaint > 250){
      renderLiveBars(preview,sessionMs);
      renderRecentHistory();
      state.lastLivePaint = now;
    }
    state.raf = requestAnimationFrame(tick);
  }

  function previewTotals(now=performance.now()){
    const totals = {...state.totals};
    if(state.active) totals[state.active.id] = (totals[state.active.id] || 0) + Math.max(0,now-state.activeStart);
    return totals;
  }

  function updateActiveHeader(){
    els.activeActivity.textContent = state.active ? state.active.name : 'Ingen vald';
  }

  function aggregateCategories(totals){
    const out = Object.fromEntries(CATEGORIES.map(c => [c.key,0]));
    Object.entries(totals).forEach(([id,ms]) => {
      const info = activityInfoFromId(id);
      if(info) out[info.category] += ms;
    });
    return out;
  }

  function renderLiveBars(totals=previewTotals(),sessionMs=(state.mode==='analysis'?performance.now()-state.sessionStart:state.sessionEnd-state.sessionStart)){
    const agg = aggregateCategories(totals);
    const tracked = Object.values(agg).reduce((a,b)=>a+b,0);
    const denom = tracked || 1;
    els.liveCategoryBars.innerHTML = CATEGORIES.map(c => {
      const ms = agg[c.key] || 0;
      const pct = ms / denom * 100;
      return `<div class="live-bar-row" style="--cat-color:${c.color}">
        <div><div class="live-bar-label"><span class="live-bar-name">${escapeHtml(c.key)}</span><span>${pct.toFixed(0)}%</span></div><div class="bar-track"><div class="bar-fill" style="width:${pct}%;--cat-color:${c.color};background:${c.color}"></div></div></div>
        <div class="live-bar-time">${fmtShort(ms)}</div>
      </div>`;
    }).join('');
    els.trackedShare.textContent = sessionMs > 0 ? `${Math.min(100,tracked/sessionMs*100).toFixed(0)} %` : '0 %';
  }

  function renderRecentHistory(){
    const rows = state.history.slice(-5).reverse();
    if(!rows.length){ els.recentHistory.textContent = 'Ingen aktivitet registrerad ännu.'; return; }
    els.recentHistory.innerHTML = rows.map(h => `<div class="recent-row"><div><div class="recent-name">${escapeHtml(h.activity)}</div><div>${escapeHtml(h.category)}</div></div><div class="recent-time">${fmtShort(h.duration)}</div></div>`).join('');
  }

  function openAddModal(){
    els.addCategorySelect.innerHTML = CATEGORIES.map(c => `<option value="${escapeAttr(c.key)}"${c.key===state.selectedCategory?' selected':''}>${escapeHtml(c.key)}</option>`).join('');
    els.newActivitiesInput.value = '';
    els.modalBackdrop.classList.remove('hidden');
    setTimeout(() => els.newActivitiesInput.focus(),30);
  }

  function closeAddModal(){ els.modalBackdrop.classList.add('hidden'); }

  function addActivitiesDuringAnalysis(){
    const category = els.addCategorySelect.value;
    const raw = els.newActivitiesInput.value.split(/\r?\n/).map(cleanActivity).filter(Boolean);
    const names = dedupe(raw);
    if(!names.length){ toast('Skriv minst en aktivitet.'); return; }
    const existing = new Set(state.activities[category].map(normalize));
    let added = 0;
    names.forEach(name => {
      if(existing.has(normalize(name))) return;
      state.activities[category].push(name);
      const id = activityId(category,name);
      state.totals[id] = 0;
      state.counts[id] = 0;
      existing.add(normalize(name));
      added++;
    });
    state.selectedCategory = category;
    closeAddModal();
    renderCategoryTabs();
    renderActivityPanel();
    renderLiveBars();
    persistDraft();
    toast(added ? `${added} aktivitet${added===1?'':'er'} tillagda.` : 'Aktiviteterna fanns redan.');
  }

  function finishSession(){
    if(!confirm('Avsluta analysen och lås resultatet?')) return;
    const now = performance.now();
    closeActiveSegment(now);
    state.active = null;
    state.activeStart = 0;
    state.sessionEnd = now;
    state.mode = 'summary';
    cancelAnimationFrame(state.raf);
    clearDraft();
    showView('summary');
    renderSummary();
  }

  function renderSummary(){
    const totalMs = Math.max(0,state.sessionEnd-state.sessionStart);
    const trackedMs = Object.values(state.totals).reduce((a,b)=>a+b,0);
    const used = Object.values(state.totals).filter(v=>v>0).length;
    els.summaryMeta.textContent = [state.header.leader || 'Ej angiven observatör', state.header.flow || 'Ej angivet flöde', state.header.date, state.header.source].join(' · ');
    els.sumTotal.textContent = fmt(totalMs);
    els.sumTracked.textContent = fmt(trackedMs);
    els.sumSwitches.textContent = String(state.history.length);
    els.sumUsed.textContent = String(used);
    renderSummaryChart();
    renderRanking();
  }

  function renderSummaryChart(){
    const agg = aggregateCategories(state.totals);
    const labels = CATEGORIES.map(c => c.key);
    const minutes = CATEGORIES.map(c => +(agg[c.key]/60000).toFixed(2));
    if(state.chart) state.chart.destroy();
    state.chart = new Chart(els.categoryChart.getContext('2d'),{
      type:'bar',
      data:{labels,datasets:[{label:'Minuter',data:minutes,backgroundColor:CATEGORIES.map(c=>c.color),borderWidth:0,borderRadius:4}]},
      options:{
        responsive:true,maintainAspectRatio:false,
        plugins:{legend:{display:false},tooltip:{callbacks:{label:ctx=>`${ctx.raw} min`}}},
        scales:{x:{grid:{display:false},ticks:{color:getCss('--muted'),maxRotation:28,minRotation:0}},y:{beginAtZero:true,grid:{color:getCss('--line')},ticks:{color:getCss('--muted')}}}
      }
    });
  }

  function renderRanking(){
    const rows = [];
    CATEGORIES.forEach(c => state.activities[c.key].forEach(a => {
      const id = activityId(c.key,a), ms = state.totals[id] || 0;
      if(ms > 0) rows.push({category:c.key,name:a,ms,count:state.counts[id]||0});
    }));
    rows.sort((a,b)=>b.ms-a.ms);
    const total = rows.reduce((s,r)=>s+r.ms,0) || 1;
    els.activityRanking.innerHTML = rows.length ? rows.map((r,i)=>`<div class="rank-row"><div class="rank-no">${i+1}</div><div><div class="rank-name">${escapeHtml(r.name)}</div><div class="rank-cat">${escapeHtml(r.category)} · ${r.count} tillfällen</div></div><div class="rank-val">${fmtShort(r.ms)}<span>${(r.ms/total*100).toFixed(1)} %</span></div></div>`).join('') : '<div class="empty-state">Ingen aktivitetstid registrerades.</div>';
  }

  function downloadExcel(){
    if(!state.sessionStart){ toast('Ingen analys att exportera.'); return; }
    const wb = XLSX.utils.book_new();
    const base = Date.now() - (performance.now() - state.sessionStart);
    const endPerf = state.mode === 'analysis' ? performance.now() : state.sessionEnd;
    const totals = state.mode === 'analysis' ? previewTotals(endPerf) : state.totals;
    const hist = state.history.slice();
    if(state.mode === 'analysis' && state.active){
      hist.push({id:state.active.id,category:state.active.category,activity:state.active.name,start:state.activeStart,end:endPerf,duration:endPerf-state.activeStart});
    }

    const meta = [
      ['Frekvensanalys'],
      ['Lagledare / observatör',state.header.leader],
      ['Flöde / område',state.header.flow],
      ['Datum',state.header.date],
      ['Källa',state.header.source],
      ['Exporterat',new Date().toLocaleString('sv-SE')],
      []
    ];
    const details = meta.concat([['Kategori','Aktivitet','Start','Slut','Sekunder']]);
    hist.forEach(h => details.push([
      h.category,h.activity,
      new Date(base + (h.start-state.sessionStart)).toLocaleString('sv-SE'),
      new Date(base + (h.end-state.sessionStart)).toLocaleString('sv-SE'),
      Math.round(h.duration/1000)
    ]));
    XLSX.utils.book_append_sheet(wb,XLSX.utils.aoa_to_sheet(details),'Detaljer');

    const agg = aggregateCategories(totals);
    const tracked = Object.values(agg).reduce((a,b)=>a+b,0) || 1;
    const catRows = [['Kategori','Sekunder','Minuter','Andel %']];
    CATEGORIES.forEach(c => catRows.push([c.key,Math.round(agg[c.key]/1000),+(agg[c.key]/60000).toFixed(2),+(agg[c.key]/tracked*100).toFixed(1)]));
    XLSX.utils.book_append_sheet(wb,XLSX.utils.aoa_to_sheet(catRows),'Kategorier');

    const actRows = [['Kategori','Aktivitet','Sekunder','Minuter','Tillfällen']];
    CATEGORIES.forEach(c => state.activities[c.key].forEach(a => {
      const id=activityId(c.key,a); actRows.push([c.key,a,Math.round((totals[id]||0)/1000),+((totals[id]||0)/60000).toFixed(2),state.counts[id]||0]);
    }));
    XLSX.utils.book_append_sheet(wb,XLSX.utils.aoa_to_sheet(actRows),'Aktiviteter');

    const safeFlow = (state.header.flow || 'analys').replace(/[^a-zA-Z0-9åäöÅÄÖ_-]+/g,'_');
    XLSX.writeFile(wb,`frekvensanalys_${safeFlow}_${state.header.date || todayISO()}.xlsx`);
  }

  function resetAll(){
    if(state.mode === 'analysis' && !confirm('En analys pågår. Nollställa allt?')) return;
    if(state.mode !== 'analysis' && state.sessionStart && !confirm('Nollställa aktuell analys?')) return;
    cancelAnimationFrame(state.raf);
    clearDraft();
    state.selectedSource = null;
    state.activities = emptyActivityMap();
    state.totals={};state.counts={};state.active=null;state.history=[];state.sessionStart=0;state.sessionEnd=0;state.mode='setup';
    els.sourceSelect.value='';
    els.exportBtnTop.disabled=true;
    els.startBtn.disabled=true;
    renderCategoryPreview();
    setSourceStatus('Välj en källa.','');
    showView('setup');
  }

  function newAnalysis(){
    if(!confirm('Starta en ny analys?')) return;
    cancelAnimationFrame(state.raf);clearDraft();
    state.totals={};state.counts={};state.active=null;state.history=[];state.sessionStart=0;state.sessionEnd=0;state.mode='setup';
    els.exportBtnTop.disabled=true;
    showView('setup');
    els.startBtn.disabled = !state.selectedSource || totalConfiguredActivities()===0;
  }

  function activityId(category,name){ return `${category}|||${name}`; }
  function activityInfoFromId(id){
    const pos = String(id).indexOf('|||');
    if(pos < 0) return null;
    return {category:id.slice(0,pos),name:id.slice(pos+3)};
  }
  function getCategory(key){ return CATEGORIES.find(c=>c.key===key) || CATEGORIES[0]; }
  function safeDomId(s){ let h=2166136261; for(let i=0;i<s.length;i++){ h^=s.charCodeAt(i); h=Math.imul(h,16777619); } return (h>>>0).toString(36); }
  function fmt(ms){ const sec=Math.max(0,Math.floor(ms/1000)); const h=String(Math.floor(sec/3600)).padStart(2,'0'); const m=String(Math.floor((sec%3600)/60)).padStart(2,'0'); const s=String(sec%60).padStart(2,'0'); return `${h}:${m}:${s}`; }
  function fmtShort(ms){ const sec=Math.max(0,Math.round(ms/1000)); if(sec<60) return `${sec}s`; const m=Math.floor(sec/60),s=sec%60; if(m<60) return `${m}:${String(s).padStart(2,'0')}`; return `${Math.floor(m/60)}:${String(m%60).padStart(2,'0')}`; }
  function todayISO(){ const d=new Date(); const off=d.getTimezoneOffset(); return new Date(d.getTime()-off*60000).toISOString().slice(0,10); }
  function getCss(name){ return getComputedStyle(document.documentElement).getPropertyValue(name).trim(); }
  function escapeHtml(s){ return String(s ?? '').replace(/[&<>"']/g,ch=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#039;'}[ch])); }
  function escapeAttr(s){ return escapeHtml(s); }
  let toastTimer=0;
  function toast(msg){ els.toast.textContent=msg;els.toast.classList.remove('hidden');clearTimeout(toastTimer);toastTimer=setTimeout(()=>els.toast.classList.add('hidden'),2600); }

  function setTheme(theme){
    document.documentElement.dataset.theme = theme;
    localStorage.setItem('fa_theme',theme);
    els.themeBtn.textContent = theme === 'dark' ? '☀︎' : '◐';
    document.querySelector('meta[name="theme-color"]')?.setAttribute('content',theme==='dark'?'#0b1020':'#ffffff');
    if(state.mode==='summary') renderSummaryChart();
  }

  function persistDraft(){
    if(state.mode!=='analysis') return;
    try{
      const now=performance.now();
      const payload={
        version:2,savedAt:Date.now(),header:state.header,activities:state.activities,selectedCategory:state.selectedCategory,
        totals:previewTotals(now),counts:state.counts,history:state.history.map(h=>({...h})),
        active:state.active,elapsedSession:now-state.sessionStart,elapsedActive:state.active?now-state.activeStart:0
      };
      localStorage.setItem('fa_draft_v2',JSON.stringify(payload));
    }catch(e){ console.warn('Kunde inte autospara',e); }
  }
  function clearDraft(){ localStorage.removeItem('fa_draft_v2'); }

  function wire(){
    els.dateInput.value = todayISO();
    renderCategoryPreview();
    CATEGORIES.forEach(c => els.addCategorySelect.insertAdjacentHTML('beforeend',`<option value="${escapeAttr(c.key)}">${escapeHtml(c.key)}</option>`));
    els.sourceSelect.addEventListener('change',loadSelectedSource);
    els.refreshSourcesBtn.addEventListener('click',loadSources);
    els.startBtn.addEventListener('click',startSession);
    els.addActivitiesBtn.addEventListener('click',openAddModal);
    els.closeModalBtn.addEventListener('click',closeAddModal);
    els.cancelModalBtn.addEventListener('click',closeAddModal);
    els.saveActivitiesBtn.addEventListener('click',addActivitiesDuringAnalysis);
    els.modalBackdrop.addEventListener('click',e=>{if(e.target===els.modalBackdrop)closeAddModal();});
    els.finishBtn.addEventListener('click',finishSession);
    els.exportBtnTop.addEventListener('click',downloadExcel);
    els.summaryExportBtn.addEventListener('click',downloadExcel);
    els.newAnalysisBtn.addEventListener('click',newAnalysis);
    els.resetBtnTop.addEventListener('click',resetAll);
    els.themeBtn.addEventListener('click',()=>setTheme(document.documentElement.dataset.theme==='dark'?'light':'dark'));
    window.addEventListener('beforeunload',e=>{
      if(state.mode==='analysis'){
        persistDraft();
        e.preventDefault();
        e.returnValue='';
      }
    });
    setInterval(()=>{if(state.mode==='analysis')persistDraft();},15000);
  }

  wire();
  setTheme(localStorage.getItem('fa_theme') || 'dark');
  loadSources();
})();
