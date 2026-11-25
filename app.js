(() => {
  const $ = id => document.getElementById(id);
  const PRICEBOOK = { effective: 'November 2, 2026', minimumCoreOrder: 96, vsan: { sku: 'VCF-VSAN-8', annualPerTiB: 250 }, vvf: { sku: 'VCF-VSP-FND-B', annualPerCore: 275 }, vcf: { sku: 'VCF-CLD-FND-B', annualPerCore: 485 } };
  const state = { session: { id: 'session', name: 'Estimate', imports: [], capacity: {}, scope: null }, selectedVcenters: new Set(), selectedClusters: new Set(), selectedHosts: new Set(), hosts: [], clusters: [], warnings: [] };
  const fmt = (n, digits = 0) => Number(n).toLocaleString(undefined, { maximumFractionDigits: digits, minimumFractionDigits: digits });
  const usd = n => Number(n).toLocaleString(undefined, { style: 'currency', currency: 'USD', minimumFractionDigits: 2, maximumFractionDigits: 2 });
  const clean = v => String(v ?? '').trim();
  const normalize = s => String(s ?? '').toLowerCase().replace(/[^a-z0-9]/g, '');
  const uid = () => (crypto.randomUUID ? crypto.randomUUID() : `${Date.now()}-${Math.random().toString(16).slice(2)}`);
  const escapeHtml = v => String(v ?? '').replace(/[&<>"']/g, c => ({ '&':'&amp;', '<':'&lt;', '>':'&gt;', '"':'&quot;', "'":'&#39;' }[c]));
  const num = v => { if (typeof v === 'number') return Number.isFinite(v) ? v : NaN; const s = clean(v).replace(/,/g, '').replace(/%$/, ''); if (!s || s === '-' || s.toLowerCase() === 'n/a') return NaN; const n = Number(s); return Number.isFinite(n) ? n : NaN; };
  function setStatus(message, error = false) { const el = $('statusMessage'); el.textContent = message; el.classList.toggle('error', error); }
  function findHeader(headers, patterns) { const normalized = headers.map(normalize); for (const pattern of patterns) { const i = normalized.indexOf(normalize(pattern)); if (i >= 0) return i; } return -1; }
  function rowsFromSheet(sheet) { return XLSX.utils.sheet_to_json(sheet, { header: 1, defval: '', raw: false, blankrows: false }); }
  function findHostTable(rows) {
    for (let i = 0; i < Math.min(rows.length, 15); i++) {
      const h = rows[i].map(clean), host = findHeader(h, ['Host', 'ESX Host', 'VMHost', 'Host Name']), sockets = findHeader(h, ['# CPU', 'Num CPU Sockets', 'CPU Sockets', 'Sockets', 'Number of CPU Sockets']), cores = findHeader(h, ['Cores per CPU', 'Cores Per Socket', 'CPU Cores per Socket', 'Cores/CPU']);
      if (host >= 0 && sockets >= 0 && cores >= 0) return { data: rows.slice(i + 1), indexes: { host, sockets, cores, total: findHeader(h, ['# Cores', 'Total Cores', 'Cores']), cluster: findHeader(h, ['Cluster']), esxiVersion: findHeader(h, ['ESX Version', 'ESXi Version']) }, headerRow: i + 1 };
    }
    return null;
  }
  function parseHostRows(rows) {
    const table = findHostTable(rows);
    if (!table) throw new Error('Could not identify the host, # CPU (sockets), and Cores per CPU columns. Import an RVTools vHost worksheet/export.');
    const { indexes } = table, hosts = [], warnings = [];
    for (const row of table.data) {
      const name = clean(row[indexes.host]); if (!name) continue;
      const sockets = num(row[indexes.sockets]), coresPerSocket = num(row[indexes.cores]);
      if (!(sockets > 0) || !(coresPerSocket > 0)) { warnings.push(`${name}: missing or invalid socket/core count; excluded from calculations.`); continue; }
      const reportedTotal = indexes.total >= 0 ? num(row[indexes.total]) : NaN, actualCores = sockets * coresPerSocket;
      if (Number.isFinite(reportedTotal) && reportedTotal > 0 && reportedTotal !== actualCores) warnings.push(`${name}: RVTools total core count (${reportedTotal}) differs from sockets × cores/socket (${actualCores}); socket calculation used.`);
      const cluster = indexes.cluster >= 0 ? clean(row[indexes.cluster]) || '(No cluster)' : '(No cluster)';
      hosts.push({ name, cluster, sockets, coresPerSocket, esxiVersion: indexes.esxiVersion >= 0 ? clean(row[indexes.esxiVersion]) : '', actualCores, licensedCores: sockets * Math.max(16, coresPerSocket), minimumApplied: coresPerSocket < 16 });
    }
    if (!hosts.length) throw new Error('The vHost sheet was found, but no hosts with valid socket and core counts were present.');
    const seen = new Set(); for (const h of hosts) { const key = h.name.toLowerCase(); if (seen.has(key)) warnings.push(`Duplicate host row in this export: ${h.name}. Review the inventory before licensing.`); seen.add(key); }
    return { hosts, warnings, headerRow: table.headerRow };
  }
  async function parseFile(file) {
    if (!window.XLSX) throw new Error('Workbook reader did not load. Refresh and try again.');
    const wb = XLSX.read(await file.arrayBuffer(), { type: 'array', cellDates: false });
    let selected, sheetName = '', vcenterName = '', metadataServer = '', collectionAt = '', rvtoolsVersion = '', vcenterVersion = '';
    // RVTools' vMetaData.Server identifies the vCenter that produced the export.
    // Prefer it over the friendlier vInfo label so imports and scope filters use
    // the stable server identity when it is available.
    for (const name of wb.SheetNames) {
      if (normalize(name) !== 'vmetadata') continue;
      const metadataRows = rowsFromSheet(wb.Sheets[name]);
      for (let i = 0; i < Math.min(metadataRows.length, 10); i++) {
        const headers = metadataRows[i].map(clean), serverIndex = findHeader(headers, ['Server']), dateIndex = findHeader(headers, ['xlsx creation datetime', 'creation datetime']), versionIndex = findHeader(headers, ['RVTools version']);
        if (serverIndex >= 0 || dateIndex >= 0 || versionIndex >= 0) {
          const values = metadataRows[i + 1] || [];
          if (serverIndex >= 0) metadataServer = clean(values[serverIndex]);
          if (dateIndex >= 0) collectionAt = clean(values[dateIndex]);
          if (versionIndex >= 0) rvtoolsVersion = clean(values[versionIndex]);
          if (metadataServer || collectionAt || rvtoolsVersion) break;
        }
      }
      if (metadataServer) break;
    }
    for (const name of wb.SheetNames) {
      if (normalize(name) !== 'vsource') continue;
      const sourceRows = rowsFromSheet(wb.Sheets[name]);
      for (let i = 0; i < Math.min(sourceRows.length, 15); i++) {
        const headers = sourceRows[i].map(clean), versionIndex = findHeader(headers, ['Version']), apiVersionIndex = findHeader(headers, ['API version']), buildIndex = findHeader(headers, ['Build']);
        if (versionIndex >= 0 || apiVersionIndex >= 0) {
          const values = sourceRows[i + 1] || [], version = clean(values[versionIndex >= 0 ? versionIndex : apiVersionIndex]), build = buildIndex >= 0 ? clean(values[buildIndex]) : '';
          vcenterVersion = version ? `${version}${build ? ` (build ${build})` : ''}` : '';
          break;
        }
      }
      break;
    }
    for (const name of wb.SheetNames) {
      if (!['vinfo','info'].includes(normalize(name))) continue;
      const infoRows = rowsFromSheet(wb.Sheets[name]);
      for (let i = 0; i < Math.min(infoRows.length, 10); i++) {
        const headers = infoRows[i].map(clean), index = findHeader(headers, ['vCenter', 'vCenter Name', 'vCenter Server']);
        if (index >= 0) { vcenterName = clean(infoRows[i + 1]?.[index]); if (vcenterName) break; }
      }
      if (vcenterName) break;
    }
    if (metadataServer) vcenterName = metadataServer;
    for (const name of wb.SheetNames) if (['vhost','hosts','host'].includes(normalize(name))) { selected = wb.Sheets[name]; sheetName = name; break; }
    if (!selected) for (const name of wb.SheetNames) { if (findHostTable(rowsFromSheet(wb.Sheets[name]))) { selected = wb.Sheets[name]; sheetName = name; break; } }
    if (!selected) { selected = wb.Sheets[wb.SheetNames[0]]; sheetName = wb.SheetNames[0] || 'CSV'; }
    return { ...parseHostRows(rowsFromSheet(selected)), sheetName, vcenterName, collectionAt, rvtoolsVersion, vcenterVersion };
  }
  function initializeSelections() {
    const all = getAllHosts(state.session);
    state.selectedVcenters = new Set(all.map(h => h.vcenter));
    state.selectedClusters = new Set(all.map(clusterKey));
    state.selectedHosts = new Set(all.map(hostKey)); reconcileSelections();
  }
  const getAllHosts = p => (p?.imports || []).flatMap(i => i.hosts.map((h, index) => ({ ...h, importId: i.id, importFile: i.fileName, sheetName: i.sheetName, vcenter: i.vcenter, vcenterVersion: i.vcenterVersion || '', hostIndex: index, warnings: i.warnings || [] })));
  const hostKey = h => `${h.importId}:${h.hostIndex}`;
  const clusterKey = h => `${h.importId}:${h.cluster}`;
  function reconcileSelections() {
    const all = getAllHosts(state.session);
    const validV = new Set(all.map(h => h.vcenter)); state.selectedVcenters = new Set([...state.selectedVcenters].filter(x => validV.has(x)));
    const validC = new Set(all.map(clusterKey)); state.selectedClusters = new Set([...state.selectedClusters].filter(x => validC.has(x)));
    const validH = new Set(all.map(hostKey)); state.selectedHosts = new Set([...state.selectedHosts].filter(x => validH.has(x)));
    const inV = all.filter(h => state.selectedVcenters.has(h.vcenter)), inC = inV.filter(h => state.selectedClusters.has(clusterKey(h)));
    state.hosts = inC.filter(h => state.selectedHosts.has(hostKey(h)));
    const byCluster = new Map();
    for (const h of state.hosts) { const key = clusterKey(h); if (!byCluster.has(key)) byCluster.set(key, { key, name: h.cluster, vcenter: h.vcenter, hosts: [] }); byCluster.get(key).hosts.push(h); }
    state.clusters = [...byCluster.values()].sort((a,b) => `${a.vcenter} ${a.name}`.localeCompare(`${b.vcenter} ${b.name}`));
    state.warnings = [...new Set(state.hosts.flatMap(h => h.warnings))];
    const identities = new Map(); for (const h of state.hosts) { const key = `${h.vcenter.toLowerCase()}\u0000${h.name.toLowerCase()}`; identities.set(key,(identities.get(key)||0)+1); }
    for (const [key,count] of identities) if(count>1){const [vcenter,host]=key.split('\u0000');state.warnings.push(`${host} appears in ${count} imported rows for ${vcenter}; check that overlapping exports are not double-counted.`);}
  }
  async function persistScope() { state.session.scope = { vcenters:[...state.selectedVcenters], clusters:[...state.selectedClusters], hosts:[...state.selectedHosts] }; }
  function renderImports() {
    const imports=state.session.imports; $('importCount').textContent=`${imports.length} export${imports.length===1?'':'s'}`;
    $('importsRows').innerHTML=imports.length?imports.map(i=>`<tr><td><strong>${escapeHtml(i.vcenter)}</strong></td><td>${escapeHtml(i.vcenterVersion||'Not recorded in export')}</td><td>${escapeHtml(i.collectionAt||'Not recorded in export')}</td><td>${escapeHtml(i.rvtoolsVersion||'Not recorded in export')}</td><td>${fmt(i.hosts.length)}</td></tr>`).join(''):'<tr><td colspan="5" class="cluster-empty">No workbooks selected in this session.</td></tr>';
  }
  function renderFilters() {
    if (!state.session) return;
    const all = getAllHosts(state.session), vcenters = [...new Set(all.map(h => h.vcenter))].sort((a,b)=>a.localeCompare(b));
    const scoped = all.filter(h => state.selectedVcenters.has(h.vcenter));
    const clusters = new Map(); scoped.forEach(h => clusters.set(clusterKey(h), h.cluster));
    const inClusters = scoped.filter(h => state.selectedClusters.has(clusterKey(h)));
    renderChecks('vcenterFilters', vcenters.map(v => [v,v]), state.selectedVcenters, 'vcenter');
    renderChecks('clusterFilters', [...clusters.keys()].map(key => [key, scoped.find(h => clusterKey(h) === key).cluster]), state.selectedClusters, 'cluster');
    renderChecks('hostFilters', inClusters.map(h => [hostKey(h), h.name]), state.selectedHosts, 'host');
  }
  function renderChecks(id, options, selected, kind) {
    const el = $(id);
    el.innerHTML = options.length ? options.map(([key,label]) => `<label class="filter-option"><input type="checkbox" data-filter="${kind}" value="${escapeHtml(key)}" ${selected.has(key) ? 'checked' : ''}><span>${escapeHtml(label)}</span></label>`).join('') : '<span class="empty-filter">No items</span>';
    el.querySelectorAll('input').forEach(input => input.addEventListener('change', async () => {
      const target = input.dataset.filter === 'vcenter' ? state.selectedVcenters : input.dataset.filter === 'cluster' ? state.selectedClusters : state.selectedHosts;
      input.checked ? target.add(input.value) : target.delete(input.value);
      reconcileSelections(); await persistScope(); renderFilters(); render();
    }));
  }
  const totalCores = () => state.hosts.reduce((sum,h) => sum + h.licensedCores, 0);
  function pricingEstimate(product, cores, entitlement, rawTotal, complete) {
    const pricedCores = Math.max(cores, PRICEBOOK.minimumCoreOrder), coreSubtotal = pricedCores * product.annualPerCore;
    const addonTiB = complete ? Math.ceil(Math.max(0, rawTotal - entitlement)) : null;
    const vsanSubtotal = addonTiB === null ? null : addonTiB * PRICEBOOK.vsan.annualPerTiB;
    return { pricedCores, coreSubtotal, addonTiB, vsanSubtotal, annualTotal: vsanSubtotal === null ? null : coreSubtotal + vsanSubtotal };
  }
  function renderPricing(cores, rawTotal, complete) {
    for (const [prefix, product, entitlement] of [['vvf', PRICEBOOK.vvf, Math.ceil(cores * .25)], ['vcf', PRICEBOOK.vcf, cores]]) {
      const estimate = pricingEstimate(product, cores, entitlement, rawTotal, complete);
      $(`${prefix}UnitPrice`).textContent = `${usd(product.annualPerCore)} / core / year`;
      $(`${prefix}CoreSubtotal`).textContent = `${usd(estimate.coreSubtotal)} / year · ${fmt(estimate.pricedCores)} cores`;
      $(`${prefix}VsanSubtotal`).textContent = estimate.vsanSubtotal === null ? 'Enter raw TiB' : `${usd(estimate.vsanSubtotal)} / year · ${fmt(estimate.addonTiB)} TiB`;
      $(`${prefix}AnnualTotal`).textContent = estimate.annualTotal === null ? 'Enter raw TiB' : `${usd(estimate.annualTotal)} / year`;
    }
  }
  function setCapacity(cluster, field, value) {
    const cap = state.session.capacity || (state.session.capacity = {}), current = cap[cluster.key] || {};
    cap[cluster.key] = { ...current, [field]: value };
    if (!Number.isFinite(cap[cluster.key].raw)) delete cap[cluster.key].raw;

  }
  function render() {
    if (!state.session) return;
    const hosts = state.hosts, clusters = state.clusters, cores = totalCores(), sockets = hosts.reduce((s,h)=>s+h.sockets,0), actual = hosts.reduce((s,h)=>s+h.actualCores,0);
    $('results').classList.toggle('is-hidden', hosts.length === 0);
    if (!hosts.length) { $('hostCount').textContent='0'; $('socketCount').textContent='0'; $('actualCoreCount').textContent='0'; $('licensedCoreCount').textContent='0'; return; }
    $('hostCount').textContent=fmt(hosts.length); $('socketCount').textContent=fmt(sockets); $('actualCoreCount').textContent=fmt(actual); $('licensedCoreCount').textContent=fmt(cores);
    $('vvfCores').textContent=$('vcfCores').textContent=fmt(cores);
    const vvfEnt=Math.ceil(cores*.25), vcfEnt=cores; $('vvfEntitlement').textContent=`${fmt(vvfEnt)} TiB`; $('vcfEntitlement').textContent=`${fmt(vcfEnt)} TiB`;
    $('sourceLabel').textContent=`${new Set(hosts.map(h=>h.vcenter)).size} vCenter${new Set(hosts.map(h=>h.vcenter)).size===1?'':'s'} · ${hosts.length} selected hosts`;
    const capacities=clusters.map(c=>state.session.capacity?.[c.key]).filter(x=>x && Number.isFinite(x.raw)), complete=clusters.length>0 && capacities.length===clusters.length, rawTotal=capacities.reduce((s,x)=>s+x.raw,0);
    $('vvfAddon').textContent=complete?`${fmt(Math.ceil(Math.max(0,rawTotal-vvfEnt)))} TiB`:'Enter raw TiB'; $('vcfAddon').textContent=complete?`${fmt(Math.ceil(Math.max(0,rawTotal-vcfEnt)))} TiB`:'Enter raw TiB';
    renderPricing(cores, rawTotal, complete);
    $('totalRaw').textContent=capacities.length?`${fmt(rawTotal,3)} TiB${complete?'':' · partial'}`:'Not provided'; $('capacityStatus').textContent=complete?(capacities.every(x=>x.verified)?'PowerCLI verified':'Estimate / mixed'):'Awaiting capacity data';
    $('clusterRows').innerHTML=clusters.map((c,i)=>{ const item=state.session.capacity?.[c.key]||{}, cCores=c.hosts.reduce((s,h)=>s+h.licensedCores,0); return `<tr><td><strong>${escapeHtml(c.name)}</strong></td><td>${fmt(c.hosts.length)}</td><td>${fmt(cCores)}</td><td><input class="capacity-input" type="number" min="0" step="0.001" data-capacity="${i}" value="${Number.isFinite(item.raw)?item.raw:''}" placeholder="e.g. 24.5" aria-label="Raw capacity TiB"></td><td><select class="confidence-select" data-confidence="${i}" aria-label="Capacity source"><option value="estimate" ${!item.verified?'selected':''}>Estimate / unknown</option><option value="verified" ${item.verified?'selected':''}>PowerCLI verified</option></select></td></tr>`; }).join('');
    $('hostRows').innerHTML=hosts.map(h=>`<tr><td>${escapeHtml(h.vcenter)}</td><td><strong>${escapeHtml(h.name)}</strong></td><td>${escapeHtml(h.cluster)}</td><td>${fmt(h.sockets)}</td><td>${fmt(h.coresPerSocket)}</td><td>${fmt(h.actualCores)}</td><td><strong>${fmt(h.licensedCores)}</strong></td><td>${escapeHtml(h.esxiVersion || 'Not recorded in export')}</td><td>${h.minimumApplied?'16-core/socket minimum applied':'Actual physical cores'}</td></tr>`).join('');
    $('warnings').classList.toggle('is-hidden',state.warnings.length===0); $('warnings').innerHTML=state.warnings.length?`<b>Review inventory notes</b><ul>${state.warnings.map(w=>`<li>${escapeHtml(w)}</li>`).join('')}</ul>`:'';
    document.querySelectorAll('[data-capacity]').forEach(el=>el.addEventListener('input',()=>{const c=clusters[Number(el.dataset.capacity)],raw=el.value===''?NaN:Number(el.value),previous=state.session.capacity?.[c.key]||{};setCapacity(c,'raw',raw);if(!Number.isFinite(raw))delete state.session.capacity[c.key].raw;state.session.capacity[c.key].verified=previous.verified||false;refreshSummary();}));
    document.querySelectorAll('[data-confidence]').forEach(el=>el.addEventListener('change',()=>{const c=clusters[Number(el.dataset.confidence)];setCapacity(c,'verified',el.value==='verified');refreshSummary();}));
  }
  function refreshSummary() {
    if (!state.session) return;
    const caps=state.clusters.map(c=>state.session.capacity?.[c.key]).filter(x=>x&&Number.isFinite(x.raw));
    const complete=state.clusters.length>0&&caps.length===state.clusters.length, rawTotal=caps.reduce((s,x)=>s+x.raw,0), cores=totalCores();
    $('totalRaw').textContent=caps.length?`${fmt(rawTotal,3)} TiB${complete?'':' · partial'}`:'Not provided';
    $('capacityStatus').textContent=complete?(caps.every(x=>x.verified)?'PowerCLI verified':'Estimate / mixed'):'Awaiting capacity data';
    $('vvfAddon').textContent=complete?`${fmt(Math.ceil(Math.max(0,rawTotal-Math.ceil(cores*.25))))} TiB`:'Enter raw TiB';
    $('vcfAddon').textContent=complete?`${fmt(Math.ceil(Math.max(0,rawTotal-cores)))} TiB`:'Enter raw TiB';
    renderPricing(cores, rawTotal, complete);
  }
  function csvCell(v) { return `"${String(v ?? '').replace(/"/g,'""')}"`; }
  function download() {
    const cores=totalCores(), caps=state.clusters.map(c=>state.session.capacity?.[c.key]).filter(x=>x&&Number.isFinite(x.raw)), complete=caps.length===state.clusters.length, rawTotal=complete?caps.reduce((s,x)=>s+x.raw,0):null;
    const vvfPrice=pricingEstimate(PRICEBOOK.vvf,cores,Math.ceil(cores*.25),rawTotal,complete),vcfPrice=pricingEstimate(PRICEBOOK.vcf,cores,cores,rawTotal,complete);
    const lines=[['VMware Foundation Licensing Planner'],['vCenters',[...new Set(state.hosts.map(h=>h.vcenter))].join('; ')],['Selected hosts',state.hosts.length],['Physical sockets',state.hosts.reduce((s,h)=>s+h.sockets,0)],['Reported physical cores',state.hosts.reduce((s,h)=>s+h.actualCores,0)],['Licensable cores',cores],['VVF included vSAN TiB',Math.ceil(cores*.25)],['VCF included vSAN TiB',cores],['Raw vSAN TiB entered',rawTotal??'Incomplete'],['Capacity evidence',complete&&caps.every(x=>x.verified)?'PowerCLI verified':'Estimate / incomplete'],['Pricebook effective',PRICEBOOK.effective],['Price basis','Annual reseller subscription pricing (USD)'],['VVF SKU',PRICEBOOK.vvf.sku],['VVF annual price / core',PRICEBOOK.vvf.annualPerCore],['VVF priced core quantity',vvfPrice.pricedCores],['VVF annual core subtotal',vvfPrice.coreSubtotal],['VVF annual vSAN add-on estimate',vvfPrice.vsanSubtotal??'Capacity incomplete'],['VVF estimated annual total',vvfPrice.annualTotal??'Capacity incomplete'],['VCF SKU',PRICEBOOK.vcf.sku],['VCF annual price / core',PRICEBOOK.vcf.annualPerCore],['VCF priced core quantity',vcfPrice.pricedCores],['VCF annual core subtotal',vcfPrice.coreSubtotal],['VCF annual vSAN add-on estimate',vcfPrice.vsanSubtotal??'Capacity incomplete'],['VCF estimated annual total',vcfPrice.annualTotal??'Capacity incomplete'],['vSAN add-on SKU',PRICEBOOK.vsan.sku],['vSAN add-on annual price / TiB',PRICEBOOK.vsan.annualPerTiB],[],['vCenter','vCenter version','RVTools file','Host','ESXi version','Cluster','Sockets','Cores per socket','Actual cores','Licensable cores','Rule']];
    for(const h of state.hosts) lines.push([h.vcenter,h.vcenterVersion,h.importFile,h.name,h.esxiVersion,h.cluster,h.sockets,h.coresPerSocket,h.actualCores,h.licensedCores,h.minimumApplied?'16-core minimum per socket':'Actual physical cores']);
    lines.push([],['vCenter','Cluster','Hosts','Licensable cores','Raw vSAN TiB','Evidence']); for(const c of state.clusters){const x=state.session.capacity?.[c.key]||{};lines.push([c.vcenter,c.name,c.hosts.length,c.hosts.reduce((s,h)=>s+h.licensedCores,0),Number.isFinite(x.raw)?x.raw:'',x.verified?'PowerCLI verified':'Estimate / unknown']);}
    const blob=new Blob(['\ufeff'+lines.map(row=>row.map(csvCell).join(',')).join('\r\n')],{type:'text/csv;charset=utf-8'}),url=URL.createObjectURL(blob),a=document.createElement('a');a.href=url;a.download='vmware-licensing-estimate.csv';a.click();URL.revokeObjectURL(url);
  }
  async function importFiles(files) {
    const selected=[...files]; if(!selected.length)return; const explicit=clean($('vcenterName').value);
    if(selected.length>1&&explicit){setStatus('For multiple files, leave the vCenter override blank; names are read from RVTools metadata (or the filename when metadata is unavailable).',true);return;}
    setStatus(`Parsing ${selected.length} export${selected.length===1?'':'s'} locally…`); const imported=[];
    for(const file of selected){try{
      const parsed=await parseFile(file), fallback=file.name.replace(/\.[^.]+$/,'')||'vCenter';
      const record={id:uid(),vcenter:parsed.vcenterName||(selected.length===1&&explicit)||fallback,fileName:file.name,importedAt:new Date().toISOString(),collectionAt:parsed.collectionAt,rvtoolsVersion:parsed.rvtoolsVersion,vcenterVersion:parsed.vcenterVersion,hosts:parsed.hosts,sheetName:parsed.sheetName,headerRow:parsed.headerRow,warnings:parsed.warnings};
      state.session.imports.push(record); imported.push(record);
    }catch(e){setStatus(`${file.name}: ${e.message}`,true);}}
    if(imported.length){state.session.scope=null;for(const i of imported){i.hosts.forEach((h,index)=>{state.selectedVcenters.add(i.vcenter);state.selectedClusters.add(`${i.id}:${h.cluster}`);state.selectedHosts.add(`${i.id}:${index}`);});}
      reconcileSelections();renderImports();renderFilters();render();$('vcenterName').value='';$('rvFile').value='';setStatus(`Read ${imported.length} workbook${imported.length===1?'':'s'} locally. Files and calculations remain in this tab and are not saved.`);}
  }
  $('rvFile').addEventListener('change',()=>importFiles($('rvFile').files).catch(e=>setStatus(e.message,true)));
  $('clearSession').addEventListener('click',()=>{state.session.imports=[];state.session.capacity={};state.session.scope=null;state.selectedVcenters.clear();state.selectedClusters.clear();state.selectedHosts.clear();reconcileSelections();renderImports();renderFilters();render();$('rvFile').value='';$('vcenterName').value='';setStatus('Session cleared. Nothing was saved.');});
  $('selectAllFilters').addEventListener('click',async()=>{const all=getAllHosts(state.session);state.selectedVcenters=new Set(all.map(h=>h.vcenter));state.selectedClusters=new Set(all.map(clusterKey));state.selectedHosts=new Set(all.map(hostKey));reconcileSelections();await persistScope();renderFilters();render();});
  $('downloadReport').addEventListener('click',download);
  initializeSelections();renderImports();renderFilters();render();setStatus('Choose RVTools workbooks to calculate licensing locally. Nothing is uploaded or saved.');
})();
