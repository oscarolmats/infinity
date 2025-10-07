(function(){
  const fileInput = document.getElementById('fileInput');
  const filterInput = document.getElementById('filterInput');
  const toggleAllBtn = document.getElementById('toggleAllBtn');
  const groupBySelect = document.getElementById('groupBy');
  let lastRows = null; // cache of parsed rows for re-rendering
  // Layer modal refs
  const layerModal = document.getElementById('layerModal');
  const layerCountInput = document.getElementById('layerCount');
  const layerThicknessesInput = document.getElementById('layerThicknesses');
  const layerCancelBtn = document.getElementById('layerCancel');
  const layerApplyBtn = document.getElementById('layerApply');
  let layerTarget = null; // { type: 'row'|'group', key?: string, rowEl?: HTMLTableRowElement }
  const output = document.getElementById('output');

  function parseDelimited(text){
    const lines = text.replace(/\r\n?/g, '\n').split('\n').filter(l => l.length > 0);
    if(lines.length === 0) return [];
    const sample = lines.slice(0, 10).join('\n');
    const comma = (sample.match(/,/g) || []).length;
    const semicolon = (sample.match(/;/g) || []).length;
    const tab = (sample.match(/\t/g) || []).length;
    const delimiter = comma >= semicolon && comma >= tab ? ',' : (semicolon >= tab ? ';' : '\t');
    return lines.map(line => line.split(delimiter).map(cell => cell.trim()));
  }

  async function uploadExcel(file){
    const form = new FormData();
    form.append('file', file);
    const res = await fetch('/upload', { method: 'POST', body: form });
    if(!res.ok){ throw new Error('Uppladdning misslyckades'); }
    return await res.text();
  }

  function getTable(){ return output.querySelector('table'); }

  function setAllGroups(open){
    const table = getTable(); if(!table) return;
    const parents = Array.from(table.querySelectorAll('tbody tr.group-parent'));
    parents.forEach(function(parent){
      parent.setAttribute('data-open', String(open));
    });
    const children = Array.from(table.querySelectorAll('tbody tr[data-group-child-of]'));
    children.forEach(function(ch){ ch.style.display = open ? '' : 'none'; });
  }

  function ensureColumnFilters(){
    const table = getTable(); if(!table) return [];
    const thead = table.querySelector('thead') || table.createTHead();
    const headerRow = thead.querySelector('tr'); if(!headerRow) return [];
    const colCount = headerRow.children.length;
    let filterRow = thead.querySelector('tr[data-filter-row="true"]');
    if(!filterRow){
      filterRow = document.createElement('tr');
      filterRow.setAttribute('data-filter-row', 'true');
      for(let i=0;i<colCount;i++){
        const th = document.createElement('th');
        const input = document.createElement('input');
        input.type = 'search'; input.placeholder = 'Filter...'; input.style.width = '100%';
        input.dataset.colIndex = String(i);
        th.appendChild(input); filterRow.appendChild(th);
      }
      thead.appendChild(filterRow);
    }
    return Array.from(filterRow.querySelectorAll('input'));
  }

  function applyFilters(){
    const table = getTable(); if(!table) return;
    const globalQ = (filterInput && filterInput.value || '').toLowerCase().trim();
    const colInputs = ensureColumnFilters();
    const colQueries = colInputs.map(inp => (inp.value || '').toLowerCase().trim());
    const rows = Array.from(table.querySelectorAll('tbody tr'));

    const parents = rows.filter(r => r.hasAttribute('data-group-key'));
    const childrenByGroup = new Map();
    rows.forEach(r => {
      const of = r.getAttribute('data-group-child-of');
      if(of){ if(!childrenByGroup.has(of)) childrenByGroup.set(of, []); childrenByGroup.get(of).push(r); }
    });

    function rowMatches(tr){
      const cells = Array.from(tr.children);
      const text = tr.textContent.toLowerCase();
      const globalOk = !globalQ || text.includes(globalQ);
      const colsOk = colQueries.every((q, idx) => {
        if(!q) return true; const td = cells[idx]; const cellText = (td ? td.textContent : '').toLowerCase(); return cellText.includes(q);
      });
      return globalOk && colsOk;
    }

    rows.forEach(tr => { tr.style.display = ''; });

    parents.forEach(parent => {
      const key = parent.getAttribute('data-group-key');
      const kids = childrenByGroup.get(key) || [];
      const parentMatch = rowMatches(parent);
      const anyChildMatch = kids.some(rowMatches);
      const showParent = parentMatch || anyChildMatch;
      parent.style.display = showParent ? '' : 'none';
      kids.forEach(k => { k.style.display = showParent && rowMatches(k) ? '' : 'none'; });
    });

    rows.filter(r => !r.hasAttribute('data-group-key') && !r.hasAttribute('data-group-child-of'))
        .forEach(tr => { tr.style.display = rowMatches(tr) ? '' : 'none'; });
  }

  function parseNumberLike(v){
    if(v == null) return NaN;
    const s = String(v).trim().replace(/\s+/g, '').replace(',', '.');
    const n = parseFloat(s);
    return Number.isFinite(n) ? n : NaN;
  }

  function buildGroupedTable(headers, bodyRows, groupColIndex){
    const table = document.createElement('table');
    const thead = document.createElement('thead');
    const headerTr = document.createElement('tr');
    // Add an empty header for actions (layer split)
    const actionTh = document.createElement('th'); actionTh.textContent = '';
    headerTr.appendChild(actionTh);
    headers.forEach(h => { const th = document.createElement('th'); th.textContent = h; headerTr.appendChild(th); });
    thead.appendChild(headerTr); table.appendChild(thead);
    const tbody = document.createElement('tbody');

    const idxType = groupColIndex;
    const idxNetArea = headers.findIndex(h => String(h).toLowerCase() === 'net area');
    const idxVolume = headers.findIndex(h => String(h).toLowerCase() === 'volume');
    const idxCount = headers.findIndex(h => String(h).toLowerCase() === 'count');

    const groups = new Map();
    bodyRows.forEach(r => {
      const key = r[idxType] || '';
      if(!groups.has(key)) groups.set(key, []);
      groups.get(key).push(r);
    });

    groups.forEach((rows, key) => {
      let sumNetArea = 0, sumVolume = 0, sumCount = 0;
      const hasNetArea = idxNetArea !== -1;
      const hasVolume = idxVolume !== -1;
      const hasCount = idxCount !== -1;
      rows.forEach(r => {
        if(hasNetArea){ const n = parseNumberLike(r[idxNetArea]); if(Number.isFinite(n)) sumNetArea += n; }
        if(hasVolume){ const n = parseNumberLike(r[idxVolume]); if(Number.isFinite(n)) sumVolume += n; }
        if(hasCount){ const n = parseNumberLike(r[idxCount]); if(Number.isFinite(n)) sumCount += n; }
      });

      const parentTr = document.createElement('tr');
      parentTr.className = 'group-parent';
      parentTr.setAttribute('data-group-key', String(key));
      // Create one cell per column so sums align under headers
      // Parent action cell (group layer)
      const actionTd = document.createElement('td');
      const groupBtn = document.createElement('button'); groupBtn.type = 'button'; groupBtn.textContent = 'Skikta grupp';
      groupBtn.addEventListener('click', function(ev){ ev.stopPropagation(); openLayerModal({ type: 'group', key: String(key) }); });
      actionTd.appendChild(groupBtn); parentTr.appendChild(actionTd);
      for(let i = 0; i < headers.length; i++){
        const td = document.createElement('td');
        if(i === idxType){
          const toggle = document.createElement('span'); toggle.className = 'group-toggle';
          toggle.innerHTML = '<svg viewBox="0 0 24 24" aria-hidden="true"><path fill="currentColor" d="M8.59 16.59L13.17 12 8.59 7.41 10 6l6 6-6 6z"/></svg>';
          td.appendChild(toggle);
          const label = document.createElement('span');
          label.textContent = (key || '(tom)') + ' (' + rows.length + ')';
          td.appendChild(label);
        } else if(hasNetArea && i === idxNetArea){
          td.textContent = String(sumNetArea);
        } else if(hasVolume && i === idxVolume){
          td.textContent = String(sumVolume);
        } else if(hasCount && i === idxCount){
          td.textContent = String(sumCount);
        } else {
          td.textContent = '';
        }
        parentTr.appendChild(td);
      }
      tbody.appendChild(parentTr);

      rows.forEach(r => {
        const tr = document.createElement('tr'); tr.setAttribute('data-group-child-of', String(key));
        // Row action cell
        const actionTd = document.createElement('td');
        const rowBtn = document.createElement('button'); rowBtn.type = 'button'; rowBtn.textContent = 'Skikta';
        rowBtn.addEventListener('click', function(ev){ ev.stopPropagation(); openLayerModal({ type: 'row', rowEl: tr }); });
        actionTd.appendChild(rowBtn); tr.appendChild(actionTd);
        r.forEach(c => { const td = document.createElement('td'); td.textContent = c; tr.appendChild(td); });
        tbody.appendChild(tr);
      });
    });

    table.appendChild(tbody);
    return table;
  }

  function populateGroupBy(headers){
    if(!groupBySelect) return;
    const previous = groupBySelect.value;
    groupBySelect.innerHTML = '';
    const noneOpt = document.createElement('option'); noneOpt.value = ''; noneOpt.textContent = '(ingen)';
    groupBySelect.appendChild(noneOpt);
    headers.forEach((h, idx) => {
      const opt = document.createElement('option');
      opt.value = String(idx);
      opt.textContent = h;
      groupBySelect.appendChild(opt);
    });
    // Preserve previous selection if still valid; otherwise default to no grouping
    const hasPrev = Array.from(groupBySelect.options).some(o => o.value === previous);
    if(previous && hasPrev){
      groupBySelect.value = previous;
    } else {
      // Default to no grouping (empty value)
      groupBySelect.value = '';
    }
  }

  function renderTableWithOptionalGrouping(rows){
    if(!rows || rows.length === 0){ output.innerHTML = '<div>Ingen data att visa.</div>'; return; }
    const headers = rows[0];
    const bodyRows = rows.slice(1);
    if(groupBySelect){ populateGroupBy(headers); }
    const selected = groupBySelect ? groupBySelect.value : '';
    const groupIdx = selected === '' ? -1 : parseInt(selected, 10);

    if(groupIdx === -1 || Number.isNaN(groupIdx)){
      const table = document.createElement('table');
      const thead = document.createElement('thead');
      const headerTr = document.createElement('tr');
      const actionTh = document.createElement('th'); actionTh.textContent = '';
      headerTr.appendChild(actionTh);
      headers.forEach(h => { const th = document.createElement('th'); th.textContent = h; headerTr.appendChild(th); });
      thead.appendChild(headerTr); table.appendChild(thead);
      const tbody = document.createElement('tbody');
      bodyRows.forEach(r => {
        const tr = document.createElement('tr');
        const actionTd = document.createElement('td');
        const rowBtn = document.createElement('button'); rowBtn.type = 'button'; rowBtn.textContent = 'Skikta';
        rowBtn.addEventListener('click', function(ev){ ev.stopPropagation(); openLayerModal({ type: 'row', rowEl: tr }); });
        actionTd.appendChild(rowBtn); tr.appendChild(actionTd);
        r.forEach(c => { const td = document.createElement('td'); td.textContent = c; tr.appendChild(td); });
        tbody.appendChild(tr);
      });
      table.appendChild(tbody);
      output.innerHTML = ''; output.appendChild(table);
    } else {
      const table = buildGroupedTable(headers, bodyRows, groupIdx);
      output.innerHTML = ''; output.appendChild(table);
      output.addEventListener('click', function(e){
        const tr = e.target && e.target.closest && e.target.closest('tr.group-parent');
        if(!tr) return;
        const key = tr.getAttribute('data-group-key');
        const isOpen = tr.getAttribute('data-open') !== 'false';
        const nextOpen = !isOpen;
        tr.setAttribute('data-open', String(nextOpen));
        // svg rotation handled via CSS data-open selector
        const table = getTable(); if(!table) return;
        const children = table.querySelectorAll(`tbody tr[data-group-child-of="${CSS.escape(key)}"]`);
        children.forEach(ch => { ch.style.display = nextOpen ? '' : 'none'; });
      });
    }
    ensureColumnFilters();
    applyFilters();
  }

  function handleFile(file){
    const ext = (file.name.split('.').pop() || '').toLowerCase();
    if(ext === 'xlsx'){
      output.textContent = 'Laser Excel...';
      uploadExcel(file)
        .then(html => { output.innerHTML = html;
          const table = getTable(); if(!table){ return; }
          const headers = Array.from(table.querySelectorAll('thead th')).map(th => th.textContent);
          const bodyRows = Array.from(table.querySelectorAll('tbody tr')).map(tr => Array.from(tr.children).map(td => td.textContent));
          lastRows = [headers, ...bodyRows];
          renderTableWithOptionalGrouping(lastRows);
        })
        .catch(() => { output.textContent = 'Kunde inte lasa Excel-filen.'; });
      return;
    }
    const reader = new FileReader();
    reader.onload = function(e){
      const text = e.target.result;
      const rows = parseDelimited(text);
      lastRows = rows;
      renderTableWithOptionalGrouping(lastRows);
    };
    reader.onerror = function(){ output.textContent = 'Fel vid lasning av filen.'; };
    reader.readAsText(file, 'utf-8');
  }

  fileInput.addEventListener('change', function(){ const file = this.files && this.files[0]; if(!file) return; handleFile(file); });
  if(filterInput){ filterInput.addEventListener('input', applyFilters); }
  if(toggleAllBtn){
    toggleAllBtn.addEventListener('click', function(){
      const table = getTable();
      if(!table){ return; }
      // Determine current overall state: if any group appears open, collapse all; else expand all
      const anyOpen = Array.from(table.querySelectorAll('tbody tr.group-parent'))
        .some(tr => tr.getAttribute('data-open') !== 'false');
      const nextOpen = !anyOpen;
      setAllGroups(nextOpen);
      toggleAllBtn.textContent = nextOpen ? 'Fäll ihop alla' : 'Fäll ut alla';
    });
  }
  if(groupBySelect){
    groupBySelect.addEventListener('change', function(){
      if(!lastRows){ return; }
      renderTableWithOptionalGrouping(lastRows);
    });
  }

  // Layer modal behavior
  function openLayerModal(target){
    layerTarget = target;
    if(layerModal){ layerModal.style.display = 'flex'; }
  }
  function closeLayerModal(){
    layerTarget = null;
    if(layerModal){ layerModal.style.display = 'none'; }
  }
  if(layerCancelBtn){ layerCancelBtn.addEventListener('click', closeLayerModal); }
  if(layerModal){ layerModal.addEventListener('click', function(e){ if(e.target === layerModal) closeLayerModal(); }); }
  if(layerApplyBtn){
    layerApplyBtn.addEventListener('click', function(){
      const count = Math.max(1, parseInt(layerCountInput && layerCountInput.value || '1', 10));
      const raw = (layerThicknessesInput && layerThicknessesInput.value || '').trim();
      const thicknesses = raw ? raw.split(',').map(s => parseFloat(s.trim().replace(',', '.'))).filter(n => Number.isFinite(n) && n > 0) : [];
      applyLayerSplit(count, thicknesses);
      closeLayerModal();
    });
  }

  function applyLayerSplit(count, thicknesses){
    if(!layerTarget){ return; }
    const table = getTable(); if(!table) return;
    const tbody = table.querySelector('tbody'); if(!tbody) return;

    function cloneRowWithMultiplier(srcTr, multiplier, layerIndex, totalLayers){
      const clone = srcTr.cloneNode(true);
      // Do not carry action button listeners
      const btn = clone.querySelector('button'); if(btn){ btn.replaceWith(btn.cloneNode(true)); }
      clone.classList.add('is-new');
      // Try to scale numeric cells for Net Area, Volume, Count
      const headerTexts = Array.from(table.querySelectorAll('thead th')).map(th => th.textContent);
      const idxNetArea = headerTexts.findIndex(h => String(h).toLowerCase() === 'net area');
      const idxVolume = headerTexts.findIndex(h => String(h).toLowerCase() === 'volume');
      const idxCount = headerTexts.findIndex(h => String(h).toLowerCase() === 'count');
      const tds = Array.from(clone.children);
      // Add badge into first data cell (after action column)
      const firstDataTd = tds[1];
      if(firstDataTd){
        const badge = document.createElement('span'); badge.className = 'badge-new'; badge.textContent = 'Skikt ' + (layerIndex + 1) + '/' + totalLayers;
        firstDataTd.insertBefore(badge, firstDataTd.firstChild);
      }
      function scaleCell(idx){
        if(idx < 0) return;
        const td = tds[idx + 1] || null; // +1 offset for action column
        if(!td) return;
        const n = parseNumberLike(td.textContent);
        if(Number.isFinite(n)){ td.textContent = String(n * multiplier); }
      }
      scaleCell(idxNetArea); scaleCell(idxVolume); scaleCell(idxCount);
      return clone;
    }

    function splitRow(tr){
      // Even split if no thicknesses provided
      const multipliers = thicknesses.length > 0
        ? thicknesses.map(t => t / thicknesses.reduce((a,b)=>a+b,0))
        : Array(count).fill(1 / count);
      const fragments = multipliers.map((m, i) => cloneRowWithMultiplier(tr, m, i, multipliers.length));
      fragments.forEach(f => tbody.insertBefore(f, tr.nextSibling));
      tr.remove();
    }

    if(layerTarget.type === 'row' && layerTarget.rowEl){
      splitRow(layerTarget.rowEl);
    } else if(layerTarget.type === 'group' && layerTarget.key != null){
      const rows = Array.from(tbody.querySelectorAll('tr[data-group-child-of="' + CSS.escape(layerTarget.key) + '"]'));
      rows.forEach(splitRow);
    }
    // Re-apply filters to keep visibility consistent
    applyFilters();
  }
  output.addEventListener('input', function(e){ const t = e.target; if(t && t.closest && t.closest('thead') && t.tagName === 'INPUT'){ applyFilters(); } });
})();
