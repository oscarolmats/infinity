(function(){
  const fileInput = document.getElementById('fileInput');
  const filterInput = document.getElementById('filterInput');
  const toggleAllBtn = document.getElementById('toggleAllBtn');
  const groupBySelect = document.getElementById('groupBy');
  let lastRows = null; // cache of parsed rows for re-rendering
  
  // Storage for layers and climate resources
  let layerData = new Map(); // key: row signature -> { count, thicknesses, layerKey }
  let climateData = new Map(); // key: row signature or layerKey -> resourceName
  
  // Layer modal refs
  const layerModal = document.getElementById('layerModal');
  const layerCountInput = document.getElementById('layerCount');
  const layerThicknessesInput = document.getElementById('layerThicknesses');
  const layerCancelBtn = document.getElementById('layerCancel');
  const layerApplyBtn = document.getElementById('layerApply');
  let layerTarget = null; // { type: 'row'|'group', key?: string, rowEl?: HTMLTableRowElement }
  
  // Climate resource modal refs
  const climateModal = document.getElementById('climateModal');
  const climateResourceSelect = document.getElementById('climateResourceSelect');
  const climateCancelBtn = document.getElementById('climateCancel');
  const climateApplyBtn = document.getElementById('climateApply');
  let climateTarget = null; // { type: 'row'|'group', key?: string, rowEl?: HTMLTableRowElement }
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

  // Generate unique signature for a row based on its original data and layer position
  // rowData should be the original array data, not the DOM elements
  function getRowSignature(rowData, layerChildOf = null){
    const baseSignature = JSON.stringify(rowData);
    // Include layer hierarchy in signature to differentiate nested layers
    return layerChildOf ? `${baseSignature}::layer::${layerChildOf}` : baseSignature;
  }
  
  // Get row data from a TR element, excluding action column and climate column
  function getRowDataFromTr(tr){
    const table = tr.closest('table');
    if(!table) return null;
    const headers = Array.from(table.querySelectorAll('thead tr:first-child th')).map(th => th.textContent);
    const climateColIndex = headers.findIndex(h => h === 'Klimatresurs');
    
    const cells = Array.from(tr.children);
    const rowData = [];
    
    // Skip first cell (action column) and climate column if it exists
    for(let i = 1; i < cells.length; i++){
      if(climateColIndex !== -1 && i === climateColIndex){
        continue; // Skip climate column
      }
      
      // Clone the cell and remove badges and decorations
      const tempDiv = document.createElement('div');
      tempDiv.innerHTML = cells[i].innerHTML;
      
      // Remove badges
      tempDiv.querySelectorAll('.badge-new').forEach(b => b.remove());
      // Remove toggle icons
      tempDiv.querySelectorAll('.group-toggle').forEach(t => t.remove());
      
      // Get clean text and remove layer indicators
      let text = tempDiv.textContent.trim();
      text = text.replace(/^\[Skikt \d+\/\d+\]\s*/, ''); // Remove "Skikt X/Y" prefix
      text = text.replace(/\s*\[\d+\s+skikt\]\s*$/, ''); // Remove "[N skikt]" suffix
      text = text.replace(/\s*\(\d+\)\s*$/, ''); // Remove "(N)" count suffix from groups
      
      rowData.push(text.trim());
    }
    
    return rowData;
  }

  // Apply saved layers to a row
  function applySavedLayers(tr, rowData){
    if(!rowData || !Array.isArray(rowData)){
      return;
    }
    
    const layerChildOf = tr.getAttribute('data-layer-child-of');
    const signature = getRowSignature(rowData, layerChildOf);
    const saved = layerData.get(signature);
    
    console.log('🔍 LOOKUP:', rowData[1]?.substring(0,10), '- LayerChild:', layerChildOf?.substring(0,10) || 'none', '- Found:', !!saved, '- Signature:', signature.substring(0, 60));
    
    if(saved){
      console.log('✅ RESTORE:', rowData[1]?.substring(0,10), '- applying', saved.count, 'layers');
      // Trigger layer split with saved parameters and layerKey
      layerTarget = { type: 'row', rowEl: tr };
      const tempCount = saved.count;
      const tempThicknesses = saved.thicknesses;
      const tempLayerKey = saved.layerKey;
      
      const table = tr.closest('table');
      const tbody = table ? table.querySelector('tbody') : null;
      if(tbody){
        applyLayerSplitWithKey(tr, tbody, tempCount, tempThicknesses, tempLayerKey);
      }
      layerTarget = null;
    }
  }
  
  // Helper function for applying layers with a specific key
  function applyLayerSplitWithKey(tr, tbody, count, thicknesses, layerKey, isNested = false){
    const table = tr.closest('table');
    
    function splitRowWithKey(tr, savedLayerKey){
      const multipliers = thicknesses.length > 0
        ? thicknesses.map(t => t / thicknesses.reduce((a,b)=>a+b,0))
        : Array(count).fill(1 / count);
      
      // Convert original row to parent row
      tr.classList.remove('is-new');
      tr.classList.add('layer-parent');
      // Don't add 'group-parent' class, only layer-parent
      tr.setAttribute('data-layer-key', savedLayerKey);
      tr.setAttribute('data-open', 'true');
      
      // Update action buttons on parent row
      const actionTd = tr.querySelector('td:first-child');
      if(actionTd){
        actionTd.innerHTML = '';
        const parentLayerBtn = document.createElement('button');
        parentLayerBtn.type = 'button';
        parentLayerBtn.textContent = 'Skikta skikt';
        parentLayerBtn.addEventListener('click', function(ev){
          ev.stopPropagation();
          openLayerModal({ type: 'group', key: savedLayerKey });
        });
        actionTd.appendChild(parentLayerBtn);
        
        const parentClimateBtn = document.createElement('button');
        parentClimateBtn.type = 'button';
        parentClimateBtn.textContent = 'Mappa klimatresurs';
        parentClimateBtn.addEventListener('click', function(ev){
          ev.stopPropagation();
          openClimateModal({ type: 'group', key: savedLayerKey });
        });
        actionTd.appendChild(parentClimateBtn);
      }
      
      // Add toggle to first data cell
      const firstDataTd = tr.querySelector('td:nth-child(2)');
      if(firstDataTd){
        const toggle = document.createElement('span');
        toggle.className = 'group-toggle';
        toggle.innerHTML = '<svg viewBox="0 0 24 24" aria-hidden="true"><path fill="currentColor" d="M8.59 16.59L13.17 12 8.59 7.41 10 6l6 6-6 6z"/></svg>';
        firstDataTd.insertBefore(toggle, firstDataTd.firstChild);
        
        const layerLabel = document.createElement('span');
        layerLabel.textContent = ' [' + multipliers.length + ' skikt]';
        layerLabel.style.marginLeft = '4px';
        firstDataTd.appendChild(layerLabel);
      }
      
      // Preserve parent's group membership for children
      const parentGroupKey = tr.getAttribute('data-group-child-of');
      
      // Preserve original row data for children
      const originalRowData = tr._originalRowData;
      
      // Create child layer rows
      const fragments = multipliers.map((m, i) => {
        const clone = tr.cloneNode(true);
        // Preserve original row data
        if(originalRowData){
          clone._originalRowData = originalRowData;
        }
        // Replace action buttons with new ones (without old listeners)
        const actionTd = clone.querySelector('td:first-child');
        if(actionTd){
          actionTd.innerHTML = ''; // Clear old buttons
          
          // Add "Skikta" button
          const layerBtn = document.createElement('button');
          layerBtn.type = 'button';
          layerBtn.textContent = 'Skikta';
          layerBtn.addEventListener('click', function(ev){ 
            ev.stopPropagation(); 
            openLayerModal({ type: 'row', rowEl: clone }); 
          });
          actionTd.appendChild(layerBtn);
          
          // Add "Mappa klimatresurs" button
          const climateBtn = document.createElement('button');
          climateBtn.type = 'button';
          climateBtn.textContent = 'Mappa klimatresurs';
          climateBtn.addEventListener('click', function(ev){ 
            ev.stopPropagation(); 
            openClimateModal({ type: 'row', rowEl: clone }); 
          });
          actionTd.appendChild(climateBtn);
        }
        
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
          const badge = document.createElement('span'); 
          badge.className = 'badge-new'; 
          badge.textContent = 'Skikt ' + (i + 1) + '/' + multipliers.length;
          firstDataTd.insertBefore(badge, firstDataTd.firstChild);
        }
        function scaleCell(idx){
          if(idx < 0) return;
          const td = tds[idx + 1] || null; // +1 offset for action column
          if(!td) return;
          const n = parseNumberLike(td.textContent);
          if(Number.isFinite(n)){ td.textContent = String(n * m); }
        }
        scaleCell(idxNetArea); scaleCell(idxVolume); scaleCell(idxCount);
        
        // Mark as child of this layer
        clone.setAttribute('data-layer-child-of', savedLayerKey);
        // Also inherit parent's group membership if it exists
        if(parentGroupKey){
          clone.setAttribute('data-group-child-of', parentGroupKey);
        }
        // Set immediate parent for toggle
        clone.setAttribute('data-parent-key', savedLayerKey);
        return clone;
      });
      
      // Insert layer children right after the parent
      let insertAfter = tr;
      fragments.forEach(f => {
        tbody.insertBefore(f, insertAfter.nextSibling);
        insertAfter = f;
        
        // Apply saved climate to this child row using original data
        const childRowData = f._originalRowData;
        if(childRowData){
          applySavedClimate(f, childRowData);
          
          // Only check for nested layers if this is NOT already a nested restoration
          // (to prevent infinite recursion when restoring multi-level hierarchies)
          if(!isNested){
            const childLayerChildOf = f.getAttribute('data-layer-child-of');
            const childSignature = getRowSignature(childRowData, childLayerChildOf);
            const childSaved = layerData.get(childSignature);
            if(childSaved){
              setTimeout(() => {
                applyLayerSplitWithKey(f, tbody, childSaved.count, childSaved.thicknesses, childSaved.layerKey, true);
              }, 0);
            }
          }
        }
      });
    }
    
    splitRowWithKey(tr, layerKey);
  }

  // Apply saved climate resource to a row
  function applySavedClimate(tr, rowData){
    const layerChildOf = tr.getAttribute('data-layer-child-of');
    const signature = getRowSignature(rowData, layerChildOf);
    const resourceName = climateData.get(signature);
    if(resourceName){
      const table = getTable(); if(!table) return;
      const thead = table.querySelector('thead'); if(!thead) return;
      
      const headerRow = thead.querySelector('tr');
      const existingClimateHeader = Array.from(headerRow.children).find(th => th.textContent === 'Klimatresurs');
      
      if(!existingClimateHeader){
        const climateTh = document.createElement('th');
        climateTh.textContent = 'Klimatresurs';
        headerRow.appendChild(climateTh);
      }
      
      const existingClimateCell = tr.querySelector('td[data-climate-cell="true"]');
      if(existingClimateCell){
        existingClimateCell.textContent = resourceName;
      } else {
        const climateTd = document.createElement('td');
        climateTd.textContent = resourceName;
        climateTd.setAttribute('data-climate-cell', 'true');
        tr.appendChild(climateTd);
      }
    }
  }

  function toggleDescendants(parentTr, show, visited = new Set()){
    const table = getTable(); if(!table) return;
    const key = parentTr.getAttribute('data-layer-key');
    if(!key) return;
    
    // Prevent infinite recursion
    if(visited.has(key)) return;
    visited.add(key);
    
    const children = table.querySelectorAll(`tbody tr[data-parent-key="${CSS.escape(key)}"]`);
    children.forEach(ch => {
      ch.style.display = show ? '' : 'none';
      if(ch.classList.contains('layer-parent')){
        ch.setAttribute('data-open', String(show));
        toggleDescendants(ch, show, visited);
      }
    });
  }
  
  // Unified toggle function for all parent types
  function toggleParentRow(parentTr){
    if(!parentTr) return;
    
    const hasGroupParentClass = parentTr.classList.contains('group-parent');
    const hasLayerParentClass = parentTr.classList.contains('layer-parent');
    const hasGroupKey = parentTr.hasAttribute('data-group-key');
    const hasLayerKey = parentTr.hasAttribute('data-layer-key');
    
    // Determine the type based on what keys are present
    // Priority: if it has a layer-key, it's primarily a layer parent
    const isLayerParent = hasLayerKey;
    const isGroupParent = hasGroupKey && !hasLayerKey;
    
    if(!isGroupParent && !isLayerParent) return;
    
    const isOpen = parentTr.getAttribute('data-open') !== 'false';
    const nextOpen = !isOpen;
    parentTr.setAttribute('data-open', String(nextOpen));
    
    const table = getTable(); if(!table) return;
    
    if(isGroupParent){
      // For pure group parents (no layer-key), toggle all direct group children
      const key = parentTr.getAttribute('data-group-key');
      const children = table.querySelectorAll(`tbody tr[data-group-child-of="${CSS.escape(key)}"]:not([data-parent-key])`);
      children.forEach(ch => { 
        ch.style.display = nextOpen ? '' : 'none';
        // If closing and child is a layer parent, also close it
        if(!nextOpen && ch.classList.contains('layer-parent')){
          ch.setAttribute('data-open', 'false');
          toggleDescendants(ch, false);
        }
      });
    } else if(isLayerParent){
      // For layer parents, toggle only direct layer children
      const key = parentTr.getAttribute('data-layer-key');
      const children = table.querySelectorAll(`tbody tr[data-parent-key="${CSS.escape(key)}"]`);
      children.forEach(ch => { 
        ch.style.display = nextOpen ? '' : 'none';
        // If closing and child is a layer parent, also close it
        if(!nextOpen && ch.classList.contains('layer-parent')){
          ch.setAttribute('data-open', 'false');
          toggleDescendants(ch, false);
        }
      });
    }
  }

  function setAllGroups(open){
    const table = getTable(); if(!table) return;
    
    // Get all parent rows (both group and layer parents)
    const allParents = Array.from(table.querySelectorAll('tbody tr.group-parent, tbody tr.layer-parent'));
    
    // Set their state
    allParents.forEach(parent => {
      const currentState = parent.getAttribute('data-open') !== 'false';
      // Only toggle if state is different
      if(currentState !== open){
        toggleParentRow(parent);
      }
    });
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

    const groupParents = rows.filter(r => r.hasAttribute('data-group-key'));
    const layerParents = rows.filter(r => r.hasAttribute('data-layer-key'));
    const childrenByGroup = new Map();
    const childrenByLayer = new Map();
    
    rows.forEach(r => {
      const groupOf = r.getAttribute('data-group-child-of');
      const parentKey = r.getAttribute('data-parent-key');
      if(groupOf && !r.hasAttribute('data-layer-child-of')){ 
        if(!childrenByGroup.has(groupOf)) childrenByGroup.set(groupOf, []); 
        childrenByGroup.get(groupOf).push(r); 
      }
      if(parentKey){ 
        if(!childrenByLayer.has(parentKey)) childrenByLayer.set(parentKey, []); 
        childrenByLayer.get(parentKey).push(r); 
      }
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

    // Handle group parents
    groupParents.forEach(parent => {
      const key = parent.getAttribute('data-group-key');
      const kids = childrenByGroup.get(key) || [];
      const parentMatch = rowMatches(parent);
      const anyChildMatch = kids.some(rowMatches);
      const showParent = parentMatch || anyChildMatch;
      parent.style.display = showParent ? '' : 'none';
      kids.forEach(k => { k.style.display = showParent && rowMatches(k) ? '' : 'none'; });
    });

    // Handle layer parents
    layerParents.forEach(parent => {
      const key = parent.getAttribute('data-layer-key');
      const kids = childrenByLayer.get(key) || [];
      const parentMatch = rowMatches(parent);
      const anyChildMatch = kids.some(rowMatches);
      const showParent = parentMatch || anyChildMatch;
      if(parent.style.display !== 'none'){ // Don't override group visibility
        parent.style.display = showParent ? '' : 'none';
      }
      kids.forEach(k => { 
        if(k.style.display !== 'none'){ // Don't override group visibility
          k.style.display = showParent && rowMatches(k) ? '' : 'none'; 
        }
      });
    });

    // Handle rows that are neither parents nor grouped children
    rows.filter(r => !r.hasAttribute('data-group-key') && !r.hasAttribute('data-group-child-of') && !r.hasAttribute('data-layer-key') && !r.hasAttribute('data-parent-key'))
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
      actionTd.appendChild(groupBtn);
      
      const groupClimateBtn = document.createElement('button'); groupClimateBtn.type = 'button'; groupClimateBtn.textContent = 'Mappa klimatresurs';
      groupClimateBtn.addEventListener('click', function(ev){ ev.stopPropagation(); openClimateModal({ type: 'group', key: String(key) }); });
      actionTd.appendChild(groupClimateBtn);
      
      parentTr.appendChild(actionTd);
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
        // Store original row data as a custom property for later use
        tr._originalRowData = r;
        
        // Row action cell
        const actionTd = document.createElement('td');
        const rowBtn = document.createElement('button'); rowBtn.type = 'button'; rowBtn.textContent = 'Skikta';
        rowBtn.addEventListener('click', function(ev){ ev.stopPropagation(); openLayerModal({ type: 'row', rowEl: tr }); });
        actionTd.appendChild(rowBtn);
        
        const rowClimateBtn = document.createElement('button'); rowClimateBtn.type = 'button'; rowClimateBtn.textContent = 'Mappa klimatresurs';
        rowClimateBtn.addEventListener('click', function(ev){ ev.stopPropagation(); openClimateModal({ type: 'row', rowEl: tr }); });
        actionTd.appendChild(rowClimateBtn);
        
        tr.appendChild(actionTd);
        r.forEach(c => { const td = document.createElement('td'); td.textContent = c; tr.appendChild(td); });
        tbody.appendChild(tr);
      });

      
    });

    table.appendChild(tbody);
    
    // Apply saved layers and climate after table is fully assembled
    const allRows = Array.from(tbody.querySelectorAll('tr[data-group-child-of]'));
    allRows.forEach(tr => {
      // Use stored original row data instead of reading from DOM
      const rowData = tr._originalRowData;
      if(rowData){
        applySavedLayers(tr, rowData);
        applySavedClimate(tr, rowData);
      }
    });
    
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
        // Store original row data as a custom property for later use
        tr._originalRowData = r;
        
        const actionTd = document.createElement('td');
        const rowBtn = document.createElement('button'); rowBtn.type = 'button'; rowBtn.textContent = 'Skikta';
        rowBtn.addEventListener('click', function(ev){ ev.stopPropagation(); openLayerModal({ type: 'row', rowEl: tr }); });
        actionTd.appendChild(rowBtn);
        
        const rowClimateBtn = document.createElement('button'); rowClimateBtn.type = 'button'; rowClimateBtn.textContent = 'Mappa klimatresurs';
        rowClimateBtn.addEventListener('click', function(ev){ ev.stopPropagation(); openClimateModal({ type: 'row', rowEl: tr }); });
        actionTd.appendChild(rowClimateBtn);
        
        tr.appendChild(actionTd);
        r.forEach(c => { const td = document.createElement('td'); td.textContent = c; tr.appendChild(td); });
        tbody.appendChild(tr);
      });
      table.appendChild(tbody);
      output.innerHTML = ''; output.appendChild(table);
      
      // Apply saved layers and climate after table is fully assembled
      const allRows = Array.from(tbody.querySelectorAll('tr'));
      allRows.forEach(tr => {
        // Use stored original row data instead of reading from DOM
        const rowData = tr._originalRowData;
        if(rowData){
          applySavedLayers(tr, rowData);
          applySavedClimate(tr, rowData);
        }
      });
    } else {
      const table = buildGroupedTable(headers, bodyRows, groupIdx);
      output.innerHTML = ''; output.appendChild(table);
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
  
  // Climate resource modal behavior
  function openClimateModal(target){
    climateTarget = target;
    if(climateModal){ climateModal.style.display = 'flex'; }
  }
  function closeClimateModal(){
    climateTarget = null;
    if(climateModal){ climateModal.style.display = 'none'; }
  }
  if(climateCancelBtn){ climateCancelBtn.addEventListener('click', closeClimateModal); }
  if(climateModal){ climateModal.addEventListener('click', function(e){ if(e.target === climateModal) closeClimateModal(); }); }
  if(layerApplyBtn){
    layerApplyBtn.addEventListener('click', function(){
      const count = Math.max(1, parseInt(layerCountInput && layerCountInput.value || '1', 10));
      const raw = (layerThicknessesInput && layerThicknessesInput.value || '').trim();
      const thicknesses = raw ? raw.split(',').map(s => parseFloat(s.trim().replace(',', '.'))).filter(n => Number.isFinite(n) && n > 0) : [];
      applyLayerSplit(count, thicknesses);
      closeLayerModal();
    });
  }
  if(climateApplyBtn){
    climateApplyBtn.addEventListener('click', function(){
      const selectedIndex = climateResourceSelect && climateResourceSelect.value;
      if(selectedIndex !== '' && window.climateResources && window.climateResources[selectedIndex]){
        const resource = window.climateResources[selectedIndex];
        applyClimateResource(resource);
      }
      closeClimateModal();
    });
  }

  function applyLayerSplit(count, thicknesses){
    if(!layerTarget){ return; }
    const table = getTable(); if(!table) return;
    const tbody = table.querySelector('tbody'); if(!tbody) return;

    function cloneRowWithMultiplier(srcTr, multiplier, layerIndex, totalLayers){
      const clone = srcTr.cloneNode(true);
      // Replace action buttons with new ones (without old listeners)
      const actionTd = clone.querySelector('td:first-child');
      if(actionTd){
        actionTd.innerHTML = ''; // Clear old buttons
        
        // Add "Skikta" button
        const layerBtn = document.createElement('button');
        layerBtn.type = 'button';
        layerBtn.textContent = 'Skikta';
        layerBtn.addEventListener('click', function(ev){ 
          ev.stopPropagation(); 
          openLayerModal({ type: 'row', rowEl: clone }); 
        });
        actionTd.appendChild(layerBtn);
        
        // Add "Mappa klimatresurs" button
        const climateBtn = document.createElement('button');
        climateBtn.type = 'button';
        climateBtn.textContent = 'Mappa klimatresurs';
        climateBtn.addEventListener('click', function(ev){ 
          ev.stopPropagation(); 
          openClimateModal({ type: 'row', rowEl: clone }); 
        });
        actionTd.appendChild(climateBtn);
      }
      
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

    function splitRow(tr, savedLayerKey){
      // Use saved layer key if provided, otherwise generate new one
      const layerKey = savedLayerKey || 'layer-' + Date.now() + '-' + Math.random().toString(36).substr(2, 9);
      
      // Save layer data for this row
      if(!savedLayerKey){
        // Use original row data if available, otherwise extract from DOM
        const hasOriginal = !!tr._originalRowData;
        const rowData = tr._originalRowData || getRowDataFromTr(tr);
        if(rowData && Array.isArray(rowData)){
          const layerChildOf = tr.getAttribute('data-layer-child-of');
          const signature = getRowSignature(rowData, layerChildOf);
          const beforeSize = layerData.size;
          layerData.set(signature, { count, thicknesses, layerKey });
          const afterSize = layerData.size;
          console.log('💾 SAVE:', rowData[1]?.substring(0,10), '- Has _originalRowData:', hasOriginal, '- LayerChild:', layerChildOf?.substring(0,10) || 'none', '- Size:', beforeSize, '→', afterSize, '- Signature:', signature.substring(0, 60));
        }
      }
      
      // Even split if no thicknesses provided
      const multipliers = thicknesses.length > 0
        ? thicknesses.map(t => t / thicknesses.reduce((a,b)=>a+b,0))
        : Array(count).fill(1 / count);
      
      // Convert original row to parent row
      tr.classList.remove('is-new');
      tr.classList.add('layer-parent');
      // Don't add 'group-parent' class, only layer-parent
      tr.setAttribute('data-layer-key', layerKey);
      tr.setAttribute('data-open', 'true');
      
      // Update action buttons on parent row
      const actionTd = tr.querySelector('td:first-child');
      if(actionTd){
        actionTd.innerHTML = '';
        const parentLayerBtn = document.createElement('button');
        parentLayerBtn.type = 'button';
        parentLayerBtn.textContent = 'Skikta skikt';
        parentLayerBtn.addEventListener('click', function(ev){
          ev.stopPropagation();
          openLayerModal({ type: 'group', key: layerKey });
        });
        actionTd.appendChild(parentLayerBtn);
        
        const parentClimateBtn = document.createElement('button');
        parentClimateBtn.type = 'button';
        parentClimateBtn.textContent = 'Mappa klimatresurs';
        parentClimateBtn.addEventListener('click', function(ev){
          ev.stopPropagation();
          openClimateModal({ type: 'group', key: layerKey });
        });
        actionTd.appendChild(parentClimateBtn);
      }
      
      // Add toggle to first data cell
      const firstDataTd = tr.querySelector('td:nth-child(2)');
      if(firstDataTd){
        const toggle = document.createElement('span');
        toggle.className = 'group-toggle';
        toggle.innerHTML = '<svg viewBox="0 0 24 24" aria-hidden="true"><path fill="currentColor" d="M8.59 16.59L13.17 12 8.59 7.41 10 6l6 6-6 6z"/></svg>';
        firstDataTd.insertBefore(toggle, firstDataTd.firstChild);
        
        const layerLabel = document.createElement('span');
        layerLabel.textContent = ' [' + multipliers.length + ' skikt]';
        layerLabel.style.marginLeft = '4px';
        firstDataTd.appendChild(layerLabel);
      }
      
      // Preserve parent's group membership and original data for children
      const parentGroupKey = tr.getAttribute('data-group-child-of');
      const originalRowData = tr._originalRowData;
      
      // Create child layer rows
      const fragments = multipliers.map((m, i) => {
        const clone = cloneRowWithMultiplier(tr, m, i, multipliers.length);
        // Mark as child of this layer
        clone.setAttribute('data-layer-child-of', layerKey);
        // Also inherit parent's group membership if it exists
        if(parentGroupKey){
          clone.setAttribute('data-group-child-of', parentGroupKey);
        }
        // Set immediate parent for toggle
        clone.setAttribute('data-parent-key', layerKey);
        // Preserve original row data
        if(originalRowData){
          clone._originalRowData = originalRowData;
        }
        return clone;
      });
      
      // Insert layer children right after the parent
      let insertAfter = tr;
      fragments.forEach(f => {
        tbody.insertBefore(f, insertAfter.nextSibling);
        insertAfter = f;
      });
    }

    if(layerTarget.type === 'row' && layerTarget.rowEl){
      splitRow(layerTarget.rowEl);
    } else if(layerTarget.type === 'group' && layerTarget.key != null){
      const rows = Array.from(tbody.querySelectorAll('tr[data-group-child-of="' + CSS.escape(layerTarget.key) + '"]'));
      console.log('🔧 Skiktar grupp - antal rader:', rows.length);
      rows.forEach((row, index) => {
        console.log('🔧 Skiktar rad', index + 1, 'av', rows.length);
        splitRow(row);
      });
    }
    // Re-apply filters to keep visibility consistent
    applyFilters();
  }

  function applyClimateResource(resource){
    if(!climateTarget){ return; }
    const table = getTable(); if(!table) return;
    const thead = table.querySelector('thead'); if(!thead) return;
    const tbody = table.querySelector('tbody'); if(!tbody) return;
    
    // Check if "Klimatresurs" column already exists
    const headerRow = thead.querySelector('tr');
    const existingClimateHeader = Array.from(headerRow.children).find(th => th.textContent === 'Klimatresurs');
    
    if(!existingClimateHeader){
      // Add "Klimatresurs" header
      const climateTh = document.createElement('th');
      climateTh.textContent = 'Klimatresurs';
      headerRow.appendChild(climateTh);
    }
    
    const resourceName = resource.Name || 'Namnlös resurs';
    
    function addClimateToRow(tr){
      // Check if row already has a climate cell
      const existingClimateCell = tr.querySelector('td[data-climate-cell="true"]');
      if(existingClimateCell){
        existingClimateCell.textContent = resourceName;
      } else {
        // Add new climate cell
        const climateTd = document.createElement('td');
        climateTd.textContent = resourceName;
        climateTd.setAttribute('data-climate-cell', 'true');
        tr.appendChild(climateTd);
      }
      
      // Save climate data for this row
      // Use original row data if available, otherwise extract from DOM
      const rowData = tr._originalRowData || getRowDataFromTr(tr);
      if(rowData){
        const layerChildOf = tr.getAttribute('data-layer-child-of');
        const signature = getRowSignature(rowData, layerChildOf);
        climateData.set(signature, resourceName);
      }
    }
    
    if(climateTarget.type === 'row' && climateTarget.rowEl){
      addClimateToRow(climateTarget.rowEl);
    } else if(climateTarget.type === 'group' && climateTarget.key != null){
      const rows = Array.from(tbody.querySelectorAll('tr[data-group-child-of="' + CSS.escape(climateTarget.key) + '"]'));
      rows.forEach(addClimateToRow);
    }
    
    // Re-apply filters to keep visibility consistent
    applyFilters();
  }
  
  output.addEventListener('input', function(e){ const t = e.target; if(t && t.closest && t.closest('thead') && t.tagName === 'INPUT'){ applyFilters(); } });
  
  // Centralized event listener for toggling parent rows
  output.addEventListener('click', function(e){
    // Find if the click was on a parent row (group-parent or layer-parent)
    const parentTr = e.target && e.target.closest && (
      e.target.closest('tr.group-parent') || 
      e.target.closest('tr.layer-parent')
    );
    
    if(!parentTr) return;
    
    // Don't toggle if clicking on a button
    if(e.target.tagName === 'BUTTON' || e.target.closest('button')) return;
    
    toggleParentRow(parentTr);
  });
})();
