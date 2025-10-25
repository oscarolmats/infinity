(function(){
  const fileInput = document.getElementById('fileInput');
  const filterInput = document.getElementById('filterInput');
  const toggleAllBtn = document.getElementById('toggleAllBtn');
  const exportBtn = document.getElementById('exportBtn');
  const saveProjectBtn = document.getElementById('saveProjectBtn');
  const saveAsProjectBtn = document.getElementById('saveAsProjectBtn');
  const loadProjectBtn = document.getElementById('loadProjectBtn');
  const projectFileInput = document.getElementById('projectFileInput');
  const groupBySelect = document.getElementById('groupBy');
  let lastRows = null; // cache of parsed rows for re-rendering
  let lastHeaders = null; // cache of headers for project save/load
  let originalFileName = null; // store original data file name
  let savedFileHandle = null; // store file handle for quick save
  
  // Storage for layers and climate resources
  let layerData = new Map(); // key: row signature -> { count, thicknesses, layerKey }
  let climateData = new Map(); // key: row signature or layerKey -> resourceName
  
  // Layer modal refs
  const layerModal = document.getElementById('layerModal');
  const layerCountInput = document.getElementById('layerCount');
  const layerThicknessesInput = document.getElementById('layerThicknesses');
  const layerCancelBtn = document.getElementById('layerCancel');
  const layerApplyBtn = document.getElementById('layerApply');
  const layerMapClimateBtn = document.getElementById('layerMapClimate');
  // Mixed layer refs
  const mixedLayerCheckboxes = document.getElementById('mixedLayerCheckboxes');
  const mixedLayerDetails = document.getElementById('mixedLayerDetails');
  const mixedLayerConfigs = document.getElementById('mixedLayerConfigs');
  const layerNamesSection = document.getElementById('layerNamesSection');
  const layerNamesContainer = document.getElementById('layerNamesContainer');
  let layerTarget = null; // { type: 'row'|'group', key?: string, rowEl?: HTMLTableRowElement }
  
  // Climate resource modal refs
  const climateModal = document.getElementById('climateModal');
  const climateResourceSelect = document.getElementById('climateResourceSelect');
  const climateCancelBtn = document.getElementById('climateCancel');
  const climateApplyBtn = document.getElementById('climateApply');
  let climateTarget = null; // { type: 'row'|'group', key?: string, rowEl?: HTMLTableRowElement }
  
  // Multi-layer climate modal refs
  const multiLayerClimateModal = document.getElementById('multiLayerClimateModal');
  const multiLayerClimateContent = document.getElementById('multiLayerClimateContent');
  const multiLayerClimateCancelBtn = document.getElementById('multiLayerClimateCancel');
  const multiLayerClimateApplyBtn = document.getElementById('multiLayerClimateApply');
  let multiLayerClimateTarget = null; // { layerRows: [], groupKey: string }
  
  // Manual factor modal refs
  const manualFactorModal = document.getElementById('manualFactorModal');
  const manualFactorResourceName = document.getElementById('manualFactorResourceName');
  const manualFactorValue = document.getElementById('manualFactorValue');
  const manualFactorUnit = document.getElementById('manualFactorUnit');
  const manualFactorCancelBtn = document.getElementById('manualFactorCancel');
  const manualFactorApplyBtn = document.getElementById('manualFactorApply');
  let manualFactorCallback = null; // Callback function to continue with manual values
  
  const output = document.getElementById('output');
  
  // Undo/Redo functionality
  const undoBtn = document.getElementById('undoBtn');
  const redoBtn = document.getElementById('redoBtn');
  let undoStack = [];
  let redoStack = [];
  const maxUndoSteps = 50; // Limit history to prevent memory issues
  let isRestoringState = false; // Flag to prevent saving during restore

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
    const res = await fetch('/api/upload', { method: 'POST', body: form });
    if(!res.ok){ throw new Error('Uppladdning misslyckades'); }
    return await res.text();
  }

  function getTable(){ return output.querySelector('table'); }
  
  // Save current state for undo/redo
  function saveState(){
    // Don't save state if we're currently restoring
    if(isRestoringState) return;
    
    const state = {
      outputHTML: output.innerHTML,
      layerData: new Map(layerData),
      climateData: new Map(climateData),
      lastRows: lastRows ? JSON.parse(JSON.stringify(lastRows)) : null,
      lastHeaders: lastHeaders ? [...lastHeaders] : null,
      groupByValue: groupBySelect ? groupBySelect.value : ''
    };
    
    undoStack.push(state);
    
    // Limit stack size
    if(undoStack.length > maxUndoSteps){
      undoStack.shift();
    }
    
    // Clear redo stack when new action is performed
    redoStack = [];
    
    updateUndoRedoButtons();
  }
  
  // Restore state
  function restoreState(state){
    if(!state) return;
    
    isRestoringState = true;
    
    output.innerHTML = state.outputHTML;
    layerData = new Map(state.layerData);
    climateData = new Map(state.climateData);
    lastRows = state.lastRows ? JSON.parse(JSON.stringify(state.lastRows)) : null;
    lastHeaders = state.lastHeaders ? [...state.lastHeaders] : null;
    
    if(groupBySelect && state.groupByValue !== undefined){
      groupBySelect.value = state.groupByValue;
    }
    
    // Re-attach event listeners to the restored table
    reattachTableEventListeners();
    
    // Update climate summary
    setTimeout(() => updateClimateSummary(), 100);
    
    isRestoringState = false;
  }
  
  // Update undo/redo button states
  function updateUndoRedoButtons(){
    if(undoBtn){
      undoBtn.disabled = undoStack.length === 0;
    }
    if(redoBtn){
      redoBtn.disabled = redoStack.length === 0;
    }
  }
  
  // Perform undo
  function performUndo(){
    if(undoStack.length === 0) return;
    
    // Save current state to redo stack
    const currentState = {
      outputHTML: output.innerHTML,
      layerData: new Map(layerData),
      climateData: new Map(climateData),
      lastRows: lastRows ? JSON.parse(JSON.stringify(lastRows)) : null,
      lastHeaders: lastHeaders ? [...lastHeaders] : null,
      groupByValue: groupBySelect ? groupBySelect.value : ''
    };
    redoStack.push(currentState);
    
    // Restore previous state
    const previousState = undoStack.pop();
    restoreState(previousState);
    
    updateUndoRedoButtons();
    console.log('↶ Undo performed');
  }
  
  // Perform redo
  function performRedo(){
    if(redoStack.length === 0) return;
    
    // Save current state to undo stack
    const currentState = {
      outputHTML: output.innerHTML,
      layerData: new Map(layerData),
      climateData: new Map(climateData),
      lastRows: lastRows ? JSON.parse(JSON.stringify(lastRows)) : null,
      lastHeaders: lastHeaders ? [...lastHeaders] : null,
      groupByValue: groupBySelect ? groupBySelect.value : ''
    };
    undoStack.push(currentState);
    
    // Restore next state
    const nextState = redoStack.pop();
    restoreState(nextState);
    
    updateUndoRedoButtons();
    console.log('↷ Redo performed');
  }
  
  // Re-attach event listeners after restoring state
  function reattachTableEventListeners(){
    const table = getTable();
    if(!table) return;
    const tbody = table.querySelector('tbody');
    if(!tbody) return;
    
    console.log('🔗 Re-attaching event listeners after state restore');
    
    // Re-attach toggle listeners for group/layer parents
    const parents = Array.from(table.querySelectorAll('tr.group-parent, tr.layer-parent'));
    parents.forEach(parent => {
      parent.onclick = function(e){
        if(e.target.closest('button')) return;
        toggleParentRow(parent);
      };
    });
    
    // Re-attach button listeners for all rows
    const allButtons = tbody.querySelectorAll('button');
    let buttonCount = { skikta: 0, skiktaGrupp: 0, skiktaSkikt: 0, klimat: 0 };
    
    allButtons.forEach(button => {
      const buttonText = button.textContent.trim();
      
      // Skikta buttons
      if(buttonText === 'Skikta'){
        buttonCount.skikta++;
        const row = button.closest('tr');
        button.onclick = function(ev){
          ev.stopPropagation();
          openLayerModal({ type: 'row', rowEl: row });
        };
      }
      // Skikta grupp buttons
      else if(buttonText === 'Skikta grupp'){
        buttonCount.skiktaGrupp++;
        const row = button.closest('tr');
        const groupKey = row.getAttribute('data-group-key');
        button.onclick = function(ev){
          ev.stopPropagation();
          openLayerModal({ type: 'group', key: String(groupKey) });
        };
      }
      // Skikta skikt buttons
      else if(buttonText === 'Skikta skikt'){
        buttonCount.skiktaSkikt++;
        const row = button.closest('tr');
        const layerKey = row.getAttribute('data-layer-key');
        button.onclick = function(ev){
          ev.stopPropagation();
          openLayerModal({ type: 'group', key: layerKey });
        };
      }
      // Mappa klimatresurs buttons for rows
      else if(buttonText === 'Mappa klimatresurs'){
        buttonCount.klimat++;
        const row = button.closest('tr');
        
        // Check if this is a group parent or layer parent
        if(row.classList.contains('group-parent') && row.hasAttribute('data-group-key')){
          const groupKey = row.getAttribute('data-group-key');
          button.onclick = function(ev){
            ev.stopPropagation();
            // Check if this group has been layered
            const layerRows = Array.from(tbody.querySelectorAll(`tr[data-group-child-of="${CSS.escape(String(groupKey))}"][data-layer-key]`));
            if(layerRows.length > 0){
              // Group is layered, open multi-layer climate modal
              saveState();
              openMultiLayerClimateModal(String(groupKey));
            } else {
              // Group is not layered, open regular climate modal
              openClimateModal({ type: 'group', key: String(groupKey) });
            }
          };
        } else if(row.classList.contains('layer-parent') && row.hasAttribute('data-layer-key')){
          const layerKey = row.getAttribute('data-layer-key');
          button.onclick = function(ev){
            ev.stopPropagation();
            openClimateModal({ type: 'group', key: layerKey });
          };
        } else {
          // Regular row
          button.onclick = function(ev){
            ev.stopPropagation();
            openClimateModal({ type: 'row', rowEl: row });
          };
        }
      }
    });
    
    console.log('✅ Re-attached listeners:', buttonCount);
    
    // Re-attach _originalRowData to rows that had it
    // This is needed for layering functionality to work correctly
    const allRows = Array.from(tbody.querySelectorAll('tr[data-group-child-of]'));
    allRows.forEach(tr => {
      // Try to restore original row data by reading it from the DOM
      const rowData = getRowDataFromTr(tr);
      if(rowData){
        tr._originalRowData = rowData;
      }
    });
    
    // Re-apply filters
    applyFilters();
  }

  // Generate unique signature for a row based on its original data and layer position
  // rowData should be the original array data, not the DOM elements
  function getRowSignature(rowData, layerChildOf = null){
    const baseSignature = JSON.stringify(rowData);
    // Include layer hierarchy in signature to differentiate nested layers
    return layerChildOf ? `${baseSignature}::layer::${layerChildOf}` : baseSignature;
  }
  
  // Get row data from a TR element, excluding action column and climate/factor columns
  function getRowDataFromTr(tr){
    const table = tr.closest('table');
    if(!table) return null;
    const headers = Array.from(table.querySelectorAll('thead tr:first-child th')).map(th => th.textContent);
    const climateColIndex = headers.findIndex(h => h === 'Klimatresurs');
    const factorColIndex = headers.findIndex(h => h === 'Omräkningsfaktor');
      const unitColIndex = headers.findIndex(h => h === 'Omräkningsfaktor enhet');
      const wasteColIndex = headers.findIndex(h => h === 'Spillfaktor');
      const A1_A3ColIndex = headers.findIndex(h => h === 'Emissionsfaktor A1-A3');
      const A4ColIndex = headers.findIndex(h => h === 'Emissionsfaktor A4');
      const A5ColIndex = headers.findIndex(h => h === 'Emissionsfaktor A5');
      const inbyggdViktColIndex = headers.findIndex(h => h === 'Inbyggd vikt');
      const inkoptViktColIndex = headers.findIndex(h => h === 'Inköpt vikt');
      const klimatA1A3ColIndex = headers.findIndex(h => h === 'Klimatpåverkan A1-A3');
      const klimatA4ColIndex = headers.findIndex(h => h === 'Klimatpåverkan A4');
      const klimatA5ColIndex = headers.findIndex(h => h === 'Klimatpåverkan A5');
      
      const cells = Array.from(tr.children);
      const rowData = [];
      
      // Skip first cell (action column) and climate/factor/unit/waste columns if they exist
      for(let i = 1; i < cells.length; i++){
        if((climateColIndex !== -1 && i === climateColIndex) || 
           (factorColIndex !== -1 && i === factorColIndex) || 
           (unitColIndex !== -1 && i === unitColIndex) ||
           (wasteColIndex !== -1 && i === wasteColIndex) ||
           (A1_A3ColIndex !== -1 && i === A1_A3ColIndex) ||
           (A4ColIndex !== -1 && i === A4ColIndex) ||
           (A5ColIndex !== -1 && i === A5ColIndex) ||
           (inbyggdViktColIndex !== -1 && i === inbyggdViktColIndex) ||
           (inkoptViktColIndex !== -1 && i === inkoptViktColIndex) ||
           (klimatA1A3ColIndex !== -1 && i === klimatA1A3ColIndex) ||
           (klimatA4ColIndex !== -1 && i === klimatA4ColIndex) ||
           (klimatA5ColIndex !== -1 && i === klimatA5ColIndex)
           ){
          continue; // Skip climate, factor, unit, waste, emission, weight and climate impact columns
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
      tr.setAttribute('data-open', 'false'); // Start collapsed by default
      
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
        
        // Read Net Area BEFORE scaling it (for volume calculation)
        let originalNetArea = null;
        if(idxNetArea >= 0){
          const netAreaTd = tds[idxNetArea + 1];
          if(netAreaTd){
            originalNetArea = parseNumberLike(netAreaTd.textContent);
          }
        }
        
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
        
        // Scale Net Area and Count with multiplier
        scaleCell(idxNetArea); 
        scaleCell(idxCount);
        
        // For Volume: if we have thickness specified, calculate Volume = Net Area × thickness (in meters)
        const layerThickness = thicknesses.length > 0 ? thicknesses[i] : undefined;
        console.log('🔍 Saved layer', i + 1, '- layerThickness:', layerThickness, 'idxVolume:', idxVolume, 'originalNetArea:', originalNetArea);
        
        if(layerThickness && idxVolume >= 0 && originalNetArea !== null && Number.isFinite(originalNetArea)){
          const volumeTd = tds[idxVolume + 1]; // +1 offset for action column
          console.log('🔍 volumeTd found:', !!volumeTd);
          if(volumeTd){
            // Thickness is always in mm, convert to meters
            const thicknessInMeters = layerThickness / 1000;
            console.log('🔍 Converting from mm:', layerThickness, '→', thicknessInMeters, 'm');
            
            const newVolume = originalNetArea * thicknessInMeters;
            console.log('✅ Calculated volume:', originalNetArea, '×', thicknessInMeters, '=', newVolume);
            volumeTd.textContent = String(newVolume);
            console.log('📝 Set volumeTd.textContent to:', volumeTd.textContent, '(cell:', volumeTd, ')');
          }
        } else {
          console.log('❌ Falling back to multiplier. layerThickness:', layerThickness, 'idxVolume:', idxVolume, 'originalNetArea:', originalNetArea);
          if(idxVolume >= 0){
            // No thickness specified, use multiplier
            scaleCell(idxVolume);
          }
        }
        
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
    const climateInfo = climateData.get(signature);
    if(climateInfo){
      const table = getTable(); if(!table) return;
      const thead = table.querySelector('thead'); if(!thead) return;
      
      const headerRow = thead.querySelector('tr');
      const existingClimateHeader = Array.from(headerRow.children).find(th => th.textContent === 'Klimatresurs');
      const existingFactorHeader = Array.from(headerRow.children).find(th => th.textContent === 'Omräkningsfaktor');
      const existingUnitHeader = Array.from(headerRow.children).find(th => th.textContent === 'Omräkningsfaktor enhet');
      const existingWasteHeader = Array.from(headerRow.children).find(th => th.textContent === 'Spillfaktor');
      const existingA1_A3Header = Array.from(headerRow.children).find(th => th.textContent === 'Emissionsfaktor A1-A3');
      const existingA4Header = Array.from(headerRow.children).find(th => th.textContent === 'Emissionsfaktor A4');
      const existingA5Header = Array.from(headerRow.children).find(th => th.textContent === 'Emissionsfaktor A5');
      
      if(!existingClimateHeader){
        const climateTh = document.createElement('th');
        climateTh.textContent = 'Klimatresurs';
        headerRow.appendChild(climateTh);
      }
      
      if(!existingFactorHeader){
        const factorTh = document.createElement('th');
        factorTh.textContent = 'Omräkningsfaktor';
        headerRow.appendChild(factorTh);
      }
      
      if(!existingUnitHeader){
        const unitTh = document.createElement('th');
        unitTh.textContent = 'Omräkningsfaktor enhet';
        headerRow.appendChild(unitTh);
      }
      
      if(!existingWasteHeader){
        const wasteTh = document.createElement('th');
        wasteTh.textContent = 'Spillfaktor';
        headerRow.appendChild(wasteTh);
      }
      
      if(!existingA1_A3Header){
        const a1a3Th = document.createElement('th');
        a1a3Th.textContent = 'Emissionsfaktor A1-A3';
        headerRow.appendChild(a1a3Th);
      }
      
      if(!existingA4Header){
        const a4Th = document.createElement('th');
        a4Th.textContent = 'Emissionsfaktor A4';
        headerRow.appendChild(a4Th);
      }
      
      if(!existingA5Header){
        const a5Th = document.createElement('th');
        a5Th.textContent = 'Emissionsfaktor A5';
        headerRow.appendChild(a5Th);
      }
      
      const existingInbyggdViktHeader = Array.from(headerRow.children).find(th => th.textContent === 'Inbyggd vikt');
      if(!existingInbyggdViktHeader){
        const inbyggdTh = document.createElement('th');
        inbyggdTh.textContent = 'Inbyggd vikt';
        headerRow.appendChild(inbyggdTh);
      }
      
      const existingInkoptViktHeader = Array.from(headerRow.children).find(th => th.textContent === 'Inköpt vikt');
      if(!existingInkoptViktHeader){
        const inkoptTh = document.createElement('th');
        inkoptTh.textContent = 'Inköpt vikt';
        headerRow.appendChild(inkoptTh);
      }
      
      // Add climate impact columns
      const existingKlimatA1A3Header = Array.from(headerRow.children).find(th => th.textContent === 'Klimatpåverkan A1-A3');
      if(!existingKlimatA1A3Header){
        const klimatA1A3Th = document.createElement('th');
        klimatA1A3Th.textContent = 'Klimatpåverkan A1-A3';
        headerRow.appendChild(klimatA1A3Th);
      }
      
      const existingKlimatA4Header = Array.from(headerRow.children).find(th => th.textContent === 'Klimatpåverkan A4');
      if(!existingKlimatA4Header){
        const klimatA4Th = document.createElement('th');
        klimatA4Th.textContent = 'Klimatpåverkan A4';
        headerRow.appendChild(klimatA4Th);
      }
      
      const existingKlimatA5Header = Array.from(headerRow.children).find(th => th.textContent === 'Klimatpåverkan A5');
      if(!existingKlimatA5Header){
        const klimatA5Th = document.createElement('th');
        klimatA5Th.textContent = 'Klimatpåverkan A5';
        headerRow.appendChild(klimatA5Th);
      }
      
      // Handle both old format (string) and new format (object)
      const resourceName = typeof climateInfo === 'string' ? climateInfo : climateInfo.name;
      const conversionFactor = typeof climateInfo === 'object' ? climateInfo.factor : 'N/A';
      const conversionUnit = typeof climateInfo === 'object' ? climateInfo.unit : 'N/A';
      const wasteFactor = typeof climateInfo === 'object' ? climateInfo.waste : 'N/A';
      const a1a3Factor = typeof climateInfo === 'object' ? climateInfo.a1a3 : 'N/A';
      const a4Factor = typeof climateInfo === 'object' ? climateInfo.a4 : 'N/A';
      const a5Factor = typeof climateInfo === 'object' ? climateInfo.a5 : 'N/A';
      
      // Calculate Inbyggd vikt and Inköpt vikt
      let inbyggdVikt = 'N/A';
      let inkoptVikt = 'N/A';
      
      // Get headers to find Volume and Net Area columns
      const allHeaders = Array.from(headerRow.children).map(th => th.textContent);
      const volumeColIndex = allHeaders.findIndex(h => String(h).toLowerCase() === 'volume');
      const netAreaColIndex = allHeaders.findIndex(h => String(h).toLowerCase() === 'net area');
      
      console.log('🔍 Beräknar vikt - Unit:', conversionUnit, 'Factor:', conversionFactor, 'Waste:', wasteFactor);
      console.log('🔍 Column indices - Volume:', volumeColIndex, 'NetArea:', netAreaColIndex);
      console.log('🔍 Headers:', allHeaders);
      
      if(conversionFactor !== 'N/A' && Number.isFinite(parseFloat(conversionFactor))){
        const factor = parseFloat(conversionFactor);
        const cells = Array.from(tr.children);
        
        console.log('🔍 Factor is valid:', factor);
        
        // Normalize unit to handle both kg/m3 and kg/m³ (with superscript)
        const normalizedUnit = String(conversionUnit).replace(/[²³]/g, function(match){
          return match === '²' ? '2' : '3';
        });
        console.log('🔍 Normalized unit:', normalizedUnit);
        
        if(normalizedUnit === 'kg/m3' && volumeColIndex !== -1){
          // Inbyggd vikt = Omräkningsfaktor × Volume
          const volumeCell = cells[volumeColIndex];
          console.log('🔍 Volume cell:', volumeCell?.textContent, 'at index:', volumeColIndex);
          if(volumeCell){
            const volume = parseNumberLike(volumeCell.textContent);
            console.log('🔍 Parsed volume:', volume);
            if(Number.isFinite(volume)){
              inbyggdVikt = factor * volume;
              console.log('✅ Inbyggd vikt calculated:', inbyggdVikt);
            }
          }
        } else if(normalizedUnit === 'kg/m2' && netAreaColIndex !== -1){
          // Inbyggd vikt = Omräkningsfaktor × Net Area
          const netAreaCell = cells[netAreaColIndex];
          console.log('🔍 NetArea cell:', netAreaCell?.textContent, 'at index:', netAreaColIndex);
          if(netAreaCell){
            const netArea = parseNumberLike(netAreaCell.textContent);
            console.log('🔍 Parsed netArea:', netArea);
            if(Number.isFinite(netArea)){
              inbyggdVikt = factor * netArea;
              console.log('✅ Inbyggd vikt calculated:', inbyggdVikt);
            }
          }
        } else {
          console.log('❌ Unit mismatch or column not found. Unit:', conversionUnit, 'Normalized:', normalizedUnit, 'VolumeIdx:', volumeColIndex, 'NetAreaIdx:', netAreaColIndex);
        }
        
        // Calculate Inköpt vikt = Inbyggd vikt × Spillfaktor
        if(inbyggdVikt !== 'N/A' && wasteFactor !== 'N/A' && Number.isFinite(parseFloat(wasteFactor))){
          const waste = parseFloat(wasteFactor);
          inkoptVikt = inbyggdVikt * waste;
          console.log('✅ Inköpt vikt calculated:', inkoptVikt);
        }
      } else {
        console.log('❌ Conversion factor not valid:', conversionFactor);
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
      
      const existingFactorCell = tr.querySelector('td[data-factor-cell="true"]');
      if(existingFactorCell){
        existingFactorCell.textContent = conversionFactor;
      } else {
        const factorTd = document.createElement('td');
        factorTd.textContent = conversionFactor;
        factorTd.setAttribute('data-factor-cell', 'true');
        tr.appendChild(factorTd);
      }
      
      const existingUnitCell = tr.querySelector('td[data-unit-cell="true"]');
      if(existingUnitCell){
        existingUnitCell.textContent = conversionUnit;
      } else {
        const unitTd = document.createElement('td');
        unitTd.textContent = conversionUnit;
        unitTd.setAttribute('data-unit-cell', 'true');
        tr.appendChild(unitTd);
      }
      
      const existingWasteCell = tr.querySelector('td[data-waste-cell="true"]');
      if(existingWasteCell){
        existingWasteCell.textContent = wasteFactor;
      } else {
        const wasteTd = document.createElement('td');
        wasteTd.textContent = wasteFactor;
        wasteTd.setAttribute('data-waste-cell', 'true');
        tr.appendChild(wasteTd);
      }
      
      const existingA1_A3Cell = tr.querySelector('td[data-A1_A3-cell="true"]');
      if(existingA1_A3Cell){
        existingA1_A3Cell.textContent = a1a3Factor;
      } else {
        const a1a3Td = document.createElement('td');
        a1a3Td.textContent = a1a3Factor;
        a1a3Td.setAttribute('data-A1_A3-cell', 'true');
        tr.appendChild(a1a3Td);
      }
      
      const existingA4Cell = tr.querySelector('td[data-A4-cell="true"]');
      if(existingA4Cell){
        existingA4Cell.textContent = a4Factor;
      } else {
        const a4Td = document.createElement('td');
        a4Td.textContent = a4Factor;
        a4Td.setAttribute('data-A4-cell', 'true');
        tr.appendChild(a4Td);
      }
      
      const existingA5Cell = tr.querySelector('td[data-A5-cell="true"]');
      if(existingA5Cell){
        existingA5Cell.textContent = a5Factor;
      } else {
        const a5Td = document.createElement('td');
        a5Td.textContent = a5Factor;
        a5Td.setAttribute('data-A5-cell', 'true');
        tr.appendChild(a5Td);
      }
      
      const existingInbyggdViktCell = tr.querySelector('td[data-inbyggd-vikt-cell="true"]');
      if(existingInbyggdViktCell){
        existingInbyggdViktCell.textContent = inbyggdVikt !== 'N/A' ? inbyggdVikt.toFixed(2) : 'N/A';
      } else {
        const inbyggdViktTd = document.createElement('td');
        inbyggdViktTd.textContent = inbyggdVikt !== 'N/A' ? inbyggdVikt.toFixed(2) : 'N/A';
        inbyggdViktTd.setAttribute('data-inbyggd-vikt-cell', 'true');
        tr.appendChild(inbyggdViktTd);
      }
      
      const existingInkoptViktCell = tr.querySelector('td[data-inkopt-vikt-cell="true"]');
      if(existingInkoptViktCell){
        existingInkoptViktCell.textContent = inkoptVikt !== 'N/A' ? inkoptVikt.toFixed(2) : 'N/A';
      } else {
        const inkoptViktTd = document.createElement('td');
        inkoptViktTd.textContent = inkoptVikt !== 'N/A' ? inkoptVikt.toFixed(2) : 'N/A';
        inkoptViktTd.setAttribute('data-inkopt-vikt-cell', 'true');
        tr.appendChild(inkoptViktTd);
      }
      
      // Calculate climate impact columns
      let klimatA1A3 = 'N/A';
      let klimatA4 = 'N/A';
      let klimatA5 = 'N/A';
      
      // Klimatpåverkan A1-A3 = Inbyggd vikt * Emissionsfaktor A1-A3
      if(inbyggdVikt !== 'N/A' && a1a3Factor !== 'N/A' && Number.isFinite(parseFloat(a1a3Factor))){
        klimatA1A3 = inbyggdVikt * parseFloat(a1a3Factor);
      }
      
      // Klimatpåverkan A4 = Inköpt vikt * Emissionsfaktor A4
      if(inkoptVikt !== 'N/A' && a4Factor !== 'N/A' && Number.isFinite(parseFloat(a4Factor))){
        klimatA4 = inkoptVikt * parseFloat(a4Factor);
      }
      
      // Klimatpåverkan A5 = Inköpt vikt * Emissionsfaktor A5
      if(inkoptVikt !== 'N/A' && a5Factor !== 'N/A' && Number.isFinite(parseFloat(a5Factor))){
        klimatA5 = inkoptVikt * parseFloat(a5Factor);
      }
      
      const existingKlimatA1A3Cell = tr.querySelector('td[data-klimat-a1a3-cell="true"]');
      if(existingKlimatA1A3Cell){
        existingKlimatA1A3Cell.textContent = klimatA1A3 !== 'N/A' ? klimatA1A3.toFixed(2) : 'N/A';
      } else {
        const klimatA1A3Td = document.createElement('td');
        klimatA1A3Td.textContent = klimatA1A3 !== 'N/A' ? klimatA1A3.toFixed(2) : 'N/A';
        klimatA1A3Td.setAttribute('data-klimat-a1a3-cell', 'true');
        tr.appendChild(klimatA1A3Td);
      }
      
      const existingKlimatA4Cell = tr.querySelector('td[data-klimat-a4-cell="true"]');
      if(existingKlimatA4Cell){
        existingKlimatA4Cell.textContent = klimatA4 !== 'N/A' ? klimatA4.toFixed(2) : 'N/A';
      } else {
        const klimatA4Td = document.createElement('td');
        klimatA4Td.textContent = klimatA4 !== 'N/A' ? klimatA4.toFixed(2) : 'N/A';
        klimatA4Td.setAttribute('data-klimat-a4-cell', 'true');
        tr.appendChild(klimatA4Td);
      }
      
      const existingKlimatA5Cell = tr.querySelector('td[data-klimat-a5-cell="true"]');
      if(existingKlimatA5Cell){
        existingKlimatA5Cell.textContent = klimatA5 !== 'N/A' ? klimatA5.toFixed(2) : 'N/A';
      } else {
        const klimatA5Td = document.createElement('td');
        klimatA5Td.textContent = klimatA5 !== 'N/A' ? klimatA5.toFixed(2) : 'N/A';
        klimatA5Td.setAttribute('data-klimat-a5-cell', 'true');
        tr.appendChild(klimatA5Td);
      }
    }
    
    // Update climate summary after changes
    setTimeout(() => updateClimateSummary(), 100);
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
    
    // Update climate summary after filtering
    setTimeout(() => updateClimateSummary(), 50);
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
    
    // Get existing table headers to preserve dynamically added columns
    const existingTable = getTable();
    let allHeaders = [...headers];
    if(existingTable){
      const existingHeaders = Array.from(existingTable.querySelectorAll('thead tr:first-child th')).map(th => th.textContent);
      // Add any new headers that aren't in the original headers
      const newHeaders = existingHeaders.slice(1); // Skip action column
      newHeaders.forEach(h => {
        if(!headers.includes(h)){
          allHeaders.push(h);
        }
      });
    }
    
    allHeaders.forEach(h => { const th = document.createElement('th'); th.textContent = h; headerTr.appendChild(th); });
    thead.appendChild(headerTr); table.appendChild(thead);
    const tbody = document.createElement('tbody');

    const idxType = groupColIndex;
    const idxNetArea = headers.findIndex(h => String(h).toLowerCase() === 'net area');
    const idxVolume = headers.findIndex(h => String(h).toLowerCase() === 'volume');
    const idxCount = headers.findIndex(h => String(h).toLowerCase() === 'count');
    const idxInbyggdVikt = allHeaders.findIndex(h => h === 'Inbyggd vikt');
    const idxInkoptVikt = allHeaders.findIndex(h => h === 'Inköpt vikt');

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
      parentTr.setAttribute('data-open', 'false'); // Start collapsed by default
      // Create one cell per column so sums align under headers
      // Parent action cell (group layer)
      const actionTd = document.createElement('td');
      const groupBtn = document.createElement('button'); groupBtn.type = 'button'; groupBtn.textContent = 'Skikta grupp';
      groupBtn.addEventListener('click', function(ev){ ev.stopPropagation(); openLayerModal({ type: 'group', key: String(key) }); });
      actionTd.appendChild(groupBtn);
      
      const groupClimateBtn = document.createElement('button'); groupClimateBtn.type = 'button'; groupClimateBtn.textContent = 'Mappa klimatresurs';
      groupClimateBtn.addEventListener('click', function(ev){ 
        ev.stopPropagation(); 
        // Check if this group has been layered (has layer children with layer keys)
        const table = groupBtn.closest('table');
        if(table){
          const tbody = table.querySelector('tbody');
          if(tbody){
            const layerRows = Array.from(tbody.querySelectorAll(`tr[data-group-child-of="${CSS.escape(String(key))}"][data-layer-key]`));
            if(layerRows.length > 0){
              // Group is layered, open multi-layer climate modal
              console.log('🔍 Opening multi-layer climate modal for layered group:', key);
              saveState(); // Save state before opening climate modal
              openMultiLayerClimateModal(String(key));
              return;
            }
          }
        }
        // Group is not layered, open regular climate modal
        openClimateModal({ type: 'group', key: String(key) }); 
      });
      actionTd.appendChild(groupClimateBtn);
      
      parentTr.appendChild(actionTd);
      for(let i = 0; i < allHeaders.length; i++){
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
        } else if(i === idxInbyggdVikt){
          // Mark as placeholder - will be calculated after rows are added
          td.setAttribute('data-sum-inbyggd-vikt', 'true');
          td.textContent = '';
        } else if(i === idxInkoptVikt){
          // Mark as placeholder - will be calculated after rows are added
          td.setAttribute('data-sum-inkopt-vikt', 'true');
          td.textContent = '';
        } else if(allHeaders[i] === 'Klimatpåverkan A1-A3'){
          // Mark as placeholder - will be calculated after rows are added
          td.setAttribute('data-sum-klimat-a1a3', 'true');
          td.textContent = '';
        } else if(allHeaders[i] === 'Klimatpåverkan A4'){
          // Mark as placeholder - will be calculated after rows are added
          td.setAttribute('data-sum-klimat-a4', 'true');
          td.textContent = '';
        } else if(allHeaders[i] === 'Klimatpåverkan A5'){
          // Mark as placeholder - will be calculated after rows are added
          td.setAttribute('data-sum-klimat-a5', 'true');
          td.textContent = '';
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
        // Add cells for original data
        r.forEach(c => { const td = document.createElement('td'); td.textContent = c; tr.appendChild(td); });
        // Add empty cells for any new columns that were added dynamically
        // Including climate columns with proper data attributes
        for(let i = r.length; i < allHeaders.length; i++){
          const headerName = allHeaders[i];
          const td = document.createElement('td');
          td.textContent = '';
          // Mark cells with data attributes so they can be found and updated by applySavedClimate
          if(headerName === 'Klimatresurs'){
            td.setAttribute('data-climate-cell', 'true');
          } else if(headerName === 'Omräkningsfaktor'){
            td.setAttribute('data-factor-cell', 'true');
          } else if(headerName === 'Omräkningsfaktor enhet'){
            td.setAttribute('data-unit-cell', 'true');
          } else if(headerName === 'Spillfaktor'){
            td.setAttribute('data-waste-cell', 'true');
          } else if(headerName === 'Emissionsfaktor A1-A3'){
            td.setAttribute('data-A1_A3-cell', 'true');
          } else if(headerName === 'Emissionsfaktor A4'){
            td.setAttribute('data-A4-cell', 'true');
          } else if(headerName === 'Emissionsfaktor A5'){
            td.setAttribute('data-A5-cell', 'true');
          } else if(headerName === 'Inbyggd vikt'){
            td.setAttribute('data-inbyggd-vikt-cell', 'true');
          } else if(headerName === 'Inköpt vikt'){
            td.setAttribute('data-inkopt-vikt-cell', 'true');
          } else if(headerName === 'Klimatpåverkan A1-A3'){
            td.setAttribute('data-klimat-a1a3-cell', 'true');
          } else if(headerName === 'Klimatpåverkan A4'){
            td.setAttribute('data-klimat-a4-cell', 'true');
          } else if(headerName === 'Klimatpåverkan A5'){
            td.setAttribute('data-klimat-a5-cell', 'true');
          }
          tr.appendChild(td);
        }
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
    
    // Calculate weight sums for group parents after climate data is applied
    const groupParents = Array.from(tbody.querySelectorAll('tr.group-parent'));
    groupParents.forEach(parentTr => {
      const groupKey = parentTr.getAttribute('data-group-key');
      const childRows = Array.from(tbody.querySelectorAll(`tr[data-group-child-of="${CSS.escape(groupKey)}"]`));
      
      let sumInbyggdVikt = 0;
      let sumInkoptVikt = 0;
      let countInbyggd = 0;
      let countInkopt = 0;
      
      let sumKlimatA1A3 = 0;
      let sumKlimatA4 = 0;
      let sumKlimatA5 = 0;
      let countKlimatA1A3 = 0;
      let countKlimatA4 = 0;
      let countKlimatA5 = 0;
      
      childRows.forEach(childTr => {
        const inbyggdCell = childTr.querySelector('td[data-inbyggd-vikt-cell="true"]');
        const inkoptCell = childTr.querySelector('td[data-inkopt-vikt-cell="true"]');
        
        if(inbyggdCell){
          const val = parseNumberLike(inbyggdCell.textContent);
          if(Number.isFinite(val)){
            sumInbyggdVikt += val;
            countInbyggd++;
          }
        }
        
        if(inkoptCell){
          const val = parseNumberLike(inkoptCell.textContent);
          if(Number.isFinite(val)){
            sumInkoptVikt += val;
            countInkopt++;
          }
        }
        
        const klimatA1A3Cell = childTr.querySelector('td[data-klimat-a1a3-cell="true"]');
        if(klimatA1A3Cell){
          const val = parseNumberLike(klimatA1A3Cell.textContent);
          if(Number.isFinite(val)){
            sumKlimatA1A3 += val;
            countKlimatA1A3++;
          }
        }
        
        const klimatA4Cell = childTr.querySelector('td[data-klimat-a4-cell="true"]');
        if(klimatA4Cell){
          const val = parseNumberLike(klimatA4Cell.textContent);
          if(Number.isFinite(val)){
            sumKlimatA4 += val;
            countKlimatA4++;
          }
        }
        
        const klimatA5Cell = childTr.querySelector('td[data-klimat-a5-cell="true"]');
        if(klimatA5Cell){
          const val = parseNumberLike(klimatA5Cell.textContent);
          if(Number.isFinite(val)){
            sumKlimatA5 += val;
            countKlimatA5++;
          }
        }
      });
      
      // Update parent's sum cells
      const parentInbyggdCell = parentTr.querySelector('td[data-sum-inbyggd-vikt="true"]');
      if(parentInbyggdCell){
        parentInbyggdCell.textContent = countInbyggd > 0 ? sumInbyggdVikt.toFixed(2) : '';
      }
      
      const parentInkoptCell = parentTr.querySelector('td[data-sum-inkopt-vikt="true"]');
      if(parentInkoptCell){
        parentInkoptCell.textContent = countInkopt > 0 ? sumInkoptVikt.toFixed(2) : '';
      }
      
      const parentKlimatA1A3Cell = parentTr.querySelector('td[data-sum-klimat-a1a3="true"]');
      if(parentKlimatA1A3Cell){
        parentKlimatA1A3Cell.textContent = countKlimatA1A3 > 0 ? sumKlimatA1A3.toFixed(2) : '';
      }
      
      const parentKlimatA4Cell = parentTr.querySelector('td[data-sum-klimat-a4="true"]');
      if(parentKlimatA4Cell){
        parentKlimatA4Cell.textContent = countKlimatA4 > 0 ? sumKlimatA4.toFixed(2) : '';
      }
      
      const parentKlimatA5Cell = parentTr.querySelector('td[data-sum-klimat-a5="true"]');
      if(parentKlimatA5Cell){
        parentKlimatA5Cell.textContent = countKlimatA5 > 0 ? sumKlimatA5.toFixed(2) : '';
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
    
    // Get all rows to check if columns contain only numbers
    const table = getTable();
    const allRows = table ? Array.from(table.querySelectorAll('tbody tr')) : [];
    
    headers.forEach((h, idx) => {
      // Skip if this is a numeric-only column
      if(isNumericOnlyColumn(allRows, idx)){
        return;
      }
      
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
  
  function isNumericOnlyColumn(rows, columnIndex){
    if(rows.length === 0) return false;
    
    // Check more rows to get a better sample (up to 20 rows or all rows if fewer)
    const sampleRows = rows.slice(0, Math.min(20, rows.length));
    let numericCount = 0;
    let totalCount = 0;
    let uniqueValues = new Set();
    
    for(const row of sampleRows){
      const cells = Array.from(row.children);
      // Make sure we have enough cells and account for action column
      if(cells.length <= columnIndex + 1) continue;
      
      const cell = cells[columnIndex + 1]; // +1 to account for action column
      
      if(cell){
        const cellText = cell.textContent.trim();
        // Remove layer indicators and badges for clean text
        const cleanText = cellText
          .replace(/^\[Skikt \d+\/\d+\]\s*/, '')
          .replace(/\s*\[\d+\s+skikt\]\s*$/, '')
          .replace(/\s*\(\d+\)\s*$/, '')
          .trim();
        
        if(cleanText !== ''){
          totalCount++;
          uniqueValues.add(cleanText);
          const num = parseNumberLike(cleanText);
          if(Number.isFinite(num)){
            numericCount++;
          }
        }
      }
    }
    
    // Additional check: if column has very few unique values and they're all numbers, it's likely numeric-only
    const isLowDiversityNumeric = uniqueValues.size <= 3 && numericCount === totalCount && totalCount > 0;
    
    // Consider it numeric-only if at least 90% of non-empty values are numbers OR if it's low diversity numeric
    const isHighPercentageNumeric = totalCount > 0 && (numericCount / totalCount) >= 0.9;
    
    const result = isHighPercentageNumeric || isLowDiversityNumeric;
    
    // Debug logging
    if(result){
      console.log(`Column ${columnIndex} filtered out - Numeric: ${numericCount}/${totalCount}, Unique values: ${Array.from(uniqueValues).join(', ')}`);
    }
    
    return result;
  }

  function renderTableWithOptionalGrouping(rows){
    if(!rows || rows.length === 0){ output.innerHTML = '<div>Ingen data att visa.</div>'; return; }
    const headers = rows[0];
    lastHeaders = headers; // Store headers for project save/load
    const bodyRows = rows.slice(1);
    const selected = groupBySelect ? groupBySelect.value : '';
    const groupIdx = selected === '' ? -1 : parseInt(selected, 10);

    if(groupIdx === -1 || Number.isNaN(groupIdx)){
      const table = document.createElement('table');
      const thead = document.createElement('thead');
      const headerTr = document.createElement('tr');
      const actionTh = document.createElement('th'); actionTh.textContent = '';
      headerTr.appendChild(actionTh);
      
      // Get existing table headers to preserve dynamically added columns
      const existingTable = getTable();
      let allHeaders = [...headers];
      if(existingTable){
        const existingHeaders = Array.from(existingTable.querySelectorAll('thead tr:first-child th')).map(th => th.textContent);
        // Add any new headers that aren't in the original headers
        const newHeaders = existingHeaders.slice(1); // Skip action column
        newHeaders.forEach(h => {
          if(!headers.includes(h)){
            allHeaders.push(h);
          }
        });
      }
      
      allHeaders.forEach(h => { const th = document.createElement('th'); th.textContent = h; headerTr.appendChild(th); });
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
        // Add cells for original data
        r.forEach(c => { const td = document.createElement('td'); td.textContent = c; tr.appendChild(td); });
        // Add empty cells for any new columns that were added dynamically
        // Including climate columns with proper data attributes
        for(let i = r.length; i < allHeaders.length; i++){
          const headerName = allHeaders[i];
          const td = document.createElement('td');
          td.textContent = '';
          // Mark cells with data attributes so they can be found and updated by applySavedClimate
          if(headerName === 'Klimatresurs'){
            td.setAttribute('data-climate-cell', 'true');
          } else if(headerName === 'Omräkningsfaktor'){
            td.setAttribute('data-factor-cell', 'true');
          } else if(headerName === 'Omräkningsfaktor enhet'){
            td.setAttribute('data-unit-cell', 'true');
          } else if(headerName === 'Spillfaktor'){
            td.setAttribute('data-waste-cell', 'true');
          } else if(headerName === 'Emissionsfaktor A1-A3'){
            td.setAttribute('data-A1_A3-cell', 'true');
          } else if(headerName === 'Emissionsfaktor A4'){
            td.setAttribute('data-A4-cell', 'true');
          } else if(headerName === 'Emissionsfaktor A5'){
            td.setAttribute('data-A5-cell', 'true');
          } else if(headerName === 'Inbyggd vikt'){
            td.setAttribute('data-inbyggd-vikt-cell', 'true');
          } else if(headerName === 'Inköpt vikt'){
            td.setAttribute('data-inkopt-vikt-cell', 'true');
          } else if(headerName === 'Klimatpåverkan A1-A3'){
            td.setAttribute('data-klimat-a1a3-cell', 'true');
          } else if(headerName === 'Klimatpåverkan A4'){
            td.setAttribute('data-klimat-a4-cell', 'true');
          } else if(headerName === 'Klimatpåverkan A5'){
            td.setAttribute('data-klimat-a5-cell', 'true');
          }
          tr.appendChild(td);
        }
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
      
      // Populate group by options after table is created
      if(groupBySelect){ populateGroupBy(headers); }
    } else {
      const table = buildGroupedTable(headers, bodyRows, groupIdx);
      output.innerHTML = ''; output.appendChild(table);
      
      // Populate group by options after table is created
      if(groupBySelect){ populateGroupBy(headers); }
    }
    ensureColumnFilters();
    applyFilters();
    
    // Hide all child rows initially since parents start collapsed
    const table = getTable();
    if(table){
      const tbody = table.querySelector('tbody');
      if(tbody){
        // Hide all group children
        const groupChildren = tbody.querySelectorAll('tr[data-group-child-of]');
        groupChildren.forEach(child => {
          const parentKey = child.getAttribute('data-group-child-of');
          const parent = tbody.querySelector(`tr[data-group-key="${CSS.escape(parentKey)}"]`);
          if(parent && parent.getAttribute('data-open') === 'false'){
            child.style.display = 'none';
          }
        });
        
        // Hide all layer children
        const layerChildren = tbody.querySelectorAll('tr[data-parent-key]');
        layerChildren.forEach(child => {
          const parentKey = child.getAttribute('data-parent-key');
          const parent = tbody.querySelector(`tr[data-layer-key="${CSS.escape(parentKey)}"]`);
          if(parent && parent.getAttribute('data-open') === 'false'){
            child.style.display = 'none';
          }
        });
      }
    }
    
    // Update climate summary after rendering
    setTimeout(() => updateClimateSummary(), 100);
    
    // Save initial state after table is rendered (but not during restore)
    if(!isRestoringState && undoStack.length === 0){
      setTimeout(() => saveState(), 200);
    }
  }

  function handleFile(file){
    originalFileName = file.name; // Store original file name
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
  
  if(exportBtn){
    exportBtn.addEventListener('click', function(){
      exportTableToExcel();
    });
  }
  
  // Add row button
  const addRowBtn = document.getElementById('addRowBtn');
  if(addRowBtn){
    addRowBtn.addEventListener('click', function(){
      addNewRow();
    });
  }
  
  // Undo/Redo buttons
  if(undoBtn){
    undoBtn.addEventListener('click', performUndo);
  }
  
  if(redoBtn){
    redoBtn.addEventListener('click', performRedo);
  }
  
  // Keyboard shortcuts for undo/redo
  document.addEventListener('keydown', function(e){
    // Ctrl+Z for undo (or Cmd+Z on Mac)
    if((e.ctrlKey || e.metaKey) && e.key === 'z' && !e.shiftKey){
      e.preventDefault();
      performUndo();
    }
    // Ctrl+Y for redo (or Cmd+Y on Mac, or Ctrl+Shift+Z)
    else if((e.ctrlKey || e.metaKey) && (e.key === 'y' || (e.key === 'z' && e.shiftKey))){
      e.preventDefault();
      performRedo();
    }
  });
  
  function addNewRow(){
    const table = getTable();
    if(!table) {
      alert('Ladda en tabell först');
      return;
    }
    
    // Save state before adding new row
    saveState();
    
    const tbody = table.querySelector('tbody');
    const thead = table.querySelector('thead');
    if(!tbody || !thead) return;
    
    // Get headers
    const headers = Array.from(thead.querySelectorAll('th')).map(th => th.textContent);
    
    // Filter out climate-related headers - new rows should not have these initially
    const climateHeaders = ['Klimatresurs', 'Omräkningsfaktor', 'Omräkningsfaktor enhet', 'Spillfaktor', 
                           'Emissionsfaktor A1-A3', 'Emissionsfaktor A4', 'Emissionsfaktor A5',
                           'Inbyggd vikt', 'Inköpt vikt', 'Klimatpåverkan A1-A3', 'Klimatpåverkan A4', 'Klimatpåverkan A5'];
    const baseHeaders = headers.filter(h => !climateHeaders.includes(h) && h.trim() !== '');
    
    // Get all existing values for autocomplete (only from base headers)
    const columnValues = new Map();
    baseHeaders.forEach((header, idx) => {
      columnValues.set(idx, new Set());
    });
    
    // Collect unique values from each column
    Array.from(tbody.querySelectorAll('tr')).forEach(tr => {
      // Skip parent rows
      if(tr.classList.contains('group-parent') || tr.classList.contains('layer-parent')) return;
      
      const dataCells = Array.from(tr.querySelectorAll('td')).slice(0, baseHeaders.length + 1); // +1 for action
      dataCells.forEach((td, idx) => {
        if(idx === 0) return; // Skip action column
        const value = td.textContent.trim();
        if(value){
          columnValues.get(idx - 1)?.add(value);
        }
      });
    });
    
    // Create new row
    const newTr = document.createElement('tr');
    newTr.classList.add('is-new');
    
    // Action cell with save button
    const actionTd = document.createElement('td');
    const saveBtn = document.createElement('button');
    saveBtn.type = 'button';
    saveBtn.textContent = 'Spara';
    saveBtn.style.background = '#4caf50';
    saveBtn.style.color = 'white';
    saveBtn.addEventListener('click', function(){
      finalizeNewRow(newTr);
    });
    actionTd.appendChild(saveBtn);
    
    const cancelBtn = document.createElement('button');
    cancelBtn.type = 'button';
    cancelBtn.textContent = 'Avbryt';
    cancelBtn.addEventListener('click', function(){
      newTr.remove();
    });
    actionTd.appendChild(cancelBtn);
    
    newTr.appendChild(actionTd);
    
    // Create editable cells ONLY for base columns (not climate columns)
    baseHeaders.forEach((header, idx) => {
      const td = document.createElement('td');
      td.classList.add('editable');
      td.textContent = '';
      
      // Make cell clickable to edit
      if(td.classList.contains('editable')){
        td.addEventListener('click', function(){
          if(td.querySelector('input')) return; // Already editing
        
        const currentValue = td.textContent;
        const input = document.createElement('input');
        input.type = 'text';
        input.value = currentValue;
        
        // Create autocomplete list
        const autocompleteDiv = document.createElement('div');
        autocompleteDiv.className = 'autocomplete-list';
        autocompleteDiv.style.display = 'none';
        
        // Show autocomplete on input
        input.addEventListener('input', function(){
          const searchTerm = input.value.toLowerCase();
          const values = Array.from(columnValues.get(idx) || []);
          const filtered = values.filter(v => v.toLowerCase().includes(searchTerm));
          
          if(filtered.length > 0 && searchTerm){
            autocompleteDiv.innerHTML = '';
            filtered.slice(0, 10).forEach(value => {
              const item = document.createElement('div');
              item.className = 'autocomplete-item';
              item.textContent = value;
              item.addEventListener('click', function(){
                input.value = value;
                autocompleteDiv.style.display = 'none';
              });
              autocompleteDiv.appendChild(item);
            });
            autocompleteDiv.style.display = 'block';
          } else {
            autocompleteDiv.style.display = 'none';
          }
        });
        
        // Save on Enter
        input.addEventListener('keydown', function(e){
          if(e.key === 'Enter'){
            td.textContent = input.value;
            td.innerHTML = td.textContent; // Remove input
            autocompleteDiv.remove();
          } else if(e.key === 'Escape'){
            td.innerHTML = currentValue;
            autocompleteDiv.remove();
          }
        });
        
        // Save on blur
        input.addEventListener('blur', function(){
          setTimeout(() => {
            td.textContent = input.value;
            td.innerHTML = td.textContent; // Remove input
            autocompleteDiv.remove();
          }, 200); // Delay to allow clicking autocomplete
        });
        
        td.innerHTML = '';
        td.appendChild(input);
        td.appendChild(autocompleteDiv);
        input.focus();
        });
      }
      
      newTr.appendChild(td);
    });
    
    // Insert at top of tbody
    tbody.insertBefore(newTr, tbody.firstChild);
  }
  
  function finalizeNewRow(tr){
    // Remove editable class and convert to normal row
    tr.classList.remove('is-new');
    
    // Get all cell values
    const cells = Array.from(tr.querySelectorAll('td.editable'));
    const rowData = cells.map(td => td.textContent.trim());
    
    // Store as original row data
    tr._originalRowData = rowData;
    
    // Add to lastRows so it persists when re-grouping
    if(lastRows){
      lastRows.push(rowData);
      console.log('✅ Added to lastRows, new count:', lastRows.length);
    }
    
    // Replace action buttons with standard ones
    const actionTd = tr.querySelector('td:first-child');
    if(actionTd){
      actionTd.innerHTML = '';
      
      const layerBtn = document.createElement('button');
      layerBtn.type = 'button';
      layerBtn.textContent = 'Skikta';
      layerBtn.addEventListener('click', function(ev){ 
        ev.stopPropagation(); 
        openLayerModal({ type: 'row', rowEl: tr }); 
      });
      actionTd.appendChild(layerBtn);
      
      const climateBtn = document.createElement('button');
      climateBtn.type = 'button';
      climateBtn.textContent = 'Mappa klimatresurs';
      climateBtn.addEventListener('click', function(ev){ 
        ev.stopPropagation(); 
        openClimateModal({ type: 'row', rowEl: tr }); 
      });
      actionTd.appendChild(climateBtn);
    }
    
    // Remove editable class from cells
    cells.forEach(td => td.classList.remove('editable'));
    
    console.log('✅ New row saved:', rowData);
  }
  
  function exportTableToExcel(){
    const table = getTable();
    if(!table){
      alert('Ingen tabell att exportera');
      return;
    }
    
    if(!window.XLSX){
      alert('Excel-export biblioteket kunde inte laddas');
      return;
    }
    
    // Create a workbook
    const wb = window.XLSX.utils.book_new();
    
    // Collect all visible data
    const exportData = [];
    
    // Get headers
    const thead = table.querySelector('thead');
    const headerRow = thead ? thead.querySelector('tr:first-child') : null;
    if(headerRow){
      const headers = [];
      Array.from(headerRow.children).forEach((th, index) => {
        // Skip first column (action buttons)
        if(index === 0) return;
        headers.push(th.textContent);
      });
      exportData.push(headers);
    }
    
    // Get all visible rows (not hidden by filters)
    const tbody = table.querySelector('tbody');
    if(tbody){
      const rows = Array.from(tbody.querySelectorAll('tr'));
      rows.forEach(tr => {
        // Skip hidden rows
        if(tr.style.display === 'none') return;
        
        const rowData = [];
        const cells = Array.from(tr.children);
        cells.forEach((td, index) => {
          // Skip first column (action buttons)
          if(index === 0) return;
          
          // Get clean text content
          let text = td.textContent.trim();
          
          // Try to parse as number for better Excel formatting
          const num = parseNumberLike(text);
          if(Number.isFinite(num)){
            rowData.push(num);
          } else {
            rowData.push(text);
          }
        });
        
        if(rowData.length > 0){
          exportData.push(rowData);
        }
      });
    }
    
    // Create worksheet from data
    const ws = window.XLSX.utils.aoa_to_sheet(exportData);
    
    // Add worksheet to workbook
    window.XLSX.utils.book_append_sheet(wb, ws, 'Data');
    
    // Generate filename with timestamp
    const date = new Date();
    const timestamp = date.toISOString().slice(0, 19).replace(/:/g, '-');
    const filename = `export_${timestamp}.xlsx`;
    
    // Write and download
    window.XLSX.writeFile(wb, filename);
  }

  // Layer modal behavior
  function openLayerModal(target){
    layerTarget = target;
    
    // Pre-fill with existing values if editing an already layered item
    if(target.type === 'row' && target.rowEl){
      // For a single row, check if it's a layer parent
      const layerKey = target.rowEl.getAttribute('data-layer-key');
      if(layerKey){
        // This row is already layered, get its layer data
        const rowData = target.rowEl._originalRowData || getRowDataFromTr(target.rowEl);
        if(rowData){
          const layerChildOf = target.rowEl.getAttribute('data-layer-child-of');
          const signature = getRowSignature(rowData, layerChildOf);
          const saved = layerData.get(signature);
          if(saved){
            if(layerCountInput) layerCountInput.value = saved.count;
            if(layerThicknessesInput) layerThicknessesInput.value = saved.thicknesses.join(', ');
            // Load existing layer names and climate resources
            if(saved.layerNames && saved.climateResources){
              loadExistingLayerData(saved.layerNames, saved.climateResources);
            }
          }
        }
      } else {
        // Not yet layered, clear the inputs
        if(layerCountInput) layerCountInput.value = '2';
        if(layerThicknessesInput) layerThicknessesInput.value = '';
        updateLayerNamesContainer();
      }
    } else if(target.type === 'group' && target.key){
      // Find existing layer data for this group
      const table = getTable();
      if(table){
        const tbody = table.querySelector('tbody');
        if(tbody){
          // Find a child row to get the layer data
          const firstChild = tbody.querySelector(`tr[data-layer-child-of="${CSS.escape(target.key)}"]`);
          if(firstChild){
            const rowData = firstChild._originalRowData || getRowDataFromTr(firstChild);
            if(rowData){
              const signature = getRowSignature(rowData, target.key);
              const saved = layerData.get(signature);
              if(saved){
                if(layerCountInput) layerCountInput.value = saved.count;
                if(layerThicknessesInput) layerThicknessesInput.value = saved.thicknesses.join(', ');
                // Load existing layer names and climate resources
                if(saved.layerNames && saved.climateResources){
                  loadExistingLayerData(saved.layerNames, saved.climateResources);
                }
              }
            }
          } else {
            // Not yet layered, clear the inputs
            if(layerCountInput) layerCountInput.value = '2';
            if(layerThicknessesInput) layerThicknessesInput.value = '';
            updateLayerNamesContainer();
          }
        }
      }
    }
    
    // Initialize mixed layer checkboxes
    updateMixedLayerCheckboxes();
    updateMixedLayerDetails();
    updateLayerNamesContainer();
    
    if(layerModal){ layerModal.style.display = 'flex'; }
  }
  function closeLayerModal(){
    layerTarget = null;
    // Reset mixed layer controls
    if(mixedLayerCheckboxes){ mixedLayerCheckboxes.innerHTML = ''; }
    if(mixedLayerDetails){ mixedLayerDetails.style.display = 'none'; }
    if(mixedLayerConfigs){ mixedLayerConfigs.innerHTML = ''; }
    // Clear layer names container
    if(layerNamesContainer){ layerNamesContainer.innerHTML = ''; }
    if(layerModal){ layerModal.style.display = 'none'; }
  }
  if(layerCancelBtn){ layerCancelBtn.addEventListener('click', closeLayerModal); }
  if(layerModal){ layerModal.addEventListener('click', function(e){ if(e.target === layerModal) closeLayerModal(); }); }
  
  // Update mixed layer checkboxes when layer count changes
  if(layerCountInput && mixedLayerCheckboxes){
    layerCountInput.addEventListener('input', function(){
      updateMixedLayerCheckboxes();
      updateLayerNamesContainer();
    });
  }
  
  // Update mixed layer checkboxes based on layer count
  function updateMixedLayerCheckboxes(){
    if(!mixedLayerCheckboxes || !layerCountInput) return;
    
    const count = Math.max(1, parseInt(layerCountInput.value || '1', 10));
    mixedLayerCheckboxes.innerHTML = '';
    
    for(let i = 1; i <= count; i++){
      const checkboxDiv = document.createElement('div');
      checkboxDiv.style.cssText = 'display:flex; align-items:center; gap:8px; padding:8px; border:1px solid #ddd; border-radius:4px; background:white;';
      
      const checkbox = document.createElement('input');
      checkbox.type = 'checkbox';
      checkbox.id = `mixedLayer${i}`;
      checkbox.addEventListener('change', function(){
        updateMixedLayerDetails();
        updateLayerNamesContainer();
      });
      
      const label = document.createElement('label');
      label.htmlFor = `mixedLayer${i}`;
      label.textContent = `Skikt ${i} är blandning`;
      label.style.cssText = 'cursor:pointer; margin:0; font-weight:500;';
      
      checkboxDiv.appendChild(checkbox);
      checkboxDiv.appendChild(label);
      mixedLayerCheckboxes.appendChild(checkboxDiv);
    }
  }
  
  // Update mixed layer details section
  function updateMixedLayerDetails(){
    if(!mixedLayerDetails || !mixedLayerConfigs) return;
    
    const count = Math.max(1, parseInt(layerCountInput.value || '1', 10));
    const mixedLayers = [];
    
    // Find which layers are marked as mixed
    for(let i = 1; i <= count; i++){
      const checkbox = document.getElementById(`mixedLayer${i}`);
      if(checkbox && checkbox.checked){
        mixedLayers.push(i);
      }
    }
    
    if(mixedLayers.length === 0){
      mixedLayerDetails.style.display = 'none';
      return;
    }
    
    mixedLayerDetails.style.display = 'block';
    mixedLayerConfigs.innerHTML = '';
    
    mixedLayers.forEach(layerNum => {
      const configDiv = document.createElement('div');
      configDiv.style.cssText = 'margin-bottom:16px; padding:12px; border:1px solid #ddd; border-radius:4px; background:white;';
      
      configDiv.innerHTML = `
        <h4 style="margin:0 0 12px 0; font-size:0.9rem; color:#333;">Skikt ${layerNum} - Blandat material</h4>
        
        <div style="display:grid; grid-template-columns: 1fr 1fr; gap:12px; margin-bottom:12px;">
          <div>
            <label style="display:block; margin-bottom:4px; font-size:0.85rem; font-weight:600;">Material 1:</label>
            <input type="text" id="mixedMat1Name${layerNum}" placeholder="t.ex. Betong C30/37" style="width:100%; padding:6px; border:1px solid #ddd; border-radius:4px; font-size:0.85rem;" />
          </div>
          <div>
            <label style="display:block; margin-bottom:4px; font-size:0.85rem; font-weight:600;">Andel (%):</label>
            <input type="number" id="mixedMat1Percent${layerNum}" min="0" max="100" value="50" style="width:100%; padding:6px; border:1px solid #ddd; border-radius:4px; font-size:0.85rem;" />
          </div>
        </div>
        
        <div style="display:grid; grid-template-columns: 1fr 1fr; gap:12px; margin-bottom:12px;">
          <div>
            <label style="display:block; margin-bottom:4px; font-size:0.85rem; font-weight:600;">Material 2:</label>
            <input type="text" id="mixedMat2Name${layerNum}" placeholder="t.ex. Stål" style="width:100%; padding:6px; border:1px solid #ddd; border-radius:4px; font-size:0.85rem;" />
          </div>
          <div>
            <label style="display:block; margin-bottom:4px; font-size:0.85rem; font-weight:600;">Andel (%):</label>
            <input type="number" id="mixedMat2Percent${layerNum}" min="0" max="100" value="50" style="width:100%; padding:6px; border:1px solid #ddd; border-radius:4px; font-size:0.85rem;" />
          </div>
        </div>
        
        <div style="display:grid; grid-template-columns: 1fr 1fr; gap:12px;">
          <div>
            <label style="display:block; margin-bottom:4px; font-size:0.85rem; font-weight:600;">Klimatresurs Material 1:</label>
            <select id="mixedMat1Climate${layerNum}" style="width:100%; padding:6px; border:1px solid #ddd; border-radius:4px; font-size:0.85rem;">
              <option value="">Välj resurs...</option>
            </select>
          </div>
          <div>
            <label style="display:block; margin-bottom:4px; font-size:0.85rem; font-weight:600;">Klimatresurs Material 2:</label>
            <select id="mixedMat2Climate${layerNum}" style="width:100%; padding:6px; border:1px solid #ddd; border-radius:4px; font-size:0.85rem;">
              <option value="">Välj resurs...</option>
            </select>
          </div>
        </div>
      `;
      
      // Populate climate resource dropdowns
      const mat1Climate = configDiv.querySelector(`#mixedMat1Climate${layerNum}`);
      const mat2Climate = configDiv.querySelector(`#mixedMat2Climate${layerNum}`);
      
      if(mat1Climate && mat2Climate && window.climateResources){
        window.climateResources.forEach((resource, index) => {
          const option1 = document.createElement('option');
          option1.value = index;
          option1.textContent = resource.Name || 'Namnlös resurs';
          mat1Climate.appendChild(option1);
          
          const option2 = document.createElement('option');
          option2.value = index;
          option2.textContent = resource.Name || 'Namnlös resurs';
          mat2Climate.appendChild(option2);
        });
      }
      
      // Auto-adjust percentages
      const mat1Percent = configDiv.querySelector(`#mixedMat1Percent${layerNum}`);
      const mat2Percent = configDiv.querySelector(`#mixedMat2Percent${layerNum}`);
      
      if(mat1Percent && mat2Percent){
        mat1Percent.addEventListener('input', function(){
          const val1 = parseFloat(this.value) || 0;
          const val2 = 100 - val1;
          if(val2 >= 0 && val2 <= 100){
            mat2Percent.value = val2.toString();
          }
        });
        
        mat2Percent.addEventListener('input', function(){
          const val2 = parseFloat(this.value) || 0;
          const val1 = 100 - val2;
          if(val1 >= 0 && val1 <= 100){
            mat1Percent.value = val1.toString();
          }
        });
      }
      
      mixedLayerConfigs.appendChild(configDiv);
    });
  }
  
  // Layer names and climate resources functions
  function updateLayerNamesContainer(){
    if(!layerNamesContainer || !layerCountInput) return;
    
    const count = Math.max(1, parseInt(layerCountInput.value || '1', 10));
    layerNamesContainer.innerHTML = '';
    
    for(let i = 1; i <= count; i++){
      // Check if this layer is marked as mixed
      const isMixedLayer = document.getElementById(`mixedLayer${i}`) && document.getElementById(`mixedLayer${i}`).checked;
      
      if(isMixedLayer){
        // Skip mixed layers - they're handled in the mixed layer details section
        continue;
      }
      
      const layerDiv = document.createElement('div');
      layerDiv.style.cssText = 'display:grid; grid-template-columns: 1fr 2fr; gap:12px; margin-bottom:12px; padding:12px; border:1px solid #ddd; border-radius:4px; background:white;';
      
      // Regular layer
      layerDiv.innerHTML = `
        <div>
          <label style="display:block; margin-bottom:4px; font-weight:600;">Skikt ${i} namn:</label>
          <input type="text" id="layerName${i}" placeholder="t.ex. Betong C30/37" style="width:100%; padding:6px; border:1px solid #ddd; border-radius:4px;" />
        </div>
        <div>
          <label style="display:block; margin-bottom:4px; font-weight:600;">Klimatresurs:</label>
          <select id="layerClimate${i}" style="width:100%; padding:6px; border:1px solid #ddd; border-radius:4px;">
            <option value="">Välj resurs...</option>
          </select>
        </div>
      `;
      
      layerNamesContainer.appendChild(layerDiv);
      
      // Populate climate resources dropdown
      const climateSelect = layerDiv.querySelector(`#layerClimate${i}`);
      if(climateSelect && window.climateResources){
        window.climateResources.forEach((resource, index) => {
          const option = document.createElement('option');
          option.value = index;
          option.textContent = resource.Name || 'Namnlös resurs';
          climateSelect.appendChild(option);
        });
      }
    }
  }
  
  function loadExistingLayerData(layerNames, climateResources){
    if(!layerNamesContainer) return;
    
    // Clear and rebuild container
    updateLayerNamesContainer();
    
    // Fill in existing data
    const count = Math.max(1, parseInt(layerCountInput.value || '1', 10));
    
    for(let i = 1; i <= count; i++){
      // Check if this layer is marked as mixed
      const isMixedLayer = document.getElementById(`mixedLayer${i}`) && document.getElementById(`mixedLayer${i}`).checked;
      
      if(isMixedLayer){
        // Mixed layer - fill material inputs from mixed layer details
        const material1Name = document.getElementById(`mixedMat1Name${i}`);
        const material2Name = document.getElementById(`mixedMat2Name${i}`);
        const material1Climate = document.getElementById(`mixedMat1Climate${i}`);
        const material2Climate = document.getElementById(`mixedMat2Climate${i}`);
        
        if(material1Name && layerNames[i-1]){
          material1Name.value = layerNames[i-1];
        }
        if(material2Name && layerNames[i-1]){
          material2Name.value = layerNames[i-1];
        }
        if(material1Climate && climateResources[i-1] !== undefined){
          material1Climate.value = climateResources[i-1];
        }
        if(material2Climate && climateResources[i-1] !== undefined){
          material2Climate.value = climateResources[i-1];
        }
      } else {
        // Regular layer - fill normal inputs
        const nameInput = document.getElementById(`layerName${i}`);
        const climateSelect = document.getElementById(`layerClimate${i}`);
        
        if(nameInput && layerNames[i-1]){
          nameInput.value = layerNames[i-1];
        }
        
        if(climateSelect && climateResources[i-1] !== undefined){
          climateSelect.value = climateResources[i-1];
        }
      }
    }
  }
  
  function getLayerNamesAndClimateResources(){
    if(!layerNamesContainer) return { layerNames: [], climateResources: [] };
    
    const layerNames = [];
    const climateResources = [];
    
    const count = Math.max(1, parseInt(layerCountInput.value || '1', 10));
    
    console.log('🔍 [getLayerNamesAndClimateResources] Getting data for', count, 'layers');
    
    for(let i = 1; i <= count; i++){
      // Check if this layer is marked as mixed
      const isMixedLayer = document.getElementById(`mixedLayer${i}`) && document.getElementById(`mixedLayer${i}`).checked;
      
      if(isMixedLayer){
        // Mixed layer - get data from mixed layer details
        const mat1Name = document.getElementById(`mixedMat1Name${i}`);
        const mat2Name = document.getElementById(`mixedMat2Name${i}`);
        const mat1Climate = document.getElementById(`mixedMat1Climate${i}`);
        const mat2Climate = document.getElementById(`mixedMat2Climate${i}`);
        
        const mat1NameValue = mat1Name ? mat1Name.value.trim() : '';
        const mat2NameValue = mat2Name ? mat2Name.value.trim() : '';
        const mat1ClimateValue = mat1Climate ? mat1Climate.value : '';
        const mat2ClimateValue = mat2Climate ? mat2Climate.value : '';
        
        // For mixed layer, we use the first material's name and climate as the base
        layerNames.push(mat1NameValue || `Skikt ${i}`);
        climateResources.push(mat1ClimateValue);
        
        console.log(`🔍 [getLayerNamesAndClimateResources] Mixed Layer ${i}: mat1="${mat1NameValue}" (${mat1ClimateValue}), mat2="${mat2NameValue}" (${mat2ClimateValue})`);
      } else {
        // Regular layer
        const nameInput = document.getElementById(`layerName${i}`);
        const climateSelect = document.getElementById(`layerClimate${i}`);
        
        const layerName = nameInput ? nameInput.value.trim() : '';
        const climateResource = climateSelect ? climateSelect.value : '';
        
        layerNames.push(layerName);
        climateResources.push(climateResource);
        
        console.log(`🔍 [getLayerNamesAndClimateResources] Layer ${i}: name="${layerName}", climate="${climateResource}"`);
      }
    }
    
    console.log('🔍 [getLayerNamesAndClimateResources] Final data:', { layerNames, climateResources });
    return { layerNames, climateResources };
  }
  
  
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
      
      // Check for mixed layers using new structure
      const mixedLayerConfigs = [];
      const mixedLayers = [];
      
      // Find which layers are marked as mixed
      for(let i = 1; i <= count; i++){
        const checkbox = document.getElementById(`mixedLayer${i}`);
        if(checkbox && checkbox.checked){
          mixedLayers.push(i);
        }
      }
      
      // Create config for each mixed layer
      mixedLayers.forEach(layerIndex => {
        const mat1Name = document.getElementById(`mixedMat1Name${layerIndex}`);
        const mat2Name = document.getElementById(`mixedMat2Name${layerIndex}`);
        const mat1Percent = document.getElementById(`mixedMat1Percent${layerIndex}`);
        const mat2Percent = document.getElementById(`mixedMat2Percent${layerIndex}`);
        const mat1Climate = document.getElementById(`mixedMat1Climate${layerIndex}`);
        const mat2Climate = document.getElementById(`mixedMat2Climate${layerIndex}`);
        
        const mat1NameValue = mat1Name ? mat1Name.value.trim() : '';
        const mat2NameValue = mat2Name ? mat2Name.value.trim() : '';
        const mat1PercentValue = parseFloat(mat1Percent ? mat1Percent.value : '50');
        const mat2PercentValue = parseFloat(mat2Percent ? mat2Percent.value : '50');
        const mat1ClimateValue = mat1Climate ? mat1Climate.value : '';
        const mat2ClimateValue = mat2Climate ? mat2Climate.value : '';
        
        if(mat1NameValue && mat2NameValue){
          mixedLayerConfigs.push({
            layerIndex: layerIndex,
            material1: { 
              name: mat1NameValue, 
              percent: mat1PercentValue,
              climateResource: mat1ClimateValue
            },
            material2: { 
              name: mat2NameValue, 
              percent: mat2PercentValue,
              climateResource: mat2ClimateValue
            }
          });
        }
      });
      
      // Save state before layering
      saveState();
      
      // Get layer names and climate resources
      const { layerNames, climateResources } = getLayerNamesAndClimateResources();
      
      applyLayerSplit(count, thicknesses, mixedLayerConfigs, layerNames, climateResources);
      closeLayerModal();
    });
  }
  if(layerMapClimateBtn){
    layerMapClimateBtn.addEventListener('click', function(){
      console.log('🔧 [LayerMapClimate] Button clicked');
      console.log('🔧 [LayerMapClimate] layerTarget:', layerTarget);
      
      // Save layerTarget before closing modal (which nulls it)
      const savedTarget = layerTarget;
      
      // First apply the layering
      const count = Math.max(1, parseInt(layerCountInput && layerCountInput.value || '1', 10));
      const raw = (layerThicknessesInput && layerThicknessesInput.value || '').trim();
      const thicknesses = raw ? raw.split(',').map(s => parseFloat(s.trim().replace(',', '.'))).filter(n => Number.isFinite(n) && n > 0) : [];
      
      // Check for mixed layers using new structure
      const mixedLayerConfigs = [];
      const mixedLayers = [];
      
      // Find which layers are marked as mixed
      for(let i = 1; i <= count; i++){
        const checkbox = document.getElementById(`mixedLayer${i}`);
        if(checkbox && checkbox.checked){
          mixedLayers.push(i);
        }
      }
      
      // Create config for each mixed layer
      mixedLayers.forEach(layerIndex => {
        const mat1Name = document.getElementById(`mixedMat1Name${layerIndex}`);
        const mat2Name = document.getElementById(`mixedMat2Name${layerIndex}`);
        const mat1Percent = document.getElementById(`mixedMat1Percent${layerIndex}`);
        const mat2Percent = document.getElementById(`mixedMat2Percent${layerIndex}`);
        const mat1Climate = document.getElementById(`mixedMat1Climate${layerIndex}`);
        const mat2Climate = document.getElementById(`mixedMat2Climate${layerIndex}`);
        
        const mat1NameValue = mat1Name ? mat1Name.value.trim() : '';
        const mat2NameValue = mat2Name ? mat2Name.value.trim() : '';
        const mat1PercentValue = parseFloat(mat1Percent ? mat1Percent.value : '50');
        const mat2PercentValue = parseFloat(mat2Percent ? mat2Percent.value : '50');
        const mat1ClimateValue = mat1Climate ? mat1Climate.value : '';
        const mat2ClimateValue = mat2Climate ? mat2Climate.value : '';
        
        if(mat1NameValue && mat2NameValue){
          mixedLayerConfigs.push({
            layerIndex: layerIndex,
            material1: { 
              name: mat1NameValue, 
              percent: mat1PercentValue,
              climateResource: mat1ClimateValue
            },
            material2: { 
              name: mat2NameValue, 
              percent: mat2PercentValue,
              climateResource: mat2ClimateValue
            }
          });
        }
      });
      
      // Save state before layering
      saveState();
      
      console.log('🔧 [LayerMapClimate] Applying layers - count:', count, 'thicknesses:', thicknesses, 'mixedLayerConfig:', mixedLayerConfig);
      
      // Get layer names and climate resources
      const { layerNames, climateResources } = getLayerNamesAndClimateResources();
      
      applyLayerSplit(count, thicknesses, mixedLayerConfigs, layerNames, climateResources);
      closeLayerModal();
      
      // Then open climate modal for the layers using saved target
      if(savedTarget && savedTarget.type === 'group' && savedTarget.key){
        console.log('🔧 [LayerMapClimate] Opening multi-layer climate modal for group:', savedTarget.key);
        setTimeout(() => {
          console.log('🔧 [LayerMapClimate] Timeout executed, opening multi-layer climate modal');
          openMultiLayerClimateModal(savedTarget.key);
        }, 100);
      } else {
        console.log('❌ [LayerMapClimate] Cannot open climate modal - savedTarget:', savedTarget);
      }
    });
  }
  
  // Multi-layer climate modal functions
  function openMultiLayerClimateModal(groupKey){
    const table = getTable();
    if(!table) return;
    const tbody = table.querySelector('tbody');
    if(!tbody) return;
    
    // Get all layer rows for this group
    // These are rows that have data-group-child-of matching the groupKey
    // AND have a data-layer-key (which identifies their layer number)
    const layerRows = Array.from(tbody.querySelectorAll(`tr[data-group-child-of="${CSS.escape(groupKey)}"][data-layer-key]`));
    
    console.log('🔍 [openMultiLayerClimate] GroupKey:', groupKey, 'Found rows:', layerRows.length);
    
    if(layerRows.length === 0){
      console.log('❌ No layer rows found for group:', groupKey);
      return;
    }
    
    // Group layer rows by their layer key to find unique layers
    const layerKeySet = new Set();
    layerRows.forEach(row => {
      const layerKey = row.dataset.layerKey || '';
      if(layerKey){
        layerKeySet.add(layerKey);
      }
    });
    
    // Extract layer numbers from keys (e.g., "Wall_Layer_1" -> 1)
    const uniqueLayers = Array.from(layerKeySet)
      .map(key => {
        const match = key.match(/_Layer_(\d+)$/);
        return match ? parseInt(match[1], 10) : null;
      })
      .filter(num => num !== null)
      .sort((a, b) => a - b);
    
    console.log('🔍 [openMultiLayerClimate] Unique layers:', uniqueLayers);
    
    if(uniqueLayers.length === 0){
      console.log('❌ No valid layer numbers found');
      return;
    }
    
    // Get thickness information for each layer
    const layerInfo = uniqueLayers.map(layerNum => {
      const layerKeyPattern = groupKey + '_Layer_' + layerNum;
      const layerRow = layerRows.find(row => row.dataset.layerKey === layerKeyPattern);
      
      let thickness = null;
      let existingResource = null;
      let existingLayerName = null;
      
      if(layerRow){
        // Try to read thickness from the row
        const table = getTable();
        if(table){
          const headers = Array.from(table.querySelectorAll('thead th')).map(th => th.textContent);
          const thicknessIdx = headers.findIndex(h => String(h).toLowerCase() === 'thickness');
          if(thicknessIdx >= 0){
            const cells = Array.from(layerRow.children);
            const thicknessCell = cells[thicknessIdx + 1]; // +1 for action column
            if(thicknessCell){
              const value = parseNumberLike(thicknessCell.textContent);
              if(Number.isFinite(value)){
                thickness = value * 1000; // Convert m to mm for display
              }
            }
          }
        }
        
        // Check for existing climate resource
        const climateCell = layerRow.querySelector('td[data-climate-cell="true"]');
        if(climateCell && climateCell.textContent){
          existingResource = climateCell.textContent;
        }
        
        // Check for existing layer name in badge
        const firstDataCell = layerRow.querySelector('td:nth-child(2)');
        if(firstDataCell){
          const badge = firstDataCell.querySelector('.badge-new');
          if(badge && badge.textContent){
            // Extract name from badge text like "Puts" or "Skikt 1/3"
            const badgeText = badge.textContent;
            // Only extract if it doesn't look like a generic "Skikt X/Y" pattern
            if(!badgeText.match(/^Skikt \d+\/\d+$/)){
              existingLayerName = badgeText;
            }
          }
        }
      }
      
      return {
        layerNum,
        layerKeyPattern,
        thickness,
        existingResource,
        existingLayerName
      };
    });
    
    multiLayerClimateTarget = { 
      layerRows, 
      groupKey, 
      uniqueLayers,
      layerInfo,
      selectedResources: new Map() // layerNumber -> resourceIndex
    };
    
    showAllLayersSelection();
  }
  
  function showAllLayersSelection(){
    if(!multiLayerClimateTarget) return;
    
    const { layerInfo } = multiLayerClimateTarget;
    
    // Clear existing content
    multiLayerClimateContent.innerHTML = '';
    
    // Create a card for each layer
    layerInfo.forEach(info => {
      const card = document.createElement('div');
      card.style.cssText = 'border:2px solid #2196f3; padding:16px; border-radius:8px; background:#f5f5f5;';
      
      // Layer header with thickness
      const header = document.createElement('div');
      header.style.cssText = 'display:flex; justify-content:space-between; align-items:center; margin-bottom:12px;';
      
      const layerTitle = document.createElement('h3');
      layerTitle.style.cssText = 'margin:0; color:#1565c0; font-size:16px;';
      layerTitle.textContent = `Skikt ${info.layerNum}`;
      
      const thicknessLabel = document.createElement('span');
      thicknessLabel.style.cssText = 'background:#2196f3; color:white; padding:4px 12px; border-radius:12px; font-size:13px; font-weight:600;';
      thicknessLabel.textContent = info.thickness ? `${info.thickness.toFixed(1)} mm` : 'Okänd tjocklek';
      
      header.appendChild(layerTitle);
      header.appendChild(thicknessLabel);
      card.appendChild(header);
      
      // Layer name input
      const nameLabel = document.createElement('label');
      nameLabel.style.cssText = 'display:block; margin-bottom:12px;';
      nameLabel.innerHTML = '<span style="font-weight:600; font-size:13px; color:#555; display:block; margin-bottom:6px;">Skiktnamn:</span>';
      
      const nameInput = document.createElement('input');
      nameInput.type = 'text';
      nameInput.placeholder = `t.ex. Puts, Betong, Isolering...`;
      nameInput.style.cssText = 'width:100%; padding:8px; border:1px solid #ddd; border-radius:4px; font-size:14px; box-sizing:border-box;';
      nameInput.dataset.layerNum = info.layerNum;
      nameInput.dataset.nameInput = 'true';
      
      // Try to get existing layer name from badge
      if(info.existingLayerName){
        nameInput.value = info.existingLayerName;
      }
      
      nameLabel.appendChild(nameInput);
      card.appendChild(nameLabel);
      
      // Search input for climate resource
      const climateLabel = document.createElement('label');
      climateLabel.style.cssText = 'display:block;';
      climateLabel.innerHTML = '<span style="font-weight:600; font-size:13px; color:#555; display:block; margin-bottom:6px;">Klimatresurs:</span>';
      
      const searchWrapper = document.createElement('div');
      searchWrapper.style.cssText = 'position:relative;';
      
      const searchInput = document.createElement('input');
      searchInput.type = 'text';
      searchInput.placeholder = 'Sök klimatresurs...';
      searchInput.style.cssText = 'width:100%; padding:10px; border:1px solid #ddd; border-radius:4px; font-size:14px; box-sizing:border-box;';
      searchInput.dataset.layerNum = info.layerNum;
      
      // Pre-fill with existing resource if any
      if(info.existingResource){
        searchInput.value = info.existingResource;
      }
      
      searchWrapper.appendChild(searchInput);
      
      // Dropdown for search results
      const dropdown = document.createElement('div');
      dropdown.style.cssText = 'position:absolute; top:100%; left:0; right:0; background:white; border:1px solid #ddd; border-top:none; max-height:200px; overflow-y:auto; display:none; z-index:1000; box-shadow:0 4px 8px rgba(0,0,0,0.1);';
      dropdown.dataset.layerNum = info.layerNum;
      
      searchWrapper.appendChild(dropdown);
      climateLabel.appendChild(searchWrapper);
      card.appendChild(climateLabel);
      
      // Add search functionality
      searchInput.addEventListener('input', function(){
        const query = this.value.toLowerCase();
        dropdown.innerHTML = '';
        
        if(!query){
          dropdown.style.display = 'none';
          return;
        }
        
        // Filter resources
        const matches = window.climateResources.filter(resource => 
          resource.Name && resource.Name.toLowerCase().includes(query)
        ).slice(0, 20); // Limit to 20 results
        
        if(matches.length === 0){
          dropdown.innerHTML = '<div style="padding:12px; color:#999;">Inga resultat</div>';
          dropdown.style.display = 'block';
          return;
        }
        
        matches.forEach((resource, idx) => {
          const item = document.createElement('div');
          item.style.cssText = 'padding:10px; cursor:pointer; border-bottom:1px solid #eee;';
          item.textContent = resource.Name;
          item.dataset.resourceIndex = window.climateResources.indexOf(resource);
          
          item.addEventListener('mouseenter', function(){
            this.style.background = '#e3f2fd';
          });
          
          item.addEventListener('mouseleave', function(){
            this.style.background = 'white';
          });
          
          item.addEventListener('click', function(){
            searchInput.value = resource.Name;
            searchInput.dataset.selectedIndex = this.dataset.resourceIndex;
            dropdown.style.display = 'none';
          });
          
          dropdown.appendChild(item);
        });
        
        dropdown.style.display = 'block';
      });
      
      // Close dropdown when clicking outside
      document.addEventListener('click', function(e){
        if(!searchWrapper.contains(e.target)){
          dropdown.style.display = 'none';
        }
      });
      
      multiLayerClimateContent.appendChild(card);
    });
    
    multiLayerClimateModal.style.display = 'flex';
  }
  
  function showNextLayerSelection(){
    if(!multiLayerClimateTarget) return;
    
    const { uniqueLayers, currentLayerIndex, selectedResources, layerRows, groupKey } = multiLayerClimateTarget;
    
    if(currentLayerIndex >= uniqueLayers.length){
      // All layers have been selected, now apply them all
      applyAllLayerResources();
      return;
    }
    
    const layerNum = uniqueLayers[currentLayerIndex];
    
    // Try to find existing climate data for this layer
    let existingResourceIndex = null;
    const layerKeyPattern = groupKey + '_Layer_' + layerNum;
    const layerRow = layerRows.find(row => row.dataset.layerKey === layerKeyPattern);
    
    if(layerRow){
      // Check if this row has climate data
      const climateCell = layerRow.querySelector('td[data-climate-cell="true"]');
      if(climateCell && climateCell.textContent){
        const resourceName = climateCell.textContent;
        // Find matching resource in climateResources
        if(window.climateResources){
          existingResourceIndex = window.climateResources.findIndex(r => r.Name === resourceName);
        }
      }
    }
    
    // Update modal title
    const modalTitle = multiLayerClimateModal.querySelector('h2');
    if(modalTitle){
      const editText = existingResourceIndex !== null && existingResourceIndex >= 0 ? ' (redigera)' : '';
      modalTitle.textContent = `Välj klimatresurs för Skikt ${layerNum}${editText} (${currentLayerIndex + 1}/${uniqueLayers.length})`;
    }
    
    // Show which layers have been selected
    multiLayerClimateContent.innerHTML = '';
    
    // Show summary of previously selected layers
    if(currentLayerIndex > 0){
      const summaryDiv = document.createElement('div');
      summaryDiv.style.cssText = 'margin-bottom:16px; padding:12px; background:#e8f5e9; border-radius:4px; border:1px solid #4caf50;';
      
      const summaryTitle = document.createElement('div');
      summaryTitle.style.cssText = 'font-weight:600; margin-bottom:8px; color:#2e7d32;';
      summaryTitle.textContent = 'Tidigare valda:';
      summaryDiv.appendChild(summaryTitle);
      
      uniqueLayers.slice(0, currentLayerIndex).forEach(prevLayerNum => {
        const resource = selectedResources.get(prevLayerNum);
        if(resource){
          const item = document.createElement('div');
          item.style.cssText = 'padding:4px 0; color:#1b5e20;';
          item.textContent = `Skikt ${prevLayerNum}: ${resource.Name || 'Namnlös'}`;
          summaryDiv.appendChild(item);
        }
      });
      
      multiLayerClimateContent.appendChild(summaryDiv);
    }
    
    // Current layer selection
    const currentDiv = document.createElement('div');
    currentDiv.style.cssText = 'border:2px solid #2196f3; padding:16px; border-radius:4px; background:#e3f2fd;';
    
    const label = document.createElement('label');
    label.style.cssText = 'display:block; margin-bottom:8px; font-weight:600; font-size:16px; color:#1565c0;';
    label.textContent = `Skikt ${layerNum}`;
    
    const select = document.createElement('select');
    select.id = 'currentLayerSelect';
    select.style.cssText = 'width:100%; padding:10px; border:1px solid #ddd; border-radius:4px; font-size:14px;';
    
    // Populate with climate resources
    const defaultOption = document.createElement('option');
    defaultOption.value = '';
    defaultOption.textContent = 'Välj klimatresurs...';
    select.appendChild(defaultOption);
    
    if(window.climateResources && Array.isArray(window.climateResources)){
      window.climateResources.forEach((resource, resIndex) => {
        const option = document.createElement('option');
        option.value = resIndex;
        option.textContent = resource.Name || 'Namnlös resurs';
        // Pre-select if this is the existing resource
        if(existingResourceIndex !== null && resIndex === existingResourceIndex){
          option.selected = true;
        }
        select.appendChild(option);
      });
    }
    
    currentDiv.appendChild(label);
    currentDiv.appendChild(select);
    multiLayerClimateContent.appendChild(currentDiv);
    
    multiLayerClimateModal.style.display = 'flex';
  }
  
  function updateLayerBadge(row, layerName){
    if(!row || !layerName) return;
    
    const firstDataCell = row.querySelector('td:nth-child(2)');
    if(!firstDataCell) return;
    
    // Find or create badge
    let badge = firstDataCell.querySelector('.badge-new');
    if(!badge){
      // Create new badge
      badge = document.createElement('span');
      badge.className = 'badge-new';
      firstDataCell.insertBefore(badge, firstDataCell.firstChild);
    }
    
    // Update badge text with layer name
    badge.textContent = layerName;
    
    console.log('🏷️ Updated layer badge:', layerName);
  }
  
  function applyAllLayerResources(){
    if(!multiLayerClimateTarget) return;
    
    const { layerRows, groupKey, selectedResources, selectedNames, uniqueLayers } = multiLayerClimateTarget;
    
    // Build a map from layer number to the layer key pattern
    const layerKeyPatterns = new Map();
    uniqueLayers.forEach(layerNum => {
      layerKeyPatterns.set(layerNum, groupKey + '_Layer_' + layerNum);
    });
    
    console.log('🔍 [applyAllLayerResources] Layer key patterns:', Array.from(layerKeyPatterns.entries()));
    
    // Apply the appropriate resource and name to each layer row based on its layer key
    layerRows.forEach(row => {
      const layerKey = row.dataset.layerKey || '';
      if(layerKey){
        // Find which layer number this key corresponds to
        const match = layerKey.match(/_Layer_(\d+)$/);
        if(match){
          const layerNumber = parseInt(match[1], 10);
          
          // Apply layer name if provided
          if(selectedNames && selectedNames.has(layerNumber)){
            const layerName = selectedNames.get(layerNumber);
            updateLayerBadge(row, layerName);
          }
          
          // Apply climate resource if provided
          const resource = selectedResources.get(layerNumber);
          if(resource){
            console.log('🔍 [applyAllLayerResources] Applying resource to layerKey:', layerKey, 'layerNum:', layerNumber, 'resource:', resource.Name);
            climateTarget = { type: 'row', rowEl: row };
            applyClimateResource(resource);
          }
        }
      }
    });
    
    // Update group weight sums after all mappings
    if(groupKey){
      const table = getTable();
      if(table){
        const tbody = table.querySelector('tbody');
        if(tbody){
          updateGroupWeightSums(groupKey, tbody);
        }
      }
    }
    
    closeMultiLayerClimateModal();
    
    // Update climate summary
    setTimeout(() => updateClimateSummary(), 100);
  }
  
  function closeMultiLayerClimateModal(){
    multiLayerClimateTarget = null;
    if(multiLayerClimateModal){
      multiLayerClimateModal.style.display = 'none';
    }
  }
  
  if(multiLayerClimateCancelBtn){
    multiLayerClimateCancelBtn.addEventListener('click', closeMultiLayerClimateModal);
  }
  
  if(multiLayerClimateModal){
    multiLayerClimateModal.addEventListener('click', function(e){
      if(e.target === multiLayerClimateModal) closeMultiLayerClimateModal();
    });
  }
  
  if(multiLayerClimateApplyBtn){
    multiLayerClimateApplyBtn.addEventListener('click', function(){
      if(!multiLayerClimateTarget) return;
      
      const { layerInfo } = multiLayerClimateTarget;
      
      // Collect all selected resources and names from the inputs
      const selectedResources = new Map();
      const selectedNames = new Map();
      let hasError = false;
      
      layerInfo.forEach(info => {
        // Get layer name input
        const nameInput = multiLayerClimateContent.querySelector(`input[data-name-input="true"][data-layer-num="${info.layerNum}"]`);
        if(nameInput){
          const layerName = nameInput.value.trim();
          if(layerName){
            selectedNames.set(info.layerNum, layerName);
          }
        }
        
        // Get climate resource input
        const searchInput = multiLayerClimateContent.querySelector(`input[data-layer-num="${info.layerNum}"]:not([data-name-input])`);
        if(!searchInput) return;
        
        const resourceName = searchInput.value.trim();
        if(!resourceName){
          // Skip empty selections
          return;
        }
        
        // Find the resource by name or index
        let resourceIndex = searchInput.dataset.selectedIndex;
        if(resourceIndex !== undefined){
          const resource = window.climateResources[resourceIndex];
          if(resource){
            selectedResources.set(info.layerNum, resource);
          }
        } else {
          // Try to find exact match by name
          const resource = window.climateResources.find(r => r.Name === resourceName);
          if(resource){
            selectedResources.set(info.layerNum, resource);
          } else {
            alert(`Kunde inte hitta klimatresurs för Skikt ${info.layerNum}: "${resourceName}"\nVälj från sökresultaten.`);
            hasError = true;
          }
        }
      });
      
      if(hasError) return;
      
      if(selectedResources.size === 0 && selectedNames.size === 0){
        alert('Välj minst en klimatresurs eller namnge minst ett skikt');
        return;
      }
      
      // Update the target with selections
      multiLayerClimateTarget.selectedResources = selectedResources;
      multiLayerClimateTarget.selectedNames = selectedNames;
      
      // Apply all resources and names
      applyAllLayerResources();
    });
  }
  if(climateApplyBtn){
    climateApplyBtn.addEventListener('click', function(){
      const selectedIndex = climateResourceSelect && climateResourceSelect.value;
      if(selectedIndex !== '' && window.climateResources && window.climateResources[selectedIndex]){
        const resource = window.climateResources[selectedIndex];
        // Save state before applying climate resource
        saveState();
        // applyClimateResource will handle closing the modal (either immediately or after manual input)
        applyClimateResource(resource);
      } else {
        closeClimateModal();
      }
    });
  }
  
  // Manual factor modal behavior
  function openManualFactorModal(resourceName, callback){
    manualFactorCallback = callback;
    if(manualFactorResourceName) manualFactorResourceName.textContent = resourceName;
    if(manualFactorValue) manualFactorValue.value = '';
    if(manualFactorUnit) manualFactorUnit.value = 'kg/m3';
    if(manualFactorModal) manualFactorModal.style.display = 'flex';
  }
  
  function closeManualFactorModal(){
    manualFactorCallback = null;
    if(manualFactorModal) manualFactorModal.style.display = 'none';
  }
  
  if(manualFactorCancelBtn){
    manualFactorCancelBtn.addEventListener('click', closeManualFactorModal);
  }
  
  if(manualFactorModal){
    manualFactorModal.addEventListener('click', function(e){
      if(e.target === manualFactorModal) closeManualFactorModal();
    });
  }
  
  if(manualFactorApplyBtn){
    manualFactorApplyBtn.addEventListener('click', function(){
      const value = manualFactorValue && parseFloat(manualFactorValue.value);
      const unit = manualFactorUnit && manualFactorUnit.value;
      
      if(!value || !Number.isFinite(value) || value <= 0){
        alert('Vänligen ange en giltig omräkningsfaktor');
        return;
      }
      
      if(manualFactorCallback){
        manualFactorCallback({
          factor: value,
          unit: unit
        });
      }
      
      closeManualFactorModal();
    });
  }

  function applyLayerSplit(count, thicknesses, mixedLayerConfigs = [], layerNames = [], climateResources = []){
    if(!layerTarget){ return; }
    const table = getTable(); if(!table) return;
    const tbody = table.querySelector('tbody'); if(!tbody) return;

    function cloneRowWithMultiplier(srcTr, multiplier, layerIndex, totalLayers, layerThickness, tableRef){
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
      
      // FIRST: Read the tds before modifying anything
      let tds = Array.from(clone.children);
      
      // Try to scale numeric cells for Net Area, Volume, Count
      // Header texts will be used later after columns are added
      
      // Add layer name to a dedicated column instead of badge
      const layerName = layerNames[layerIndex] || `Skikt ${layerIndex + 1}`;
      
      // Check if "Skiktnamn" column exists, if not create it
      if(tableRef){
        const headerRow = tableRef.querySelector('thead tr');
        if(headerRow){
          const existingLayerNameHeader = Array.from(headerRow.children).find(th => th.textContent === 'Skiktnamn');
          if(!existingLayerNameHeader){
            // Add "Skiktnamn" header at the end to avoid shifting other columns
            const layerNameTh = document.createElement('th');
            layerNameTh.textContent = 'Skiktnamn';
            headerRow.appendChild(layerNameTh);
            
            // Add empty cells to all existing rows at the end
            const tbody = tableRef.querySelector('tbody');
            if(tbody){
              const allRows = Array.from(tbody.querySelectorAll('tr'));
              allRows.forEach(row => {
                const newCell = document.createElement('td');
                newCell.textContent = '';
                row.appendChild(newCell);
              });
            }
          }
          
          // Find the layer name column index
          const headerTexts = Array.from(headerRow.children).map(th => th.textContent);
          const layerNameColumnIndex = headerTexts.findIndex(h => h === 'Skiktnamn');
          
          if(layerNameColumnIndex >= 0){
            // Update tds array to include the new column if it was just added
            tds = Array.from(clone.children);
            
            console.log('🔍 [LayerName] Setting layer name for layer', layerIndex, 'name:', layerName, 'columnIndex:', layerNameColumnIndex, 'tdsLength:', tds.length);
            
            // If the column index is out of bounds, add a new cell
            if(layerNameColumnIndex >= tds.length){
              const newCell = document.createElement('td');
              newCell.textContent = '';
              clone.appendChild(newCell);
              tds = Array.from(clone.children);
              console.log('➕ [LayerName] Added new cell, tdsLength now:', tds.length);
            }
            
            if(tds[layerNameColumnIndex]){
              // Set the layer name in the dedicated column
              tds[layerNameColumnIndex].textContent = layerName;
              console.log('✅ [LayerName] Set layer name:', layerName, 'in column:', layerNameColumnIndex);
            } else {
              console.log('❌ [LayerName] Column index out of bounds:', layerNameColumnIndex, 'tdsLength:', tds.length);
            }
          }
        }
      }
      
      // NOW calculate column indices AFTER all columns have been added
      // Use header text to find the correct column indices
      const currentHeaderTexts = Array.from(tableRef.querySelectorAll('thead th')).map(th => th.textContent);
      const idxNetArea = currentHeaderTexts.findIndex(h => String(h).toLowerCase() === 'net area');
      const idxThickness = currentHeaderTexts.findIndex(h => String(h).toLowerCase() === 'thickness');
      const idxVolume = currentHeaderTexts.findIndex(h => String(h).toLowerCase() === 'volume');
      
      let originalNetArea = null;
      if(idxNetArea >= 0 && tds[idxNetArea]){
        originalNetArea = parseNumberLike(tds[idxNetArea].textContent);
      }
      
      // Don't scale Net Area or Count - they remain unchanged
      // Instead, update Thickness column with the layer thickness
      
      // For Volume: if we have thickness specified, calculate Volume = Net Area × thickness (in meters)
      if(layerThickness && originalNetArea !== null && Number.isFinite(originalNetArea)){
        // Update Thickness cell with the layer thickness (convert from mm to m)
        if(idxThickness >= 0 && tds[idxThickness]){
          const thicknessInMeters = layerThickness / 1000;
          tds[idxThickness].textContent = String(thicknessInMeters);
        }
        
        // Calculate and update Volume
        if(idxVolume >= 0 && tds[idxVolume]){
          // Check if this row already has a volume that looks like a mixed layer volume
          // Mixed layer volumes are typically smaller than Net Area × Thickness
          const currentVolume = parseNumberLike(tds[idxVolume].textContent);
          const expectedVolume = originalNetArea * (layerThickness / 1000);
          const isMixedLayerVolume = currentVolume < expectedVolume * 0.9; // 90% threshold
          
          console.log('🔧 [cloneRowWithMultiplier] Volume check - current:', currentVolume, 'expected:', expectedVolume, 'isMixedLayer:', isMixedLayerVolume);
          
          if(isMixedLayerVolume){
            console.log('🔧 [cloneRowWithMultiplier] Skipping volume update - appears to be mixed layer volume');
          } else {
            // Thickness is in mm, convert to meters for volume calculation
            const thicknessInMeters = layerThickness / 1000;
            const newVolume = originalNetArea * thicknessInMeters;
            tds[idxVolume].textContent = String(newVolume);
            console.log('🔧 [cloneRowWithMultiplier] Updated volume to:', newVolume);
          }
        }
      } else {
        // No thickness specified, scale all numeric cells with multiplier
        function scaleCell(tdIndex){
          if(tdIndex < 0 || tdIndex >= tds.length) return;
          const td = tds[tdIndex];
          if(!td) return;
          const n = parseNumberLike(td.textContent);
          if(Number.isFinite(n)){ td.textContent = String(n * multiplier); }
        }
        if(idxNetArea >= 0) scaleCell(idxNetArea);
        if(idxVolume >= 0) scaleCell(idxVolume);
        // Find count column index
        const idxCount = currentHeaderTexts.findIndex(h => String(h).toLowerCase() === 'count');
        if(idxCount >= 0) scaleCell(idxCount);
      }
      
      // Apply climate resource if provided
      if(climateResources[layerIndex] !== undefined && climateResources[layerIndex] !== ''){
        const resourceIndex = parseInt(climateResources[layerIndex]);
        console.log('🔍 [LayerSplit] Checking climate resource for layer', layerIndex, 'resourceIndex:', resourceIndex, 'climateResources:', climateResources);
        
        if(!isNaN(resourceIndex) && window.climateResources && window.climateResources[resourceIndex]){
          const resource = window.climateResources[resourceIndex];
          console.log('🌍 [LayerSplit] Applying climate resource to layer:', layerIndex, 'resource:', resource.Name);
          
          // Use the existing applyClimateResource function
          // Set climateTarget to the clone and apply the resource
          const originalClimateTarget = climateTarget;
          climateTarget = { type: 'row', rowEl: clone };
          applyClimateResource(resource);
          climateTarget = originalClimateTarget; // Restore original target
          
          console.log('✅ [LayerSplit] Climate resource applied to layer:', layerIndex);
          
          // Re-set the layer name after climate resource application
          // as it might have been overwritten
          // But don't overwrite mixed layer names that were set by mixed layer processing
          const isMixedLayer = mixedLayerConfigs && mixedLayerConfigs.some(config => config.layerIndex === layerIndex + 1);
          
          if(!isMixedLayer){
            const layerName = layerNames[layerIndex] || `Skikt ${layerIndex + 1}`;
            const headerTexts = Array.from(tableRef.querySelectorAll('thead th')).map(th => th.textContent);
            const layerNameColumnIndex = headerTexts.findIndex(h => h === 'Skiktnamn');
            
            if(layerNameColumnIndex >= 0){
              const updatedTds = Array.from(clone.children);
              if(updatedTds[layerNameColumnIndex]){
                updatedTds[layerNameColumnIndex].textContent = layerName;
                console.log('🔄 [LayerName] Re-set layer name after climate resource:', layerName);
              }
            }
          } else {
            console.log('🔄 [LayerName] Skipping layer name reset for mixed layer:', layerIndex + 1);
          }
        } else {
          console.log('❌ [LayerSplit] No valid climate resource found for layer:', layerIndex, 'resourceIndex:', resourceIndex);
        }
      } else {
        console.log('ℹ️ [LayerSplit] No climate resource specified for layer:', layerIndex);
      }
      
      return clone;
    }

    function splitRow(tr, savedLayerKey){
      // Use saved layer key if provided, otherwise generate new one
      const layerKey = savedLayerKey || 'layer-' + Date.now() + '-' + Math.random().toString(36).substr(2, 9);
      
      // Check if this row is already layered (has existing layer children)
      const existingLayerKey = tr.getAttribute('data-layer-key');
      if(existingLayerKey && !savedLayerKey){
        // Re-layering: remove all existing layer children first
        const existingChildren = Array.from(tbody.querySelectorAll(`tr[data-layer-child-of="${CSS.escape(existingLayerKey)}"]`));
        console.log('🔧 Omskiktar rad - tar bort', existingChildren.length, 'befintliga skikt');
        existingChildren.forEach(child => child.remove());
        
        // Remove layer label from first data cell
        const firstDataTd = tr.querySelector('td:nth-child(2)');
        if(firstDataTd){
          const toggle = firstDataTd.querySelector('.group-toggle');
          const layerLabel = Array.from(firstDataTd.childNodes).find(node => 
            node.nodeType === Node.TEXT_NODE && node.textContent.includes('skikt')
          );
          if(toggle) toggle.remove();
          if(layerLabel) layerLabel.remove();
          // Also remove any span with layer count
          const spans = firstDataTd.querySelectorAll('span');
          spans.forEach(span => {
            if(span.textContent.includes('skikt')) span.remove();
          });
        }
      }
      
      // Save layer data for this row
      if(!savedLayerKey){
        // Use original row data if available, otherwise extract from DOM
        const hasOriginal = !!tr._originalRowData;
        const rowData = tr._originalRowData || getRowDataFromTr(tr);
        if(rowData && Array.isArray(rowData)){
          const layerChildOf = tr.getAttribute('data-layer-child-of');
          const signature = getRowSignature(rowData, layerChildOf);
          const beforeSize = layerData.size;
          layerData.set(signature, { count, thicknesses, layerKey: existingLayerKey || layerKey });
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
      tr.setAttribute('data-open', 'false'); // Start collapsed by default
      
      // Update action buttons on parent row
      const actionTd = tr.querySelector('td:first-child');
      if(actionTd){
        actionTd.innerHTML = '';
        const parentLayerBtn = document.createElement('button');
        parentLayerBtn.type = 'button';
        parentLayerBtn.textContent = 'Skikta skikt';
        parentLayerBtn.addEventListener('click', function(ev){
          ev.stopPropagation();
          // Open as 'row' type so it can be re-layered
          openLayerModal({ type: 'row', rowEl: tr });
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
      // Adjust layer index for mixed layers - if previous layers were mixed, they take 2 slots each
      let adjustedThicknesses = [...thicknesses];
      
      // Adjust thicknesses for mixed layers
      if(mixedLayerConfigs && mixedLayerConfigs.length > 0){
        // Create a new thickness array that accounts for mixed layers
        adjustedThicknesses = [];
        for(let i = 0; i < count; i++){
          const isMixedLayer = mixedLayerConfigs.some(config => config.layerIndex === i + 1);
          if(isMixedLayer){
            // For mixed layers, both materials get the same thickness
            adjustedThicknesses.push(thicknesses[i]);
            adjustedThicknesses.push(thicknesses[i]);
          } else {
            // For regular layers, use the original thickness
            adjustedThicknesses.push(thicknesses[i]);
          }
        }
      }
      
      const fragments = multipliers.map((m, i) => {
        // Always use the original thickness for each layer
        // The adjustedThicknesses array is only used for mixed layer splitting later
        const layerThickness = thicknesses.length > 0 ? thicknesses[i] : undefined;
        const clone = cloneRowWithMultiplier(tr, m, i, multipliers.length, layerThickness, table);
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
        
        // Set layer name for this layer
        const layerName = layerNames[i] || `Skikt ${i + 1}`;
        const headerTexts = Array.from(table.querySelectorAll('thead th')).map(th => th.textContent);
        const layerNameColumnIndex = headerTexts.findIndex(h => h === 'Skiktnamn');
        
        if(layerNameColumnIndex >= 0){
          const cloneCells = Array.from(clone.children);
          if(cloneCells[layerNameColumnIndex]){
            cloneCells[layerNameColumnIndex].textContent = layerName;
          }
        }
        
        return clone;
      });
      
      // Insert layer children right after the parent
      let insertAfter = tr;
      fragments.forEach((f, idx) => {
        tbody.insertBefore(f, insertAfter.nextSibling);
        insertAfter = f;
      });
      
      // Update parent row's Volume to show sum of all layers (AFTER creating children)
      if(thicknesses.length > 0){
        const parentTds = Array.from(tr.children);
        const parentOriginalCellCount = parentTds.length; // Parent may have more cells after climate columns added
        
        // Use same backward counting as in cloneRowWithMultiplier
        // Find where the original data ends (before any added climate columns)
        let parentDataCellCount = parentOriginalCellCount;
        // If parent has been extended with climate columns, find original end
        for(let i = parentTds.length - 1; i >= 0; i--){
          if(parentTds[i].hasAttribute('data-climate-cell') || 
             parentTds[i].hasAttribute('data-factor-cell') ||
             parentTds[i].textContent === ''){
            parentDataCellCount = i;
          } else {
            break;
          }
        }
        
        const parentCountCellIdx = Math.min(12, parentDataCellCount - 1); // Count is typically at index 12
        const parentVolumeCellIdx = parentCountCellIdx - 1; // Volume is before Count
        const parentNetAreaCellIdx = parentVolumeCellIdx - 4; // Net Area is 4 cells before Volume
        
        const parentVolumeTd = parentTds[parentVolumeCellIdx];
        const parentNetAreaTd = parentTds[parentNetAreaCellIdx];
        
        if(parentVolumeTd && parentNetAreaTd){
          const netArea = parseNumberLike(parentNetAreaTd.textContent);
          if(Number.isFinite(netArea)){
            // Calculate total volume from all layers
            let totalVolume = 0;
            for(let i = 0; i < thicknesses.length; i++){
              const thicknessInMeters = thicknesses[i] / 1000;
              totalVolume += netArea * thicknessInMeters;
            }
            parentVolumeTd.textContent = String(totalVolume);
          }
        }
      }
    }

    if(layerTarget.type === 'row' && layerTarget.rowEl){
      splitRow(layerTarget.rowEl);
    } else if(layerTarget.type === 'group' && layerTarget.key != null){
      // Check if this group is already layered (has layer children)
      const existingChildren = Array.from(tbody.querySelectorAll(`tr[data-layer-child-of="${CSS.escape(layerTarget.key)}"]`));
      
      if(existingChildren.length > 0){
        // Re-layering: remove all existing layer children first
        console.log('🔧 Omskiktar grupp - tar bort', existingChildren.length, 'befintliga skikt');
        existingChildren.forEach(child => child.remove());
        
        // Also remove layer-parent class and attributes from parent
        const parentTr = tbody.querySelector(`tr[data-layer-key="${CSS.escape(layerTarget.key)}"]`);
        if(parentTr){
          parentTr.classList.remove('layer-parent');
          parentTr.removeAttribute('data-layer-key');
          parentTr.removeAttribute('data-open');
          
          // Reset action buttons to original state
          const actionTd = parentTr.querySelector('td:first-child');
          if(actionTd){
            actionTd.innerHTML = '';
            const layerBtn = document.createElement('button');
            layerBtn.type = 'button';
            layerBtn.textContent = 'Skikta skikt';
            layerBtn.addEventListener('click', function(ev){
              ev.stopPropagation();
              openLayerModal({ type: 'group', key: layerTarget.key });
            });
            actionTd.appendChild(layerBtn);
            
            const climateBtn = document.createElement('button');
            climateBtn.type = 'button';
            climateBtn.textContent = 'Mappa klimatresurs';
            climateBtn.addEventListener('click', function(ev){
              ev.stopPropagation();
              openClimateModal({ type: 'group', key: layerTarget.key });
            });
            actionTd.appendChild(climateBtn);
          }
          
          // Remove layer label from first data cell
          const firstDataTd = parentTr.querySelector('td:nth-child(2)');
          if(firstDataTd){
            const toggle = firstDataTd.querySelector('.group-toggle');
            const layerLabel = firstDataTd.querySelector('span:last-child');
            if(toggle) toggle.remove();
            if(layerLabel && layerLabel.textContent.includes('skikt')) layerLabel.remove();
          }
        }
      }
      
      // Now split the group rows (which are no longer layer children)
      const rows = Array.from(tbody.querySelectorAll('tr[data-group-child-of="' + CSS.escape(layerTarget.key) + '"]:not([data-layer-child-of])'));
      console.log('🔧 Skiktar grupp - antal rader:', rows.length);
      
      // Generate unique layer keys for each row to avoid conflicts with other objects
      // Each row gets its own unique layer keys
      console.log('🔧 Genererade layerKeys för grupp:', layerTarget.key);
      
      rows.forEach((row, rowIndex) => {
        console.log('🔧 Skiktar rad', rowIndex + 1, 'av', rows.length);
        
        // Check if this row is already layered and remove existing children
        const existingRowLayerKey = row.getAttribute('data-layer-key');
        if(existingRowLayerKey){
          // This row is already layered - remove its children
          const existingRowChildren = Array.from(tbody.querySelectorAll(`tr[data-layer-child-of="${CSS.escape(existingRowLayerKey)}"]`));
          console.log('🔧 Omskiktar rad inom grupp - tar bort', existingRowChildren.length, 'befintliga skikt');
          existingRowChildren.forEach(child => child.remove());
          
          // Remove layer label from first data cell
          const firstDataTd = row.querySelector('td:nth-child(2)');
          if(firstDataTd){
            const toggle = firstDataTd.querySelector('.group-toggle');
            const spans = firstDataTd.querySelectorAll('span');
            if(toggle) toggle.remove();
            spans.forEach(span => {
              if(span.textContent.includes('skikt')) span.remove();
            });
          }
        }
        
        // Split this row, but we need to set individual layerKeys for each child
        // So we'll do it manually here instead of calling splitRow
        const rowLayerKey = existingRowLayerKey || 'row-' + Date.now() + '-' + Math.random().toString(36).substr(2, 9);
        
        // Generate unique layer keys for this specific row
        const rowLayerKeys = Array.from({ length: count }, (_, i) => 
          rowLayerKey + '_Layer_' + (i + 1)
        );
        
        // Save layer data for this row
        const hasOriginal = !!row._originalRowData;
        const rowData = row._originalRowData || getRowDataFromTr(row);
        if(rowData && Array.isArray(rowData)){
          const layerChildOf = row.getAttribute('data-layer-child-of');
          const signature = getRowSignature(rowData, layerChildOf);
          const beforeSize = layerData.size;
          layerData.set(signature, { count, thicknesses, layerKey: rowLayerKey });
          const afterSize = layerData.size;
          console.log('💾 SAVE:', rowData[1]?.substring(0,10), '- Has _originalRowData:', hasOriginal, '- LayerChild:', layerChildOf?.substring(0,10) || 'none', '- Size:', beforeSize, '→', afterSize, '- Signature:', signature.substring(0, 60));
        }
        
        // Even split if no thicknesses provided
        const multipliers = thicknesses.length > 0
          ? thicknesses.map(t => t / thicknesses.reduce((a,b)=>a+b,0))
          : Array(count).fill(1 / count);
        
        // Convert original row to parent row
        row.classList.remove('is-new');
        row.classList.add('layer-parent');
        row.setAttribute('data-layer-key', rowLayerKey);
        row.setAttribute('data-open', 'false'); // Start collapsed by default
        
        // Update action buttons on parent row
        const actionTd = row.querySelector('td:first-child');
        if(actionTd){
          actionTd.innerHTML = '';
          const parentLayerBtn = document.createElement('button');
          parentLayerBtn.type = 'button';
          parentLayerBtn.textContent = 'Skikta skikt';
          parentLayerBtn.addEventListener('click', function(ev){
            ev.stopPropagation();
            openLayerModal({ type: 'group', key: rowLayerKey });
          });
          actionTd.appendChild(parentLayerBtn);
          
          const parentClimateBtn = document.createElement('button');
          parentClimateBtn.type = 'button';
          parentClimateBtn.textContent = 'Mappa klimatresurs';
          parentClimateBtn.addEventListener('click', function(ev){
            ev.stopPropagation();
            openClimateModal({ type: 'group', key: rowLayerKey });
          });
          actionTd.appendChild(parentClimateBtn);
        }
        
        // Add toggle to first data cell
        const firstDataTd = row.querySelector('td:nth-child(2)');
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
        const parentGroupKey = row.getAttribute('data-group-child-of');
        const originalRowData = row._originalRowData;
        
        // Create child layer rows
        // Adjust layer index for mixed layers - if previous layers were mixed, they take 2 slots each
        let adjustedThicknesses = [...thicknesses];
        
        // Adjust thicknesses for mixed layers
        if(mixedLayerConfigs && mixedLayerConfigs.length > 0){
          // Create a new thickness array that accounts for mixed layers
          adjustedThicknesses = [];
          for(let i = 0; i < count; i++){
            const isMixedLayer = mixedLayerConfigs.some(config => config.layerIndex === i + 1);
            if(isMixedLayer){
              // For mixed layers, both materials get the same thickness
              adjustedThicknesses.push(thicknesses[i]);
              adjustedThicknesses.push(thicknesses[i]);
            } else {
              // For regular layers, use the original thickness
              adjustedThicknesses.push(thicknesses[i]);
            }
          }
        }
        
        const fragments = multipliers.map((m, i) => {
          // Always use the original thickness for each layer
          // The adjustedThicknesses array is only used for mixed layer splitting later
          const layerThickness = thicknesses.length > 0 ? thicknesses[i] : undefined;
          const clone = cloneRowWithMultiplier(row, m, i, multipliers.length, layerThickness, table);
          // Mark as child of this row's layer
          clone.setAttribute('data-layer-child-of', rowLayerKey);
          // Also inherit parent's group membership if it exists
          if(parentGroupKey){
            clone.setAttribute('data-group-child-of', parentGroupKey);
          }
          // Set immediate parent for toggle
          clone.setAttribute('data-parent-key', rowLayerKey);
          // Set the unique layer key for this specific row and layer
          clone.setAttribute('data-layer-key', rowLayerKeys[i]);
          // Preserve original row data
          if(originalRowData){
            clone._originalRowData = originalRowData;
          }
          
          // Set layer name for this layer
          const layerName = layerNames[i] || `Skikt ${i + 1}`;
          const headerTexts = Array.from(table.querySelectorAll('thead th')).map(th => th.textContent);
          const layerNameColumnIndex = headerTexts.findIndex(h => h === 'Skiktnamn');
          
          if(layerNameColumnIndex >= 0){
            const cloneCells = Array.from(clone.children);
            if(cloneCells[layerNameColumnIndex]){
              cloneCells[layerNameColumnIndex].textContent = layerName;
            }
          }
          
          return clone;
        });
        
        // Insert layer children right after the parent
        let insertAfter = row;
        fragments.forEach(f => {
          tbody.insertBefore(f, insertAfter.nextSibling);
          insertAfter = f;
        });
        
        // Update parent row's Volume to show sum of all layers (AFTER creating children)
        if(thicknesses.length > 0){
          const groupParentTds = Array.from(row.children);
          const groupParentOriginalCellCount = groupParentTds.length;
          
          // Find where original data ends
          let groupParentDataCellCount = groupParentOriginalCellCount;
          for(let i = groupParentTds.length - 1; i >= 0; i--){
            if(groupParentTds[i].hasAttribute('data-climate-cell') || 
               groupParentTds[i].hasAttribute('data-factor-cell') ||
               groupParentTds[i].textContent === ''){
              groupParentDataCellCount = i;
            } else {
              break;
            }
          }
          
          const groupParentCountCellIdx = Math.min(12, groupParentDataCellCount - 1);
          const groupParentVolumeCellIdx = groupParentCountCellIdx - 1;
          const groupParentNetAreaCellIdx = groupParentVolumeCellIdx - 4;
          
          const groupParentVolumeTd = groupParentTds[groupParentVolumeCellIdx];
          const groupParentNetAreaTd = groupParentTds[groupParentNetAreaCellIdx];
          
          if(groupParentVolumeTd && groupParentNetAreaTd){
            const groupNetArea = parseNumberLike(groupParentNetAreaTd.textContent);
            if(Number.isFinite(groupNetArea)){
              // Calculate total volume from all layers
              let groupTotalVolume = 0;
              for(let i = 0; i < thicknesses.length; i++){
                const groupThicknessInMeters = thicknesses[i] / 1000;
                groupTotalVolume += groupNetArea * groupThicknessInMeters;
              }
              groupParentVolumeTd.textContent = String(groupTotalVolume);
            }
          }
        }
      });
    }
    
    // Handle mixed layer splitting (after all layers have been created)
    if(mixedLayerConfigs && mixedLayerConfigs.length > 0){
      console.log('🔧 Processing mixed layer configs:', mixedLayerConfigs);
      
      // Process each mixed layer
      mixedLayerConfigs.forEach((mixedLayerConfig, configIndex) => {
        let targetLayerIndex = mixedLayerConfig.layerIndex - 1; // Convert from 1-based to 0-based
        console.log(`🔧 Processing mixed layer config ${configIndex + 1}/${mixedLayerConfigs.length}:`, mixedLayerConfig, 'targetLayerIndex:', targetLayerIndex);
      
      // Find the layer parent(s) that were just created for THIS specific object only
      // Only get top-level parents (not children that happen to have layer-parent class)
      let layerParents;
      if(layerTarget.type === 'group'){
        // For groups, only process rows that belong to this group
        layerParents = Array.from(tbody.querySelectorAll(`tr[data-group-child-of="${CSS.escape(layerTarget.key)}"].layer-parent[data-layer-key]`));
      } else {
        // For individual rows, only process this specific row
        layerParents = Array.from(tbody.querySelectorAll(`tr.layer-parent[data-layer-key]`)).filter(parent => 
          parent === layerTarget.rowEl
        );
      }
      
      // Filter out any children that happen to have layer-parent class
      layerParents = layerParents.filter(parent => {
        const layerKey = parent.getAttribute('data-layer-key');
        const layerChildOf = parent.getAttribute('data-layer-child-of');
        // A true parent should not also be a child of the same layer
        return !layerChildOf || layerChildOf !== layerKey;
      });
      
      console.log('🔍 Found', layerParents.length, 'parent rows to process');
      
      layerParents.forEach(parentTr => {
        const layerKey = parentTr.getAttribute('data-layer-key');
        if(!layerKey) return;
        
        // Find all layer children for this parent
        const layerChildren = Array.from(tbody.querySelectorAll(`tr[data-layer-child-of="${CSS.escape(layerKey)}"]`));
        console.log('🔍 [MixedLayer] All layer children for parent:', layerKey, 'count:', layerChildren.length);
        
        // Check if the target layer index is valid for this parent
        if(targetLayerIndex < 0 || targetLayerIndex >= layerChildren.length){
          console.log('❌ Invalid layer index:', targetLayerIndex, 'for parent with', layerChildren.length, 'layers');
          console.log('🔍 Available layer children:', layerChildren.map((child, idx) => `${idx}: ${child.textContent?.substring(0, 50)}...`));
          return;
        }

        // Get reference to target layer
        let targetLayer = layerChildren[targetLayerIndex];
        
        // Check if this layer is already a mixed layer (has been processed before)
        if(targetLayer && targetLayer.hasAttribute('data-mixed-layer')){
          console.log('⚠️ [MixedLayer] Layer', targetLayerIndex, 'is already a mixed layer, skipping');
          console.log('🔍 [MixedLayer] Available layers after skipping:', layerChildren.map((child, idx) => `${idx}: ${child.textContent?.substring(0, 30)}... (mixed: ${child.hasAttribute('data-mixed-layer')})`));
          
          // Find the next available layer that is not a mixed layer
          let nextAvailableIndex = -1;
          for(let i = 0; i < layerChildren.length; i++){
            if(!layerChildren[i].hasAttribute('data-mixed-layer')){
              nextAvailableIndex = i;
              break;
            }
          }
          
          if(nextAvailableIndex >= 0){
            console.log('🔄 [MixedLayer] Found next available layer at index:', nextAvailableIndex);
            // Update targetLayerIndex to use the next available layer
            targetLayerIndex = nextAvailableIndex;
            targetLayer = layerChildren[nextAvailableIndex];
            console.log('✅ Found next target layer:', targetLayer);
            
            // Continue processing with the next available layer
            // (fall through to the rest of the processing)
          } else {
            console.log('❌ [MixedLayer] No available layers found for mixed layer config');
            return;
          }
        }
        console.log('✅ Found target layer:', targetLayer);
        console.log('🔍 [MixedLayer] Layer children count:', layerChildren.length);
        console.log('🔍 [MixedLayer] Target layer index:', targetLayerIndex);
        console.log('🔍 [MixedLayer] All layer children:', layerChildren.map((child, idx) => `${idx}: ${child.textContent?.substring(0, 50)}...`));
        
        // Remove layer-parent class from target if it has it (it shouldn't as a child)
        targetLayer.classList.remove('layer-parent');
        
        // Clone the target layer to create the second material row
        // This adds an extra row - mixed layers expand the total count
        const material2Row = targetLayer.cloneNode(true);
        
        // Ensure cloned row doesn't have layer-parent class
        material2Row.classList.remove('layer-parent');
        material2Row.removeAttribute('data-layer-key'); // Only the parent should have this as a unique key
        
        // Mark both rows as mixed layers to prevent re-processing
        targetLayer.setAttribute('data-mixed-layer', 'true');
        material2Row.setAttribute('data-mixed-layer', 'true');

        // Update layer names in the dedicated column for both materials
        const headerTexts = Array.from(table.querySelectorAll('thead th')).map(th => th.textContent);
        const layerNameColumnIndex = headerTexts.findIndex(h => h === 'Skiktnamn');
        console.log('🔧 [MixedLayer] Header texts:', headerTexts);
        console.log('🔧 [MixedLayer] Layer name column index:', layerNameColumnIndex);

        if(layerNameColumnIndex >= 0){
          // Update material 1 layer name
          const mat1Cells = Array.from(targetLayer.children);
          if(mat1Cells[layerNameColumnIndex]){
            const material1Name = mixedLayerConfig.material1.name + ' (' + mixedLayerConfig.material1.percent + '%)';
            mat1Cells[layerNameColumnIndex].textContent = material1Name;
            console.log('🔧 [MixedLayer] Set material 1 name to:', material1Name);
            console.log('🔧 [MixedLayer] Material 1 cell content after setting:', mat1Cells[layerNameColumnIndex].textContent);
          }

          // Update material 2 layer name
          const mat2Cells = Array.from(material2Row.children);
          if(mat2Cells[layerNameColumnIndex]){
            const material2Name = mixedLayerConfig.material2.name + ' (' + mixedLayerConfig.material2.percent + '%)';
            mat2Cells[layerNameColumnIndex].textContent = material2Name;
            console.log('🔧 [MixedLayer] Set material 2 name to:', material2Name);
            console.log('🔧 [MixedLayer] Material 2 cell content after setting:', mat2Cells[layerNameColumnIndex].textContent);
          }
        }
        
        // Calculate thickness and volume - thickness stays the same, volume uses percentage
        const idxNetArea = headerTexts.findIndex(h => String(h).toLowerCase() === 'net area');
        const idxThickness = headerTexts.findIndex(h => String(h).toLowerCase() === 'thickness');
        const idxVolume = headerTexts.findIndex(h => String(h).toLowerCase() === 'volume');
        
        console.log('🔧 [MixedLayer] Volume calculation check:');
        console.log('  idxNetArea:', idxNetArea, 'idxVolume:', idxVolume, 'thicknesses.length:', thicknesses.length, 'targetLayerIndex:', targetLayerIndex);
        console.log('  Condition result:', idxNetArea >= 0 && idxVolume >= 0 && thicknesses.length > targetLayerIndex);
        
        // For mixed layers, we need to use the original layer index for thickness lookup
        // targetLayerIndex might be higher than thicknesses.length if we found a "next available" layer
        const originalLayerIndex = mixedLayerConfig.layerIndex - 1; // Convert from 1-based to 0-based
        const thicknessIndex = Math.min(originalLayerIndex, thicknesses.length - 1);
        console.log('🔧 [MixedLayer] Using thickness index:', thicknessIndex, 'for original layer:', originalLayerIndex);
        
        if(idxNetArea >= 0 && idxVolume >= 0 && thicknesses.length > thicknessIndex){
          const mat1Cells = Array.from(targetLayer.children);
          const mat2Cells = Array.from(material2Row.children);
          
          console.log('🔧 [MixedLayer] Processing mixed layer config:', mixedLayerConfig.layerIndex, 'targetLayerIndex:', targetLayerIndex);
          console.log('🔧 [MixedLayer] Thicknesses array:', thicknesses);
          console.log('🔧 [MixedLayer] Target layer thickness:', thicknesses[thicknessIndex]);
          
          // Find cells using header text (more reliable than backward counting)
          const mat1NetAreaCell = mat1Cells[idxNetArea];
          const mat1ThicknessCell = mat1Cells[idxThickness];
          const mat1VolumeCell = mat1Cells[idxVolume];
          
          const mat2NetAreaCell = mat2Cells[idxNetArea];
          const mat2ThicknessCell = mat2Cells[idxThickness];
          const mat2VolumeCell = mat2Cells[idxVolume];
          
          if(mat1NetAreaCell && mat1VolumeCell && mat2VolumeCell){
            const netArea = parseNumberLike(mat1NetAreaCell.textContent);
            const layerThickness = thicknesses[thicknessIndex]; // in mm
            
            if(Number.isFinite(netArea) && Number.isFinite(layerThickness)){
              const thicknessInMeters = layerThickness / 1000; // Convert mm to m
              const mat1Percent = mixedLayerConfig.material1.percent / 100;
              const mat2Percent = mixedLayerConfig.material2.percent / 100;
              
              // Thickness remains the same for both materials (same as original layer)
              if(mat1ThicknessCell){
                mat1ThicknessCell.textContent = String(thicknessInMeters);
              }
              if(mat2ThicknessCell){
                mat2ThicknessCell.textContent = String(thicknessInMeters);
              }
              
              // Calculate volume for each material: Net Area × thickness × percent
              // This represents the volume proportion of each material within the same thickness
              const mat1Volume = netArea * thicknessInMeters * mat1Percent;
              const mat2Volume = netArea * thicknessInMeters * mat2Percent;
              
              mat1VolumeCell.textContent = String(mat1Volume);
              mat2VolumeCell.textContent = String(mat2Volume);
              
              console.log('🔧 [MixedLayer] Set material 1 volume to:', mat1Volume, 'm³');
              console.log('🔧 [MixedLayer] Set material 2 volume to:', mat2Volume, 'm³');
              
              // Add a watcher to detect if volume changes
              const originalMat1Volume = mat1Volume;
              const originalMat2Volume = mat2Volume;
              
              // Check volumes after a short delay
              setTimeout(() => {
                const currentMat1Volume = parseNumberLike(mat1VolumeCell.textContent);
                const currentMat2Volume = parseNumberLike(mat2VolumeCell.textContent);
                
                if(Math.abs(currentMat1Volume - originalMat1Volume) > 0.0001){
                  console.log('⚠️ [VolumeWatcher] Material 1 volume changed from', originalMat1Volume, 'to', currentMat1Volume);
                  console.log('🔍 [VolumeWatcher] Material 1 cell element:', mat1VolumeCell);
                  console.log('🔍 [VolumeWatcher] Material 1 cell parent row:', mat1VolumeCell.parentElement);
                }
                if(Math.abs(currentMat2Volume - originalMat2Volume) > 0.0001){
                  console.log('⚠️ [VolumeWatcher] Material 2 volume changed from', originalMat2Volume, 'to', currentMat2Volume);
                  console.log('🔍 [VolumeWatcher] Material 2 cell element:', mat2VolumeCell);
                  console.log('🔍 [VolumeWatcher] Material 2 cell parent row:', mat2VolumeCell.parentElement);
                }
              }, 100);
              
              console.log('📊 Calculated mixed layer volumes:');
              console.log('   Net Area:', netArea, 'm²');
              console.log('   Thickness (same for both):', layerThickness, 'mm (', thicknessInMeters, 'm)');
              console.log('   Material 1 (' + mixedLayerConfig.material1.percent + '%):', 'Volume:', mat1Volume, 'm³');
              console.log('   Material 2 (' + mixedLayerConfig.material2.percent + '%):', 'Volume:', mat2Volume, 'm³');
              console.log('   Total Volume:', (mat1Volume + mat2Volume), 'm³');
            }
          }
        }
        
        // Reset action buttons for material2Row with fresh event listeners
        const mat2ActionTd = material2Row.querySelector('td:first-child');
        if(mat2ActionTd){
          mat2ActionTd.innerHTML = '';
          
          const layerBtn = document.createElement('button');
          layerBtn.type = 'button';
          layerBtn.textContent = 'Skikta';
          layerBtn.addEventListener('click', function(ev){
            ev.stopPropagation();
            openLayerModal({ type: 'row', rowEl: material2Row });
          });
          mat2ActionTd.appendChild(layerBtn);
          
          const climateBtn = document.createElement('button');
          climateBtn.type = 'button';
          climateBtn.textContent = 'Mappa klimatresurs';
          climateBtn.addEventListener('click', function(ev){
            ev.stopPropagation();
            openClimateModal({ type: 'row', rowEl: material2Row });
          });
          mat2ActionTd.appendChild(climateBtn);
        }
        
        // Apply climate resources to each material separately
        const material1ClimateResource = mixedLayerConfig.material1.climateResource;
        const material2ClimateResource = mixedLayerConfig.material2.climateResource;
        
        if(material1ClimateResource && material1ClimateResource !== ''){
          const resourceIndex1 = parseInt(material1ClimateResource);
          console.log('🔍 [MixedLayer] Material 1 climate resource:', resourceIndex1);
          
          if(!isNaN(resourceIndex1) && window.climateResources && window.climateResources[resourceIndex1]){
            const resource1 = window.climateResources[resourceIndex1];
            console.log('🌍 [MixedLayer] Applying climate resource to material 1:', resource1.Name);
            
            const originalClimateTarget = climateTarget;
            climateTarget = { type: 'row', rowEl: targetLayer };
            applyClimateResource(resource1);
            climateTarget = originalClimateTarget;
            
            console.log('✅ [MixedLayer] Climate resource applied to material 1');
          }
        }
        
        if(material2ClimateResource && material2ClimateResource !== ''){
          const resourceIndex2 = parseInt(material2ClimateResource);
          console.log('🔍 [MixedLayer] Material 2 climate resource:', resourceIndex2);
          
          if(!isNaN(resourceIndex2) && window.climateResources && window.climateResources[resourceIndex2]){
            const resource2 = window.climateResources[resourceIndex2];
            console.log('🌍 [MixedLayer] Applying climate resource to material 2:', resource2.Name);
            
            const originalClimateTarget = climateTarget;
            climateTarget = { type: 'row', rowEl: material2Row };
            applyClimateResource(resource2);
            climateTarget = originalClimateTarget;
            
            console.log('✅ [MixedLayer] Climate resource applied to material 2');
          }
        }
        
        // Insert material2Row right after targetLayer
        tbody.insertBefore(material2Row, targetLayer.nextSibling);
        
        console.log('✅ Mixed layer split completed for layer', targetLayerIndex + 1);
        
        // Update parent row's layer count label to reflect the added row
        const parentFirstDataCell = parentTr.querySelector('td:nth-child(2)');
        if(parentFirstDataCell){
          // Find and update the layer count span
          const layerLabelSpan = Array.from(parentFirstDataCell.childNodes).find(node => 
            node.nodeType === Node.TEXT_NODE && node.textContent.includes('skikt')
          );
          if(!layerLabelSpan){
            // Try to find a span element
            const layerSpan = Array.from(parentFirstDataCell.querySelectorAll('span')).find(span => 
              span.textContent.includes('skikt')
            );
            if(layerSpan){
              // Update count to include the extra row from mixed layer
              const newCount = count + 1; // +1 for the mixed layer split
              layerSpan.textContent = ' [' + newCount + ' skikt]';
              console.log('📊 Updated parent label to show', newCount, 'layers');
            }
          }
        }
        
        // Update parent row's total volume to reflect the mixed layer split
        if(thicknesses.length > 0){
          const parentTds = Array.from(parentTr.children);
          const headerTexts = Array.from(table.querySelectorAll('thead th')).map(th => th.textContent);
          const idxNetArea = headerTexts.findIndex(h => String(h).toLowerCase() === 'net area');
          const idxVolume = headerTexts.findIndex(h => String(h).toLowerCase() === 'volume');
          
          if(idxNetArea >= 0 && idxVolume >= 0){
            const parentNetAreaCell = parentTds[idxNetArea + 1];
            const parentVolumeCell = parentTds[idxVolume + 1];
            
            if(parentNetAreaCell && parentVolumeCell){
              const netArea = parseNumberLike(parentNetAreaCell.textContent);
              if(Number.isFinite(netArea)){
                // Recalculate total volume from all layer children
                let totalVolume = 0;
                const layerChildren = Array.from(tbody.querySelectorAll(`tr[data-layer-child-of="${CSS.escape(layerKey)}"]`));
                layerChildren.forEach(child => {
                  const cells = Array.from(child.children);
                  const volumeCell = cells[idxVolume];
                  if(volumeCell){
                    const volume = parseNumberLike(volumeCell.textContent);
                    if(Number.isFinite(volume)){
                      totalVolume += volume;
                    }
                  }
                });
                parentVolumeCell.textContent = String(totalVolume);
                console.log('📊 Updated parent volume to:', totalVolume, 'm³');
              }
            }
          }
        }
      });
      }); // Close forEach loop for mixedLayerConfigs
    }
    
    // Debug: Check layer names after mixed layer processing
    console.log('🔍 [FinalCheck] Checking layer names after mixed layer processing:');
    const headerTexts = Array.from(table.querySelectorAll('thead th')).map(th => th.textContent);
    const allRows = Array.from(tbody.querySelectorAll('tr[data-layer-child-of]'));
    allRows.forEach((row, index) => {
      const cells = Array.from(row.children);
      const layerNameColumnIndex = headerTexts.findIndex(h => h === 'Skiktnamn');
      if(layerNameColumnIndex >= 0 && cells[layerNameColumnIndex]){
        console.log(`🔍 [FinalCheck] Row ${index + 1} layer name:`, cells[layerNameColumnIndex].textContent);
      }
    });
    
    // Re-apply filters to keep visibility consistent
    applyFilters();
    
    // Debug: Check final volumes after all processing
    console.log('🔍 [FinalVolumeCheck] Checking final volumes after all processing:');
    const allLayerChildren = Array.from(tbody.querySelectorAll('tr[data-layer-child-of]'));
    const finalHeaderTexts = Array.from(table.querySelectorAll('thead th')).map(th => th.textContent);
    console.log('🔍 [FinalVolumeCheck] Header texts:', finalHeaderTexts);
    const volumeColumnIndex = finalHeaderTexts.findIndex(h => h.toLowerCase() === 'volume');
    console.log('🔍 [FinalVolumeCheck] Volume column index:', volumeColumnIndex);
    
    allLayerChildren.forEach((row, index) => {
      const cells = Array.from(row.children);
      console.log(`🔍 [FinalVolumeCheck] Row ${index + 1} has ${cells.length} cells`);
      if(volumeColumnIndex >= 0 && cells[volumeColumnIndex]){
        console.log(`🔍 [FinalVolumeCheck] Row ${index + 1} volume:`, cells[volumeColumnIndex].textContent, 'm³');
      } else {
        console.log(`🔍 [FinalVolumeCheck] Row ${index + 1} - no volume cell found at index ${volumeColumnIndex}`);
      }
    });
    
    // Debug: Check layer names after applyFilters
    console.log('🔍 [AfterFilters] Checking layer names after applyFilters:');
    const allRowsAfterFilters = Array.from(tbody.querySelectorAll('tr[data-layer-child-of]'));
    allRowsAfterFilters.forEach((row, index) => {
      const cells = Array.from(row.children);
      const layerNameColumnIndex = headerTexts.findIndex(h => h === 'Skiktnamn');
      if(layerNameColumnIndex >= 0 && cells[layerNameColumnIndex]){
        console.log(`🔍 [AfterFilters] Row ${index + 1} layer name:`, cells[layerNameColumnIndex].textContent);
      }
    });
    
    // Update climate summary after layering
    setTimeout(() => updateClimateSummary(), 100);
  }

  function applyClimateResource(resource){
    if(!climateTarget){ return; }
    
    // Save climateTarget because it might be cleared if modals close
    const savedClimateTarget = climateTarget;
    
    const table = getTable(); if(!table) return;
    const thead = table.querySelector('thead'); if(!thead) return;
    const tbody = table.querySelector('tbody'); if(!tbody) return;
    
    // Check if "Klimatresurs", "Omräkningsfaktor", "Omräkningsfaktor enhet" and "Avfallsfaktor" columns already exist
    const headerRow = thead.querySelector('tr');
    const existingClimateHeader = Array.from(headerRow.children).find(th => th.textContent === 'Klimatresurs');
    const existingFactorHeader = Array.from(headerRow.children).find(th => th.textContent === 'Omräkningsfaktor');
    const FactorUnit = Array.from(headerRow.children).find(th => th.textContent === 'Omräkningsfaktor enhet');
    const existingWasteHeader = Array.from(headerRow.children).find(th => th.textContent === 'Spillfaktor');
    const existingA1_A3Header = Array.from(headerRow.children).find(th => th.textContent === 'Emissionsfaktor A1-A3');
    const existingA4Header = Array.from(headerRow.children).find(th => th.textContent === 'Emissionsfaktor A4');
    const existingA5Header = Array.from(headerRow.children).find(th => th.textContent === 'Emissionsfaktor A5');
    const existingInbyggdViktHeader = Array.from(headerRow.children).find(th => th.textContent === 'Inbyggd vikt');
    const existingInkoptViktHeader = Array.from(headerRow.children).find(th => th.textContent === 'Inköpt vikt');
    
    if(!existingClimateHeader){
      // Add "Klimatresurs" header
      const climateTh = document.createElement('th');
      climateTh.textContent = 'Klimatresurs';
      headerRow.appendChild(climateTh);
    }
    
    if(!existingFactorHeader){
      // Add "Omräkningsfaktor" header
      const factorTh = document.createElement('th');
      factorTh.textContent = 'Omräkningsfaktor';
      headerRow.appendChild(factorTh);
    }
    
    if(!FactorUnit){
      // Add "Omräkningsfaktor enhet" header  
      const factorTh = document.createElement('th');
      factorTh.textContent = 'Omräkningsfaktor enhet';
      headerRow.appendChild(factorTh);
    }
    
    if(!existingWasteHeader){
      // Add "Spillfaktor" header
      const wasteTh = document.createElement('th');
      wasteTh.textContent = 'Spillfaktor';
      headerRow.appendChild(wasteTh);
    }
    
    if(!existingA1_A3Header){
      // Add "Emissionsfaktor A1-A3" header
      const a1a3Th = document.createElement('th');
      a1a3Th.textContent = 'Emissionsfaktor A1-A3';
      headerRow.appendChild(a1a3Th);
    }
    
    if(!existingA4Header){
      // Add "Emissionsfaktor A4" header
      const a4Th = document.createElement('th');
      a4Th.textContent = 'Emissionsfaktor A4';
      headerRow.appendChild(a4Th);
    }
    
    if(!existingA5Header){
      // Add "Emissionsfaktor A5" header
      const a5Th = document.createElement('th');
      a5Th.textContent = 'Emissionsfaktor A5';
      headerRow.appendChild(a5Th);
    }
    
    if(!existingInbyggdViktHeader){
      // Add "Inbyggd vikt" header
      const inbyggdTh = document.createElement('th');
      inbyggdTh.textContent = 'Inbyggd vikt';
      headerRow.appendChild(inbyggdTh);
    }
    
    if(!existingInkoptViktHeader){
      // Add "Inköpt vikt" header
      const inkoptTh = document.createElement('th');
      inkoptTh.textContent = 'Inköpt vikt';
      headerRow.appendChild(inkoptTh);
    }
    
    // Add climate impact headers
    const existingKlimatA1A3Header = Array.from(headerRow.children).find(th => th.textContent === 'Klimatpåverkan A1-A3');
    if(!existingKlimatA1A3Header){
      const klimatA1A3Th = document.createElement('th');
      klimatA1A3Th.textContent = 'Klimatpåverkan A1-A3';
      headerRow.appendChild(klimatA1A3Th);
    }
    
    const existingKlimatA4Header = Array.from(headerRow.children).find(th => th.textContent === 'Klimatpåverkan A4');
    if(!existingKlimatA4Header){
      const klimatA4Th = document.createElement('th');
      klimatA4Th.textContent = 'Klimatpåverkan A4';
      headerRow.appendChild(klimatA4Th);
    }
    
    const existingKlimatA5Header = Array.from(headerRow.children).find(th => th.textContent === 'Klimatpåverkan A5');
    if(!existingKlimatA5Header){
      const klimatA5Th = document.createElement('th');
      klimatA5Th.textContent = 'Klimatpåverkan A5';
      headerRow.appendChild(klimatA5Th);
    }

    const resourceName = resource.Name || 'Namnlös resurs';
    let conversionFactor = (resource.Conversions && resource.Conversions[0] && resource.Conversions[0].Value) || 
                            resource.ConservativeDataConversionFactor || 
                            resource.ConversionFactor || 
                            resource.Factor || 
                            resource.Omräkningsfaktor || 
                            'N/A';
    let conversionUnit = (resource.Conversions && resource.Conversions[0] && resource.Conversions[0].Unit) || 'N/A';
    let wasteFactor = resource.WasteFactor || 'N/A';
    
    // Check if conversion factor or unit is missing - if so, prompt user for manual input
    if(conversionFactor === 'N/A' || conversionUnit === 'N/A'){
      // Close climate modal before opening manual factor modal
      closeClimateModal();
      
      openManualFactorModal(resourceName, function(manualData){
        // User provided manual values, update and continue
        conversionFactor = manualData.factor;
        conversionUnit = manualData.unit;
        // wasteFactor stays as it was from API (or N/A)
        
        // Now continue with the rest of the function
        continueApplyClimateResource(resource, resourceName, conversionFactor, conversionUnit, wasteFactor, tbody, thead, headerRow, savedClimateTarget);
      });
      return; // Exit and wait for user input
    }
    
    // If we have all values, close climate modal and continue immediately
    closeClimateModal();
    continueApplyClimateResource(resource, resourceName, conversionFactor, conversionUnit, wasteFactor, tbody, thead, headerRow, savedClimateTarget);
  }
  
  function continueApplyClimateResource(resource, resourceName, conversionFactor, conversionUnit, wasteFactor, tbody, thead, headerRow, savedClimateTarget){
    // Extract emission factors from DataItems
    let a1a3Conservative = 'N/A';
    let a4Value = 'N/A';
    let a5Value = 'N/A';
    if(resource.DataItems && Array.isArray(resource.DataItems) && resource.DataItems.length > 0){
      const dataItem = resource.DataItems[0];
      if(dataItem.DataValueItems && Array.isArray(dataItem.DataValueItems)){
        const conservativeItem = dataItem.DataValueItems.find(item => item.DataModuleCode === 'A1-A3 Conservative');
        if(conservativeItem && conservativeItem.Value !== undefined){
          a1a3Conservative = conservativeItem.Value;
        }
        
        const a4Item = dataItem.DataValueItems.find(item => item.DataModuleCode === 'A4');
        if(a4Item && a4Item.Value !== undefined){
          a4Value = a4Item.Value;
        }
        
        const a5Item = dataItem.DataValueItems.find(item => item.DataModuleCode === 'A5.1');
        if(a5Item && a5Item.Value !== undefined){
          a5Value = a5Item.Value;
        }
      }
    }
    
    function addClimateToRow(tr){
      // Check if row already has climate cells
      const existingClimateCell = tr.querySelector('td[data-climate-cell="true"]');
      const existingFactorCell = tr.querySelector('td[data-factor-cell="true"]');
      const existingUnitCell = tr.querySelector('td[data-unit-cell="true"]');
      const existingWasteCell = tr.querySelector('td[data-waste-cell="true"]');
      const existingA1_A3Cell = tr.querySelector('td[data-A1_A3-cell="true"]');
      const existingA4Cell = tr.querySelector('td[data-A4-cell="true"]');
      const existingA5Cell = tr.querySelector('td[data-A5-cell="true"]');
      
      if(existingClimateCell){
        existingClimateCell.textContent = resourceName;
      } else {
        // Add new climate cell
        const climateTd = document.createElement('td');
        climateTd.textContent = resourceName;
        climateTd.setAttribute('data-climate-cell', 'true');
        tr.appendChild(climateTd);
      }
      
      if(existingFactorCell){
        existingFactorCell.textContent = conversionFactor;
      } else {
        // Add new factor cell
        const factorTd = document.createElement('td');
        factorTd.textContent = conversionFactor;
        factorTd.setAttribute('data-factor-cell', 'true');
        tr.appendChild(factorTd);
      }
      
      if(existingUnitCell){
        existingUnitCell.textContent = conversionUnit;
      } else {
        // Add new unit cell
        const unitTd = document.createElement('td');
        unitTd.textContent = conversionUnit;
        unitTd.setAttribute('data-unit-cell', 'true');
        tr.appendChild(unitTd);
      }
      
      if(existingWasteCell){
        existingWasteCell.textContent = wasteFactor;
      } else {
        // Add new waste factor cell
        const wasteTd = document.createElement('td');
        wasteTd.textContent = wasteFactor;
        wasteTd.setAttribute('data-waste-cell', 'true');
        tr.appendChild(wasteTd);
      }
      
      if(existingA1_A3Cell){
        existingA1_A3Cell.textContent = a1a3Conservative;
      } else {
        // Add new A1-A3 cell
        const a1a3Td = document.createElement('td');
        a1a3Td.textContent = a1a3Conservative;
        a1a3Td.setAttribute('data-A1_A3-cell', 'true');
        tr.appendChild(a1a3Td);
      }
      
      if(existingA4Cell){
        existingA4Cell.textContent = a4Value;
      } else {
        // Add new A4 cell
        const a4Td = document.createElement('td');
        a4Td.textContent = a4Value;
        a4Td.setAttribute('data-A4-cell', 'true');
        tr.appendChild(a4Td);
      }
      
      if(existingA5Cell){
        existingA5Cell.textContent = a5Value;
      } else {
        // Add new A5 cell
        const a5Td = document.createElement('td');
        a5Td.textContent = a5Value;
        a5Td.setAttribute('data-A5-cell', 'true');
        tr.appendChild(a5Td);
      }
      
      // Calculate Inbyggd vikt and Inköpt vikt for this row
      let inbyggdVikt = 'N/A';
      let inkoptVikt = 'N/A';
      
      // Get headers to find Volume and Net Area columns
      const allHeaders = Array.from(headerRow.children).map(th => th.textContent);
      const volumeColIndex = allHeaders.findIndex(h => String(h).toLowerCase() === 'volume');
      const netAreaColIndex = allHeaders.findIndex(h => String(h).toLowerCase() === 'net area');
      
      console.log('🔍 [applyClimate] Beräknar vikt - Unit:', conversionUnit, 'Factor:', conversionFactor, 'Waste:', wasteFactor);
      console.log('🔍 [applyClimate] Column indices - Volume:', volumeColIndex, 'NetArea:', netAreaColIndex);
      console.log('🔍 [applyClimate] Headers:', allHeaders);
      console.log('🔍 [applyClimate] Row cells before calculation:', Array.from(tr.children).length);
      
      if(conversionFactor !== 'N/A' && Number.isFinite(parseFloat(conversionFactor))){
        const factor = parseFloat(conversionFactor);
        const cells = Array.from(tr.children);
        
        console.log('🔍 [applyClimate] Factor is valid:', factor);
        console.log('🔍 [applyClimate] Number of cells:', cells.length);
        
        // Normalize unit to handle both kg/m3 and kg/m³ (with superscript)
        const normalizedUnit = String(conversionUnit).replace(/[²³]/g, function(match){
          return match === '²' ? '2' : '3';
        });
        console.log('🔍 [applyClimate] Normalized unit:', normalizedUnit);
        
        if(normalizedUnit === 'kg/m3' && volumeColIndex !== -1){
          // Inbyggd vikt = Omräkningsfaktor × Volume
          const volumeCell = cells[volumeColIndex];
          console.log('🔍 [applyClimate] Volume cell:', volumeCell?.textContent, 'at index:', volumeColIndex);
          if(volumeCell){
            const volume = parseNumberLike(volumeCell.textContent);
            console.log('🔍 [applyClimate] Parsed volume:', volume);
            if(Number.isFinite(volume)){
              inbyggdVikt = factor * volume;
              console.log('✅ [applyClimate] Inbyggd vikt calculated:', inbyggdVikt);
            }
          }
        } else if(normalizedUnit === 'kg/m2' && netAreaColIndex !== -1){
          // Inbyggd vikt = Omräkningsfaktor × Net Area
          const netAreaCell = cells[netAreaColIndex];
          console.log('🔍 [applyClimate] NetArea cell:', netAreaCell?.textContent, 'at index:', netAreaColIndex);
          if(netAreaCell){
            const netArea = parseNumberLike(netAreaCell.textContent);
            console.log('🔍 [applyClimate] Parsed netArea:', netArea);
            if(Number.isFinite(netArea)){
              inbyggdVikt = factor * netArea;
              console.log('✅ [applyClimate] Inbyggd vikt calculated:', inbyggdVikt);
            }
          }
        } else {
          console.log('❌ [applyClimate] Unit mismatch or column not found. Unit:', conversionUnit, 'Normalized:', normalizedUnit, 'VolumeIdx:', volumeColIndex, 'NetAreaIdx:', netAreaColIndex);
        }
        
        // Calculate Inköpt vikt = Inbyggd vikt × Spillfaktor
        if(inbyggdVikt !== 'N/A' && wasteFactor !== 'N/A' && Number.isFinite(parseFloat(wasteFactor))){
          const waste = parseFloat(wasteFactor);
          inkoptVikt = inbyggdVikt * waste;
          console.log('✅ [applyClimate] Inköpt vikt calculated:', inkoptVikt);
        }
      } else {
        console.log('❌ [applyClimate] Conversion factor not valid:', conversionFactor);
      }
      
      const existingInbyggdViktCell = tr.querySelector('td[data-inbyggd-vikt-cell="true"]');
      if(existingInbyggdViktCell){
        existingInbyggdViktCell.textContent = inbyggdVikt !== 'N/A' ? inbyggdVikt.toFixed(2) : 'N/A';
      } else {
        const inbyggdViktTd = document.createElement('td');
        inbyggdViktTd.textContent = inbyggdVikt !== 'N/A' ? inbyggdVikt.toFixed(2) : 'N/A';
        inbyggdViktTd.setAttribute('data-inbyggd-vikt-cell', 'true');
        tr.appendChild(inbyggdViktTd);
      }
      
      const existingInkoptViktCell = tr.querySelector('td[data-inkopt-vikt-cell="true"]');
      if(existingInkoptViktCell){
        existingInkoptViktCell.textContent = inkoptVikt !== 'N/A' ? inkoptVikt.toFixed(2) : 'N/A';
      } else {
        const inkoptViktTd = document.createElement('td');
        inkoptViktTd.textContent = inkoptVikt !== 'N/A' ? inkoptVikt.toFixed(2) : 'N/A';
        inkoptViktTd.setAttribute('data-inkopt-vikt-cell', 'true');
        tr.appendChild(inkoptViktTd);
      }
      
      // Calculate climate impact columns
      let klimatA1A3 = 'N/A';
      let klimatA4 = 'N/A';
      let klimatA5 = 'N/A';
      
      // Klimatpåverkan A1-A3 = Inbyggd vikt * Emissionsfaktor A1-A3
      if(inbyggdVikt !== 'N/A' && a1a3Conservative !== 'N/A' && Number.isFinite(parseFloat(a1a3Conservative))){
        klimatA1A3 = inbyggdVikt * parseFloat(a1a3Conservative);
      }
      
      // Klimatpåverkan A4 = Inköpt vikt * Emissionsfaktor A4
      if(inkoptVikt !== 'N/A' && a4Value !== 'N/A' && Number.isFinite(parseFloat(a4Value))){
        klimatA4 = inkoptVikt * parseFloat(a4Value);
      }
      
      // Klimatpåverkan A5 = Inköpt vikt * Emissionsfaktor A5
      if(inkoptVikt !== 'N/A' && a5Value !== 'N/A' && Number.isFinite(parseFloat(a5Value))){
        klimatA5 = inkoptVikt * parseFloat(a5Value);
      }
      
      const existingKlimatA1A3Cell = tr.querySelector('td[data-klimat-a1a3-cell="true"]');
      if(existingKlimatA1A3Cell){
        existingKlimatA1A3Cell.textContent = klimatA1A3 !== 'N/A' ? klimatA1A3.toFixed(2) : 'N/A';
      } else {
        const klimatA1A3Td = document.createElement('td');
        klimatA1A3Td.textContent = klimatA1A3 !== 'N/A' ? klimatA1A3.toFixed(2) : 'N/A';
        klimatA1A3Td.setAttribute('data-klimat-a1a3-cell', 'true');
        tr.appendChild(klimatA1A3Td);
      }
      
      const existingKlimatA4Cell = tr.querySelector('td[data-klimat-a4-cell="true"]');
      if(existingKlimatA4Cell){
        existingKlimatA4Cell.textContent = klimatA4 !== 'N/A' ? klimatA4.toFixed(2) : 'N/A';
      } else {
        const klimatA4Td = document.createElement('td');
        klimatA4Td.textContent = klimatA4 !== 'N/A' ? klimatA4.toFixed(2) : 'N/A';
        klimatA4Td.setAttribute('data-klimat-a4-cell', 'true');
        tr.appendChild(klimatA4Td);
      }
      
      const existingKlimatA5Cell = tr.querySelector('td[data-klimat-a5-cell="true"]');
      if(existingKlimatA5Cell){
        existingKlimatA5Cell.textContent = klimatA5 !== 'N/A' ? klimatA5.toFixed(2) : 'N/A';
      } else {
        const klimatA5Td = document.createElement('td');
        klimatA5Td.textContent = klimatA5 !== 'N/A' ? klimatA5.toFixed(2) : 'N/A';
        klimatA5Td.setAttribute('data-klimat-a5-cell', 'true');
        tr.appendChild(klimatA5Td);
      }
      
      // Save climate data for this row
      // Use original row data if available, otherwise extract from DOM
      const rowData = tr._originalRowData || getRowDataFromTr(tr);
      if(rowData){
        const layerChildOf = tr.getAttribute('data-layer-child-of');
        const signature = getRowSignature(rowData, layerChildOf);
        climateData.set(signature, { name: resourceName, factor: conversionFactor, unit: conversionUnit, waste: wasteFactor, a1a3: a1a3Conservative, a4: a4Value, a5: a5Value });
      }
    }
    
    if(savedClimateTarget.type === 'row' && savedClimateTarget.rowEl){
      addClimateToRow(savedClimateTarget.rowEl);
      
      // Update parent row's weight sums if this row belongs to a group
      const groupKey = savedClimateTarget.rowEl.getAttribute('data-group-child-of');
      if(groupKey){
        updateGroupWeightSums(groupKey, tbody);
      }
    } else if(savedClimateTarget.type === 'group' && savedClimateTarget.key != null){
      const rows = Array.from(tbody.querySelectorAll('tr[data-group-child-of="' + CSS.escape(savedClimateTarget.key) + '"]'));
      rows.forEach(addClimateToRow);
      
      // Update parent row's weight sums after applying climate to all children
      updateGroupWeightSums(savedClimateTarget.key, tbody);
    }
    
    // Re-apply filters to keep visibility consistent
    applyFilters();
    
    // Update climate summary
    setTimeout(() => updateClimateSummary(), 100);
  }
  
  // Helper function to update weight sums for a group parent
  function updateGroupWeightSums(groupKey, tbody){
    console.log('🔍 [updateGroupWeightSums] Called with groupKey:', groupKey);
    const parentTr = tbody.querySelector(`tr.group-parent[data-group-key="${CSS.escape(groupKey)}"]`);
    console.log('🔍 [updateGroupWeightSums] Found parent:', !!parentTr);
    if(!parentTr) return;
    
    const childRows = Array.from(tbody.querySelectorAll(`tr[data-group-child-of="${CSS.escape(groupKey)}"]`));
    console.log('🔍 [updateGroupWeightSums] Number of children:', childRows.length);
    
    let sumInbyggdVikt = 0;
    let sumInkoptVikt = 0;
    let countInbyggd = 0;
    let countInkopt = 0;
    
    let sumKlimatA1A3 = 0;
    let sumKlimatA4 = 0;
    let sumKlimatA5 = 0;
    let countKlimatA1A3 = 0;
    let countKlimatA4 = 0;
    let countKlimatA5 = 0;
    
    childRows.forEach(childTr => {
      const inbyggdCell = childTr.querySelector('td[data-inbyggd-vikt-cell="true"]');
      const inkoptCell = childTr.querySelector('td[data-inkopt-vikt-cell="true"]');
      
      if(inbyggdCell){
        const val = parseNumberLike(inbyggdCell.textContent);
        console.log('🔍 [updateGroupWeightSums] Inbyggd cell value:', inbyggdCell.textContent, 'Parsed:', val);
        if(Number.isFinite(val)){
          sumInbyggdVikt += val;
          countInbyggd++;
        }
      }
      
      if(inkoptCell){
        const val = parseNumberLike(inkoptCell.textContent);
        console.log('🔍 [updateGroupWeightSums] Inkopt cell value:', inkoptCell.textContent, 'Parsed:', val);
        if(Number.isFinite(val)){
          sumInkoptVikt += val;
          countInkopt++;
        }
      }
      
      const klimatA1A3Cell = childTr.querySelector('td[data-klimat-a1a3-cell="true"]');
      if(klimatA1A3Cell){
        const val = parseNumberLike(klimatA1A3Cell.textContent);
        if(Number.isFinite(val)){
          sumKlimatA1A3 += val;
          countKlimatA1A3++;
        }
      }
      
      const klimatA4Cell = childTr.querySelector('td[data-klimat-a4-cell="true"]');
      if(klimatA4Cell){
        const val = parseNumberLike(klimatA4Cell.textContent);
        if(Number.isFinite(val)){
          sumKlimatA4 += val;
          countKlimatA4++;
        }
      }
      
      const klimatA5Cell = childTr.querySelector('td[data-klimat-a5-cell="true"]');
      if(klimatA5Cell){
        const val = parseNumberLike(klimatA5Cell.textContent);
        if(Number.isFinite(val)){
          sumKlimatA5 += val;
          countKlimatA5++;
        }
      }
    });
    
    console.log('🔍 [updateGroupWeightSums] Sums - Inbyggd:', sumInbyggdVikt, 'count:', countInbyggd, 'Inkopt:', sumInkoptVikt, 'count:', countInkopt);
    console.log('🔍 [updateGroupWeightSums] Climate Sums - A1-A3:', sumKlimatA1A3, 'A4:', sumKlimatA4, 'A5:', sumKlimatA5);
    
    // Find the column indices for Inbyggd vikt and Inköpt vikt from headers
    const table = parentTr.closest('table');
    if(!table) return;
    const thead = table.querySelector('thead');
    if(!thead) return;
    const headerRow = thead.querySelector('tr:first-child');
    if(!headerRow) return;
    
    const headers = Array.from(headerRow.children).map(th => th.textContent);
    const inbyggdViktColIndex = headers.findIndex(h => h === 'Inbyggd vikt');
    const inkoptViktColIndex = headers.findIndex(h => h === 'Inköpt vikt');
    const klimatA1A3ColIndex = headers.findIndex(h => h === 'Klimatpåverkan A1-A3');
    const klimatA4ColIndex = headers.findIndex(h => h === 'Klimatpåverkan A4');
    const klimatA5ColIndex = headers.findIndex(h => h === 'Klimatpåverkan A5');
    
    console.log('🔍 [updateGroupWeightSums] Column indices - Inbyggd:', inbyggdViktColIndex, 'Inkopt:', inkoptViktColIndex);
    console.log('🔍 [updateGroupWeightSums] Climate Column indices - A1-A3:', klimatA1A3ColIndex, 'A4:', klimatA4ColIndex, 'A5:', klimatA5ColIndex);
    
    // Update parent's cells by index
    const parentCells = Array.from(parentTr.children);
    console.log('🔍 [updateGroupWeightSums] Parent has', parentCells.length, 'cells, need at least', Math.max(inbyggdViktColIndex, inkoptViktColIndex, klimatA1A3ColIndex, klimatA4ColIndex, klimatA5ColIndex) + 1);
    
    // If parent doesn't have enough cells, add them
    const neededCells = Math.max(inbyggdViktColIndex, inkoptViktColIndex, klimatA1A3ColIndex, klimatA4ColIndex, klimatA5ColIndex) + 1;
    while(parentCells.length < neededCells){
      const td = document.createElement('td');
      td.textContent = '';
      parentTr.appendChild(td);
      parentCells.push(td);
      console.log('🔧 [updateGroupWeightSums] Added missing cell, now has', parentCells.length, 'cells');
    }
    
    if(inbyggdViktColIndex !== -1 && parentCells[inbyggdViktColIndex]){
      const cell = parentCells[inbyggdViktColIndex];
      cell.textContent = countInbyggd > 0 ? sumInbyggdVikt.toFixed(2) : '';
      // Also add the attribute for future lookups
      cell.setAttribute('data-sum-inbyggd-vikt', 'true');
      console.log('✅ [updateGroupWeightSums] Updated Inbyggd cell to:', cell.textContent);
    }
    
    if(inkoptViktColIndex !== -1 && parentCells[inkoptViktColIndex]){
      const cell = parentCells[inkoptViktColIndex];
      cell.textContent = countInkopt > 0 ? sumInkoptVikt.toFixed(2) : '';
      // Also add the attribute for future lookups
      cell.setAttribute('data-sum-inkopt-vikt', 'true');
      console.log('✅ [updateGroupWeightSums] Updated Inkopt cell to:', cell.textContent);
    }
    
    if(klimatA1A3ColIndex !== -1 && parentCells[klimatA1A3ColIndex]){
      const cell = parentCells[klimatA1A3ColIndex];
      cell.textContent = countKlimatA1A3 > 0 ? sumKlimatA1A3.toFixed(2) : '';
      // Also add the attribute for future lookups
      cell.setAttribute('data-sum-klimat-a1a3', 'true');
      console.log('✅ [updateGroupWeightSums] Updated Klimat A1-A3 cell to:', cell.textContent);
    }
    
    if(klimatA4ColIndex !== -1 && parentCells[klimatA4ColIndex]){
      const cell = parentCells[klimatA4ColIndex];
      cell.textContent = countKlimatA4 > 0 ? sumKlimatA4.toFixed(2) : '';
      // Also add the attribute for future lookups
      cell.setAttribute('data-sum-klimat-a4', 'true');
      console.log('✅ [updateGroupWeightSums] Updated Klimat A4 cell to:', cell.textContent);
    }
    
    if(klimatA5ColIndex !== -1 && parentCells[klimatA5ColIndex]){
      const cell = parentCells[klimatA5ColIndex];
      cell.textContent = countKlimatA5 > 0 ? sumKlimatA5.toFixed(2) : '';
      // Also add the attribute for future lookups
      cell.setAttribute('data-sum-klimat-a5', 'true');
      console.log('✅ [updateGroupWeightSums] Updated Klimat A5 cell to:', cell.textContent);
    }
  }
  
  // Function to update climate impact summary
  function updateClimateSummary(){
    const table = getTable();
    if(!table){
      // Hide summary if no table
      const climateSummary = document.getElementById('climateSummary');
      if(climateSummary) climateSummary.style.display = 'none';
      return;
    }
    
    const tbody = table.querySelector('tbody');
    if(!tbody) return;
    
    // Get all rows (including parents)
    const allRows = Array.from(tbody.querySelectorAll('tr'));
    
    // Keep track of which rows we've already counted (via their parent's sum)
    const countedViaParent = new Set();
    
    let totalA1A3 = 0;
    let totalA4 = 0;
    let totalA5 = 0;
    let hasAnyData = false;
    
    // First pass: identify all parent rows with visible children and mark their children as counted
    allRows.forEach(tr => {
      if(tr.style.display === 'none') return;
      
      const isParent = tr.classList.contains('group-parent') || tr.classList.contains('layer-parent');
      if(!isParent) return;
      
      const layerKey = tr.getAttribute('data-layer-key');
      const groupKey = tr.getAttribute('data-group-key');
      
      let hasVisibleChildren = false;
      let childrenList = [];
      
      if(layerKey){
        // Check for layer children
        const children = Array.from(tbody.querySelectorAll(`tr[data-parent-key="${CSS.escape(layerKey)}"]`));
        childrenList = children;
        hasVisibleChildren = children.some(child => child.style.display !== 'none');
      }
      if(groupKey && !hasVisibleChildren){
        // Check for group children (only direct children, not grandchildren)
        const children = Array.from(tbody.querySelectorAll(`tr[data-group-child-of="${CSS.escape(groupKey)}"]:not([data-parent-key])`));
        childrenList = children;
        hasVisibleChildren = children.some(child => child.style.display !== 'none');
      }
      
      if(hasVisibleChildren){
        // Mark all children (and their descendants) as counted via this parent
        childrenList.forEach(child => {
          countedViaParent.add(child);
          // Also mark any descendants of this child
          const childLayerKey = child.getAttribute('data-layer-key');
          if(childLayerKey){
            const grandchildren = tbody.querySelectorAll(`tr[data-parent-key="${CSS.escape(childLayerKey)}"]`);
            grandchildren.forEach(gc => countedViaParent.add(gc));
          }
        });
      }
    });
    
    // Second pass: count climate data
    allRows.forEach(tr => {
      // Skip hidden rows
      if(tr.style.display === 'none') return;
      
      // Skip rows that are already counted via their parent's sum
      if(countedViaParent.has(tr)) return;
      
      // Check if this is a parent row
      const isParent = tr.classList.contains('group-parent') || tr.classList.contains('layer-parent');
      
      if(isParent){
        const layerKey = tr.getAttribute('data-layer-key');
        const groupKey = tr.getAttribute('data-group-key');
        
        let hasVisibleChildren = false;
        if(layerKey){
          const children = tbody.querySelectorAll(`tr[data-parent-key="${CSS.escape(layerKey)}"]`);
          hasVisibleChildren = Array.from(children).some(child => child.style.display !== 'none');
        }
        if(groupKey && !hasVisibleChildren){
          const children = tbody.querySelectorAll(`tr[data-group-child-of="${CSS.escape(groupKey)}"]:not([data-parent-key])`);
          hasVisibleChildren = Array.from(children).some(child => child.style.display !== 'none');
        }
        
        if(hasVisibleChildren){
          // Use sum cells if they exist
          const a1a3SumCell = tr.querySelector('td[data-sum-klimat-a1a3="true"]');
          const a4SumCell = tr.querySelector('td[data-sum-klimat-a4="true"]');
          const a5SumCell = tr.querySelector('td[data-sum-klimat-a5="true"]');
          
          if(a1a3SumCell){
            const val = parseNumberLike(a1a3SumCell.textContent);
            if(Number.isFinite(val)){
              totalA1A3 += val;
              hasAnyData = true;
            }
          }
          
          if(a4SumCell){
            const val = parseNumberLike(a4SumCell.textContent);
            if(Number.isFinite(val)){
              totalA4 += val;
              hasAnyData = true;
            }
          }
          
          if(a5SumCell){
            const val = parseNumberLike(a5SumCell.textContent);
            if(Number.isFinite(val)){
              totalA5 += val;
              hasAnyData = true;
            }
          }
        } else {
          // No visible children, use parent's own climate data if it exists
          console.log('📊 Parent with no visible children, using own data');
          
          // Try sum cells first (for group parents), then regular climate cells
          let a1a3Cell = tr.querySelector('td[data-sum-klimat-a1a3="true"]');
          if(!a1a3Cell) a1a3Cell = tr.querySelector('td[data-klimat-a1a3-cell="true"]');
          
          let a4Cell = tr.querySelector('td[data-sum-klimat-a4="true"]');
          if(!a4Cell) a4Cell = tr.querySelector('td[data-klimat-a4-cell="true"]');
          
          let a5Cell = tr.querySelector('td[data-sum-klimat-a5="true"]');
          if(!a5Cell) a5Cell = tr.querySelector('td[data-klimat-a5-cell="true"]');
          
          if(a1a3Cell){
            const val = parseNumberLike(a1a3Cell.textContent);
            console.log('  A1-A3:', a1a3Cell.textContent, '→', val);
            if(Number.isFinite(val)){
              totalA1A3 += val;
              hasAnyData = true;
            }
          }
          
          if(a4Cell){
            const val = parseNumberLike(a4Cell.textContent);
            console.log('  A4:', a4Cell.textContent, '→', val);
            if(Number.isFinite(val)){
              totalA4 += val;
              hasAnyData = true;
            }
          }
          
          if(a5Cell){
            const val = parseNumberLike(a5Cell.textContent);
            console.log('  A5:', a5Cell.textContent, '→', val);
            if(Number.isFinite(val)){
              totalA5 += val;
              hasAnyData = true;
            }
          }
        }
      } else {
        // For non-parent rows that weren't counted via parent, count them directly
        const a1a3Cell = tr.querySelector('td[data-klimat-a1a3-cell="true"]');
        const a4Cell = tr.querySelector('td[data-klimat-a4-cell="true"]');
        const a5Cell = tr.querySelector('td[data-klimat-a5-cell="true"]');
        
        if(a1a3Cell){
          const val = parseNumberLike(a1a3Cell.textContent);
          if(Number.isFinite(val)){
            totalA1A3 += val;
            hasAnyData = true;
          }
        }
        
        if(a4Cell){
          const val = parseNumberLike(a4Cell.textContent);
          if(Number.isFinite(val)){
            totalA4 += val;
            hasAnyData = true;
          }
        }
        
        if(a5Cell){
          const val = parseNumberLike(a5Cell.textContent);
          if(Number.isFinite(val)){
            totalA5 += val;
            hasAnyData = true;
          }
        }
      }
    });
    
    const total = totalA1A3 + totalA4 + totalA5;
    
    // Update summary display
    const climateSummary = document.getElementById('climateSummary');
    const summaryA1A3 = document.getElementById('summaryA1A3');
    const summaryA4 = document.getElementById('summaryA4');
    const summaryA5 = document.getElementById('summaryA5');
    const summaryTotal = document.getElementById('summaryTotal');
    
    if(hasAnyData){
      // Show and update summary
      if(climateSummary) climateSummary.style.display = 'block';
      if(summaryA1A3) summaryA1A3.textContent = totalA1A3.toFixed(2) + ' kg CO₂e';
      if(summaryA4) summaryA4.textContent = totalA4.toFixed(2) + ' kg CO₂e';
      if(summaryA5) summaryA5.textContent = totalA5.toFixed(2) + ' kg CO₂e';
      if(summaryTotal) summaryTotal.textContent = total.toFixed(2) + ' kg CO₂e';
    } else {
      // Hide summary if no data
      if(climateSummary) climateSummary.style.display = 'none';
    }
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

  // ============ PROJECT SAVE/LOAD FUNCTIONS ============
  
  // Clean up old badges and layer data before applying new ones
  function cleanupOldBadgesAndLayers(){
    const table = getTable();
    if(!table) return;
    
    const tbody = table.querySelector('tbody');
    if(!tbody) return;
    
    // Remove all existing badges
    const existingBadges = tbody.querySelectorAll('.badge-new');
    existingBadges.forEach(badge => badge.remove());
    
    // Remove all existing layer children
    const layerChildren = tbody.querySelectorAll('tr[data-layer-child-of]');
    layerChildren.forEach(child => child.remove());
    
    // Remove layer attributes from parent rows
    const parentRows = tbody.querySelectorAll('tr[data-layer-key]');
    parentRows.forEach(tr => {
      tr.removeAttribute('data-layer-key');
      tr.classList.remove('layer-parent');
      
      // Remove layer labels from first data cell
      const firstDataTd = tr.querySelector('td:nth-child(2)');
      if(firstDataTd){
        const toggle = firstDataTd.querySelector('.group-toggle');
        const spans = firstDataTd.querySelectorAll('span');
        if(toggle) toggle.remove();
        spans.forEach(span => {
          if(span.textContent.includes('skikt')) span.remove();
        });
      }
    });
    
    console.log('🧹 Rensade bort gamla badges och skikt');
  }
  
  // Apply saved layers and climate to all rows
  function applySavedLayersAndClimate(){
    const table = getTable();
    if(!table) return;
    
    const tbody = table.querySelector('tbody');
    if(!tbody) return;
    
    // Apply layers and climate to all rows
    const allRows = Array.from(tbody.querySelectorAll('tr'));
    allRows.forEach(tr => {
      const rowData = tr._originalRowData;
      if(rowData){
        applySavedLayers(tr, rowData);
        applySavedClimate(tr, rowData);
      }
    });
    
    console.log('✅ Applicerade sparade skikt och klimatdata');
  }
  
  function getProjectData(){
    if(!lastRows || !lastHeaders){
      return null;
    }
    
    // Get the current table HTML structure
    const table = getTable();
    const tableHTML = table ? table.outerHTML : '';
    
    // Create project data object
    return {
      version: '1.2',
      timestamp: new Date().toISOString(),
      originalFileName: originalFileName,
      headers: lastHeaders,
      rows: lastRows.slice(1), // Exclude header row
      tableHTML: tableHTML, // Save the actual table structure
      layerData: Array.from(layerData.entries()).map(([key, value]) => ({
        key,
        count: value.count,
        thicknesses: value.thicknesses,
        layerKey: value.layerKey
      })),
      climateData: Array.from(climateData.entries()).map(([key, value]) => ({
        key,
        resourceName: value
      })),
      // Include undo/redo state
      undoStack: undoStack.slice(-10), // Save last 10 undo steps
      redoStack: redoStack.slice(-10), // Save last 10 redo steps
      // Include current filter and group by values
      filterValue: filterInput ? filterInput.value : '',
      groupByValue: groupBySelect ? groupBySelect.value : ''
    };
  }
  
  async function saveProjectWithDialog(){
    const projectData = getProjectData();
    if(!projectData){
      alert('Ingen data att spara. Ladda en fil först.');
      return;
    }
    
    const jsonStr = JSON.stringify(projectData, null, 2);
    
    // Check if File System Access API is available
    if('showSaveFilePicker' in window){
      try {
        const baseName = originalFileName ? originalFileName.replace(/\.[^.]+$/, '') : 'projekt';
        const fileHandle = await window.showSaveFilePicker({
          suggestedName: `${baseName}_projekt.json`,
          types: [{
            description: 'JSON Project File',
            accept: { 'application/json': ['.json'] }
          }]
        });
        
        const writable = await fileHandle.createWritable();
        await writable.write(jsonStr);
        await writable.close();
        
        // Store file handle for quick save
        savedFileHandle = fileHandle;
        
        // Update save button text to show file is saved
        if(saveProjectBtn){
          const fileName = fileHandle.name;
          saveProjectBtn.textContent = `💾 Spara`;
          saveProjectBtn.title = `Sparad som: ${fileName}`;
        }
        
        console.log('✅ Projekt sparat:', fileHandle.name);
      } catch(err){
        if(err.name !== 'AbortError'){
          console.error('Fel vid sparande:', err);
          alert('Kunde inte spara filen: ' + err.message);
        }
      }
    } else {
      // Fallback for browsers that don't support File System Access API
      const blob = new Blob([jsonStr], { type: 'application/json' });
      const url = URL.createObjectURL(blob);
      const a = document.createElement('a');
      a.href = url;
      const baseName = originalFileName ? originalFileName.replace(/\.[^.]+$/, '') : 'projekt';
      a.download = `${baseName}_projekt.json`;
      document.body.appendChild(a);
      a.click();
      document.body.removeChild(a);
      URL.revokeObjectURL(url);
      console.log('✅ Projekt sparat (fallback)');
    }
  }
  
  async function saveProject(){
    // Quick save to existing file
    if(savedFileHandle){
      try {
        const projectData = getProjectData();
        if(!projectData){
          alert('Ingen data att spara. Ladda en fil först.');
          return;
        }
        
        const jsonStr = JSON.stringify(projectData, null, 2);
        const writable = await savedFileHandle.createWritable();
        await writable.write(jsonStr);
        await writable.close();
        
        console.log('✅ Projekt sparat snabbt:', savedFileHandle.name);
        
        // Visual feedback
        const originalText = saveProjectBtn.textContent;
        saveProjectBtn.textContent = '✅ Sparat!';
        setTimeout(() => {
          saveProjectBtn.textContent = originalText;
        }, 1500);
      } catch(err){
        console.error('Fel vid snabbsparande:', err);
        // If quick save fails, fall back to save dialog
        saveProjectWithDialog();
      }
    } else {
      // No saved file yet, show save dialog
      saveProjectWithDialog();
    }
  }
  
  function loadProject(file){
    const reader = new FileReader();
    
    reader.onload = function(e){
      try {
        const projectData = JSON.parse(e.target.result);
        
        // Validate project data
        if(!projectData.version || !projectData.headers || !projectData.rows){
          alert('Ogiltig projektfil. Kontrollera att filen är en giltig JSON-projektfil.');
          return;
        }
        
        console.log('📂 Laddar projekt:', projectData);
        
        // Restore basic data
        originalFileName = projectData.originalFileName || 'unknown';
        lastHeaders = projectData.headers;
        lastRows = [projectData.headers, ...projectData.rows];
        
        // Clear existing data
        layerData.clear();
        climateData.clear();
        
        // Restore layer data
        if(projectData.layerData && Array.isArray(projectData.layerData)){
          projectData.layerData.forEach(item => {
            layerData.set(item.key, {
              count: item.count,
              thicknesses: item.thicknesses,
              layerKey: item.layerKey
            });
          });
        }
        
        // Restore climate data
        if(projectData.climateData && Array.isArray(projectData.climateData)){
          projectData.climateData.forEach(item => {
            climateData.set(item.key, item.resourceName);
          });
        }
        
        // Restore undo/redo stacks (if available)
        if(projectData.undoStack && Array.isArray(projectData.undoStack)){
          undoStack = projectData.undoStack;
        } else {
          undoStack = [];
        }
        
        if(projectData.redoStack && Array.isArray(projectData.redoStack)){
          redoStack = projectData.redoStack;
        } else {
          redoStack = [];
        }
        
        // Restore filter and group by values
        if(projectData.filterValue && filterInput){
          filterInput.value = projectData.filterValue;
        }
        
        if(projectData.groupByValue && groupBySelect){
          groupBySelect.value = projectData.groupByValue;
        }
        
        console.log('✅ Projekt laddat. Rader:', lastRows.length, 'Skikt:', layerData.size, 'Klimat:', climateData.size, 'Undo:', undoStack.length, 'Redo:', redoStack.length);
        
        // Clear saved file handle when loading a different file
        savedFileHandle = null;
        if(saveProjectBtn){
          saveProjectBtn.textContent = '💾 Spara';
          saveProjectBtn.title = '';
        }
        
        // Update undo/redo button states
        updateUndoRedoButtons();
        
        // Restore the table structure if available (version 1.2+)
        if(projectData.tableHTML){
          console.log('🔄 Återställer sparad tabellstruktur');
          output.innerHTML = projectData.tableHTML;
          
          // Re-attach event listeners to the restored table
          reattachTableEventListeners();
          
          // Update climate summary
          setTimeout(() => updateClimateSummary(), 100);
        } else {
          // Fallback for older project files - render table normally
          console.log('🔄 Återställer tabell från rådata (äldre format)');
          renderTableWithOptionalGrouping(lastRows);
          
          // Clean up old badges and layers before applying new ones
          setTimeout(() => {
            cleanupOldBadgesAndLayers();
            applySavedLayersAndClimate();
            updateClimateSummary();
          }, 100);
        }
        
      } catch(err){
        console.error('Fel vid laddning av projekt:', err);
        alert('Kunde inte ladda projektfilen. Felmeddelande: ' + err.message);
      }
    };
    
    reader.readAsText(file);
  }
  
  // Event listeners for save/load buttons
  if(saveProjectBtn){
    saveProjectBtn.addEventListener('click', saveProject);
  }
  
  if(saveAsProjectBtn){
    saveAsProjectBtn.addEventListener('click', saveProjectWithDialog);
  }
  
  if(loadProjectBtn){
    loadProjectBtn.addEventListener('click', function(){
      projectFileInput.click();
    });
  }
  
  if(projectFileInput){
    projectFileInput.addEventListener('change', function(e){
      const file = e.target.files[0];
      if(file){
        loadProject(file);
      }
      // Reset input so same file can be loaded again
      e.target.value = '';
    });
  }
})();
