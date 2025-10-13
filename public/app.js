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
  const multiLayerClimateNextBtn = document.getElementById('multiLayerClimateNext');
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
  
  function addNewRow(){
    const table = getTable();
    if(!table) {
      alert('Ladda en tabell först');
      return;
    }
    
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
          }
        }
      } else {
        // Not yet layered, clear the inputs
        if(layerCountInput) layerCountInput.value = '2';
        if(layerThicknessesInput) layerThicknessesInput.value = '';
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
              }
            }
          } else {
            // Not yet layered, clear the inputs
            if(layerCountInput) layerCountInput.value = '2';
            if(layerThicknessesInput) layerThicknessesInput.value = '';
          }
        }
      }
    }
    
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
      
      console.log('🔧 [LayerMapClimate] Applying layers - count:', count, 'thicknesses:', thicknesses);
      applyLayerSplit(count, thicknesses);
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
    
    multiLayerClimateTarget = { 
      layerRows, 
      groupKey, 
      uniqueLayers,
      currentLayerIndex: 0,
      selectedResources: new Map() // layerNumber -> resource
    };
    
    showNextLayerSelection();
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
  
  function applyAllLayerResources(){
    if(!multiLayerClimateTarget) return;
    
    const { layerRows, groupKey, selectedResources, uniqueLayers } = multiLayerClimateTarget;
    
    // Build a map from layer number to the layer key pattern
    const layerKeyPatterns = new Map();
    uniqueLayers.forEach(layerNum => {
      layerKeyPatterns.set(layerNum, groupKey + '_Layer_' + layerNum);
    });
    
    console.log('🔍 [applyAllLayerResources] Layer key patterns:', Array.from(layerKeyPatterns.entries()));
    
    // Apply the appropriate resource to each layer row based on its layer key
    layerRows.forEach(row => {
      const layerKey = row.dataset.layerKey || '';
      if(layerKey){
        // Find which layer number this key corresponds to
        const match = layerKey.match(/_Layer_(\d+)$/);
        if(match){
          const layerNumber = parseInt(match[1], 10);
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
    // Reset button text
    if(multiLayerClimateNextBtn){
      multiLayerClimateNextBtn.textContent = 'Nästa';
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
  
  if(multiLayerClimateNextBtn){
    multiLayerClimateNextBtn.addEventListener('click', function(){
      if(!multiLayerClimateTarget) return;
      
      const select = document.getElementById('currentLayerSelect');
      if(!select) return;
      
      const selectedIndex = select.value;
      if(selectedIndex === ''){
        alert('Välj en klimatresurs först');
        return;
      }
      
      const { uniqueLayers, currentLayerIndex, selectedResources } = multiLayerClimateTarget;
      const currentLayerNum = uniqueLayers[currentLayerIndex];
      const resource = window.climateResources[selectedIndex];
      
      if(resource){
        // Save the selection
        selectedResources.set(currentLayerNum, resource);
        
        // Move to next layer
        multiLayerClimateTarget.currentLayerIndex++;
        
        // Show next layer or finish
        if(multiLayerClimateTarget.currentLayerIndex < uniqueLayers.length){
          // Update button text for last layer
          if(multiLayerClimateTarget.currentLayerIndex === uniqueLayers.length - 1){
            multiLayerClimateNextBtn.textContent = 'Klar';
          }
          showNextLayerSelection();
        } else {
          // All done, apply all resources
          applyAllLayerResources();
        }
      }
    });
  }
  if(climateApplyBtn){
    climateApplyBtn.addEventListener('click', function(){
      const selectedIndex = climateResourceSelect && climateResourceSelect.value;
      if(selectedIndex !== '' && window.climateResources && window.climateResources[selectedIndex]){
        const resource = window.climateResources[selectedIndex];
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

  function applyLayerSplit(count, thicknesses){
    if(!layerTarget){ return; }
    const table = getTable(); if(!table) return;
    const tbody = table.querySelector('tbody'); if(!tbody) return;

    function cloneRowWithMultiplier(srcTr, multiplier, layerIndex, totalLayers, layerThickness){
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
      const headerTexts = Array.from(table.querySelectorAll('thead th')).map(th => th.textContent);
      const idxNetArea = headerTexts.findIndex(h => String(h).toLowerCase() === 'net area');
      const idxVolume = headerTexts.findIndex(h => String(h).toLowerCase() === 'volume');
      const idxCount = headerTexts.findIndex(h => String(h).toLowerCase() === 'count');
      
      // Read Net Area BEFORE scaling it (for volume calculation)
      // Count backwards from the original end to find the right cells
      const originalCellCount = tds.length;
      const countCellIdx = originalCellCount - 1; // Last original cell = Count (tds[12])
      const volumeCellIdx = countCellIdx - 1; // Second to last = Volume (tds[11])
      const netAreaCellIdx = volumeCellIdx - 4; // Net Area is 4 cells before Volume (tds[7])
      
      let originalNetArea = null;
      const netAreaTd = tds[netAreaCellIdx];
      if(netAreaTd){
        originalNetArea = parseNumberLike(netAreaTd.textContent);
      }
      
      // Add badge into first data cell (after action column)
      const firstDataTd = tds[1];
      if(firstDataTd){
        const badge = document.createElement('span'); badge.className = 'badge-new'; badge.textContent = 'Skikt ' + (layerIndex + 1) + '/' + totalLayers;
        firstDataTd.insertBefore(badge, firstDataTd.firstChild);
      }
      
      // Don't scale Net Area or Count - they remain unchanged
      // Instead, update Thickness column with the layer thickness
      
      // Find Thickness column: Net Area(7) -> Length(8) -> Thickness(9) -> Height(10) -> Volume(11)
      const thicknessCellIdx = volumeCellIdx - 2; // Thickness is 2 cells before Volume
      
      // For Volume: if we have thickness specified, calculate Volume = Net Area × thickness (in meters)
      if(layerThickness && originalNetArea !== null && Number.isFinite(originalNetArea)){
        // Update Thickness cell with the layer thickness (convert from mm to m)
        const thicknessTd = tds[thicknessCellIdx];
        if(thicknessTd){
          const thicknessInMeters = layerThickness / 1000;
          thicknessTd.textContent = String(thicknessInMeters);
        }
        
        // Calculate and update Volume
        const volumeTd = tds[volumeCellIdx];
        if(volumeTd){
          // Thickness is in mm, convert to meters for volume calculation
          const thicknessInMeters = layerThickness / 1000;
          const newVolume = originalNetArea * thicknessInMeters;
          volumeTd.textContent = String(newVolume);
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
        scaleCell(netAreaCellIdx);
        scaleCell(volumeCellIdx);
        scaleCell(countCellIdx);
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
      const fragments = multipliers.map((m, i) => {
        // Pass the actual thickness for this layer if available
        const layerThickness = thicknesses.length > 0 ? thicknesses[i] : undefined;
        const clone = cloneRowWithMultiplier(tr, m, i, multipliers.length, layerThickness);
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
      
      // Generate layer keys for each layer NUMBER (not each row)
      // So all rows in layer 1 share the same key, all rows in layer 2 share another key, etc.
      const layerKeys = Array.from({ length: count }, (_, i) => 
        layerTarget.key + '_Layer_' + (i + 1)
      );
      console.log('🔧 Genererade layerKeys:', layerKeys);
      
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
        const fragments = multipliers.map((m, i) => {
          // Pass the actual thickness for this layer if available
          const layerThickness = thicknesses.length > 0 ? thicknesses[i] : undefined;
          const clone = cloneRowWithMultiplier(row, m, i, multipliers.length, layerThickness);
          // Mark as child of this row's layer
          clone.setAttribute('data-layer-child-of', rowLayerKey);
          // Also inherit parent's group membership if it exists
          if(parentGroupKey){
            clone.setAttribute('data-group-child-of', parentGroupKey);
          }
          // Set immediate parent for toggle
          clone.setAttribute('data-parent-key', rowLayerKey);
          // Set the SHARED layer key for this layer number
          clone.setAttribute('data-layer-key', layerKeys[i]);
          // Preserve original row data
          if(originalRowData){
            clone._originalRowData = originalRowData;
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
    // Re-apply filters to keep visibility consistent
    applyFilters();
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
  
  function getProjectData(){
    if(!lastRows || !lastHeaders){
      return null;
    }
    
    // Create project data object
    return {
      version: '1.0',
      timestamp: new Date().toISOString(),
      originalFileName: originalFileName,
      headers: lastHeaders,
      rows: lastRows.slice(1), // Exclude header row
      layerData: Array.from(layerData.entries()).map(([key, value]) => ({
        key,
        count: value.count,
        thicknesses: value.thicknesses,
        layerKey: value.layerKey
      })),
      climateData: Array.from(climateData.entries()).map(([key, value]) => ({
        key,
        resourceName: value
      }))
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
        
        console.log('✅ Projekt laddat. Rader:', lastRows.length, 'Skikt:', layerData.size, 'Klimat:', climateData.size);
        
        // Clear saved file handle when loading a different file
        savedFileHandle = null;
        if(saveProjectBtn){
          saveProjectBtn.textContent = '💾 Spara';
          saveProjectBtn.title = '';
        }
        
        // Render the table
        renderTableWithOptionalGrouping(lastRows);
        
        // Apply saved layers and climate mappings
        setTimeout(() => {
          applySavedLayers();
          applySavedClimate();
          updateClimateSummary();
        }, 100);
        
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
