$(document).ready(function() {
  let table = null, allData = [], nameConflicts = new Set(), dupLicUsers = new Set();

  const DEBOUNCE_DELAY = 350;
  let multiSearchDebounceTimer;

  const searchableColumnsConfig = [
    { index: 0, title: 'ID', dataProp: 'Id', useDropdown: false },
    { index: 1, title: 'Name', dataProp: 'DisplayName', useDropdown: false },
    { index: 2, title: 'Email', dataProp: 'Email', useDropdown: false },
    { index: 3, title: 'Job Title', dataProp: 'JobTitle', useDropdown: true },
    { index: 4, title: 'Location', dataProp: 'OfficeLocation', useDropdown: true },
    { index: 5, title: 'Phones', dataProp: 'BusinessPhones', useDropdown: false },
    { index: 6, title: 'Licenses', dataProp: 'Licenses', useDropdown: true }
  ];

  let uniqueFieldValues = {};

  function showLoader() {
    if ($('#loadingOverlay').length === 0 && $('body').length > 0) {
        $('body').append('<div id="loadingOverlay" style="display:none; position:fixed; top:0; left:0; width:100%; height:100%; background:rgba(0,0,0,0.5); z-index:10000; display:flex; align-items:center; justify-content:center;"><div style="background:white; color:black; padding:20px; border-radius:5px;">Processing...</div></div>');
    }
    $('#loadingOverlay').show();
  }

  function hideLoader() {
    $('#loadingOverlay').hide();
  }

  const nameKey = u => `${u.DisplayName || ''}|||${u.OfficeLocation || ''}`;
  const escapeHtml = s => typeof s === 'string' ? s.replace(/[&<>"']/g, m => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' })[m]) : '';

  function escapeCsvValue(value, forceQuotesOnSemicolon = false) {
    if (value == null) return '';
    let stringValue = String(value);
    const regex = forceQuotesOnSemicolon ? /[,"\n\r;]/ : /[,"\n\r]/;
    if (regex.test(stringValue)) {
      stringValue = `"${stringValue.replace(/"/g, '""')}"`;
    }
    return stringValue;
  }

  function validateJson(data) {
    if (!Array.isArray(data)) {
      alert('Invalid JSON: Must be an array of objects.');
      return [];
    }
    return data.map((u, i) => u && typeof u === 'object' ? {
      Id: u.Id || `unknown_${i}`,
      DisplayName: u.DisplayName || 'Unknown',
      OfficeLocation: u.OfficeLocation || 'Unknown',
      Email: u.Email || '',
      JobTitle: u.JobTitle || 'Unknown',
      BusinessPhones: Array.isArray(u.BusinessPhones) ? u.BusinessPhones : (typeof u.BusinessPhones === 'string' ? u.BusinessPhones.split('; ').filter(p => p) : []),
      Licenses: Array.isArray(u.Licenses) ? u.Licenses.map(l => ({
        LicenseName: l.LicenseName || `Lic_${i}_${Math.random().toString(36).substr(2, 5)}`,
        SkuId: l.SkuId || ''
      })).filter(l => l.LicenseName) : []
    } : null).filter(x => x);
  }

  function findIssues(data) {
    const nameMap = new Map();
    const dupSet = new Set();
    const officeLic = new Set([
      'Microsoft 365 E3', 'Microsoft 365 E5',
      'Microsoft 365 Business Standard', 'Microsoft 365 Business Premium',
      'Office 365 E3', 'Office 365 E5'
    ]);
    data.forEach(u => {
      const k = nameKey(u);
      nameMap.set(k, (nameMap.get(k) || 0) + 1);
      const licCount = new Map();
      (u.Licenses || []).forEach(l => {
        const baseName = (l.LicenseName || '').match(/^(Microsoft 365|Office 365)/)?.[0] ?
          (l.LicenseName.match(/^(Microsoft 365 E3|Microsoft 365 E5|Microsoft 365 Business Standard|Microsoft 365 Business Premium|Office 365 E3|Office 365 E5)/)?.[0] || l.LicenseName)
          : l.LicenseName;
        if (baseName) {
          licCount.set(baseName, (licCount.get(baseName) || 0) + 1);
        }
      });
      if ([...licCount].some(([lic, c]) => officeLic.has(lic) && c > 1)) {
        dupSet.add(u.Id);
      }
    });
    const conflictingNameKeys = new Set([...nameMap].filter(([, count]) => count > 1).map(([key]) => key));
    return { nameConflicts: conflictingNameKeys, dupLicUsers: dupSet };
  }

  function renderAlerts() {
    const $p = $('#alertPanel').empty();
    if (nameConflicts.size) {
      $p.append(`<div class="alert-badge">
        <button id="filterNameConflicts" style="background: none; border: none; color: inherit; cursor: pointer; text-decoration: underline;" class="underline"><i class="fas fa-users-slash" style="margin-right: 0.3rem;"></i>Name+Location Conflicts: ${nameConflicts.size}</button>
      </div>`);
    }
    if (dupLicUsers.size) {
      const list = allData.filter(u => dupLicUsers.has(u.Id)).map(u => {
        const licCount = {}, duplicateLicNames = [];
        (u.Licenses || []).forEach(l => { licCount[l.LicenseName] = (licCount[l.LicenseName] || 0) + 1; });
        Object.entries(licCount).forEach(([licName, count]) => count > 1 && duplicateLicNames.push(licName));
        const hasPaid = (u.Licenses || []).some(l => !(l.LicenseName || '').toLowerCase().includes('free'));
        return `<li>${escapeHtml(u.DisplayName)} (${escapeHtml(u.OfficeLocation)}): ${escapeHtml(duplicateLicNames.join(', '))} — Paid: ${hasPaid ? 'Yes' : 'No'}</li>`;
      }).join('');
      $p.append(`<div class="alert-badge">
        <span><i class="fas fa-copy" style="margin-right: 0.3rem;"></i>Duplicate Licenses: ${dupLicUsers.size}</span>
        <button class="underline toggle-details" data-target="dupDetails" style="background: none; border: none; color: inherit; cursor: pointer; text-decoration: underline;">Details</button>
      </div>
      <div id="dupDetails" class="alert-details"><ul>${list}</ul></div>`);
    }
    $('#filterNameConflicts').off('click').on('click', function() {
      if (!table) return;
      const conflictUserIds = allData.filter(u => nameConflicts.has(nameKey(u))).map(u => u.Id);
      table.search('').columns().search('');
      table.column(0).search('^(' + conflictUserIds.join('|') + ')$', true, false).draw();
    });
    $('.toggle-details').off('click').on('click', function() {
      const targetId = $(this).data('target');
      $(`#${targetId}`).toggleClass('show');
      $(this).text($(`#${targetId}`).hasClass('show') ? 'Hide' : 'Details');
    });
  }

  function initTable(data) {
    allData = data;
    ({ nameConflicts, dupLicUsers } = findIssues(allData));
    uniqueFieldValues = {};
    searchableColumnsConfig.forEach(colConfig => {
      if (colConfig.useDropdown) {
        if (colConfig.dataProp === 'Licenses') {
          const allLicenseObjects = allData.flatMap(user => user.Licenses || []).filter(l => l.LicenseName);
          uniqueFieldValues.Licenses = [...new Set(allLicenseObjects.map(l => l.LicenseName))].sort();
        } else {
          uniqueFieldValues[colConfig.dataProp] = [...new Set(allData.map(user => user[colConfig.dataProp]).filter(value => value && String(value).trim() !== ''))].sort();
        }
      }
    });
    if ($('#licenseDatalist').length === 0) {
      $('body').append('<datalist id="licenseDatalist"></datalist>');
    }
    const $licenseDatalist = $('#licenseDatalist').empty();
    if (uniqueFieldValues.Licenses) {
      uniqueFieldValues.Licenses.forEach(name => {
        $licenseDatalist.append($('<option>').attr('value', escapeHtml(name)));
      });
    }
    if (table) {
      $(table.table().node()).off('preDraw.dt draw.dt');
      table.destroy();
    }
    table = $('#licenseTable').DataTable({
      data: allData,
      deferRender: true, pageLength: 25, orderCellsTop: true,
      columns: [
        { data: 'Id', title: 'ID', visible: false },
        { data: 'DisplayName', title: 'Name', visible: true },
        { data: 'Email', title: 'Email', visible: true },
        { data: 'JobTitle', title: 'Job Title', visible: true },
        { data: 'OfficeLocation', title: 'Location', visible: false },
        { data: 'BusinessPhones', title: 'Phones', visible: false, render: p => Array.isArray(p) ? p.join('; ') : (p || '') },
        { data: 'Licenses', title: 'Licenses', visible: true, render: l => Array.isArray(l) ? l.map(x => x.LicenseName || '').filter(name => name).join(', ') : '' }
      ],
      initComplete: function() {
        const api = this.api();
        api.columns().every(function(colIdx) {
          const column = this;
          const input = $(api.table().header()).find('tr:eq(1) th:eq(' + colIdx + ') input');
          if (input.length > 0) {
            input.off('keyup change clear').on('keyup change clear', function() {
              if (column.search() !== this.value) { column.search(this.value).draw(); }
            });
          }
        });
        $('#colContainer .col-vis').each(function() {
          const idx = +$(this).data('col');
          try {
            if (idx >= 0 && idx < api.columns().nodes().length) { $(this).prop('checked', api.column(idx).visible()); }
            else { $(this).prop('disabled', true); }
          } catch (e) { console.warn("Error checking column visibility:", idx, e); }
        });
        $('.col-vis').off('change').on('change', function() {
          const idx = +$(this).data('col');
          try {
            const col = api.column(idx);
            if (col && col.visible) { col.visible(!col.visible()); }
            else { console.warn("Column object not found for index:", idx); }
          } catch (e) { console.warn("Error toggling column visibility:", idx, e); }
        });
        $('#multiSearchFields .multi-search-row').each(function() { updateSearchFieldUI($(this)); });
        applyMultiSearch();
      },
      rowCallback: function(row, data) {
        $(row).removeClass('conflict dup-license');
        if (nameConflicts.has(nameKey(data))) $(row).addClass('conflict');
        if (dupLicUsers.has(data.Id)) $(row).addClass('dup-license');
      },
      drawCallback: renderAlerts
    });
    $(table.table().node()).on('preDraw.dt', showLoader).on('draw.dt', hideLoader);
  }

  // MODIFIED function to integrate Select2
  function updateSearchFieldUI($row) {
    const selectedColIndex = $row.find('.column-select').val();
    const columnConfig = searchableColumnsConfig.find(c => c.index == selectedColIndex);

    const $textInput = $row.find('.search-input');
    const $dropdownInput = $row.find('.search-value-select');

    // If Select2 was previously initialized, destroy it to prevent conflicts
    if ($dropdownInput.hasClass("select2-hidden-accessible")) {
      $dropdownInput.select2('destroy');
    }
    // Ensure the original select is visible before Select2 potentially hides it again
    // This also helps reset its state if switching from text input.
    $dropdownInput.show();


    if (columnConfig && columnConfig.useDropdown) {
      $textInput.hide();
      $dropdownInput.show().empty(); // Show and clear previous options

      // Add a default empty option which Select2 can use for its placeholder
      $dropdownInput.append($('<option>').val('').text('')); // Placeholder managed by Select2

      const values = uniqueFieldValues[columnConfig.dataProp] || [];
      values.forEach(val => {
        $dropdownInput.append($('<option>').val(escapeHtml(val)).text(escapeHtml(val)));
      });

      // Initialize Select2
      $dropdownInput.select2({
        placeholder: `Digite ou selecione um(a) ${columnConfig.title.toLowerCase()}`,
        allowClear: true,
        width: 'style', // Adjusts width to the style of the original element
                      // or use '100%' if parent container controls width
        minimumInputLength: 0 // Show dropdown on focus, can be set to 1 to require typing
      });
      // Ensure the Select2 container is visible (it might be hidden if the original select was display:none)
      $dropdownInput.next(".select2-container").show();

    } else {
      $textInput.show();
      $dropdownInput.hide().empty(); // Hide and empty if not used
      // Also explicitly hide any Select2 container that might have been associated
      $dropdownInput.next(".select2-container").hide();
      $textInput.attr('placeholder', 'Termo...');
    }
  }

  function setupMultiSearch() {
    const $container = $('#multiSearchFields');
    function addSearchField() {
      const columnOptions = searchableColumnsConfig
        .map(c => `<option value="${c.index}">${escapeHtml(c.title)}</option>`)
        .join('');
      const $row = $(`<div class="multi-search-row">
        <select class="column-select">${columnOptions}</select>
        <input class="search-input" placeholder="Termo..." />
        <select class="search-value-select" style="display: none;"></select>
        <button class="remove-field" title="Remover filtro"><i class="fas fa-trash-alt"></i></button>
      </div>`);
      $container.append($row);
      updateSearchFieldUI($row);
      $row.find('.column-select').on('change', function() {
        updateSearchFieldUI($row);
        $row.find('.search-input').val('');
        $row.find('.search-value-select').val('').trigger('change'); // Clear Select2 value properly
        applyMultiSearch();
      });
      // For text input, 'input' event is good for immediate feedback.
      // For Select2, 'change' event is standard when a selection is made.
      $row.find('.search-input').on('input change', applyMultiSearch);
      $row.find('.search-value-select').on('change', applyMultiSearch);

      $row.find('.remove-field').on('click', function() {
        // If Select2 is active on the select being removed, destroy it first
        const $selectToRemove = $(this).closest('.multi-search-row').find('.search-value-select');
        if ($selectToRemove.hasClass("select2-hidden-accessible")) {
            $selectToRemove.select2('destroy');
        }
        $(this).closest('.multi-search-row').remove();
        applyMultiSearch();
      });
    }
    $('#addSearchField').off('click').on('click', addSearchField);
    $('#multiSearchOperator').off('change').on('change', applyMultiSearch);
    if ($container.children().length === 0) { addSearchField(); }
    _executeMultiSearchLogic();
  }

  function _executeMultiSearchLogic() {
    const operator = $('#multiSearchOperator').val();
    const $searchCriteriaText = $('#searchCriteria');
    if (!table) {
      $searchCriteriaText.text(allData.length === 0 ? 'No data loaded.' : 'Table not available/initialized.');
      return;
    }
    table.search('').columns().search('');
    while ($.fn.dataTable.ext.search.length > 0) { $.fn.dataTable.ext.search.pop(); }
    const filters = [];
    $('#multiSearchFields .multi-search-row').each(function() {
      const colIndex = $(this).find('.column-select').val();
      const columnConfig = searchableColumnsConfig.find(c => c.index == colIndex);
      let searchTerm = '';
      if (columnConfig) {
        searchTerm = columnConfig.useDropdown ?
          $(this).find('.search-value-select').val() :
          $(this).find('.search-input').val().trim();
        if (searchTerm) {
          filters.push({
            col: parseInt(colIndex, 10), term: searchTerm,
            dataProp: columnConfig.dataProp, isDropdown: columnConfig.useDropdown
          });
        }
      }
    });
    let criteriaText = operator === 'AND' ? 'Critérios: Todos os filtros devem corresponder (E)' : 'Critérios: Qualquer filtro deve corresponder (OU)';
    if (filters.length > 0) {
      criteriaText += ` (${filters.length} filtro(s) ativo(s))`;
      $.fn.dataTable.ext.search.push(
        function(settings, apiData, dataIndex) {
          if (settings.nTable.id !== table.table().node().id) return true;
          const rowData = table.row(dataIndex).data();
          if (!rowData) return false;
          const logicFn = operator === 'OR' ? filters.some.bind(filters) : filters.every.bind(filters);
          return logicFn(filter => {
            if (filter.dataProp === 'Licenses') {
              return (rowData.Licenses && Array.isArray(rowData.Licenses)) ?
                rowData.Licenses.some(l => (l.LicenseName || '').toLowerCase() === filter.term.toLowerCase()) : false;
            } else if (filter.isDropdown) {
              const cellValue = rowData[filter.dataProp] || '';
              return cellValue.toString().toLowerCase() === filter.term.toLowerCase();
            } else {
              const cellValue = apiData[filter.col] || '';
              return cellValue.toString().toLowerCase().includes(filter.term.toLowerCase());
            }
          });
        }
      );
    } else {
      criteriaText = 'Critérios: Todos os resultados (sem filtros ativos)';
    }
    $searchCriteriaText.text(criteriaText);
    table.draw();
  }

  function applyMultiSearch() {
    clearTimeout(multiSearchDebounceTimer);
    multiSearchDebounceTimer = setTimeout(_executeMultiSearchLogic, DEBOUNCE_DELAY);
  }

  $('#clearFilters').on('click', () => {
    if (table) {
      $(table.table().header()).find('tr:eq(1) th input').val('');
      table.search('').columns().search('');
    }
    // Properly clear Select2 instances in multi-search rows before removing them
    $('#multiSearchFields .multi-search-row').each(function() {
        const $selectToClear = $(this).find('.search-value-select');
        if ($selectToClear.hasClass("select2-hidden-accessible")) {
            $selectToClear.select2('destroy');
        }
    });
    $('#multiSearchFields').empty();

    if (allData.length > 0) { setupMultiSearch(); }
    else { $('#searchCriteria').text('Nenhum dado carregado.'); }
    while ($.fn.dataTable.ext.search.length > 0) { $.fn.dataTable.ext.search.pop(); }
    if(table) table.draw();
    $('#alertPanel').empty();
    const defaultVisibleCols = [1, 2, 3, 6];
    $('#colContainer .col-vis').each(function() {
      const idx = +$(this).data('col');
      const isDefaultVisible = defaultVisibleCols.includes(idx);
      if (table && idx >= 0 && idx < table.columns().nodes().length) {
        try { table.column(idx).visible(isDefaultVisible); } catch(e) { console.warn("Error resetting column visibility:", idx, e); }
      }
      $(this).prop('checked', isDefaultVisible);
    });
  });

  function downloadCsv(csvContent, fileName) {
    const bom = "\uFEFF";
    const blob = new Blob([bom, csvContent], { type: 'text/csv;charset=utf-8;' });
    const link = document.createElement('a');
    if (link.download !== undefined) {
      const url = URL.createObjectURL(blob);
      link.setAttribute('href', url);
      link.setAttribute('download', fileName);
      link.style.visibility = 'hidden';
      document.body.appendChild(link);
      link.click();
      document.body.removeChild(link);
      URL.revokeObjectURL(url);
    } else {
      alert("Seu navegador não suporta o download direto de arquivos.");
    }
  }

  $('#exportCsv').on('click', () => {
    if (!table) return alert('Tabela não inicializada. Nenhum dado carregado.');
    const rowsToExport = table.rows({ search: 'applied' }).data().toArray();
    if (!rowsToExport.length) return alert('Nenhum registro para exportar com os filtros atuais.');
    const visibleColumns = [];
    table.columns(':visible').every(function() {
      const columnTitle = $(table.table().header()).find('tr:eq(0) th').eq(this.index()).text();
      const dataProp = table.settings()[0].aoColumns[this.index()].mData;
      visibleColumns.push({ title: columnTitle, dataProp: dataProp });
    });
    const headerRow = visibleColumns.map(col => escapeCsvValue(col.title)).join(',');
    const csvRows = rowsToExport.map(rowData => {
      return visibleColumns.map(col => {
        let cellData = rowData[col.dataProp];
        let shouldForceQuotes = false;
        if (col.dataProp === 'BusinessPhones') {
          cellData = Array.isArray(cellData) ? cellData.join('; ') : (cellData || '');
          if (String(cellData).match(/[,"\n\r;]/)) shouldForceQuotes = true;
        } else if (col.dataProp === 'Licenses') {
          const licensesArray = (rowData.Licenses && Array.isArray(rowData.Licenses)) ?
            rowData.Licenses.map(l => l.LicenseName || '').filter(name => name) : [];
          cellData = licensesArray.length > 0 ? licensesArray.join('; ') : '';
          if (String(cellData).match(/[,"\n\r;]/)) shouldForceQuotes = true;
        } else {
          if (String(cellData).match(/[,"\n\r]/)) shouldForceQuotes = true;
        }
        return escapeCsvValue(cellData, shouldForceQuotes);
      }).join(',');
    });
    const csvContent = [headerRow, ...csvRows].join('\n');
    downloadCsv(csvContent, 'license_report.csv');
  });

  $('#exportIssues').on('click', () => {
    if (!allData.length) return alert('Nenhum dado carregado para gerar o relatório de problemas.');
    const lines = [];
    if (nameConflicts.size) {
      lines.push(['CONFLITOS DE NOME+LOCALIZAÇÃO']);
      lines.push(['Nome', 'Localização'].map(h => escapeCsvValue(h)));
      nameConflicts.forEach(key => lines.push(key.split('|||').map(value => escapeCsvValue(value))));
      lines.push([]);
    }
    if (dupLicUsers.size) {
      lines.push(['USUÁRIOS com Licenças Duplicadas']);
      lines.push(['Nome', 'Localização', 'Licenças Duplicadas', 'Possui Licença Paga?'].map(h => escapeCsvValue(h)));
      allData.filter(user => dupLicUsers.has(user.Id)).forEach(user => {
        const licCount = {}, duplicateLicNames = [];
        (user.Licenses || []).forEach(l => licCount[l.LicenseName] = (licCount[l.LicenseName] || 0) + 1);
        Object.entries(licCount).forEach(([licName, count]) => count > 1 && duplicateLicNames.push(licName));
        const joinedDups = duplicateLicNames.join('; ');
        const hasPaid = (user.Licenses || []).some(l => !(l.LicenseName || '').toLowerCase().includes('free'));
        lines.push([
          escapeCsvValue(user.DisplayName), escapeCsvValue(user.OfficeLocation),
          escapeCsvValue(joinedDups, joinedDups.includes(';') || joinedDups.match(/[,"\n\r]/)),
          escapeCsvValue(hasPaid ? 'Yes' : 'No')
        ]);
      });
    }
    if (!lines.length) { lines.push(['Nenhum problema detectado.']); }
    const csvContent = lines.map(rowArray => rowArray.join(',')).join('\n');
    downloadCsv(csvContent, 'issues_report.csv');
  });

  try {
    if (typeof userData === 'undefined') {
      throw new Error("A variável 'userData' não está definida. Os dados JSON não foram corretamente incorporados no HTML.");
    }
    const validatedData = validateJson(userData);
    if (validatedData.length > 0) {
      initTable(validatedData);
      setupMultiSearch();
      $('#searchCriteria').text(`Dados carregados. (${validatedData.length} usuários) Use filtros para refinar.`);
    } else {
      alert('Nenhum usuário válido encontrado nos dados.');
      $('#searchCriteria').text('Nenhum dado válido carregado.');
    }
  } catch (error) {
    alert('Erro ao processar dados: ' + error.message);
    console.error("Erro ao processar dados:", error);
    $('#searchCriteria').text('Erro ao carregar dados.');
  }
});
