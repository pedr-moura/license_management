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
        $('body').append('<div id="loadingOverlay" style="display:none; position:fixed; top:0; left:0; width:100%; height:100%; background:rgba(0,0,0,0.5); z-index:10000; display:flex; align-items:center; justify-content:center;"><div style="background:white; color:black; padding:20px; border-radius:5px;">Processando...</div></div>');
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
      alert('JSON Inválido: Deve ser um array de objetos.');
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
        if (baseName) { licCount.set(baseName, (licCount.get(baseName) || 0) + 1); }
      });
      if ([...licCount].some(([lic, c]) => officeLic.has(lic) && c > 1)) { dupSet.add(u.Id); }
    });
    const conflictingNameKeys = new Set([...nameMap].filter(([, count]) => count > 1).map(([key]) => key));
    return { nameConflicts: conflictingNameKeys, dupLicUsers: dupSet };
  }

  function renderAlerts() {
    const $p = $('#alertPanel').empty();
    if (nameConflicts.size) {
      $p.append(`<div class="alert-badge"><button id="filterNameConflicts" style="background: none; border: none; color: inherit; cursor: pointer; text-decoration: underline;" class="underline"><i class="fas fa-users-slash" style="margin-right: 0.3rem;"></i>Conflitos Nome+Local: ${nameConflicts.size}</button></div>`);
    }
    if (dupLicUsers.size) {
      const list = allData.filter(u => dupLicUsers.has(u.Id)).map(u => {
        const licCount = {}, duplicateLicNames = [];
        (u.Licenses || []).forEach(l => { licCount[l.LicenseName] = (licCount[l.LicenseName] || 0) + 1; });
        Object.entries(licCount).forEach(([licName, count]) => count > 1 && duplicateLicNames.push(licName));
        const hasPaid = (u.Licenses || []).some(l => !(l.LicenseName || '').toLowerCase().includes('free'));
        return `<li>${escapeHtml(u.DisplayName)} (${escapeHtml(u.OfficeLocation)}): ${escapeHtml(duplicateLicNames.join(', '))} — Pago: ${hasPaid ? 'Sim' : 'Não'}</li>`;
      }).join('');
      $p.append(`<div class="alert-badge"><span><i class="fas fa-copy" style="margin-right: 0.3rem;"></i>Licenças Duplicadas: ${dupLicUsers.size}</span> <button class="underline toggle-details" data-target="dupDetails" style="background: none; border: none; color: inherit; cursor: pointer; text-decoration: underline;">Detalhes</button></div><div id="dupDetails" class="alert-details"><ul>${list}</ul></div>`);
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
      $(this).text($(`#${targetId}`).hasClass('show') ? 'Esconder' : 'Detalhes');
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
          uniqueFieldValues.Licenses = [...new Set(allLicenseObjects.map(l => l.LicenseName))].sort((a, b) => a.toLowerCase().localeCompare(b.toLowerCase()));
        } else {
          uniqueFieldValues[colConfig.dataProp] = [...new Set(allData.map(user => user[colConfig.dataProp]).filter(value => value && String(value).trim() !== ''))].sort((a, b) => a.toLowerCase().localeCompare(b.toLowerCase()));
        }
      }
    });
    if ($('#licenseDatalist').length === 0) { $('body').append('<datalist id="licenseDatalist"></datalist>'); }
    const $licenseDatalist = $('#licenseDatalist').empty();
    if (uniqueFieldValues.Licenses) {
      uniqueFieldValues.Licenses.forEach(name => { $licenseDatalist.append($('<option>').attr('value', escapeHtml(name))); });
    }
    if (table) { $(table.table().node()).off('preDraw.dt draw.dt'); table.destroy(); }
    table = $('#licenseTable').DataTable({
      data: allData, deferRender: true, pageLength: 25, orderCellsTop: true,
      columns: [
        { data: 'Id', title: 'ID', visible: false }, { data: 'DisplayName', title: 'Nome', visible: true },
        { data: 'Email', title: 'Email', visible: true }, { data: 'JobTitle', title: 'Cargo', visible: true },
        { data: 'OfficeLocation', title: 'Local', visible: false },
        { data: 'BusinessPhones', title: 'Telefones', visible: false, render: p => Array.isArray(p) ? p.join('; ') : (p || '') },
        { data: 'Licenses', title: 'Licenças', visible: true, render: l => Array.isArray(l) ? l.map(x => x.LicenseName || '').filter(name => name).join(', ') : '' }
      ],
      initComplete: function() {
        const api = this.api();
        api.columns().every(function(colIdx) {
          const column = this; $(api.table().header()).find('tr:eq(1) th:eq(' + colIdx + ') input')
            .off('keyup change clear').on('keyup change clear', function() { if (column.search() !== this.value) { column.search(this.value).draw(); } });
        });
        $('#colContainer .col-vis').each(function() {
          const idx = +$(this).data('col'); try { if (idx >= 0 && idx < api.columns().nodes().length) { $(this).prop('checked', api.column(idx).visible()); } else { $(this).prop('disabled', true); } } catch (e) { console.warn("Erro ao checar visibilidade da coluna:", idx, e); }
        });
        $('.col-vis').off('change').on('change', function() {
          const idx = +$(this).data('col'); try { const col = api.column(idx); if (col && col.visible) { col.visible(!col.visible()); } else { console.warn("Coluna não encontrada para índice:", idx); } } catch (e) { console.warn("Erro ao alternar visibilidade da coluna:", idx, e); }
        });
        $('#multiSearchFields .multi-search-row').each(function() { updateSearchFieldUI($(this)); });
        applyMultiSearch();
      },
      rowCallback: function(row, data) { $(row).removeClass('conflict dup-license'); if (nameConflicts.has(nameKey(data))) $(row).addClass('conflict'); if (dupLicUsers.has(data.Id)) $(row).addClass('dup-license'); },
      drawCallback: renderAlerts
    });
    $(table.table().node()).on('preDraw.dt', showLoader).on('draw.dt', hideLoader);
  }

  // --- NOVO: Função para UI de dropdown customizado ---
  function updateSearchFieldUI($row) {
    const selectedColIndex = $row.find('.column-select').val();
    const columnConfig = searchableColumnsConfig.find(c => c.index == selectedColIndex);

    const $searchInput = $row.find('.search-input'); // Input de texto padrão
    const $customDropdownContainer = $row.find('.custom-dropdown-container');
    const $customDropdownTextInput = $row.find('.custom-dropdown-text-input'); // Input para dropdown customizado
    const $customOptionsList = $row.find('.custom-options-list'); // Lista de opções do dropdown customizado
    const $hiddenValueSelect = $row.find('.search-value-select'); // Select oculto que guarda o valor

    // Limpa ouvintes de evento anteriores para evitar duplicatas
    $customDropdownTextInput.off();
    $customOptionsList.off(); // Para eventos delegados

    if (columnConfig && columnConfig.useDropdown) {
      $searchInput.hide(); // Esconde input padrão
      $customDropdownContainer.show(); // Mostra container do dropdown customizado
      $customDropdownTextInput.val(''); // Limpa texto do input customizado
      $customDropdownTextInput.attr('placeholder', `Digite ou selecione ${columnConfig.title.toLowerCase()}`);
      $customOptionsList.hide().empty(); // Esconde e limpa lista de opções
      $hiddenValueSelect.show().empty(); // Garante que o select oculto exista e esteja limpo (mas ele não será visível)

      // Popula o select oculto com todas as opções únicas (fonte da verdade)
      $hiddenValueSelect.append($('<option>').val('').text('')); // Opção vazia padrão
      const allUniqueOptions = uniqueFieldValues[columnConfig.dataProp] || [];
      allUniqueOptions.forEach(val => {
        $hiddenValueSelect.append($('<option>').val(escapeHtml(val)).text(escapeHtml(val)));
      });
      $hiddenValueSelect.val(''); // Garante que valor inicial seja limpo

      // Evento: Digitando no input do dropdown customizado
      $customDropdownTextInput.on('input', function() {
        const searchTerm = $(this).val().toLowerCase();
        $customOptionsList.empty().show();
        const filteredOptions = allUniqueOptions.filter(opt => String(opt).toLowerCase().includes(searchTerm));

        if (filteredOptions.length === 0) {
          $customOptionsList.append('<div class="custom-option-item no-results">Nenhum resultado encontrado</div>');
        } else {
          filteredOptions.forEach(opt => {
            const $optionEl = $('<div class="custom-option-item"></div>').text(opt).data('value', opt);
            $customOptionsList.append($optionEl);
          });
        }
      });

      // Evento: Foco no input do dropdown customizado
      $customDropdownTextInput.on('focus', function() {
        $(this).trigger('input'); // Dispara o 'input' para popular a lista
        $customOptionsList.show(); // Garante que a lista seja exibida
      });

      // Evento: Clique em um item da lista de opções customizada (delegação)
      $customOptionsList.on('mousedown', '.custom-option-item', function(e) { // Usar mousedown para registrar antes do blur
        e.preventDefault(); // Previne que o blur do input esconda a lista antes do click
        if ($(this).hasClass('no-results')) return;

        const selectedText = $(this).text();
        const selectedValue = $(this).data('value');

        $customDropdownTextInput.val(selectedText);
        $hiddenValueSelect.val(selectedValue).trigger('change'); // ATUALIZA O SELECT OCULTO E DISPARA A BUSCA
        $customOptionsList.hide();
      });

      // Evento: Esconder a lista ao perder o foco (blur)
      let blurTimeout;
      $customDropdownTextInput.on('blur', function() {
        clearTimeout(blurTimeout);
        blurTimeout = setTimeout(() => { // Atraso para permitir clique no item da lista
          $customOptionsList.hide();
        }, 150); // Ajuste o tempo conforme necessário
      });

    } else { // Configuração para input de texto padrão
      $searchInput.show();
      $customDropdownContainer.hide();
      $hiddenValueSelect.empty().hide(); // Esconde e limpa select oculto (não usado para texto)
      $searchInput.attr('placeholder', 'Termo...');
    }
  }


  function setupMultiSearch() {
    const $container = $('#multiSearchFields');
    function addSearchField() {
      const columnOptions = searchableColumnsConfig
        .map(c => `<option value="${c.index}">${escapeHtml(c.title)}</option>`)
        .join('');

      // --- NOVO: Estrutura HTML da linha de busca com placeholders para dropdown customizado ---
      const $row = $(`
        <div class="multi-search-row">
            <select class="column-select">${columnOptions}</select>

            <input class="search-input" placeholder="Termo..." style="display:none;" />

            <div class="custom-dropdown-container" style="position: relative; display:none;">
                <input type="text" class="custom-dropdown-text-input" autocomplete="off" />
                <div class="custom-options-list" style="display:none;"></div>
            </div>
            <select class="search-value-select" style="display: none;"></select>

            <button class="remove-field" title="Remover filtro"><i class="fas fa-trash-alt"></i></button>
        </div>
      `);

      $container.append($row);
      updateSearchFieldUI($row); // Configura o tipo de input correto

      $row.find('.column-select').on('change', function() {
        // Limpa os valores dos inputs visuais e do select oculto antes de reconfigurar
        $row.find('.search-input').val('');
        $row.find('.custom-dropdown-text-input').val('');
        // $hiddenValueSelect é limpo dentro de updateSearchFieldUI se for dropdown

        updateSearchFieldUI($row); // Reconfigura a UI para o novo tipo de coluna

        // Dispara a busca com o valor limpo para o novo tipo de input
        const selectedColIndex = $row.find('.column-select').val();
        const columnConfig = searchableColumnsConfig.find(c => c.index == selectedColIndex);

        if (columnConfig && columnConfig.useDropdown) {
            // O select oculto já foi limpo e re-populado em updateSearchFieldUI
            // Dispara 'change' para que applyMultiSearch use o valor vazio (ou o placeholder)
            $row.find('.search-value-select').val('').trigger('change');
        } else {
            // Para input de texto padrão, dispara 'input' para que applyMultiSearch use o valor vazio
            $row.find('.search-input').val('').trigger('input');
        }
      });

      // Ouvintes para disparar a busca
      $row.find('.search-input').on('input change', applyMultiSearch); // Para input de texto padrão
      $row.find('.search-value-select').on('change', applyMultiSearch); // Para o select oculto (atualizado pelo dropdown customizado)

      $row.find('.remove-field').on('click', function() {
        // Limpa ouvintes específicos do dropdown customizado antes de remover
        const $multiSearchRow = $(this).closest('.multi-search-row');
        $multiSearchRow.find('.custom-dropdown-text-input').off();
        $multiSearchRow.find('.custom-options-list').off();

        $multiSearchRow.remove();
        applyMultiSearch();
      });
    }
    $('#addSearchField').off('click').on('click', addSearchField);
    $('#multiSearchOperator').off('change').on('change', applyMultiSearch);
    if ($container.children().length === 0 && allData.length > 0) { addSearchField(); } // Adiciona um campo se houver dados
    else if ($container.children().length === 0 && allData.length === 0) { $('#searchCriteria').text('Nenhum dado carregado.');}
    _executeMultiSearchLogic();
  }

  function _executeMultiSearchLogic() {
    const operator = $('#multiSearchOperator').val();
    const $searchCriteriaText = $('#searchCriteria');
    if (!table) {
      $searchCriteriaText.text(allData.length === 0 ? 'Nenhum dado carregado.' : 'Tabela não inicializada.');
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
          $(this).find('.search-value-select').val() : // Lê do select oculto
          $(this).find('.search-input').val().trim();
        if (searchTerm) {
          filters.push({
            col: parseInt(colIndex, 10), term: searchTerm,
            dataProp: columnConfig.dataProp, isDropdown: columnConfig.useDropdown
          });
        }
      }
    });
    let criteriaText = operator === 'AND' ? 'Critérios: Todos os filtros (E)' : 'Critérios: Qualquer filtro (OU)';
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
              return String(cellValue).toLowerCase() === filter.term.toLowerCase();
            } else {
              const cellValue = apiData[filter.col] || '';
              return String(cellValue).toLowerCase().includes(filter.term.toLowerCase());
            }
          });
        }
      );
    } else { criteriaText = 'Critérios: Todos os resultados (sem filtros ativos)'; }
    $searchCriteriaText.text(criteriaText);
    table.draw();
  }

  function applyMultiSearch() {
    clearTimeout(multiSearchDebounceTimer);
    multiSearchDebounceTimer = setTimeout(_executeMultiSearchLogic, DEBOUNCE_DELAY);
  }

  $('#clearFilters').on('click', () => {
    if (table) { $(table.table().header()).find('tr:eq(1) th input').val(''); table.search('').columns().search(''); }
    // Limpa ouvintes de dropdown customizado antes de esvaziar
    $('#multiSearchFields .multi-search-row').each(function() {
        $(this).find('.custom-dropdown-text-input').off();
        $(this).find('.custom-options-list').off();
    });
    $('#multiSearchFields').empty();
    if (allData.length > 0) { setupMultiSearch(); } else { $('#searchCriteria').text('Nenhum dado carregado.'); }
    while ($.fn.dataTable.ext.search.length > 0) { $.fn.dataTable.ext.search.pop(); }
    if(table) table.draw();
    $('#alertPanel').empty();
    const defaultVisibleCols = [1, 2, 3, 6];
    $('#colContainer .col-vis').each(function() {
      const idx = +$(this).data('col'); const isDefaultVisible = defaultVisibleCols.includes(idx);
      if (table && idx >= 0 && idx < table.columns().nodes().length) { try { table.column(idx).visible(isDefaultVisible); } catch(e) { console.warn("Erro ao resetar visibilidade da coluna:", idx, e); } }
      $(this).prop('checked', isDefaultVisible);
    });
  });

  function downloadCsv(csvContent, fileName) { /* ... (sem alterações) ... */ }
  $('#exportCsv').on('click', () => { /* ... (sem alterações) ... */ });
  $('#exportIssues').on('click', () => { /* ... (sem alterações) ... */ });

  try {
    if (typeof userData === 'undefined') { throw new Error("Variável 'userData' não definida."); }
    const validatedData = validateJson(userData);
    if (validatedData.length > 0) {
      initTable(validatedData); setupMultiSearch();
      $('#searchCriteria').text(`Dados carregados (${validatedData.length} usuários). Use filtros para refinar.`);
    } else { alert('Nenhum usuário válido nos dados.'); $('#searchCriteria').text('Nenhum dado válido carregado.'); }
  } catch (error) { alert('Erro ao processar dados: ' + error.message); console.error("Erro:", error); $('#searchCriteria').text('Erro ao carregar dados.'); }
});
