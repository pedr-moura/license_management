$(document).ready(function() {
  let table = null, allData = [], nameConflicts = new Set(), dupLicUsers = new Set();
  let uniqueLicenseNames = [];

  const DEBOUNCE_DELAY = 350;
  let multiSearchDebounceTimer;

  // Funções para controlar o overlay de carregamento
  function showLoader() {
    // Certifique-se de que o elemento #loadingOverlay existe no seu HTML
    if ($('#loadingOverlay').length === 0 && $('body').length > 0) { // Adiciona apenas se não existir e o body estiver pronto
        $('body').append('<div id="loadingOverlay" style="display:none; position:fixed; top:0; left:0; width:100%; height:100%; background:rgba(0,0,0,0.5); z-index:10000; display:flex; align-items:center; justify-content:center;"><div style="background:white; color:black; padding:20px; border-radius:5px;">Processando...</div></div>');
    }
    $('#loadingOverlay').show();
  }

  function hideLoader() {
    $('#loadingOverlay').hide();
  }

  const nameKey = u => `${u.DisplayName || ''}|||${u.OfficeLocation || ''}`;
  const escapeHtml = s => typeof s === 'string' ? s.replace(/[&<>"']/g, m => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' })[m]) : '';

  function validateJson(data) {
    if (!Array.isArray(data)) {
      alert('JSON inválido: Deve ser um array de objetos.');
      return [];
    }
    return data.map((u, i) => u && typeof u === 'object' ? {
      Id: u.Id || `unknown_${i}`,
      DisplayName: u.DisplayName || 'Desconhecido',
      OfficeLocation: u.OfficeLocation || 'Desconhecido',
      Email: u.Email || '',
      JobTitle: u.JobTitle || '',
      BusinessPhones: Array.isArray(u.BusinessPhones) ? u.BusinessPhones : (typeof u.BusinessPhones === 'string' ? u.BusinessPhones.split('; ') : []),
      Licenses: Array.isArray(u.Licenses) ? u.Licenses.map(l => ({
        LicenseName: l.LicenseName || `Lic_${i}_${Math.random().toString(36).substr(2, 5)}`,
        SkuId: l.SkuId || ''
      })) : []
    } : null).filter(x => x);
  }

  function findIssues(data) {
    const nameMap = new Map(), dupSet = new Set(), officeLic = new Set([
      'Microsoft 365 E3', 'Microsoft 365 E5',
      'Microsoft 365 Business Standard', 'Microsoft 365 Business Premium',
      'Office 365 E3', 'Office 365 E5'
    ]);
    data.forEach(u => {
      const k = nameKey(u);
      nameMap.set(k, (nameMap.get(k) || 0) + 1);
      const cnt = new Map();
      (u.Licenses || []).forEach(l => {
        const base = (l.LicenseName || '').match(/^(Microsoft 365|Office 365)/)?.[0] || l.LicenseName;
        cnt.set(base, (cnt.get(base) || 0) + 1);
      });
      if ([...cnt].some(([lic, c]) => officeLic.has(lic) && c > 1)) dupSet.add(u.Id);
    });
    return { nameConflicts: new Set([...nameMap].filter(([, c]) => c > 1).map(([k]) => k)), dupLicUsers: dupSet };
  }

  function renderAlerts() {
    const $p = $('#alertPanel').empty();
    if (nameConflicts.size) {
      $p.append(`<div class="alert-badge">
        <button id="filterNameConflicts" style="background: none; border: none; color: inherit; cursor: pointer; text-decoration: underline;" class="underline"><i class="fas fa-users-slash" style="margin-right: 0.3rem;"></i>Conflitos Nome+Unidade: ${nameConflicts.size}</button>
      </div>`);
    }
    if (dupLicUsers.size) {
      const list = allData.filter(u => dupLicUsers.has(u.Id)).map(u => {
        const cnt = {}, dups = [];
        (u.Licenses || []).forEach(l => cnt[l.LicenseName] = (cnt[l.LicenseName] || 0) + 1);
        Object.entries(cnt).forEach(([lic, c]) => c > 1 && dups.push(lic));
        const hasPaid = (u.Licenses || []).some(l => !(l.LicenseName || '').toLowerCase().includes('free'));
        return `<li>${escapeHtml(u.DisplayName)} (${escapeHtml(u.OfficeLocation)}): ${escapeHtml(dups.join(', '))} — Paga: ${hasPaid ? 'Sim' : 'Não'}</li>`;
      }).join('');
      $p.append(`<div class="alert-badge">
        <span><i class="fas fa-copy" style="margin-right: 0.3rem;"></i>Licenças duplicadas: ${dupLicUsers.size}</span>
        <button class="underline toggle-details" data-target="dupDetails" style="background: none; border: none; color: inherit; cursor: pointer; text-decoration: underline;">Detalhes</button>
      </div>
      <div id="dupDetails" class="alert-details"><ul>${list}</ul></div>`);
    }

    $('#filterNameConflicts').off('click').on('click', function() {
      if (!table) return;
      const ids = allData.filter(u => nameConflicts.has(nameKey(u))).map(u => u.Id);
      table.search('').columns().search(''); // Limpa busca global e por coluna antes de aplicar nova
      table.column(0).search(ids.join('|'), true, false).draw();
    });

    $('.toggle-details').off('click').on('click', function() {
      const tgt = $(this).data('target');
      $(`#${tgt}`).toggleClass('show');
      $(this).text($(`#${tgt}`).hasClass('show') ? 'Ocultar' : 'Detalhes');
    });
  }

  function initTable(data) {
    allData = data;
    ({ nameConflicts, dupLicUsers } = findIssues(allData));

    const allLicenseObjects = allData.flatMap(user => user.Licenses || []).filter(l => l.LicenseName);
    uniqueLicenseNames = [...new Set(allLicenseObjects.map(l => l.LicenseName))].sort();

    if ($('#licenseDatalist').length === 0) {
      $('body').append('<datalist id="licenseDatalist"></datalist>');
    }
    const $licenseDatalist = $('#licenseDatalist').empty();
    uniqueLicenseNames.forEach(name => {
      $licenseDatalist.append($('<option>').attr('value', name));
    });

    if (table) {
      $(table.table().node()).off('preDraw.dt draw.dt'); // Desvincular eventos antigos
      table.destroy();
    }

    table = $('#licenseTable').DataTable({
      data: allData,
      deferRender: true,
      pageLength: 25,
      orderCellsTop: true,
      language: {
        "emptyTable": "Nenhum registro encontrado",
        "info": "Mostrando de _START_ até _END_ de _TOTAL_ registros",
        "infoEmpty": "Mostrando 0 até 0 de 0 registros",
        "infoFiltered": "(Filtrados de _MAX_ registros)",
        "infoThousands": ".",
        "lengthMenu": "_MENU_ resultados por página",
        "loadingRecords": "Carregando...",
        "processing": "Processando...", // Este é um texto padrão do DataTables, o nosso loader é customizado
        "zeroRecords": "Nenhum registro encontrado",
        "search": "Pesquisar:",
        "paginate": {
          "next": "Próximo",
          "previous": "Anterior",
          "first": "Primeiro",
          "last": "Último"
        },
        "aria": {
          "sortAscending": ": Ordenar colunas de forma ascendente",
          "sortDescending": ": Ordenar colunas de forma descendente"
        },
        // ... (resto do seu objeto de 'language', se houver mais)
      },
      columnDefs: [
        { targets: '_all', visible: false },
        { targets: [1, 2, 3, 6], visible: true } // Nome, Email, Cargo, Licenças
      ],
      columns: [
        { data: 'Id', title: 'ID' },
        { data: 'DisplayName', title: 'Nome' },
        { data: 'Email', title: 'Email' },
        { data: 'JobTitle', title: 'Cargo' },
        { data: 'OfficeLocation', title: 'Unidade' },
        { data: 'BusinessPhones', title: 'Telefones', render: p => Array.isArray(p) ? p.join('; ') : (p || '') },
        { data: 'Licenses', title: 'Licenças', render: l => Array.isArray(l) ? l.map(x => x.LicenseName || '').join(', ') : '' }
      ],
      initComplete: function() {
        const api = this.api();
        api.columns().every(function(colIdx) {
          const column = this;
          // Supondo que seus inputs de filtro por coluna estejam na segunda linha do header (tr:eq(1))
          const input = $(api.table().header()).find('tr:eq(1) th:eq(' + colIdx + ') input');
          if (input.length > 0) {
            input.off('keyup change clear').on('keyup change clear', function() {
              if (column.search() !== this.value) {
                column.search(this.value).draw();
              }
            });
          }
        });

        $('#colContainer .col-vis').each(function() {
          const idx = +$(this).data('col');
          try {
            $(this).prop('checked', api.column(idx).visible());
          } catch (e) {
            console.warn("Erro ao verificar visibilidade da coluna:", idx, e);
          }
        });
        $('.col-vis').off('change').on('change', function() {
          const idx = +$(this).data('col');
          try {
            const col = api.column(idx);
            col.visible(!col.visible()); // Isso também dispara um draw
          } catch (e) {
            console.warn("Erro ao alternar visibilidade da coluna:", idx, e);
          }
        });
        // Garante que o estado dos inputs de multi-pesquisa seja atualizado se já existirem
        $('#multiSearchFields .multi-search-row').each(function() {
          updateLicenseInputState($(this));
        });
      },
      rowCallback: function(row, data) {
        $(row).removeClass('conflict dup-license');
        if (nameConflicts.has(nameKey(data))) $(row).addClass('conflict');
        if (dupLicUsers.has(data.Id)) $(row).addClass('dup-license');
      },
      drawCallback: renderAlerts // Sua função renderAlerts continua aqui
    });

    // Vincular eventos preDraw e draw para o loader
    $(table.table().node()).on('preDraw.dt', function() {
      showLoader();
    });

    $(table.table().node()).on('draw.dt', function() {
      hideLoader();
    });
  }

  function updateLicenseInputState($row) {
    const $select = $row.find('.column-select');
    const $input = $row.find('.search-input');
    const $licenseSelect = $row.find('.license-select');
    if ($select.val() == 6) { // Coluna de Licenças (índice 6 na definição de 'columns')
      $input.hide();
      $licenseSelect.show();
    } else {
      $input.show();
      $licenseSelect.hide();
      $input.attr('placeholder', 'Termo...');
    }
  }

  function setupMultiSearch() {
    const cols = [
      { v: 0, text: 'ID' }, { v: 1, text: 'Nome' }, { v: 2, text: 'Email' },
      { v: 3, text: 'Cargo' }, { v: 4, text: 'Unidade' }, { v: 5, text: 'Telefones' },
      { v: 6, text: 'Licenças' }
    ];
    const $ct = $('#multiSearchFields');

    function addSearchField() {
      const $r = $(`<div class="multi-search-row">
        <select class="column-select">
          ${cols.map(c => `<option value="${c.v}">${c.text}</option>`).join('')}
        </select>
        <input class="search-input" placeholder="Termo..." />
        <select class="license-select" style="display: none;">
          <option value="">Selecione uma licença...</option>
          ${uniqueLicenseNames.map(name => `<option value="${escapeHtml(name)}">${escapeHtml(name)}</option>`).join('')}
        </select>
        <button class="remove-field" title="Remover filtro"><i class="fas fa-trash-alt"></i></button>
      </div>`);
      $ct.append($r);
      updateLicenseInputState($r);
      $r.find('.column-select').on('change', function() {
        updateLicenseInputState($r);
        $r.find('.search-input').val('');
        $r.find('.license-select').val('');
        applyMultiSearch();
      });
      $r.find('.search-input, .license-select').on('input change', applyMultiSearch);
      $r.find('.remove-field').on('click', function() {
        $(this).closest('.multi-search-row').remove();
        applyMultiSearch();
      });
    }

    $('#addSearchField').off('click').on('click', addSearchField);
    $('#multiSearchOperator').off('change').on('change', applyMultiSearch);

    if ($ct.children().length === 0) {
      addSearchField();
    }
    // Chamada inicial para aplicar qualquer filtro padrão ou estado
    // applyMultiSearch(); // ou _executeMultiSearchLogic(); se não precisar de debounce na carga inicial
    // Considerando que _executeMultiSearchLogic limpa filtros e redesenha, pode ser melhor chamá-lo diretamente se não houver inputs preenchidos
     _executeMultiSearchLogic(); // Para garantir que a lógica de filtro seja aplicada ao carregar
  }

  function _executeMultiSearchLogic() {
    const op = $('#multiSearchOperator').val();
    const $searchCriteria = $('#searchCriteria');

    if (!table && allData.length === 0) {
      $searchCriteria.text('Nenhum dado carregado.');
      return;
    }
    if (!table && allData.length > 0) {
      $searchCriteria.text('Tabela não inicializada, mas dados carregados. Filtros serão aplicados ao carregar a tabela.');
      return;
    }
    if (!table) {
      $searchCriteria.text('Tabela não disponível.');
      return;
    }

    table.search('').columns().search('');
    
    while ($.fn.dataTable.ext.search.length > 0) {
      $.fn.dataTable.ext.search.pop();
    }

    const rows = $('#multiSearchFields .multi-search-row');
    const filters = [];
    rows.each(function() {
      const idx = $(this).find('.column-select').val();
      let val;
      if (idx == 6) { // Coluna de Licenças
        val = $(this).find('.license-select').val();
      } else {
        val = $(this).find('.search-input').val().trim();
      }
      if (val) {
        filters.push({ col: parseInt(idx, 10), term: val });
      }
    });
    
    let criteriaText = op === 'AND' ? 'Critério: Todos os filtros devem corresponder' : 'Critério: Qualquer filtro deve corresponder';

    if (filters.length > 0) {
      criteriaText += ` (${filters.length} filtro(s) ativo(s))`;
      $.fn.dataTable.ext.search.push(
        function(settings, data, dataIndex) {
          if (settings.nTable.id !== table.table().node().id) return true;
          
          const rowData = table.row(dataIndex).data();
          if (!rowData) return false;

          if (op === 'OR') {
            return filters.some(filter => {
              if (filter.col === 6) { // Filtro de Licenças
                return (rowData.Licenses && Array.isArray(rowData.Licenses)) ?
                  rowData.Licenses.some(l => (l.LicenseName || '').toLowerCase() === filter.term.toLowerCase()) : false;
              }
              // Para outras colunas, usa os dados da célula como string (conforme DataTables os passa)
              const cellValue = data[filter.col] || ''; // 'data' aqui são os dados da linha como array de strings renderizadas
              return cellValue.toString().toLowerCase().includes(filter.term.toLowerCase());
            });
          } else { // AND
            return filters.every(filter => {
              if (filter.col === 6) { // Filtro de Licenças
                return (rowData.Licenses && Array.isArray(rowData.Licenses)) ?
                  rowData.Licenses.some(l => (l.LicenseName || '').toLowerCase() === filter.term.toLowerCase()) : false;
              }
              const cellValue = data[filter.col] || '';
              return cellValue.toString().toLowerCase().includes(filter.term.toLowerCase());
            });
          }
        }
      );
    } else {
        criteriaText = op === 'AND' ? 'Critério: Todos os resultados (nenhum filtro AND ativo)' : 'Critério: Todos os resultados (nenhum filtro OR ativo)';
    }
    
    $searchCriteria.text(criteriaText);
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
    $('#multiSearchFields').empty(); 
    if (allData.length > 0) {
        setupMultiSearch(); // Reconfigura e adiciona o primeiro campo, chamando _executeMultiSearchLogic internamente
    } else {
        $('#searchCriteria').text('Nenhum dado carregado.');
    }
    
    // Garante que filtros customizados sejam limpos
    while ($.fn.dataTable.ext.search.length > 0) {
      $.fn.dataTable.ext.search.pop();
    }

    if(table) table.draw(); // Redesenha para aplicar a limpeza
    
    $('#alertPanel').empty();
    // Reseta visibilidade das colunas para o padrão
    $('#colContainer .col-vis').each(function() {
      const idx = +$(this).data('col');
      const isDefaultVisible = [1, 2, 3, 6].includes(idx);
      if (table) {
        try { table.column(idx).visible(isDefaultVisible); } catch(e) { console.warn("Erro ao resetar visibilidade da coluna:", idx, e); }
      }
      $(this).prop('checked', isDefaultVisible);
    });
  });

  function escapeCsvValue(value) {
    if (value == null) return '';
    let stringValue = String(value);
    if (/[,"\n\r]/.test(stringValue)) {
      stringValue = `"${stringValue.replace(/"/g, '""')}"`;
    }
    return stringValue;
  }

  function downloadCsv(csvContent, fileName) {
    const bom = "\uFEFF"; // BOM para UTF-8 no Excel
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
    }
  }

  $('#exportCsv').on('click', () => {
    if (!table) return alert('Tabela não iniciada. Dados não carregados.');
    
    const rowsToExport = table.rows({ search: 'applied' }).data().toArray();
    if (!rowsToExport.length) return alert('Nenhum registro para exportar com os filtros atuais.');

    const visibleColumns = [];
    table.columns(':visible').every(function() {
      // Pega o título da primeira linha do cabeçalho para a coluna atual
      const columnTitle = $(table.table().header()).find('tr:eq(0) th:eq(' + this.index() + ')').text();
      visibleColumns.push({
        title: columnTitle,
        dataProp: table.settings()[0].aoColumns[this.index()].mData // Propriedade de dados da coluna
      });
    });

    const headerRow = visibleColumns.map(col => escapeCsvValue(col.title)).join(',');
    const csvRows = rowsToExport.map(rowData => {
      return visibleColumns.map(col => {
        let cellData = rowData[col.dataProp];
        // Tratamento especial para colunas que são arrays ou objetos
        if (col.dataProp === 'BusinessPhones') {
          cellData = Array.isArray(cellData) ? cellData.join('; ') : (cellData || '');
        } else if (col.dataProp === 'Licenses') {
          cellData = (Array.isArray(rowData.Licenses) && rowData.Licenses.length > 0) ? 
            rowData.Licenses.map(l => l.LicenseName || '').join('; ') : '';
        }
        return escapeCsvValue(cellData);
      }).join(',');
    });

    const csvContent = [headerRow, ...csvRows].join('\n');
    downloadCsv(csvContent, 'relatorio_licencas.csv');
  });

  $('#exportIssues').on('click', () => {
    if (!allData.length) return alert('Nenhum dado carregado para gerar relatório de falhas.');
    let lines = [];
    if (nameConflicts.size) {
      lines.push(['CONFLITOS Nome+Unidade']);
      lines.push(['Nome', 'Unidade']);
      nameConflicts.forEach(k => lines.push(k.split('|||').map(escapeCsvValue)));
      lines.push([]); // Linha em branco para separar seções
    }
    if (dupLicUsers.size) {
      lines.push(['USUÁRIOS com Licenças Duplicadas']);
      lines.push(['Nome', 'Unidade', 'Licenças Duplicadas', 'Possui Licença Paga?']);
      allData.filter(u => dupLicUsers.has(u.Id)).forEach(u => {
        const cnt = {}, dups = [];
        (u.Licenses || []).forEach(l => cnt[l.LicenseName] = (cnt[l.LicenseName] || 0) + 1);
        Object.entries(cnt).forEach(([l, c]) => c > 1 && dups.push(l));
        const hasPaid = (u.Licenses || []).some(l => !(l.LicenseName || '').toLowerCase().includes('free'));
        lines.push([u.DisplayName, u.OfficeLocation, dups.join('; '), hasPaid ? 'Sim' : 'Não'].map(escapeCsvValue));
      });
    }
    if (!lines.length) return alert('Nenhuma falha detectada.');
    downloadCsv(lines.map(r => r.join(',')).join('\n'), 'relatorio_falhas.csv');
  });

  // Inicialização com dados embutidos
  try {
    // A variável 'userData' deve ser definida globalmente (ex: por um script PowerShell)
    // antes deste script ser carregado.
    if (typeof userData === 'undefined') {
        throw new Error("A variável 'userData' não foi definida. O JSON não foi embutido corretamente no HTML.");
    }
    const validatedData = validateJson(userData);  
    if (validatedData.length > 0) {
      initTable(validatedData);
      setupMultiSearch();
      $('#searchCriteria').text(`Dados carregados. (${validatedData.length} usuários) Use os filtros para refinar.`);
    } else {
      alert('Nenhum usuário válido encontrado nos dados.');
      $('#searchCriteria').text('Nenhum dado válido carregado.');
    }
  } catch (er) {
    alert('Erro ao processar dados: ' + er.message);
    console.error("Erro ao processar dados:", er);
    $('#searchCriteria').text('Erro ao carregar dados.');
  }
});
