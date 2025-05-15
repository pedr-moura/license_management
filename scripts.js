$(document).ready(function() {
  let table = null, allData = [], nameConflicts = new Set(), dupLicUsers = new Set();
  let uniqueLicenseNames = [];

  const DEBOUNCE_DELAY = 350;
  let multiSearchDebounceTimer;

  // Funções para controlar o overlay de carregamento
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
      table.search('').columns().search('');
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
      $licenseDatalist.append($('<option>').attr('value', escapeHtml(name)));
    });

    if (table) {
      $(table.table().node()).off('preDraw.dt draw.dt');
      table.destroy();
    }

    table = $('#licenseTable').DataTable({
      data: allData,
      deferRender: true,
      pageLength: 25,
      orderCellsTop: true,
      // language: { // OBJETO LANGUAGE COMPLETO DO SEU ORIGINAL
      //   "emptyTable": "Nenhum registro encontrado",
      //   "info": "Mostrando de _START_ até _END_ de _TOTAL_ registros",
      //   "infoEmpty": "Mostrando 0 até 0 de 0 registros",
      //   "infoFiltered": "(Filtrados de _MAX_ registros)",
      //   "infoThousands": ".",
      //   "lengthMenu": "_MENU_ resultados por página",
      //   "loadingRecords": "Carregando...",
      //   "processing": "Processando...",
      //   "zeroRecords": "Nenhum registro encontrado",
      //   "search": "Pesquisar:",
      //   "paginate": {
      //     "next": "Próximo",
      //     "previous": "Anterior",
      //     "first": "Primeiro",
      //     "last": "Último"
      //   },
      //   "aria": {
      //     "sortAscending": ": Ordenar colunas de forma ascendente",
      //     "sortDescending": ": Ordenar colunas de forma descendente"
      //   },
      //   "select": {
      //     "rows": {
      //       "_": "Selecionado %d linhas",
      //       "0": "Nenhuma linha selecionada",
      //       "1": "Selecionado 1 linha"
      //     },
      //     "cells": {
      //       "1": "1 célula selecionada",
      //       "_": "%d células selecionadas"
      //     },
      //     "columns": {
      //       "1": "1 coluna selecionada",
      //       "_": "%d colunas selecionadas"
      //     }
      //   },
      //   "buttons": {
      //     "copySuccess": {
      //       "1": "Uma linha copiada para a área de transferência",
      //       "_": "%d linhas copiadas para a área de transferência"
      //     },
      //     "collection": "Coleção <span class=\"ui-button-icon-primary ui-icon ui-icon-triangle-1-s\"></span>",
      //     "colvis": "Visibilidade da Coluna",
      //     "colvisRestore": "Restaurar Visibilidade",
      //     "copy": "Copiar",
      //     "copyKeys": "Pressione ctrl ou u2318 + C para copiar os dados da tabela para a área de transferência do sistema. Para cancelar, clique nesta mensagem ou pressione Esc.",
      //     "copyTitle": "Copiar para área de transferência",
      //     "csv": "CSV",
      //     "excel": "Excel",
      //     "pageLength": {
      //       "-1": "Mostrar todos os registros",
      //       "_": "Mostrar %d registros"
      //     },
      //     "pdf": "PDF",
      //     "print": "Imprimir",
      //     "createState": "Criar estado",
      //     "removeAllStates": "Remover todos os estados",
      //     "removeState": "Remover",
      //     "renameState": "Renomear",
      //     "savedStates": "Estados salvos",
      //     "stateRestore": "Estado %d",
      //     "updateState": "Atualizar"
      //   },
      //   "searchBuilder": {
      //     "add": "Adicionar Condição",
      //     "button": {
      //       "0": "Construtor de Pesquisa",
      //       "_": "Construtor de Pesquisa (%d)"
      //     },
      //     "clearAll": "Limpar Tudo",
      //     "condition": "Condição",
      //     "conditions": {
      //       "date": {
      //         "after": "Depois",
      //         "before": "Antes",
      //         "between": "Entre",
      //         "empty": "Vazio",
      //         "equals": "Igual",
      //         "not": "Não",
      //         "notBetween": "Não Entre",
      //         "notEmpty": "Não Vazio"
      //       },
      //       "number": {
      //         "between": "Entre",
      //         "empty": "Vazio",
      //         "equals": "Igual",
      //         "gt": "Maior Que",
      //         "gte": "Maior Ou Igual Que",
      //         "lt": "Menor Que",
      //         "lte": "Menor Ou Igual Que",
      //         "not": "Não",
      //         "notBetween": "Não Entre",
      //         "notEmpty": "Não Vazio"
      //       },
      //       "string": {
      //         "contains": "Contém",
      //         "empty": "Vazio",
      //         "endsWith": "Termina Com",
      //         "equals": "Igual",
      //         "not": "Não",
      //         "notEmpty": "Não Vazio",
      //         "startsWith": "Começa Com",
      //         "notContains": "Não contém",
      //         "notStartsWith": "Não começa com",
      //         "notEndsWith": "Não termina com"
      //       },
      //       "array": {
      //         "without": "Sem",
      //         "notEmpty": "Não Vazio",
      //         "not": "Não",
      //         "contains": "Contém",
      //         "empty": "Vazio",
      //         "equals": "Igual"
      //       }
      //     },
      //     "data": "Data",
      //     "deleteTitle": "Excluir regra de filtragem",
      //     "logicAnd": "E",
      //     "logicOr": "Ou",
      //     "title": {
      //       "0": "Construtor de Pesquisa",
      //       "_": "Construtor de Pesquisa (%d)"
      //     },
      //     "value": "Valor",
      //     "leftTitle": "Critérios Externos",
      //     "rightTitle": "Critérios Internos"
      //   },
      //   "searchPanes": {
      //     "clearMessage": "Limpar Tudo",
      //     "collapse": {
      //       "0": "Painéis de Pesquisa",
      //       "_": "Painéis de Pesquisa (%d)"
      //     },
      //     "count": "{total}",
      //     "countFiltered": "{shown} ({total})",
      //     "emptyPanes": "Nenhum Painel de Pesquisa",
      //     "loadMessage": "Carregando Painéis de Pesquisa...",
      //     "title": "Filtros Ativos",
      //     "showMessage": "Mostrar todos",
      //     "collapseMessage": "Fechar todos"
      //   },
      //   "thousands": ".",
      //   "datetime": {
      //     "previous": "Anterior",
      //     "next": "Próximo",
      //     "hours": "Hora",
      //     "minutes": "Minuto",
      //     "seconds": "Segundo",
      //     "unknown": "-",
      //     "amPm": ["am", "pm"],
      //     "weekdays": ["Dom", "Seg", "Ter", "Qua", "Qui", "Sex", "Sáb"],
      //     "months": ["Janeiro", "Fevereiro", "Março", "Abril", "Maio", "Junho", "Julho", "Agosto", "Setembro", "Outubro", "Novembro", "Dezembro"]
      //   },
      //   "editor": {
      //     "close": "Fechar",
      //     "create": {
      //       "button": "Novo",
      //       "title": "Criar novo registro",
      //       "submit": "Criar"
      //     },
      //     "edit": {
      //       "button": "Editar",
      //       "title": "Editar registro",
      //       "submit": "Atualizar"
      //     },
      //     "remove": {
      //       "button": "Remover",
      //       "title": "Remover registro",
      //       "submit": "Remover",
      //       "confirm": {
      //         "_": "Deseja realmente remover %d registros?",
      //         "1": "Deseja realmente remover 1 registro?"
      //       }
      //     },
      //     "error": {
      //       "system": "Ocorreu um erro no sistema (<a target=\"\\\" rel=\"nofollow\" href=\"\\\">Mais informações</a>)."
      //     },
      //     "multi": {
      //       "title": "Múltiplos valores",
      //       "info": "Os itens selecionados contêm valores diferentes para esta entrada. Para editar e definir todos os itens para esta entrada com o mesmo valor, clique ou toque aqui, caso contrário, eles manterão seus valores individuais.",
      //       "restore": "Desfazer alterações",
      //       "noMulti": "Esta entrada pode ser editada individualmente, mas não como parte de um grupo."
      //     }
      //   },
      //   "stateRestore": {
      //     "creationModal": {
      //       "button": "Criar",
      //       "columns": {
      //         "search": "Busca de colunas",
      //         "visible": "Visibilidade da coluna"
      //       },
      //       "name": "Nome:",
      //       "order": "Ordernar",
      //       "paging": "Paginação",
      //       "scroller": "Posição da barra de rolagem",
      //       "search": "Busca",
      //       "searchBuilder": "Construtor de pesquisa",
      //       "select": "Selecionar",
      //       "title": "Criar novo estado",
      //       "toggleLabel": "Inclui:"
      //     },
      //     "duplicateError": "Já existe um estado com este nome.",
      //     "emptyError": "Não pode ser vazio.",
      //     "emptyStates": "Nenhum estado salvo",
      //     "removeConfirm": "Confirma remover %s?",
      //     "removeError": "Falha ao remover estado.",
      //     "removeJoiner": "e",
      //     "removeSubmit": "Remover",
      //     "removeTitle": "Remover estado",
      //     "renameButton": "Renomear",
      //     "renameLabel": "Novo nome para %s:",
      //     "renameTitle": "Renomear estado"
      //   },
      //   "decimal": ","
      // }, // FIM DO OBJETO LANGUAGE COMPLETO
      columnDefs: [
        { targets: '_all', visible: false },
        { targets: [1, 2, 3, 6], visible: true }
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
          try { $(this).prop('checked', api.column(idx).visible()); } catch (e) { console.warn("Erro visibilidade coluna:", idx, e); }
        });
        $('.col-vis').off('change').on('change', function() {
          const idx = +$(this).data('col');
          try { const col = api.column(idx); col.visible(!col.visible()); } catch (e) { console.warn("Erro alternar visibilidade:", idx, e); }
        });
        $('#multiSearchFields .multi-search-row').each(function() { updateLicenseInputState($(this)); });
      },
      rowCallback: function(row, data) {
        $(row).removeClass('conflict dup-license');
        if (nameConflicts.has(nameKey(data))) $(row).addClass('conflict');
        if (dupLicUsers.has(data.Id)) $(row).addClass('dup-license');
      },
      drawCallback: renderAlerts
    });

    $(table.table().node()).on('preDraw.dt', function() { showLoader(); });
    $(table.table().node()).on('draw.dt', function() { hideLoader(); });
  }

  function updateLicenseInputState($row) {
    const $select = $row.find('.column-select');
    const $input = $row.find('.search-input');
    const $licenseSelect = $row.find('.license-select');
    if ($select.val() == 6) {
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

    if ($ct.children().length === 0) { addSearchField(); }
    _executeMultiSearchLogic();
  }

  function _executeMultiSearchLogic() {
    const op = $('#multiSearchOperator').val();
    const $searchCriteria = $('#searchCriteria');

    if (!table) {
      $searchCriteria.text(allData.length === 0 ? 'Nenhum dado carregado.' : 'Tabela não disponível/inicializada.');
      return;
    }

    table.search('').columns().search('');
    while ($.fn.dataTable.ext.search.length > 0) { $.fn.dataTable.ext.search.pop(); }

    const rows = $('#multiSearchFields .multi-search-row');
    const filters = [];
    rows.each(function() {
      const idx = $(this).find('.column-select').val();
      let val = (idx == 6) ? $(this).find('.license-select').val() : $(this).find('.search-input').val().trim();
      if (val) { filters.push({ col: parseInt(idx, 10), term: val }); }
    });
    
    let criteriaText = op === 'AND' ? 'Critério: Todos os filtros devem corresponder' : 'Critério: Qualquer filtro deve corresponder';
    if (filters.length > 0) {
      criteriaText += ` (${filters.length} filtro(s) ativo(s))`;
      $.fn.dataTable.ext.search.push(
        function(settings, data, dataIndex) {
          if (settings.nTable.id !== table.table().node().id) return true;
          const rowData = table.row(dataIndex).data();
          if (!rowData) return false;
          const logicFn = op === 'OR' ? filters.some.bind(filters) : filters.every.bind(filters);
          return logicFn(filter => {
            if (filter.col === 6) {
              return (rowData.Licenses && Array.isArray(rowData.Licenses)) ?
                rowData.Licenses.some(l => (l.LicenseName || '').toLowerCase() === filter.term.toLowerCase()) : false;
            }
            const cellValue = data[filter.col] || '';
            return cellValue.toString().toLowerCase().includes(filter.term.toLowerCase());
          });
        }
      );
    } else {
      criteriaText = 'Critério: Todos os resultados (nenhum filtro ativo)';
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
    if (allData.length > 0) { setupMultiSearch(); } else { $('#searchCriteria').text('Nenhum dado carregado.'); }
    while ($.fn.dataTable.ext.search.length > 0) { $.fn.dataTable.ext.search.pop(); }
    if(table) table.draw();
    $('#alertPanel').empty();
    $('#colContainer .col-vis').each(function() {
      const idx = +$(this).data('col');
      const isDefaultVisible = [1, 2, 3, 6].includes(idx);
      if (table) { try { table.column(idx).visible(isDefaultVisible); } catch(e) {} }
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
    }
  }

 $('#exportCsv').on('click', () => {
    console.log("[Export CSV] Iniciando exportação de licenças...");
    if (!table) {
      console.log("[Export CSV] Tabela não iniciada.");
      return alert('Tabela não iniciada. Dados não carregados.');
    }
    
    const rowsToExport = table.rows({ search: 'applied' }).data().toArray();
    console.log("[Export CSV] Número de linhas para exportar (após filtros):", rowsToExport.length);

    if (!rowsToExport.length) {
      console.log("[Export CSV] Nenhuma linha para exportar com filtros atuais.");
      return alert('Nenhum registro para exportar com os filtros atuais.');
    }

    const visibleColumns = [];
    table.columns(':visible').every(function() {
      const columnTitle = $(table.table().header()).find('tr:eq(0) th:eq(' + this.index() + ')').text();
      visibleColumns.push({
        title: columnTitle,
        dataProp: table.settings()[0].aoColumns[this.index()].mData
      });
    });
    console.log("[Export CSV] Colunas visíveis para exportação:", JSON.stringify(visibleColumns));

    const headerRow = visibleColumns.map(col => escapeCsvValue(col.title)).join(',');
    console.log("[Export CSV] Linha de cabeçalho:", headerRow);

    const csvRows = rowsToExport.map(rowData => {
      return visibleColumns.map(col => {
        let cellData = rowData[col.dataProp];
        let shouldForceQuotesForSemicolon = false;

        if (col.dataProp === 'BusinessPhones') {
          cellData = Array.isArray(cellData) ? cellData.join('; ') : (cellData || '');
          if (String(cellData).includes(';')) {
            shouldForceQuotesForSemicolon = true;
          }
        } else if (col.dataProp === 'Licenses') {
          const licensesArray = (rowData.Licenses && Array.isArray(rowData.Licenses)) ?
                                rowData.Licenses.map(l => l.LicenseName || '') : [];
          if (licensesArray.length > 0) {
            cellData = licensesArray.join('; ');
            if (licensesArray.length > 1 || String(cellData).includes(';')) {
              shouldForceQuotesForSemicolon = true;
            }
          } else {
            cellData = '';
          }
        }
        return escapeCsvValue(cellData, shouldForceQuotesForSemicolon);
      }).join(',');
    });
    // Para ver algumas linhas de dados processadas:
    // console.log("[Export CSV] Primeiras 3 linhas de dados CSV:", csvRows.slice(0, 3));


    const csvContent = [headerRow, ...csvRows].join('\n');
    console.log("[Export CSV] Conteúdo CSV gerado (primeiros 500 caracteres):", csvContent.substring(0, 500));
    console.log("[Export CSV] Comprimento total do csvContent:", csvContent.length);

    if (csvContent.trim() === "" || (csvContent.trim() === headerRow.trim() && csvRows.length === 0 && headerRow.trim() !== "")) {
        // Se o conteúdo for vazio, ou apenas o cabeçalho (e o cabeçalho em si não for vazio)
        console.warn("[Export CSV] csvContent está vazio ou contém apenas cabeçalho (se houver dados no headerRow)!");
    } else if (csvContent.trim() === "" && headerRow.trim() === "") {
        console.warn("[Export CSV] csvContent está completamente vazio (sem cabeçalho, sem dados)!");
    }


    downloadCsv(csvContent, 'relatorio_licencas.csv');
    console.log("[Export CSV] Download solicitado.");
  });

  $('#exportIssues').on('click', () => {
    if (!allData.length) return alert('Nenhum dado carregado para gerar relatório de falhas.');
    let lines = [];
    if (nameConflicts.size) {
      lines.push(['CONFLITOS Nome+Unidade']);
      lines.push(['Nome', 'Unidade'].map(h => escapeCsvValue(h)));
      nameConflicts.forEach(k => lines.push(k.split('|||').map(v => escapeCsvValue(v))));
      lines.push([]);
    }
    if (dupLicUsers.size) {
      lines.push(['USUÁRIOS com Licenças Duplicadas']);
      lines.push(['Nome', 'Unidade', 'Licenças Duplicadas', 'Possui Licença Paga?'].map(h => escapeCsvValue(h)));
      allData.filter(u => dupLicUsers.has(u.Id)).forEach(u => {
        const cnt = {}, dups = [];
        (u.Licenses || []).forEach(l => cnt[l.LicenseName] = (cnt[l.LicenseName] || 0) + 1);
        Object.entries(cnt).forEach(([l, c]) => c > 1 && dups.push(l));
        const joinedDups = dups.join('; '); // Junção para CSV
        const hasPaid = (u.Licenses || []).some(l => !(l.LicenseName || '').toLowerCase().includes('free'));
        // Força aspas para o campo de licenças duplicadas se contiver ';' (ou seja, múltiplas duplicatas)
        lines.push([
            escapeCsvValue(u.DisplayName), 
            escapeCsvValue(u.OfficeLocation), 
            escapeCsvValue(joinedDups, dups.length > 1 || joinedDups.includes(';')), 
            escapeCsvValue(hasPaid ? 'Sim' : 'Não')
        ]);
      });
    }
    if (!lines.length) return alert('Nenhuma falha detectada.');
    const csvContent = lines.map(rowArray => rowArray.join(',')).join('\n');
    downloadCsv(csvContent, 'relatorio_falhas.csv');
  });

  try {
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
