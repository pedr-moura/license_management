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
    const MAX_DROPDOWN_OPTIONS_DISPLAYED = 50; // Optimization for custom dropdown

    function showLoader(message = 'Processando...') {
        if ($('#loadingOverlay').length === 0 && $('body').length > 0) {
            $('body').append(`<div id="loadingOverlay" style="display:none; position:fixed; top:0; left:0; width:100%; height:100%; background:rgba(0,0,0,0.5); z-index:10000; display:flex; align-items:center; justify-content:center;"><div id="loaderMessage" style="background:white; color:black; padding:20px; border-radius:5px;">${message}</div></div>`);
        }
        $('#loaderMessage').text(message);
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

    // --- Functions to be potentially moved to or called by a Web Worker ---
    function validateJsonForWorker(data) { // Renamed to avoid conflict if worker is inlined
        if (!Array.isArray(data)) {
            // In a real worker, you'd postMessage back an error
            console.error('JSON Inválido: Deve ser um array de objetos.');
            return { error: 'JSON Inválido: Deve ser um array de objetos.', validatedData: [] };
        }
        const validatedData = data.map((u, i) => u && typeof u === 'object' ? {
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
        return { validatedData };
    }

    function findIssuesForWorker(data) { // Renamed
        const nameMap = new Map();
        const dupSet = new Set();
        const officeLic = new Set([
            'Microsoft 365 E3', 'Microsoft 365 E5',
            'Microsoft 365 Business Standard', 'Microsoft 365 Business Premium',
            'Office 365 E3', 'Office 365 E5'
        ]);
        data.forEach(u => {
            const k = `${u.DisplayName || ''}|||${u.OfficeLocation || ''}`; // Inlined nameKey for worker context
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

    function calculateUniqueFieldValuesForWorker(data, config) { // Renamed
        const localUniqueFieldValues = {};
        config.forEach(colConfig => {
            if (colConfig.useDropdown) {
                if (colConfig.dataProp === 'Licenses') {
                    const allLicenseObjects = data.flatMap(user => user.Licenses || []).filter(l => l.LicenseName);
                    localUniqueFieldValues.Licenses = [...new Set(allLicenseObjects.map(l => l.LicenseName))].sort((a, b) => a.toLowerCase().localeCompare(b.toLowerCase()));
                } else {
                    localUniqueFieldValues[colConfig.dataProp] = [...new Set(data.map(user => user[colConfig.dataProp]).filter(value => value && String(value).trim() !== ''))].sort((a, b) => a.toLowerCase().localeCompare(b.toLowerCase()));
                }
            }
        });
        return { uniqueFieldValues: localUniqueFieldValues };
    }
    // --- End of functions for Web Worker ---

    function renderAlerts() {
        const $p = $('#alertPanel').empty();
        if (nameConflicts.size) {
            $p.append(`<div class="alert-badge"><button id="filterNameConflicts" style="background: none; border: none; color: inherit; cursor: pointer; text-decoration: underline;" class="underline"><i class="fas fa-users-slash" style="margin-right: 0.3rem;"></i>Conflitos Nome+Local: ${nameConflicts.size}</button></div>`);
        }
        if (dupLicUsers.size) {
            // Optimization: Debounce or paginate if this list is huge. For now, keep as is.
            // Consider only showing a count and a button to "View All in Table" or a modal for very long lists.
            const usersToList = allData.filter(u => dupLicUsers.has(u.Id));
            let listHtml = '';
            const maxPreview = 10; // Show first N items
            usersToList.slice(0, maxPreview).forEach(u => {
                const licCount = {}, duplicateLicNames = [];
                (u.Licenses || []).forEach(l => { licCount[l.LicenseName] = (licCount[l.LicenseName] || 0) + 1; });
                Object.entries(licCount).forEach(([licName, count]) => count > 1 && duplicateLicNames.push(licName));
                const hasPaid = (u.Licenses || []).some(l => !(l.LicenseName || '').toLowerCase().includes('free'));
                listHtml += `<li>${escapeHtml(u.DisplayName)} (${escapeHtml(u.OfficeLocation)}): ${escapeHtml(duplicateLicNames.join(', '))} — Pago: ${hasPaid ? 'Sim' : 'Não'}</li>`;
            });
             if (usersToList.length > maxPreview) {
                listHtml += `<li>E mais ${usersToList.length - maxPreview} usuário(s)...</li>`;
            }

            $p.append(`<div class="alert-badge"><span><i class="fas fa-copy" style="margin-right: 0.3rem;"></i>Licenças Duplicadas: ${dupLicUsers.size}</span> <button class="underline toggle-details" data-target="dupDetails" style="background: none; border: none; color: inherit; cursor: pointer; text-decoration: underline;">Detalhes</button></div><div id="dupDetails" class="alert-details"><ul>${listHtml}</ul></div>`);
        }
        $('#filterNameConflicts').off('click').on('click', function() {
            if (!table) return;
            showLoader('Filtrando conflitos...');
            setTimeout(() => { // Allow loader to show
                const conflictUserIds = allData.filter(u => nameConflicts.has(nameKey(u))).map(u => u.Id);
                table.search('').columns().search('');
                // Optimize regex for large number of IDs by chunking or finding a more performant DataTables search method if this becomes slow.
                // For now, this is standard.
                if (conflictUserIds.length > 0) {
                    table.column(0).search('^(' + conflictUserIds.join('|') + ')$', true, false).draw();
                } else {
                    table.column(0).search('').draw(); // Clear search if no conflicts (should not happen if button is shown)
                }
                // hideLoader is called by table draw.dt event
            }, 50);
        });
        $('.toggle-details').off('click').on('click', function() {
            const targetId = $(this).data('target');
            $(`#${targetId}`).toggleClass('show');
            $(this).text($(`#${targetId}`).hasClass('show') ? 'Esconder' : 'Detalhes');
        });
    }

    function initTable(data) {
        allData = data; // uniqueFieldValues, nameConflicts, dupLicUsers are now set by worker callback

        if ($('#licenseDatalist').length === 0) { $('body').append('<datalist id="licenseDatalist"></datalist>'); }
        const $licenseDatalist = $('#licenseDatalist').empty();
        if (uniqueFieldValues.Licenses) {
            uniqueFieldValues.Licenses.forEach(name => { $licenseDatalist.append($('<option>').attr('value', escapeHtml(name))); });
        }

        if (table) { $(table.table().node()).off('preDraw.dt draw.dt'); table.destroy(); $('#licenseTable').empty(); } // Empty to clear old header/footer

        table = $('#licenseTable').DataTable({
            data: allData,
            deferRender: true, // Crucial for performance with large datasets
            pageLength: 25, // Keep page length reasonable
            orderCellsTop: true,
            columns: [
                { data: 'Id', title: 'ID', visible: false },
                { data: 'DisplayName', title: 'Nome', visible: true },
                { data: 'Email', title: 'Email', visible: true },
                { data: 'JobTitle', title: 'Cargo', visible: true },
                { data: 'OfficeLocation', title: 'Local', visible: false },
                { data: 'BusinessPhones', title: 'Telefones', visible: false, render: p => Array.isArray(p) ? p.join('; ') : (p || '') },
                { data: 'Licenses', title: 'Licenças', visible: true, render: l => Array.isArray(l) ? l.map(x => x.LicenseName || '').filter(name => name).join(', ') : '' }
            ],
            // language: { searchPlaceholder: "Filtrar tabela..." }, // Example for search placeholder
            initComplete: function() {
                const api = this.api();
                // Column individual search (header inputs) - this is standard DataTables functionality
                // No major changes here, but ensure it's not causing re-draws too frequently if complex
                api.columns().every(function(colIdx) {
                    const column = this;
                    // Debounce individual column search if needed, though DataTables might have its own debounce
                    // For simplicity, keeping original logic. If typing in these is slow, add debounce.
                    $(api.table().header()).find('tr:eq(1) th:eq(' + colIdx + ') input')
                        .off('keyup change clear').on('keyup change clear', function() {
                            if (column.search() !== this.value) {
                                column.search(this.value).draw();
                            }
                        });
                });

                $('#colContainer .col-vis').each(function() {
                    const idx = +$(this).data('col');
                    try {
                        if (idx >= 0 && idx < api.columns().nodes().length) {
                            $(this).prop('checked', api.column(idx).visible());
                        } else { $(this).prop('disabled', true); }
                    } catch (e) { console.warn("Erro ao checar visibilidade da coluna:", idx, e); }
                });
                $('.col-vis').off('change').on('change', function() {
                    const idx = +$(this).data('col');
                    try {
                        const col = api.column(idx);
                        if (col && col.visible) { col.visible(!col.visible()); }
                        else { console.warn("Coluna não encontrada para índice:", idx); }
                    } catch (e) { console.warn("Erro ao alternar visibilidade da coluna:", idx, e); }
                });

                // Multi-search related UI updates
                $('#multiSearchFields .multi-search-row').each(function() { updateSearchFieldUI($(this)); });
                applyMultiSearch(); // Initial application of multi-search if any fields exist
            },
            rowCallback: function(row, data) {
                // Minimize class manipulation if possible, but this is generally fine.
                let classes = '';
                if (nameConflicts.has(nameKey(data))) classes += ' conflict';
                if (dupLicUsers.has(data.Id)) classes += ' dup-license';
                if (classes) $(row).addClass(classes.trim());
                else $(row).removeClass('conflict dup-license'); // Ensure removal if state changes
            },
            drawCallback: function() {
                renderAlerts(); // Update alerts based on current data/filters
                hideLoader(); // Hide loader after table draw
            }
        });
        $(table.table().node()).on('preDraw.dt', () => showLoader('Atualizando tabela...'));
        // 'draw.dt' will call hideLoader via drawCallback
    }


    function updateSearchFieldUI($row) {
        const selectedColIndex = $row.find('.column-select').val();
        const columnConfig = searchableColumnsConfig.find(c => c.index == selectedColIndex);

        const $searchInput = $row.find('.search-input');
        const $customDropdownContainer = $row.find('.custom-dropdown-container');
        const $customDropdownTextInput = $row.find('.custom-dropdown-text-input');
        const $customOptionsList = $row.find('.custom-options-list');
        const $hiddenValueSelect = $row.find('.search-value-select');

        $customDropdownTextInput.off();
        $customOptionsList.off();

        if (columnConfig && columnConfig.useDropdown) {
            $searchInput.hide();
            $customDropdownContainer.show();
            $customDropdownTextInput.val('');
            $customDropdownTextInput.attr('placeholder', `Digite ou selecione ${columnConfig.title.toLowerCase()}`);
            $customOptionsList.hide().empty();

            $hiddenValueSelect.empty().append($('<option>').val('').text(''));
            const allUniqueOptions = uniqueFieldValues[columnConfig.dataProp] || [];
            // No need to populate hidden select with all options if we are not using it directly for display
            // It will be updated when an option is selected from the custom list.
            $hiddenValueSelect.val('');

            let filterDebounce;
            $customDropdownTextInput.on('input', function() {
                clearTimeout(filterDebounce);
                const $input = $(this);
                filterDebounce = setTimeout(() => {
                    const searchTerm = $input.val().toLowerCase();
                    $customOptionsList.empty().show();
                    const filteredOptions = allUniqueOptions.filter(opt => String(opt).toLowerCase().includes(searchTerm));

                    if (filteredOptions.length === 0) {
                        $customOptionsList.append('<div class="custom-option-item no-results">Nenhum resultado encontrado</div>');
                    } else {
                        // OPTIMIZATION: Limit the number of options displayed to avoid huge DOM
                        filteredOptions.slice(0, MAX_DROPDOWN_OPTIONS_DISPLAYED).forEach(opt => {
                            const $optionEl = $('<div class="custom-option-item"></div>').text(opt).data('value', opt);
                            $customOptionsList.append($optionEl);
                        });
                        if (filteredOptions.length > MAX_DROPDOWN_OPTIONS_DISPLAYED) {
                             $customOptionsList.append(`<div class="custom-option-item no-results" style="font-style:italic; color: #555;">Mais ${filteredOptions.length - MAX_DROPDOWN_OPTIONS_DISPLAYED} opções ocultas...</div>`);
                        }
                    }
                }, 200); // Debounce filtering within dropdown input
            });

            $customDropdownTextInput.on('focus', function() {
                $(this).trigger('input');
                $customOptionsList.show();
            });

            $customOptionsList.on('mousedown', '.custom-option-item', function(e) {
                e.preventDefault();
                if ($(this).hasClass('no-results')) return;
                const selectedText = $(this).text();
                const selectedValue = $(this).data('value');
                $customDropdownTextInput.val(selectedText);
                $hiddenValueSelect.val(selectedValue).trigger('change'); // Crucial: update hidden select
                $customOptionsList.hide();
            });

            let blurTimeout;
            $customDropdownTextInput.on('blur', function() {
                clearTimeout(blurTimeout);
                blurTimeout = setTimeout(() => { $customOptionsList.hide(); }, 150);
            });

        } else {
            $searchInput.show().val('');
            $customDropdownContainer.hide();
            $hiddenValueSelect.empty().hide(); // Ensure it's hidden if not used
            $searchInput.attr('placeholder', 'Termo...');
        }
    }

    function setupMultiSearch() {
        const $container = $('#multiSearchFields');
        function addSearchField() {
            const columnOptions = searchableColumnsConfig
                .map(c => `<option value="${c.index}">${escapeHtml(c.title)}</option>`)
                .join('');
            const $row = $(`
                <div class="multi-search-row">
                    <select class="column-select">${columnOptions}</select>
                    <input class="search-input" placeholder="Termo..." style="display:none;" />
                    <div class="custom-dropdown-container" style="position: relative; display:none;">
                        <input type="text" class="custom-dropdown-text-input" autocomplete="off" />
                        <div class="custom-options-list" style="display:none; position:absolute; background:white; border:1px solid #ccc; z-index:100; max-height: 200px; overflow-y:auto;"></div>
                    </div>
                    <select class="search-value-select" style="display: none;"></select>
                    <button class="remove-field" title="Remover filtro"><i class="fas fa-trash-alt"></i></button>
                </div>
            `);
            $container.append($row);
            updateSearchFieldUI($row);

            $row.find('.column-select').on('change', function() {
                updateSearchFieldUI($row);
                const selectedColIndex = $row.find('.column-select').val();
                const columnConfig = searchableColumnsConfig.find(c => c.index == selectedColIndex);
                if (columnConfig && columnConfig.useDropdown) {
                    $row.find('.search-value-select').val('').trigger('change');
                } else {
                    $row.find('.search-input').val('').trigger('input'); // ensure input for non-dropdown is also triggered
                }
            });

            // Bind to the elements that actually trigger search logic
            $row.find('.search-input').on('input change', applyMultiSearch); // For text inputs
            $row.find('.search-value-select').on('change', applyMultiSearch); // For dropdowns (via hidden select)

            $row.find('.remove-field').on('click', function() {
                const $multiSearchRow = $(this).closest('.multi-search-row');
                $multiSearchRow.find('.custom-dropdown-text-input').off();
                $multiSearchRow.find('.custom-options-list').off();
                $multiSearchRow.remove();
                applyMultiSearch();
            });
        }
        $('#addSearchField').off('click').on('click', addSearchField);
        $('#multiSearchOperator').off('change').on('change', applyMultiSearch);

        if ($container.children().length === 0) {
            if (allData && allData.length > 0) {
                addSearchField();
            } else {
                $('#searchCriteria').text('Nenhum dado carregado. Carregue um arquivo JSON para começar.');
            }
        }
       // _executeMultiSearchLogic(); // This will be called by applyMultiSearch if needed
    }

    function _executeMultiSearchLogic() {
        const operator = $('#multiSearchOperator').val();
        const $searchCriteriaText = $('#searchCriteria');
        if (!table) {
            $searchCriteriaText.text(allData.length === 0 ? 'Nenhum dado carregado.' : 'Tabela não inicializada.');
            hideLoader(); // Ensure loader is hidden if table isn't ready
            return;
        }

        // Clear previous global search and custom filters
        table.search(''); // Clear global table search
        while ($.fn.dataTable.ext.search.length > 0) { $.fn.dataTable.ext.search.pop(); }

        const filters = [];
        $('#multiSearchFields .multi-search-row').each(function() {
            const colIndex = $(this).find('.column-select').val();
            const columnConfig = searchableColumnsConfig.find(c => c.index == colIndex);
            let searchTerm = '';
            if (columnConfig) {
                searchTerm = columnConfig.useDropdown ?
                    $(this).find('.search-value-select').val() : // Value from hidden select
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
                function(settings, apiData, dataIndex) { // apiData is search data, rowData is original
                    if (settings.nTable.id !== table.table().node().id) return true;
                    const rowData = table.row(dataIndex).data(); // Get full data object for the row
                    if (!rowData) return false;

                    const logicFn = operator === 'OR' ? filters.some.bind(filters) : filters.every.bind(filters);
                    return logicFn(filter => {
                        let cellValue;
                        // For 'Licenses', use the array of objects directly from rowData
                        if (filter.dataProp === 'Licenses') {
                            return (rowData.Licenses && Array.isArray(rowData.Licenses)) ?
                                rowData.Licenses.some(l => (l.LicenseName || '').toLowerCase() === filter.term.toLowerCase()) : false;
                        }
                        // For other dropdowns, use direct property access and exact match
                        else if (filter.isDropdown) {
                            cellValue = rowData[filter.dataProp] || '';
                            return String(cellValue).toLowerCase() === filter.term.toLowerCase();
                        }
                        // For text inputs, use DataTables' prepared search data (apiData) for "contains" match
                        else {
                             // apiData[filter.col] is the rendered data for the column, good for text search
                            cellValue = apiData[filter.col] || '';
                            return String(cellValue).toLowerCase().includes(filter.term.toLowerCase());
                        }
                    });
                }
            );
        } else { criteriaText = 'Critérios: Todos os resultados (sem filtros ativos)'; }

        $searchCriteriaText.text(criteriaText);
        table.draw(); // This will trigger preDraw and drawCallback (which includes hideLoader)
    }

    function applyMultiSearch() {
        clearTimeout(multiSearchDebounceTimer);
        showLoader('Aplicando filtros...'); // Show loader immediately
        multiSearchDebounceTimer = setTimeout(_executeMultiSearchLogic, DEBOUNCE_DELAY);
    }

    $('#clearFilters').on('click', () => {
        showLoader('Limpando filtros...');
        setTimeout(() => { // Allow loader to show
            if (table) {
                $(table.table().header()).find('tr:eq(1) th input').val('');
                table.search('').columns().search(''); // Clear individual column searches
            }
            $('#multiSearchFields .multi-search-row').each(function() {
                $(this).find('.custom-dropdown-text-input').off();
                $(this).find('.custom-options-list').off();
            });
            $('#multiSearchFields').empty();
            if (allData && allData.length > 0) {
                setupMultiSearch(); // This will add one default search field and apply empty search
            } else {
                $('#searchCriteria').text('Nenhum dado carregado.');
            }
            while ($.fn.dataTable.ext.search.length > 0) { $.fn.dataTable.ext.search.pop(); }

            if (table) table.draw(); // This will trigger hideLoader via drawCallback
            else hideLoader();

            $('#alertPanel').empty(); // Clear alerts
            const defaultVisibleCols = [1, 2, 3, 6];
            $('#colContainer .col-vis').each(function() {
                const idx = +$(this).data('col');
                const isDefaultVisible = defaultVisibleCols.includes(idx);
                if (table && idx >= 0 && idx < table.columns().nodes().length) {
                    try { table.column(idx).visible(isDefaultVisible); }
                    catch (e) { console.warn("Erro ao resetar visibilidade da coluna:", idx, e); }
                }
                $(this).prop('checked', isDefaultVisible);
            });
             if (!table && !(allData && allData.length > 0)) {
                 $('#searchCriteria').text('Nenhum dado carregado. Carregue um arquivo JSON para começar.');
             }
        }, 50);
    });


    function downloadCsv(csvContent, fileName) {
        const bom = "\uFEFF";
        const blob = new Blob([bom, csvContent], { type: 'text/csv;charset=utf-8;' });
        const link = document.createElement('a');
        if (link.download !== undefined) {
            const url = URL.createObjectURL(blob);
            link.setAttribute('href', url); link.setAttribute('download', fileName);
            link.style.visibility = 'hidden'; document.body.appendChild(link);
            link.click(); document.body.removeChild(link); URL.revokeObjectURL(url);
        } else { alert("Seu navegador não suporta o download direto de arquivos."); }
    }

    $('#exportCsv').on('click', () => {
        if (!table) return alert('Tabela não inicializada. Nenhum dado carregado.');
        showLoader('Exportando CSV...');
        setTimeout(() => { // Allow loader to show
            const rowsToExport = table.rows({ search: 'applied' }).data().toArray();
            if (!rowsToExport.length) {
                hideLoader();
                return alert('Nenhum registro para exportar com os filtros atuais.');
            }
            const visibleColumns = [];
            table.columns(':visible').every(function() {
                const columnConfig = table.settings()[0].aoColumns[this.index()];
                // Ensure title is from the config or header, and mData is the correct prop
                const colTitle = $(table.column(this.index()).header()).text() || columnConfig.title;
                const dataProp = columnConfig.mData;
                visibleColumns.push({ title: colTitle, dataProp: dataProp });
            });

            const headerRow = visibleColumns.map(col => escapeCsvValue(col.title)).join(',');
            const csvRows = rowsToExport.map(rowData => {
                return visibleColumns.map(col => {
                    let cellData = rowData[col.dataProp]; let shouldForceQuotes = false;
                    if (col.dataProp === 'BusinessPhones') {
                        cellData = Array.isArray(cellData) ? cellData.join('; ') : (cellData || '');
                        if (String(cellData).includes(';')) shouldForceQuotes = true;
                    } else if (col.dataProp === 'Licenses') {
                        const licensesArray = (rowData.Licenses && Array.isArray(rowData.Licenses)) ?
                            rowData.Licenses.map(l => l.LicenseName || '').filter(name => name) : [];
                        cellData = licensesArray.length > 0 ? licensesArray.join('; ') : '';
                         if (String(cellData).includes(';')) shouldForceQuotes = true;
                    }
                    return escapeCsvValue(cellData, shouldForceQuotes || String(cellData).match(/[,"\n\r]/));
                }).join(',');
            });
            const csvContent = [headerRow, ...csvRows].join('\n');
            downloadCsv(csvContent, 'license_report.csv');
            hideLoader();
        }, 50);
    });

    $('#exportIssues').on('click', () => {
        if (!allData.length) return alert('Nenhum dado carregado para gerar o relatório de problemas.');
        showLoader('Gerando relatório de problemas...');
        setTimeout(() => { // Allow loader to show
            const lines = [];
            if (nameConflicts.size) {
                lines.push(['CONFLITOS DE NOME+LOCALIZAÇÃO']);
                lines.push(['Nome', 'Localização'].map(h => escapeCsvValue(h)));
                nameConflicts.forEach(key => lines.push(key.split('|||').map(value => escapeCsvValue(value))));
                lines.push([]); // Empty line for separation
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
                        escapeCsvValue(joinedDups, true), // Force quotes if it contains semicolons for clarity
                        escapeCsvValue(hasPaid ? 'Sim' : 'Não')
                    ]);
                });
            }
            if (!lines.length) { lines.push(['Nenhum problema detectado.']); }
            const csvContent = lines.map(rowArray => rowArray.join(',')).join('\n');
            downloadCsv(csvContent, 'issues_report.csv');
            hideLoader();
        }, 50);
    });


    // --- Initial Data Loading with Web Worker ---
    function processDataWithWorker(rawData) {
        showLoader('Validando e processando dados (isso pode levar alguns instantes)...');

        // Create a Blob URL for the worker script
        const workerScript = `
            // Minimal nameKey for worker context
            const nameKeyInternal = u => \`\${u.DisplayName || ''}|||\${u.OfficeLocation || ''}\`;

            // --- validateJsonForWorker (copied from main thread) ---
            function validateJsonForWorker(data) {
                if (!Array.isArray(data)) {
                    return { error: 'JSON Inválido: Deve ser um array de objetos.', validatedData: [] };
                }
                const validatedData = data.map((u, i) => u && typeof u === 'object' ? {
                    Id: u.Id || \`unknown_\${i}\`,
                    DisplayName: u.DisplayName || 'Unknown',
                    OfficeLocation: u.OfficeLocation || 'Unknown',
                    Email: u.Email || '',
                    JobTitle: u.JobTitle || 'Unknown',
                    BusinessPhones: Array.isArray(u.BusinessPhones) ? u.BusinessPhones : (typeof u.BusinessPhones === 'string' ? u.BusinessPhones.split('; ').filter(p => p) : []),
                    Licenses: Array.isArray(u.Licenses) ? u.Licenses.map(l => ({
                        LicenseName: l.LicenseName || \`Lic_\${i}_\${Math.random().toString(36).substr(2, 5)}\`,
                        SkuId: l.SkuId || ''
                    })).filter(l => l.LicenseName) : []
                } : null).filter(x => x);
                return { validatedData };
            }

            // --- findIssuesForWorker (copied from main thread) ---
            function findIssuesForWorker(data) {
                const nameMap = new Map();
                const dupSet = new Set();
                const officeLic = new Set([
                    'Microsoft 365 E3', 'Microsoft 365 E5',
                    'Microsoft 365 Business Standard', 'Microsoft 365 Business Premium',
                    'Office 365 E3', 'Office 365 E5'
                ]);
                data.forEach(u => {
                    const k = nameKeyInternal(u);
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
                // Convert Set to Array for transferring from worker
                const conflictingNameKeysArray = [...nameMap].filter(([, count]) => count > 1).map(([key]) => key);
                return { nameConflictsArray: conflictingNameKeysArray, dupLicUsersArray: Array.from(dupSet) };
            }

            // --- calculateUniqueFieldValuesForWorker (copied from main thread) ---
            function calculateUniqueFieldValuesForWorker(data, config) {
                const localUniqueFieldValues = {};
                config.forEach(colConfig => {
                    if (colConfig.useDropdown) {
                        if (colConfig.dataProp === 'Licenses') {
                            const allLicenseObjects = data.flatMap(user => user.Licenses || []).filter(l => l.LicenseName);
                            localUniqueFieldValues.Licenses = [...new Set(allLicenseObjects.map(l => l.LicenseName))].sort((a, b) => a.toLowerCase().localeCompare(b.toLowerCase()));
                        } else {
                             // Filter out empty or null values before creating the Set
                            localUniqueFieldValues[colConfig.dataProp] = [...new Set(data.map(user => user[colConfig.dataProp]).filter(value => value && String(value).trim() !== ''))].sort((a, b) => a.toLowerCase().localeCompare(b.toLowerCase()));
                        }
                    }
                });
                return { uniqueFieldValues: localUniqueFieldValues };
            }

            self.onmessage = function(e) {
                const { rawData, searchableColumnsConfig } = e.data;
                try {
                    const validationResult = validateJsonForWorker(rawData);
                    if (validationResult.error) {
                        self.postMessage({ error: validationResult.error });
                        return;
                    }
                    const validatedData = validationResult.validatedData;
                    if (validatedData.length === 0) {
                         self.postMessage({ validatedData: [], nameConflicts: new Set(), dupLicUsers: new Set(), uniqueFieldValues: {} });
                         return;
                    }

                    const issues = findIssuesForWorker(validatedData);
                    const uniqueValues = calculateUniqueFieldValuesForWorker(validatedData, searchableColumnsConfig);
                    
                    self.postMessage({
                        validatedData: validatedData,
                        nameConflictsArray: issues.nameConflictsArray, // Send as Array
                        dupLicUsersArray: issues.dupLicUsersArray,     // Send as Array
                        uniqueFieldValues: uniqueValues.uniqueFieldValues,
                        error: null
                    });
                } catch (err) {
                    self.postMessage({ error: 'Erro no Web Worker: ' + err.message + '\\n' + err.stack });
                } finally {
                    self.close(); // Close worker after processing
                }
            };
        `;
        const blob = new Blob([workerScript], { type: 'application/javascript' });
        const worker = new Worker(URL.createObjectURL(blob));

        worker.onmessage = function(e) {
            URL.revokeObjectURL(blob); // Clean up blob URL
            const { validatedData: processedData, nameConflictsArray, dupLicUsersArray, uniqueFieldValues: uFValues, error } = e.data;

            if (error) {
                hideLoader();
                alert('Erro ao processar dados no Worker: ' + error);
                console.error("Erro do Worker:", error);
                $('#searchCriteria').text('Erro ao carregar dados.');
                return;
            }

            // Convert arrays back to Sets
            nameConflicts = new Set(nameConflictsArray);
            dupLicUsers = new Set(dupLicUsersArray);
            uniqueFieldValues = uFValues;

            if (processedData && processedData.length > 0) {
                showLoader('Renderizando tabela...'); // Update loader message
                // Use setTimeout to allow the loader message to update before heavy table init
                setTimeout(() => {
                    initTable(processedData);
                    setupMultiSearch();
                    $('#searchCriteria').text(`Dados carregados (${processedData.length} usuários). Use filtros para refinar.`);
                    // hideLoader() will be called by initTable's drawCallback
                }, 50);
            } else {
                hideLoader();
                alert('Nenhum usuário válido encontrado nos dados. Verifique o arquivo JSON.');
                $('#searchCriteria').text('Nenhum dado válido carregado.');
                // Ensure UI is reset or shown as empty
                 if (table) { table.clear().draw(); } // Clear table if it exists
                 $('#multiSearchFields').empty();
                 $('#alertPanel').empty();
            }
        };

        worker.onerror = function(e) {
            URL.revokeObjectURL(blob); // Clean up blob URL
            hideLoader();
            console.error(`Erro no Web Worker: Linha ${e.lineno} em ${e.filename}: ${e.message}`);
            alert('Ocorreu um erro crítico durante o processamento de dados. Verifique o console.');
            $('#searchCriteria').text('Erro crítico ao carregar dados.');
        };

        // Send data to worker
        worker.postMessage({ rawData: rawData, searchableColumnsConfig: searchableColumnsConfig });
    }


    // --- Initial Load Trigger ---
    try {
        // Ensure userData is available globally or loaded via an input, e.g. file input
        if (typeof userData !== 'undefined' && Array.isArray(userData)) { // userData should be defined in your HTML
            processDataWithWorker(userData);
        } else {
             // Allow user to load JSON via file input as a fallback or primary method
            $('#jsonFileInput').on('change', function(event) {
                const file = event.target.files[0];
                if (file) {
                    showLoader('Lendo arquivo JSON...');
                    const reader = new FileReader();
                    reader.onload = function(e) {
                        try {
                            const jsonData = JSON.parse(e.target.result);
                            processDataWithWorker(jsonData);
                        } catch (err) {
                            hideLoader();
                            alert('Erro ao parsear o arquivo JSON: ' + err.message);
                            console.error("JSON Parse Error:", err);
                            $('#searchCriteria').text('Falha ao ler JSON.');
                        }
                    };
                    reader.onerror = function() {
                        hideLoader();
                        alert('Erro ao ler o arquivo.');
                         $('#searchCriteria').text('Falha ao ler arquivo.');
                    };
                    reader.readAsText(file);
                }
            });
            if ($('#jsonFileInput').length === 0) {
                 console.warn("Variável 'userData' não definida e input #jsonFileInput não encontrado. Adicione um input de arquivo ou defina 'userData'.");
                 $('#searchCriteria').html('Por favor, carregue um arquivo JSON de usuários. <br/> (Variável <code>userData</code> global não encontrada.)');
            } else {
                 $('#searchCriteria').text('Por favor, carregue um arquivo JSON de usuários.');
            }
            hideLoader(); // Hide if no initial userData
        }
    } catch (error) {
        hideLoader();
        alert('Erro ao iniciar o carregamento de dados: ' + error.message);
        console.error("Erro de carregamento inicial:", error);
        $('#searchCriteria').text('Erro ao carregar dados.');
    }
});
