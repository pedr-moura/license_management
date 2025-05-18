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
    const MAX_DROPDOWN_OPTIONS_DISPLAYED = 50;

    function showLoader(message = 'Processing...') {
        let $overlay = $('#loadingOverlay');
        if ($overlay.length === 0 && $('body').length > 0) {
            $('body').append('<div id="loadingOverlay"><div class="loader-content"><p id="loaderMessageText"></p></div></div>');
            $overlay = $('#loadingOverlay');
        }
        $overlay.find('#loaderMessageText').text(message);
        $overlay.css('display', 'flex');
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

    function renderAlerts() {
        const $alertPanel = $('#alertPanel').empty();
        if (nameConflicts.size) {
            $alertPanel.append(`<div class="alert-badge"><button id="filterNameConflicts" style="background: none; border: none; color: inherit; cursor: pointer; text-decoration: underline;" class="underline"><i class="fas fa-users-slash" style="margin-right: 0.3rem;"></i>Name+Location Conflicts: ${nameConflicts.size}</button></div>`);
        }
        if (dupLicUsers.size) {
            const usersToList = allData.filter(u => dupLicUsers.has(u.Id));
            let listHtml = '';
            const maxPreview = 10;
            usersToList.slice(0, maxPreview).forEach(u => {
                const licCount = {}, duplicateLicNames = [];
                (u.Licenses || []).forEach(l => { licCount[l.LicenseName] = (licCount[l.LicenseName] || 0) + 1; });
                Object.entries(licCount).forEach(([licName, count]) => count > 1 && duplicateLicNames.push(licName));
                const hasPaid = (u.Licenses || []).some(l => !(l.LicenseName || '').toLowerCase().includes('free'));
                listHtml += `<li>${escapeHtml(u.DisplayName)} (${escapeHtml(u.OfficeLocation)}): ${escapeHtml(duplicateLicNames.join(', '))} — Paid: ${hasPaid ? 'Yes' : 'No'}</li>`;
            });
             if (usersToList.length > maxPreview) {
                listHtml += `<li>And ${usersToList.length - maxPreview} more user(s)...</li>`;
            }
            $alertPanel.append(`<div class="alert-badge"><span><i class="fas fa-copy" style="margin-right: 0.3rem;"></i>Duplicate Licenses: ${dupLicUsers.size}</span> <button class="underline toggle-details" data-target="dupDetails" style="background: none; border: none; color: inherit; cursor: pointer; text-decoration: underline;">Details</button></div><div id="dupDetails" class="alert-details"><ul>${listHtml}</ul></div>`);
        }

        $('#filterNameConflicts').off('click').on('click', function() {
            if (!table) return;
            showLoader('Filtering conflicts...');
            setTimeout(() => {
                const conflictUserIds = allData.filter(u => nameConflicts.has(nameKey(u))).map(u => u.Id);
                table.search('').columns().search('');
                if (conflictUserIds.length > 0) {
                    table.column(0).search('^(' + conflictUserIds.join('|') + ')$', true, false).draw();
                } else {
                    table.column(0).search('').draw();
                }
            }, 50);
        });

        $('.toggle-details').off('click').on('click', function() {
            const targetId = $(this).data('target');
            $(`#${targetId}`).toggleClass('show');
            $(this).text($(`#${targetId}`).hasClass('show') ? 'Hide' : 'Details');
        });
    }

    function initTable(data) {
        allData = data;
        if ($('#licenseDatalist').length === 0) { $('body').append('<datalist id="licenseDatalist"></datalist>'); }
        const $licenseDatalist = $('#licenseDatalist').empty();
        if (uniqueFieldValues.Licenses) {
            uniqueFieldValues.Licenses.forEach(name => { $licenseDatalist.append($('<option>').attr('value', escapeHtml(name))); });
        }

        if (table) { $(table.table().node()).off('preDraw.dt draw.dt'); table.destroy(); $('#licenseTable').empty(); }

        table = $('#licenseTable').DataTable({
            data: allData,
            deferRender: true,
            pageLength: 25,
            orderCellsTop: true,
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
                $('#colContainer .col-vis').each(function() {
                    const idx = +$(this).data('col');
                    try {
                        if (idx >= 0 && idx < api.columns().nodes().length) {
                            $(this).prop('checked', api.column(idx).visible());
                        } else { $(this).prop('disabled', true); }
                    } catch (e) { console.warn("Error checking column visibility:", idx, e); }
                });
                $('.col-vis').off('change').on('change', function() {
                    const idx = +$(this).data('col');
                    try {
                        const col = api.column(idx);
                        if (col && col.visible) { col.visible(!col.visible()); }
                        else { console.warn("Column not found for index:", idx); }
                    } catch (e) { console.warn("Error toggling column visibility:", idx, e); }
                });

                if ($('#multiSearchFields .multi-search-row').length > 0) {
                    applyMultiSearch();
                }
                 // Inicializa o status do botão da IA
                updateAiFeatureStatus();
            },
            rowCallback: function(row, data) {
                let classes = '';
                if (nameConflicts.has(nameKey(data))) classes += ' conflict';
                if (dupLicUsers.has(data.Id)) classes += ' dup-license';
                if (classes) $(row).addClass(classes.trim());
                else $(row).removeClass('conflict dup-license');
            },
            drawCallback: function() {
                renderAlerts();
                hideLoader();
            }
        });
        $(table.table().node()).on('preDraw.dt', () => showLoader('Updating table...'));
    }

    function updateSearchFieldUI($row) {
        const selectedColIndex = $row.find('.column-select').val();
        const columnConfig = searchableColumnsConfig.find(c => c.index == selectedColIndex);

        const $searchInput = $row.find('.search-input');
        const $customDropdownContainer = $row.find('.custom-dropdown-container');
        const $customDropdownTextInput = $row.find('.custom-dropdown-text-input');
        const $customOptionsList = $row.find('.custom-options-list');
        const $hiddenValueInput = $row.find('.search-value-input');

        $customDropdownTextInput.off('input focus blur');
        $customOptionsList.off('mousedown', '.custom-option-item');

        if (columnConfig && columnConfig.useDropdown) {
            $searchInput.hide();
            $customDropdownContainer.show();
            $customDropdownTextInput.val('');
            $customDropdownTextInput.attr('placeholder', `Type or select ${columnConfig.title.toLowerCase()}`);
            $customOptionsList.hide().empty();
            $hiddenValueInput.val('');
            const allUniqueOptions = uniqueFieldValues[columnConfig.dataProp] || [];

            let filterDebounce;
            $customDropdownTextInput.on('input', function() {
                clearTimeout(filterDebounce);
                const $input = $(this);
                filterDebounce = setTimeout(() => {
                    const searchTerm = $input.val().toLowerCase();
                    $customOptionsList.empty().show();
                    const filteredOptions = allUniqueOptions.filter(opt => String(opt).toLowerCase().includes(searchTerm));

                    if (filteredOptions.length === 0) {
                        $customOptionsList.append('<div class="custom-option-item no-results">No results found</div>');
                    } else {
                        filteredOptions.slice(0, MAX_DROPDOWN_OPTIONS_DISPLAYED).forEach(opt => {
                            const $optionEl = $('<div class="custom-option-item"></div>').text(opt).data('value', opt);
                            $customOptionsList.append($optionEl);
                        });
                        if (filteredOptions.length > MAX_DROPDOWN_OPTIONS_DISPLAYED) {
                             $customOptionsList.append(`<div class="custom-option-item no-results" style="font-style:italic; color: #aaa;">${filteredOptions.length - MAX_DROPDOWN_OPTIONS_DISPLAYED} more options hidden...</div>`);
                        }
                    }
                }, 200);
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
                $hiddenValueInput.val(selectedValue).trigger('change');
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
            $hiddenValueInput.val('').hide();
            $searchInput.attr('placeholder', 'Term...');
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
                    <input class="search-input" placeholder="Term..." />
                    <div class="custom-dropdown-container" style="display:none;">
                        <input type="text" class="custom-dropdown-text-input" autocomplete="off" />
                        <div class="custom-options-list"></div>
                    </div>
                    <input type="hidden" class="search-value-input" />
                    <button class="remove-field" title="Remove filter"><i class="fas fa-trash-alt"></i></button>
                </div>
            `);
            $container.append($row);
            updateSearchFieldUI($row);

            $row.find('.column-select').on('change', function() {
                updateSearchFieldUI($row);
                const selectedColIndex = $row.find('.column-select').val();
                const columnConfig = searchableColumnsConfig.find(c => c.index == selectedColIndex);
                if (columnConfig && columnConfig.useDropdown) {
                    $row.find('.search-value-input').val('').trigger('change');
                } else {
                    $row.find('.search-input').val('').trigger('input');
                }
            });

            $row.find('.search-input').on('input change', applyMultiSearch);
            $row.find('.search-value-input').on('change', applyMultiSearch);

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

        if ($container.children().length === 0 && allData && allData.length > 0) {
            addSearchField();
        } else if (!(allData && allData.length > 0)) {
             $('#searchCriteria').text('No data loaded. Cannot set up search fields.');
        }
    }

    function applyMultiSearch() {
        clearTimeout(multiSearchDebounceTimer);
        showLoader('Applying filters...');
        multiSearchDebounceTimer = setTimeout(_executeMultiSearchLogic, DEBOUNCE_DELAY);
    }

    function _executeMultiSearchLogic() {
        const operator = $('#multiSearchOperator').val();
        const $searchCriteriaText = $('#searchCriteria');
        if (!table) {
            $searchCriteriaText.text(allData.length === 0 ? 'No data loaded.' : 'Table not initialized.');
            hideLoader();
            return;
        }

        table.search('');
        $.fn.dataTable.ext.search.pop();

        const filters = [];
        $('#multiSearchFields .multi-search-row').each(function() {
            const colIndex = $(this).find('.column-select').val();
            const columnConfig = searchableColumnsConfig.find(c => c.index == colIndex);
            let searchTerm = '';
            if (columnConfig) {
                searchTerm = columnConfig.useDropdown ?
                    $(this).find('.search-value-input').val() :
                    $(this).find('.search-input').val().trim();
                if (searchTerm) {
                    filters.push({
                        col: parseInt(colIndex, 10),
                        term: searchTerm,
                        dataProp: columnConfig.dataProp,
                        isDropdown: columnConfig.useDropdown
                    });
                }
            }
        });

        let criteriaText = operator === 'AND' ? 'Criteria: All filters (AND)' : 'Criteria: Any filter (OR)';
        if (filters.length > 0) {
            criteriaText += ` (${filters.length} active filter(s))`;
            $.fn.dataTable.ext.search.push(
                function(settings, apiData, dataIndex) {
                    if (settings.nTable.id !== table.table().node().id) return true;
                    const rowData = table.row(dataIndex).data();
                    if (!rowData) return false;

                    const logicFn = operator === 'OR' ? filters.some.bind(filters) : filters.every.bind(filters);
                    return logicFn(filter => {
                        let cellDataToTest;
                        if (filter.dataProp === 'Licenses') {
                            return (rowData.Licenses && Array.isArray(rowData.Licenses)) ?
                                rowData.Licenses.some(l => (l.LicenseName || '').toLowerCase() === filter.term.toLowerCase()) : false;
                        } else if (filter.isDropdown) {
                            cellDataToTest = rowData[filter.dataProp] || '';
                            return String(cellDataToTest).toLowerCase() === filter.term.toLowerCase();
                        } else {
                            cellDataToTest = apiData[filter.col] || '';
                            return String(cellDataToTest).toLowerCase().includes(filter.term.toLowerCase());
                        }
                    });
                }
            );
        } else { criteriaText = 'Criteria: All results (no active filters)'; }

        $searchCriteriaText.text(criteriaText);
        table.draw();
    }

    $('#clearFilters').on('click', () => {
        showLoader('Clearing filters...');
        setTimeout(() => {
            if (table) {
                table.search('').columns().search('');
            }
            $('#multiSearchFields .multi-search-row').each(function() {
                $(this).find('.custom-dropdown-text-input').off();
                $(this).find('.custom-options-list').off();
            });
            $('#multiSearchFields').empty();

            if (allData && allData.length > 0) {
                setupMultiSearch();
            } else {
                $('#searchCriteria').text('No data loaded.');
            }
            $.fn.dataTable.ext.search.pop();

            if (table) table.draw(); else hideLoader();

            $('#alertPanel').empty();
            const defaultVisibleCols = [1, 2, 3, 6];
            $('#colContainer .col-vis').each(function() {
                const idx = +$(this).data('col');
                const isDefaultVisible = defaultVisibleCols.includes(idx);
                if (table && idx >= 0 && idx < table.columns().nodes().length) {
                    try { table.column(idx).visible(isDefaultVisible); }
                    catch (e) { console.warn("Error resetting column visibility:", idx, e); }
                }
                $(this).prop('checked', isDefaultVisible);
            });
             if (!table && !(allData && allData.length > 0)) {
                 $('#searchCriteria').text('No data loaded. Please load a JSON file to start.');
             }
        }, 50);
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
            alert("Your browser does not support direct file downloads.");
        }
    }

    $('#exportCsv').on('click', () => {
        if (!table) { alert('Table not initialized. No data loaded.'); return; }
        showLoader('Exporting CSV...');
        setTimeout(() => {
            const rowsToExport = table.rows({ search: 'applied' }).data().toArray();
            if (!rowsToExport.length) {
                hideLoader();
                alert('No records to export with the current filters.');
                return;
            }
            const visibleColumns = [];
            table.columns(':visible').every(function() {
                const columnConfig = table.settings()[0].aoColumns[this.index()];
                const colTitle = $(table.column(this.index()).header()).text() || columnConfig.title;
                const dataProp = columnConfig.mData;
                visibleColumns.push({ title: colTitle, dataProp: dataProp });
            });

            const headerRow = visibleColumns.map(col => escapeCsvValue(col.title)).join(',');
            const csvRows = rowsToExport.map(rowData => {
                return visibleColumns.map(col => {
                    let cellData = rowData[col.dataProp];
                    let shouldForceQuotes = false;
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
        if (!allData.length) { alert('No data loaded to generate the issues report.'); return; }
        showLoader('Generating issues report...');
        setTimeout(() => {
            const lines = [];
            if (nameConflicts.size) {
                lines.push(['NAME+LOCATION CONFLICTS']);
                lines.push(['Name', 'Location'].map(h => escapeCsvValue(h)));
                nameConflicts.forEach(key => lines.push(key.split('|||').map(value => escapeCsvValue(value))));
                lines.push([]);
            }
            if (dupLicUsers.size) {
                lines.push(['USERS with Duplicate Licenses']);
                lines.push(['Name', 'Location', 'Duplicate Licenses', 'Has Paid License?'].map(h => escapeCsvValue(h)));
                allData.filter(user => dupLicUsers.has(user.Id)).forEach(user => {
                    const licCount = {}, duplicateLicNames = [];
                    (user.Licenses || []).forEach(l => licCount[l.LicenseName] = (licCount[l.LicenseName] || 0) + 1);
                    Object.entries(licCount).forEach(([licName, count]) => count > 1 && duplicateLicNames.push(licName));
                    const joinedDups = duplicateLicNames.join('; ');
                    const hasPaid = (user.Licenses || []).some(l => !(l.LicenseName || '').toLowerCase().includes('free'));
                    lines.push([
                        escapeCsvValue(user.DisplayName), escapeCsvValue(user.OfficeLocation),
                        escapeCsvValue(joinedDups, true),
                        escapeCsvValue(hasPaid ? 'Yes' : 'No')
                    ]);
                });
            }
            if (!lines.length) { lines.push(['No issues detected.']); }
            const csvContent = lines.map(rowArray => rowArray.join(',')).join('\n');
            downloadCsv(csvContent, 'issues_report.csv');
            hideLoader();
        }, 50);
    });

    function processDataWithWorker(rawData) {
        showLoader('Validating and processing data (this may take a moment)...');
        const workerScript = `
            const nameKeyInternal = u => \`\${u.DisplayName || ''}|||\${u.OfficeLocation || ''}\`;
            function validateJsonForWorker(data) {
                if (!Array.isArray(data)) {
                    return { error: 'Invalid JSON: Must be an array of objects.', validatedData: [] };
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
                        const baseName = (l.LicenseName || '');
                        if (baseName) { licCount.set(baseName, (licCount.get(baseName) || 0) + 1); }
                    });
                    if ([...licCount].some(([lic, c]) => officeLic.has(lic) && c > 1)) { dupSet.add(u.Id); }
                });
                const conflictingNameKeysArray = [...nameMap].filter(([, count]) => count > 1).map(([key]) => key);
                return { nameConflictsArray: conflictingNameKeysArray, dupLicUsersArray: Array.from(dupSet) };
            }
            function calculateUniqueFieldValuesForWorker(data, config) {
                const localUniqueFieldValues = {};
                config.forEach(colConfig => {
                    if (colConfig.useDropdown) {
                        if (colConfig.dataProp === 'Licenses') {
                            const allLicenseObjects = data.flatMap(user => user.Licenses || []).filter(l => l.LicenseName);
                            localUniqueFieldValues.Licenses = [...new Set(allLicenseObjects.map(l => l.LicenseName))].sort((a, b) => a.toLowerCase().localeCompare(b.toLowerCase()));
                        } else {
                            localUniqueFieldValues[colConfig.dataProp] = [...new Set(data.map(user => user[colConfig.dataProp]).filter(value => value && String(value).trim() !== ''))].sort((a, b) => String(a).toLowerCase().localeCompare(String(b).toLowerCase()));
                        }
                    }
                });
                return { uniqueFieldValues: localUniqueFieldValues };
            }
            self.onmessage = function(e) {
                const { rawData, searchableColumnsConfig: workerSearchableColumnsConfig } = e.data;
                try {
                    const validationResult = validateJsonForWorker(rawData);
                    if (validationResult.error) {
                        self.postMessage({ error: validationResult.error });
                        return;
                    }
                    const validatedData = validationResult.validatedData;
                    if (validatedData.length === 0) {
                         self.postMessage({ validatedData: [], nameConflictsArray: [], dupLicUsersArray: [], uniqueFieldValues: {} });
                         return;
                    }
                    const issues = findIssuesForWorker(validatedData);
                    const uniqueValues = calculateUniqueFieldValuesForWorker(validatedData, workerSearchableColumnsConfig);
                    self.postMessage({
                        validatedData: validatedData,
                        nameConflictsArray: issues.nameConflictsArray,
                        dupLicUsersArray: issues.dupLicUsersArray,
                        uniqueFieldValues: uniqueValues.uniqueFieldValues,
                        error: null
                    });
                } catch (err) {
                    self.postMessage({ error: 'Error in Web Worker: ' + err.message + '\\\\n' + err.stack });
                } finally {
                    self.close();
                }
            };
        `;
        const blob = new Blob([workerScript], { type: 'application/javascript' });
        const worker = new Worker(URL.createObjectURL(blob));

        worker.onmessage = function(e) {
            URL.revokeObjectURL(blob);
            const { validatedData: processedData, nameConflictsArray, dupLicUsersArray, uniqueFieldValues: uFValues, error } = e.data;

            if (error) {
                hideLoader();
                alert('Error processing data in Worker: ' + error);
                console.error("Worker Error:", error);
                $('#searchCriteria').text('Error loading data.');
                return;
            }

            nameConflicts = new Set(nameConflictsArray);
            dupLicUsers = new Set(dupLicUsersArray);
            uniqueFieldValues = uFValues;

            if (processedData && processedData.length > 0) {
                showLoader('Rendering table...');
                setTimeout(() => {
                    initTable(processedData); // initTable também chamará updateAiFeatureStatus()
                    setupMultiSearch();
                    $('#searchCriteria').text(`Data loaded (${processedData.length} users). Use filters to refine.`);
                }, 50);
            } else {
                hideLoader();
                alert('No valid users found in the data. Please check the JSON file or PowerShell script output.');
                $('#searchCriteria').text('No valid data loaded.');
                if (table) { table.clear().draw(); }
                $('#multiSearchFields').empty();
                $('#alertPanel').empty();
                updateAiFeatureStatus(); // Atualiza status mesmo sem dados
            }
        };

        worker.onerror = function(e) {
            URL.revokeObjectURL(blob);
            hideLoader();
            console.error(`Error in Web Worker: Line ${e.lineno} in ${e.filename}: ${e.message}`);
            alert('A critical error occurred during data processing. Please check the console.');
            $('#searchCriteria').text('Critical error loading data.');
            updateAiFeatureStatus(); // Atualiza status em caso de erro crítico
        };
        worker.postMessage({ rawData: rawData, searchableColumnsConfig: JSON.parse(JSON.stringify(searchableColumnsConfig)) });
    }

    try {
        if (typeof userData !== 'undefined' && ( (Array.isArray(userData) && userData.length > 0) || (userData.data && Array.isArray(userData.data) && userData.data.length > 0) ) ) {
            let dataToProcess = Array.isArray(userData) ? userData : userData.data;
            if (userData.error || (Array.isArray(dataToProcess) && dataToProcess.length === 0 && !userData.message) ) {
                 $('#searchCriteria').text(userData.error || 'JSON data is empty or invalid as provided by PowerShell.');
                 console.error("Error in userData from PowerShell or empty data:", userData);
                 hideLoader();
                 updateAiFeatureStatus();
            } else if (userData.message && Array.isArray(dataToProcess) && dataToProcess.length === 0) {
                 $('#searchCriteria').text(userData.message);
                 hideLoader();
                 updateAiFeatureStatus();
            }
            else {
                processDataWithWorker(dataToProcess);
            }
        } else if (typeof userData !== 'undefined' && userData.message) {
             $('#searchCriteria').text(userData.message);
             hideLoader();
             updateAiFeatureStatus();
        }
        else {
            console.warn("Global variable 'userData' is not defined as expected or indicates an error state from PowerShell.");
            $('#searchCriteria').html('Failed to load data from PowerShell or data is empty. Check PowerShell script output. <br/> (If running HTML directly, <code>userData</code> is not defined).');
            hideLoader();
            updateAiFeatureStatus();
        }
    } catch (error) {
        hideLoader();
        alert('Error initiating data loading: ' + error.message);
        console.error("Initial loading error:", error);
        $('#searchCriteria').text('Error loading data.');
        updateAiFeatureStatus();
    }

   // ===========================================================
    // FUNCIONALIDADE DE IA - (Pré-processamento dinâmico por pergunta e controle de tokens)
    // ===========================================================
    const AI_API_KEY_STORAGE_KEY = 'licenseAiApiKey_v1_google_prompt_efficient';
    const $manageAiApiKeyButton = $('#manageAiApiKeyButton');
    const $askAiButton = $('#askAiButton');
    const $aiQuestionInput = $('#aiQuestionInput');
    const $aiResponseArea = $('#aiResponseArea');
    const $aiLoadingIndicator = $('#aiLoadingIndicator');
    const $aiTokenInfo = $('#aiTokenInfo'); // Novo elemento para mostrar info de tokens

    let aiConversationHistory = [];
    const MAX_CONVERSATION_HISTORY_TURNS = 3; // Reduzido para economizar tokens, manter últimos 3 pares + system
    const TARGET_MAX_CHARS_FOR_API = 750000; // Alvo de ~750k caracteres (aproximadamente 180k-200k tokens, bem abaixo de 800k)
                                           // Ajuste conforme necessário. 800k tokens é muito, talvez seja 800k caracteres?
                                           // Gemini 1.5 Flash tem 1M de tokens de contexto.
                                           // Um token é ~4 caracteres em inglês. Para português, pode ser um pouco mais.
                                           // 800.000 tokens * ~3.5 chars/token (PT-BR) = ~2.800.000 caracteres.
                                           // Vou usar TARGET_MAX_CHARS_FOR_API como um limite prático para o CONTEÚDO gerado.

    function getAiApiKey() {
        return localStorage.getItem(AI_API_KEY_STORAGE_KEY);
    }

    function updateAiApiKeyStatusDisplay() {
        if (getAiApiKey()) {
            $manageAiApiKeyButton.text('API IA Config.');
            $manageAiApiKeyButton.css('background-color', '#28a745');
            $manageAiApiKeyButton.attr('title', 'Chave da API da IA (Google Gemini) configurada. Clique para alterar ou remover.');
        } else {
            $manageAiApiKeyButton.text('Configurar API IA');
            $manageAiApiKeyButton.css('background-color', ''); // Volta para o estilo padrão
            $manageAiApiKeyButton.attr('title', 'Configurar chave da API da IA (Google Gemini) para análise.');
        }
        updateAiFeatureStatus(); // Atualiza o botão de perguntar à IA
    }

    function handleApiKeyInputViaPrompt() {
        const currentKey = getAiApiKey() || "";
        const newKey = window.prompt("Por favor, insira sua chave de API do Google Gemini:", currentKey);

        if (newKey === null) { // Usuário cancelou
            // Nenhuma mudança, apenas atualiza a exibição
        } else if (newKey.trim() === "" && currentKey !== "") {
            if (confirm("Você tem certeza que deseja remover a chave da API salva?")) {
                localStorage.removeItem(AI_API_KEY_STORAGE_KEY);
                alert("Chave da API removida.");
            }
        } else if (newKey.trim() !== "") {
            localStorage.setItem(AI_API_KEY_STORAGE_KEY, newKey.trim());
            alert("Chave da API da IA (Google Gemini) salva!");
        }
        updateAiApiKeyStatusDisplay();
        return getAiApiKey() !== null && getAiApiKey().trim() !== "";
    }

    $manageAiApiKeyButton.on('click', function() {
        handleApiKeyInputViaPrompt();
    });

    // Função para estimar tokens (muito simplificada, baseada em caracteres)
    function estimateChars(text) {
        return text.length;
    }

    // Atualiza o estado do botão Ask AI e a informação de tokens
    async function updateAiFeatureStatus() {
        const apiKey = getAiApiKey();
        const hasData = allData && allData.length > 0;

        if (!apiKey || !hasData) {
            $askAiButton.prop('disabled', true);
            $askAiButton.attr('title', !apiKey ? 'Configure a chave da API primeiro.' : 'Carregue os dados primeiro.');
            $aiTokenInfo.text('');
            return;
        }

        // Simula a preparação dos dados para verificar o tamanho ANTES de realmente perguntar
        // Esta é uma verificação leve, a verificação completa ocorre no $askAiButton.on('click')
        const estimatedContext = await prepareContextForAI("", false); // false = não é para envio real
        const totalEstimatedChars = estimateChars(estimatedContext.context) +
                                   estimateChars(JSON.stringify(aiConversationHistory));

        if (totalEstimatedChars > TARGET_MAX_CHARS_FOR_API) {
            $askAiButton.prop('disabled', true);
            $askAiButton.attr('title', `Os dados preparados excedem o limite de caracteres (${Math.round(totalEstimatedChars/1000)}k > ${Math.round(TARGET_MAX_CHARS_FOR_API/1000)}k). Tente uma pergunta mais específica ou menos histórico.`);
            $aiTokenInfo.text(`Info: Dados (${Math.round(totalEstimatedChars/1000)}k chars) podem exceder o limite.`).css('color', 'orange');
        } else {
            $askAiButton.prop('disabled', false);
            $askAiButton.attr('title', 'Perguntar à IA');
            $aiTokenInfo.text(`Pronto para IA. Dados estimados: ${Math.round(totalEstimatedChars/1000)}k / ${Math.round(TARGET_MAX_CHARS_FOR_API/1000)}k caracteres.`).css('color', '');
        }
    }


    function extractMentionedKeywords(userQuestionLower, availableValues, keywordsToTrigger) {
        const mentionedItems = [];
        if (availableValues && keywordsToTrigger.some(kw => userQuestionLower.includes(kw))) {
            availableValues.forEach(val => {
                if (userQuestionLower.includes(String(val).toLowerCase())) {
                    mentionedItems.push(val);
                }
            });
        }
        return [...new Set(mentionedItems)];
    }

    async function prepareContextForAI(userQuestion, isActualSendOperation) {
        const userQuestionLower = userQuestion.toLowerCase();
        let contextForThisQuestion = "";
        let sampleDataForPrompt = [];
        let contextNoteForSample = "";

        // Estimativa de caracteres para o histórico e a pergunta do usuário (aproximado)
        let currentChars = estimateChars(JSON.stringify(aiConversationHistory)) + estimateChars(userQuestion);
        const MAX_SAMPLE_USERS_FOR_QUESTION = 10; // Amostra pode ser maior se o resto couber
        const MAX_CHARS_FOR_USER_SAMPLE = TARGET_MAX_CHARS_FOR_API * 0.3; // 30% do total para amostra de usuários
        const MAX_CHARS_FOR_AGGREGATES = TARGET_MAX_CHARS_FOR_API * 0.4; // 40% para agregados

        let systemMessage = `Você é um assistente especialista em análise de licenciamento Microsoft 365.
Sua tarefa é responder à pergunta do usuário com base nos dados fornecidos.
Os dados podem consistir em:
1. Estatísticas Agregadas Gerais: Um resumo do uso de licenças.
2. Amostra Contextual de Dados de Usuários: Se a pergunta for específica, uma amostra de usuários relevantes PODE ser fornecida.

Instruções:
- Baseie-se ESTRITAMENTE nos dados fornecidos.
- Se a pergunta exigir dados de nível de usuário que não estão na amostra, INDIQUE CLARAMENTE que você pode responder com base nas estatísticas e na amostra limitada, mas uma análise completa de todos os usuários para aquela pergunta específica não é possível com os dados fornecidos. NÃO invente dados de usuários.
- Seja conciso. Use Markdown simples. Não inclua saudações/despedidas.`;
        currentChars += estimateChars(systemMessage);

        if (allData && allData.length > 0) {
            const totalUsers = allData.length;
            const licenseMap = new Map();
            allData.forEach(user => (user.Licenses || []).forEach(lic => licenseMap.set(lic.LicenseName, (licenseMap.get(lic.LicenseName) || 0) + 1)));
            
            let aggregatedStats = `Estatísticas Agregadas (Total de ${totalUsers} usuários no dataset):\n`;
            aggregatedStats += `- Licenças Únicas Distribuídas: ${licenseMap.size}\n`;
            
            const topLicensesOverall = [...licenseMap.entries()].sort(([,a],[,b]) => b-a).slice(0,10).map(([n,c])=>`  - ${n}: ${c} usuários`).join('\n');
            aggregatedStats += `- Top 10 Licenças (Geral):\n${topLicensesOverall}\n`;

            let usersWithMultipleHighValueLicenses = 0;
            const highValueLicenseKeywords = ['E5', 'Premium', 'Copilot', 'P2', 'Enterprise']; // Adicionado 'Enterprise'
            const departmentLicenseCounts = {};

            allData.forEach(user => {
                let highValueLicenseCount = 0;
                const location = user.OfficeLocation || "Não Especificado";
                if (!departmentLicenseCounts[location]) { departmentLicenseCounts[location] = { count: 0, licenses: new Map() }; }
                departmentLicenseCounts[location].count++;

                (user.Licenses || []).forEach(lic => {
                    const licName = lic.LicenseName;
                    const currentDeptLicMap = departmentLicenseCounts[location].licenses;
                    currentDeptLicMap.set(licName, (currentDeptLicMap.get(licName) || 0) + 1);
                    if (highValueLicenseKeywords.some(keyword => licName.toLowerCase().includes(keyword.toLowerCase()))) {
                        highValueLicenseCount++;
                    }
                });
                if (highValueLicenseCount > 1) usersWithMultipleHighValueLicenses++;
            });
            aggregatedStats += `- Usuários com Múltiplas Licenças de Alto Valor: ${usersWithMultipleHighValueLicenses}\n`;

            let departmentSummary = "\nDistribuição de Licenças por Localização/Departamento (Top 3 locais com mais usuários):\n";
            const sortedDepartments = Object.entries(departmentLicenseCounts)
                .map(([dept, data]) => ({ name: dept, userCount: data.count, licMap: data.licenses }))
                .sort((a,b) => b.userCount - a.userCount).slice(0, 3);

            sortedDepartments.forEach(deptInfo => {
                departmentSummary += `  Local/Depto: ${deptInfo.name} (${deptInfo.userCount} usuários)\n`;
                const topDeptLicenses = [...deptInfo.licMap.entries()].sort(([,a],[,b]) => b-a).slice(0,3).map(([name, count]) => `    - ${name}: ${count}`).join('\n');
                departmentSummary += `${topDeptLicenses}\n`;
            });
            aggregatedStats += departmentSummary;
            
            // Truncar estatísticas agregadas se necessário
            if (estimateChars(aggregatedStats) > MAX_CHARS_FOR_AGGREGATES) {
                aggregatedStats = aggregatedStats.substring(0, MAX_CHARS_FOR_AGGREGATES) + "\n...(estatísticas agregadas truncadas)...";
            }
            contextForThisQuestion += aggregatedStats;
            currentChars += estimateChars(contextForThisQuestion);

            // Lógica para Amostragem Contextual Baseada na Pergunta
            const detailKeywords = ["quais usuários", "liste usuários", "usuários com", "quem tem", "usuário específico", "detalhes de"];
            const isDetailedUserQuestion = userQuestion && detailKeywords.some(kw => userQuestionLower.includes(kw));

            if (isDetailedUserQuestion) {
                let filteredSamplePool = [...allData]; // Começa com todos os dados
                const mentionedLics = extractMentionedKeywords(userQuestionLower, uniqueFieldValues.Licenses, ["licença", "sku"]);
                const mentionedLocs = extractMentionedKeywords(userQuestionLower, uniqueFieldValues.OfficeLocation, ["local", "departamento", "escritório"]);
                const mentionedJobTitles = extractMentionedKeywords(userQuestionLower, uniqueFieldValues.JobTitle, ["cargo", "função"]);

                if (mentionedLics.length > 0) {
                    filteredSamplePool = filteredSamplePool.filter(u =>
                        (u.Licenses || []).some(l => mentionedLics.includes(l.LicenseName))
                    );
                }
                if (mentionedLocs.length > 0) {
                    filteredSamplePool = filteredSamplePool.filter(u =>
                        mentionedLocs.includes(u.OfficeLocation)
                    );
                }
                if (mentionedJobTitles.length > 0) {
                     filteredSamplePool = filteredSamplePool.filter(u =>
                        mentionedJobTitles.includes(u.JobTitle)
                    );
                }
                // Se nenhum filtro específico foi aplicado pela pergunta mas é uma pergunta detalhada,
                // pegue uma amostra aleatória de todo o dataset como fallback.
                if (filteredSamplePool.length === allData.length && isDetailedUserQuestion && mentionedLics.length === 0 && mentionedLocs.length === 0 && mentionedJobTitles.length === 0) {
                    // Amostra aleatória geral se nenhum critério específico foi encontrado
                     filteredSamplePool = [...allData].sort(() => 0.5 - Math.random());
                }


                if (filteredSamplePool.length > 0) {
                    let tempSample = [];
                    let sampleChars = 0;
                    for (const user of filteredSamplePool) {
                        if (tempSample.length >= MAX_SAMPLE_USERS_FOR_QUESTION) break;
                        const userSample = {
                            DisplayName: user.DisplayName,
                            OfficeLocation: user.OfficeLocation,
                            JobTitle: user.JobTitle,
                            Licenses: (user.Licenses || []).map(l => l.LicenseName)
                        };
                        const userSampleStr = JSON.stringify(userSample);
                        if (currentChars + sampleChars + estimateChars(userSampleStr) < TARGET_MAX_CHARS_FOR_API && sampleChars + estimateChars(userSampleStr) < MAX_CHARS_FOR_USER_SAMPLE) {
                            tempSample.push(userSample);
                            sampleChars += estimateChars(userSampleStr);
                        } else {
                            break; // Para de adicionar se exceder o limite de caracteres para amostra ou total
                        }
                    }
                    sampleDataForPrompt = tempSample;
                }

                if (sampleDataForPrompt.length > 0) {
                    contextNoteForSample = `\nAmostra Contextual de Usuários (${sampleDataForPrompt.length} usuário(s) relevantes para a pergunta):\n${JSON.stringify(sampleDataForPrompt, null, 2)}\n`;
                } else {
                    contextNoteForSample = "\nNão foi possível gerar uma amostra de usuários relevante para os detalhes específicos da pergunta dentro dos limites de dados, ou nenhum usuário correspondeu. A resposta será baseada principalmente nas estatísticas agregadas.\n";
                }
                // Adiciona nota da amostra apenas se houver espaço
                if (currentChars + estimateChars(contextNoteForSample) < TARGET_MAX_CHARS_FOR_API) {
                    contextForThisQuestion += contextNoteForSample;
                    currentChars += estimateChars(contextNoteForSample);
                } else {
                     contextForThisQuestion += "\n...(nota sobre amostra de usuários omitida por limite de dados)...";
                }
            } else {
                const generalNote = "\nNenhuma amostra detalhada de usuário foi incluída para esta pergunta geral. Responda com base nas estatísticas agregadas.\n";
                 if (currentChars + estimateChars(generalNote) < TARGET_MAX_CHARS_FOR_API){
                    contextForThisQuestion += generalNote;
                    currentChars += estimateChars(generalNote);
                 }
            }
        } else {
            contextForThisQuestion = 'Não há dados de licença carregados para analisar.';
            if (isActualSendOperation) { // Só mostra alerta se for um envio real
                 $aiResponseArea.text(contextForThisQuestion);
            }
            return { context: contextForThisQuestion, systemMessage: systemMessage, history: aiConversationHistory, canSend: false };
        }
        
        const canSendData = currentChars < TARGET_MAX_CHARS_FOR_API;
        if (isActualSendOperation) {
            $aiTokenInfo.text(`Enviando: ~${Math.round(currentChars/1000)}k / ${Math.round(TARGET_MAX_CHARS_FOR_API/1000)}k caracteres.`).css('color', canSendData ? 'green' : 'red');
        }

        return {
            context: contextForThisQuestion,
            systemMessage: systemMessage,
            history: aiConversationHistory,
            canSend: canSendData
        };
    }


    $askAiButton.on('click', async function() {
        let apiKey = getAiApiKey();
        if (!apiKey) {
            alert('Por favor, configure sua chave de API do Google Gemini clicando em "Configurar API IA".');
            if (!handleApiKeyInputViaPrompt()) return;
            apiKey = getAiApiKey();
            if (!apiKey) return;
        }

        const userQuestion = $aiQuestionInput.val().trim();
        if (!userQuestion) {
            alert("Por favor, digite sua pergunta para a IA.");
            return;
        }

        if (!allData || allData.length === 0) {
            $aiResponseArea.text('Não há dados de licença carregados para analisar.');
            return;
        }

        $aiLoadingIndicator.removeClass('hidden');
        $askAiButton.prop('disabled', true);
        $aiResponseArea.html(`Preparando dados e analisando com IA (Google Gemini)... <i class="fas fa-spinner fa-spin"></i>`);

        const preparedPayload = await prepareContextForAI(userQuestion, true); // true = para envio real

        if (!preparedPayload.canSend) {
            $aiResponseArea.text('Os dados a serem enviados para a IA excedem o limite de tamanho configurado. Por favor, tente uma pergunta mais específica, reduza o histórico de chat (se aplicável limpando a página e começando uma nova conversa) ou revise os filtros de dados.');
            $aiLoadingIndicator.addClass('hidden');
            $askAiButton.prop('disabled', false); // Reabilita para tentar de novo
            updateAiFeatureStatus(); // Atualiza o status e o botão
            return;
        }

        const GEMINI_MODEL_TO_USE = "gemini-1.5-flash-latest";
        const AI_PROVIDER_ENDPOINT_BASE = `https://generativelanguage.googleapis.com/v1beta/models/${GEMINI_MODEL_TO_USE}:generateContent`;
        const finalEndpoint = `${AI_PROVIDER_ENDPOINT_BASE}?key=${apiKey}`;

        try {
            const requestHeaders = { 'Content-Type': 'application/json', };
            
            let conversationTurnsForAPI = [];
            if (aiConversationHistory.length === 0) { // Primeira pergunta da sessão de chat
                conversationTurnsForAPI.push({ "role": "user", "parts": [{ "text": preparedPayload.systemMessage }] });
                conversationTurnsForAPI.push({ "role": "model", "parts": [{ "text": "Entendido. Estou pronto para ajudar com a análise de licenciamento, respeitando os dados fornecidos."}] });
            } else {
                conversationTurnsForAPI = [...aiConversationHistory];
            }
            
            const userTurn = { "role": "user", "parts": [{ "text": `${preparedPayload.context}\n\nPERGUNTA DO USUÁRIO: ${userQuestion}` }] };
            conversationTurnsForAPI.push(userTurn);
            
            const requestBody = {
              contents: conversationTurnsForAPI,
              generationConfig: { "maxOutputTokens": 2048, "temperature": 0.4 } // Max output pode ser ajustado
            };

            // console.log(`Enviando para Google Gemini. Turnos no histórico (incluindo system prompt implícito): ${conversationTurnsForAPI.length}`);
            // console.log("Corpo da Requisição (semelhante):", JSON.stringify(requestBody.contents.map(c => ({role: c.role, textLength: c.parts[0].text.length})), null, 2));
            // console.log("Contexto enviado para esta pergunta (aprox.):", preparedPayload.context.length, "caracteres.");


            const response = await fetch(finalEndpoint, {
                method: 'POST',
                headers: requestHeaders,
                body: JSON.stringify(requestBody)
            });
            
            const responseBodyText = await response.text();
            let data;
            try { data = JSON.parse(responseBodyText); }
            catch (e) {
                console.error("Erro ao fazer parse da resposta JSON da IA:", e);
                console.error("Resposta recebida (texto):", responseBodyText);
                $aiResponseArea.text(`Erro ao processar resposta da IA. Detalhes no console. Status: ${response.status}`);
                throw new Error(`JSON Parse Error: ${e.message}. Response Status: ${response.status}.`);
            }

            if (!response.ok) {
                console.error("Erro da API Google Gemini:", data);
                let errorMsg = `Erro da API da IA: ${response.status}. `;
                if (data && data.error && data.error.message) {
                    errorMsg += data.error.message;
                } else {
                    errorMsg += "Detalhes não disponíveis.";
                }
                $aiResponseArea.text(errorMsg);
                throw new Error(errorMsg);
            }

            // console.log("Resposta da API Google Gemini:", data);

            let aiTextResponse = "Não foi possível extrair a resposta da IA ou a resposta estava vazia.";
             if (data.candidates && data.candidates[0]) {
                const candidate = data.candidates[0];
                if (candidate.content && candidate.content.parts && candidate.content.parts.length > 0 && candidate.content.parts[0].text) {
                    aiTextResponse = candidate.content.parts.map(part => part.text).join("").trim();
                } else if (candidate.finishReason && candidate.finishReason !== "STOP") {
                    aiTextResponse = `A IA parou de gerar a resposta devido a: ${candidate.finishReason}.`;
                    if (candidate.safetyRatings) {
                         aiTextResponse += ` Avaliações de segurança: ${JSON.stringify(candidate.safetyRatings)}`;
                    }
                }
            }
            
            // Adicionar pergunta atual (com contexto) e resposta da IA ao histórico
            // Adiciona o turno do usuário que acabou de ser enviado
            aiConversationHistory.push(userTurn);
            // Adiciona resposta da IA
            aiConversationHistory.push({ "role": "model", "parts": [{ "text": aiTextResponse }] });

            // Truncar histórico para manter os últimos N turnos (1 turno = 1 user + 1 model)
            // +2 para o system prompt inicial e resposta da IA simulada
            const maxHistoryItems = (MAX_CONVERSATION_HISTORY_TURNS * 2) + (aiConversationHistory.find(turn => turn.parts[0].text.startsWith("Você é um assistente especialista")) ? 2 : 0);
            if (aiConversationHistory.length > maxHistoryItems) {
                 // Se o system prompt original estiver lá, preserva-o e a primeira resposta do modelo.
                if (aiConversationHistory[0].parts[0].text.startsWith("Você é um assistente especialista")) {
                    const systemPromptAndFirstModelResponse = aiConversationHistory.slice(0, 2);
                    const recentTurns = aiConversationHistory.slice(-( (MAX_CONVERSATION_HISTORY_TURNS -1) * 2 )); // -1 porque já temos um par (system)
                    aiConversationHistory = [...systemPromptAndFirstModelResponse, ...recentTurns];
                } else { // Caso contrário, apenas pega os últimos N pares.
                    aiConversationHistory = aiConversationHistory.slice(-(MAX_CONVERSATION_HISTORY_TURNS * 2));
                }
                // console.log(`Histórico truncado. Mantendo ${aiConversationHistory.length} itens.`);
            }
            
            $aiQuestionInput.val('');
            function simpleMarkdownToHtml(mdText) {
                if (typeof mdText !== 'string') return '';
                // Negrito: **texto** ou __texto__
                mdText = mdText.replace(/\*\*(.*?)\*\*|__(.*?)__/g, '<strong>$1$2</strong>');
                // Itálico: *texto* ou _texto_
                mdText = mdText.replace(/\*(.*?)\*|_(.*?)_/g, '<em>$1$2</em>');
                // Listas não ordenadas: - item ou * item
                mdText = mdText.replace(/^(\s*)- (.*)/gm, '$1<li>$2</li>');
                mdText = mdText.replace(/^(\s*)\* (.*)/gm, '$1<li>$2</li>');
                mdText = mdText.replace(/(\<\/li\>\n)+<li>/g, '</li><li>'); // Agrupa <li> adjacentes
                mdText = mdText.replace(/(<li>.*<\/li>)/gs, '<ul>$1</ul>');
                mdText = mdText.replace(/<\/ul>\s*<ul>/gs, ''); // Remove <ul> aninhados desnecessariamente

                // Títulos (simples): # Título, ## Título
                mdText = mdText.replace(/^# (.*$)/gim, '<h1>$1</h1>');
                mdText = mdText.replace(/^## (.*$)/gim, '<h2>$1</h2>');
                mdText = mdText.replace(/^### (.*$)/gim, '<h3>$1</h3>');
                // Quebras de linha
                mdText = mdText.replace(/\n/g, '<br>');
                // Remover <br> dentro de <ul> e antes/depois de <li> (efeito colateral da conversão de lista)
                mdText = mdText.replace(/<ul><br>/g, '<ul>').replace(/<br><\/ul>/g, '</ul>');
                mdText = mdText.replace(/<li><br>/g, '<li>').replace(/<br><\/li>/g, '</li>');
                return mdText;
            }
            $aiResponseArea.html(simpleMarkdownToHtml(aiTextResponse));

        } catch (error) {
            console.error("Erro na interação com a IA:", error);
            if (!$aiResponseArea.text().startsWith("Erro da API da IA:") && !$aiResponseArea.text().startsWith("Erro ao processar resposta da IA.")) {
                 $aiResponseArea.text('Ocorreu um erro ao tentar comunicar com a IA. Verifique o console para mais detalhes.');
            }
        }
        finally {
            $aiLoadingIndicator.addClass('hidden');
            $askAiButton.prop('disabled', false); // Reabilita mesmo em erro para nova tentativa
            updateAiFeatureStatus(); // Atualiza o status e o botão
        }
    });
   
    updateAiApiKeyStatusDisplay(); // Chamada inicial para configurar o botão da API
    // ===========================================================
    // FUNCIONALIDADE DE IA - FIM
    // ===========================================================

}); // Fim do $(document).ready()
