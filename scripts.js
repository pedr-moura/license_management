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
                    initTable(processedData);
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
            }
        };

        worker.onerror = function(e) {
            URL.revokeObjectURL(blob);
            hideLoader();
            console.error(`Error in Web Worker: Line ${e.lineno} in ${e.filename}: ${e.message}`);
            alert('A critical error occurred during data processing. Please check the console.');
            $('#searchCriteria').text('Critical error loading data.');
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
            } else if (userData.message && Array.isArray(dataToProcess) && dataToProcess.length === 0) {
                 $('#searchCriteria').text(userData.message);
                 hideLoader();
            }
            else {
                processDataWithWorker(dataToProcess);
            }
        } else if (typeof userData !== 'undefined' && userData.message) {
             $('#searchCriteria').text(userData.message);
             hideLoader();
        }
        else {
            console.warn("Global variable 'userData' is not defined as expected or indicates an error state from PowerShell.");
            $('#searchCriteria').html('Failed to load data from PowerShell or data is empty. Check PowerShell script output. <br/> (If running HTML directly, <code>userData</code> is not defined).');
            hideLoader();
        }
    } catch (error) {
        hideLoader();
        alert('Error initiating data loading: ' + error.message);
        console.error("Initial loading error:", error);
        $('#searchCriteria').text('Error loading data.');
    }

   // ===========================================================
    // FUNCIONALIDADE DE IA - (Input da API Key via prompt())
    // ===========================================================
    const AI_API_KEY_STORAGE_KEY = 'licenseAiApiKey_v1_google_prompt'; // Nova chave para localStorage

    const $manageAiApiKeyButton = $('#manageAiApiKeyButton');
    // Variáveis do modal HTML removidas: $aiApiKeyModal, $closeAiModalButton, $aiApiKeyInput, $saveAiApiKeyButton, $aiApiKeyMessage

    const $startAiProcessingButton = $('#startAiProcessingButton');
    const $aiProcessingStatus = $('#aiProcessingStatus');
    const $aiInteractionArea = $('#aiInteractionArea');

    const $askAiButton = $('#askAiButton');
    const $aiQuestionInput = $('#aiQuestionInput');
    const $aiResponseArea = $('#aiResponseArea');
    const $aiLoadingIndicator = $('#aiLoadingIndicator');

    let aiConversationHistory = []; 
    let blockSummaries = [];
    let totalChunks = 0;
    let chunksProcessed = 0;

    const USERS_PER_CHUNK = 2500; 
    const MAX_TOKENS_PER_CHUNK_ANALYSIS = 2048;
    const MAX_TOKENS_FINAL_QUESTION = 2048;


    // --- Gerenciamento da Chave API da IA via prompt() ---
    function getAiApiKey() {
        return localStorage.getItem(AI_API_KEY_STORAGE_KEY);
    }

    function updateAiApiKeyStatusDisplay() {
        if (getAiApiKey()) {
            $manageAiApiKeyButton.text('API IA Config.');
            $manageAiApiKeyButton.css('background-color', '#28a745'); // Verde
            $manageAiApiKeyButton.attr('title', 'Chave da API da IA (Google Gemini) configurada. Clique para alterar.');
        } else {
            $manageAiApiKeyButton.text('Configurar API IA');
            $manageAiApiKeyButton.css('background-color', ''); // Volta à cor original
            $manageAiApiKeyButton.attr('title', 'Configurar chave da API da IA (Google Gemini) para análise.');
        }
    }
   
    function handleApiKeyInputViaPrompt() {
        const currentKey = getAiApiKey() || "";
        const newKey = window.prompt("Por favor, insira sua chave de API do Google Gemini (deixe em branco e cancele para não alterar):", currentKey);

        if (newKey === null) { // Usuário clicou em Cancelar
            alert("Configuração da chave da API cancelada.");
            updateAiApiKeyStatusDisplay(); // Garante que o status do botão reflita a chave existente (ou nenhuma)
            return false; // Indica que a configuração foi cancelada ou falhou
        } else if (newKey.trim() === "" && currentKey !== "") { // Usuário apagou a chave e clicou OK
            if (confirm("Você tem certeza que deseja remover a chave da API salva?")) {
                localStorage.removeItem(AI_API_KEY_STORAGE_KEY);
                alert("Chave da API removida.");
                updateAiApiKeyStatusDisplay();
            } else {
                 alert("Remoção da chave cancelada. A chave anterior foi mantida.");
            }
            return false;
        } else if (newKey.trim() !== "") { // Usuário inseriu uma nova chave
            localStorage.setItem(AI_API_KEY_STORAGE_KEY, newKey.trim());
            alert("Chave da API da IA (Google Gemini) salva!");
            updateAiApiKeyStatusDisplay();
            return true; // Indica que a chave foi salva com sucesso
        } else { // Usuário deixou em branco e clicou OK, sem chave anterior
            alert("Nenhuma chave inserida. A configuração permanece inalterada.");
            updateAiApiKeyStatusDisplay();
            return false;
        }
    }

    $manageAiApiKeyButton.on('click', function() {
        handleApiKeyInputViaPrompt();
    });


    // --- Lógica de Processamento em Blocos e Interação com a IA ---
    $startAiProcessingButton.on('click', async function() {
        let apiKey = getAiApiKey();
        if (!apiKey) {
            alert('Por favor, configure sua chave de API do Google Gemini primeiro clicando no botão "Configurar API IA".');
            // Opcionalmente, chamar handleApiKeyInputViaPrompt() diretamente:
            // if (!handleApiKeyInputViaPrompt()) return; // Se o usuário cancelar, para aqui
            // apiKey = getAiApiKey(); // Pega a chave recém-inserida
            // if (!apiKey) return; // Para se ainda não houver chave
            return;
        }
        if (!allData || allData.length === 0) {
            $aiProcessingStatus.text('Não há dados de licença carregados para processar.');
            return;
        }

        // ... (resto da lógica do $startAiProcessingButton.on('click') como na versão anterior)
        // ... (calcula chunks, itera, chama a API para cada chunk, armazena blockSummaries)
        // ... (habilita a área de interação no final)
        $(this).prop('disabled', true).html('<i class="fas fa-spinner fa-spin"></i> Processando Blocos...');
        $aiInteractionArea.addClass('hidden'); 
        aiConversationHistory = []; 
        blockSummaries = [];
        chunksProcessed = 0;
        totalChunks = Math.ceil(allData.length / USERS_PER_CHUNK);
        $aiProcessingStatus.text(`Iniciando processamento de ${totalChunks} blocos de dados...`);

        const GEMINI_MODEL_TO_USE = "gemini-1.5-flash-latest"; 
        const AI_PROVIDER_ENDPOINT_BASE = `https://generativelanguage.googleapis.com/v1beta/models/${GEMINI_MODEL_TO_USE}:generateContent`;
        
        for (let i = 0; i < totalChunks; i++) {
            chunksProcessed = i + 1;
            $aiProcessingStatus.html(`Processando bloco ${chunksProcessed} de ${totalChunks}... <i class="fas fa-spinner fa-spin"></i>`);
            
            const startIndex = i * USERS_PER_CHUNK;
            const endIndex = startIndex + USERS_PER_CHUNK;
            const currentChunkData = allData.slice(startIndex, endIndex);

            const chunkSampleForAI = currentChunkData.map(u => ({
                DisplayName: u.DisplayName, OfficeLocation: u.OfficeLocation, JobTitle: u.JobTitle, Licenses: (u.Licenses || []).map(l => l.LicenseName)
            }));

            const systemMessageForChunk = `Você é um assistente de análise de licenças. Este é o bloco de dados ${chunksProcessed} de um total de ${totalChunks} blocos. Os dados contêm informações de aproximadamente ${USERS_PER_CHUNK} usuários (DisplayName, OfficeLocation, JobTitle, Licenses).
Seu objetivo é ANALISAR APENAS ESTE BLOCO e EXTRAIR os seguintes insights de forma concisa (use markdown):
1.  Principais licenças atribuídas NESTE BLOCO (top 3-5 com contagens).
2.  Observações sobre possíveis atribuições excessivas ou insuficientes de licenças NESTE BLOCO (se houver).
3.  Padrões de licenciamento por JobTitle ou OfficeLocation NESTE BLOCO (se houver).
4.  Quaisquer anomalias ou pontos de atenção específicos APENAS deste bloco.
Seja breve e objetivo. Não faça saudações. Limite sua resposta a alguns parágrafos ou uma lista de pontos principais.`;
            
            const chunkPayload = `${systemMessageForChunk}\n\nDADOS DO BLOCO ${chunksProcessed}:\n${JSON.stringify(chunkSampleForAI, null, 2)}`;
            
            console.log(`Enviando Bloco ${chunksProcessed} para Gemini. Tamanho aprox: ${chunkPayload.length} caracteres.`);

            try {
                const finalEndpoint = `${AI_PROVIDER_ENDPOINT_BASE}?key=${apiKey}`;
                const requestHeaders = { 'Content-Type': 'application/json' };
                const requestBody = {
                  contents: [{"role": "user", "parts": [{ "text": chunkPayload }]}],
                  generationConfig: { "maxOutputTokens": MAX_TOKENS_PER_CHUNK_ANALYSIS, "temperature": 0.5 }
                };

                const response = await fetch(finalEndpoint, { method: 'POST', headers: requestHeaders, body: JSON.stringify(requestBody) });
                const responseBodyText = await response.text();
                let data;
                try { data = JSON.parse(responseBodyText); } 
                catch (e) { throw new Error(`Resposta do Bloco ${chunksProcessed} não é JSON. Status: ${response.status}. Resposta: ${responseBodyText.substring(0,200)}`); }

                if (!response.ok) {
                    let detailedMessage = data.error?.message || JSON.stringify(data.error) || responseBodyText;
                    throw new Error(`Erro da API Bloco ${chunksProcessed} (${response.status}): ${detailedMessage}`);
                }

                let chunkSummaryText = `Resumo/Análise do Bloco ${chunksProcessed} não pôde ser extraído.`;
                if (data.candidates && data.candidates[0]?.content?.parts?.[0]?.text) {
                    chunkSummaryText = data.candidates[0].content.parts.map(p => p.text).join("").trim();
                } else if (data.candidates && data.candidates[0]?.finishReason !== "STOP") {
                    chunkSummaryText = `Bloco ${chunksProcessed} processado pela IA com motivo: ${data.candidates[0].finishReason}.`;
                }
                blockSummaries.push({ chunk: chunksProcessed, summary: chunkSummaryText });
                console.log(`Resumo do Bloco ${chunksProcessed} recebido.`);

            } catch (error) {
                console.error(`Erro ao processar bloco ${chunksProcessed}:`, error);
                blockSummaries.push({ chunk: chunksProcessed, summary: `Erro ao processar bloco ${chunksProcessed}: ${error.message}` });
            }
        } 

        $aiProcessingStatus.text(`Todos os ${totalChunks} blocos foram processados! ${blockSummaries.length} resumos gerados. Agora você pode fazer perguntas sobre os dados analisados.`);
        $startAiProcessingButton.prop('disabled', false).html('<i class="fas fa-cogs"></i> Reprocessar Dados para IA');
        
        aiConversationHistory = []; 
        const aggregatedSummaries = blockSummaries.map(bs => `--- INSIGHTS DO BLOCO DE DADOS ${bs.chunk} ---\n${bs.summary}`).join("\n\n");
        
        const initialSystemPromptForQuestions = `Você é um assistente especialista em análise de licenciamento Microsoft 365.
O processamento inicial dos dados de ${allData.length} usuários (divididos em ${totalChunks} blocos) foi concluído.
A seguir estão os resumos/análises gerados pela IA para cada bloco de dados. Use estas informações para responder às perguntas do usuário.
Se a pergunta exigir um detalhe muito específico que não está nos resumos, indique que a análise se baseia nos insights extraídos de cada bloco e que o detalhe individual exato pode não estar disponível.
Seja conciso e use Markdown. Não faça saudações.

RESUMOS DOS BLOCOS DE DADOS PROCESSADOS:
${aggregatedSummaries}
--- FIM DOS RESUMOS DOS BLOCOS ---

Agora, estou pronto para suas perguntas sobre esses dados analisados.`;

        aiConversationHistory.push({ "role": "user", "parts": [{ "text": initialSystemPromptForQuestions }] });
        $aiResponseArea.html("<strong>Contexto dos blocos processado e carregado.</strong><br>Por favor, faça sua pergunta sobre os dados de licença analisados.");
        $aiQuestionInput.prop('disabled', false).attr('placeholder', 'Faça sua pergunta sobre os dados analisados...');
        $askAiButton.prop('disabled', false);
        $aiInteractionArea.removeClass('hidden');
    }); 


    $askAiButton.on('click', async function() {
        let apiKey = getAiApiKey(); // Renomeado para não conflitar com a variável apiKey no escopo do click
        if (!apiKey) {
            alert('Por favor, configure sua chave de API do Google Gemini primeiro.');
            if (!handleApiKeyInputViaPrompt()) return; // Tenta obter a chave via prompt
            apiKey = getAiApiKey(); // Pega a chave recém-inserida
            if (!apiKey) return; // Para se ainda não houver chave
        }

        const userQuestion = $aiQuestionInput.val().trim();
        if (!userQuestion) {
            alert("Por favor, digite sua pergunta.");
            return;
        }
        if (aiConversationHistory.length === 0) {
            alert("Por favor, inicie o processamento dos dados primeiro usando o botão 'Iniciar Processamento'.");
            return;
        }
        
        $aiLoadingIndicator.removeClass('hidden');
        $(this).prop('disabled', true);
        $aiResponseArea.html('Analisando sua pergunta... <i class="fas fa-spinner fa-spin"></i>');

        aiConversationHistory.push({ "role": "user", "parts": [{ "text": userQuestion }] });

        const GEMINI_MODEL_TO_USE = "gemini-1.5-flash-latest"; 
        const AI_PROVIDER_ENDPOINT_BASE = `https://generativelanguage.googleapis.com/v1beta/models/${GEMINI_MODEL_TO_USE}:generateContent`;
        const finalEndpoint = `${AI_PROVIDER_ENDPOINT_BASE}?key=${apiKey}`;

        try {
            const requestHeaders = { 'Content-Type': 'application/json' };
            const requestBody = {
              contents: aiConversationHistory,
              generationConfig: { "maxOutputTokens": MAX_TOKENS_FINAL_QUESTION, "temperature": 0.5 }
            };

            console.log(`Enviando pergunta final para Gemini. Turnos no histórico: ${aiConversationHistory.length}`);

            const response = await fetch(finalEndpoint, { method: 'POST', headers: requestHeaders, body: JSON.stringify(requestBody) });
            const responseBodyText = await response.text();
            let data;
            try { data = JSON.parse(responseBodyText); } 
            catch (e) { throw new Error(`Resposta da pergunta final não é JSON. Status: ${response.status}. Resposta: ${responseBodyText.substring(0,200)}`);}

            if (!response.ok) {
                let detailedMessage = data.error?.message || JSON.stringify(data.error) || responseBodyText;
                throw new Error(`Erro da API Gemini na pergunta final (${response.status}): ${detailedMessage}`);
            }

            let aiTextResponse = "Não foi possível extrair a resposta da IA.";
            if (data.candidates && data.candidates[0]?.content?.parts?.[0]?.text) {
                aiTextResponse = data.candidates[0].content.parts.map(p => p.text).join("").trim();
                aiConversationHistory.push({ "role": "model", "parts": [{ "text": aiTextResponse }] });
            } else if (data.candidates && data.candidates[0]?.finishReason !== "STOP") {
                aiTextResponse = `A IA finalizou com motivo: ${data.candidates[0].finishReason}.`;
            } else {
                aiTextResponse = "A IA processou, mas não retornou texto.";
            }
            
            const MAX_HISTORY_ITEMS = 20; 
            if (aiConversationHistory.length > MAX_HISTORY_ITEMS) {
                const itemsToRemove = aiConversationHistory.length - MAX_HISTORY_ITEMS;
                aiConversationHistory.splice(1, itemsToRemove); 
                console.log(`Histórico de perguntas truncado. Mantendo ${aiConversationHistory.length} itens.`);
            }

            function simpleMarkdownToHtml(mdText) { /* ... (função como antes) ... */ }
            $aiResponseArea.html(simpleMarkdownToHtml(aiTextResponse));
            $aiQuestionInput.val('');

        } catch (error) {
            console.error('Erro ao interagir com IA para pergunta final:', error);
            $aiResponseArea.html(`<strong>Erro na análise:</strong><br>${escapeHtml(error.message)}`);
            if (aiConversationHistory.length > 0 && aiConversationHistory[aiConversationHistory.length -1].role === "user") {
                aiConversationHistory.pop();
            }
        } finally {
            $aiLoadingIndicator.addClass('hidden');
            $(this).prop('disabled', false);
        }
    });
   
    updateAiApiKeyStatusDisplay(); // Atualiza status inicial do botão
    // ===========================================================
    // FUNCIONALIDADE DE IA - FIM
    // ===========================================================

}); // Fim do $(document).ready()
