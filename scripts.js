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
    // FUNCIONALIDADE DE IA - INÍCIO (Configurado para Google Gemini)
    // ===========================================================
    const AI_API_KEY_STORAGE_KEY = 'licenseAiApiKey_v1_google';

    const $manageAiApiKeyButton = $('#manageAiApiKeyButton');
    const $aiApiKeyModal = $('#aiApiKeyModal');
    const $closeAiModalButton = $('#closeAiModalButton');
    const $aiApiKeyInput = $('#aiApiKeyInput');
    const $saveAiApiKeyButton = $('#saveAiApiKeyButton');
    const $aiApiKeyMessage = $('#aiApiKeyMessage');

    const $askAiButton = $('#askAiButton');
    const $aiQuestionInput = $('#aiQuestionInput');
    const $aiResponseArea = $('#aiResponseArea');
    const $aiLoadingIndicator = $('#aiLoadingIndicator');

    function getAiApiKey() {
        return localStorage.getItem(AI_API_KEY_STORAGE_KEY);
    }

    function updateAiApiKeyStatusDisplay() {
        if (getAiApiKey()) {
            $manageAiApiKeyButton.text('API IA Config.');
            $manageAiApiKeyButton.css('background-color', '#28a745');
            $manageAiApiKeyButton.attr('title', 'Chave da API da IA (Google Gemini) configurada. Clique para alterar.');
        } else {
            $manageAiApiKeyButton.text('Configurar API IA');
            $manageAiApiKeyButton.css('background-color', '');
            $manageAiApiKeyButton.attr('title', 'Configurar chave da API da IA (Google Gemini) para análise.');
        }
    }
   
    $manageAiApiKeyButton.on('click', function() {
        $aiApiKeyInput.val(getAiApiKey() || '');
        $aiApiKeyModal.addClass('is-active');
        $aiApiKeyInput.focus();
    });

    $closeAiModalButton.on('click', function() {
        $aiApiKeyModal.removeClass('is-active');
    });

    $aiApiKeyModal.on('click', function(event) {
        if (event.target === $aiApiKeyModal[0]) {
            $aiApiKeyModal.removeClass('is-active');
        }
    });

    $saveAiApiKeyButton.on('click', function() {
        const key = $aiApiKeyInput.val().trim();
        if (key) {
            localStorage.setItem(AI_API_KEY_STORAGE_KEY, key);
            $aiApiKeyMessage.text('Chave da API da IA (Google Gemini) salva!').removeClass('error').addClass('success');
            updateAiApiKeyStatusDisplay();
            setTimeout(() => {
                $aiApiKeyMessage.text('').removeClass('success error');
                $aiApiKeyModal.removeClass('is-active');
            }, 1500);
        } else {
            $aiApiKeyMessage.text('Por favor, insira uma chave válida.').removeClass('success').addClass('error');
        }
    });

    // --- Interação com a IA (Configurado para Google Gemini com Amostragem Contextual) ---
    $askAiButton.on('click', async function() {
        const apiKey = getAiApiKey();
        if (!apiKey) {
            alert('Por favor, configure sua chave de API do Google Gemini primeiro.');
            $aiApiKeyModal.addClass('is-active');
            return;
        }

        const userQuestion = $aiQuestionInput.val().trim();
        const userQuestionLower = userQuestion.toLowerCase();
        let promptContext = "";
        const MAX_SAMPLE_USERS_FOR_AI = 5; // Limite de usuários na amostra para a IA

        // System message ajustado para refletir a nova lógica de amostragem
        let systemMessage = `Você é um assistente especialista em análise de licenciamento Microsoft 365.
Você receberá estatísticas agregadas sobre o uso de licenças e, PARA PERGUNTAS ESPECÍFICAS QUE O EXIJAM, uma PEQUENA AMOSTRA de dados de usuários individuais (no máximo ${MAX_SAMPLE_USERS_FOR_AI} usuários) que podem ser relevantes para a pergunta.
Se a pergunta puder ser respondida com as estatísticas agregadas, use-as primariamente.
Se uma amostra de dados de usuário for fornecida para uma pergunta específica, use essa amostra para detalhar sua resposta.
Se as estatísticas e/ou a amostra não forem suficientes para responder completamente à pergunta, INDIQUE CLARAMENTE quais informações adicionais seriam necessárias ou por que não pode responder com os dados fornecidos.
Seja conciso e foque em observações acionáveis. Use formatação Markdown simples (negrito, itálico, listas). Não inclua saudações ou despedidas.`;

        if (allData && allData.length > 0) {
            const totalUsers = allData.length;
            const licenseMap = new Map();
            let usersWithMultipleHighValueLicenses = 0;
            const highValueLicenseKeywords = ['E5', 'Premium', 'Copilot', 'P2'];
            const departmentLicenseCounts = {};

            allData.forEach(user => {
                let highValueLicenseCount = 0;
                const location = user.OfficeLocation || "Não Especificado";
                if (!departmentLicenseCounts[location]) {
                    departmentLicenseCounts[location] = new Map();
                }
                (user.Licenses || []).forEach(lic => {
                    const licName = lic.LicenseName;
                    licenseMap.set(licName, (licenseMap.get(licName) || 0) + 1);
                    const currentDeptLicMap = departmentLicenseCounts[location];
                    currentDeptLicMap.set(licName, (currentDeptLicMap.get(licName) || 0) + 1);
                    if (highValueLicenseKeywords.some(keyword => licName.toLowerCase().includes(keyword.toLowerCase()))) {
                        highValueLicenseCount++;
                    }
                });
                if (highValueLicenseCount > 1) {
                    usersWithMultipleHighValueLicenses++;
                }
            });

            const uniqueLicenseCount = licenseMap.size;
            const topLicensesOverall = [...licenseMap.entries()]
                .sort(([,a],[,b]) => b-a)
                .slice(0, 10)
                .map(([name, count]) => `  - ${name}: ${count} atribuições`)
                .join('\n');
            
            let departmentSummary = "\nDistribuição de Licenças por Localização/Departamento (Top 3 locais com mais usuários):\n";
            const sortedDepartments = Object.entries(departmentLicenseCounts)
                .map(([dept, licMap]) => ({ name: dept, userCount: new Set(allData.filter(u => (u.OfficeLocation || "Não Especificado") === dept).map(u => u.Id)).size, licMap: licMap }))
                .sort((a,b) => b.userCount - a.userCount)
                .slice(0, 3);
            
            sortedDepartments.forEach(deptInfo => {
                departmentSummary += `  Local/Depto: ${deptInfo.name} (${deptInfo.userCount} usuários)\n`;
                const topDeptLicenses = [...deptInfo.licMap.entries()]
                    .sort(([,a],[,b]) => b-a)
                    .slice(0,3)
                    .map(([name, count]) => `    - ${name}: ${count}`)
                    .join('\n');
                departmentSummary += `${topDeptLicenses}\n`;
            });

            promptContext = `Estatísticas Agregadas do Ambiente de Licenciamento Microsoft 365:
- Total de Usuários Analisados: ${totalUsers}
- Número de Tipos de Licenças Únicas Encontradas: ${uniqueLicenseCount}
- Usuários com Múltiplas Licenças de Alto Valor (ex: E5, Premium, Copilot, P2): ${usersWithMultipleHighValueLicenses}
- Distribuição das Top 10 Licenças Mais Atribuídas (Geral):
${topLicensesOverall}
${departmentSummary}\n`;
            
            // Lógica para decidir se envia amostra detalhada
            let sendDetailedSample = false;
            const detailKeywords = ["quais usuários", "liste usuários", "usuários com", "quem tem", "mesmo produto", "licença duplicada", "licenças pagas"];
            if (userQuestion && detailKeywords.some(kw => userQuestionLower.includes(kw))) {
                sendDetailedSample = true;
            }

            if (sendDetailedSample) {
                let sampleDataForPrompt = [];
                let contextNote = "";

                // Palavras-chave para heurística de "licença paga"
                const paidKeywords = ['E3', 'E5', 'P1', 'P2', 'Premium', 'Copilot', 'Standard', 'Voice', 'Calling', 'Pro', 'Plan 1', 'Plan 2', 'Plan 3', 'Visio', 'Project'];
                const freeOrTrialKeywords = ['free', 'trial', 'developer', 'community', 'essentials', 'fabric (free)', 'student', 'faculty', 'education'];

                // Tentativa de pré-filtragem para "licença paga do mesmo produto"
                if (userQuestionLower.includes("licença paga") && (userQuestionLower.includes("mesmo produto") || userQuestionLower.includes("duplicada"))) {
                    const usersWithPotentialDupPaidLicenses = [];
                    for (const user of allData) {
                        if (usersWithPotentialDupPaidLicenses.length >= MAX_SAMPLE_USERS_FOR_AI) break;
                        const userLicenses = (user.Licenses || []);
                        const paidLicenseCounts = new Map();
                        let hasAnyRelevantPaidLicense = false;

                        userLicenses.forEach(lic => {
                            const licNameClean = (lic.LicenseName || '').trim();
                            const licNameLower = licNameClean.toLowerCase();
                            
                            const isLikelyPaid = paidKeywords.some(pk => licNameLower.includes(pk.toLowerCase())) &&
                                             !freeOrTrialKeywords.some(fk => licNameLower.includes(fk.toLowerCase()));
                            
                            if (isLikelyPaid) {
                                hasAnyRelevantPaidLicense = true;
                                paidLicenseCounts.set(licNameClean, (paidLicenseCounts.get(licNameClean) || 0) + 1);
                            }
                        });

                        if (hasAnyRelevantPaidLicense && Array.from(paidLicenseCounts.values()).some(count => count > 1)) {
                            usersWithPotentialDupPaidLicenses.push({
                                DisplayName: user.DisplayName,
                                OfficeLocation: user.OfficeLocation,
                                Licenses: userLicenses.map(l => l.LicenseName)
                            });
                        }
                    }
                    if (usersWithPotentialDupPaidLicenses.length > 0) {
                        sampleDataForPrompt = usersWithPotentialDupPaidLicenses;
                        contextNote = `\nAMOSTRA DE DADOS DE USUÁRIOS RELEVANTES PARA A PERGUNTA:\nBaseado na pergunta sobre licenças pagas duplicadas do mesmo produto, aqui está uma amostra de até ${MAX_SAMPLE_USERS_FOR_AI} usuários que PARECEM se encaixar nesse critério (identificados por uma pré-filtragem no lado do cliente). Por favor, use esta amostra para sua análise e resposta:\n`;
                    } else {
                        contextNote = `\nUma pergunta sobre licenças pagas duplicadas foi feita, mas uma varredura preliminar no cliente não encontrou usuários que se encaixassem claramente nesse critério para formar uma amostra. Por favor, responda com base nas estatísticas agregadas, se possível, ou indique que os dados detalhados para confirmar isso não foram fornecidos.\n`;
                    }
                } else { 
                    // Amostra genérica para outras perguntas detalhadas ou se a pré-filtragem não for específica
                    // Tenta encontrar usuários com base em palavras-chave da pergunta (licenças, local, cargo)
                    let filteredSample = allData;
                    const questionTokens = userQuestionLower.split(/\s+/);
                    
                    // Tenta filtrar por nomes de licenças mencionados na pergunta
                    const mentionedLicenses = [];
                    if(uniqueFieldValues.Licenses) { // Certifica que uniqueFieldValues.Licenses existe
                        uniqueFieldValues.Licenses.forEach(licName => {
                            if (userQuestionLower.includes(licName.toLowerCase())) {
                                mentionedLicenses.push(licName.toLowerCase());
                            }
                        });
                    }

                    if (mentionedLicenses.length > 0) {
                        filteredSample = filteredSample.filter(u => 
                            (u.Licenses || []).some(l => mentionedLicenses.includes(l.LicenseName.toLowerCase()))
                        );
                    }
                    // Poderia adicionar mais filtros aqui para JobTitle, OfficeLocation se mencionados...

                    if(filteredSample.length > MAX_SAMPLE_USERS_FOR_AI) { // Se o filtro ainda resultar em muitos usuários, pegue uma amostra aleatória deles
                        for(let i = 0; i < MAX_SAMPLE_USERS_FOR_AI; i++) {
                             sampleDataForPrompt.push(filteredSample[Math.floor(Math.random() * filteredSample.length)]);
                        }
                    } else if (filteredSample.length > 0) {
                        sampleDataForPrompt = filteredSample.slice(0, MAX_SAMPLE_USERS_FOR_AI);
                    } else { // Se nenhum filtro pegou, amostra aleatória geral
                         for(let i = 0; i < Math.min(MAX_SAMPLE_USERS_FOR_AI, allData.length); i++) {
                             sampleDataForPrompt.push(allData[Math.floor(Math.random() * allData.length)]);
                        }
                    }
                     sampleDataForPrompt = sampleDataForPrompt.map(u => ({ // Mapeia para os campos desejados
                        DisplayName: u.DisplayName,
                        OfficeLocation: u.OfficeLocation,
                        Licenses: (u.Licenses || []).map(l => l.LicenseName)
                    }));

                    contextNote = `\nAMOSTRA DE DADOS DE USUÁRIOS (POSSIVELMENTE RELEVANTE À PERGUNTA):\nPara ajudar a responder perguntas que podem necessitar de detalhes, segue uma amostra de até ${MAX_SAMPLE_USERS_FOR_AI} usuários (filtrados se possível, ou aleatórios):\n`;
                }
                
                if (sampleDataForPrompt.length > 0) {
                     promptContext += contextNote + `${JSON.stringify(sampleDataForPrompt, null, 2)}\n`;
                } else if (!contextNote.includes("não encontrou usuários")) { // Evita duplicar a mensagem se a pré-filtragem já disse que não achou
                     promptContext += "\nNão foi possível gerar uma amostra de dados de usuário relevante para esta pergunta específica com os filtros automáticos.\n";
                }

            } else { // Se não for uma pergunta detalhada
                 promptContext += "\nNenhuma amostra de dados de usuário individual foi incluída neste prompt, apenas estatísticas agregadas.";
            }
            
        } else {
            $aiResponseArea.text('Não há dados de licença carregados para analisar.');
            return;
        }

        let userQueryForAI = userQuestion || "Com base nas estatísticas de licença e na amostra de dados (se fornecida), quais são os principais insights, possíveis otimizações de custos ou anomalias que você pode identificar?";
       
        const GEMINI_MODEL = "gemini-1.5-flash-latest"; 
        const AI_PROVIDER_ENDPOINT_BASE = `https://generativelanguage.googleapis.com/v1beta/models/${GEMINI_MODEL}:generateContent`;
        const finalEndpoint = `${AI_PROVIDER_ENDPOINT_BASE}?key=${apiKey}`;

        $aiLoadingIndicator.removeClass('hidden');
        $askAiButton.prop('disabled', true);
        $aiResponseArea.html('Analisando com IA (Google Gemini)... <i class="fas fa-spinner fa-spin"></i>');

        try {
            const requestHeaders = { 'Content-Type': 'application/json', };
            const fullPromptForGemini = `${systemMessage}\n\nCONTEXTO DOS DADOS FORNECIDOS (ESTATÍSTICAS E/OU AMOSTRA):\n${promptContext}\n\nPERGUNTA/TAREFA DO USUÁRIO: ${userQueryForAI}`;
            const requestBody = {
              contents: [{"role": "user", "parts": [{ "text": fullPromptForGemini }]}],
              generationConfig: { "maxOutputTokens": 1500, "temperature": 0.4, "topP": 0.95, "topK": 40 }
            };
            
            console.log("Enviando para Google Gemini API (Prompt Completo - Tamanho Aproximado):", fullPromptForGemini.length, "caracteres");
            // console.log("Objeto de Requisição:", JSON.stringify(requestBody, null, 2));
            // console.log("Endpoint:", finalEndpoint);

            const response = await fetch(finalEndpoint, {
                method: 'POST',
                headers: requestHeaders,
                body: JSON.stringify(requestBody)
            });
            
            const responseBodyText = await response.text();
            let data;
            try { data = JSON.parse(responseBodyText); } 
            catch (e) {
                console.error("Falha ao parsear JSON da resposta da API:", responseBodyText);
                throw new Error(`Resposta da API não é um JSON válido. Status: ${response.status}. Resposta: ${responseBodyText.substring(0, 500)}...`);
            }

            if (!response.ok) {
                let detailedMessage = data.error?.message || JSON.stringify(data.error) || responseBodyText;
                console.error("Erro da API Google Gemini:", response.status, detailedMessage);
                throw new Error(`Erro da API Google Gemini (${response.status}): ${detailedMessage}`);
            }

            console.log("Resposta completa da API Google Gemini:", data);

            let aiTextResponse = "Não foi possível extrair uma resposta da IA. Verifique o console.";
            if (data.candidates && data.candidates[0]) {
                const candidate = data.candidates[0];
                if (candidate.content && candidate.content.parts && candidate.content.parts.length > 0 && candidate.content.parts[0].text) {
                    aiTextResponse = candidate.content.parts.map(part => part.text).join("").trim();
                } else if (candidate.finishReason && candidate.finishReason !== "STOP") {
                    aiTextResponse = `A IA finalizou o processamento com o motivo: ${candidate.finishReason}.`;
                    if(data.promptFeedback && data.promptFeedback.blockReason) {
                        aiTextResponse += ` Feedback do prompt: ${data.promptFeedback.blockReason}. Verifique se o conteúdo enviado é permitido.`;
                    } else if (candidate.safetyRatings) {
                         aiTextResponse += ` Safety Ratings: ${JSON.stringify(candidate.safetyRatings)}`;
                    }
                    console.warn("Resposta da IA pode estar bloqueada ou incompleta:", data);
                }
            }
            
            function simpleMarkdownToHtml(mdText) {
                if (typeof mdText !== 'string') return '';
                let html = escapeHtml(mdText);
                html = html.replace(/^### (.*$)/gim, '<h3>$1</h3>');
                html = html.replace(/^## (.*$)/gim, '<h2>$1</h2>');
                html = html.replace(/^# (.*$)/gim, '<h1>$1</h1>');
                html = html.replace(/\*\*(.*?)\*\*|__(.*?)__/g, '<strong>$1$2</strong>');
                html = html.replace(/\*(.*?)\*|_(.*?)_/g, '<em>$1$2</em>');
                
                html = html.replace(/^(?:<br>)?\s*[-*+]\s+(.*(?:<br>\s*[-*+]\s+.*)*)/gm, (match, listItems) => {
                    const items = listItems.split(/<br>\s*[-*+]\s+/);
                    return '<ul>' + items.map(item => `<li>${item.trim()}</li>`).join('') + '</ul>';
                });
                html = html.replace(/<\/ul>\s*<br>\s*<ul>/g, '');
                html = html.replace(/<\/ul><ul>/g, '');

                html = html.replace(/\n/g, '<br>');
                return html;
            }
            $aiResponseArea.html(simpleMarkdownToHtml(aiTextResponse));

        } catch (error) {
            console.error('Erro ao chamar ou processar resposta da API Google Gemini:', error);
            $aiResponseArea.html(`<strong>Erro ao obter análise da IA:</strong><br>${escapeHtml(error.message)}`);
        } finally {
            $aiLoadingIndicator.addClass('hidden');
            $askAiButton.prop('disabled', false);
        }
    });
   
    updateAiApiKeyStatusDisplay();
    // ===========================================================
    // FUNCIONALIDADE DE IA - FIM
    // ===========================================================

}); // Fim do $(document).ready()
