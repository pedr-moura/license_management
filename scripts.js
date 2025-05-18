$(document).ready(function() {
    let table = null, allData = [], nameConflicts = new Set(), dupLicUsers = new Set();
    const DEBOUNCE_DELAY = 350; // Milliseconds
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
                // Column visibility checkboxes
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

                // Apply initial multi-search if fields are present
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

        $customDropdownTextInput.off();
        $customOptionsList.off();

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
                $(this).trigger('input'); // Populate options on focus
                $customOptionsList.show();
            });

            $customOptionsList.on('mousedown', '.custom-option-item', function(e) { // mousedown to fire before blur
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
                blurTimeout = setTimeout(() => { $customOptionsList.hide(); }, 150); // Delay to allow click on options
            });

        } else { // Not a dropdown field
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
            updateSearchFieldUI($row); // Initial setup based on default selected column

            $row.find('.column-select').on('change', function() {
                updateSearchFieldUI($row); // Update UI when column selection changes
                const selectedColIndex = $row.find('.column-select').val();
                const columnConfig = searchableColumnsConfig.find(c => c.index == selectedColIndex);
                if (columnConfig && columnConfig.useDropdown) {
                    $row.find('.search-value-input').val('').trigger('change'); // Clear hidden value for dropdown
                } else {
                    $row.find('.search-input').val('').trigger('input'); // Clear text input
                }
            });

            $row.find('.search-input').on('input change', applyMultiSearch);
            $row.find('.search-value-input').on('change', applyMultiSearch);

            $row.find('.remove-field').on('click', function() {
                const $multiSearchRow = $(this).closest('.multi-search-row');
                $multiSearchRow.find('.custom-dropdown-text-input').off(); // Clean up listeners
                $multiSearchRow.find('.custom-options-list').off();
                $multiSearchRow.remove();
                applyMultiSearch();
            });
        }

        $('#addSearchField').off('click').on('click', addSearchField);
        $('#multiSearchOperator').off('change').on('change', applyMultiSearch);

        // Add one search field by default if data is loaded
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

        table.search(''); // Clear global search
        while ($.fn.dataTable.ext.search.length > 0) { $.fn.dataTable.ext.search.pop(); } // Clear custom search functions

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
                        col: parseInt(colIndex, 10), // Used for direct apiData access for non-dropdown text search
                        term: searchTerm,
                        dataProp: columnConfig.dataProp, // Used for object property access for dropdowns
                        isDropdown: columnConfig.useDropdown
                    });
                }
            }
        });

        let criteriaText = operator === 'AND' ? 'Criteria: All filters (AND)' : 'Criteria: Any filter (OR)';
        if (filters.length > 0) {
            criteriaText += ` (${filters.length} active filter(s))`;
            $.fn.dataTable.ext.search.push(
                function(settings, apiData, dataIndex) { // apiData is the array of raw string values for the row's cells
                    if (settings.nTable.id !== table.table().node().id) return true; // Ensure this is for our table
                    const rowData = table.row(dataIndex).data(); // Full object data for the row
                    if (!rowData) return false;

                    const logicFn = operator === 'OR' ? filters.some.bind(filters) : filters.every.bind(filters);
                    return logicFn(filter => {
                        let cellDataToTest;
                        if (filter.dataProp === 'Licenses') { // Special handling for Licenses array
                            return (rowData.Licenses && Array.isArray(rowData.Licenses)) ?
                                rowData.Licenses.some(l => (l.LicenseName || '').toLowerCase() === filter.term.toLowerCase()) : false;
                        } else if (filter.isDropdown) { // Exact match for dropdowns using object property
                            cellDataToTest = rowData[filter.dataProp] || '';
                            return String(cellDataToTest).toLowerCase() === filter.term.toLowerCase();
                        } else { // Contains match for text inputs using DataTables' cell data array
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
                $(table.table().header()).find('tr:eq(1) th input').val(''); // Clear individual column filters if they existed
                table.search('').columns().search(''); // Clear global and all column searches
            }
            $('#multiSearchFields .multi-search-row').each(function() {
                $(this).find('.custom-dropdown-text-input').off();
                $(this).find('.custom-options-list').off();
            });
            $('#multiSearchFields').empty();

            if (allData && allData.length > 0) {
                setupMultiSearch(); // Re-add one default search field
            } else {
                $('#searchCriteria').text('No data loaded.');
            }
            while ($.fn.dataTable.ext.search.length > 0) { $.fn.dataTable.ext.search.pop(); } // Clear custom search functions

            if (table) table.draw(); else hideLoader();

            $('#alertPanel').empty(); // Clear alerts
            // Reset column visibility to default
            const defaultVisibleCols = [1, 2, 3, 6]; // Name, Email, Job Title, Licenses
            $('#colContainer .col-vis').each(function() {
                const idx = +$(this).data('col');
                const isDefaultVisible = defaultVisibleCols.includes(idx);
                if (table && idx >= 0 && idx < table.columns().nodes().length) {
                    try { table.column(idx).visible(isDefaultVisible); }
                    catch (e) { console.warn("Error resetting column visibility:", idx, e); }
                }
                $(this).prop('checked', isDefaultVisible);
            });
             if (!table && !(allData && allData.length > 0)) { // If table wasn't even init and no data
                 $('#searchCriteria').text('No data loaded. Please load a JSON file to start.');
             }
        }, 50);
    });

    function downloadCsv(csvContent, fileName) {
        const bom = "\uFEFF"; // BOM for UTF-8
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
            table.columns(':visible').every(function() { // Iterate only over visible columns
                const columnConfig = table.settings()[0].aoColumns[this.index()];
                const colTitle = $(table.column(this.index()).header()).text() || columnConfig.title; // Get header text
                const dataProp = columnConfig.mData; // Get the 'data' property name
                visibleColumns.push({ title: colTitle, dataProp: dataProp });
            });

            const headerRow = visibleColumns.map(col => escapeCsvValue(col.title)).join(',');
            const csvRows = rowsToExport.map(rowData => {
                return visibleColumns.map(col => {
                    let cellData = rowData[col.dataProp]; // Access data using the 'data' property
                    let shouldForceQuotes = false;
                    if (col.dataProp === 'BusinessPhones') {
                        cellData = Array.isArray(cellData) ? cellData.join('; ') : (cellData || '');
                        if (String(cellData).includes(';')) shouldForceQuotes = true;
                    } else if (col.dataProp === 'Licenses') { // Correctly handle licenses array
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
                lines.push([]); // Empty line separator
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
                        escapeCsvValue(joinedDups, true), // Force quotes if semicolons are present
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
                const officeLic = new Set([ // Define licenses that are typically "base" or "suite" licenses
                    'Microsoft 365 E3', 'Microsoft 365 E5',
                    'Microsoft 365 Business Standard', 'Microsoft 365 Business Premium',
                    'Office 365 E3', 'Office 365 E5'
                    // Add other relevant "suite" licenses here if needed
                ]);
                data.forEach(u => {
                    const k = nameKeyInternal(u);
                    nameMap.set(k, (nameMap.get(k) || 0) + 1); // Count occurrences for name+location conflicts

                    const licCount = new Map();
                    (u.Licenses || []).forEach(l => {
                        // More robust way to check for base "Office/Microsoft 365" type licenses
                        const baseName = (l.LicenseName || '');
                        if (baseName) { // Count all licenses
                           licCount.set(baseName, (licCount.get(baseName) || 0) + 1);
                        }
                    });
                    // Check if any of the typical "suite" licenses are duplicated
                    if ([...licCount].some(([lic, c]) => officeLic.has(lic) && c > 1)) {
                        dupSet.add(u.Id);
                    }
                    // Alternative: Check if *any* license is duplicated (more general)
                    // if ([...licCount].some(([,c]) => c > 1)) { dupSet.add(u.Id); }
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
                    if (validatedData.length === 0) { // If no valid data after validation
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
                    self.close(); // Important to close the worker
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
                alert('Error processing data in Worker: ' + error);
                console.error("Worker Error:", error);
                $('#searchCriteria').text('Error loading data.');
                return;
            }

            nameConflicts = new Set(nameConflictsArray);
            dupLicUsers = new Set(dupLicUsersArray);
            uniqueFieldValues = uFValues;

            if (processedData && processedData.length > 0) {
                showLoader('Rendering table...'); // Show loader before table init
                setTimeout(() => { // Allow UI to update loader message
                    initTable(processedData);
                    setupMultiSearch(); // Setup search after table is initialized
                    $('#searchCriteria').text(`Data loaded (${processedData.length} users). Use filters to refine.`);
                    // hideLoader() is called in table.drawCallback
                }, 50); // Small delay for UI update
            } else {
                hideLoader();
                alert('No valid users found in the data. Please check the JSON file or PowerShell script output.');
                $('#searchCriteria').text('No valid data loaded.');
                if (table) { table.clear().draw(); } // Clear table if it exists
                $('#multiSearchFields').empty();
                $('#alertPanel').empty();
            }
        };

        worker.onerror = function(e) {
            URL.revokeObjectURL(blob); // Clean up blob URL
            hideLoader();
            console.error(`Error in Web Worker: Line ${e.lineno} in ${e.filename}: ${e.message}`);
            alert('A critical error occurred during data processing. Please check the console.');
            $('#searchCriteria').text('Critical error loading data.');
        };
        // Pass a structured clone of searchableColumnsConfig to the worker
        worker.postMessage({ rawData: rawData, searchableColumnsConfig: JSON.parse(JSON.stringify(searchableColumnsConfig)) });
    }

    // Initial Data Loading Logic
    try {
        // userData is injected by PowerShell script
        if (typeof userData !== 'undefined' && ( (Array.isArray(userData) && userData.length > 0) || (userData.data && Array.isArray(userData.data) && userData.data.length > 0) ) ) {
            let dataToProcess = Array.isArray(userData) ? userData : userData.data;
            // Check for PowerShell specific error/message structure
            if (userData.error || (Array.isArray(dataToProcess) && dataToProcess.length === 0 && !userData.message) ) {
                 $('#searchCriteria').text(userData.error || 'JSON data is empty or invalid as provided by PowerShell.');
                 console.error("Error in userData from PowerShell or empty data:", userData);
                 hideLoader();
            } else if (userData.message && Array.isArray(dataToProcess) && dataToProcess.length === 0) {
                 $('#searchCriteria').text(userData.message); // E.g., "No user data found or processed."
                 hideLoader();
            }
            else { // Assumed valid data array
                processDataWithWorker(dataToProcess);
            }
        } else if (typeof userData !== 'undefined' && userData.message) { // Handles { "message": "No data...", "data": [] }
             $('#searchCriteria').text(userData.message);
             hideLoader();
        }
        else { // Fallback if userData is not in expected format or not defined
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
    // NOVA FUNCIONALIDADE DE IA - INÍCIO
    // ===========================================================
    const AI_API_KEY_STORAGE_KEY = 'licenseAiApiKey_v1';

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

    // --- Gerenciamento da Chave API da IA ---
    function getAiApiKey() {
        return localStorage.getItem(AI_API_KEY_STORAGE_KEY);
    }

    function updateAiApiKeyStatusDisplay() {
        if (getAiApiKey()) {
            $manageAiApiKeyButton.text('API IA Config.');
            $manageAiApiKeyButton.css('background-color', '#28a745'); // Verde sucesso
            $manageAiApiKeyButton.attr('title', 'Chave da API da IA configurada. Clique para alterar.');
        } else {
            $manageAiApiKeyButton.text('Configurar API IA');
            $manageAiApiKeyButton.css('background-color', ''); // Volta à cor original do botão (button-blue)
            $manageAiApiKeyButton.attr('title', 'Configurar chave da API da IA para análise.');
        }
    }
   
    $manageAiApiKeyButton.on('click', function() {
        $aiApiKeyInput.val(getAiApiKey() || '');
        $aiApiKeyModal.addClass('is-active'); // MOSTRAR MODAL
        $aiApiKeyInput.focus();
    });

    $closeAiModalButton.on('click', function() {
        $aiApiKeyModal.removeClass('is-active'); // ESCONDER MODAL
    });

    $(window).on('click', function(event) {
        if (event.target == $aiApiKeyModal[0]) {
            $aiApiKeyModal.removeClass('is-active'); // ESCONDER MODAL
        }
    });

    $saveAiApiKeyButton.on('click', function() {
        const key = $aiApiKeyInput.val().trim();
        if (key) {
            localStorage.setItem(AI_API_KEY_STORAGE_KEY, key);
            $aiApiKeyMessage.text('Chave da API da IA salva!').removeClass('error').addClass('success');
            updateAiApiKeyStatusDisplay();
            setTimeout(() => {
                $aiApiKeyMessage.text('').removeClass('success error');
                $aiApiKeyModal.removeClass('is-active'); // ESCONDER MODAL
            }, 1500);
        } else {
            $aiApiKeyMessage.text('Por favor, insira uma chave válida.').removeClass('success').addClass('error');
        }
    });

    // --- Interação com a IA ---
    $askAiButton.on('click', async function() {
        const apiKey = getAiApiKey();
        if (!apiKey) {
            alert('Por favor, configure sua chave de API da IA primeiro.');
            $aiApiKeyModal.addClass('is-active'); // MOSTRAR MODAL para configuração
            return;
        }

        const userQuestion = $aiQuestionInput.val().trim();
        let promptContext = "";
        let systemMessage = `Você é um assistente especialista em análise de licenciamento Microsoft 365. Analise os dados fornecidos e responda à pergunta do usuário ou forneça insights gerais se nenhuma pergunta específica for feita. Seja conciso e foque em observações acionáveis. Os dados de licença de usuários são fornecidos em formato JSON. Quando fornecer uma resposta, use formatação Markdown simples (negrito, itálico, listas). Não inclua saudações ou despedidas genéricas, vá direto ao ponto da análise.`;

        if (allData && allData.length > 0) {
            const totalUsers = allData.length;
            const licenseCounts = {};
            allData.forEach(user => {
                (user.Licenses || []).forEach(lic => {
                    licenseCounts[lic.LicenseName] = (licenseCounts[lic.LicenseName] || 0) + 1;
                });
            });
            const topLicenses = Object.entries(licenseCounts)
                .sort(([,a],[,b]) => b-a)
                .slice(0, 5)
                .map(([name, count]) => `${name}: ${count} atribuições`)
                .join('\n  - ');
           
            promptContext = `Resumo dos dados de licença:\n- Total de Usuários: ${totalUsers}\n- Principais Licenças (Top 5 por atribuição):\n  - ${topLicenses}`;
            
            let sampleDataForPrompt = [];
            const maxSample = 5; // Limite de usuários na amostra
            if (allData.length > 0) {
                sampleDataForPrompt = allData.slice(0, maxSample).map(u => ({
                    DisplayName: u.DisplayName, 
                    JobTitle: u.JobTitle, 
                    OfficeLocation: u.OfficeLocation,
                    Licenses: (u.Licenses || []).map(l => l.LicenseName)
                }));
            }
            
            if (allData.length > maxSample) {
                promptContext += `\n\nNota: Os dados completos contêm ${allData.length} usuários. Uma amostra de até ${maxSample} usuários é fornecida abaixo para detalhamento, focando em nome, cargo, local e licenças.`;
            }
            promptContext += `\n\nAmostra de Dados (Nome, Cargo, Local e Licenças):\n${JSON.stringify(sampleDataForPrompt, null, 2)}`;

        } else {
            $aiResponseArea.text('Não há dados de licença carregados para analisar.');
            return;
        }

        let userQueryForAI = userQuestion || "Com base no resumo e na amostra de dados de licença fornecidos, quais são os principais insights, possíveis otimizações de custos ou anomalias que você pode identificar?";
       
        // << ====================================================================== >>
        // << ATENÇÃO: VOCÊ PRECISA CONFIGURAR AS PRÓXIMAS LINHAS PARA SEU PROVEDOR DE IA >>
        // << ====================================================================== >>
        // 1. Defina o AI_PROVIDER_ENDPOINT com a URL correta da API da IA.
        // 2. Adapte os HEADERS (especialmente a Autorização com sua API Key).
        // 3. Adapte o BODY da requisição para o formato esperado pela sua IA (OpenAI, Gemini, Claude, etc.).
        // 4. Adapte a EXTRAÇÃO DA RESPOSTA para pegar o texto da IA corretamente do JSON de resposta.
        // ------------------------------------------------------------------------------

        const AI_PROVIDER_ENDPOINT = 'URL_DA_SUA_API_DE_IA_ESPECIFICA_AQUI'; // <<== CONFIGURE ISTO!

        if (AI_PROVIDER_ENDPOINT === 'URL_DA_SUA_API_DE_IA_ESPECIFICA_AQUI') {
            $aiResponseArea.html('<strong>CONFIGURAÇÃO NECESSÁRIA:</strong><br>O endpoint da API da IA (<code>AI_PROVIDER_ENDPOINT</code>) e a lógica de chamada no arquivo <code>scripts.js</code> precisam ser definidos para o seu provedor de IA específico.');
            return;
        }

        $aiLoadingIndicator.removeClass('hidden');
        $askAiButton.prop('disabled', true);
        $aiResponseArea.html('Analisando com IA... <i class="fas fa-spinner fa-spin"></i>'); // Adiciona ícone de spinner

        try {
            const requestHeaders = {
                'Content-Type': 'application/json',
                // Exemplo para OpenAI (substitua 'apiKey' pela sua chave real):
                // 'Authorization': `Bearer ${apiKey}`,

                // Exemplo para Google AI Studio (Gemini API) - a chave normalmente vai na URL
                // Se a API usar um header específico:
                // 'x-api-key': apiKey, // Exemplo para algumas APIs
                // 'x-goog-api-key': apiKey, // Exemplo para APIs do Google que usam este header
            };
             // ADICIONE AQUI A LÓGICA PARA INCLUIR A CHAVE NO HEADER SE NECESSÁRIO
             // Exemplo:
             // if (AI_PROVIDER_ENDPOINT.toLowerCase().includes("api.openai.com")) {
             //     requestHeaders['Authorization'] = `Bearer ${apiKey}`;
             // } else if (AI_PROVIDER_ENDPOINT.toLowerCase().includes("generativelanguage.googleapis") && !AI_PROVIDER_ENDPOINT.includes("?key=")) {
             //     // Para algumas APIs do Google (não Gemini via REST padrão que usa key na URL)
             //     // requestHeaders['x-goog-api-key'] = apiKey;
             // }


            let finalEndpoint = AI_PROVIDER_ENDPOINT;
            // Exemplo para Google Gemini, onde a chave vai na URL:
            // if (AI_PROVIDER_ENDPOINT.toLowerCase().includes("generativelanguage.googleapis.com")) { // Gemini
            //     if (!finalEndpoint.includes("?key=")) { // Adiciona a chave apenas se não estiver já na URL
            //        finalEndpoint = `${AI_PROVIDER_ENDPOINT}?key=${apiKey}`;
            //     }
            // }

            // Adapte o corpo da requisição para o formato esperado pela sua IA
            let requestBody = {};

            // Exemplo para OpenAI (modelos GPT):
            // requestBody = {
            //     model: "gpt-3.5-turbo", // Ou "gpt-4", "gpt-4o" etc.
            //     messages: [
            //         { role: "system", content: systemMessage },
            //         { role: "user", content: `${promptContext}\n\nPergunta do usuário: ${userQueryForAI}` }
            //     ],
            //     max_tokens: 800, // Ajuste conforme necessário
            //     temperature: 0.5
            // };

            // Exemplo para Google Gemini (Content API):
            // requestBody = {
            //   contents: [{
            //     role: "user", // Ou pode omitir role para prompts simples
            //     parts: [{ text: `${systemMessage}\n\nContexto dos Dados:\n${promptContext}\n\nPergunta do usuário: ${userQueryForAI}` }]
            //   }],
            //   generationConfig: { "maxOutputTokens": 800, "temperature": 0.5 } // Ajuste conforme necessário
            // };

            // !! IMPORTANTE: Se você não definir `requestBody` acima com um exemplo VÁLIDO para sua API, a IA não funcionará.
            if (Object.keys(requestBody).length === 0) { // Verifica se o objeto está vazio
                 $aiResponseArea.html('<strong>ERRO DE CONFIGURAÇÃO:</strong><br>O `requestBody` para a API da IA não foi configurado corretamente no `scripts.js`. Verifique os exemplos comentados e adapte para seu provedor de IA.');
                 $aiLoadingIndicator.addClass('hidden');
                 $askAiButton.prop('disabled', false);
                 return; // Interrompe a execução se o corpo não estiver configurado
            }

            const response = await fetch(finalEndpoint, {
                method: 'POST',
                headers: requestHeaders,
                body: JSON.stringify(requestBody)
            });

            if (!response.ok) {
                const errorText = await response.text(); 
                let errorJson = null;
                try { errorJson = JSON.parse(errorText); } catch (e) { /* não é JSON */ }
                
                let detailedMessage = errorJson ? (errorJson.error?.message || JSON.stringify(errorJson)) : errorText;
                console.error("Erro da API da IA:", response.status, detailedMessage);
                throw new Error(`Erro da API (${response.status}): ${detailedMessage}`);
            }

            const data = await response.json();
            console.log("Resposta completa da API da IA:", data);

            // Extraia a resposta da IA (ALTAMENTE DEPENDENTE DO PROVEDOR)
            let aiTextResponse = "Não foi possível extrair uma resposta de texto da IA. Verifique o console para a resposta completa.";
            
            // Exemplo para OpenAI:
            // if (data.choices && data.choices[0] && data.choices[0].message && data.choices[0].message.content) {
            //     aiTextResponse = data.choices[0].message.content.trim();
            // }
            // Exemplo para Google Gemini:
            // if (data.candidates && data.candidates[0] && data.candidates[0].content && data.candidates[0].content.parts && data.candidates[0].content.parts[0] && data.candidates[0].content.parts[0].text) {
            //    aiTextResponse = data.candidates[0].content.parts.map(part => part.text).join("").trim(); // Concatena todas as partes sem espaço extra
            // }
            
            // Para exibir Markdown de forma simples (substituindo newlines e alguns formatos básicos)
            // Você pode usar uma biblioteca de Markdown para JS (como Showdown.js ou Marked.js) para uma renderização mais rica
            let formattedHtml = escapeHtml(aiTextResponse) // Escapa HTML primeiro para segurança
                                .replace(/\n/g, '<br>')
                                .replace(/\*\*(.*?)\*\*/g, '<strong>$1</strong>') // Negrito
                                .replace(/\*(.*?)\*/g, '<em>$1</em>')     // Itálico
                                .replace(/^- (.*)/gm, '<ul><li>$1</li></ul>') // Listas simples (needs improvement for consecutive items)
                                .replace(/<\/ul>\s*<ul>/g, ''); // Merge consecutive lists

            $aiResponseArea.html(formattedHtml);

        } catch (error) {
            console.error('Erro ao chamar ou processar resposta da API da IA:', error);
            $aiResponseArea.html(`<strong>Erro ao obter análise da IA:</strong><br>${escapeHtml(error.message)}`);
        } finally {
            $aiLoadingIndicator.addClass('hidden');
            $askAiButton.prop('disabled', false);
        }
    });
   
    updateAiApiKeyStatusDisplay(); // Atualiza o status do botão ao carregar
    // ===========================================================
    // NOVA FUNCIONALIDADE DE IA - FIM
    // ===========================================================

}); // Fim do $(document).ready()
