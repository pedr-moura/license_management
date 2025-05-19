$(document).ready(function() {
    let table = null, allData = [], nameConflicts = new Set(), dupLicUsers = new Set();
    const DEBOUNCE_DELAY = 350; // Milliseconds
    let multiSearchDebounceTimer;

    // Definition of operator types in English
    const operatorTypes = {
        IS: 'is',
        IS_NOT: 'is not',
        CONTAINS: 'contains',
        DOES_NOT_CONTAIN: 'does not contain',
    };

    const searchableColumnsConfig = [
        { index: 0, title: 'ID', dataProp: 'Id', useDropdown: false },
        { index: 1, title: 'Name', dataProp: 'DisplayName', useDropdown: false },
        { index: 2, title: 'Email', dataProp: 'Email', useDropdown: false },
        { index: 3, title: 'Job Title', dataProp: 'JobTitle', useDropdown: true },
        { index: 4, title: 'Location', dataProp: 'OfficeLocation', useDropdown: true },
        { index: 5, title: 'Phones', dataProp: 'BusinessPhones', useDropdown: false },
        { index: 6, title: 'Licenses', dataProp: 'Licenses', useDropdown: true }
    ];

    let uniqueFieldValues = {}; // To store unique values for dropdown filters
    const MAX_DROPDOWN_OPTIONS_DISPLAYED = 50; // Max options to show in custom dropdowns

    // Function to show a loading overlay
    function showLoader(message = 'Processing...') {
        let $overlay = $('#loadingOverlay');
        if ($overlay.length === 0 && $('body').length > 0) { // Create overlay if it doesn't exist
            $('body').append('<div id="loadingOverlay"><div class="loader-content"><p id="loaderMessageText"></p></div></div>');
            $overlay = $('#loadingOverlay');
        }
        $overlay.find('#loaderMessageText').text(message);
        $overlay.css('display', 'flex');
    }

    // Function to hide the loading overlay
    function hideLoader() {
        $('#loadingOverlay').hide();
    }

    // Generates a unique key for a user based on DisplayName and OfficeLocation
    const nameKey = u => `${u.DisplayName || ''}|||${u.OfficeLocation || ''}`;

    // Escapes HTML special characters in a string
    const escapeHtml = s => typeof s === 'string' ? s.replace(/[&<>"']/g, m => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' })[m]) : '';

    // Escapes a value for CSV export, quoting if necessary
    function escapeCsvValue(value, forceQuotesOnSemicolon = false) {
        if (value == null) return ''; // Handle null or undefined
        let stringValue = String(value);
        const regex = forceQuotesOnSemicolon ? /[,"\n\r;]/ : /[,"\n\r]/; // Characters that require quoting
        if (regex.test(stringValue)) {
            stringValue = `"${stringValue.replace(/"/g, '""')}"`; // Quote and escape existing quotes
        }
        return stringValue;
    }

    // Renders alerts for name conflicts and duplicate licenses
    function renderAlerts() {
        const $alertPanel = $('#alertPanel').empty(); // Clear previous alerts
        if (nameConflicts.size) {
            $alertPanel.append(`<div class="alert-badge"><button id="filterNameConflicts" style="background: none; border: none; color: inherit; cursor: pointer; text-decoration: underline;" class="underline"><i class="fas fa-users-slash" style="margin-right: 0.3rem;"></i>Name+Location Conflicts: ${nameConflicts.size}</button></div>`);
        }
        if (dupLicUsers.size) {
            const usersToList = allData.filter(u => dupLicUsers.has(u.Id));
            let listHtml = '';
            const maxPreview = 10; // Max users to list in the preview
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

        // Event listener for filtering name conflicts
        $('#filterNameConflicts').off('click').on('click', function() {
            if (!table) return;
            showLoader('Filtering conflicts...');
            setTimeout(() => {
                const conflictUserIds = allData.filter(u => nameConflicts.has(nameKey(u))).map(u => u.Id);
                table.search('').columns().search(''); // Clear existing searches
                if (conflictUserIds.length > 0) {
                    // Search for specific user IDs
                    table.column(0).search('^(' + conflictUserIds.join('|') + ')$', true, false).draw();
                } else {
                    table.column(0).search('').draw(); // Clear search if no conflicts (should not happen if button is clicked)
                }
            }, 50); // Timeout to allow UI update
        });

        // Event listener for toggling alert details
        $('.toggle-details').off('click').on('click', function() {
            const targetId = $(this).data('target');
            $(`#${targetId}`).toggleClass('show');
            $(this).text($(`#${targetId}`).hasClass('show') ? 'Hide' : 'Details');
        });
    }

    // Initializes the DataTable
    function initTable(data) {
        allData = data; // Store all data for global access

        // Populate datalist for license suggestions (if used by browser)
        if ($('#licenseDatalist').length === 0) { $('body').append('<datalist id="licenseDatalist"></datalist>'); }
        const $licenseDatalist = $('#licenseDatalist').empty();
        if (uniqueFieldValues.Licenses) {
            uniqueFieldValues.Licenses.forEach(name => { $licenseDatalist.append($('<option>').attr('value', escapeHtml(name))); });
        }

        // Destroy existing table if it exists, and clean up listeners
        if (table) {
            $(table.table().node()).off('preDraw.dt draw.dt'); // Remove previous DataTables event listeners
            table.destroy();
            $('#licenseTable').empty(); // Clear table HTML to prevent conflicts
        }

        table = $('#licenseTable').DataTable({
            data: allData,
            deferRender: true, // Improves performance for large datasets
            pageLength: 25,    // Default number of rows per page
            orderCellsTop: true, // Enables sorting on header cells
            columns: [ // Column definitions
                { data: 'Id', title: 'ID', visible: false }, // Hidden by default
                { data: 'DisplayName', title: 'Name', visible: true },
                { data: 'Email', title: 'Email', visible: true },
                { data: 'JobTitle', title: 'Job Title', visible: true },
                { data: 'OfficeLocation', title: 'Location', visible: false }, // Hidden by default
                { data: 'BusinessPhones', title: 'Phones', visible: false, render: p => Array.isArray(p) ? p.join('; ') : (p || '') }, // Hidden by default, format array
                { data: 'Licenses', title: 'Licenses', visible: true, render: l => Array.isArray(l) ? l.map(x => x.LicenseName || '').filter(name => name).join(', ') : '' } // Format array of license objects
            ],
            initComplete: function() { // Called when table is fully initialized
                const api = this.api();
                // Logic for individual column text search (if header inputs were used, not primary here)
                api.columns().every(function(colIdx) {
                    const column = this;
                    // Example if you had text inputs in a second header row for DataTables' native search per column
                    // $(api.table().header()).find('tr:eq(1) th:eq(' + colIdx + ') input')
                    //     .off('keyup change clear').on('keyup change clear', function() {
                    //         if (column.search() !== this.value) {
                    //             column.search(this.value).draw();
                    //         }
                    //     });
                });

                // Set initial visibility of columns based on checkboxes
                $('#colContainer .col-vis').each(function() {
                    const idx = +$(this).data('col');
                    try {
                        if (idx >= 0 && idx < api.columns().nodes().length) {
                            $(this).prop('checked', api.column(idx).visible());
                        } else { $(this).prop('disabled', true); } // Disable checkbox if column index is invalid
                    } catch (e) { console.warn("Error checking column visibility:", idx, e); }
                });
                // Event listener for column visibility checkboxes
                $('.col-vis').off('change').on('change', function() {
                    const idx = +$(this).data('col');
                    try {
                        const col = api.column(idx);
                        if (col && col.visible) { col.visible(!col.visible()); }
                        else { console.warn("Column not found for index:", idx); }
                    } catch (e) { console.warn("Error toggling column visibility:", idx, e); }
                });

                // Initialize UI for any pre-existing multi-search fields
                $('#multiSearchFields .multi-search-row').each(function() { updateSearchFieldUI($(this)); });
                applyMultiSearch(); // Apply initial multi-search state
            },
            rowCallback: function(row, data) { // Called for each row drawn
                let classes = '';
                if (nameConflicts.has(nameKey(data))) classes += ' conflict'; // Add class for name conflicts
                if (dupLicUsers.has(data.Id)) classes += ' dup-license';    // Add class for duplicate licenses
                if (classes) $(row).addClass(classes.trim());
                else $(row).removeClass('conflict dup-license'); // Ensure classes are removed if not applicable
            },
            drawCallback: function() { // Called after every table draw
                renderAlerts(); // Update alerts panel
                hideLoader();   // Hide loader after table is drawn
            }
        });
        // Show loader before drawing the table
        $(table.table().node()).on('preDraw.dt', () => showLoader('Updating table...'));
    }

    // Updates the UI of a single multi-search filter row (operators, input type)
    function updateSearchFieldUI($row) {
        const selectedColIndex = $row.find('.column-select').val();
        const columnConfig = searchableColumnsConfig.find(c => c.index == selectedColIndex);

        const $searchInput = $row.find('.search-input'); // Text input for non-dropdown fields
        const $customDropdownContainer = $row.find('.custom-dropdown-container'); // Container for custom dropdown
        const $customDropdownTextInput = $row.find('.custom-dropdown-text-input'); // Text input for custom dropdown
        const $customOptionsList = $row.find('.custom-options-list'); // List of options for custom dropdown
        const $hiddenValueInput = $row.find('.search-value-input'); // Hidden input to store selected value from custom dropdown
        const $conditionOperatorSelect = $row.find('.condition-operator-select'); // Operator select (is, is not, etc.)

        $customDropdownTextInput.off(); // Remove previous event listeners
        $customOptionsList.off();    // Remove previous event listeners
        $conditionOperatorSelect.empty(); // Clear previous operator options

        if (columnConfig) {
            if (columnConfig.useDropdown) { // For columns like Licenses, Job Title, Location
                $searchInput.hide(); // Hide general text input
                $customDropdownContainer.show(); // Show custom dropdown input
                $customDropdownTextInput.attr('placeholder', `Select or type ${columnConfig.title.toLowerCase()}`);
                $customOptionsList.hide().empty();

                // Operators for dropdown/selectable value columns
                $conditionOperatorSelect.append(`<option value="IS" selected>${operatorTypes.IS}</option>`);
                $conditionOperatorSelect.append(`<option value="IS_NOT">${operatorTypes.IS_NOT}</option>`);

                const allUniqueOptions = uniqueFieldValues[columnConfig.dataProp] || [];
                let filterDebounce;
                // Event listener for typing in the custom dropdown's text input
                $customDropdownTextInput.on('input', function() {
                    clearTimeout(filterDebounce);
                    const $input = $(this);
                    filterDebounce = setTimeout(() => {
                        const searchTerm = $input.val().toLowerCase();
                        $customOptionsList.empty().show(); // Clear and show options list
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
                    }, 200); // Debounce input
                });

                // Show options list on focus
                $customDropdownTextInput.on('focus', function() {
                    $(this).trigger('input'); // Populate options on focus
                    $customOptionsList.show();
                });

                // Handle selection from custom dropdown list
                $customOptionsList.on('mousedown', '.custom-option-item', function(e) {
                    e.preventDefault(); // Prevent blur on text input when clicking an option
                    if ($(this).hasClass('no-results')) return;

                    const selectedText = $(this).text();
                    const selectedValue = $(this).data('value');
                    
                    $customDropdownTextInput.val(selectedText); // Update visible text input
                    $hiddenValueInput.val(selectedValue).trigger('change'); // Set hidden input and trigger change for applyMultiSearch
                    $customOptionsList.hide(); // Hide options list
                });

                // Hide options list on blur (with a delay to allow mousedown on options)
                let blurTimeout;
                $customDropdownTextInput.on('blur', function() {
                    clearTimeout(blurTimeout);
                    blurTimeout = setTimeout(() => { $customOptionsList.hide(); }, 150);
                });

            } else { // For free text columns like ID, Name, Email, Phones
                $searchInput.show(); // Show general text input
                $customDropdownContainer.hide(); // Hide custom dropdown
                $hiddenValueInput.val('').hide(); // Ensure hidden input is cleared and hidden
                $searchInput.attr('placeholder', 'Search term...');

                // Operators for free text columns
                $conditionOperatorSelect.append(`<option value="CONTAINS" selected>${operatorTypes.CONTAINS}</option>`);
                $conditionOperatorSelect.append(`<option value="DOES_NOT_CONTAIN">${operatorTypes.DOES_NOT_CONTAIN}</option>`);
                $conditionOperatorSelect.append(`<option value="IS">${operatorTypes.IS}</option>`);
                $conditionOperatorSelect.append(`<option value="IS_NOT">${operatorTypes.IS_NOT}</option>`);
            }
        } else {
            // Fallback or error state - hide inputs and operator select
            $searchInput.hide();
            $customDropdownContainer.hide();
            $conditionOperatorSelect.hide();
        }
    }

    // Sets up the multi-search functionality (add/remove filter rows)
    function setupMultiSearch() {
        const $container = $('#multiSearchFields'); // Container for filter rows
        
        // Function to add a new filter row
        function addSearchField() {
            const columnOptions = searchableColumnsConfig
                .map(c => `<option value="${c.index}">${escapeHtml(c.title)}</option>`)
                .join('');

            // HTML for a new filter row, including the condition operator select
            const $row = $(`
                <div class="multi-search-row">
                    <select class="column-select">${columnOptions}</select>
                    <select class="condition-operator-select"></select> 
                    <input class="search-input" placeholder="Search term..." />
                    <div class="custom-dropdown-container">
                        <input type="text" class="custom-dropdown-text-input" autocomplete="off" />
                        <div class="custom-options-list"></div>
                    </div>
                    <input type="hidden" class="search-value-input" /> 
                    <button class="remove-field" title="Remove filter"><i class="fas fa-trash-alt"></i></button>
                </div>
            `);
            $container.append($row);
            updateSearchFieldUI($row); // Initialize operators and input type for the new row

            // Event listeners for the new row's elements
            $row.find('.column-select').on('change', function() {
                updateSearchFieldUI($row); // Update operators and input type based on new column
                const selectedColConfig = searchableColumnsConfig.find(c => c.index == $(this).val());
                // Clear previous values when column changes to avoid mismatch
                if (selectedColConfig) {
                    if (selectedColConfig.useDropdown) {
                        $row.find('.custom-dropdown-text-input').val('');
                        $row.find('.search-value-input').val('').trigger('change'); // Trigger applyMultiSearch
                    } else {
                        $row.find('.search-input').val('').trigger('input'); // Trigger applyMultiSearch
                    }
                }
            });

            $row.find('.condition-operator-select').off('change').on('change', applyMultiSearch);
            $row.find('.search-input').off('input change').on('input change', applyMultiSearch);
            $row.find('.search-value-input').off('change').on('change', function() { // Hidden input for custom dropdowns
                applyMultiSearch();
            });

            // Event listener for removing a filter row
            $row.find('.remove-field').on('click', function() {
                const $multiSearchRow = $(this).closest('.multi-search-row');
                // Clean up event listeners on child elements before removing
                $multiSearchRow.find('.custom-dropdown-text-input').off();
                $multiSearchRow.find('.custom-options-list').off();
                $multiSearchRow.remove();
                applyMultiSearch(); // Re-apply filters after removal
                // If all filter rows are removed, add a new one if data exists
                if ($container.children().length === 0 && allData && allData.length > 0) {
                     addSearchField();
                }
            });
        }

        $('#addSearchField').off('click').on('click', addSearchField); // "Add Filter" button
        $('#multiSearchOperator').off('change').on('change', applyMultiSearch); // Global AND/OR operator

        // Add an initial search field if data is present and no fields exist yet
        if ($container.children().length === 0) {
            if (allData && allData.length > 0) {
                addSearchField();
            } else {
                $('#searchCriteria').text('No data loaded. Please load a JSON file to start.');
            }
        }
    }

    // Debounced function to apply multi-search filters
    function applyMultiSearch() {
        clearTimeout(multiSearchDebounceTimer);
        showLoader('Applying filters...'); // Consider if loader is too frequent here
        multiSearchDebounceTimer = setTimeout(_executeMultiSearchLogic, DEBOUNCE_DELAY);
    }

    // Core logic for applying multi-search filters to the DataTable
    function _executeMultiSearchLogic() {
        console.log('DEBUG: _executeMultiSearchLogic called');
        const globalOperator = $('#multiSearchOperator').val(); // Global AND or OR operator
        const $searchCriteriaText = $('#searchCriteria');
        
        if (!table) { // Ensure table is initialized
            $searchCriteriaText.text(allData.length === 0 ? 'No data loaded.' : 'Table not initialized.');
            hideLoader();
            return;
        }

        table.search(''); // Clear DataTables' global search
        while ($.fn.dataTable.ext.search.length > 0) { $.fn.dataTable.ext.search.pop(); } // Clear any custom search functions

        // Collect all active filters from the UI
        const filters = [];
        $('#multiSearchFields .multi-search-row').each(function() {
            const $row = $(this);
            const colIndex = $row.find('.column-select').val();
            const columnConfig = searchableColumnsConfig.find(c => c.index == colIndex);
            const conditionOperator = $row.find('.condition-operator-select').val(); // Get selected operator for this row
            let searchTerm = '';

            if (columnConfig) {
                searchTerm = columnConfig.useDropdown ?
                    $row.find('.search-value-input').val() : // Value from hidden input for dropdowns
                    $row.find('.search-input').val().trim();    // Value from text input for others

                // Add filter if a search term is present (even empty string for IS/IS_NOT checks)
                if (searchTerm !== null && searchTerm !== undefined) {
                     filters.push({
                        col: parseInt(colIndex, 10), // Column index in DataTable
                        term: searchTerm,             // Search term
                        dataProp: columnConfig.dataProp, // Property name in rowData object
                        isDropdown: columnConfig.useDropdown, // Is this a dropdown-type field?
                        condition: conditionOperator    // Selected condition (IS, IS_NOT, etc.)
                    });
                }
            }
        });
        console.log('DEBUG: Active filters collected:', filters);

        let criteriaText = globalOperator === 'AND' ? 'Criteria: All filters (AND)' : 'Criteria: Any filter (OR)';
        if (filters.length > 0) {
            criteriaText += ` (${filters.length} active filter(s))`;
            // Push a new custom search function to DataTables
            $.fn.dataTable.ext.search.push(
                function(settings, apiData, dataIndex) { // apiData is the array of data for the row by DataTables
                    if (settings.nTable.id !== table.table().node().id) return true; // Ensure correct table context
                    const rowData = table.row(dataIndex).data(); // Get full data object for the row
                    if (!rowData) return false;

                    // Determine if 'every' (AND) or 'some' (OR) filter conditions must be met
                    const logicFn = globalOperator === 'OR' ? filters.some.bind(filters) : filters.every.bind(filters);
                    
                    return logicFn(filter => { // Apply each filter
                        let cellDataString;
                        // Normalize filter term for comparison, handling null/undefined.
                        let filterTermString = String(filter.term || '').toLowerCase();

                        if (filter.dataProp === 'Licenses') { // Specific logic for 'Licenses' field
                            const userLicenses = (rowData.Licenses && Array.isArray(rowData.Licenses)) ?
                                                 rowData.Licenses.map(l => (l.LicenseName || '').toLowerCase()) : [];
                            
                            let licenseMatch = userLicenses.includes(filterTermString);

                            if (filter.condition === 'IS') return licenseMatch;
                            if (filter.condition === 'IS_NOT') return !licenseMatch;
                            return false; // Default/fallback if condition is unexpected

                        } else if (filter.isDropdown) { // Logic for other dropdown fields (JobTitle, OfficeLocation)
                            cellDataString = String(rowData[filter.dataProp] || '').toLowerCase();
                            
                            if (filter.condition === 'IS') return cellDataString === filterTermString;
                            if (filter.condition === 'IS_NOT') return cellDataString !== filterTermString;
                            return cellDataString === filterTermString; // Default to IS for dropdowns if condition is odd

                        } else { // Logic for free text input fields (ID, Name, Email, Phones)
                            // For non-dropdown, non-license fields, use apiData (DataTables' internal representation)
                            cellDataString = String(apiData[filter.col] || '').toLowerCase();

                            if (filter.condition === 'CONTAINS') return cellDataString.includes(filterTermString);
                            if (filter.condition === 'DOES_NOT_CONTAIN') return !cellDataString.includes(filterTermString);
                            if (filter.condition === 'IS') return cellDataString === filterTermString;
                            if (filter.condition === 'IS_NOT') return cellDataString !== filterTermString;
                            return cellDataString.includes(filterTermString); // Default to CONTAINS for text fields
                        }
                    });
                }
            );
        } else { criteriaText = 'Criteria: All results (no active filters)'; }

        $searchCriteriaText.text(criteriaText); // Update criteria display text
        table.draw(); // Redraw the table to apply filters
        // hideLoader() is called in table.drawCallback
    }

    // Event listener for "Clear Filters" button
    $('#clearFilters').on('click', () => {
        showLoader('Clearing filters...');
        setTimeout(() => {
            if (table) {
                // $(table.table().header()).find('tr:eq(1) th input').val(''); // Clear individual column filters (if used)
                table.search('').columns().search(''); // Clear DT's own search states
            }
            
            // Remove all multi-search rows and their listeners
            $('#multiSearchFields .multi-search-row').each(function() {
                $(this).find('.custom-dropdown-text-input').off();
                $(this).find('.custom-options-list').off();
            });
            $('#multiSearchFields').empty();

            // Clear custom DataTable search functions
            while ($.fn.dataTable.ext.search.length > 0) { $.fn.dataTable.ext.search.pop(); }

            // Re-add a single default search field if data exists
            if (allData && allData.length > 0) {
                // setupMultiSearch() will be called effectively if it detects an empty container and adds a field
                 if ($('#multiSearchFields').children().length === 0) {
                    setupMultiSearch(); // Explicitly call if needed, or ensure addSearchField is called
                }
            } else {
                $('#searchCriteria').text('No data loaded.');
            }
            
            if (table) {
                table.draw(); // This will trigger hideLoader via drawCallback
            } else {
                hideLoader();
            }

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
             if (!table && !(allData && allData.length > 0)) { // If no table and no data
                 $('#searchCriteria').text('No data loaded. Please load a JSON file to start.');
             }
        }, 50); // Timeout to allow UI update
    });

    // Function to download content as a CSV file
    function downloadCsv(csvContent, fileName) {
        const bom = "\uFEFF"; // BOM for UTF-8 Excel compatibility
        const blob = new Blob([bom, csvContent], { type: 'text/csv;charset=utf-8;' });
        const link = document.createElement('a');
        if (link.download !== undefined) { // Check browser support for download attribute
            const url = URL.createObjectURL(blob);
            link.setAttribute('href', url);
            link.setAttribute('download', fileName);
            link.style.visibility = 'hidden';
            document.body.appendChild(link);
            link.click(); // Trigger download
            document.body.removeChild(link); // Clean up
            URL.revokeObjectURL(url); // Release object URL
        } else {
            alert("Your browser does not support direct file downloads.");
        }
    }

    // Event listener for "Export CSV" button
    $('#exportCsv').on('click', () => {
        if (!table) { alert('Table not initialized. No data loaded.'); return; }
        showLoader('Exporting CSV...');
        setTimeout(() => {
            const rowsToExport = table.rows({ search: 'applied' }).data().toArray(); // Get filtered data
            if (!rowsToExport.length) {
                hideLoader();
                alert('No records to export with the current filters.');
                return;
            }
            const visibleColumns = [];
            // Get visible columns in their current display order
            table.columns(':visible').every(function() {
                const columnConfig = table.settings()[0].aoColumns[this.index()];
                const colTitle = $(table.column(this.index()).header()).text() || columnConfig.title || columnConfig.mData;
                const dataProp = columnConfig.mData;
                visibleColumns.push({ title: colTitle, dataProp: dataProp });
            });

            const headerRow = visibleColumns.map(col => escapeCsvValue(col.title)).join(',');
            const csvRows = rowsToExport.map(rowData => {
                return visibleColumns.map(col => {
                    let cellData = rowData[col.dataProp];
                    let shouldForceQuotes = false; // Flag to force quotes if data contains semicolons (for Excel)
                    if (col.dataProp === 'BusinessPhones') {
                        cellData = Array.isArray(cellData) ? cellData.join('; ') : (cellData || '');
                        if (String(cellData).includes(';')) shouldForceQuotes = true;
                    } else if (col.dataProp === 'Licenses') {
                        const licensesArray = (rowData.Licenses && Array.isArray(rowData.Licenses)) ?
                            rowData.Licenses.map(l => l.LicenseName || '').filter(name => name) : [];
                        cellData = licensesArray.length > 0 ? licensesArray.join('; ') : '';
                         if (String(cellData).includes(';')) shouldForceQuotes = true;
                    }
                    // Ensure all values that might contain problematic characters (or semicolons) are quoted
                    return escapeCsvValue(cellData, shouldForceQuotes || String(cellData).match(/[,"\n\r;]/));
                }).join(',');
            });
            const csvContent = [headerRow, ...csvRows].join('\n');
            downloadCsv(csvContent, 'license_report.csv');
            hideLoader();
        }, 50); // Timeout for UI update
    });

    // Event listener for "Issues Report" button
    $('#exportIssues').on('click', () => {
        if (!allData.length) { alert('No data loaded to generate the issues report.'); return; }
        showLoader('Generating issues report...');
        setTimeout(() => {
            const lines = [];
            if (nameConflicts.size) {
                lines.push(['NAME+LOCATION CONFLICTS']);
                lines.push(['Name', 'Location'].map(h => escapeCsvValue(h)));
                nameConflicts.forEach(key => lines.push(key.split('|||').map(value => escapeCsvValue(value))));
                lines.push([]); // Empty line for separation
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
                        escapeCsvValue(joinedDups, true), // Force quotes due to potential semicolons
                        escapeCsvValue(hasPaid ? 'Yes' : 'No')
                    ]);
                });
            }
            if (!lines.length) { lines.push(['No issues detected.']); } // Message if no issues found
            const csvContent = lines.map(rowArray => rowArray.join(',')).join('\n');
            downloadCsv(csvContent, 'issues_report.csv');
            hideLoader();
        }, 50); // Timeout for UI update
    });

    // Processes data using a Web Worker for performance
    function processDataWithWorker(rawData) {
        showLoader('Validating and processing data (this may take a moment)...');
        // Web Worker script content
        const workerScript = `
            const nameKeyInternal = u => \`\${u.DisplayName || ''}|||\${u.OfficeLocation || ''}\`;
            // Validates and normalizes user data
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
                        LicenseName: l.LicenseName || \`Lic_\${i}_\${Math.random().toString(36).substr(2, 5)}\`, // Generate a temp name if missing
                        SkuId: l.SkuId || ''
                    })).filter(l => l.LicenseName) : [] // Ensure licenses have a name
                } : null).filter(x => x); // Filter out any null entries from bad data
                return { validatedData };
            }
            // Finds issues like name conflicts and duplicate licenses
            function findIssuesForWorker(data) {
                const nameMap = new Map(); // For tracking name+location occurrences
                const dupSet = new Set();  // For tracking users with duplicate licenses
                const officeLicBases = [ // List of "Office suite" type licenses often considered exclusive
                    'microsoft 365 e3', 'microsoft 365 e5', 
                    'office 365 e3', 'office 365 e5',
                    'microsoft 365 business standard', 'microsoft 365 business premium',
                ];
                data.forEach(u => {
                    const k = nameKeyInternal(u); // Generate name+location key
                    nameMap.set(k, (nameMap.get(k) || 0) + 1); // Count occurrences
                    const licCount = new Map(); // Count licenses for current user
                    (u.Licenses || []).forEach(l => {
                        const licNameLower = (l.LicenseName || '').toLowerCase();
                        if (officeLicBases.some(base => licNameLower.includes(base))) {
                            const primarySuiteName = officeLicBases.find(base => licNameLower.includes(base)) || licNameLower;
                            licCount.set(primarySuiteName, (licCount.get(primarySuiteName) || 0) + 1);
                        } else {
                             if(l.LicenseName) licCount.set(licNameLower, (licCount.get(licNameLower) || 0) + 1);
                        }
                    });
                    // Mark user if any license (especially suite types) is counted more than once
                    if ([...licCount.values()].some(c => c > 1)) {
                        dupSet.add(u.Id);
                    }
                });
                // Identify conflicting name+location keys
                const conflictingNameKeysArray = [...nameMap].filter(([, count]) => count > 1).map(([key]) => key);
                return { nameConflictsArray: conflictingNameKeysArray, dupLicUsersArray: Array.from(dupSet) };
            }
            // Calculates unique values for fields configured to use dropdowns
            function calculateUniqueFieldValuesForWorker(data, config) {
                const localUniqueFieldValues = {};
                config.forEach(colConfig => {
                    if (colConfig.useDropdown) { // Only for columns configured to use dropdowns
                        if (colConfig.dataProp === 'Licenses') {
                            const allLicenseObjects = data.flatMap(user => user.Licenses || []).filter(l => l.LicenseName);
                            localUniqueFieldValues.Licenses = [...new Set(allLicenseObjects.map(l => l.LicenseName))].sort((a, b) => String(a).toLowerCase().localeCompare(String(b).toLowerCase()));
                        } else {
                            // For other dropdown columns like JobTitle, OfficeLocation
                            localUniqueFieldValues[colConfig.dataProp] = [...new Set(data.map(user => user[colConfig.dataProp]).filter(value => value && String(value).trim() !== ''))].sort((a, b) => String(a).toLowerCase().localeCompare(String(b).toLowerCase()));
                        }
                    }
                });
                return { uniqueFieldValues: localUniqueFieldValues };
            }
            // Web Worker message handler
            self.onmessage = function(e) {
                const { rawData, searchableColumnsConfig: workerSearchableColumnsConfig } = e.data;
                try {
                    const validationResult = validateJsonForWorker(rawData);
                    if (validationResult.error) { // Handle validation errors
                        self.postMessage({ error: validationResult.error });
                        return;
                    }
                    const validatedData = validationResult.validatedData;
                    if (validatedData.length === 0) { // Handle no valid data
                         self.postMessage({ validatedData: [], nameConflictsArray: [], dupLicUsersArray: [], uniqueFieldValues: {} });
                         return;
                    }
                    const issues = findIssuesForWorker(validatedData); // Find data issues
                    const uniqueValues = calculateUniqueFieldValuesForWorker(validatedData, workerSearchableColumnsConfig); // Calculate unique values for filters
                    // Post processed data back to main thread
                    self.postMessage({
                        validatedData: validatedData,
                        nameConflictsArray: issues.nameConflictsArray,
                        dupLicUsersArray: issues.dupLicUsersArray,
                        uniqueFieldValues: uniqueValues.uniqueFieldValues,
                        error: null // No error
                    });
                } catch (err) { // Handle unexpected errors in worker
                    self.postMessage({ error: 'Error in Web Worker: ' + err.message + '\\\\n' + err.stack });
                } finally {
                    self.close(); // Close worker after processing
                }
            };
        `;
        const blob = new Blob([workerScript], { type: 'application/javascript' });
        const worker = new Worker(URL.createObjectURL(blob)); // Create worker

        // Handler for messages received from Web Worker
        worker.onmessage = function(e) {
            URL.revokeObjectURL(blob); // Clean up blob URL
            const { validatedData: processedData, nameConflictsArray, dupLicUsersArray, uniqueFieldValues: uFValues, error } = e.data;

            if (error) { // Handle errors from worker
                hideLoader();
                alert('Error processing data in Worker: ' + error);
                console.error("Worker Error:", error);
                $('#searchCriteria').text('Error loading data.');
                return;
            }

            // Store processed data and issues globally
            nameConflicts = new Set(nameConflictsArray);
            dupLicUsers = new Set(dupLicUsersArray);
            uniqueFieldValues = uFValues;

            if (processedData && processedData.length > 0) { // If data is valid and processed
                showLoader('Rendering table...');
                setTimeout(() => { // Allow UI to update loader message
                    initTable(processedData); // Initialize DataTable with processed data
                    setupMultiSearch();       // Setup multi-search UI
                    $('#searchCriteria').text(`Data loaded (${processedData.length} users). Use filters to refine.`);
                    // Note: hideLoader() is called in initTable's drawCallback
                }, 50);
            } else { // Handle no valid data after processing
                hideLoader();
                alert('No valid users found in the data. Please check the JSON file.');
                $('#searchCriteria').text('No valid data loaded.');
                if (table) { table.clear().draw(); } // Clear table if it exists
                $('#multiSearchFields').empty();     // Clear filter fields
                $('#alertPanel').empty();            // Clear alerts
            }
        };

        // Handler for errors in Web Worker itself
        worker.onerror = function(e) {
            URL.revokeObjectURL(blob); // Clean up blob URL
            hideLoader();
            console.error(`Error in Web Worker: Line ${e.lineno} in ${e.filename}: ${e.message}`);
            alert('A critical error occurred during data processing. Please check the console.');
            $('#searchCriteria').text('Critical error loading data.');
        };
        // Post raw data and config to Web Worker for processing
        worker.postMessage({ rawData: rawData, searchableColumnsConfig: searchableColumnsConfig });
    }

    // Initial data loading logic (tries to use 'userData' global, then falls back to file input)
    try {
        // Check if 'userData' is provided (likely injected by PowerShell script)
        if (typeof userData !== 'undefined' && Array.isArray(userData) && userData.length > 0) {
            processDataWithWorker(userData);
        } else if (typeof userData !== 'undefined' && Array.isArray(userData) && userData.length === 0 && userData.hasOwnProperty('message')) {
            // Handle case where PowerShell script returns a message for no data
             hideLoader();
             $('#searchCriteria').text(userData.message || 'No user data found or processed.');
        }
        else { // Fallback for file input if userData is not available or not valid
            // This part assumes an input#jsonFileInput exists in the HTML if userData is not provided.
            $('#jsonFileInput').on('change', function(event) { // Event listener for file input
                const file = event.target.files[0];
                if (file) {
                    showLoader('Reading JSON file...');
                    const reader = new FileReader();
                    reader.onload = function(e_reader) { // File read successfully
                        try {
                            const jsonData = JSON.parse(e_reader.target.result);
                            processDataWithWorker(jsonData); // Process JSON data
                        } catch (err) { // Handle JSON parsing errors
                            hideLoader();
                            alert('Error parsing JSON file: ' + err.message);
                            console.error("JSON Parse Error:", err);
                            $('#searchCriteria').text('Failed to read JSON.');
                        }
                    };
                    reader.onerror = function() { // Handle file reading errors
                        hideLoader();
                        alert('Error reading file.');
                         $('#searchCriteria').text('Failed to read file.');
                    };
                    reader.readAsText(file); // Read file as text
                }
            });
            // Display message if no data source is readily available
            if ($('#jsonFileInput').length === 0 && (typeof userData === 'undefined' || !Array.isArray(userData))) {
                 console.warn("Global variable 'userData' not defined or empty and input #jsonFileInput not found.");
                 $('#searchCriteria').html('Please load a user JSON file. <br/> (Global <code>userData</code> variable not found or empty).');
            } else if (typeof userData !== 'undefined' && Array.isArray(userData) && userData.length === 0 && !userData.hasOwnProperty('message')) {
                 // If userData is an empty array without a message from PowerShell
                 $('#searchCriteria').text('No user data provided by the global userData variable.');
            }
            // Hide loader if no immediate data processing starts (e.g., waiting for file input)
            if (typeof userData === 'undefined' || (Array.isArray(userData) && userData.length === 0 && !userData.hasOwnProperty('message'))) {
                 hideLoader();
            }
        }
    } catch (error) { // Handle errors during initial loading setup
        hideLoader();
        alert('Error initiating data loading: ' + error.message);
        console.error("Initial loading error:", error);
        $('#searchCriteria').text('Error loading data.');
    }
});
