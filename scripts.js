$(document).ready(function() {
    let table = null, allData = [], nameConflicts = new Set(), dupLicUsers = new Set();
    const DEBOUNCE_DELAY = 350; // Milliseconds
    let multiSearchDebounceTimer;

    // Configuration for searchable columns, titles now in English
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
    const MAX_DROPDOWN_OPTIONS_DISPLAYED = 50; // Optimization for custom dropdown list length

    function showLoader(message = 'Processing...') {
        let $overlay = $('#loadingOverlay');
        if ($overlay.length === 0 && $('body').length > 0) {
            // Create structure expected by CSS, without inline styles for presentation
            $('body').append('<div id="loadingOverlay"><div class="loader-content"><p id="loaderMessageText"></p></div></div>');
            $overlay = $('#loadingOverlay'); // Re-select after appending
        }
        $overlay.find('#loaderMessageText').text(message);
        $overlay.css('display', 'flex'); // Use flex for centering defined in your CSS potentially
    }

    function hideLoader() {
        $('#loadingOverlay').hide();
    }

    // Generates a unique key for a user based on DisplayName and OfficeLocation
    const nameKey = u => `${u.DisplayName || ''}|||${u.OfficeLocation || ''}`;

    // Escapes HTML special characters in a string
    const escapeHtml = s => typeof s === 'string' ? s.replace(/[&<>"']/g, m => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' })[m]) : '';

    // Escapes a value for CSV export, optionally forcing quotes if it contains a semicolon
    function escapeCsvValue(value, forceQuotesOnSemicolon = false) {
        if (value == null) return ''; // Handle null or undefined by returning an empty string
        let stringValue = String(value);
        const regex = forceQuotesOnSemicolon ? /[,"\n\r;]/ : /[,"\n\r]/; // Add semicolon to regex if forced
        if (regex.test(stringValue)) {
            stringValue = `"${stringValue.replace(/"/g, '""')}"`; // Double up existing quotes and wrap in quotes
        }
        return stringValue;
    }

    // Renders alert messages for conflicts and duplicate licenses
    function renderAlerts() {
        const $alertPanel = $('#alertPanel').empty();
        if (nameConflicts.size) {
            $alertPanel.append(`<div class="alert-badge"><button id="filterNameConflicts" style="background: none; border: none; color: inherit; cursor: pointer; text-decoration: underline;" class="underline"><i class="fas fa-users-slash" style="margin-right: 0.3rem;"></i>Name+Location Conflicts: ${nameConflicts.size}</button></div>`);
        }
        if (dupLicUsers.size) {
            const usersToList = allData.filter(u => dupLicUsers.has(u.Id));
            let listHtml = '';
            const maxPreview = 10; // Show first N items in the alert
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

        // Event handler to filter by name conflicts
        $('#filterNameConflicts').off('click').on('click', function() {
            if (!table) return;
            showLoader('Filtering conflicts...');
            setTimeout(() => { // Allow loader to show
                const conflictUserIds = allData.filter(u => nameConflicts.has(nameKey(u))).map(u => u.Id);
                table.search('').columns().search(''); // Clear existing searches
                if (conflictUserIds.length > 0) {
                    // Search by ID, using regex for multiple IDs
                    table.column(0).search('^(' + conflictUserIds.join('|') + ')$', true, false).draw();
                } else {
                    table.column(0).search('').draw(); // Clear search if no conflicts
                }
            }, 50); // Small delay for UI update
        });

        // Event handler to toggle details visibility
        $('.toggle-details').off('click').on('click', function() {
            const targetId = $(this).data('target');
            $(`#${targetId}`).toggleClass('show');
            $(this).text($(`#${targetId}`).hasClass('show') ? 'Hide' : 'Details');
        });
    }

    // Initializes the DataTable
    function initTable(data) {
        allData = data; // Store the full dataset

        // Populate datalist for license suggestions (if used by any input, e.g. for accessibility)
        if ($('#licenseDatalist').length === 0) { $('body').append('<datalist id="licenseDatalist"></datalist>'); }
        const $licenseDatalist = $('#licenseDatalist').empty();
        if (uniqueFieldValues.Licenses) {
            uniqueFieldValues.Licenses.forEach(name => { $licenseDatalist.append($('<option>').attr('value', escapeHtml(name))); });
        }

        // Destroy existing table if it exists, and empty the container
        if (table) { $(table.table().node()).off('preDraw.dt draw.dt'); table.destroy(); $('#licenseTable').empty(); }

        table = $('#licenseTable').DataTable({
            data: allData,
            deferRender: true, // Essential for performance with large datasets
            pageLength: 25,    // Default number of rows per page
            orderCellsTop: true, // Allows sorting by clicking on the header cells if search inputs are in a second row
            columns: [ // Column definitions with English titles
                { data: 'Id', title: 'ID', visible: false }, // Hidden by default
                { data: 'DisplayName', title: 'Name', visible: true },
                { data: 'Email', title: 'Email', visible: true },
                { data: 'JobTitle', title: 'Job Title', visible: true },
                { data: 'OfficeLocation', title: 'Location', visible: false }, // Hidden by default
                { data: 'BusinessPhones', title: 'Phones', visible: false, render: p => Array.isArray(p) ? p.join('; ') : (p || '') }, // Hidden by default
                { data: 'Licenses', title: 'Licenses', visible: true, render: l => Array.isArray(l) ? l.map(x => x.LicenseName || '').filter(name => name).join(', ') : '' }
            ],
            initComplete: function() { // Called when the table has been fully initialized
                const api = this.api();
                // Individual column search input setup (if using header search inputs)
                api.columns().every(function(colIdx) {
                    const column = this;
                    // Assuming search inputs are in the second row of the header (thead tr:eq(1))
                    $(api.table().header()).find('tr:eq(1) th:eq(' + colIdx + ') input')
                        .off('keyup change clear').on('keyup change clear', function() {
                            if (column.search() !== this.value) {
                                column.search(this.value).draw();
                            }
                        });
                });

                // Set initial state of column visibility checkboxes
                $('#colContainer .col-vis').each(function() {
                    const idx = +$(this).data('col');
                    try {
                        if (idx >= 0 && idx < api.columns().nodes().length) {
                            $(this).prop('checked', api.column(idx).visible());
                        } else { $(this).prop('disabled', true); } // Disable checkbox if column index is invalid
                    } catch (e) { console.warn("Error checking column visibility:", idx, e); }
                });
                // Event handler for column visibility checkboxes
                $('.col-vis').off('change').on('change', function() {
                    const idx = +$(this).data('col');
                    try {
                        const col = api.column(idx);
                        if (col && col.visible) { col.visible(!col.visible()); } // Toggle visibility
                        else { console.warn("Column not found for index:", idx); }
                    } catch (e) { console.warn("Error toggling column visibility:", idx, e); }
                });

                // Initialize UI for any existing multi-search fields
                $('#multiSearchFields .multi-search-row').each(function() { updateSearchFieldUI($(this)); });
                applyMultiSearch(); // Apply any default or restored multi-search filters
            },
            rowCallback: function(row, data) { // Called for each row created
                // Add classes for highlighting rows with issues
                let classes = '';
                if (nameConflicts.has(nameKey(data))) classes += ' conflict';
                if (dupLicUsers.has(data.Id)) classes += ' dup-license';
                if (classes) $(row).addClass(classes.trim());
                else $(row).removeClass('conflict dup-license'); // Ensure classes are removed if not applicable
            },
            drawCallback: function() { // Called every time the table is drawn
                renderAlerts(); // Update alert messages
                hideLoader();   // Hide loader after table is drawn
            }
        });
        // Show loader before drawing, hideLoader is called by drawCallback
        $(table.table().node()).on('preDraw.dt', () => showLoader('Updating table...'));
    }

    // Updates the UI of a multi-search field (text input vs custom dropdown)
    function updateSearchFieldUI($row) {
        const selectedColIndex = $row.find('.column-select').val();
        const columnConfig = searchableColumnsConfig.find(c => c.index == selectedColIndex);

        const $searchInput = $row.find('.search-input'); // Standard text input
        const $customDropdownContainer = $row.find('.custom-dropdown-container');
        const $customDropdownTextInput = $row.find('.custom-dropdown-text-input'); // Text input for custom dropdown
        const $customOptionsList = $row.find('.custom-options-list'); // List for custom dropdown options
        const $hiddenValueSelect = $row.find('.search-value-select'); // Hidden select to store chosen dropdown value

        // Clear previous event handlers to prevent accumulation
        $customDropdownTextInput.off();
        $customOptionsList.off();

        if (columnConfig && columnConfig.useDropdown) {
            $searchInput.hide(); // Hide standard text input
            $customDropdownContainer.show(); // Show custom dropdown UI
            $customDropdownTextInput.val(''); // Clear any previous text
            $customDropdownTextInput.attr('placeholder', `Type or select ${columnConfig.title.toLowerCase()}`);
            $customOptionsList.hide().empty(); // Ensure options list is initially hidden and empty

            $hiddenValueSelect.empty().append($('<option>').val('').text('')); // Add default empty option
            const allUniqueOptions = uniqueFieldValues[columnConfig.dataProp] || [];
            $hiddenValueSelect.val(''); // Ensure hidden select is cleared

            let filterDebounce;
            $customDropdownTextInput.on('input', function() { // When typing in the custom dropdown input
                clearTimeout(filterDebounce);
                const $input = $(this);
                filterDebounce = setTimeout(() => {
                    const searchTerm = $input.val().toLowerCase();
                    $customOptionsList.empty().show(); // Clear and show list
                    const filteredOptions = allUniqueOptions.filter(opt => String(opt).toLowerCase().includes(searchTerm));

                    if (filteredOptions.length === 0) {
                        $customOptionsList.append('<div class="custom-option-item no-results">No results found</div>');
                    } else {
                        // Display a limited number of options for performance
                        filteredOptions.slice(0, MAX_DROPDOWN_OPTIONS_DISPLAYED).forEach(opt => {
                            const $optionEl = $('<div class="custom-option-item"></div>').text(opt).data('value', opt);
                            $customOptionsList.append($optionEl);
                        });
                        if (filteredOptions.length > MAX_DROPDOWN_OPTIONS_DISPLAYED) {
                             $customOptionsList.append(`<div class="custom-option-item no-results" style="font-style:italic; color: #aaa;">${filteredOptions.length - MAX_DROPDOWN_OPTIONS_DISPLAYED} more options hidden...</div>`);
                        }
                    }
                }, 200); // Debounce filtering
            });

            $customDropdownTextInput.on('focus', function() {
                $(this).trigger('input'); // Trigger input to populate/show list
                $customOptionsList.show();
            });

            // When an option is clicked (mousedown to register before blur)
            $customOptionsList.on('mousedown', '.custom-option-item', function(e) {
                e.preventDefault(); // Prevent blur from closing list before click registers
                if ($(this).hasClass('no-results')) return; // Do nothing if "no results" message is clicked

                const selectedText = $(this).text();
                const selectedValue = $(this).data('value');

                $customDropdownTextInput.val(selectedText); // Update visible input
                $hiddenValueSelect.val(selectedValue).trigger('change'); // Update hidden select and trigger change
                $customOptionsList.hide(); // Hide options list
            });

            let blurTimeout; // To handle hiding the list on blur
            $customDropdownTextInput.on('blur', function() {
                clearTimeout(blurTimeout); // Clear any previous timeout
                // Delay hiding to allow mousedown on an item to be processed
                blurTimeout = setTimeout(() => { $customOptionsList.hide(); }, 150);
            });

        } else { // For standard text input (not a dropdown)
            $searchInput.show().val(''); // Show and clear standard input
            $customDropdownContainer.hide(); // Hide custom dropdown UI
            $hiddenValueSelect.empty().hide(); // Clear and ensure hidden select is hidden (CSS also hides it)
            $searchInput.attr('placeholder', 'Term...'); // Set placeholder for text input
        }
    }

    // Sets up the multi-search functionality and UI
    function setupMultiSearch() {
        const $container = $('#multiSearchFields'); // Container for multi-search rows
        function addSearchField() { // Function to add a new search field row
            // Create options for the column selector
            const columnOptions = searchableColumnsConfig
                .map(c => `<option value="${c.index}">${escapeHtml(c.title)}</option>`)
                .join('');

            // HTML structure for a multi-search row. Relies on CSS for styling.
            const $row = $(`
                <div class="multi-search-row">
                    <select class="column-select">${columnOptions}</select>
                    <input class="search-input" placeholder="Term..." />
                    <div class="custom-dropdown-container">
                        <input type="text" class="custom-dropdown-text-input" autocomplete="off" />
                        <div class="custom-options-list"></div>
                    </div>
                    <select class="search-value-select" style="display: none;"></select> {/* Hidden select stores actual value for dropdowns */}
                    <button class="remove-field" title="Remove filter"><i class="fas fa-trash-alt"></i></button>
                </div>
            `);
            $container.append($row); // Add the new row to the container
            updateSearchFieldUI($row); // Initialize UI for this new row

            // When the column select changes, update the input type UI
            $row.find('.column-select').on('change', function() {
                updateSearchFieldUI($row);
                // Reset value of the input field when column changes
                const selectedColIndex = $row.find('.column-select').val();
                const columnConfig = searchableColumnsConfig.find(c => c.index == selectedColIndex);
                if (columnConfig && columnConfig.useDropdown) {
                    $row.find('.search-value-select').val('').trigger('change'); // Reset hidden select
                } else {
                    $row.find('.search-input').val('').trigger('input'); // Reset text input
                }
            });

            // Event handlers to apply search when input values change
            $row.find('.search-input').on('input change', applyMultiSearch); // For text inputs
            // DEBUG: Modified handler for .search-value-select to include console.log
            $row.find('.search-value-select').on('change', function() {
                console.log('DEBUG: Hidden select .search-value-select changed. Value:', $(this).val(), 'Applying multi-search...');
                applyMultiSearch();
            });

            // Event handler to remove a search field row
            $row.find('.remove-field').on('click', function() {
                const $multiSearchRow = $(this).closest('.multi-search-row');
                // Clean up specific event handlers for custom dropdown before removing
                $multiSearchRow.find('.custom-dropdown-text-input').off();
                $multiSearchRow.find('.custom-options-list').off();
                $multiSearchRow.remove(); // Remove the row
                applyMultiSearch(); // Re-apply filters
            });
        }

        $('#addSearchField').off('click').on('click', addSearchField); // Button to add a new search field
        $('#multiSearchOperator').off('change').on('change', applyMultiSearch); // Operator (AND/OR) change

        // Add an initial search field if none exist and data is loaded
        if ($container.children().length === 0) {
            if (allData && allData.length > 0) {
                addSearchField();
            } else {
                $('#searchCriteria').text('No data loaded. Please load a JSON file to start.');
            }
        }
    }

    // Debounced function to trigger the actual search logic
    function applyMultiSearch() {
        console.log('DEBUG: applyMultiSearch called'); // DEBUG Line
        clearTimeout(multiSearchDebounceTimer); // Clear previous debounce timer
        showLoader('Applying filters...');    // Show loader immediately
        // Set a new timer to execute the search logic after a delay
        multiSearchDebounceTimer = setTimeout(_executeMultiSearchLogic, DEBOUNCE_DELAY);
    }

    // Core logic for executing the multi-search filter on the DataTable
    function _executeMultiSearchLogic() {
        console.log('DEBUG: _executeMultiSearchLogic called'); // DEBUG Line
        const operator = $('#multiSearchOperator').val(); // "AND" or "OR"
        const $searchCriteriaText = $('#searchCriteria'); // Element to display filter criteria summary
        if (!table) { // Ensure table is initialized
            $searchCriteriaText.text(allData.length === 0 ? 'No data loaded.' : 'Table not initialized.');
            hideLoader();
            return;
        }

        table.search(''); // Clear global DataTables search
        while ($.fn.dataTable.ext.search.length > 0) { $.fn.dataTable.ext.search.pop(); } // Clear previous custom search functions

        const filters = []; // Array to hold active filter criteria
        // Iterate over each multi-search row to collect filter terms
        $('#multiSearchFields .multi-search-row').each(function() {
            const colIndex = $(this).find('.column-select').val();
            const columnConfig = searchableColumnsConfig.find(c => c.index == colIndex);
            let searchTerm = '';
            if (columnConfig) {
                // Get search term based on whether it's a dropdown or text input
                searchTerm = columnConfig.useDropdown ?
                    $(this).find('.search-value-select').val() : // Value from hidden select for dropdowns
                    $(this).find('.search-input').val().trim();   // Value from text input

                // DEBUG: Log details of the filter field being inspected
                console.log('DEBUG: Inspecting filter field: Column Title=', columnConfig.title, 'Is Dropdown=', columnConfig.useDropdown, 'Search Term from UI=', searchTerm);

                if (searchTerm) { // Only add filter if a search term is provided
                    filters.push({
                        col: parseInt(colIndex, 10), // DataTable column index (not necessarily the config index)
                        term: searchTerm,
                        dataProp: columnConfig.dataProp, // Property name in the source data object
                        isDropdown: columnConfig.useDropdown
                    });
                }
            }
        });
        // DEBUG: Log the array of collected filters
        console.log('DEBUG: Active filters collected:', filters);

        let criteriaText = operator === 'AND' ? 'Criteria: All filters (AND)' : 'Criteria: Any filter (OR)';
        if (filters.length > 0) {
            criteriaText += ` (${filters.length} active filter(s))`;
            // Add the custom search function to DataTables
            $.fn.dataTable.ext.search.push(
                function(settings, apiData, dataIndex) { // settings, searchDataForRow, rowIndex, originalDataForRow (if serverSide=false)
                    // Ensure this filter applies only to the current table
                    if (settings.nTable.id !== table.table().node().id) return true;

                    const rowData = table.row(dataIndex).data(); // Get the original data object for the row
                    if (!rowData) return false; // Should not happen if table is populated

                    // Determine if row matches based on AND/OR logic
                    const logicFn = operator === 'OR' ? filters.some.bind(filters) : filters.every.bind(filters);
                    return logicFn(filter => { // Check if the row matches this specific filter
                        let cellValue;
                        // Special handling for 'Licenses' as it's an array of objects
                        if (filter.dataProp === 'Licenses') {
                            return (rowData.Licenses && Array.isArray(rowData.Licenses)) ?
                                rowData.Licenses.some(l => (l.LicenseName || '').toLowerCase() === filter.term.toLowerCase()) : false;
                        }
                        // Exact match for other dropdowns
                        else if (filter.isDropdown) {
                            cellValue = rowData[filter.dataProp] || ''; // Get value from original data object
                            return String(cellValue).toLowerCase() === filter.term.toLowerCase();
                        }
                        // "Contains" match for text inputs, using DataTables' prepared search data
                        else {
                            cellValue = apiData[filter.col] || ''; // Use DataTables' search-optimised data for the column
                            return String(cellValue).toLowerCase().includes(filter.term.toLowerCase());
                        }
                    });
                }
            );
        } else { criteriaText = 'Criteria: All results (no active filters)'; }

        $searchCriteriaText.text(criteriaText); // Update criteria summary text
        table.draw(); // Redraw the table to apply filters (will trigger preDraw and drawCallback)
    }

    // Clears all applied filters and resets the UI
    $('#clearFilters').on('click', () => {
        showLoader('Clearing filters...');
        setTimeout(() => { // Allow loader to show
            if (table) {
                $(table.table().header()).find('tr:eq(1) th input').val(''); // Clear individual column header inputs
                table.search('').columns().search(''); // Clear DataTables internal searches
            }
            // Clean up event handlers on custom dropdowns before removing rows
            $('#multiSearchFields .multi-search-row').each(function() {
                $(this).find('.custom-dropdown-text-input').off();
                $(this).find('.custom-options-list').off();
            });
            $('#multiSearchFields').empty(); // Remove all multi-search rows

            if (allData && allData.length > 0) {
                setupMultiSearch(); // Re-adds one default search field and applies empty search
            } else {
                $('#searchCriteria').text('No data loaded.');
            }
            while ($.fn.dataTable.ext.search.length > 0) { $.fn.dataTable.ext.search.pop(); } // Clear custom search functions

            if (table) table.draw(); // Redraw table (will call hideLoader via drawCallback)
            else hideLoader(); // Hide loader if no table to draw

            $('#alertPanel').empty(); // Clear alert messages
            // Reset column visibility to default states
            const defaultVisibleCols = [1, 2, 3, 6]; // Default visible: Name, Email, Job Title, Licenses
            $('#colContainer .col-vis').each(function() {
                const idx = +$(this).data('col');
                const isDefaultVisible = defaultVisibleCols.includes(idx);
                if (table && idx >= 0 && idx < table.columns().nodes().length) {
                    try { table.column(idx).visible(isDefaultVisible); } // Set visibility
                    catch (e) { console.warn("Error resetting column visibility:", idx, e); }
                }
                $(this).prop('checked', isDefaultVisible); // Update checkbox state
            });
             if (!table && !(allData && allData.length > 0)) { // Update criteria if no data and no table
                 $('#searchCriteria').text('No data loaded. Please load a JSON file to start.');
             }
        }, 50);
    });

    // Function to trigger CSV download
    function downloadCsv(csvContent, fileName) {
        const bom = "\uFEFF"; // Byte Order Mark for UTF-8 Excel compatibility
        const blob = new Blob([bom, csvContent], { type: 'text/csv;charset=utf-8;' });
        const link = document.createElement('a');
        if (link.download !== undefined) { // Check for browser support
            const url = URL.createObjectURL(blob);
            link.setAttribute('href', url);
            link.setAttribute('download', fileName);
            link.style.visibility = 'hidden';
            document.body.appendChild(link);
            link.click(); // Simulate click to trigger download
            document.body.removeChild(link); // Clean up
            URL.revokeObjectURL(url); // Release object URL
        } else {
            alert("Your browser does not support direct file downloads.");
        }
    }

    // Event handler for exporting current table view to CSV
    $('#exportCsv').on('click', () => {
        if (!table) { alert('Table not initialized. No data loaded.'); return; }
        showLoader('Exporting CSV...');
        setTimeout(() => { // Allow loader to display
            const rowsToExport = table.rows({ search: 'applied' }).data().toArray(); // Get filtered and sorted data
            if (!rowsToExport.length) {
                hideLoader();
                alert('No records to export with the current filters.');
                return;
            }
            const visibleColumns = [];
            // Get configuration of visible columns
            table.columns(':visible').every(function() {
                const columnConfig = table.settings()[0].aoColumns[this.index()];
                const colTitle = $(table.column(this.index()).header()).text() || columnConfig.title; // Get title from header or config
                const dataProp = columnConfig.mData; // Get data property name
                visibleColumns.push({ title: colTitle, dataProp: dataProp });
            });

            const headerRow = visibleColumns.map(col => escapeCsvValue(col.title)).join(',');
            // Map data rows to CSV format
            const csvRows = rowsToExport.map(rowData => {
                return visibleColumns.map(col => {
                    let cellData = rowData[col.dataProp];
                    let shouldForceQuotes = false; // Flag to force quotes for values with semicolons
                    // Special handling for BusinessPhones and Licenses (arrays/joined strings)
                    if (col.dataProp === 'BusinessPhones') {
                        cellData = Array.isArray(cellData) ? cellData.join('; ') : (cellData || '');
                        if (String(cellData).includes(';')) shouldForceQuotes = true;
                    } else if (col.dataProp === 'Licenses') {
                        const licensesArray = (rowData.Licenses && Array.isArray(rowData.Licenses)) ?
                            rowData.Licenses.map(l => l.LicenseName || '').filter(name => name) : [];
                        cellData = licensesArray.length > 0 ? licensesArray.join('; ') : '';
                         if (String(cellData).includes(';')) shouldForceQuotes = true;
                    }
                    // Escape CSV value, forcing quotes if needed or if special chars exist
                    return escapeCsvValue(cellData, shouldForceQuotes || String(cellData).match(/[,"\n\r]/));
                }).join(',');
            });
            const csvContent = [headerRow, ...csvRows].join('\n');
            downloadCsv(csvContent, 'license_report.csv');
            hideLoader();
        }, 50);
    });

    // Event handler for exporting identified issues to CSV
    $('#exportIssues').on('click', () => {
        if (!allData.length) { alert('No data loaded to generate the issues report.'); return; }
        showLoader('Generating issues report...');
        setTimeout(() => { // Allow loader to display
            const lines = []; // Array to hold CSV lines
            // Section for Name+Location Conflicts
            if (nameConflicts.size) {
                lines.push(['NAME+LOCATION CONFLICTS']); // CSV Section Header
                lines.push(['Name', 'Location'].map(h => escapeCsvValue(h))); // Column Headers
                nameConflicts.forEach(key => lines.push(key.split('|||').map(value => escapeCsvValue(value))));
                lines.push([]); // Empty line for separation
            }
            // Section for Users with Duplicate Licenses
            if (dupLicUsers.size) {
                lines.push(['USERS with Duplicate Licenses']); // CSV Section Header
                lines.push(['Name', 'Location', 'Duplicate Licenses', 'Has Paid License?'].map(h => escapeCsvValue(h))); // Column Headers
                allData.filter(user => dupLicUsers.has(user.Id)).forEach(user => {
                    const licCount = {}, duplicateLicNames = [];
                    (user.Licenses || []).forEach(l => licCount[l.LicenseName] = (licCount[l.LicenseName] || 0) + 1);
                    Object.entries(licCount).forEach(([licName, count]) => count > 1 && duplicateLicNames.push(licName));
                    const joinedDups = duplicateLicNames.join('; '); // Join duplicate license names with semicolon
                    const hasPaid = (user.Licenses || []).some(l => !(l.LicenseName || '').toLowerCase().includes('free'));
                    lines.push([
                        escapeCsvValue(user.DisplayName), escapeCsvValue(user.OfficeLocation),
                        escapeCsvValue(joinedDups, true), // Force quotes if it contains semicolons
                        escapeCsvValue(hasPaid ? 'Yes' : 'No')
                    ]);
                });
            }
            if (!lines.length) { lines.push(['No issues detected.']); } // Message if no issues found
            const csvContent = lines.map(rowArray => rowArray.join(',')).join('\n');
            downloadCsv(csvContent, 'issues_report.csv');
            hideLoader();
        }, 50);
    });


    // --- Initial Data Loading with Web Worker ---
    // Processes raw data using a Web Worker to avoid freezing the main thread
    function processDataWithWorker(rawData) {
        showLoader('Validating and processing data (this may take a moment)...');

        // Define the worker script as a string
        const workerScript = `
            // Helper function inside worker for generating name keys
            const nameKeyInternal = u => \`\${u.DisplayName || ''}|||\${u.OfficeLocation || ''}\`;

            // --- validateJsonForWorker (self-contained for worker) ---
            function validateJsonForWorker(data) {
                if (!Array.isArray(data)) {
                    return { error: 'Invalid JSON: Must be an array of objects.', validatedData: [] };
                }
                // Map and sanitize user data
                const validatedData = data.map((u, i) => u && typeof u === 'object' ? {
                    Id: u.Id || \`unknown_\${i}\`,
                    DisplayName: u.DisplayName || 'Unknown',
                    OfficeLocation: u.OfficeLocation || 'Unknown',
                    Email: u.Email || '',
                    JobTitle: u.JobTitle || 'Unknown',
                    BusinessPhones: Array.isArray(u.BusinessPhones) ? u.BusinessPhones : (typeof u.BusinessPhones === 'string' ? u.BusinessPhones.split('; ').filter(p => p) : []),
                    Licenses: Array.isArray(u.Licenses) ? u.Licenses.map(l => ({
                        LicenseName: l.LicenseName || \`Lic_\${i}_\${Math.random().toString(36).substr(2, 5)}\`, // Generate random name if missing
                        SkuId: l.SkuId || ''
                    })).filter(l => l.LicenseName) : [] // Ensure license has a name
                } : null).filter(x => x); // Filter out any null entries from malformed objects
                return { validatedData };
            }

            // --- findIssuesForWorker (self-contained for worker) ---
            function findIssuesForWorker(data) {
                const nameMap = new Map(); // For tracking name+location uniqueness
                const dupSet = new Set();  // For tracking users with duplicate MS Office licenses
                // Set of common paid Microsoft Office/365 licenses to check for duplicates
                const officeLic = new Set([
                    'Microsoft 365 E3', 'Microsoft 365 E5',
                    'Microsoft 365 Business Standard', 'Microsoft 365 Business Premium',
                    'Office 365 E3', 'Office 365 E5'
                ]);
                data.forEach(u => {
                    const k = nameKeyInternal(u); // Generate unique key for name+location
                    nameMap.set(k, (nameMap.get(k) || 0) + 1); // Count occurrences
                    
                    const licCount = new Map(); // Count occurrences of each license for the current user
                    (u.Licenses || []).forEach(l => {
                        // Normalize license names to base product (e.g., "Microsoft 365 E3" from variations)
                        const baseName = (l.LicenseName || '').match(/^(Microsoft 365|Office 365)/)?.[0] ?
                            (l.LicenseName.match(/^(Microsoft 365 E3|Microsoft 365 E5|Microsoft 365 Business Standard|Microsoft 365 Business Premium|Office 365 E3|Office 365 E5)/)?.[0] || l.LicenseName)
                            : l.LicenseName;
                        if (baseName) { licCount.set(baseName, (licCount.get(baseName) || 0) + 1); }
                    });
                    // Check if any of the specified office licenses are duplicated
                    if ([...licCount].some(([lic, c]) => officeLic.has(lic) && c > 1)) {
                        dupSet.add(u.Id); // Add user ID to the set of users with duplicate licenses
                    }
                });
                // Identify name+location keys that are not unique
                const conflictingNameKeysArray = [...nameMap].filter(([, count]) => count > 1).map(([key]) => key);
                return { nameConflictsArray: conflictingNameKeysArray, dupLicUsersArray: Array.from(dupSet) };
            }

            // --- calculateUniqueFieldValuesForWorker (self-contained for worker) ---
            function calculateUniqueFieldValuesForWorker(data, config) {
                const localUniqueFieldValues = {};
                config.forEach(colConfig => {
                    if (colConfig.useDropdown) { // Only for columns configured to use dropdowns
                        if (colConfig.dataProp === 'Licenses') { // Special handling for Licenses
                            // Flatten all license names from all users into a single array
                            const allLicenseObjects = data.flatMap(user => user.Licenses || []).filter(l => l.LicenseName);
                            // Get unique, sorted license names
                            localUniqueFieldValues.Licenses = [...new Set(allLicenseObjects.map(l => l.LicenseName))].sort((a, b) => a.toLowerCase().localeCompare(b.toLowerCase()));
                        } else { // For other dropdown-able properties
                            // Get unique, sorted values for the property, filtering out empty/null values
                            localUniqueFieldValues[colConfig.dataProp] = [...new Set(data.map(user => user[colConfig.dataProp]).filter(value => value && String(value).trim() !== ''))].sort((a, b) => a.toLowerCase().localeCompare(b.toLowerCase()));
                        }
                    }
                });
                return { uniqueFieldValues: localUniqueFieldValues };
            }

            // Worker's main message handler
            self.onmessage = function(e) {
                const { rawData, searchableColumnsConfig: workerSearchableColumnsConfig } = e.data; // searchableColumnsConfig received from main thread
                try {
                    const validationResult = validateJsonForWorker(rawData);
                    if (validationResult.error) { // If validation fails, post error back
                        self.postMessage({ error: validationResult.error });
                        return;
                    }
                    const validatedData = validationResult.validatedData;
                    if (validatedData.length === 0) { // If no valid data, post empty results
                         self.postMessage({ validatedData: [], nameConflictsArray: [], dupLicUsersArray: [], uniqueFieldValues: {} });
                         return;
                    }

                    // Process data to find issues and unique values
                    const issues = findIssuesForWorker(validatedData);
                    const uniqueValues = calculateUniqueFieldValuesForWorker(validatedData, workerSearchableColumnsConfig);
                    
                    // Post successful results back to the main thread
                    self.postMessage({
                        validatedData: validatedData,
                        nameConflictsArray: issues.nameConflictsArray, // Send as Array
                        dupLicUsersArray: issues.dupLicUsersArray,     // Send as Array
                        uniqueFieldValues: uniqueValues.uniqueFieldValues,
                        error: null // No error
                    });
                } catch (err) { // Catch any unexpected errors during worker processing
                    self.postMessage({ error: 'Error in Web Worker: ' + err.message + '\\n' + err.stack });
                } finally {
                    self.close(); // Terminate the worker after processing is complete
                }
            };
        `;
        const blob = new Blob([workerScript], { type: 'application/javascript' });
        const worker = new Worker(URL.createObjectURL(blob)); // Create worker from blob URL

        // Handler for messages received from the worker
        worker.onmessage = function(e) {
            URL.revokeObjectURL(blob); // Clean up blob URL as it's no longer needed
            const { validatedData: processedData, nameConflictsArray, dupLicUsersArray, uniqueFieldValues: uFValues, error } = e.data;

            if (error) { // If worker posted an error
                hideLoader();
                alert('Error processing data in Worker: ' + error);
                console.error("Worker Error:", error);
                $('#searchCriteria').text('Error loading data.');
                return;
            }

            // Store processed data and issues (convert arrays back to Sets for efficient lookup)
            nameConflicts = new Set(nameConflictsArray);
            dupLicUsers = new Set(dupLicUsersArray);
            uniqueFieldValues = uFValues;

            if (processedData && processedData.length > 0) { // If data was successfully processed
                showLoader('Rendering table...');
                // Use setTimeout to allow the loader message to update before heavy table initialization
                setTimeout(() => {
                    initTable(processedData);    // Initialize the DataTable
                    setupMultiSearch();          // Setup multi-search UI
                    $('#searchCriteria').text(`Data loaded (${processedData.length} users). Use filters to refine.`);
                    // hideLoader() will be called by initTable's drawCallback
                }, 50);
            } else { // If no valid users were found
                hideLoader();
                alert('No valid users found in the data. Please check the JSON file.');
                $('#searchCriteria').text('No valid data loaded.');
                // Clear UI elements if data is empty
                if (table) { table.clear().draw(); }
                $('#multiSearchFields').empty();
                $('#alertPanel').empty();
            }
        };

        // Handler for errors occurring in the worker itself
        worker.onerror = function(e) {
            URL.revokeObjectURL(blob); // Clean up blob URL
            hideLoader();
            console.error(`Error in Web Worker: Line ${e.lineno} in ${e.filename}: ${e.message}`);
            alert('A critical error occurred during data processing. Please check the console.');
            $('#searchCriteria').text('Critical error loading data.');
        };

        // Send raw data and column configuration to the worker to start processing
        worker.postMessage({ rawData: rawData, searchableColumnsConfig: searchableColumnsConfig });
    }


    // --- Initial Load Trigger ---
    // Tries to load data from global 'userData' or prompts for file input
    try {
        // Check if 'userData' is defined globally (e.g., in a <script> tag in the HTML)
        if (typeof userData !== 'undefined' && Array.isArray(userData)) {
            processDataWithWorker(userData); // Process pre-loaded data
        } else {
            // Fallback or primary method: allow user to load JSON via file input
            $('#jsonFileInput').on('change', function(event) {
                const file = event.target.files[0];
                if (file) {
                    showLoader('Reading JSON file...');
                    const reader = new FileReader();
                    reader.onload = function(e_reader) { // Renamed event var to avoid conflict with worker's e
                        try {
                            const jsonData = JSON.parse(e_reader.target.result);
                            processDataWithWorker(jsonData); // Process data from file
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
                    reader.readAsText(file); // Start reading the file
                }
            });
            // UI feedback if no initial data and no file input found/used
            if ($('#jsonFileInput').length === 0) {
                 console.warn("Global variable 'userData' not defined and input #jsonFileInput not found. Please add a file input or define 'userData'.");
                 $('#searchCriteria').html('Please load a user JSON file. <br/> (Global <code>userData</code> variable not found.)');
            } else {
                 $('#searchCriteria').text('Please load a user JSON file.');
            }
            hideLoader(); // Hide if no initial userData and no file input interaction yet
        }
    } catch (error) { // Catch errors during the initial setup phase
        hideLoader();
        alert('Error initiating data loading: ' + error.message);
        console.error("Initial loading error:", error);
        $('#searchCriteria').text('Error loading data.');
    }
});
