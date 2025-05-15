$(document).ready(function() {
  let table = null, allData = [], nameConflicts = new Set(), dupLicUsers = new Set();
  let uniqueLicenseNames = [];

  const DEBOUNCE_DELAY = 350;
  let multiSearchDebounceTimer;

  // Functions to control the loading overlay
  function showLoader() {
    if ($('#loadingOverlay').length === 0 && $('body').length > 0) {
        $('body').append('<div id="loadingOverlay" style="display:none; position:fixed; top:0; left:0; width:100%; height:100%; background:rgba(0,0,0,0.5); z-index:10000; display:flex; align-items:center; justify-content:center;"><div style="background:white; color:black; padding:20px; border-radius:5px;">Processing...</div></div>');
    }
    $('#loadingOverlay').show();
  }

  function hideLoader() {
    $('#loadingOverlay').hide();
  }

  // Key generator for Name + OfficeLocation conflicts
  const nameKey = u => `${u.DisplayName || ''}|||${u.OfficeLocation || ''}`;

  // Escape HTML entities
  const escapeHtml = s => typeof s === 'string' ? s.replace(/[&<>"']/g, m => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' })[m]) : '';

  // Escape value for CSV, optionally forcing quotes if a semicolon is present
  function escapeCsvValue(value, forceQuotesOnSemicolon = false) {
    if (value == null) return '';
    let stringValue = String(value);
    // Quote if contains comma, double quote, newline, carriage return
    // Optionally, also quote if contains semicolon
    const regex = forceQuotesOnSemicolon ? /[,"\n\r;]/ : /[,"\n\r]/;

    if (regex.test(stringValue)) {
      stringValue = `"${stringValue.replace(/"/g, '""')}"`;
    }
    return stringValue;
  }

  // Validate and normalize input JSON data
  function validateJson(data) {
    if (!Array.isArray(data)) {
      alert('Invalid JSON: Must be an array of objects.');
      return [];
    }
    return data.map((u, i) => u && typeof u === 'object' ? {
      Id: u.Id || `unknown_${i}`,
      DisplayName: u.DisplayName || 'Unknown',
      OfficeLocation: u.OfficeLocation || 'Unknown',
      Email: u.Email || '',
      JobTitle: u.JobTitle || '',
      // Ensure BusinessPhones is an array; split string by semicolon if not already
      BusinessPhones: Array.isArray(u.BusinessPhones) ? u.BusinessPhones : (typeof u.BusinessPhones === 'string' ? u.BusinessPhones.split('; ').filter(p => p) : []),
      // Ensure Licenses is an array of objects with basic properties
      Licenses: Array.isArray(u.Licenses) ? u.Licenses.map(l => ({
        LicenseName: l.LicenseName || `Lic_${i}_${Math.random().toString(36).substr(2, 5)}`, // Generate a unique name if missing
        SkuId: l.SkuId || ''
      })).filter(l => l.LicenseName) : [] // Filter out licenses without a name
    } : null).filter(x => x); // Filter out any null entries from invalid objects
  }


  // Find data issues (Name+Location conflicts, duplicate Office licenses)
  function findIssues(data) {
    const nameMap = new Map(); // Map to count occurrences of Name+Location
    const dupSet = new Set(); // Set of User IDs with duplicate main Office licenses
    // List of key Office license names to check for duplicates per user
    const officeLic = new Set([
      'Microsoft 365 E3', 'Microsoft 365 E5',
      'Microsoft 365 Business Standard', 'Microsoft 365 Business Premium',
      'Office 365 E3', 'Office 365 E5'
    ]);

    data.forEach(u => {
      // Check for Name+Location conflicts
      const k = nameKey(u);
      nameMap.set(k, (nameMap.get(k) || 0) + 1);

      // Check for duplicate Office licenses for this user
      const licCount = new Map();
      (u.Licenses || []).forEach(l => {
        // Normalize license name for counting (e.g., 'Microsoft 365 E3 (Add-on)' -> 'Microsoft 365 E3')
        // Or just use the full name if it doesn't match the major suites
        const baseName = (l.LicenseName || '').match(/^(Microsoft 365|Office 365)/)?.[0] ?
                         (l.LicenseName.match(/^(Microsoft 365 E3|Microsoft 365 E5|Microsoft 365 Business Standard|Microsoft 365 Business Premium|Office 365 E3|Office 365 E5)/)?.[0] || l.LicenseName)
                        : l.LicenseName;

        if (baseName) {
          licCount.set(baseName, (licCount.get(baseName) || 0) + 1);
        }
      });

      // If any of the main Office licenses appear more than once for this user
      if ([...licCount].some(([lic, c]) => officeLic.has(lic) && c > 1)) {
        dupSet.add(u.Id);
      }
    });

    // Collect Name+Location keys where count > 1
    const conflictingNameKeys = new Set([...nameMap].filter(([, count]) => count > 1).map(([key]) => key));

    return { nameConflicts: conflictingNameKeys, dupLicUsers: dupSet };
  }


  // Render alert badges for detected issues
  function renderAlerts() {
    const $p = $('#alertPanel').empty(); // Clear previous alerts

    // Name+Location Conflicts alert
    if (nameConflicts.size) {
      $p.append(`<div class="alert-badge">
        <button id="filterNameConflicts" style="background: none; border: none; color: inherit; cursor: pointer; text-decoration: underline;" class="underline"><i class="fas fa-users-slash" style="margin-right: 0.3rem;"></i>Name+Location Conflicts: ${nameConflicts.size}</button>
      </div>`);
    }

    // Duplicate Licenses alert
    if (dupLicUsers.size) {
      // Build list of users with duplicate licenses for the details section
      const list = allData.filter(u => dupLicUsers.has(u.Id)).map(u => {
        const licCount = {}, duplicateLicNames = [];
        (u.Licenses || []).forEach(l => {
            // Count *all* licenses for details list, not just the main office ones
            licCount[l.LicenseName] = (licCount[l.LicenseName] || 0) + 1
        });
        // Find which specific licenses are duplicated for this user
        Object.entries(licCount).forEach(([licName, count]) => count > 1 && duplicateLicNames.push(licName));

        const hasPaid = (u.Licenses || []).some(l => !(l.LicenseName || '').toLowerCase().includes('free'));
        return `<li>${escapeHtml(u.DisplayName)} (${escapeHtml(u.OfficeLocation)}): ${escapeHtml(duplicateLicNames.join(', '))} — Paid: ${hasPaid ? 'Yes' : 'No'}</li>`;
      }).join('');

      $p.append(`<div class="alert-badge">
        <span><i class="fas fa-copy" style="margin-right: 0.3rem;"></i>Duplicate Licenses: ${dupLicUsers.size}</span>
        <button class="underline toggle-details" data-target="dupDetails" style="background: none; border: none; color: inherit; cursor: pointer; text-decoration: underline;">Details</button>
      </div>
      <div id="dupDetails" class="alert-details"><ul>${list}</ul></div>`);
    }

    // Attach click handler for Name+Location filter button
    $('#filterNameConflicts').off('click').on('click', function() {
      if (!table) return;
      // Filter table to show only users whose Name+Location combination has conflicts
      const conflictUserIds = allData.filter(u => nameConflicts.has(nameKey(u))).map(u => u.Id);
      // Clear existing DataTables global/column searches before applying this filter
      table.search('').columns().search('');
      // Use a regex search on the hidden ID column (column 0) to match any of the conflict IDs
      // The regex `^(${ids.join('|')})$` matches if the cell value is exactly one of the IDs
      table.column(0).search('^(' + conflictUserIds.join('|') + ')$', true, false).draw();
      // Note: This filter replaces any active multi-search or other column filters
      // Clearing multi-search fields here might be desired, but let's keep it simple for now.
    });

    // Attach click handler for toggling details visibility
    $('.toggle-details').off('click').on('click', function() {
      const targetId = $(this).data('target');
      $(`#${targetId}`).toggleClass('show');
      $(this).text($(`#${targetId}`).hasClass('show') ? 'Hide' : 'Details');
    });
  }

  // Initialize the DataTables table
  function initTable(data) {
    allData = data; // Store processed data

    // Find initial issues
    ({ nameConflicts, dupLicUsers } = findIssues(allData));

    // Extract and sort unique license names for the datalist and multi-search select
    const allLicenseObjects = allData.flatMap(user => user.Licenses || []).filter(l => l.LicenseName);
    uniqueLicenseNames = [...new Set(allLicenseObjects.map(l => l.LicenseName))].sort();

    // Populate the datalist for license name input suggestions (if used elsewhere)
    if ($('#licenseDatalist').length === 0) {
      $('body').append('<datalist id="licenseDatalist"></datalist>');
    }
    const $licenseDatalist = $('#licenseDatalist').empty();
    uniqueLicenseNames.forEach(name => {
      $licenseDatalist.append($('<option>').attr('value', escapeHtml(name)));
    });

    // Destroy existing table if it exists
    if (table) {
      // Unbind event listeners before destroying
      $(table.table().node()).off('preDraw.dt draw.dt');
      table.destroy();
    }

    // Initialize DataTables
    table = $('#licenseTable').DataTable({
      data: allData,
      deferRender: true, // Improve performance with large data
      pageLength: 25,
      orderCellsTop: true, // Needed for column filters in header row
      // language: { ... } // <-- Your full language object goes here if needed
      columns: [
        // Hidden ID column for internal filtering (like Name+Location conflicts)
        { data: 'Id', title: 'ID', visible: false },
        { data: 'DisplayName', title: 'Name', visible: true },
        { data: 'Email', title: 'Email', visible: true },
        { data: 'JobTitle', title: 'Job Title', visible: true },
        { data: 'OfficeLocation', title: 'Location', visible: false },
        {
          data: 'BusinessPhones',
          title: 'Phones',
          visible: false,
          render: p => Array.isArray(p) ? p.join('; ') : (p || '') // Render array as semicolon-separated string
        },
        {
          data: 'Licenses',
          title: 'Licenses',
          visible: true,
          render: l => Array.isArray(l) ? l.map(x => x.LicenseName || '').filter(name => name).join(', ') : '' // Render array of license names
        }
      ],
      initComplete: function() {
        const api = this.api();

        // Setup individual column search inputs in the header row
        api.columns().every(function(colIdx) {
          const column = this;
          // Find the input element in the second header row for this column
          const input = $(api.table().header()).find('tr:eq(1) th:eq(' + colIdx + ') input');
          if (input.length > 0) {
            input.off('keyup change clear').on('keyup change clear', function() {
              // Apply search on keyup, change, or clear
              if (column.search() !== this.value) {
                column.search(this.value).draw();
              }
            });
          }
        });

        // Sync column visibility checkboxes with initial table state
        $('#colContainer .col-vis').each(function() {
          const idx = +$(this).data('col');
          try {
            // Check if the column index is valid before accessing column()
            if (idx >= 0 && idx < api.columns().nodes().length) {
              $(this).prop('checked', api.column(idx).visible());
            } else {
              console.warn("Invalid column index for visibility toggle:", idx);
              $(this).prop('disabled', true); // Disable checkbox if index is invalid
            }
          } catch (e) { console.warn("Error checking column visibility:", idx, e); }
        });

        // Attach change listener to column visibility checkboxes
        $('.col-vis').off('change').on('change', function() {
          const idx = +$(this).data('col');
          try {
            const col = api.column(idx);
            // Ensure column exists before toggling
            if (col && col.visible) {
              col.visible(!col.visible());
            } else {
              console.warn("Column object not found or visible method missing for index:", idx);
            }
          } catch (e) { console.warn("Error toggling column visibility:", idx, e); }
        });

        // Initialize state for multi-search license input fields
        $('#multiSearchFields .multi-search-row').each(function() { updateLicenseInputState($(this)); });
        // Apply initial multi-search criteria if any fields were pre-filled (e.g., from saved state)
        applyMultiSearch();
      },
      // Callback function for each row to add CSS classes based on issues
      rowCallback: function(row, data) {
        $(row).removeClass('conflict dup-license'); // Remove previous classes
        if (nameConflicts.has(nameKey(data))) $(row).addClass('conflict');
        if (dupLicUsers.has(data.Id)) $(row).addClass('dup-license');
      },
      // Callback function after each table draw to update alerts
      drawCallback: renderAlerts
    });

    // Show/hide loader during table draws
    $(table.table().node()).on('preDraw.dt', function() { showLoader(); });
    $(table.table().node()).on('draw.dt', function() { hideLoader(); });
  }

  // Toggle visibility of text input vs. license select in multi-search rows
  function updateLicenseInputState($row) {
    const $select = $row.find('.column-select');
    const $input = $row.find('.search-input');
    const $licenseSelect = $row.find('.license-select');

    // Column index 6 is 'Licenses'
    if ($select.val() == 6) {
      $input.hide();
      $licenseSelect.show();
    } else {
      $input.show();
      $licenseSelect.hide();
      $input.attr('placeholder', 'Term...');
    }
  }

  // Set up multi-search functionality UI and listeners
  function setupMultiSearch() {
    // Define available columns for multi-search dropdown
    const cols = [
      { v: 0, text: 'ID' }, { v: 1, text: 'Name' }, { v: 2, text: 'Email' },
      { v: 3, text: 'Job Title' }, { v: 4, text: 'Location' }, { v: 5, text: 'Phones' },
      { v: 6, text: 'Licenses' }
    ];
    const $container = $('#multiSearchFields');

    // Function to add a new search field row
    function addSearchField() {
      const $row = $(`<div class="multi-search-row">
        <select class="column-select">
          ${cols.map(c => `<option value="${c.v}">${c.text}</option>`).join('')}
        </select>
        <input class="search-input" placeholder="Term..." />
        <select class="license-select" style="display: none;">
          <option value="">Select a license...</option>
          ${uniqueLicenseNames.map(name => `<option value="${escapeHtml(name)}">${escapeHtml(name)}</option>`).join('')}
        </select>
        <button class="remove-field" title="Remove filter"><i class="fas fa-trash-alt"></i></button>
      </div>`);

      $container.append($row);
      updateLicenseInputState($row); // Set initial input state based on default column select

      // Attach event listeners to the newly added row
      $row.find('.column-select').on('change', function() {
        updateLicenseInputState($row); // Update input based on column change
        // Clear values when column changes
        $row.find('.search-input').val('');
        $row.find('.license-select').val('');
        applyMultiSearch(); // Re-apply search after change
      });

      // Debounced listener for input/change events on search fields
      $row.find('.search-input, .license-select').on('input change', applyMultiSearch);

      // Listener for remove button
      $row.find('.remove-field').on('click', function() {
        $(this).closest('.multi-search-row').remove(); // Remove the row
        applyMultiSearch(); // Re-apply search after removing a filter
      });
    }

    // Attach click listener to 'Add Search Field' button
    $('#addSearchField').off('click').on('click', addSearchField);

    // Attach change listener to the logical operator dropdown (AND/OR)
    $('#multiSearchOperator').off('change').on('change', applyMultiSearch);

    // Add one search field by default if none exist
    if ($container.children().length === 0) { addSearchField(); }

    // Execute the search logic initially to reflect the default state
    // Note: applyMultiSearch uses a debounce, so call the direct function once
    // to ensure the table state and criteria text are updated immediately on setup.
    _executeMultiSearchLogic();
  }

  // The core logic for applying multi-search filters to DataTables
  function _executeMultiSearchLogic() {
    const operator = $('#multiSearchOperator').val(); // 'AND' or 'OR'
    const $searchCriteriaText = $('#searchCriteria'); // Element to display current criteria

    if (!table) {
      $searchCriteriaText.text(allData.length === 0 ? 'No data loaded.' : 'Table not available/initialized.');
      return;
    }

    // Clear previous DataTables searches (global and column searches)
    // Note: Column searches set via `column().search()` are cleared.
    // The multi-search uses `$.fn.dataTable.ext.search`.
    table.search('').columns().search('');

    // Remove all previous custom search functions
    while ($.fn.dataTable.ext.search.length > 0) { $.fn.dataTable.ext.search.pop(); }

    // Collect current filter criteria from the UI
    const rows = $('#multiSearchFields .multi-search-row');
    const filters = [];
    rows.each(function() {
      const colIndex = $(this).find('.column-select').val();
      let searchTerm = (colIndex == 6) ? $(this).find('.license-select').val() : $(this).find('.search-input').val().trim();

      // Only add filters with a non-empty search term
      if (searchTerm) {
        filters.push({ col: parseInt(colIndex, 10), term: searchTerm });
      }
    });

    // Update the criteria text display
    let criteriaText = operator === 'AND' ? 'Criteria: All filters must match (AND)' : 'Criteria: Any filter must match (OR)';

    if (filters.length > 0) {
      criteriaText += ` (${filters.length} active filter(s))`;

      // Add the custom search function to DataTables
      $.fn.dataTable.ext.search.push(
        function(settings, apiData, dataIndex) {
          // Apply this filter only to the correct table instance
          if (settings.nTable.id !== table.table().node().id) return true;

          const rowData = table.row(dataIndex).data(); // Get the original row object
          if (!rowData) return false; // Should not happen if DataTables is working, but safety check

          // Determine the logical function (some for OR, every for AND)
          const logicFn = operator === 'OR' ? filters.some.bind(filters) : filters.every.bind(filters);

          // Test the row against all filters using the chosen logic (AND/OR)
          return logicFn(filter => {
            if (filter.col === 6) { // Special handling for the 'Licenses' column
              // Check if *any* license in the user's Licenses array matches the search term (case-insensitive)
              return (rowData.Licenses && Array.isArray(rowData.Licenses)) ?
                rowData.Licenses.some(l => (l.LicenseName || '').toLowerCase() === filter.term.toLowerCase()) : false;
            }

            // Standard text search for other columns
            // Use the stringified cell value provided by DataTables (`apiData`)
            // DataTables provides a stringified version in the `apiData` array for the standard search mechanism
            const cellValue = apiData[filter.col] || '';
            return cellValue.toString().toLowerCase().includes(filter.term.toLowerCase());
          });
        }
      );
    } else {
      criteriaText = 'Criteria: All results (no filters active)';
    }

    $searchCriteriaText.text(criteriaText); // Update display
    table.draw(); // Redraw the table to apply filters
  }

  // Debounce wrapper for executing multi-search logic
  function applyMultiSearch() {
    clearTimeout(multiSearchDebounceTimer);
    multiSearchDebounceTimer = setTimeout(_executeMultiSearchLogic, DEBOUNCE_DELAY);
  }

  // Clear all filters (individual column inputs and multi-search) and reset column visibility
  $('#clearFilters').on('click', () => {
    if (table) {
      // Clear individual column search inputs
      $(table.table().header()).find('tr:eq(1) th input').val('');
      // Clear DataTables internal column and global searches
      table.search('').columns().search('');
    }

    // Clear multi-search fields UI
    $('#multiSearchFields').empty();
    // Add a default empty search field back if data is loaded
    if (allData.length > 0) { setupMultiSearch(); }
    // Clear multi-search criteria text if no data
    else { $('#searchCriteria').text('No data loaded.'); }

    // Remove all custom search functions used by multi-search
    while ($.fn.dataTable.ext.search.length > 0) { $.fn.dataTable.ext.search.pop(); }

    // Redraw table to show all data (as filters are removed)
    if(table) table.draw();

    // Clear issue alerts
    $('#alertPanel').empty();

    // Reset column visibility to default
    const defaultVisibleCols = [1, 2, 3, 6]; // Name, Email, Job Title, Licenses
    $('#colContainer .col-vis').each(function() {
      const idx = +$(this).data('col');
      const isDefaultVisible = defaultVisibleCols.includes(idx);
      if (table && idx >= 0 && idx < table.columns().nodes().length) {
        try { table.column(idx).visible(isDefaultVisible); } catch(e) { console.warn("Error resetting column visibility:", idx, e); }
      }
      $(this).prop('checked', isDefaultVisible);
    });
  });

  // Function to download a CSV file
  function downloadCsv(csvContent, fileName) {
    // Add BOM (Byte Order Mark) for better compatibility with Excel, especially for UTF-8
    const bom = "\uFEFF";
    const blob = new Blob([bom, csvContent], { type: 'text/csv;charset=utf-8;' });
    const link = document.createElement('a');
    // Use download attribute if supported
    if (link.download !== undefined) {
      const url = URL.createObjectURL(blob);
      link.setAttribute('href', url);
      link.setAttribute('download', fileName);
      link.style.visibility = 'hidden'; // Hide the link
      document.body.appendChild(link); // Append to body
      link.click(); // Simulate click
      document.body.removeChild(link); // Clean up
      URL.revokeObjectURL(url); // Release object URL
    } else {
      // Fallback for browsers that don't support download attribute (less common now)
      alert("Your browser does not support downloading files directly. Please copy the content manually.");
    }
  }

  // Handle CSV export button click
  $('#exportCsv').on('click', () => {
    console.log("[Export CSV] Initiating license data export...");
    if (!table) {
      console.log("[Export CSV] Table not initialized.");
      return alert('Table not initialized. No data loaded.');
    }

    // Get filtered rows from DataTables API (respecting all active filters)
    const rowsToExport = table.rows({ search: 'applied' }).data().toArray();
    console.log("[Export CSV] Number of rows to export (after filters):", rowsToExport.length);

    if (!rowsToExport.length) {
      console.log("[Export CSV] No rows to export with current filters.");
      return alert('No records to export with current filters.');
    }

    // Determine which columns are currently visible and their data properties
    const visibleColumns = [];
    table.columns(':visible').every(function() {
      // Get the column title from the first header row (thead > tr:eq(0))
      const columnTitle = $(table.table().header()).find('tr:eq(0) th').eq(this.index()).text();
      // Get the data property name used in the initial `columns` definition
      const dataProp = table.settings()[0].aoColumns[this.index()].mData;
      // Exclude the hidden ID column unless it's somehow made visible later and needed
      // For a standard export, we usually only want the columns defined in the UI.
      // The original code correctly gets visible columns, which implicitly excludes the initially hidden ID.
      visibleColumns.push({
        title: columnTitle,
        dataProp: dataProp
      });
    });
    console.log("[Export CSV] Visible columns for export:", JSON.stringify(visibleColumns.map(c => c.title))); // Log titles

    // Generate the CSV header row
    const headerRow = visibleColumns.map(col => escapeCsvValue(col.title)).join(',');
    console.log("[Export CSV] Header row:", headerRow);

    // Generate CSV rows from data
    const csvRows = rowsToExport.map(rowData => {
      return visibleColumns.map(col => {
        let cellData = rowData[col.dataProp];
        let shouldForceQuotes = false; // Flag to force quotes around the cell value

        // Special handling for array data (Phones, Licenses)
        if (col.dataProp === 'BusinessPhones') {
          cellData = Array.isArray(cellData) ? cellData.join('; ') : (cellData || '');
          // Force quotes if the joined string contains a semicolon, comma, or newline
          if (String(cellData).match(/[,"\n\r;]/)) {
            shouldForceQuotes = true;
          }
        } else if (col.dataProp === 'Licenses') {
          const licensesArray = (rowData.Licenses && Array.isArray(rowData.Licenses)) ?
                                rowData.Licenses.map(l => l.LicenseName || '').filter(name => name) : [];
          if (licensesArray.length > 0) {
            cellData = licensesArray.join('; ');
            // Force quotes if the joined string contains a semicolon (common separator), comma, or newline
            if (String(cellData).match(/[,"\n\r;]/)) {
              shouldForceQuotes = true;
            }
          } else {
            cellData = '';
          }
        } else {
            // For other columns, just force quotes if it contains standard CSV special chars
             if (String(cellData).match(/[,"\n\r]/)) {
              shouldForceQuotes = true;
            }
        }

        // Use the escapeCsvValue function with the forceQuotes flag
        return escapeCsvValue(cellData, shouldForceQuotes);
      }).join(','); // Join cell values with comma
    });

    // Combine header and data rows
    const csvContent = [headerRow, ...csvRows].join('\n');

    console.log("[Export CSV] Generated CSV content (first 500 chars):", csvContent.substring(0, 500));
    console.log("[Export CSV] Total csvContent length:", csvContent.length);

    // Basic check if content seems empty (only header or nothing)
    if (csvContent.trim() === "" || (csvContent.trim() === headerRow.trim() && csvRows.length === 0 && headerRow.trim() !== "")) {
        console.warn("[Export CSV] csvContent is empty or contains only header (if header has data)!");
        // Consider adding an alert here if it indicates an unexpected state
    } else if (csvContent.trim() === "" && headerRow.trim() === "") {
        console.warn("[Export CSV] csvContent is completely empty (no header, no data)!");
        // Consider adding an alert here if it indicates an unexpected state
    }

    // Trigger download
    downloadCsv(csvContent, 'license_report.csv');
    console.log("[Export CSV] Download requested.");
  });

  // Handle Issues Report export button click
  $('#exportIssues').on('click', () => {
    if (!allData.length) return alert('No data loaded to generate issues report.');

    const lines = []; // Array to hold CSV rows

    // Add Name+Location Conflicts section if any
    if (nameConflicts.size) {
      lines.push(['NAME+LOCATION CONFLICTS']); // Section Header
      lines.push(['Name', 'Location'].map(h => escapeCsvValue(h))); // Column Headers
      // Add rows for each conflict, splitting the stored key
      nameConflicts.forEach(key => lines.push(key.split('|||').map(value => escapeCsvValue(value))));
      lines.push([]); // Add empty line for separation
    }

    // Add Users with Duplicate Licenses section if any
    if (dupLicUsers.size) {
      lines.push(['USERS with Duplicate Licenses']); // Section Header
      lines.push(['Name', 'Location', 'Duplicate Licenses', 'Has Paid License?'].map(h => escapeCsvValue(h))); // Column Headers

      // Filter and process data for users with duplicate licenses
      allData.filter(user => dupLicUsers.has(user.Id)).forEach(user => {
        const licCount = {}, duplicateLicNames = [];
        // Count *all* licenses for this user
        (user.Licenses || []).forEach(l => licCount[l.LicenseName] = (licCount[l.LicenseName] || 0) + 1);

        // Find which specific licenses appear more than once
        Object.entries(licCount).forEach(([licName, count]) => count > 1 && duplicateLicNames.push(licName));

        const joinedDups = duplicateLicNames.join('; '); // Join duplicate license names
        const hasPaid = (user.Licenses || []).some(l => !(l.LicenseName || '').toLowerCase().includes('free')); // Check for paid licenses

        // Add the user's data as a row
        lines.push([
            escapeCsvValue(user.DisplayName),
            escapeCsvValue(user.OfficeLocation),
            // Force quotes around the duplicate licenses list if it contains semicolons or other CSV special characters
            escapeCsvValue(joinedDups, joinedDups.includes(';') || joinedDups.match(/[,"\n\r]/)),
            escapeCsvValue(hasPaid ? 'Yes' : 'No')
        ]);
      });
    }

    // Alert if no issues were found
    if (!lines.length) {
      // Add a line indicating no issues before the alert, in case someone exports an empty file
      lines.push(['No issues detected.']);
      // return alert('No issues detected.'); // Keep the alert for immediate feedback
    }

    // Combine all lines into a single CSV string
    const csvContent = lines.map(rowArray => rowArray.join(',')).join('\n');

    // Trigger download
    downloadCsv(csvContent, 'issues_report.csv');
  });


  // --- Initial Data Loading and Setup ---
  try {
    // Check if 'userData' variable is defined (expected to be embedded in HTML)
    if (typeof userData === 'undefined') {
      throw new Error("The 'userData' variable is not defined. JSON data was not correctly embedded in the HTML.");
    }
    // Validate and normalize the loaded data
    const validatedData = validateJson(userData);

    // Initialize table and features if valid data is found
    if (validatedData.length > 0) {
      initTable(validatedData); // Initialize DataTables
      setupMultiSearch(); // Setup multi-search UI and logic
      // Update initial status text
      $('#searchCriteria').text(`Data loaded. (${validatedData.length} users) Use filters to refine.`);
    } else {
      // Handle case with no valid users
      alert('No valid users found in the data.');
      $('#searchCriteria').text('No valid data loaded.');
    }
  } catch (error) {
    // Handle errors during data processing
    alert('Error processing data: ' + error.message);
    console.error("Error processing data:", error);
    $('#searchCriteria').text('Error loading data.');
  }
});
