$(document).ready(function() {
  let table = null, allData = [], nameConflicts = new Set(), dupLicUsers = new Set();
  // let uniqueLicenseNames = []; // Replaced by uniqueFieldValues.Licenses

  const DEBOUNCE_DELAY = 350;
  let multiSearchDebounceTimer;

  // Central configuration for multi-searchable columns
  // index: The column index in DataTables
  // title: Display name in the dropdown
  // dataProp: The property name in the `allData` objects
  // useDropdown: Boolean, true if this column should use a select dropdown for searching
  const searchableColumnsConfig = [
    { index: 0, title: 'ID', dataProp: 'Id', useDropdown: false },
    { index: 1, title: 'Name', dataProp: 'DisplayName', useDropdown: false },
    { index: 2, title: 'Email', dataProp: 'Email', useDropdown: false },
    { index: 3, title: 'Job Title', dataProp: 'JobTitle', useDropdown: true },
    { index: 4, title: 'Location', dataProp: 'OfficeLocation', useDropdown: true },
    { index: 5, title: 'Phones', dataProp: 'BusinessPhones', useDropdown: false },
    { index: 6, title: 'Licenses', dataProp: 'Licenses', useDropdown: true } // Special handling for this dataProp
  ];

  let uniqueFieldValues = {}; // Stores unique values for dropdown-enabled fields

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
      OfficeLocation: u.OfficeLocation || 'Unknown', // Ensure this exists
      Email: u.Email || '',
      JobTitle: u.JobTitle || 'Unknown', // Ensure this exists
      BusinessPhones: Array.isArray(u.BusinessPhones) ? u.BusinessPhones : (typeof u.BusinessPhones === 'string' ? u.BusinessPhones.split('; ').filter(p => p) : []),
      Licenses: Array.isArray(u.Licenses) ? u.Licenses.map(l => ({
        LicenseName: l.LicenseName || `Lic_${i}_${Math.random().toString(36).substr(2, 5)}`,
        SkuId: l.SkuId || ''
      })).filter(l => l.LicenseName) : []
    } : null).filter(x => x);
  }

  // Find data issues (Name+Location conflicts, duplicate Office licenses)
  function findIssues(data) {
    const nameMap = new Map();
    const dupSet = new Set();
    const officeLic = new Set([
      'Microsoft 365 E3', 'Microsoft 365 E5',
      'Microsoft 365 Business Standard', 'Microsoft 365 Business Premium',
      'Office 365 E3', 'Office 365 E5'
    ]);

    data.forEach(u => {
      const k = nameKey(u);
      nameMap.set(k, (nameMap.get(k) || 0) + 1);
      const licCount = new Map();
      (u.Licenses || []).forEach(l => {
        const baseName = (l.LicenseName || '').match(/^(Microsoft 365|Office 365)/)?.[0] ?
                         (l.LicenseName.match(/^(Microsoft 365 E3|Microsoft 365 E5|Microsoft 365 Business Standard|Microsoft 365 Business Premium|Office 365 E3|Office 365 E5)/)?.[0] || l.LicenseName)
                        : l.LicenseName;
        if (baseName) {
          licCount.set(baseName, (licCount.get(baseName) || 0) + 1);
        }
      });
      if ([...licCount].some(([lic, c]) => officeLic.has(lic) && c > 1)) {
        dupSet.add(u.Id);
      }
    });
    const conflictingNameKeys = new Set([...nameMap].filter(([, count]) => count > 1).map(([key]) => key));
    return { nameConflicts: conflictingNameKeys, dupLicUsers: dupSet };
  }

  // Render alert badges for detected issues
  function renderAlerts() {
    const $p = $('#alertPanel').empty();
    if (nameConflicts.size) {
      $p.append(`<div class="alert-badge">
        <button id="filterNameConflicts" style="background: none; border: none; color: inherit; cursor: pointer; text-decoration: underline;" class="underline"><i class="fas fa-users-slash" style="margin-right: 0.3rem;"></i>Name+Location Conflicts: ${nameConflicts.size}</button>
      </div>`);
    }
    if (dupLicUsers.size) {
      const list = allData.filter(u => dupLicUsers.has(u.Id)).map(u => {
        const licCount = {}, duplicateLicNames = [];
        (u.Licenses || []).forEach(l => {
            licCount[l.LicenseName] = (licCount[l.LicenseName] || 0) + 1
        });
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
    $('#filterNameConflicts').off('click').on('click', function() {
      if (!table) return;
      const conflictUserIds = allData.filter(u => nameConflicts.has(nameKey(u))).map(u => u.Id);
      table.search('').columns().search('');
      table.column(0).search('^(' + conflictUserIds.join('|') + ')$', true, false).draw();
    });
    $('.toggle-details').off('click').on('click', function() {
      const targetId = $(this).data('target');
      $(`#${targetId}`).toggleClass('show');
      $(this).text($(`#${targetId}`).hasClass('show') ? 'Hide' : 'Details');
    });
  }

  // Initialize the DataTables table
  function initTable(data) {
    allData = data;
    ({ nameConflicts, dupLicUsers } = findIssues(allData));

    // Initialize uniqueFieldValues storage
    uniqueFieldValues = {};
    searchableColumnsConfig.forEach(colConfig => {
      if (colConfig.useDropdown) {
        if (colConfig.dataProp === 'Licenses') {
          const allLicenseObjects = allData.flatMap(user => user.Licenses || []).filter(l => l.LicenseName);
          uniqueFieldValues.Licenses = [...new Set(allLicenseObjects.map(l => l.LicenseName))].sort();
        } else {
          uniqueFieldValues[colConfig.dataProp] = [...new Set(allData.map(user => user[colConfig.dataProp]).filter(value => value && String(value).trim() !== ''))].sort();
        }
      }
    });

    // Populate the datalist for license name input suggestions (if used elsewhere, e.g. a general license search input)
    // This specific datalist might become less relevant if all license searching is through the multi-search select.
    if ($('#licenseDatalist').length === 0) {
      $('body').append('<datalist id="licenseDatalist"></datalist>');
    }
    const $licenseDatalist = $('#licenseDatalist').empty();
    if (uniqueFieldValues.Licenses) {
      uniqueFieldValues.Licenses.forEach(name => {
        $licenseDatalist.append($('<option>').attr('value', escapeHtml(name)));
      });
    }


    if (table) {
      $(table.table().node()).off('preDraw.dt draw.dt');
      table.destroy();
    }

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
        {
          data: 'BusinessPhones', title: 'Phones', visible: false,
          render: p => Array.isArray(p) ? p.join('; ') : (p || '')
        },
        {
          data: 'Licenses', title: 'Licenses', visible: true,
          render: l => Array.isArray(l) ? l.map(x => x.LicenseName || '').filter(name => name).join(', ') : ''
        }
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
          try {
            if (idx >= 0 && idx < api.columns().nodes().length) {
              $(this).prop('checked', api.column(idx).visible());
            } else {
              $(this).prop('disabled', true);
            }
          } catch (e) { console.warn("Error checking column visibility:", idx, e); }
        });
        $('.col-vis').off('change').on('change', function() {
          const idx = +$(this).data('col');
          try {
            const col = api.column(idx);
            if (col && col.visible) {
              col.visible(!col.visible());
            } else {
              console.warn("Column object not found or visible method missing for index:", idx);
            }
          } catch (e) { console.warn("Error toggling column visibility:", idx, e); }
        });
        $('#multiSearchFields .multi-search-row').each(function() { updateSearchFieldUI($(this)); });
        applyMultiSearch();
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

  // Toggle visibility of text input vs. select dropdown in multi-search rows
  function updateSearchFieldUI($row) {
    const selectedColIndex = $row.find('.column-select').val();
    const columnConfig = searchableColumnsConfig.find(c => c.index == selectedColIndex);

    const $textInput = $row.find('.search-input');
    const $dropdownInput = $row.find('.search-value-select');

    if (columnConfig && columnConfig.useDropdown) {
      $textInput.hide();
      $dropdownInput.show().empty(); // Clear previous options
      $dropdownInput.append($('<option>').val('').text(`Select a ${columnConfig.title}...`));

      const values = uniqueFieldValues[columnConfig.dataProp] || [];
      values.forEach(val => {
        $dropdownInput.append($('<option>').val(escapeHtml(val)).text(escapeHtml(val)));
      });
    } else {
      $textInput.show();
      $dropdownInput.hide().empty();
      $textInput.attr('placeholder', 'Term...');
    }
  }

  // Set up multi-search functionality UI and listeners
  function setupMultiSearch() {
    const $container = $('#multiSearchFields');

    function addSearchField() {
      const columnOptions = searchableColumnsConfig
        .map(c => `<option value="${c.index}">${escapeHtml(c.title)}</option>`)
        .join('');

      const $row = $(`<div class="multi-search-row">
        <select class="column-select">${columnOptions}</select>
        <input class="search-input" placeholder="Term..." />
        <select class="search-value-select" style="display: none;"></select> 
        <button class="remove-field" title="Remove filter"><i class="fas fa-trash-alt"></i></button>
      </div>`);

      $container.append($row);
      updateSearchFieldUI($row); // Set initial input state

      $row.find('.column-select').on('change', function() {
        updateSearchFieldUI($row);
        $row.find('.search-input').val('');
        $row.find('.search-value-select').val('');
        applyMultiSearch();
      });

      $row.find('.search-input, .search-value-select').on('input change', applyMultiSearch);

      $row.find('.remove-field').on('click', function() {
        $(this).closest('.multi-search-row').remove();
        applyMultiSearch();
      });
    }

    $('#addSearchField').off('click').on('click', addSearchField);
    $('#multiSearchOperator').off('change').on('change', applyMultiSearch);

    if ($container.children().length === 0) { addSearchField(); }
    _executeMultiSearchLogic(); // Apply initial state
  }

  // The core logic for applying multi-search filters to DataTables
  function _executeMultiSearchLogic() {
    const operator = $('#multiSearchOperator').val();
    const $searchCriteriaText = $('#searchCriteria');

    if (!table) {
      $searchCriteriaText.text(allData.length === 0 ? 'No data loaded.' : 'Table not available/initialized.');
      return;
    }

    table.search('').columns().search(''); // Clear existing searches
    while ($.fn.dataTable.ext.search.length > 0) { $.fn.dataTable.ext.search.pop(); }

    const filters = [];
    $('#multiSearchFields .multi-search-row').each(function() {
      const colIndex = $(this).find('.column-select').val();
      const columnConfig = searchableColumnsConfig.find(c => c.index == colIndex);
      let searchTerm = '';

      if (columnConfig) {
        searchTerm = columnConfig.useDropdown ?
                     $(this).find('.search-value-select').val() :
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

    let criteriaText = operator === 'AND' ? 'Criteria: All filters must match (AND)' : 'Criteria: Any filter must match (OR)';
    if (filters.length > 0) {
      criteriaText += ` (${filters.length} active filter(s))`;
      $.fn.dataTable.ext.search.push(
        function(settings, apiData, dataIndex) {
          if (settings.nTable.id !== table.table().node().id) return true;
          const rowData = table.row(dataIndex).data();
          if (!rowData) return false;

          const logicFn = operator === 'OR' ? filters.some.bind(filters) : filters.every.bind(filters);

          return logicFn(filter => {
            if (filter.dataProp === 'Licenses') { // Special handling for Licenses
              return (rowData.Licenses && Array.isArray(rowData.Licenses)) ?
                rowData.Licenses.some(l => (l.LicenseName || '').toLowerCase() === filter.term.toLowerCase()) : false;
            } else if (filter.isDropdown) { // Exact match for other dropdowns
              const cellValue = rowData[filter.dataProp] || '';
              return cellValue.toString().toLowerCase() === filter.term.toLowerCase();
            } else { // Standard text search (contains) for non-dropdown fields
              const cellValue = apiData[filter.col] || ''; // Use DataTables prepared string data
              return cellValue.toString().toLowerCase().includes(filter.term.toLowerCase());
            }
          });
        }
      );
    } else {
      criteriaText = 'Criteria: All results (no filters active)';
    }
    $searchCriteriaText.text(criteriaText);
    table.draw();
  }

  // Debounce wrapper for executing multi-search logic
  function applyMultiSearch() {
    clearTimeout(multiSearchDebounceTimer);
    multiSearchDebounceTimer = setTimeout(_executeMultiSearchLogic, DEBOUNCE_DELAY);
  }

  // Clear all filters
  $('#clearFilters').on('click', () => {
    if (table) {
      $(table.table().header()).find('tr:eq(1) th input').val('');
      table.search('').columns().search('');
    }
    $('#multiSearchFields').empty();
    if (allData.length > 0) { setupMultiSearch(); } // Re-adds one default field
    else { $('#searchCriteria').text('No data loaded.'); }

    while ($.fn.dataTable.ext.search.length > 0) { $.fn.dataTable.ext.search.pop(); }
    if(table) table.draw();
    $('#alertPanel').empty();

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
      alert("Your browser does not support downloading files directly.");
    }
  }

  // Handle CSV export button click
  $('#exportCsv').on('click', () => {
    if (!table) return alert('Table not initialized. No data loaded.');
    const rowsToExport = table.rows({ search: 'applied' }).data().toArray();
    if (!rowsToExport.length) return alert('No records to export with current filters.');

    const visibleColumns = [];
    table.columns(':visible').every(function() {
      const columnTitle = $(table.table().header()).find('tr:eq(0) th').eq(this.index()).text();
      const dataProp = table.settings()[0].aoColumns[this.index()].mData;
      visibleColumns.push({ title: columnTitle, dataProp: dataProp });
    });

    const headerRow = visibleColumns.map(col => escapeCsvValue(col.title)).join(',');
    const csvRows = rowsToExport.map(rowData => {
      return visibleColumns.map(col => {
        let cellData = rowData[col.dataProp];
        let shouldForceQuotes = false;
        if (col.dataProp === 'BusinessPhones') {
          cellData = Array.isArray(cellData) ? cellData.join('; ') : (cellData || '');
          if (String(cellData).match(/[,"\n\r;]/)) shouldForceQuotes = true;
        } else if (col.dataProp === 'Licenses') {
          const licensesArray = (rowData.Licenses && Array.isArray(rowData.Licenses)) ?
                                rowData.Licenses.map(l => l.LicenseName || '').filter(name => name) : [];
          cellData = licensesArray.length > 0 ? licensesArray.join('; ') : '';
          if (String(cellData).match(/[,"\n\r;]/)) shouldForceQuotes = true;
        } else {
          if (String(cellData).match(/[,"\n\r]/)) shouldForceQuotes = true;
        }
        return escapeCsvValue(cellData, shouldForceQuotes);
      }).join(',');
    });
    const csvContent = [headerRow, ...csvRows].join('\n');
    downloadCsv(csvContent, 'license_report.csv');
  });

  // Handle Issues Report export button click
  $('#exportIssues').on('click', () => {
    if (!allData.length) return alert('No data loaded to generate issues report.');
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
          escapeCsvValue(user.DisplayName),
          escapeCsvValue(user.OfficeLocation),
          escapeCsvValue(joinedDups, joinedDups.includes(';') || joinedDups.match(/[,"\n\r]/)),
          escapeCsvValue(hasPaid ? 'Yes' : 'No')
        ]);
      });
    }
    if (!lines.length) {
        lines.push(['No issues detected.']);
    }
    const csvContent = lines.map(rowArray => rowArray.join(',')).join('\n');
    downloadCsv(csvContent, 'issues_report.csv');
  });

  // --- Initial Data Loading and Setup ---
  try {
    if (typeof userData === 'undefined') {
      throw new Error("The 'userData' variable is not defined. JSON data was not correctly embedded in the HTML.");
    }
    const validatedData = validateJson(userData);
    if (validatedData.length > 0) {
      initTable(validatedData);
      setupMultiSearch(); // Must be called after initTable populates uniqueFieldValues
      $('#searchCriteria').text(`Data loaded. (${validatedData.length} users) Use filters to refine.`);
    } else {
      alert('No valid users found in the data.');
      $('#searchCriteria').text('No valid data loaded.');
    }
  } catch (error) {
    alert('Error processing data: ' + error.message);
    console.error("Error processing data:", error);
    $('#searchCriteria').text('Error loading data.');
  }
});
