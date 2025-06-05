$(document).ready(function() {
    let table = null, allData = [], nameConflicts = new Set(), dupLicUsers = new Set();
    let userMap = new Map();
    const DEBOUNCE_DELAY = 350;
    let multiSearchDebounceTimer;

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
        { index: 3, title: 'Job Title', dataProp: 'JobTitle', useDropdown: false },
        { index: 4, title: 'Manager', dataProp: 'ReportsTo', useDropdown: true },
        { index: 5, title: 'Total Subs.', dataProp: 'TotalSubordinates', useDropdown: false },
        { index: 6, title: 'Location', dataProp: 'OfficeLocation', useDropdown: true },
        { index: 7, title: 'Phones', dataProp: 'BusinessPhones', useDropdown: false },
        { index: 8, title: 'Licenses', dataProp: 'Licenses', useDropdown: true }
    ];

    let uniqueFieldValues = {};
    const MAX_DROPDOWN_OPTIONS_DISPLAYED = 50;

    function showLoader(message = 'Processing...') {
        let $overlay = $('#loadingOverlay');
        if ($overlay.length === 0) {
            $('body').append('<div id="loadingOverlay"><div class="loader-content"><p id="loaderMessageText"></p></div></div>');
            $overlay = $('#loadingOverlay');
        }
        $overlay.find('#loaderMessageText').text(message);
        $overlay.css('display', 'flex');
    }

    function hideLoader() { $('#loadingOverlay').hide(); }
    const nameKey = u => `${u.DisplayName || ''}|||${u.OfficeLocation || ''}`;

    function escapeCsvValue(value) {
        if (value == null) return '';
        let str = String(value);
        if (/[,"\n\r]/.test(str)) {
            str = `"${str.replace(/"/g, '""')}"`;
        }
        return str;
    }

    function downloadCsv(csvContent, fileName) {
        const bom = "\uFEFF";
        const blob = new Blob([bom, csvContent], { type: 'text/csv;charset=utf-8;' });
        const link = document.createElement('a');
        const url = URL.createObjectURL(blob);
        link.setAttribute('href', url);
        link.setAttribute('download', fileName);
        link.style.visibility = 'hidden';
        document.body.appendChild(link);
        link.click();
        document.body.removeChild(link);
        URL.revokeObjectURL(url);
    }

    function renderAlerts() {
        const $alertPanel = $('#alertPanel').empty();
        if (nameConflicts.size) {
            $alertPanel.append(`<div class="alert-badge"><i class="fas fa-users-slash"></i><button id="filterNameConflicts" class="underline-button">Name+Location Conflicts: ${nameConflicts.size}</button></div>`);
        }
        if (dupLicUsers.size) {
            $alertPanel.append(`<div class="alert-badge"><i class="fas fa-copy"></i><span>Duplicate Licenses: ${dupLicUsers.size}</span></div>`);
        }
        
        $('#filterNameConflicts').off('click').on('click', function() {
            if (!table) return;
            showLoader('Filtering...');
            setTimeout(() => {
                const conflictUserIds = allData.filter(u => nameConflicts.has(nameKey(u))).map(u => u.Id);
                table.search('').columns().search('').column(0).search(
                    conflictUserIds.length ? `^(${conflictUserIds.join('|')})$` : '', 
                    true, 
                    false
                ).draw();
            }, 50);
        });
    }

    function initTable(data) {
        allData = data;
        userMap.clear();
        allData.forEach(u => userMap.set(u.Id, u));
        if (table) { table.destroy(); $('#licenseTable').empty(); }

        table = $('#licenseTable').DataTable({
            data: allData, deferRender: true, pageLength: 25, orderCellsTop: true,
            columns: [
                { data: 'Id', visible: false }, { data: 'DisplayName', title: 'Name' }, { data: 'Email', title: 'Email' },
                { data: 'JobTitle', title: 'Job Title' }, { data: 'ReportsTo', title: 'Manager' },
                { 
                    data: 'TotalSubordinates', title: 'Total Subs.',
                    render: (data, type, row) => (type === 'display' && data > 0) ? `<a href="#" class="view-hierarchy" data-userid="${row.Id}" title="Export team of ${row.DisplayName}">${data} <i class="fas fa-file-csv"></i></a>` : data
                },
                { data: 'OfficeLocation', title: 'Location', visible: false },
                { data: 'BusinessPhones', title: 'Phones', visible: false, render: p => Array.isArray(p) ? p.join('; ') : (p || '') },
                { data: 'Licenses', title: 'Licenses', render: l => Array.isArray(l) ? l.map(x => x.LicenseName || '').join(', ') : '' }
            ],
            initComplete: function() {
                const api = this.api();
                $('#colContainer .col-vis').each(function() { const idx = +$(this).data('col'); if (api.column(idx).length) $(this).prop('checked', api.column(idx).visible()); });
                $('.col-vis').off('change').on('change', function() { const col = api.column(+$(this).data('col')); if (col.length) col.visible(!col.visible()); });
                setupMultiSearch(); applyMultiSearch();
            },
            rowCallback: (row, data) => { $(row).removeClass('conflict dup-license').addClass(`${nameConflicts.has(nameKey(data)) ? 'conflict' : ''} ${dupLicUsers.has(data.Id) ? 'dup-license' : ''}`.trim()); },
            drawCallback: () => { renderAlerts(); hideLoader(); }
        });
        $(table.table().node()).on('preDraw.dt', () => showLoader('Updating...'));
    }

    function getAllSubordinates(managerId, localUserMap) {
        const subordinates = [], manager = localUserMap.get(managerId);
        if (!manager || !manager.children) return [];
        const queue = [...manager.children], processed = new Set();
        while (queue.length > 0) {
            const userId = queue.shift();
            if (processed.has(userId)) continue;
            processed.add(userId);
            const user = localUserMap.get(userId);
            if (user) {
                subordinates.push(user);
                if (user.children && user.children.length > 0) queue.push(...user.children);
            }
        }
        return subordinates;
    }

    function setupHierarchyExport() {
        $('#licenseTable tbody').on('click', 'a.view-hierarchy', function(e) {
            e.preventDefault();
            const managerId = $(this).data('userid');
            const manager = userMap.get(managerId);
            if (!manager) return;

            showLoader(`Exporting team for ${manager.DisplayName}...`);
            setTimeout(() => {
                const subordinatesData = getAllSubordinates(managerId, userMap);
                if (subordinatesData.length === 0) {
                    hideLoader();
                    alert("This manager has no subordinates to export.");
                    return;
                }

                const headers = ["Name", "Email", "Job Title", "Direct Manager", "Location", "Licenses", "Top-Level Manager"];
                
                const csvRows = subordinatesData.map(user => {
                    const licenses = (user.Licenses || []).map(l => l.LicenseName).join(', ');
                    const rowData = [
                        user.DisplayName, user.Email, user.JobTitle,
                        user.ReportsTo, user.OfficeLocation, licenses,
                        manager.DisplayName
                    ];
                    return rowData.map(escapeCsvValue).join(',');
                });

                const csvContent = [headers.join(','), ...csvRows].join('\n');
                const fileName = `Team_${manager.DisplayName.replace(/\s/g, '_')}.csv`;
                
                downloadCsv(csvContent, fileName);
                hideLoader();
            }, 50);
        });
    }

    function processDataWithWorker(rawData) {
        showLoader('Processing data and building hierarchy...');
        const workerScript = `const nameKeyInternal=u=>\`\${u.DisplayName||""}||\${u.OfficeLocation||""}\`;function validateAndNormalize(t){return t.map((t,e)=>t&&"object"==typeof t?{Id:t.Id||\`unknown_\${e}\`,DisplayName:t.DisplayName||"Unknown",OfficeLocation:t.OfficeLocation||"Unknown",Email:t.Email||"",JobTitle:t.JobTitle||"Unknown",ReportsTo:t.ReportsTo||null,BusinessPhones:Array.isArray(t.BusinessPhones)?t.BusinessPhones:[],Licenses:Array.isArray(t.Licenses)?t.Licenses.map(t=>({LicenseName:t.LicenseName||"",SkuId:t.SkuId||""})):[],children:[],TotalSubordinates:0}:null).filter(Boolean)}function findIssues(t){const e=new Map,n=new Set;return t.forEach(t=>{e.set(nameKeyInternal(t),(e.get(nameKeyInternal(t))||0)+1);const r=new Map;(t.Licenses||[]).forEach(t=>{t.LicenseName&&r.set(t.LicenseName,(r.get(t.LicenseName)||0)+1)}),[...r.values()].some(t=>t>1)&&n.add(t.Id)}),{nameConflictsArray:[...e].filter(([t,e])=>e>1).map(([t])=>t),dupLicUsersArray:Array.from(n)}}function countAllSubordinates(t,e){const n=e.get(t);if(!n||!n.children||0===n.children.length)return 0;let r=0;const a=[...n.children],o=new Set;for(;a.length>0;){const t=a.shift();o.has(t)||o.add(t,r++,e.get(t)?.children?.length>0&&a.push(...e.get(t).children))}return r}function buildHierarchyAndCount(t){const e=new Map;t.forEach(t=>e.set(t.DisplayName,t));const n=new Map;t.forEach(t=>n.set(t.Id,t)),t.forEach(t=>{if(t.ReportsTo){const r=e.get(t.ReportsTo);r&&r.children.push(t.Id)}}),t.forEach(t=>{t.TotalSubordinates=countAllSubordinates(t.Id,n)})}function getUniqueFieldValues(t,e){const n={};return e.forEach(e=>{e.useDropdown&&("Licenses"===e.dataProp?n.Licenses=[...new Set(t.flatMap(t=>t.Licenses||[]).map(t=>t.LicenseName).filter(Boolean))].sort():n[e.dataProp]=[...new Set(t.map(t=>t[e.dataProp]).filter(Boolean))].sort())}),n}self.onmessage=function(t){const{rawData:e,searchableColumnsConfig:n}=t.data;try{const t=validateAndNormalize(e),{nameConflictsArray:r,dupLicUsersArray:a}=findIssues(t);buildHierarchyAndCount(t);const o=getUniqueFieldValues(t,n);self.postMessage({validatedData:t,nameConflictsArray:r,dupLicUsersArray:a,uniqueFieldValues:o,error:null})}catch(t){self.postMessage({error:"Error in Web Worker: "+t.message})}finally{self.close()}};`;
        const blob = new Blob([workerScript], { type: 'application/javascript' });
        const worker = new Worker(URL.createObjectURL(blob));

        worker.onmessage = function(e) {
            URL.revokeObjectURL(blob);
            const { validatedData, nameConflictsArray, dupLicUsersArray, uniqueFieldValues: uFValues, error } = e.data;
            if (error) { hideLoader(); alert(error); return; }
            nameConflicts = new Set(nameConflictsArray); dupLicUsers = new Set(dupLicUsersArray); uniqueFieldValues = uFValues;
            if (validatedData && validatedData.length > 0) {
                showLoader('Rendering table...');
                setTimeout(() => { initTable(validatedData); setupHierarchyExport(); }, 50);
            } else { hideLoader(); alert('No valid user data found.'); }
        };
        worker.onerror = function(e) { URL.revokeObjectURL(blob); hideLoader(); console.error(`Error in Web Worker: Line ${e.lineno} in ${e.filename}: ${e.message}`); alert('A critical error occurred during data processing. Please check the console.'); };
        worker.postMessage({ rawData, searchableColumnsConfig });
    }

    try {
        if (typeof userData !== 'undefined' && Array.isArray(userData) && userData.length > 0) {
            processDataWithWorker(userData);
        } else { hideLoader(); $('#searchCriteria').text('No user data found to load.'); }
    } catch (error) { hideLoader(); alert('Error initiating data loading: ' + error.message); console.error("Initial loading error:", error); }
    
    function setupMultiSearch() {
        const $container = $('#multiSearchFields');
        function addSearchField() {
            const columnOptions = searchableColumnsConfig.map(c => `<option value="${c.index}">${c.title}</option>`).join('');
            const $row = $(`<div class="multi-search-row"><select class="form-control column-select">${columnOptions}</select><select class="form-control condition-operator-select"></select><div class="custom-dropdown-container"><input type="text" class="form-control custom-dropdown-text-input" autocomplete="off" /></div><input class="form-control search-input" placeholder="Search term..." /><input type="hidden" class="search-value-input" /><button class="remove-field" title="Remove filter"><i class="fas fa-trash-alt"></i></button></div>`);
            $container.append($row); updateSearchFieldUI($row);
            $row.find('.column-select').on('change', function() { updateSearchFieldUI($row); $row.find('.search-input, .custom-dropdown-text-input, .search-value-input').val('').trigger('change'); });
            $row.find('.condition-operator-select, .search-value-input').on('change', applyMultiSearch);
            $row.find('.search-input').on('input change', applyMultiSearch);
            $row.find('.remove-field').on('click', function() { $(this).closest('.multi-search-row').remove(); applyMultiSearch(); });
        }
        $('#addSearchField').off('click').on('click', addSearchField);
        $('#multiSearchOperator').off('change').on('change', applyMultiSearch);
        if ($container.children().length === 0 && allData.length > 0) addSearchField();
    }
    
    function updateSearchFieldUI($row) {
        const selIdx = $row.find('.column-select').val(), conf = searchableColumnsConfig.find(c => c.index == selIdx),
        $si = $row.find('.search-input'), $cdc = $row.find('.custom-dropdown-container'), $cdti = $row.find('.custom-dropdown-text-input'),
        $col = $row.find('.custom-options-list'), $hvi = $row.find('.search-value-input'), $cos = $row.find('.condition-operator-select');
        $cdti.off(); $col.off(); $cos.empty();
        $cdc.hide().find('.custom-options-list').remove();
        $si.hide();
        if (conf) {
            if (conf.useDropdown) {
                $cdc.show();
                if($cdc.find('.custom-options-list').length === 0) {
                     $cdc.append('<div class="custom-options-list"></div>');
                }
                $cdti.attr('placeholder', `Select ${conf.title.toLowerCase()}`);
                $cos.append(`<option value="IS" selected>${operatorTypes.IS}</option><option value="IS_NOT">${operatorTypes.IS_NOT}</option>`);
                const opts = uniqueFieldValues[conf.dataProp] || []; let fd;
                $cdti.on('input', function() {
                    clearTimeout(fd); const $i = $(this); fd = setTimeout(() => {
                        const st = $i.val().toLowerCase(); 
                        const $list = $i.next('.custom-options-list');
                        $list.empty().show(); 
                        const fOpts = opts.filter(o => String(o).toLowerCase().includes(st));
                        if (fOpts.length === 0) $list.append('<div class="custom-option-item no-results">No results</div>');
                        else { fOpts.slice(0, 50).forEach(o => $list.append($('<div class="custom-option-item"></div>').text(o).data('value', o))); if (fOpts.length > 50) $list.append(`<div class="custom-option-item no-results">...and ${fOpts.length - 50} more</div>`); }
                    }, 200);
                }).on('focus', function() { $(this).trigger('input'); }).on('blur', function() { setTimeout(() => $(this).next('.custom-options-list').hide(), 150); });
                $cdc.on('mousedown', '.custom-option-item', function(e) { e.preventDefault(); if ($(this).hasClass('no-results')) return; $cdti.val($(this).text()); $hvi.val($(this).data('value')).trigger('change'); $(this).parent().hide(); });
            } else {
                $si.show(); $si.attr('placeholder', 'Search term...');
                $cos.append(`<option value="CONTAINS" selected>${operatorTypes.CONTAINS}</option><option value="DOES_NOT_CONTAIN">${operatorTypes.DOES_NOT_CONTAIN}</option><option value="IS">${operatorTypes.IS}</option><option value="IS_NOT">${operatorTypes.IS_NOT}</option>`);
            }
        }
    }
    
    function applyMultiSearch() { clearTimeout(multiSearchDebounceTimer); multiSearchDebounceTimer = setTimeout(_executeMultiSearchLogic, DEBOUNCE_DELAY); }
    
    function _executeMultiSearchLogic() {
        if (!table) return;
        const op = $('#multiSearchOperator').val(); while ($.fn.dataTable.ext.search.length > 0) $.fn.dataTable.ext.search.pop();
        const filters = [];
        $('#multiSearchFields .multi-search-row').each(function() {
            const $r = $(this), cIdx = $r.find('.column-select').val(), cConf = searchableColumnsConfig.find(c => c.index == cIdx);
            if (cConf) { const term = cConf.useDropdown ? $r.find('.search-value-input').val() : $r.find('.search-input').val().trim(); if (term !== '') filters.push({ term, dataProp: cConf.dataProp, isDropdown: cConf.useDropdown, condition: $r.find('.condition-operator-select').val() }); }
        });
        if (filters.length > 0) {
            $.fn.dataTable.ext.search.push(function(settings, data, dataIndex) {
                if (settings.nTable.id !== table.table().node().id) return true;
                const rowData = table.row(dataIndex).data(); if (!rowData) return false;
                return (op === 'OR' ? filters.some.bind(filters) : filters.every.bind(filters))(f => {
                    const ft = String(f.term || '').toLowerCase(), cd = String(rowData[f.dataProp] || '').toLowerCase();
                    if (f.dataProp === 'Licenses') { const ul = (rowData.Licenses || []).map(l => (l.LicenseName || '').toLowerCase()), m = ul.includes(ft); return f.condition === 'IS' ? m : !m; }
                    else if (f.isDropdown) { return f.condition === 'IS' ? cd === ft : cd !== ft; }
                    else { if (f.condition === 'CONTAINS') return cd.includes(ft); if (f.condition === 'DOES_NOT_CONTAIN') return !cd.includes(ft); if (f.condition === 'IS') return cd === ft; if (f.condition === 'IS_NOT') return cd !== ft; }
                    return false;
                });
            });
        } table.draw();
    }
    
    $('#clearFilters').on('click', () => { if(table) table.search('').columns().search(''); $('#multiSearchFields').empty(); while ($.fn.dataTable.ext.search.length > 0) $.fn.dataTable.ext.search.pop(); if (allData.length > 0) setupMultiSearch(); if(table) table.draw(); });
    
    $('#exportCsv').on('click', () => {
        if (!table) return;
        showLoader('Exporting...');
        setTimeout(() => {
            const rows = table.rows({ search: 'applied' }).data().toArray();
            if (!rows.length) { hideLoader(); return; }
            const headers = ["Name","Email","Job Title","Manager","Total Subs.","Location","Phones","Licenses"];
            const csvRows = rows.map(user => {
                const licenses = (user.Licenses || []).map(l => l.LicenseName).join(', ');
                const rowData = [
                    user.DisplayName, user.Email, user.JobTitle, user.ReportsTo,
                    user.TotalSubordinates, user.OfficeLocation, 
                    Array.isArray(user.BusinessPhones) ? user.BusinessPhones.join('; ') : (user.BusinessPhones || ''),
                    licenses
                ];
                return rowData.map(escapeCsvValue).join(',');
            });
            downloadCsv([headers.join(','), ...csvRows].join('\n'), 'license_report.csv');
            hideLoader();
        }, 50);
    });

    $('#exportIssues').on('click', () => {
        if (!allData.length) return;
        showLoader('Generating report...');
        setTimeout(() => {
            const lines = [];
            if (nameConflicts.size) {
                lines.push(['NAME+LOCATION CONFLICTS']);
                lines.push(['Name', 'Location']);
                nameConflicts.forEach(k => lines.push(k.split('|||')));
                lines.push([]);
            }
            if (dupLicUsers.size) {
                lines.push(['USERS WITH DUPLICATE LICENSES']);
                lines.push(['Name', 'Location', 'Duplicate Licenses']);
                allData.filter(u => dupLicUsers.has(u.Id)).forEach(u => {
                    const c = {}, d = [];
                    (u.Licenses || []).forEach(li => c[li.LicenseName] = (c[li.LicenseName] || 0) + 1);
                    Object.entries(c).forEach(([n, ct]) => ct > 1 && d.push(n));
                    lines.push([u.DisplayName, u.OfficeLocation, d.join('; ')]);
                });
            }
            if (!lines.length) { lines.push(['No issues detected.']); }
            const csvContent = lines.map(row => row.map(escapeCsvValue).join(',')).join('\n');
            downloadCsv(csvContent, 'issues_report.csv');
            hideLoader();
        }, 50);
    });
});
