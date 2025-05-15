try {
    Connect-MgGraph -Scopes User.Read.All -NoWelcome
    Write-Host "Successfully connected to Microsoft Graph." -ForegroundColor Green
}
catch {
    Write-Error "Failed to connect to Microsoft Graph: $($_.Exception.Message)"
    exit 1
}

# Map to translate SkuIDs into understandable names
$skuIdMap = @{
    "3a349c99-ffec-43d2-a2e8-6b97fcb71103"="Bullet Chart Premium Tier 1 (1-19 users)"
    "29fcd665-d8d1-4f34-8eed-3811e3fca7b3"="Dynamics 365 Customer Insights Self-Service"
    "1e615a51-59db-4807-9957-aa83c3657351"="Dynamics 365 Customer Service Enterprise vTrial"
    "bc946dac-7877-4271-b2f7-99d2db13cd2c"="Dynamics 365 Customer Voice Trial"
    "6ec92958-3cc1-49db-95bd-bc6b3798df71"="Dynamics 365 Sales Premium Viral Trial"
    "b05e124f-c7cc-45a0-a6aa-8cf78c946968"="Enterprise Mobility + Security E5"
    "ee02fd1b-340e-4a4b-b355-4a514e4c8943"="Exchange Online Archiving for Exchange Online"
    "0c266dff-15dd-4b49-8397-2bb16070ed52"="Microsoft 365 Audio Conferencing"
    "639dec6b-bb19-468b-871c-c5c441c4b0cb"="Microsoft 365 Copilot"
    "7792674b-fa0c-4af5-b2a1-a15239f933b6"="Microsoft 365 Copilot for Finance (Preview)"
    "05e9a617-0261-4cee-bb44-138d3ef5d965"="Microsoft 365 E3"
    "2bc9d149-a1dc-4d8f-bcd8-e9c5750a59b5"="Microsoft 365 E5 Information Protection and Governance"
    "66b55226-6b4f-492c-910c-a3b7a3c9d993"="Microsoft 365 F3"
    "606b54a9-78d8-4298-ad8b-df6ef4481c80"="Microsoft Copilot Studio Viral Trial"
    "84a661c4-e949-4bd2-a560-ed7766fcaf2b"="Microsoft Entra ID P2"
    "a403ebcc-fae0-4ca2-8c8c-7a907fd6c235"="Microsoft Fabric (Free)"
    "a929cd4d-8672-47c9-8664-159c1f322ba8"="Microsoft Intune Suite"
    "5b631642-bd26-49fe-bd20-1daaa972ef80"="Microsoft Power Apps for Developer"
    "dcb1a3ae-b33f-4487-846a-a640262fadf4"="Microsoft Power Apps Plan 2 Trial"
    "f30db892-07e9-47e9-837c-80727f46fd3d"="Microsoft Power Automate Free"
    "1f2f344a-700d-42c9-9427-5cea1d5d7ba6"="Microsoft Stream Trial"
    "d3b4fe1f-9992-4930-8acb-ca6ec609365e"="Microsoft Teams Domestic and International Calling Plan"
    "440eaaa8-b3e0-484b-a8be-62870b9ba70a"="Microsoft Teams Phone Resource Account"
    "e43b5b99-8dfb-405f-9987-dc307f34bcbd"="Microsoft Teams Phone Standard"
    "6070a4c8-34c6-4937-8dfb-39bbc6397a60"="Microsoft Teams Rooms Standard"
    "3d957427-ecdc-4df2-aacd-01cc9d519da8"="Microsoft Viva Insights"
    "18181a46-0d4e-45cd-891e-60aabd171b4e"="Office 365 E1"
    "4b585984-651b-448a-9e53-3b10f069cf7f"="Office 365 F3"
    "53818b1b-4a27-454b-8896-0dba576410e6"="Planner and Project Plan 3"
    "beb6439c-caad-48d3-bf46-0c82871e12be"="Planner Plan 1"
    "b30411f5-fea1-4a59-9ad9-3db7c7ead579"="Power Apps Premium"
    "4a51bf65-409c-4a91-b845-1121b571cc9d"="Power Automate per user plan"
    "f8a1db68-be16-40ed-86d5-cb42ce701560"="Power BI Pro"
    "bf666882-9c9b-4b2e-aa2f-4789b0a52ba2"="PowerApps per app baseline access"
    "776df282-9fc0-4862-99e2-70e561b9909e"="Project Online Essentials"
    "8c4ce438-32a7-4ac5-91a6-e22ae08d9c8b"="Rights Management Adhoc"
    "52ea0e27-ae73-4983-a08f-13561efdb823"="Teams Premium (for Departments)"
    "4b244418-9658-4451-a2b8-b5e2b364e9bd"="Visio Plan 1"
    "c5928f49-12ba-48f7-ada3-0d743a3601d5"="Visio Plan 2"
}

# Define concurrent task limit
$throttleLimit = 5 
$jobs = @()
$allEntries = @() # Stores results from all jobs

# Iterate over each SkuID to fetch users
foreach ($skuId in $skuIdMap.Keys) {
    $skuName = $skuIdMap[$skuId]
    
    # Throttle job creation
    while (($jobs | Where-Object State -eq 'Running').Count -ge $throttleLimit) {
        Start-Sleep -Seconds 1
        $jobs = Get-Job # Refresh job list
    }

    $jobs += Start-ThreadJob -Name "GetUsers_$($skuName -replace '\W','_')" -ScriptBlock {
        param($skuIdParam, $skuNameParam)
        
        $mgUserParams = @{
            Filter           = "assignedLicenses/any(u:u/skuId eq $($skuIdParam))"
            ConsistencyLevel = 'eventual' # Necessary for advanced queries on Azure AD
            All              = $true       # Retrieve all users matching the filter
            Select           = 'id,userPrincipalName,displayName,jobTitle,officeLocation,businessPhones'
        }
        
        try {
            $users = Get-MgUser @mgUserParams -ErrorAction Stop
            $outputObjects = @()
            foreach ($user in $users) {
                $outputObjects += [PSCustomObject]@{
                    Id             = $user.Id
                    Email          = $user.UserPrincipalName
                    DisplayName    = $user.DisplayName
                    JobTitle       = $user.JobTitle
                    OfficeLocation = $user.OfficeLocation
                    BusinessPhones = if ($user.BusinessPhones) { $user.BusinessPhones -join "; " } else { "" }
                    LicenseName    = $skuNameParam 
                    SkuId          = $skuIdParam
                }
            }
            return $outputObjects
        }
        catch {
            Write-Error "Error in job GetUsers_$($skuNameParam -replace '\W','_') for SKU '$skuIdParam': $($_.Exception.ToString())"
            return $null # Return null on error for this job
        }
    } -ArgumentList $skuId, $skuName
    
    Write-Host "Task started for license '$skuName'" -ForegroundColor DarkCyan
}

Write-Host "Waiting for tasks to complete..." -ForegroundColor Cyan
$jobs | Wait-Job # Wait for all started jobs to finish

# Collect results from completed jobs
foreach ($job in $jobs) {
    if ($job.State -eq 'Completed') {
        $jobResult = Receive-Job -Job $job
        if ($null -ne $jobResult) {
            $allEntries += $jobResult
        }
    }
    else {
        Write-Warning "Job $($job.Name) (ID: $($job.Id)) did not complete. State: $($job.State)."
        if ($job.Error.Count -gt 0) {
            $job.Error | ForEach-Object { Write-Warning "  Error in Job: $($_.Exception.ToString())" }
        }
    }
}

# Aggregate licenses by user
$userIndex = @{} # Hashtable to store unique users and their licenses
if ($allEntries.Count -gt 0) {
    foreach ($entry in $allEntries) {
        if ($null -ne $entry -and $entry.PSObject.Properties['Id']) { # Ensure entry is valid and has an Id
            if (-not $userIndex.ContainsKey($entry.Id)) {
                $userIndex[$entry.Id] = [PSCustomObject]@{
                    Id             = $entry.Id
                    Email          = $entry.Email
                    DisplayName    = $entry.DisplayName
                    JobTitle       = $entry.JobTitle
                    OfficeLocation = $entry.OfficeLocation
                    BusinessPhones = $entry.BusinessPhones
                    Licenses       = [System.Collections.Generic.List[PSCustomObject]]::new() # List to hold multiple licenses
                }
            }
            # Add current license to the user's list of licenses
            $userIndex[$entry.Id].Licenses.Add([PSCustomObject]@{
                LicenseName = $entry.LicenseName
                SkuId       = $entry.SkuId
            })
        }
    }
}
else {
    Write-Warning "No entries were collected from jobs. Check job error logs."
}

$results = $userIndex.Values # Get the collection of user objects
Write-Host ""
Write-Host "Total unique users processed: $($results.Count)"

# Convert aggregated data to JSON for the HTML report
$json = ""
if ($results.Count -gt 0) {
    $json = $results | ConvertTo-Json -Depth 5 -Compress # Depth 5 should be enough for nested licenses
    if ($json) {
        # Write-Host "JSON conversion successful." -ForegroundColor Green # Optional success message
    }
    else {
        Write-Warning "JSON conversion produced no output. Using error JSON."
        $json = '{ "error": "Failed to generate JSON from results", "data": [] }'
    }
}
else {
    Write-Warning "No results to convert to JSON. JSON will indicate 'no data'."
    $json = '{ "message": "No user data found or processed.", "data": [] }'
}

# HTML content with translated strings
$simpleHtmlContent = @"
<html lang="en-US">
<head>
    <meta charset="UTF-8" />
    <meta name="viewport" content="width=device-width, initial-scale=1.0" />
    <title>Corporate License Management</title>
    <script src="https://code.jquery.com/jquery-3.6.0.min.js"></script>
    <script src="https://cdn.datatables.net/1.13.7/js/jquery.dataTables.min.js"></script>
    <link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.5.1/css/all.min.css" />
    <link rel="stylesheet" href="https://cdn.datatables.net/1.13.7/css/jquery.dataTables.min.css" />
    <link rel="stylesheet" href="https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700&display=swap" />
    <link rel="stylesheet" href="https://lic-management.vercel.app/style.css">
    <link href="https://cdn.jsdelivr.net/npm/select2@4.1.0-rc.0/dist/css/select2.min.css" rel="stylesheet" />
</head>
<body>

<div id="loadingOverlay" style="display: none;">
  <div class="loader-content">
    <p>Loading...</p>
    </div>
</div>

    <script>
        // PowerShell will replace $json with actual user data
        const userData = $json; 
    </script>
    <aside>
        <a href="https://github.com/pedr-moura" target="_blank" class="sidebar-link">
            <svg xmlns="http://www.w3.org/2000/svg" width="30" height="30" viewBox="0 0 16 16">
                <path d="M8 0C3.58 0 0 3.58 0 8c0 3.54 2.29 6.53 5.47 7.59.4.07.55-.17.55-.38 0-.19-.01-.82-.01-1.49-2.01.37-2.53-.49-2.69-.94-.09-.23-.48-.94-.82-1.13-.28-.15-.68-.52-.01-.53.63-.01 1.08.58 1.23.82.72 1.21 1.87.87 2.33.66.07-.52.28-.87.51-1.07-1.78-.2-3.64-.89-3.64-3.95 0-.87.31-1.59.82-2.15-.08-.2-.36-1.02.08-2.12 0 0 .67-.21 2.2.82.64-.18 1.32-.27 2-.27s1.36.09 2 .27c1.53-1.04 2.2-.82 2.2-.82.44 1.1.16 1.92.08 2.12.51.56.82 1.27.82 2.15 0 3.07-1.87 3.75-3.65 3.95.29.25.54.73.54 1.48 0 1.07-.01 1.93-.01 2.2 0 .21.15.46.55.38A8.01 8.01 0 0 0 16 8c0-4.42-3.58-8-8-8"/>
            </svg>
            <span>Pedro Moura<br>2025</span>
        </a>
    </aside>
    <div class="content-container">
        <header>
            <h1><i class="fas fa-id-card-alt" style="margin-right: 0.5rem;"></i>License Management</h1>
        </header>
        <main>
            <div class="grid-container grid-container-md section-spacing">
                <div class="card col-span-2-md">
                    <span class="label">Visible Columns in Table</span>
                    <div class="checkbox-container" id="colContainer">
                        <label class="checkbox-label"><input type="checkbox" class="checkbox col-vis" data-col="0" /><span class="checkbox-text">ID</span></label>
                        <label class="checkbox-label"><input type="checkbox" class="checkbox col-vis" data-col="1" checked /><span class="checkbox-text">Name</span></label>
                        <label class="checkbox-label"><input type="checkbox" class="checkbox col-vis" data-col="2" checked /><span class="checkbox-text">Email</span></label>
                        <label class="checkbox-label"><input type="checkbox" class="checkbox col-vis" data-col="3" checked /><span class="checkbox-text">Job Title</span></label>
                        <label class="checkbox-label"><input type="checkbox" class="checkbox col-vis" data-col="4" /><span class="checkbox-text">Location</span></label> 
                        <label class="checkbox-label"><input type="checkbox" class="checkbox col-vis" data-col="5" /><span class="checkbox-text">Phones</span></label>
                        <label class="checkbox-label"><input type="checkbox" class="checkbox col-vis" data-col="6" checked /><span class="checkbox-text">Licenses</span></label>
                    </div>
                </div>
            </div>
            <div class="grid-container grid-container-lg section-spacing">
                <div class="card">
                    <div class="multi-search-header">
                        <h2><i class="fas fa-filter" style="margin-right: 0.5rem;"></i>Advanced Search</h2>
                        <div class="multi-search-controls">
                            <label for="multiSearchOperator">Operator:</label>
                            <select id="multiSearchOperator">
                                <option value="AND" selected>AND (All)</option>
                                <option value="OR">OR (Any)</option>
                            </select>
                            <button id="addSearchField" class="button button-blue"><i class="fas fa-plus"></i> Filter</button>
                        </div>
                    </div>
                    <div id="multiSearchFields"></div>
                    <div id="searchCriteria">Search criteria will be displayed here...</div>
                    <div id="alertPanel" style="margin-top: 1rem;"></div>
                </div>
                <div style="display: flex; flex-direction: column; gap: 1rem;">
                    <div class="card" style="padding: 1rem;">
                        <h3 class="label" style="margin-bottom: 1rem; font-size: 1rem;">Actions</h3>
                        <button id="clearFilters" class="button button-gray"><i class="fas fa-eraser"></i>Clear Filters</button>
                        <button id="exportCsv" class="button button-blue" style="margin-top: 0.75rem;"><i class="fas fa-file-csv"></i>Export CSV</button>
                        <button id="exportIssues" class="button button-red" style="margin-top: 0.75rem;"><i class="fas fa-triangle-exclamation"></i>Issues Report</button>
                    </div>
                </div>
            </div>
            <div class="card">
                <table id="licenseTable" class="table-container display responsive" style="width:100%">
                    <thead class="table-header">
                        <tr>
                            <th>ID</th>
                            <th>Name</th>
                            <th>Email</th>
                            <th>Job Title</th>
                            <th>Department</th>
                            <th>Phones</th>
                            <th>Licenses</th>
                        </tr>
                    </thead>
                    <tbody></tbody>
                </table>
            </div>
        </main>
    </div>
    <script src="https://lic-management.vercel.app/scripts.js" defer></script>
</body>
</html>
"@

# Define output directory and file
$OutputDirectory = "C:\LicManagement" # Standard practice to use variable for path
$OutputFileName = "license_management_report.html" # English file name
$ReportHtmlPath = Join-Path -Path $OutputDirectory -ChildPath $OutputFileName

# Create directory if it doesn't exist
if (-not (Test-Path -Path $OutputDirectory -PathType Container)) {
    # Write-Host "Directory '$OutputDirectory' not found. Attempting to create..." -ForegroundColor Yellow # Optional
    try {
        New-Item -Path $OutputDirectory -ItemType Directory -Force -ErrorAction Stop | Out-Null
        # Write-Host "Directory '$OutputDirectory' created successfully." -ForegroundColor Green # Optional
    }
    catch {
        Write-Error "Failed to create directory '$OutputDirectory': $($_.Exception.Message)"
        # Optionally exit or handle error if directory creation is critical
    }
}
else {
    # Write-Host "Directory '$OutputDirectory' already exists." -ForegroundColor DarkGray # Optional
}

# Save the HTML file
try {
    $simpleHtmlContent | Out-File -FilePath $ReportHtmlPath -Encoding UTF8 -ErrorAction Stop
    Write-Host "Report saved to: $ReportHtmlPath" -ForegroundColor Green
}
catch {
    Write-Error "Failed to save the HTML report file: $($_.Exception.Message)" 
}

# Clean up jobs
Get-Job | Remove-Job -Force
