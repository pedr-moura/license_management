# 🔍 Listing Microsoft 365 Users and Licenses via Microsoft Graph

This PowerShell script connects to Microsoft Graph, queries all users with active licenses in the tenant, and groups the results by user, displaying detailed information such as name, email, phone, office location, and assigned licenses.

The script performs tasks in **parallel**, respecting a concurrency limit to optimize execution time without overloading the system.

## 📋 Prerequisites

* PowerShell 5.1+

* [Microsoft.Graph](https://learn.microsoft.com/en-us/powershell/microsoftgraph/overview) module:

  ```powershell
  Install-Module Microsoft.Graph -Scope CurrentUser
  ```

* Delegated or application permissions:

  * `User.Read.All`

## 🚀 How to Use

1. Authenticate to Microsoft Graph with the required scope:

   ```powershell
   Connect-MgGraph -Scopes User.Read.All
   ```

2. Run the full script below.

3. The result will be saved as a JSON file named `UsuariosComLicencas_Paralelo.json` in the current directory.

---

## 💻 PowerShell Script (Example)

```powershell
Connect-MgGraph -Scopes User.Read.All

# Example mapping of SKUs to human-readable names
$skuIdMap = @{
    "SKU001" = "Microsoft 365 E3"
    "SKU002" = "Office 365 F3"
    "SKU003" = "Enterprise Mobility + Security"
    "SKU004" = "Power BI Pro"
    "SKU005" = "Visio Plan 2"
}

$throttleLimit = 5
$jobs = @()

foreach ($skuId in $skuIdMap.Keys) {
    $skuName = $skuIdMap[$skuId]

    while (($jobs | Where-Object State -eq 'Running').Count -ge $throttleLimit) {
        Start-Sleep -Seconds 1
        $jobs = $jobs | Where-Object { $_.State -eq 'Running' }
    }

    $jobs += Start-ThreadJob -Name "GetUsers_$skuName" -ScriptBlock {
        param($skuId, $skuName)

        $mgUserParams = @{
            Filter = "assignedLicenses/any(u:u/skuId eq '$skuId')"
            ConsistencyLevel = 'eventual'
            All = $true
            Select = 'id,userPrincipalName,displayName,jobTitle,officeLocation,businessPhones'
        }

        Get-MgUser @mgUserParams | ForEach-Object {
            [PSCustomObject]@{
                Id = $_.Id
                Email = $_.UserPrincipalName
                DisplayName = $_.DisplayName
                JobTitle = $_.JobTitle
                OfficeLocation = $_.OfficeLocation
                BusinessPhones = if ($_.BusinessPhones) { $_.BusinessPhones -join "; " } else { "" }
                LicenseName = $skuName
            }
        }
    } -ArgumentList $skuId, $skuName

    Write-Host "Task started for license '$skuName'" -ForegroundColor DarkCyan
}

Write-Host "⏳ Waiting for all tasks to complete…" -ForegroundColor Cyan
$jobs | Wait-Job

$allEntries = $jobs | Receive-Job

$userIndex = @{}
foreach ($entry in $allEntries) {
    if (-not $userIndex.ContainsKey($entry.Id)) {
        $userIndex[$entry.Id] = [PSCustomObject]@{
            Id = $entry.Id
            Email = $entry.Email
            DisplayName = $entry.DisplayName
            JobTitle = $entry.JobTitle
            OfficeLocation = $entry.OfficeLocation
            BusinessPhones = $entry.BusinessPhones
            Licenses = @()
        }
    }
    $userIndex[$entry.Id].Licenses += [PSCustomObject]@{
        LicenseName = $entry.LicenseName
    }
}

$results = $userIndex.Values
$json = $results | ConvertTo-Json -Depth 5
$json | Out-File -FilePath ".\DADOS_OBTIDOS.json" -Encoding UTF8
```

---

## 📁 Output (Sample)

```json
[
  {
    "Id": "abc123",
    "Email": "maria.silva@company.com",
    "DisplayName": "Maria Silva",
    "JobTitle": "IT Analyst",
    "OfficeLocation": "São Paulo",
    "BusinessPhones": "+55 11 99999-0000",
    "Licenses": [
      {
        "LicenseName": "Microsoft 365 E3"
      },
      {
        "LicenseName": "Power BI Pro"
      }
    ]
  },
  {
    "Id": "xyz789",
    "Email": "joao.souza@company.com",
    "DisplayName": "João Souza",
    "JobTitle": "Project Manager",
    "OfficeLocation": "Rio de Janeiro",
    "BusinessPhones": "+55 21 98888-1111",
    "Licenses": [
      {
        "LicenseName": "Office 365 F3"
      }
    ]
  }
]
```
