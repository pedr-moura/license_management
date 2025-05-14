# 🔍 Listagem de Usuários e Licenças do Microsoft 365 via Microsoft Graph

Este script PowerShell conecta-se ao Microsoft Graph, consulta todos os usuários que possuem licenças ativas no tenant e agrupa os resultados por usuário, exibindo informações detalhadas como nome, e-mail, telefone, local de trabalho e as licenças atribuídas.

O script executa as tarefas de forma **paralela**, respeitando um limite de concorrência, para otimizar o tempo de execução sem sobrecarregar o sistema.

## 📋 Pré-requisitos

* PowerShell 5.1+
* Módulo [Microsoft.Graph](https://learn.microsoft.com/en-us/powershell/microsoftgraph/overview)

  ```powershell
  Install-Module Microsoft.Graph -Scope CurrentUser
  ```
* Permissões delegadas ou de aplicativo:

  * `User.Read.All`

## 🚀 Como usar

1. Autentique-se no Microsoft Graph com o escopo necessário:

   ```powershell
   Connect-MgGraph -Scopes User.Read.All
   ```

2. Execute o script completo abaixo.

3. O resultado será salvo como um arquivo JSON chamado `UsuariosComLicencas_Paralelo.json` no diretório atual.

---

## 💻 Script PowerShell (Exemplo)

```powershell
Connect-MgGraph -Scopes User.Read.All

# Exemplo fictício de mapeamento de SKUs para nomes legíveis
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

    Write-Host "Tarefa iniciada para a licença '$skuName'" -ForegroundColor DarkCyan
}

Write-Host "⏳ Aguardando a finalização das tarefas…" -ForegroundColor Cyan
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

## 📁 Saída (Exemplo Fictício)

```json
[
  {
    "Id": "abc123",
    "Email": "maria.silva@empresa.com",
    "DisplayName": "Maria Silva",
    "JobTitle": "Analista de TI",
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
    "Email": "joao.souza@empresa.com",
    "DisplayName": "João Souza",
    "JobTitle": "Gerente de Projetos",
    "OfficeLocation": "Rio de Janeiro",
    "BusinessPhones": "+55 21 98888-1111",
    "Licenses": [
      {
        "LicenseName": "Office 365 F3"
      }
    ]
  }
]
