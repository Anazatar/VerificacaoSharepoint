[CmdletBinding()]
param(
    [string]$tenant
)

if ($PSBoundParameters.ContainsKey('Verbose') -and $PSBoundParameters['Verbose']) {
    $VerbosePreference = 'Continue'
}

$modulosNecessarios = @(
    'Microsoft.Graph.Authentication',
    'Microsoft.Graph.Sites',
    'Microsoft.Graph.Files',
    'Microsoft.PowerApps.Administration.PowerShell',
    'Microsoft.Online.SharePoint.PowerShell'
)

foreach ($modulo in $modulosNecessarios) {
    if (-not (Get-Module -ListAvailable -Name $modulo)) {
        try {
            Install-Module $modulo -Scope CurrentUser -Force -ErrorAction Stop
            Write-Host "Módulo $modulo instalado com sucesso." -ForegroundColor Green
        } catch {
            Write-Host "Erro ao instalar o módulo $modulo $($_.Exception.Message)" -ForegroundColor Red
            return
        }
    }
}

foreach ($modulo in $modulosNecessarios) {
    try {
        Import-Module $modulo -Force -ErrorAction Stop
        Write-Host "Módulo $modulo importado com sucesso." -ForegroundColor Green
    } catch {
        Write-Host "Erro ao importar o módulo $modulo $($_.Exception.Message)" -ForegroundColor Red
        return
    }
}

$moduloNome = 'scriptverif.psm1'
$moduloPath = Join-Path -Path $PSScriptRoot -ChildPath $moduloNome
$moduloUrl  = 'https://raw.githubusercontent.com/Anazatar/verifmod/refs/heads/main/scriptverif.psm1'

if (-not (Test-Path $moduloPath)) {
    Write-Host "Módulo $moduloNome não encontrado localmente. Baixando de $moduloUrl ..." -ForegroundColor Yellow
    try {
        Invoke-WebRequest -Uri $moduloUrl -OutFile $moduloPath -UseBasicParsing -TimeoutSec 30
        Write-Host 'Download concluído.' -ForegroundColor Green
    } catch {
        Write-Host "Erro ao baixar ${moduloNome} $($_.Exception.Message)" -ForegroundColor Red
        return
    }
}

try {
    (Get-Content $moduloPath -Raw) | Set-Content -Path $moduloPath -Encoding utf8
    Write-Host 'Codificação UTF‑8 definida para o módulo.' -ForegroundColor Green
} catch {
    Write-Host "Erro ao definir codificação UTF‑8 no módulo: $($_.Exception.Message)" -ForegroundColor Red
}

try {
    Import-Module $moduloPath -Force
    Write-Host "Módulo $moduloNome carregado." -ForegroundColor Green
} catch {
    Write-Host "Falha ao importar ${moduloNome}: $($_.Exception.Message)" -ForegroundColor Red
    return
}

try {
    $adminUrl = "https://$($tenant)-admin.sharepoint.com"
    Connect-SPOService -Url $adminUrl
    Write-Host "Conectado ao SharePoint Admin: $adminUrl" -ForegroundColor Green
} catch {
    Write-Host "Erro ao conectar ao SharePoint Online: $($_.Exception.Message)" -ForegroundColor Red
    return
}

try {
    $clientId = $env:AZURE_CLIENT_ID
    $clientSecret = $env:AZURE_SECRET

    $secureSecret = ConvertTo-SecureString $clientSecret -AsPlainText -Force

    Connect-MgGraph -ClientId $clientId -ClientSecret $secureSecret -Scopes "https://graph.microsoft.com/.default"

    Write-Host 'Conectado à Microsoft Graph via App Registration.' -ForegroundColor Green
} catch {
    Write-Host "Erro ao conectar ao Microsoft Graph: $($_.Exception.Message)" -ForegroundColor Red
    return
}


Write-Host "Tenant recebido: $tenant"

$relatorioLimitacoesAplicaveis    = @()
$relatorioLimitacoesNaoAplicaveis = @()

Verificar-MultiGeo                                -relatorioAplicaveis ([ref]$relatorioLimitacoesAplicaveis) `
                                                  -relatorioNaoAplicaveis ([ref]$relatorioLimitacoesNaoAplicaveis)
Verificar-LimitacoesTenant -tenant $tenant        -relatorioAplicaveis ([ref]$relatorioLimitacoesAplicaveis) `
                                                  -relatorioNaoAplicaveis ([ref]$relatorioLimitacoesNaoAplicaveis)
Verificar-SitesExcluidos                          -relatorioAplicaveis ([ref]$relatorioLimitacoesAplicaveis) `
                                                  -relatorioNaoAplicaveis ([ref]$relatorioLimitacoesNaoAplicaveis)
Verificar-OneDriveSync                            -relatorioAplicaveis ([ref]$relatorioLimitacoesAplicaveis) `
                                                  -relatorioNaoAplicaveis ([ref]$relatorioLimitacoesNaoAplicaveis)
Verificar-OneNote                                 -relatorioAplicaveis ([ref]$relatorioLimitacoesAplicaveis) `
                                                  -relatorioNaoAplicaveis ([ref]$relatorioLimitacoesNaoAplicaveis)
Verificar-Redirect308     -url "https://$tenant.sharepoint.com/sites/exemplo" `
                                                  -relatorioAplicaveis ([ref]$relatorioLimitacoesAplicaveis) `
                                                  -relatorioNaoAplicaveis ([ref]$relatorioLimitacoesNaoAplicaveis)
Verificar-Delve                                   -relatorioAplicaveis ([ref]$relatorioLimitacoesAplicaveis)
Verificar-eDiscovery                              -relatorioAplicaveis ([ref]$relatorioLimitacoesAplicaveis)
Buscar-FormulariosInfoPathGraph -tenant $tenant   -relatorioAplicaveis ([ref]$relatorioLimitacoesAplicaveis) `
                                                  -relatorioNaoAplicaveis ([ref]$relatorioLimitacoesNaoAplicaveis)
Verificar-Loop                                    -relatorioAplicaveis ([ref]$relatorioLimitacoesAplicaveis)
Verificar-SitesArquivados                         -relatorioAplicaveis ([ref]$relatorioLimitacoesAplicaveis) `
                                                  -relatorioNaoAplicaveis ([ref]$relatorioLimitacoesNaoAplicaveis)
Verificar-MicrosoftFormsUpload                    -relatorioAplicaveis ([ref]$relatorioLimitacoesAplicaveis) `
                                                  -relatorioNaoAplicaveis ([ref]$relatorioLimitacoesNaoAplicaveis)
Verificar-OfficeAppsSalvamento                    -relatorioAplicaveis ([ref]$relatorioLimitacoesAplicaveis)
Verificar-OneDriveAcessoRapido                    -relatorioAplicaveis ([ref]$relatorioLimitacoesAplicaveis)
Verificar-OneDriveTeamsApp                        -relatorioAplicaveis ([ref]$relatorioLimitacoesAplicaveis)
Verificar-PowerPlatformConectoresSharePoint       -relatorioAplicaveis ([ref]$relatorioLimitacoesAplicaveis) `
                                                  -relatorioNaoAplicaveis ([ref]$relatorioLimitacoesNaoAplicaveis)
Verificar-ProjectOnlineWorkflows                  -relatorioAplicaveis ([ref]$relatorioLimitacoesAplicaveis)
Verificar-ProjectOnlinePWA                        -relatorioAplicaveis ([ref]$relatorioLimitacoesAplicaveis)
Verificar-ProjectOnlineExcelRelatorios            -relatorioAplicaveis ([ref]$relatorioLimitacoesAplicaveis)
Verificar-ProjectPro                              -relatorioAplicaveis ([ref]$relatorioLimitacoesAplicaveis)
Verificar-SitesHubSharePoint                      -relatorioAplicaveis ([ref]$relatorioLimitacoesAplicaveis) `
                                                  -relatorioNaoAplicaveis ([ref]$relatorioLimitacoesNaoAplicaveis)
Verificar-SitesBloqueados                         -relatorioAplicaveis ([ref]$relatorioLimitacoesAplicaveis) `
                                                  -relatorioNaoAplicaveis ([ref]$relatorioLimitacoesNaoAplicaveis)
Verificar-URLsAlternativos                        -relatorioAplicaveis ([ref]$relatorioLimitacoesAplicaveis) `
                                                  -relatorioNaoAplicaveis ([ref]$relatorioLimitacoesNaoAplicaveis)

$saidaHTML = Join-Path -Path $PSScriptRoot -ChildPath 'relatorio_limitacoes.html'

$css = @"
<style>
    body { font-family: Arial, sans-serif; background-color: #f4f4f4; padding: 20px; }
    h1  { text-align: center; color: #333; }
    h2  { color: #444; margin-top: 40px; }
    table { width: 100%; border-collapse: collapse; margin-top: 10px; }
    th, td { padding: 10px; border: 1px solid #ddd; text-align: left; }
    th { background-color: #0078d4; color: white; }
    tr:nth-child(even) { background-color: #f9f9f9; }
    .section { background-color: white; padding: 20px; border-radius: 8px; box-shadow: 0 2px 4px rgba(0,0,0,0.1); margin-bottom: 20px; }
</style>
"@

$header = "<html><head><meta charset='UTF-8'><title>Relatório de Limitações</title>$css</head><body><h1>Relatório de Limitações do Tenant: $tenant</h1>"
$footer = '</body></html>'

$corpoAplicaveis     = $relatorioLimitacoesAplicaveis | ConvertTo-Html -Fragment -PreContent "<div class='section'><h2>Limitações Aplicáveis</h2>"
$corpoNaoAplicaveis  = $relatorioLimitacoesNaoAplicaveis | ConvertTo-Html -Fragment -PreContent "</div><div class='section'><h2>Limitações Não Aplicáveis</h2>"

$paginaCompleta = "$header$corpoAplicaveis$corpoNaoAplicaveis</div>$footer"
$paginaCompleta | Out-File -FilePath $saidaHTML -Encoding utf8

Write-Host "Relatório HTML gerado em: $saidaHTML" -ForegroundColor Cyan


try { Disconnect-SPOService } catch { }
try { Disconnect-MgGraph }     catch { }
