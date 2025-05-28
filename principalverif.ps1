param(
    [string]$tenant
)


# ================================
# Carregar e importar módulos mínimos do Microsoft Graph e outros
# ================================

$modulosNecessarios = @(
    "Microsoft.Graph.Authentication",
    "Microsoft.Graph.Sites",
    "Microsoft.Graph.Files",
    "Microsoft.PowerApps.Administration.PowerShell",
    "Microsoft.Online.SharePoint.PowerShell"
)

foreach ($modulo in $modulosNecessarios) {
    if (-not (Get-Module -ListAvailable -Name $modulo)) {
        try {
            Install-Module $modulo -Scope CurrentUser -Force -ErrorAction Stop
            Write-Host "Módulo ${modulo} instalado com sucesso." -ForegroundColor Green
        } catch {
            Write-Host "Erro ao instalar o módulo ${modulo}: $($_.Exception.Message)" -ForegroundColor Red
            return
        }
    }
}

foreach ($modulo in $modulosNecessarios) {
    try {
        Import-Module $modulo -Force -ErrorAction Stop
        Write-Host "Módulo ${modulo} importado com sucesso." -ForegroundColor Green
    } catch {
        Write-Host "Erro ao importar o módulo ${modulo}: $($_.Exception.Message)" -ForegroundColor Red
        return
    }
}

# ================================
# Carregar módulo local scriptverif.psm1 com download se não existir
# ================================

$moduloNome = "scriptverif.psm1"
$moduloPath = Join-Path -Path $PSScriptRoot -ChildPath $moduloNome
$moduloUrl = "https://raw.githubusercontent.com/Anazatar/verifmod/refs/heads/main/scriptverif.psm1"

if (-not (Test-Path $moduloPath)) {
    Write-Host "Módulo ${moduloNome} não encontrado localmente. Baixando de ${moduloUrl} ..." -ForegroundColor Yellow
    try {
        Invoke-WebRequest -Uri $moduloUrl -OutFile $moduloPath -UseBasicParsing
        Write-Host "Módulo baixado com sucesso." -ForegroundColor Green
    } catch {
        Write-Host "Erro ao baixar o módulo ${moduloNome}: $($_.Exception.Message)" -ForegroundColor Red
        return
    }
}

try {
    $content = Get-Content $moduloPath -Raw
    Set-Content -Path $moduloPath -Value $content -Encoding UTF8
    Write-Host "Codificação UTF-8 garantida para o módulo local." -ForegroundColor Green
} catch {
    Write-Host "Erro ao forçar UTF-8: $($_.Exception.Message)" -ForegroundColor Red
}


try {
    Import-Module $moduloPath -Force
    Write-Host "Módulo local ${moduloNome} importado com sucesso." -ForegroundColor Green
} catch {
    Write-Host "Falha ao importar o módulo local ${moduloNome}: $($_.Exception.Message)" -ForegroundColor Red
    return
}


# ================================
# Autenticação Microsoft Graph
# ================================

try {
    Connect-MgGraph -Scopes "Sites.Read.All" -NoWelcome
    Write-Host "Conectado à Microsoft Graph." -ForegroundColor Green
} catch {
    Write-Host "Erro ao conectar ao Microsoft Graph: $($_.Exception.Message)" -ForegroundColor Red
    return
}



# ================================
# Entrada do usuário
# ================================

Write-Host "Tenant recebido: $tenant"

# ================================
# Inicialização dos relatórios
# ================================

$relatorioLimitacoesAplicaveis = @()
$relatorioLimitacoesNaoAplicaveis = @()

# ================================
# Execução das verificações
# ================================

Verificar-LimitacoesTenant -tenant $tenant `
    -relatorioAplicaveis ([ref]$relatorioLimitacoesAplicaveis) `
    -relatorioNaoAplicaveis ([ref]$relatorioLimitacoesNaoAplicaveis)

Verificar-OneDriveSync -relatorioAplicaveis ([ref]$relatorioLimitacoesAplicaveis) `
    -relatorioNaoAplicaveis ([ref]$relatorioLimitacoesNaoAplicaveis)

Verificar-OneNote -relatorioAplicaveis ([ref]$relatorioLimitacoesAplicaveis) `
    -relatorioNaoAplicaveis ([ref]$relatorioLimitacoesNaoAplicaveis)

Verificar-Redirect308 -url "https://$tenant.sharepoint.com/sites/exemplo" `
    -relatorioAplicaveis ([ref]$relatorioLimitacoesAplicaveis) `
    -relatorioNaoAplicaveis ([ref]$relatorioLimitacoesNaoAplicaveis)

Verificar-Delve -relatorioAplicaveis ([ref]$relatorioLimitacoesAplicaveis)

Verificar-eDiscovery -relatorioAplicaveis ([ref]$relatorioLimitacoesAplicaveis)

Buscar-FormulariosInfoPathGraph -tenant $tenant `
    -relatorioAplicaveis ([ref]$relatorioLimitacoesAplicaveis) `
    -relatorioNaoAplicaveis ([ref]$relatorioLimitacoesNaoAplicaveis)

Verificar-Loop -relatorioAplicaveis ([ref]$relatorioLimitacoesAplicaveis)

Verificar-SitesArquivados -relatorioAplicaveis ([ref]$relatorioLimitacoesAplicaveis) `
                          -relatorioNaoAplicaveis ([ref]$relatorioLimitacoesNaoAplicaveis)

Verificar-MicrosoftFormsUpload -relatorioAplicaveis ([ref]$relatorioLimitacoesAplicaveis) `
                               -relatorioNaoAplicaveis ([ref]$relatorioLimitacoesNaoAplicaveis)
          
Verificar-OfficeAppsSalvamento -relatorioAplicaveis ([ref]$relatorioLimitacoesAplicaveis)

Verificar-OneDriveAcessoRapido -relatorioAplicaveis ([ref]$relatorioLimitacoesAplicaveis)

Verificar-OneDriveTeamsApp -relatorioAplicaveis ([ref]$relatorioLimitacoesAplicaveis)

Verificar-PowerPlatformConectoresSharePoint -relatorioAplicaveis ([ref]$relatorioLimitacoesAplicaveis) `
                                            -relatorioNaoAplicaveis ([ref]$relatorioLimitacoesNaoAplicaveis)  

Verificar-ProjectOnlineWorkflows -relatorioAplicaveis ([ref]$relatorioLimitacoesAplicaveis)

Verificar-ProjectOnlinePWA -relatorioAplicaveis ([ref]$relatorioLimitacoesAplicaveis)

Verificar-ProjectOnlineExcelRelatorios -relatorioAplicaveis ([ref]$relatorioLimitacoesAplicaveis)

Verificar-ProjectPro -relatorioAplicaveis ([ref]$relatorioLimitacoesAplicaveis)
               

Verificar-SitesHubSharePoint -relatorioAplicaveis ([ref]$relatorioLimitacoesAplicaveis) `
                             -relatorioNaoAplicaveis ([ref]$relatorioLimitacoesNaoAplicaveis)

Verificar-SitesBloqueados -relatorioAplicaveis ([ref]$relatorioLimitacoesAplicaveis) `
                         -relatorioNaoAplicaveis ([ref]$relatorioLimitacoesNaoAplicaveis)



# ================================
# Gerar relatório HTML
# ================================

$saidaHTML = Join-Path -Path $PSScriptRoot -ChildPath "relatorio_limitacoes.html"

$css = @"
<style>
    body { font-family: Arial, sans-serif; background-color: #f4f4f4; padding: 20px; }
    h1 { text-align: center; color: #333; }
    h2 { color: #444; margin-top: 40px; }
    table { width: 100%; border-collapse: collapse; margin-top: 10px; }
    th, td { padding: 10px; border: 1px solid #ddd; text-align: left; }
    th { background-color: #0078d4; color: white; }
    tr:nth-child(even) { background-color: #f9f9f9; }
    .section { background-color: white; padding: 20px; border-radius: 8px; box-shadow: 0 2px 4px rgba(0,0,0,0.1); margin-bottom: 20px; }
</style>
"@

$header = "<html><head><meta charset='UTF-8'><title>Relatório de Limitações</title>$css</head><body><h1>Relatório de Limitações do Tenant: $tenant</h1>"
$footer = "</body></html>"

$corpoAplicaveis = $relatorioLimitacoesAplicaveis | ConvertTo-Html -Fragment -PreContent "<div class='section'><h2>Limitações Aplicáveis</h2>"
$corpoNaoAplicaveis = $relatorioLimitacoesNaoAplicaveis | ConvertTo-Html -Fragment -PreContent "</div><div class='section'><h2>Limitações Não Aplicáveis</h2>"

$paginaCompleta = "$header$corpoAplicaveis$corpoNaoAplicaveis</div>$footer"
$paginaCompleta | Out-File -FilePath $saidaHTML -Encoding UTF8

Write-Host "Relatório HTML gerado em: $saidaHTML" -ForegroundColor Cyan

# ================================
# Encerramento
# ================================

Disconnect-SPOService
