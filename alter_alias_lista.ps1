# Script para alterar alias de usuários usando Exchange Online PowerShell
# Este é o método CORRETO para alterar ProxyAddresses

# =====================================================
# VERIFICAR/INSTALAR MÓDULO EXCHANGE ONLINE
# =====================================================
Write-Host "========================================" -ForegroundColor Cyan
Write-Host "ALTERAÇÃO DE ALIAS - EXCHANGE ONLINE" -ForegroundColor Cyan
Write-Host "========================================" -ForegroundColor Cyan

# Verificar se o módulo está instalado
Write-Host "`nVerificando módulo Exchange Online..." -ForegroundColor Yellow
if (Get-Module -ListAvailable -Name ExchangeOnlineManagement) {
    Write-Host "✓ Módulo Exchange Online Management encontrado" -ForegroundColor Green
} else {
    Write-Host "✗ Módulo não encontrado. Instalando..." -ForegroundColor Red
    try {
        Install-Module -Name ExchangeOnlineManagement -Force -Scope CurrentUser
        Write-Host "✓ Módulo instalado com sucesso!" -ForegroundColor Green
    } catch {
        Write-Host "✗ Erro ao instalar módulo: $_" -ForegroundColor Red
        Write-Host "Execute manualmente: Install-Module -Name ExchangeOnlineManagement" -ForegroundColor Yellow
        exit 1
    }
}

# Importar módulo
Import-Module ExchangeOnlineManagement -Force

# =====================================================
# CONECTAR AO EXCHANGE ONLINE
# =====================================================
Write-Host "`nConectando ao Exchange Online..." -ForegroundColor Yellow
try {
    # Verificar se já está conectado
    $existingSession = Get-PSSession | Where-Object { $_.ConfigurationName -eq "Microsoft.Exchange" -and $_.State -eq "Opened" }
    
    if ($existingSession) {
        Write-Host "✓ Já conectado ao Exchange Online" -ForegroundColor Green
    } else {
        Connect-ExchangeOnline -ShowBanner:$false
        Write-Host "✓ Conectado com sucesso ao Exchange Online!" -ForegroundColor Green
    }
} catch {
    Write-Host "✗ Erro ao conectar: $_" -ForegroundColor Red
    exit 1
}

# =====================================================
# CONFIGURAÇÃO - USUÁRIOS PARA ALTERAR
# =====================================================
$usuariosParaAlterar = @(
    "user@dominio.com",
    "user@dominio.com",
    "user@dominio.com"
)

$oldAliasDomain = "@edu.externatost.onmicrosoft.com"
$newAliasDomain = "@externatost.onmicrosoft.com"
$logFile = "Alteracao_Alias_Exchange_$(Get-Date -Format 'yyyyMMdd_HHmmss').csv"
$logEntries = @()

Write-Host "`nTotal de usuários para processar: $($usuariosParaAlterar.Count)" -ForegroundColor Yellow
Write-Host "Alterando de: $oldAliasDomain" -ForegroundColor Yellow
Write-Host "Para: $newAliasDomain" -ForegroundColor Yellow
Write-Host ""

# =====================================================
# PROCESSAR CADA USUÁRIO
# =====================================================
foreach ($userEmail in $usuariosParaAlterar) {
    Write-Host "========================================" -ForegroundColor Cyan
    Write-Host "Processando: $userEmail" -ForegroundColor Cyan
    
    try {
        # Obter informações da caixa de correio
        $mailbox = Get-Mailbox -Identity $userEmail -ErrorAction Stop
        
        if ($null -eq $mailbox) {
            Write-Host "✗ Caixa de correio não encontrada" -ForegroundColor Red
            $logEntries += [PSCustomObject]@{
                Usuario = $userEmail
                Status = "Não encontrado"
                AliasAntigo = "N/A"
                AliasNovo = "N/A"
                Erro = "Caixa de correio não encontrada"
                DataHora = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
            }
            continue
        }
        
        Write-Host "✓ Caixa de correio encontrada: $($mailbox.DisplayName)" -ForegroundColor Green
        
        # Listar aliases atuais
        Write-Host "`nAliases atuais:" -ForegroundColor Gray
        foreach ($email in $mailbox.EmailAddresses) {
            if ($email -like "*$oldAliasDomain") {
                Write-Host "  - $email [SERÁ ALTERADO]" -ForegroundColor Yellow
            } else {
                Write-Host "  - $email" -ForegroundColor Gray
            }
        }
        
        # Encontrar aliases que precisam ser alterados
        $aliasesToChange = $mailbox.EmailAddresses | Where-Object { $_ -like "*$oldAliasDomain" }
        
        if ($aliasesToChange.Count -eq 0) {
            Write-Host "`n✓ Nenhum alias com $oldAliasDomain encontrado" -ForegroundColor Green
            
            # Verificar se já tem o novo alias
            $hasNewAlias = $mailbox.EmailAddresses | Where-Object { $_ -like "*$newAliasDomain" }
            if ($hasNewAlias) {
                Write-Host "✓ Usuário já possui alias(es) com $newAliasDomain" -ForegroundColor Green
            }
            
            $logEntries += [PSCustomObject]@{
                Usuario = $userEmail
                Status = "Sem alterações necessárias"
                AliasAntigo = "N/A"
                AliasNovo = "N/A"
                EmailAddresses = ($mailbox.EmailAddresses -join ";")
                DataHora = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
            }
            continue
        }
        
        # Preparar alterações
        $addAddresses = @()
        $removeAddresses = @()
        
        foreach ($oldAlias in $aliasesToChange) {
            # Criar novo alias substituindo o domínio
            $newAlias = $oldAlias -replace [regex]::Escape($oldAliasDomain), $newAliasDomain
            
            $removeAddresses += $oldAlias
            $addAddresses += $newAlias
            
            Write-Host "`nAlterando:" -ForegroundColor Yellow
            Write-Host "  DE: $oldAlias" -ForegroundColor Red
            Write-Host "  PARA: $newAlias" -ForegroundColor Green
        }
        
        # Executar alteração
        Write-Host "`nAplicando alterações..." -ForegroundColor Yellow
        
        try {
            Set-Mailbox -Identity $userEmail `
                       -EmailAddresses @{Add=$addAddresses; Remove=$removeAddresses} `
                       -ErrorAction Stop
            
            Write-Host "✅ Alterações aplicadas com sucesso!" -ForegroundColor Green
            
            # Verificar alterações (aguardar um pouco para propagação)
            Start-Sleep -Seconds 2
            $updatedMailbox = Get-Mailbox -Identity $userEmail
            
            Write-Host "`nAliases após alteração:" -ForegroundColor Gray
            foreach ($email in $updatedMailbox.EmailAddresses) {
                if ($email -in $addAddresses) {
                    Write-Host "  - $email [NOVO]" -ForegroundColor Green
                } else {
                    Write-Host "  - $email" -ForegroundColor Gray
                }
            }
            
            # Adicionar ao log
            $logEntries += [PSCustomObject]@{
                Usuario = $userEmail
                DisplayName = $mailbox.DisplayName
                Status = "Sucesso"
                AliasesRemovidos = ($removeAddresses -join ";")
                AliasesAdicionados = ($addAddresses -join ";")
                EmailAddressesAtuais = ($updatedMailbox.EmailAddresses -join ";")
                DataHora = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
            }
            
        } catch {
            Write-Host "❌ Erro ao aplicar alterações: $_" -ForegroundColor Red
            
            $logEntries += [PSCustomObject]@{
                Usuario = $userEmail
                DisplayName = $mailbox.DisplayName
                Status = "Erro"
                AliasesRemovidos = ($removeAddresses -join ";")
                AliasesAdicionados = "N/A"
                Erro = $_.Exception.Message
                DataHora = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
            }
        }
        
    } catch {
        Write-Host "❌ Erro ao processar usuário: $_" -ForegroundColor Red
        
        $logEntries += [PSCustomObject]@{
            Usuario = $userEmail
            Status = "Erro"
            Erro = $_.Exception.Message
            DataHora = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
        }
    }
}

# =====================================================
# SALVAR LOG E EXIBIR RESUMO
# =====================================================
Write-Host "`n========================================" -ForegroundColor Cyan
Write-Host "RESUMO FINAL" -ForegroundColor Cyan
Write-Host "========================================" -ForegroundColor Cyan

# Exportar log
if ($logEntries.Count -gt 0) {
    $logEntries | Export-Csv -Path $logFile -NoTypeInformation -Encoding UTF8
    Write-Host "📄 Log detalhado salvo em: $logFile" -ForegroundColor Green
}

# Estatísticas
$sucessos = ($logEntries | Where-Object { $_.Status -eq "Sucesso" }).Count
$erros = ($logEntries | Where-Object { $_.Status -eq "Erro" }).Count
$semAlteracoes = ($logEntries | Where-Object { $_.Status -eq "Sem alterações necessárias" }).Count
$naoEncontrados = ($logEntries | Where-Object { $_.Status -eq "Não encontrado" }).Count

Write-Host "`nEstatísticas:" -ForegroundColor White
Write-Host "  Total processado: $($usuariosParaAlterar.Count)" -ForegroundColor White
Write-Host "  ✅ Alterados com sucesso: $sucessos" -ForegroundColor Green
Write-Host "  ℹ️  Sem alterações necessárias: $semAlteracoes" -ForegroundColor Yellow
Write-Host "  ❌ Erros: $erros" -ForegroundColor Red
Write-Host "  ❓ Não encontrados: $naoEncontrados" -ForegroundColor Red

# Mostrar sucessos
if ($sucessos -gt 0) {
    Write-Host "`n✅ Usuários alterados com sucesso:" -ForegroundColor Green
    $logEntries | Where-Object { $_.Status -eq "Sucesso" } | ForEach-Object {
        Write-Host "  - $($_.Usuario) ($($_.DisplayName))" -ForegroundColor Green
        Write-Host "    Removido: $($_.AliasesRemovidos)" -ForegroundColor Gray
        Write-Host "    Adicionado: $($_.AliasesAdicionados)" -ForegroundColor Gray
    }
}

# Mostrar erros
if ($erros -gt 0) {
    Write-Host "`n❌ Erros encontrados:" -ForegroundColor Red
    $logEntries | Where-Object { $_.Status -eq "Erro" } | ForEach-Object {
        Write-Host "  - $($_.Usuario): $($_.Erro)" -ForegroundColor Red
    }
}

Write-Host "`n========================================" -ForegroundColor Cyan
Write-Host "Processamento concluído!" -ForegroundColor Cyan
Write-Host "========================================" -ForegroundColor Cyan

# Desconectar do Exchange Online
Write-Host "`nDesconectando do Exchange Online..." -ForegroundColor Yellow
Disconnect-ExchangeOnline -Confirm:$false -ErrorAction SilentlyContinue
Write-Host "✓ Desconectado" -ForegroundColor Green