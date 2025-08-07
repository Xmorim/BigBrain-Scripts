# Script para buscar usuários por UPN e alterar aliases usando Exchange Online PowerShell
# Processa todos os usuários que correspondem aos critérios definidos

# =====================================================
# VERIFICAR/INSTALAR MÓDULO EXCHANGE ONLINE
# =====================================================
Write-Host "========================================" -ForegroundColor Cyan
Write-Host "ALTERAÇÃO DE ALIAS EM MASSA - EXCHANGE" -ForegroundColor Cyan
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
# CONFIGURAÇÃO - CRITÉRIOS DE BUSCA E ALTERAÇÃO
# =====================================================
# Domínio UPN para buscar usuários
$targetUPNDomain = "@edu.salesianost.com.br"

# Configuração de alteração de alias
$oldAliasDomain = "@edu.externatost.onmicrosoft.com"
$newAliasDomain = "@externatost.onmicrosoft.com"

# Arquivos de log
$timestamp = Get-Date -Format 'yyyyMMdd_HHmmss'
$logFile = "Alteracao_Alias_Massa_${timestamp}.csv"
$errorLogFile = "Alteracao_Alias_Erros_${timestamp}.txt"

# Arrays para armazenar resultados
$logEntries = @()
$errorMessages = @()

# =====================================================
# BUSCAR USUÁRIOS
# =====================================================
Write-Host "`n========================================" -ForegroundColor Cyan
Write-Host "BUSCANDO USUÁRIOS" -ForegroundColor Cyan
Write-Host "========================================" -ForegroundColor Cyan
Write-Host "Critério: UPN contendo '$targetUPNDomain'" -ForegroundColor Yellow
Write-Host "Aguarde, isso pode levar alguns minutos..." -ForegroundColor Yellow

try {
    # Buscar todas as caixas de correio com o domínio UPN especificado
    $mailboxes = Get-Mailbox -ResultSize Unlimited -Filter "UserPrincipalName -like '*$targetUPNDomain'" | 
                 Select-Object UserPrincipalName, DisplayName, EmailAddresses, PrimarySmtpAddress, Alias
    
    Write-Host "✓ Total de usuários encontrados: $($mailboxes.Count)" -ForegroundColor Green
    
    if ($mailboxes.Count -eq 0) {
        Write-Host "⚠ Nenhum usuário encontrado com o critério especificado" -ForegroundColor Yellow
        exit 0
    }
    
} catch {
    Write-Host "✗ Erro ao buscar usuários: $_" -ForegroundColor Red
    exit 1
}

# =====================================================
# ANÁLISE PRÉVIA
# =====================================================
Write-Host "`n========================================" -ForegroundColor Cyan
Write-Host "ANÁLISE PRÉVIA" -ForegroundColor Cyan
Write-Host "========================================" -ForegroundColor Cyan

$usersWithOldAlias = 0
$usersAlreadyUpdated = 0
$usersNoChangesNeeded = 0

foreach ($mailbox in $mailboxes) {
    $hasOldAlias = $mailbox.EmailAddresses | Where-Object { $_ -like "*$oldAliasDomain" }
    $hasNewAlias = $mailbox.EmailAddresses | Where-Object { $_ -like "*$newAliasDomain" }
    
    if ($hasOldAlias) {
        $usersWithOldAlias++
    } elseif ($hasNewAlias) {
        $usersAlreadyUpdated++
    } else {
        $usersNoChangesNeeded++
    }
}

Write-Host "Usuários que precisam alteração: $usersWithOldAlias" -ForegroundColor Yellow
Write-Host "Usuários já atualizados: $usersAlreadyUpdated" -ForegroundColor Green
Write-Host "Usuários sem o alias específico: $usersNoChangesNeeded" -ForegroundColor Gray

if ($usersWithOldAlias -eq 0) {
    Write-Host "`n✓ Nenhum usuário precisa de alteração!" -ForegroundColor Green
    exit 0
}

# Perguntar se deseja continuar
Write-Host "`n⚠ Serão alterados $usersWithOldAlias usuários." -ForegroundColor Yellow
$continuar = Read-Host "Deseja continuar? (S/N)"
if ($continuar -ne 'S' -and $continuar -ne 's') {
    Write-Host "Operação cancelada pelo usuário." -ForegroundColor Red
    exit 0
}

# =====================================================
# PROCESSAR ALTERAÇÕES
# =====================================================
Write-Host "`n========================================" -ForegroundColor Cyan
Write-Host "PROCESSANDO ALTERAÇÕES" -ForegroundColor Cyan
Write-Host "========================================" -ForegroundColor Cyan

$processados = 0
$sucessos = 0
$erros = 0
$semAlteracoes = 0

foreach ($mailbox in $mailboxes) {
    $processados++
    $percentual = [math]::Round(($processados / $mailboxes.Count) * 100, 1)
    
    Write-Host "`n[$processados/$($mailboxes.Count)] ($percentual%) Processando: $($mailbox.UserPrincipalName)" -ForegroundColor Cyan
    
    try {
        # Verificar aliases que precisam ser alterados
        $aliasesToChange = $mailbox.EmailAddresses | Where-Object { $_ -like "*$oldAliasDomain" }
        
        if ($aliasesToChange.Count -eq 0) {
            # Verificar se já tem o novo alias
            $hasNewAlias = $mailbox.EmailAddresses | Where-Object { $_ -like "*$newAliasDomain" }
            
            if ($hasNewAlias) {
                Write-Host "  ✓ Usuário já possui alias atualizado" -ForegroundColor Green
                $status = "Já atualizado"
            } else {
                Write-Host "  ℹ Usuário não possui alias $oldAliasDomain" -ForegroundColor Gray
                $status = "Sem alias para alterar"
            }
            
            $semAlteracoes++
            
            # Adicionar ao log
            $logEntries += [PSCustomObject]@{
                UserPrincipalName = $mailbox.UserPrincipalName
                DisplayName = $mailbox.DisplayName
                Status = $status
                AliasesRemovidos = "N/A"
                AliasesAdicionados = "N/A"
                EmailAddressesAtuais = ($mailbox.EmailAddresses -join ";")
                DataHora = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
            }
            continue
        }
        
        # Preparar alterações
        $addAddresses = @()
        $removeAddresses = @()
        
        Write-Host "  Aliases a alterar:" -ForegroundColor Yellow
        foreach ($oldAlias in $aliasesToChange) {
            $newAlias = $oldAlias -replace [regex]::Escape($oldAliasDomain), $newAliasDomain
            $removeAddresses += $oldAlias
            $addAddresses += $newAlias
            Write-Host "    - $oldAlias → $newAlias" -ForegroundColor Yellow
        }
        
        # Executar alteração
        Write-Host "  Aplicando alterações..." -ForegroundColor Gray
        
        Set-Mailbox -Identity $mailbox.UserPrincipalName `
                   -EmailAddresses @{Add=$addAddresses; Remove=$removeAddresses} `
                   -ErrorAction Stop `
                   -WarningAction SilentlyContinue
        
        Write-Host "  ✅ Alteração realizada com sucesso!" -ForegroundColor Green
        $sucessos++
        
        # Adicionar ao log de sucesso
        $logEntries += [PSCustomObject]@{
            UserPrincipalName = $mailbox.UserPrincipalName
            DisplayName = $mailbox.DisplayName
            Status = "Sucesso"
            AliasesRemovidos = ($removeAddresses -join ";")
            AliasesAdicionados = ($addAddresses -join ";")
            EmailAddressesOriginais = ($mailbox.EmailAddresses -join ";")
            DataHora = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
        }
        
    } catch {
        Write-Host "  ❌ Erro: $_" -ForegroundColor Red
        $erros++
        
        # Adicionar ao log de erro
        $errorMessage = "[$($mailbox.UserPrincipalName)] Erro: $_"
        $errorMessages += $errorMessage
        
        $logEntries += [PSCustomObject]@{
            UserPrincipalName = $mailbox.UserPrincipalName
            DisplayName = $mailbox.DisplayName
            Status = "Erro"
            AliasesRemovidos = "N/A"
            AliasesAdicionados = "N/A"
            Erro = $_.Exception.Message
            DataHora = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
        }
    }
    
    # Mostrar progresso a cada 10 usuários
    if ($processados % 10 -eq 0) {
        Write-Host "`n--- Progresso: $processados de $($mailboxes.Count) processados ---" -ForegroundColor Cyan
        Write-Host "Sucessos: $sucessos | Erros: $erros | Sem alterações: $semAlteracoes" -ForegroundColor Gray
    }
}

# =====================================================
# SALVAR LOGS
# =====================================================
Write-Host "`n========================================" -ForegroundColor Cyan
Write-Host "SALVANDO LOGS" -ForegroundColor Cyan
Write-Host "========================================" -ForegroundColor Cyan

# Exportar log principal
if ($logEntries.Count -gt 0) {
    $logEntries | Export-Csv -Path $logFile -NoTypeInformation -Encoding UTF8
    Write-Host "✓ Log principal salvo em: $logFile" -ForegroundColor Green
}

# Exportar log de erros se houver
if ($errorMessages.Count -gt 0) {
    $errorMessages | Out-File -FilePath $errorLogFile -Encoding UTF8
    Write-Host "✓ Log de erros salvo em: $errorLogFile" -ForegroundColor Yellow
}

# =====================================================
# RESUMO FINAL
# =====================================================
Write-Host "`n========================================" -ForegroundColor Cyan
Write-Host "RESUMO FINAL" -ForegroundColor Cyan
Write-Host "========================================" -ForegroundColor Cyan

Write-Host "`nEstatísticas gerais:" -ForegroundColor White
Write-Host "  Total de usuários encontrados: $($mailboxes.Count)" -ForegroundColor White
Write-Host "  Total processado: $processados" -ForegroundColor White
Write-Host "  ✅ Alterados com sucesso: $sucessos" -ForegroundColor Green
Write-Host "  ❌ Erros: $erros" -ForegroundColor Red
Write-Host "  ℹ️  Sem alterações necessárias: $semAlteracoes" -ForegroundColor Yellow

# Taxa de sucesso
if ($usersWithOldAlias -gt 0) {
    $taxaSucesso = [math]::Round(($sucessos / $usersWithOldAlias) * 100, 1)
    Write-Host "`nTaxa de sucesso: $taxaSucesso% ($sucessos de $usersWithOldAlias usuários que precisavam alteração)" -ForegroundColor Cyan
}

# Mostrar alguns exemplos de sucesso
if ($sucessos -gt 0) {
    Write-Host "`n✅ Exemplos de alterações bem-sucedidas:" -ForegroundColor Green
    $logEntries | Where-Object { $_.Status -eq "Sucesso" } | Select-Object -First 5 | ForEach-Object {
        Write-Host "  - $($_.UserPrincipalName)" -ForegroundColor Green
        Write-Host "    Removido: $($_.AliasesRemovidos)" -ForegroundColor Gray
        Write-Host "    Adicionado: $($_.AliasesAdicionados)" -ForegroundColor Gray
    }
    
    if ($sucessos -gt 5) {
        Write-Host "  ... e mais $($sucessos - 5) usuários (veja o log completo)" -ForegroundColor Gray
    }
}

# Mostrar alguns erros se houver
if ($erros -gt 0) {
    Write-Host "`n❌ Primeiros erros encontrados:" -ForegroundColor Red
    $logEntries | Where-Object { $_.Status -eq "Erro" } | Select-Object -First 3 | ForEach-Object {
        Write-Host "  - $($_.UserPrincipalName): $($_.Erro)" -ForegroundColor Red
    }
    
    if ($erros -gt 3) {
        Write-Host "  ... e mais $($erros - 3) erros (veja o log de erros)" -ForegroundColor Gray
    }
}

Write-Host "`n========================================" -ForegroundColor Cyan
Write-Host "Processamento concluído!" -ForegroundColor Cyan
Write-Host "Logs salvos em:" -ForegroundColor White
Write-Host "  - $logFile" -ForegroundColor White
if ($errorMessages.Count -gt 0) {
    Write-Host "  - $errorLogFile" -ForegroundColor White
}
Write-Host "========================================" -ForegroundColor Cyan

# Desconectar do Exchange Online
Write-Host "`nDesconectando do Exchange Online..." -ForegroundColor Yellow
Disconnect-ExchangeOnline -Confirm:$false -ErrorAction SilentlyContinue
Write-Host "✓ Desconectado" -ForegroundColor Green