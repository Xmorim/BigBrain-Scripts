# ====================================
# DELETAR USUÁRIOS EM MASSA - AZURE AD
# COM VERIFICAÇÃO DE DESABILITAÇÃO
# ====================================

# Conectar ao Microsoft Graph
Connect-MgGraph -Scopes "User.ReadWrite.All", "AuditLog.Read.All"

# Verificar conexão
Get-MgContext

# ====================================
# CONFIGURAÇÕES
# ====================================
$csvPath = "C:\Temp\usuarios_deletar.csv"
$logPath = "C:\Temp\log_delecao_$(Get-Date -Format 'yyyyMMdd_HHmmss').txt"
$backupPath = "C:\Temp\backup_usuarios_$(Get-Date -Format 'yyyyMMdd_HHmmss').csv"

# Importar lista de usuários
$usuarios = Import-Csv -Path $csvPath

# Criar log inicial
"=== INÍCIO DA DELEÇÃO ===" | Out-File -FilePath $logPath
"Data/Hora: $(Get-Date)" | Out-File -FilePath $logPath -Append
"Total de usuários: $($usuarios.Count)" | Out-File -FilePath $logPath -Append
"" | Out-File -FilePath $logPath -Append

# ====================================
# BACKUP E VERIFICAÇÃO
# ====================================
Write-Host "🔍 Verificando usuários antes de deletar..." -ForegroundColor Yellow

$usuariosBackup = @()

foreach ($user in $usuarios) {
    try {
        $upn = $user.UserPrincipalName
        $mgUser = Get-MgUser -UserId $upn -Property Id,UserPrincipalName,DisplayName,Mail,JobTitle,Department,AccountEnabled
        
        $usuariosBackup += [PSCustomObject]@{
            UserPrincipalName = $mgUser.UserPrincipalName
            DisplayName       = $mgUser.DisplayName
            Email             = $mgUser.Mail
            JobTitle          = $mgUser.JobTitle
            Department        = $mgUser.Department
            AccountEnabled    = $mgUser.AccountEnabled
            ObjectId          = $mgUser.Id
        }
        
        Write-Host "✅ Encontrado: $upn" -ForegroundColor Green
    }
    catch {
        Write-Host "❌ Não encontrado: $upn" -ForegroundColor Red
        "ERRO - Usuário não encontrado: $upn" | Out-File -FilePath $logPath -Append
    }
}

# Salvar backup
$usuariosBackup | Export-Csv -Path $backupPath -NoTypeInformation -Encoding UTF8
Write-Host "`n💾 Backup salvo em: $backupPath" -ForegroundColor Cyan

# ====================================
# 🔒 VERIFICAR SE TODOS ESTÃO DESABILITADOS
# ====================================
Write-Host "`n🔒 VERIFICANDO STATUS DE HABILITAÇÃO DOS USUÁRIOS..." -ForegroundColor Yellow
Write-Host "━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━" -ForegroundColor Yellow

$usuariosHabilitados = @()
$usuariosDesabilitados = @()

foreach ($user in $usuariosBackup) {
    if ($user.AccountEnabled -eq $true) {
        $usuariosHabilitados += $user
        Write-Host "⚠️  HABILITADO: $($user.UserPrincipalName) - $($user.DisplayName)" -ForegroundColor Red
        "AVISO - Usuário HABILITADO encontrado: $($user.UserPrincipalName)" | Out-File -FilePath $logPath -Append
    }
    else {
        $usuariosDesabilitados += $user
        Write-Host "✅ Desabilitado: $($user.UserPrincipalName) - $($user.DisplayName)" -ForegroundColor Green
    }
}

Write-Host "━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━" -ForegroundColor Yellow

# ====================================
# DECISÃO: PROSSEGUIR OU CANCELAR
# ====================================
if ($usuariosHabilitados.Count -gt 0) {
    Write-Host "`n❌ OPERAÇÃO CANCELADA!" -ForegroundColor Red -BackgroundColor Yellow
    Write-Host "`n⚠️  ATENÇÃO: Foram encontrados $($usuariosHabilitados.Count) usuários HABILITADOS!" -ForegroundColor Red
    Write-Host "Por segurança, a operação foi cancelada automaticamente.`n" -ForegroundColor Red
    
    Write-Host "📋 Lista de usuários que ainda estão HABILITADOS:" -ForegroundColor Yellow
    $usuariosHabilitados | Format-Table UserPrincipalName, DisplayName, Department, AccountEnabled -AutoSize
    
    Write-Host "🔧 AÇÃO NECESSÁRIA:" -ForegroundColor Yellow
    Write-Host "1. Desabilite os usuários acima no Entra ID" -ForegroundColor White
    Write-Host "   Portal: Entra Admin Center > Users > Selecionar usuário > Properties > Account enabled = NO" -ForegroundColor Gray
    Write-Host "2. Ou execute o script de desabilitação em massa" -ForegroundColor White
    Write-Host "3. Execute este script novamente após desabilitar todos os usuários`n" -ForegroundColor White
    
    # Salvar lista de usuários habilitados
    $arquivoHabilitados = "C:\Temp\usuarios_HABILITADOS_$(Get-Date -Format 'yyyyMMdd_HHmmss').csv"
    $usuariosHabilitados | Export-Csv -Path $arquivoHabilitados -NoTypeInformation -Encoding UTF8
    Write-Host "📁 Lista de usuários habilitados salva em: $arquivoHabilitados`n" -ForegroundColor Cyan
    
    "OPERAÇÃO CANCELADA - Usuários habilitados encontrados: $($usuariosHabilitados.Count)" | Out-File -FilePath $logPath -Append
    "Data/Hora Cancelamento: $(Get-Date)" | Out-File -FilePath $logPath -Append
    
    Disconnect-MgGraph
    exit
}
else {
    Write-Host "`n✅ Todos os usuários estão realmente desabilitados. ✔️" -ForegroundColor Green -BackgroundColor DarkGreen
    Write-Host "Total verificado: $($usuariosDesabilitados.Count) usuários" -ForegroundColor Green
    "VERIFICAÇÃO OK - Todos os usuários estão desabilitados" | Out-File -FilePath $logPath -Append
}

# ====================================
# CONFIRMAR ANTES DE DELETAR
# ====================================
Write-Host "`n⚠️  ATENÇÃO: Você está prestes a deletar $($usuariosBackup.Count) usuários!" -ForegroundColor Yellow
Write-Host "📁 Lista de usuários que serão deletados:" -ForegroundColor Yellow
$usuariosBackup | Format-Table UserPrincipalName, DisplayName, Department

$confirmacao = Read-Host "`nDigite 'SIM' para confirmar a deleção"

if ($confirmacao -ne "SIM") {
    Write-Host "❌ Operação cancelada pelo usuário." -ForegroundColor Red
    "Operação cancelada pelo usuário" | Out-File -FilePath $logPath -Append
    Disconnect-MgGraph
    exit
}

# ====================================
# DELETAR USUÁRIOS
# ====================================
Write-Host "`n🗑️  Iniciando deleção..." -ForegroundColor Red

$sucessos = 0
$erros = 0

foreach ($user in $usuarios) {
    try {
        $upn = $user.UserPrincipalName
        
        # Deletar usuário (vai para lixeira - recuperável por 30 dias)
        Remove-MgUser -UserId $upn -Confirm:$false
        
        Write-Host "✅ DELETADO: $upn" -ForegroundColor Green
        "SUCESSO - Deletado: $upn - $(Get-Date)" | Out-File -FilePath $logPath -Append
        $sucessos++
        
        # Pausa de 1 segundo entre cada deleção (evitar throttling)
        Start-Sleep -Seconds 1
    }
    catch {
        Write-Host "❌ ERRO ao deletar: $upn - $($_.Exception.Message)" -ForegroundColor Red
        "ERRO - $upn : $($_.Exception.Message)" | Out-File -FilePath $logPath -Append
        $erros++
    }
}

# ====================================
# RELATÓRIO FINAL
# ====================================
Write-Host "`n================================" -ForegroundColor Cyan
Write-Host "✅ Usuários deletados: $sucessos" -ForegroundColor Green
Write-Host "❌ Erros: $erros" -ForegroundColor Red
Write-Host "📁 Backup salvo em: $backupPath" -ForegroundColor Cyan
Write-Host "📄 Log salvo em: $logPath" -ForegroundColor Cyan
Write-Host "================================`n" -ForegroundColor Cyan

"" | Out-File -FilePath $logPath -Append
"=== RESUMO ===" | Out-File -FilePath $logPath -Append
"Sucessos: $sucessos" | Out-File -FilePath $logPath -Append
"Erros: $erros" | Out-File -FilePath $logPath -Append
"Data/Hora Fim: $(Get-Date)" | Out-File -FilePath $logPath -Append

# Desconectar
Disconnect-MgGraph

Write-Host "⚠️  LEMBRETE: Usuários ficam na lixeira por 30 dias e podem ser restaurados!" -ForegroundColor Yellow
Write-Host "Para restaurar: Entra Admin Center > Users > Deleted users" -ForegroundColor Yellow