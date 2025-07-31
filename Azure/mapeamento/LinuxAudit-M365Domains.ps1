#Requires -Modules ExchangeOnlineManagement

<#
.SYNOPSIS
    Script de auditoria de domínios Microsoft 365 - Versão Linux
.DESCRIPTION
    Versão otimizada para PowerShell Core no Linux, usando apenas Exchange Online Management
.PARAMETER DominiosParaRemover
    Array com os domínios que serão auditados
.PARAMETER ExportPath
    Caminho onde os relatórios serão salvos
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory=$false)]
    [string[]]$DominiosParaRemover = @("[COLOQUE O DOMINIO AQUI]", "[COLOQUE O DOMINIO AQUI]"),
    
    [Parameter(Mandatory=$false)]
    [string]$ExportPath = "./M365_Domain_Audit_$(Get-Date -Format 'yyyyMMdd_HHmmss')"
)

# Configurações iniciais
$ErrorActionPreference = "Continue"
$InformationPreference = "Continue"
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

# Criar diretório para exportação
if (!(Test-Path $ExportPath)) {
    New-Item -ItemType Directory -Path $ExportPath | Out-Null
}

# Função para log
function Write-Log {
    param($Message, $Type = "INFO")
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $logMessage = "[$timestamp] [$Type] $Message"
    Write-Host $logMessage -ForegroundColor $(if($Type -eq "ERROR"){"Red"}elseif($Type -eq "WARNING"){"Yellow"}else{"Green"})
    Add-Content -Path "$ExportPath/audit_log.txt" -Value $logMessage
}

# Função para exportar dados
function Export-DataToCSV {
    param($Data, $FileName, $Description)
    if ($Data -and $Data.Count -gt 0) {
        $Data | Export-Csv "$ExportPath/$FileName" -NoTypeInformation -Encoding utf8
        Write-Log "$($Data.Count) $Description encontrados e exportados para $FileName"
    } else {
        Write-Log "Nenhum $Description encontrado" -Type "WARNING"
    }
}

# Início do script
Write-Log "=== INICIANDO AUDITORIA DE DOMÍNIOS MICROSOFT 365 (Versão Linux) ==="
Write-Log "Domínios a serem auditados: $($DominiosParaRemover -join ', ')"

# Conectar ao Exchange Online
try {
    Write-Log "Conectando ao Exchange Online..."
    
    # Verificar se já está conectado
    $existingSession = Get-PSSession | Where-Object { $_.ConfigurationName -eq "Microsoft.Exchange" -and $_.State -eq "Opened" }
    
    if (-not $existingSession) {
        Connect-ExchangeOnline -ShowBanner:$false -ErrorAction Stop
        Write-Log "Conectado ao Exchange Online com sucesso"
    } else {
        Write-Log "Já conectado ao Exchange Online"
    }
} catch {
    Write-Log "Erro ao conectar ao Exchange Online: $_" -Type "ERROR"
    Write-Log "Certifique-se de ter o módulo instalado: Install-Module -Name ExchangeOnlineManagement" -Type "ERROR"
    exit 1
}

# Criar regex para match de domínios
$domainRegex = ($DominiosParaRemover | ForEach-Object { [regex]::Escape($_) }) -join "|"

# Relatório resumo
$resumo = @{}

Write-Log "`n=== 1. CAIXAS DE CORREIO E USUÁRIOS ==="

# 1.1 Todas as caixas de correio (substitui a busca por UPN do MSOnline)
Write-Log "Buscando caixas de correio com endereços nos domínios..."
$mailboxes = Get-Mailbox -ResultSize Unlimited | Where-Object {
    ($_.PrimarySmtpAddress -match $domainRegex) -or
    ($_.UserPrincipalName -match $domainRegex) -or
    ($_.EmailAddresses -match $domainRegex)
}

# Separar por tipo
$userMailboxes = $mailboxes | Where-Object { $_.RecipientTypeDetails -eq "UserMailbox" }
$sharedMailboxes = $mailboxes | Where-Object { $_.RecipientTypeDetails -eq "SharedMailbox" }
$resourceMailboxes = $mailboxes | Where-Object { $_.RecipientTypeDetails -in @("RoomMailbox", "EquipmentMailbox") }

# Exportar usuários com UPN ou email primário nos domínios
$userMailboxData = $userMailboxes | Select-Object DisplayName, UserPrincipalName, PrimarySmtpAddress,
    @{N="RecipientType";E={$_.RecipientTypeDetails}},
    @{N="EmailsNosDominios";E={
        ($_.EmailAddresses | Where-Object { $_ -match $domainRegex }) -join "; "
    }},
    @{N="IsInactiveMailbox";E={$_.IsInactiveMailbox}},
    @{N="AccountDisabled";E={$_.AccountDisabled}}

Export-DataToCSV -Data $userMailboxData -FileName "01_usuarios_mailboxes.csv" -Description "caixas de correio de usuários"
$resumo["Caixas de Usuários"] = $userMailboxes.Count

# 1.2 Aliases em todas as caixas
Write-Log "Identificando aliases nos domínios..."
$mailboxesComAlias = $mailboxes | Where-Object {
    $aliases = $_.EmailAddresses | Where-Object { $_ -match $domainRegex -and $_ -ne $_.PrimarySmtpAddress }
    $aliases.Count -gt 0
}

$mailboxAliasData = $mailboxesComAlias | Select-Object DisplayName, PrimarySmtpAddress, 
    @{N="AliasesNosDominios";E={
        ($_.EmailAddresses | Where-Object { $_ -match $domainRegex -and $_ -ne $_.PrimarySmtpAddress }) -join "; "
    }},
    @{N="TotalAliases";E={$_.EmailAddresses.Count}}

Export-DataToCSV -Data $mailboxAliasData -FileName "02_aliases_emails.csv" -Description "caixas com aliases"
$resumo["Caixas com Aliases"] = $mailboxesComAlias.Count

Write-Log "`n=== 2. GRUPOS ==="

# 2.1 Distribution Groups
Write-Log "Buscando Distribution Groups..."
$distGroups = Get-DistributionGroup -ResultSize Unlimited | Where-Object {
    ($_.PrimarySmtpAddress -match $domainRegex) -or
    ($_.EmailAddresses -match $domainRegex)
}
$distGroupData = $distGroups | Select-Object DisplayName, PrimarySmtpAddress, GroupType,
    @{N="EmailsNosDominios";E={
        ($_.EmailAddresses | Where-Object { $_ -match $domainRegex }) -join "; "
    }},
    @{N="ManagedBy";E={$_.ManagedBy -join "; "}}
Export-DataToCSV -Data $distGroupData -FileName "03_distribution_groups.csv" -Description "distribution groups"
$resumo["Distribution Groups"] = $distGroups.Count

# 2.2 Microsoft 365 Groups
Write-Log "Buscando Microsoft 365 Groups..."
$unifiedGroups = Get-UnifiedGroup -ResultSize Unlimited -ErrorAction SilentlyContinue | Where-Object {
    ($_.PrimarySmtpAddress -match $domainRegex) -or
    ($_.EmailAddresses -match $domainRegex)
}
if ($unifiedGroups) {
    $unifiedGroupData = $unifiedGroups | Select-Object DisplayName, PrimarySmtpAddress,
        @{N="EmailsNosDominios";E={
            ($_.EmailAddresses | Where-Object { $_ -match $domainRegex }) -join "; "
        }},
        @{N="SharePointSiteUrl";E={$_.SharePointSiteUrl}},
        @{N="AccessType";E={$_.AccessType}}
    Export-DataToCSV -Data $unifiedGroupData -FileName "04_microsoft365_groups.csv" -Description "Microsoft 365 groups"
    $resumo["Microsoft 365 Groups"] = $unifiedGroups.Count
} else {
    $resumo["Microsoft 365 Groups"] = 0
}

# 2.3 Dynamic Distribution Groups
Write-Log "Buscando Dynamic Distribution Groups..."
$dynamicGroups = Get-DynamicDistributionGroup -ResultSize Unlimited | Where-Object {
    ($_.PrimarySmtpAddress -match $domainRegex) -or
    ($_.EmailAddresses -match $domainRegex)
}
Export-DataToCSV -Data $dynamicGroups -FileName "05_dynamic_distribution_groups.csv" -Description "dynamic distribution groups"
$resumo["Dynamic Distribution Groups"] = $dynamicGroups.Count

# 2.4 Mail-Enabled Security Groups
Write-Log "Buscando Mail-Enabled Security Groups..."
$mailSecurityGroups = Get-DistributionGroup -ResultSize Unlimited -RecipientTypeDetails MailUniversalSecurityGroup | Where-Object {
    ($_.PrimarySmtpAddress -match $domainRegex) -or
    ($_.EmailAddresses -match $domainRegex)
}
Export-DataToCSV -Data $mailSecurityGroups -FileName "06_mail_security_groups.csv" -Description "mail-enabled security groups"
$resumo["Mail-Enabled Security Groups"] = $mailSecurityGroups.Count

Write-Log "`n=== 3. RECURSOS E CAIXAS ESPECIAIS ==="

# 3.1 Shared Mailboxes (já capturadas anteriormente)
Export-DataToCSV -Data $sharedMailboxes -FileName "07_shared_mailboxes.csv" -Description "shared mailboxes"
$resumo["Shared Mailboxes"] = $sharedMailboxes.Count

# 3.2 Resource Mailboxes (já capturadas anteriormente)
Export-DataToCSV -Data $resourceMailboxes -FileName "08_resource_mailboxes.csv" -Description "resource mailboxes"
$resumo["Resource Mailboxes"] = $resourceMailboxes.Count

# 3.3 Mail Contacts
Write-Log "Buscando Mail Contacts..."
$mailContacts = Get-MailContact -ResultSize Unlimited | Where-Object {
    $_.ExternalEmailAddress -match $domainRegex
}
$contactData = $mailContacts | Select-Object DisplayName, ExternalEmailAddress, 
    @{N="EmailAddresses";E={$_.EmailAddresses -join "; "}}
Export-DataToCSV -Data $contactData -FileName "09_mail_contacts.csv" -Description "mail contacts"
$resumo["Mail Contacts"] = $mailContacts.Count

# 3.4 Mail Users
Write-Log "Buscando Mail Users..."
$mailUsers = Get-MailUser -ResultSize Unlimited | Where-Object {
    ($_.WindowsEmailAddress -match $domainRegex) -or
    ($_.EmailAddresses -match $domainRegex)
}
Export-DataToCSV -Data $mailUsers -FileName "10_mail_users.csv" -Description "mail users"
$resumo["Mail Users"] = $mailUsers.Count

# 3.5 Recipients (busca geral)
Write-Log "Buscando todos os recipients..."
$allRecipients = Get-Recipient -ResultSize Unlimited | Where-Object {
    ($_.PrimarySmtpAddress -match $domainRegex) -or
    ($_.EmailAddresses -match $domainRegex)
}
$recipientSummary = $allRecipients | Group-Object RecipientType | Select-Object Name, Count
Export-DataToCSV -Data $recipientSummary -FileName "11_all_recipients_summary.csv" -Description "resumo de recipients"

Write-Log "`n=== 4. CONFIGURAÇÕES DO EXCHANGE ==="

# 4.1 Accepted Domains
Write-Log "Verificando Accepted Domains..."
$acceptedDomains = Get-AcceptedDomain | Where-Object { $_.DomainName -in $DominiosParaRemover }
$acceptedDomainData = $acceptedDomains | Select-Object DomainName, DomainType, Default
Export-DataToCSV -Data $acceptedDomainData -FileName "12_accepted_domains.csv" -Description "accepted domains"
$resumo["Accepted Domains"] = $acceptedDomains.Count

# 4.2 Email Address Policies
Write-Log "Verificando Email Address Policies..."
$emailPolicies = Get-EmailAddressPolicy | Where-Object {
    $templates = $_.EnabledEmailAddressTemplates -join " "
    $templates -match $domainRegex
}
$policyData = $emailPolicies | Select-Object Name, Priority, 
    @{N="TemplatesComDominio";E={
        $templates = $_.EnabledEmailAddressTemplates -join "; "
        $templates
    }}
Export-DataToCSV -Data $policyData -FileName "13_email_address_policies.csv" -Description "email address policies"
$resumo["Email Address Policies"] = $emailPolicies.Count

# 4.3 Remote Domains
Write-Log "Verificando Remote Domains..."
$remoteDomains = Get-RemoteDomain | Where-Object { $_.DomainName -in $DominiosParaRemover }
Export-DataToCSV -Data $remoteDomains -FileName "14_remote_domains.csv" -Description "remote domains"
$resumo["Remote Domains"] = $remoteDomains.Count

# 4.4 Connectors
Write-Log "Verificando Connectors..."
$inboundConnectors = Get-InboundConnector | Where-Object {
    $senderDomains = $_.SenderDomains -join " "
    $senderDomains -match $domainRegex
}
$outboundConnectors = Get-OutboundConnector | Where-Object {
    $recipientDomains = $_.RecipientDomains -join " "
    $recipientDomains -match $domainRegex
}
Export-DataToCSV -Data $inboundConnectors -FileName "15_inbound_connectors.csv" -Description "inbound connectors"
Export-DataToCSV -Data $outboundConnectors -FileName "16_outbound_connectors.csv" -Description "outbound connectors"
$resumo["Inbound Connectors"] = $inboundConnectors.Count
$resumo["Outbound Connectors"] = $outboundConnectors.Count

Write-Log "`n=== 5. REGRAS E POLÍTICAS ==="

# 5.1 Transport Rules
Write-Log "Verificando Transport Rules..."
$transportRules = Get-TransportRule | Where-Object {
    $ruleText = ($_ | Format-List | Out-String)
    $ruleText -match $domainRegex
}
Export-DataToCSV -Data $transportRules -FileName "17_transport_rules.csv" -Description "transport rules"
$resumo["Transport Rules"] = $transportRules.Count

# 5.2 Journaling Rules
Write-Log "Verificando Journaling Rules..."
$journalRules = Get-JournalRule -ErrorAction SilentlyContinue | Where-Object {
    ($_.Recipient -match $domainRegex) -or
    ($_.JournalEmailAddress -match $domainRegex)
}
if ($journalRules) {
    Export-DataToCSV -Data $journalRules -FileName "18_journal_rules.csv" -Description "journal rules"
    $resumo["Journal Rules"] = $journalRules.Count
} else {
    $resumo["Journal Rules"] = 0
}

Write-Log "`n=== 6. VERIFICAÇÕES ADICIONAIS ==="

# 6.1 Public Folders (se habilitadas)
Write-Log "Verificando Public Folders..."
try {
    $publicFolders = Get-PublicFolder -ResultSize Unlimited -ErrorAction SilentlyContinue | 
        Where-Object { $_.MailEnabled } |
        ForEach-Object {
            $pf = Get-MailPublicFolder $_.Identity -ErrorAction SilentlyContinue
            if ($pf.PrimarySmtpAddress -match $domainRegex) { $pf }
        }
    Export-DataToCSV -Data $publicFolders -FileName "19_public_folders.csv" -Description "mail-enabled public folders"
    $resumo["Public Folders"] = $publicFolders.Count
} catch {
    Write-Log "Public Folders não disponíveis ou não configuradas" -Type "WARNING"
    $resumo["Public Folders"] = "N/A"
}

# 6.2 Retention Policies
Write-Log "Verificando Retention Policies..."
$retentionPolicies = Get-RetentionPolicy -ErrorAction SilentlyContinue
if ($retentionPolicies) {
    Export-DataToCSV -Data $retentionPolicies -FileName "20_retention_policies.csv" -Description "retention policies"
    $resumo["Retention Policies"] = $retentionPolicies.Count
} else {
    $resumo["Retention Policies"] = 0
}

Write-Log "`n=== GERANDO RELATÓRIO DE RESUMO ==="

# Criar relatório HTML
$htmlReport = @"
<!DOCTYPE html>
<html>
<head>
    <title>Relatório de Auditoria de Domínios - Microsoft 365</title>
    <style>
        body { font-family: Arial, sans-serif; margin: 20px; background-color: #f5f5f5; }
        .container { max-width: 1200px; margin: 0 auto; background-color: white; padding: 20px; box-shadow: 0 0 10px rgba(0,0,0,0.1); }
        h1 { color: #0078d4; border-bottom: 3px solid #0078d4; padding-bottom: 10px; }
        h2 { color: #106ebe; margin-top: 30px; }
        table { border-collapse: collapse; width: 100%; margin-top: 20px; }
        th, td { border: 1px solid #ddd; padding: 12px; text-align: left; }
        th { background-color: #0078d4; color: white; }
        tr:nth-child(even) { background-color: #f2f2f2; }
        tr:hover { background-color: #e8f4f8; }
        .warning { color: #ff6b00; font-weight: bold; }
        .critical { color: #d13438; font-weight: bold; }
        .success { color: #107c10; font-weight: bold; }
        .info { color: #0078d4; }
        .summary-box { background-color: #f3f2f1; padding: 20px; border-radius: 5px; margin: 20px 0; border-left: 5px solid #0078d4; }
        .stat-card { display: inline-block; margin: 10px; padding: 20px; border: 1px solid #ddd; border-radius: 5px; min-width: 200px; text-align: center; }
        .stat-number { font-size: 2em; font-weight: bold; color: #0078d4; }
        .note { background-color: #fff4ce; padding: 15px; border-radius: 5px; margin: 20px 0; border-left: 5px solid #ffb900; }
        pre { background-color: #f5f5f5; padding: 15px; border-radius: 5px; overflow-x: auto; }
        .footer { margin-top: 50px; padding-top: 20px; border-top: 1px solid #ddd; text-align: center; color: #666; }
    </style>
</head>
<body>
    <div class="container">
        <h1>Relatório de Auditoria de Domínios - Microsoft 365</h1>
        <p><strong>Data:</strong> $(Get-Date -Format "dd/MM/yyyy HH:mm:ss")</p>
        <p><strong>Domínios Auditados:</strong> <span class="info">$($DominiosParaRemover -join ', ')</span></p>
        <p><strong>Versão:</strong> Linux/PowerShell Core (sem MSOnline)</p>
        
        <div class="summary-box">
            <h2>Resumo Executivo</h2>
            <p>Total de objetos encontrados que referenciam os domínios: <span class="critical">$($resumo.Values | Where-Object {$_ -ne "N/A" -and $_ -ne 0} | Measure-Object -Sum | Select-Object -ExpandProperty Sum)</span></p>
            
            <div style="text-align: center; margin: 20px 0;">
                <div class="stat-card">
                    <div class="stat-number">$($resumo["Caixas de Usuários"] + $resumo["Shared Mailboxes"] + $resumo["Resource Mailboxes"])</div>
                    <div>Total de Caixas</div>
                </div>
                <div class="stat-card">
                    <div class="stat-number">$($resumo["Distribution Groups"] + $resumo["Microsoft 365 Groups"] + $resumo["Dynamic Distribution Groups"] + $resumo["Mail-Enabled Security Groups"])</div>
                    <div>Total de Grupos</div>
                </div>
                <div class="stat-card">
                    <div class="stat-number">$($resumo["Caixas com Aliases"])</div>
                    <div>Caixas com Aliases</div>
                </div>
            </div>
        </div>
        
        <h2>Detalhamento por Categoria</h2>
        <table>
            <tr>
                <th>Categoria</th>
                <th>Quantidade</th>
                <th>Status</th>
                <th>Arquivo CSV</th>
            </tr>
"@

$fileNumber = 1
foreach ($key in $resumo.Keys | Sort-Object) {
    $value = $resumo[$key]
    $status = if ($value -eq 0) { "<span class='success'>✓ Limpo</span>" } 
              elseif ($value -eq "N/A") { "<span class='warning'>⚠ Não verificado</span>" }
              else { "<span class='critical'>⚡ Requer ação</span>" }
    
    $fileName = "{0:D2}_{1}.csv" -f $fileNumber, ($key -replace " ", "_" -replace "-", "_").ToLower()
    $fileNumber++
    
    $htmlReport += @"
        <tr>
            <td><strong>$key</strong></td>
            <td style="text-align: center;">$value</td>
            <td style="text-align: center;">$status</td>
            <td><code>$fileName</code></td>
        </tr>
"@
}

$htmlReport += @"
        </table>
        
        <div class="note">
            <h3>⚠️ Importante - Limitações da Versão Linux</h3>
            <p>Esta versão não inclui verificações que dependem do módulo MSOnline:</p>
            <ul>
                <li>Service Principals / Enterprise Applications</li>
                <li>Detalhes de licenciamento de usuários</li>
                <li>Status de bloqueio de credenciais</li>
            </ul>
            <p>Para uma auditoria completa, execute este script em um ambiente Windows com PowerShell 5.1.</p>
        </div>
        
        <h2>Próximos Passos</h2>
        <ol>
            <li><strong>Revisar arquivos CSV:</strong> Abra cada arquivo para ver os objetos específicos</li>
            <li><strong>Planejar ações:</strong> Para cada objeto, decida:
                <ul>
                    <li>Migrar para novo domínio</li>
                    <li>Remover completamente</li>
                    <li>Manter (impedirá remoção do domínio)</li>
                </ul>
            </li>
            <li><strong>Executar correções:</strong> Use o script de remoção ou comandos manuais</li>
            <li><strong>Validar:</strong> Execute esta auditoria novamente após as correções</li>
        </ol>
        
        <h2>Comandos Úteis para Correção</h2>
        <pre>
# Alterar endereço primário de usuário
Set-Mailbox -Identity "user@olddomain.com" -WindowsEmailAddress "user@newdomain.com"

# Remover alias de email
Set-Mailbox -Identity "user" -EmailAddresses @{remove="smtp:user@olddomain.com"}

# Alterar endereço de grupo
Set-DistributionGroup -Identity "Group" -PrimarySmtpAddress "group@newdomain.com"

# Alterar UPN (requer Azure AD PowerShell no Windows)
# Set-MsolUserPrincipalName -UserPrincipalName old@domain.com -NewUserPrincipalName new@domain.com
        </pre>
        
        <div class="footer">
            <p>Relatório gerado por Audit-M365Domains.ps1 | Arquivos salvos em: <code>$ExportPath</code></p>
        </div>
    </div>
</body>
</html>
"@

$htmlReport | Out-File "$ExportPath/00_RELATORIO_RESUMO.html" -Encoding utf8
Write-Log "Relatório HTML gerado: 00_RELATORIO_RESUMO.html"

# Criar arquivo de resumo em texto
$textSummary = @"
RESUMO DA AUDITORIA DE DOMÍNIOS MICROSOFT 365
=============================================
Data: $(Get-Date -Format "dd/MM/yyyy HH:mm:ss")
Domínios: $($DominiosParaRemover -join ', ')
Versão: Linux/PowerShell Core

TOTAIS ENCONTRADOS:
"@

foreach ($key in $resumo.Keys | Sort-Object) {
    $textSummary += "`n$($key.PadRight(30)): $($resumo[$key])"
}

$textSummary += @"

OBJETOS QUE IMPEDEM A REMOÇÃO DOS DOMÍNIOS:
- Total: $($resumo.Values | Where-Object {$_ -ne "N/A" -and $_ -ne 0} | Measure-Object -Sum | Select-Object -ExpandProperty Sum)

AÇÕES NECESSÁRIAS:
1. Revisar cada arquivo CSV gerado
2. Planejar a migração ou remoção dos objetos
3. Executar as mudanças necessárias
4. Re-executar este script para validação

Todos os arquivos foram salvos em: $ExportPath

NOTA: Esta é a versão Linux que não inclui verificações do MSOnline.
Para auditoria completa, use Windows PowerShell 5.1.
"@

$textSummary | Out-File "$ExportPath/00_RESUMO.txt" -Encoding utf8

Write-Log "`n=== AUDITORIA CONCLUÍDA ==="
Write-Log "Todos os relatórios foram salvos em: $ExportPath"
Write-Log "Abra o arquivo 00_RELATORIO_RESUMO.html para visualizar o resumo completo"

# Desconectar
Write-Log "Desconectando do Exchange Online..."
Disconnect-ExchangeOnline -Confirm:$false -ErrorAction SilentlyContinue

Write-Host "`nAuditoria concluída com sucesso!" -ForegroundColor Green
Write-Host "Verifique os arquivos em: $ExportPath" -ForegroundColor Yellow
