<#
.SYNOPSIS
    Script para criação de múltiplos usuários no Azure AD com atribuição de licença A1
.DESCRIPTION
    Este script cria usuários definidos diretamente no código e atribui licenças Office 365 A1
.AUTHOR
    DevOps Team
.VERSION
    2.0
#>

# ============================================
# CONFIGURAÇÕES GLOBAIS
# ============================================

$Config = @{
    TenantDomain = "seudominio.onmicrosoft.com"  # {{TENANT_DOMAIN}}
    DefaultPassword = "SenhaInicial@2025"         # {{DEFAULT_PASSWORD}}
    ForcePasswordChange = $true
    UsageLocation = "BR"                          # {{USAGE_LOCATION}}
    SendWelcomeEmail = $true
    LogPath = ".\user_creation_log_$(Get-Date -Format 'yyyyMMdd_HHmmss').txt"
}

# ============================================
# DEFINIÇÃO DOS USUÁRIOS A CRIAR
# ============================================
# Adicione ou remova usuários conforme necessário

$UsersToCreate = @(
    @{
        FirstName = "João"                        # {{USER1_FIRSTNAME}}
        LastName = "Silva"                        # {{USER1_LASTNAME}}
        DisplayName = "João Silva"
        UserPrincipalName = "joao.silva"          # {{USER1_LOGIN}}
        Department = "TI"
        JobTitle = "Analista de Sistemas"
        City = "São Paulo"
        Country = "Brasil"
        MobilePhone = "+55 11 98765-4321"
    },
    @{
        FirstName = "Maria"                       # {{USER2_FIRSTNAME}}
        LastName = "Santos"                       # {{USER2_LASTNAME}}
        DisplayName = "Maria Santos"
        UserPrincipalName = "maria.santos"        # {{USER2_LOGIN}}
        Department = "RH"
        JobTitle = "Coordenadora de RH"
        City = "Rio de Janeiro"
        Country = "Brasil"
        MobilePhone = "+55 21 98765-4322"
    },
    @{
        FirstName = "Pedro"                       # {{USER3_FIRSTNAME}}
        LastName = "Oliveira"                     # {{USER3_LASTNAME}}
        DisplayName = "Pedro Oliveira"
        UserPrincipalName = "pedro.oliveira"      # {{USER3_LOGIN}}
        Department = "Financeiro"
        JobTitle = "Assistente Financeiro"
        City = "Belo Horizonte"
        Country = "Brasil"
        MobilePhone = "+55 31 98765-4323"
    },
    @{
        FirstName = "Ana"                         # {{USER4_FIRSTNAME}}
        LastName = "Costa"                        # {{USER4_LASTNAME}}
        DisplayName = "Ana Costa"
        UserPrincipalName = "ana.costa"           # {{USER4_LOGIN}}
        Department = "Marketing"
        JobTitle = "Designer Gráfico"
        City = "Porto Alegre"
        Country = "Brasil"
        MobilePhone = "+55 51 98765-4324"
    },
    @{
        FirstName = "Carlos"                      # {{USER5_FIRSTNAME}}
        LastName = "Ferreira"                     # {{USER5_LASTNAME}}
        DisplayName = "Carlos Ferreira"
        UserPrincipalName = "carlos.ferreira"     # {{USER5_LOGIN}}
        Department = "TI"
        JobTitle = "Desenvolvedor Senior"
        City = "Brasília"
        Country = "Brasil"
        MobilePhone = "+55 61 98765-4325"
    }
)

# ============================================
# FUNÇÕES DE LOGGING
# ============================================

function Write-Log {
    param(
        [string]$Message,
        [string]$Level = "INFO"
    )
    
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $logMessage = "$timestamp [$Level] $Message"
    
    # Escrever no arquivo de log
    $logMessage | Out-File -FilePath $Config.LogPath -Append
    
    # Exibir no console com cores
    switch ($Level) {
        "SUCCESS" { Write-Host $Message -ForegroundColor Green }
        "ERROR"   { Write-Host $Message -ForegroundColor Red }
        "WARNING" { Write-Host $Message -ForegroundColor Yellow }
        "INFO"    { Write-Host $Message -ForegroundColor Cyan }
        default   { Write-Host $Message }
    }
}

# ============================================
# INSTALAÇÃO E IMPORTAÇÃO DE MÓDULOS
# ============================================

Write-Log "🔧 Verificando módulos necessários..." "INFO"

# Função para instalar módulos
function Install-RequiredModule {
    param([string]$ModuleName)
    
    if (!(Get-Module -ListAvailable -Name $ModuleName)) {
        Write-Log "📦 Instalando módulo $ModuleName..." "WARNING"
        try {
            Install-Module $ModuleName -Force -AllowClobber -Scope CurrentUser
            Write-Log "✅ Módulo $ModuleName instalado com sucesso" "SUCCESS"
        }
        catch {
            Write-Log "❌ Erro ao instalar módulo $ModuleName: $_" "ERROR"
            exit 1
        }
    }
}

# Instalar módulos necessários
Install-RequiredModule "AzureAD"
Install-RequiredModule "MSOnline"

# Importar módulos
Import-Module AzureAD -ErrorAction SilentlyContinue
Import-Module MSOnline -ErrorAction SilentlyContinue

# ============================================
# CONEXÃO COM AZURE AD
# ============================================

Write-Log "`n🔐 Conectando ao Azure AD..." "INFO"

try {
    # Conectar ao Azure AD
    $AzureADConnection = Connect-AzureAD -ErrorAction Stop
    Write-Log "✅ Conectado ao tenant: $($AzureADConnection.TenantDomain)" "SUCCESS"
    
    # Conectar ao MSOnline para gerenciamento de licenças
    Connect-MsolService -ErrorAction Stop
    Write-Log "✅ Conectado ao MSOnline Service" "SUCCESS"
}
catch {
    Write-Log "❌ Erro ao conectar: $_" "ERROR"
    Write-Log "💡 Dica: Execute 'Connect-AzureAD' manualmente se necessário" "WARNING"
    exit 1
}

# ============================================
# FUNÇÃO PARA OBTER LICENÇA A1
# ============================================

function Get-A1License {
    Write-Log "`n🔍 Buscando licenças A1 disponíveis..." "INFO"
    
    $licenses = Get-MsolAccountSku
    
    # Possíveis SKUs de licença A1
    $a1SkuPatterns = @(
        "*A1*",
        "*STANDARDWOFFPACK_FACULTY*",
        "*STANDARDWOFFPACK_STUDENT*",
        "*STANDARDWOFFPACK_IW_FACULTY*",
        "*STANDARDWOFFPACK_IW_STUDENT*",
        "*M365EDU_A1*",
        "*OFFICESUBSCRIPTION_FACULTY*",
        "*OFFICESUBSCRIPTION_STUDENT*"
    )
    
    $a1License = $null
    foreach ($pattern in $a1SkuPatterns) {
        $found = $licenses | Where-Object { $_.SkuPartNumber -like $pattern }
        if ($found) {
            $a1License = $found[0]
            break
        }
    }
    
    if ($a1License) {
        $available = $a1License.ActiveUnits - $a1License.ConsumedUnits
        Write-Log "✅ Licença A1 encontrada: $($a1License.SkuPartNumber)" "SUCCESS"
        Write-Log "📊 Licenças disponíveis: $available de $($a1License.ActiveUnits)" "INFO"
        
        if ($available -lt $UsersToCreate.Count) {
            Write-Log "⚠️  Atenção: Apenas $available licenças disponíveis para $($UsersToCreate.Count) usuários" "WARNING"
        }
        
        return $a1License.AccountSkuId
    }
    else {
        Write-Log "❌ Nenhuma licença A1 encontrada" "ERROR"
        Write-Log "`n📋 Licenças disponíveis no tenant:" "INFO"
        
        foreach ($license in $licenses) {
            $available = $license.ActiveUnits - $license.ConsumedUnits
            Write-Log "   • $($license.SkuPartNumber): $available disponíveis" "INFO"
        }
        
        return $null
    }
}

# ============================================
# FUNÇÃO PARA CRIAR USUÁRIO
# ============================================

function Create-User {
    param(
        [hashtable]$UserData,
        [string]$Domain,
        [string]$Password,
        [string]$LicenseSku
    )
    
    $upn = "$($UserData.UserPrincipalName)@$Domain"
    
    Write-Log "`n👤 Processando usuário: $($UserData.DisplayName)" "INFO"
    
    try {
        # Verificar se usuário já existe
        $existingUser = Get-AzureADUser -Filter "userPrincipalName eq '$upn'" -ErrorAction SilentlyContinue
        
        if ($existingUser) {
            Write-Log "⚠️  Usuário já existe: $upn" "WARNING"
            return $existingUser
        }
        
        # Criar perfil de senha
        $PasswordProfile = New-Object -TypeName Microsoft.Open.AzureAD.Model.PasswordProfile
        $PasswordProfile.Password = $Password
        $PasswordProfile.ForceChangePasswordNextLogin = $Config.ForcePasswordChange
        
        # Criar usuário
        $newUserParams = @{
            DisplayName = $UserData.DisplayName
            GivenName = $UserData.FirstName
            Surname = $UserData.LastName
            UserPrincipalName = $upn
            MailNickName = $UserData.UserPrincipalName
            PasswordProfile = $PasswordProfile
            AccountEnabled = $true
            Department = $UserData.Department
            JobTitle = $UserData.JobTitle
            City = $UserData.City
            Country = $UserData.Country
            UsageLocation = $Config.UsageLocation
            Mobile = $UserData.MobilePhone
        }
        
        $newUser = New-AzureADUser @newUserParams
        Write-Log "✅ Usuário criado: $upn" "SUCCESS"
        
        # Aguardar sincronização
        Start-Sleep -Seconds 5
        
        # Atribuir licença se disponível
        if ($LicenseSku) {
            try {
                Set-MsolUserLicense -UserPrincipalName $upn -AddLicenses $LicenseSku
                Write-Log "✅ Licença A1 atribuída para: $upn" "SUCCESS"
            }
            catch {
                Write-Log "❌ Erro ao atribuir licença: $_" "ERROR"
            }
        }
        
        return $newUser
    }
    catch {
        Write-Log "❌ Erro ao criar usuário $upn : $_" "ERROR"
        return $null
    }
}

# ============================================
# FUNÇÃO PARA GERAR RELATÓRIO
# ============================================

function Generate-Report {
    param(
        [array]$CreatedUsers,
        [array]$FailedUsers
    )
    
    Write-Log "`n📊 RELATÓRIO FINAL" "INFO"
    Write-Log "=================" "INFO"
    
    Write-Log "`n✅ Usuários criados com sucesso: $($CreatedUsers.Count)" "SUCCESS"
    foreach ($user in $CreatedUsers) {
        Write-Log "   • $($user.DisplayName) - $($user.UserPrincipalName)" "INFO"
    }
    
    if ($FailedUsers.Count -gt 0) {
        Write-Log "`n❌ Falhas na criação: $($FailedUsers.Count)" "ERROR"
        foreach ($user in $FailedUsers) {
            Write-Log "   • $($user.DisplayName)" "ERROR"
        }
    }
    
    # Criar arquivo com credenciais
    $credentialsFile = ".\credenciais_$(Get-Date -Format 'yyyyMMdd_HHmmss').txt"
    $credentials = @"
===============================================
CREDENCIAIS DOS NOVOS USUÁRIOS
Data: $(Get-Date -Format 'dd/MM/yyyy HH:mm:ss')
===============================================

Senha inicial para todos: $($Config.DefaultPassword)
Forçar mudança de senha: $($Config.ForcePasswordChange)

USUÁRIOS CRIADOS:
-----------------
"@
    
    foreach ($user in $CreatedUsers) {
        $credentials += @"

Nome: $($user.DisplayName)
Email: $($user.UserPrincipalName)
Departamento: $($user.Department)
Cargo: $($user.JobTitle)
"@
    }
    
    $credentials | Out-File -FilePath $credentialsFile
    Write-Log "`n📄 Arquivo de credenciais salvo em: $credentialsFile" "SUCCESS"
}

# ============================================
# EXECUÇÃO PRINCIPAL
# ============================================

Write-Log "🚀 INICIANDO CRIAÇÃO DE USUÁRIOS EM MASSA" "INFO"
Write-Log "==========================================" "INFO"
Write-Log "Total de usuários a criar: $($UsersToCreate.Count)" "INFO"

# Obter licença A1
$licenseSku = Get-A1License

if (!$licenseSku) {
    Write-Log "⚠️  Continuando sem atribuição de licença A1" "WARNING"
    $continue = Read-Host "Deseja continuar sem licenças? (S/N)"
    if ($continue -ne 'S') {
        Write-Log "Operação cancelada pelo usuário" "WARNING"
        exit 0
    }
}

# Arrays para relatório
$createdUsers = @()
$failedUsers = @()

# Criar cada usuário
foreach ($userData in $UsersToCreate) {
    $result = Create-User -UserData $userData -Domain $Config.TenantDomain -Password $Config.DefaultPassword -LicenseSku $licenseSku
    
    if ($result) {
        $createdUsers += $result
    }
    else {
        $failedUsers += $userData
    }
    
    # Pequena pausa entre criações
    Start-Sleep -Seconds 2
}

# Gerar relatório final
Generate-Report -CreatedUsers $createdUsers -FailedUsers $failedUsers

Write-Log "`n✅ PROCESSO CONCLUÍDO!" "SUCCESS"
Write-Log "Log completo salvo em: $($Config.LogPath)" "INFO"