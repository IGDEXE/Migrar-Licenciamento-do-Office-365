# Migrar licenciamento do Office
# Ivo Dias
Clear-Host
# Funcao para validar as pastas
function Validar-Pasta {
    param (
        [parameter(position=0,Mandatory=$True)]
        $caminho
    )
    # Verifica se ja existe
    $Existe = Test-Path -Path $caminho
    # Cria a pasta
    if ($Existe -eq $false) {
        Write-Host "Configurando pasta: $caminho"
        try {
            $noReturn = New-Item -ItemType directory -Path $caminho # Cria a pasta
            Write-Host "Pasta configurada com sucesso"
        }
        catch {
            $ErrorMessage = $_.Exception.Message # Recebe o erro
            Write-Host "Ocorreu um erro durante a configuracao da pasta" # Mostra a mensagem
            Write-Host "Erro: $ErrorMessage"
        }
    }
}

# Credencial
$userADM = $env:UserName # Recebe o usuario
$userADM += '@sinqia.com.br' # Configura o e-mail
$LiveCred = Get-Credential -Message "Informe as credenciais de Administrador do Office 365" -UserName $userADM # Recebe as credenciais

# Conecta no Office
Write-Host "Conectando ao Office 365"
try {
    $Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri 'https://ps.outlook.com/powershell/' -Credential $LiveCred -Authentication Basic -AllowRedirection 
    Import-PSSession $Session # Importa a secao
    Connect-MsolService -Credential $LiveCred # Conecta a secao
    Clear-Host # Limpa a tela
}
catch {
    $ErrorMessage = $_.Exception.Message # Recebe a mensagem de erro
    Write-Host "Erro ao conectar: $ErrorMessage" # Mostra a mensagem
    Pause # Pausa para exibicao
    exit # Fecha o script
}

# Recebe o caminho do arquivo
$caminhoTxT = "$PSScriptRoot\O365\ListaAlteracaoLicenciamento.txt"

# Recebe os nomes
$usuarios = Get-Content $caminhoTxT

# Configuracoes
$LogPath = "$env:USERPROFILE\Desktop" # Define o desktop como local para os LOGs
$identificacao = Get-Date -Format LOG@ddMMyyyy # Cria um hash de identificacao
# Os codigos de licenciamento podem ser consultados em:
# https://docs.microsoft.com/pt-br/azure/active-directory/users-groups-roles/licensing-service-plan-reference
$LicencaAntiga = "ATTPS:ENTERPRISEPACK" # Esse eh o codigo da E3
$LicencaNova = "ATTPS:O365_BUSINESS_PREMIUM" # esse eh o da Business Premium

# Faz o procedimento
foreach ($usuario in $usuarios) {
    try {
        Write-Host "Fazendo o procedimento com o usuario: $usuario"
        $usuario += "@sinqia.com.br"
        Set-MsolUserLicense -UserPrincipalName "$usuario" -AddLicenses "$LicencaNova" -RemoveLicenses "$LicencaAntiga" # Faz a troca
        Add-Content -Path "$LogPath\migrarLicenciamento.$identificacao.txt" -Value "Usuario: $usuario - Licenciamento: $LicencaNova"
        Write-Host "Deu certo"
    }
    catch {
        $ErrorMessage = $_.Exception.Message # Recebe a mensagem de erro
        Add-Content -Path "$LogPath\migrarLicenciamento.$identificacao.txt" -Value "Usuario: $usuario - Erro: $ErrorMessage"
        Write-Host "Ocorreu um erro: $ErrorMessage"
    }
}

# Encerra
Clear-Host
Write-Host "Procedimento concluido"
Write-Host "Mais detalhes em: $LogPath\migrarLicenciamento.$identificacao.txt"
Pause