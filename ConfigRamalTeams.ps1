Script Configurar ramal teams

# Instala os módulos necessários
Install-Module -Name PowerShellGet -Force -AllowClobber
Install-Module -Name MicrosoftTeams -Force -AllowClobber
Import-Module MicrosoftTeams
 
# Conecta ao Teams
Connect-MicrosoftTeams
 
# Cria a interface gráfica
Add-Type -AssemblyName System.Windows.Forms
 
$form = New-Object System.Windows.Forms.Form
$form.Text = "Configurar Ramal no Teams"
$form.Size = New-Object System.Drawing.Size(400,250)
$form.StartPosition = "CenterScreen"
 
# Campo de e-mail
$labelEmail = New-Object System.Windows.Forms.Label
$labelEmail.Text = "E-mail do usuário:"
$labelEmail.Location = New-Object System.Drawing.Point(20,20)
$labelEmail.Size = New-Object System.Drawing.Size(120,20)
$form.Controls.Add($labelEmail)
 
$textEmail = New-Object System.Windows.Forms.TextBox
$textEmail.Location = New-Object System.Drawing.Point(150,18)
$textEmail.Size = New-Object System.Drawing.Size(200,20)
$form.Controls.Add($textEmail)
 
# Campo de número
$labelPhone = New-Object System.Windows.Forms.Label
$labelPhone.Text = "Número (ex: +5521...)"
$labelPhone.Location = New-Object System.Drawing.Point(20,60)
$labelPhone.Size = New-Object System.Drawing.Size(130,20)
$form.Controls.Add($labelPhone)
 
$textPhone = New-Object System.Windows.Forms.TextBox
$textPhone.Location = New-Object System.Drawing.Point(150,58)
$textPhone.Size = New-Object System.Drawing.Size(200,20)
$form.Controls.Add($textPhone)
 
# Botão
$button = New-Object System.Windows.Forms.Button
$button.Location = New-Object System.Drawing.Point(200,100)
$button.Size = New-Object System.Drawing.Size(120,30)
$button.Text = "Configurar Ramal"
$form.Controls.Add($button)
 
# Função de verificação de conexão
function Test-TeamsConnection {
    try {
        Get-CsTenant | Out-Null
        return $true
    } catch {
        return $false
    }
}
 
# Ação do botão
$button.Add_Click({
    $email = $textEmail.Text
    $phone = $textPhone.Text
    $sip = "sip:$email"
 
    if ($email -eq "" -or $phone -eq "") {
      [System.Windows.Forms.MessageBox]::Show("Preencha todos os campos.","Atenção")
    }
    elseif (-not (Test-TeamsConnection)) {
      [System.Windows.Forms.MessageBox]::Show("Você não está conectado ao Teams.","Atenção")
    }
    else {
        try {
            Set-CsPhoneNumberAssignment -Identity $sip -PhoneNumberType DirectRouting -PhoneNumber $phone
            Set-CsPhoneNumberAssignment -Identity $sip -EnterpriseVoiceEnabled $true
            Grant-CsOnlineVoiceRoutingPolicy -Identity $sip -PolicyName "Local and LDN"  # Troque por nome genérico se for necessário
      [System.Windows.Forms.MessageBox]::Show("Ramal configurado com sucesso!","Sucesso")
        }
        catch {
      [System.Windows.Forms.MessageBox]::Show("Erro: $_","Erro")
        }
    }
})
 
# Exibe a janela
$form.Add_Shown({$form.Activate()})
[void]$form.ShowDialog()
