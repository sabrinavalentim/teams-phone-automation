# teams-phone-automation
Scripts Powershell com interface gráfica para configurar e remover ramais no Microsoft Teams

Install-Module -Name PowerShellGet -Force -AllowClobber
Install-Module -Name MicrosoftTeams -Force -AllowClobber
Import-Module MicrosoftTeams
Connect-MicrosoftTeams
 
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
# Janela principal
$form = New-Object System.Windows.Forms.Form
$form.Text = "Remover Ramal no Teams"
$form.Size = New-Object System.Drawing.Size(400,250)
$form.StartPosition = "CenterScreen"
# Label e campo para e-mail
$emailLabel = New-Object System.Windows.Forms.Label
$emailLabel.Location = New-Object System.Drawing.Point(20,20)
$emailLabel.Size = New-Object System.Drawing.Size(120,20)
$emailLabel.Text = "E-mail do usuario:"
$form.Controls.Add($emailLabel)
$emailBox = New-Object System.Windows.Forms.TextBox
$emailBox.Location = New-Object System.Drawing.Point(150,18)
$emailBox.Size = New-Object System.Drawing.Size(200,20)
$form.Controls.Add($emailBox)
# Label e campo para número
$numeroLabel = New-Object System.Windows.Forms.Label
$numeroLabel.Location = New-Object System.Drawing.Point(20,60)
$numeroLabel.Size = New-Object System.Drawing.Size(130,20)
$numeroLabel.Text = "Numero (ex: +5521...):"
$form.Controls.Add($numeroLabel)
$numeroBox = New-Object System.Windows.Forms.TextBox
$numeroBox.Location = New-Object System.Drawing.Point(150,58)
$numeroBox.Size = New-Object System.Drawing.Size(200,20)
$form.Controls.Add($numeroBox)
# Botão remover ramal
$removeButton = New-Object System.Windows.Forms.Button
$removeButton.Location = New-Object System.Drawing.Point(200,100)
$removeButton.Size = New-Object System.Drawing.Size(120,30)
$removeButton.Text = "Remover Ramal"
$removeButton.Add_Click({
    $email = $emailBox.Text
    $numero = $numeroBox.Text
    if ([string]::IsNullOrWhiteSpace($email) -or [string]::IsNullOrWhiteSpace($numero)) {
        [System.Windows.Forms.MessageBox]::Show("Preencha todos os campos.","Atenção")
        return
    }
    try {
        Remove-CsPhoneNumberAssignment -Identity $email -PhoneNumber $numero -PhoneNumberType DirectRouting -ErrorAction Stop
        [System.Windows.Forms.MessageBox]::Show("Ramal $numero removido com sucesso do usuário $email.","Sucesso")
    }
    catch {
        [System.Windows.Forms.MessageBox]::Show("Erro ao remover o ramal. Verifique se está conectado ao Teams e se os dados estão corretos.`n`nErro: $_","Erro")
    }
})
$form.Controls.Add($removeButton)
# Abrir janela
$form.Add_Shown({$form.Activate()})
[void]$form.ShowDialog()
