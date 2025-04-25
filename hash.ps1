Add-Type -AssemblyName System.Windows.Forms

# --- Récupération des infos poste ---
$serial = (Get-WmiObject -Class Win32_BIOS).SerialNumber
$productId = (Get-ItemProperty "HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion").ProductId

# --- Hardware Hash : WMI MDM ou MdmDiagnosticsTool ---
$hash = $null
try {
    $client = Get-WmiObject -Namespace root\cimv2\mdm\dmmap -Class MDM_DevDetail_Ext01 -ErrorAction Stop
    $hash = $client.DeviceHardwareData
} catch {
    $csvFile = "$env:TEMP\AutoPilotHWID.csv"
    Start-Process -FilePath "C:\Windows\System32\MdmDiagnosticsTool.exe" -ArgumentList "-autopilot", "-output", "$csvFile" -Wait
    if (Test-Path $csvFile) {
        $csv = Import-Csv $csvFile
        $hash = $csv[0].'Hardware Hash'
        Remove-Item $csvFile -ErrorAction SilentlyContinue
    }
}

# --- Liste des Group Tag prédéfinis ---
$groupTags = @(
    "VIP-Devices",
    "Direction",
    "RH",
    "Comptabilité",
    "IT",
    "Stagiaire",
    "Aucun",
    "Personnalisé"
)

# --- Création du formulaire principal ---
$form = New-Object System.Windows.Forms.Form
$form.Text = "Infos Autopilot, Webhook & QR Code"
$form.Size = New-Object System.Drawing.Size(520,500)
$form.StartPosition = "CenterScreen"

# Champs affichage
$lblSerial = New-Object System.Windows.Forms.Label
$lblSerial.Text = "Serial Number : $serial"
$lblSerial.AutoSize = $true
$lblSerial.Location = New-Object System.Drawing.Point(10,20)
$form.Controls.Add($lblSerial)

$lblProduct = New-Object System.Windows.Forms.Label
$lblProduct.Text = "Windows Product ID : $productId"
$lblProduct.AutoSize = $true
$lblProduct.Location = New-Object System.Drawing.Point(10,50)
$form.Controls.Add($lblProduct)

$lblHash = New-Object System.Windows.Forms.Label
$lblHash.Text = "Hardware Hash : " + ($(if ($hash) { "OK" } else { "Non disponible" }))
$lblHash.AutoSize = $true
$lblHash.Location = New-Object System.Drawing.Point(10,80)
$form.Controls.Add($lblHash)

# Combobox Group Tag
$lblGroup = New-Object System.Windows.Forms.Label
$lblGroup.Text = "Group Tag :"
$lblGroup.Location = New-Object System.Drawing.Point(10,120)
$lblGroup.Size = New-Object System.Drawing.Size(80,20)
$form.Controls.Add($lblGroup)

$cmbGroup = New-Object System.Windows.Forms.ComboBox
$cmbGroup.Location = New-Object System.Drawing.Point(100,115)
$cmbGroup.Size = New-Object System.Drawing.Size(200,20)
$cmbGroup.Items.AddRange($groupTags)
$cmbGroup.SelectedIndex = 0
$form.Controls.Add($cmbGroup)

# Champ texte pour Group Tag personnalisé (masqué par défaut)
$txtGroup = New-Object System.Windows.Forms.TextBox
$txtGroup.Location = New-Object System.Drawing.Point(310,115)
$txtGroup.Size = New-Object System.Drawing.Size(150,20)
$txtGroup.Visible = $false
$form.Controls.Add($txtGroup)

# Label UPN
$lblUpn = New-Object System.Windows.Forms.Label
$lblUpn.Text = "User Principal Name (UPN) :"
$lblUpn.Location = New-Object System.Drawing.Point(10,160)
$lblUpn.Size = New-Object System.Drawing.Size(180,20)
$form.Controls.Add($lblUpn)

# TextBox UPN
$txtUpn = New-Object System.Windows.Forms.TextBox
$txtUpn.Location = New-Object System.Drawing.Point(190,155)
$txtUpn.Size = New-Object System.Drawing.Size(210,20)
$form.Controls.Add($txtUpn)

# Affichage résultat JSON
$txtResult = New-Object System.Windows.Forms.TextBox
$txtResult.Multiline = $true
$txtResult.ScrollBars = "Vertical"
$txtResult.Location = New-Object System.Drawing.Point(10,190)
$txtResult.Size = New-Object System.Drawing.Size(470,120)
$form.Controls.Add($txtResult)

# Bouton ENVOYER
$btnSend = New-Object System.Windows.Forms.Button
$btnSend.Text = "Envoyer au Webhook"
$btnSend.Location = New-Object System.Drawing.Point(10,330)
$btnSend.Size = New-Object System.Drawing.Size(200,30)
$form.Controls.Add($btnSend)

# Bouton QR Code
$btnQR = New-Object System.Windows.Forms.Button
$btnQR.Text = "Afficher QR Code"
$btnQR.Location = New-Object System.Drawing.Point(220,330)
$btnQR.Size = New-Object System.Drawing.Size(200,30)
$form.Controls.Add($btnQR)

# Label pour afficher la réponse
$lblResp = New-Object System.Windows.Forms.Label
$lblResp.Location = New-Object System.Drawing.Point(10,370)
$lblResp.Size = New-Object System.Drawing.Size(470,60)
$lblResp.AutoSize = $false
$form.Controls.Add($lblResp)

# Affichage du champ personnalisé si besoin
$cmbGroup.Add_SelectedIndexChanged({
    if ($cmbGroup.SelectedItem -eq "Personnalisé") {
        $txtGroup.Visible = $true
    } else {
        $txtGroup.Visible = $false
    }
})

# Action ENVOYER
$btnSend.Add_Click({
    $grp = $cmbGroup.SelectedItem
    if ($grp -eq "Aucun") { $grp = "" }
    if ($grp -eq "Personnalisé") { $grp = $txtGroup.Text }
    $upn = $txtUpn.Text

    $payload = [PSCustomObject]@{
        "Device Serial Number"  = $serial
        "Windows Product ID"    = $productId
        "Hardware Hash"         = $hash
        "Group Tag"             = $grp
        "User Principal Name"   = $upn
    }

    $json = $payload | ConvertTo-Json
    $txtResult.Text = $json

    # Webhook à personnaliser ici !
    $webhookUrl = "https://TON_WEBHOOK_URL"

    if ($hash -and $upn) {
        try {
            $resp = Invoke-RestMethod -Uri $webhookUrl -Method POST -Body $json -ContentType "application/json"
            $lblResp.Text = "Webhook envoyé ! Réponse : $resp"
        } catch {
            $lblResp.Text = "Erreur lors de l'envoi : $_"
        }
    } else {
        $lblResp.Text = "UPN ou Hardware Hash manquant."
    }
})

# Action QR Code
$btnQR.Add_Click({
    $json = $txtResult.Text
    if (![string]::IsNullOrWhiteSpace($json)) {
        $encoded = [uri]::EscapeDataString($json)
        $qrUrl = "https://api.qrserver.com/v1/create-qr-code/?size=300x300&data=$encoded"
        $qrFile = "$env:TEMP\autopilot_qrcode.png"
        try {
            Invoke-WebRequest -Uri $qrUrl -OutFile $qrFile -UseBasicParsing
            # Afficher dans une nouvelle fenêtre WinForms
            $qrForm = New-Object System.Windows.Forms.Form
            $qrForm.Text = "QR Code"
            $qrForm.Size = New-Object System.Drawing.Size(340,340)
            $pic = New-Object System.Windows.Forms.PictureBox
            $pic.Image = [System.Drawing.Image]::FromFile($qrFile)
            $pic.SizeMode = "Zoom"
            $pic.Dock = "Fill"
            $qrForm.Controls.Add($pic)
            $qrForm.TopMost = $true
            $qrForm.ShowDialog()
        } catch {
            [System.Windows.Forms.MessageBox]::Show("Erreur lors de la génération ou de l'affichage du QR code.`n$_")
        }
    } else {
        [System.Windows.Forms.MessageBox]::Show("Aucun résultat JSON à encoder.")
    }
})

# Affichage du formulaire
[void]$form.ShowDialog()
