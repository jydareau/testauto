Add-Type -AssemblyName PresentationFramework

[xml]$xaml = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        Title="AutoPilot QR - OOBE"
        WindowStartupLocation="CenterScreen"
        Width="820" Height="520"
        Background="#FFF7F0" FontFamily="Segoe UI" ResizeMode="NoResize" Topmost="True">
    <Grid>
        <Grid.ColumnDefinitions>
            <ColumnDefinition Width="270"/>
            <ColumnDefinition Width="*"/>
        </Grid.ColumnDefinitions>
        <!-- Colonne de gauche (explication) -->
        <StackPanel Grid.Column="0" Background="#FFE4D6" Margin="0,0,0,0" VerticalAlignment="Stretch">
            <TextBlock Text="Bienvenue dans l’assistant AutoPilot !" Margin="24,38,12,12" 
                Foreground="#5D90E3" FontWeight="Bold" FontSize="22"/>
            <TextBlock TextWrapping="Wrap"
                Text="Cet outil collecte les informations de votre poste pour un enregistrement AutoPilot facilité :
- Numéro de série
- Marque et modèle
- UPN (identifiant utilisateur)
- GroupTag (affectation)
- Hardware Hash (pour Intune)
Vous pouvez ensuite envoyer ces infos via webhook ou générer un QR code pour validation."
                Margin="24,0,18,0" FontSize="16" Foreground="#282828"/>
            <TextBlock Text="Assurez-vous d’être connecté à Internet." Margin="24,28,10,0" FontSize="14" Foreground="#B87C46"/>
        </StackPanel>
        <!-- Colonne de droite (formulaire) -->
        <StackPanel Grid.Column="1" Margin="0,20,0,0" HorizontalAlignment="Center">
            <Image Name="LogoImage" Width="90" Height="90" Margin="0,10,0,2"/>
            <TextBlock Text="AutoPilot" HorizontalAlignment="Center" FontSize="18" Foreground="#282828" Margin="0,0,0,16"/>
            <Border Background="#fff" CornerRadius="28" Margin="18,0,18,18" Padding="24">
                <StackPanel>
                    <TextBlock Name="LblSerial" Text="Numéro de série : ..." Foreground="#272728" FontWeight="Bold" FontSize="22" Margin="0,0,0,8"/>
                    <TextBlock Name="LblBrand"  Text="Marque : ..." Foreground="#272728" FontWeight="Bold" FontSize="22" Margin="0,0,0,8"/>
                    <TextBlock Name="LblModel"  Text="Modèle : ..." Foreground="#272728" FontWeight="Bold" FontSize="22"/>
                </StackPanel>
            </Border>
            <TextBlock Text="Group Tag" Foreground="#5D90E3" FontWeight="Bold" FontSize="16" Margin="18,2,0,2"/>
            <ComboBox  Name="GroupTagCombo" Margin="18,0,18,10" FontSize="18" Background="#fff" Foreground="#272728" Height="40">
                <ComboBoxItem>Aucun</ComboBoxItem>
                <ComboBoxItem>VIP-Devices</ComboBoxItem>
                <ComboBoxItem>Direction</ComboBoxItem>
                <ComboBoxItem>RH</ComboBoxItem>
                <ComboBoxItem>Comptabilité</ComboBoxItem>
                <ComboBoxItem>IT</ComboBoxItem>
                <ComboBoxItem>Stagiaire</ComboBoxItem>
                <ComboBoxItem>Personnalisé</ComboBoxItem>
            </ComboBox>
            <TextBox Name="CustomGroupTag" Margin="18,0,18,10" Padding="9" FontSize="18" Background="#fff" Foreground="#272728" Visibility="Collapsed" Height="38"/>
            <TextBlock Text="UPN utilisateur" Foreground="#5D90E3" FontWeight="Bold" FontSize="16" Margin="18,0,0,2"/>
            <TextBox Name="UpnBox" Margin="18,0,18,10" Padding="9" FontSize="18" Background="#fff" Foreground="#272728" Height="38"/>
            <TextBlock Text="Token API/Bearer" Foreground="#5D90E3" FontWeight="Bold" FontSize="16" Margin="18,0,0,2"/>
            <PasswordBox Name="TokenBox" Margin="18,0,18,18" Padding="9" FontSize="18" Background="#fff" Foreground="#272728" Height="38"/>
            <StackPanel Orientation="Horizontal" HorizontalAlignment="Center" Margin="0,12,0,22">
                <Button Name="BtnSend" Content="Envoyer" Width="120" Height="46" Margin="7" Padding="7"
                    Background="#FFC6A0" Foreground="#272728" FontWeight="Bold" BorderThickness="0"  Cursor="Hand"/>
                <Button Name="BtnQR" Content="QR Code" Width="100" Height="46" Margin="7" Padding="7"
                    Background="#5D90E3" Foreground="#fff" FontWeight="Bold" BorderThickness="0" Cursor="Hand"/>
                <Button Name="BtnQuit" Content="Quitter" Width="85" Height="46" Margin="7" Padding="7"
                    Background="#FFD1B2" Foreground="#B81E1E" FontWeight="Bold" BorderThickness="0" Cursor="Hand"/>
            </StackPanel>
        </StackPanel>
    </Grid>
</Window>
"@

$reader = (New-Object System.Xml.XmlNodeReader $xaml)
$window = [Windows.Markup.XamlReader]::Load($reader)

# --- Afficher le logo depuis l'URL GitHub "raw"
$logoUrl = 'https://raw.githubusercontent.com/jydareau/testauto/main/autopilot-shield.png'
$logoImage = $window.FindName("LogoImage")
if ($logoImage) {
    try {
        $img = New-Object System.Windows.Media.Imaging.BitmapImage
        $img.BeginInit()
        $img.UriSource = $logoUrl
        $img.CacheOption = [System.Windows.Media.Imaging.BitmapCacheOption]::OnLoad
        $img.EndInit()
        $logoImage.Source = $img
    } catch {}
}

$lblSerial    = $window.FindName("LblSerial")
$lblBrand     = $window.FindName("LblBrand")
$lblModel     = $window.FindName("LblModel")
$groupTagCombo= $window.FindName("GroupTagCombo")
$customTagBox = $window.FindName("CustomGroupTag")
$upnBox       = $window.FindName("UpnBox")
$tokenBox     = $window.FindName("TokenBox")
$btnSend      = $window.FindName("BtnSend")
$btnQR        = $window.FindName("BtnQR")
$btnQuit      = $window.FindName("BtnQuit")

# --- Infos device (accentué, encodé en UTF-8)
$comp = Get-WmiObject -Class Win32_ComputerSystem
$manufacturer = $comp.Manufacturer
$model = $comp.Model
$serial = (Get-WmiObject -Class Win32_BIOS).SerialNumber

$lblSerial.Text = "Numéro de série : $serial"
$lblBrand.Text  = "Marque : $manufacturer"
$lblModel.Text  = "Modèle : $model"

# --- Champ perso et option "Aucun"
$groupTagCombo.Add_SelectionChanged({
    $selected = $groupTagCombo.SelectedItem.Content
    if ($selected -eq "Personnalisé") {
        $customTagBox.Visibility = "Visible"
    } else {
        $customTagBox.Visibility = "Collapsed"
    }
})

# --- Hardware hash
function Get-HardwareHash {
    try {
        $client = Get-WmiObject -Namespace root\cimv2\mdm\dmmap -Class MDM_DevDetail_Ext01 -ErrorAction Stop
        return $client.DeviceHardwareData
    } catch {
        $csvFile = "$env:TEMP\AutoPilotHWID.csv"
        Start-Process -FilePath "C:\Windows\System32\MdmDiagnosticsTool.exe" -ArgumentList "-autopilot", "-output", "$csvFile" -Wait
        if (Test-Path $csvFile) {
            $csv = Import-Csv $csvFile
            $hash = $csv[0].'Hardware Hash'
            Remove-Item $csvFile -ErrorAction SilentlyContinue
            return $hash
        }
    }
    return $null
}

$global:LastTicketId = ""
$global:LastSerial   = $serial

# --- Envoi Webhook
$btnSend.Add_Click({
    $groupTag = $groupTagCombo.SelectedItem.Content
    if ($groupTag -eq "Aucun") { $groupTag = "" }
    if ($groupTag -eq "Personnalisé") { $groupTag = $customTagBox.Text }
    $upn = $upnBox.Text
    $token = $tokenBox.Password

    $productId = (Get-ItemProperty "HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion").ProductId
    $hash = Get-HardwareHash

    $payload = [PSCustomObject]@{
        "Device Serial Number"  = $serial
        "Windows Product ID"    = $productId
        "Hardware Hash"         = $hash
        "Group Tag"             = $groupTag
        "User Principal Name"   = $upn
        "Manufacturer"          = $manufacturer
        "Model"                 = $model
    }

    $json = $payload | ConvertTo-Json

    $webhookUrl = "https://TON_WEBHOOK_URL"   # Personnalise ici !

    if ($hash -and $upn -and $token) {
        try {
            $headers = @{ "Authorization" = "Bearer $token" }
            $resp = Invoke-RestMethod -Uri $webhookUrl -Headers $headers -Method POST -Body $json -ContentType "application/json"
            if ($resp.ticketId) {
                $global:LastTicketId = $resp.ticketId
                [System.Windows.MessageBox]::Show("Numéro de ticket : $($resp.ticketId)`nScannez le QR code pour valider/importer.", "Envoyé", "OK", "Info")
            } else {
                [System.Windows.MessageBox]::Show("Webhook envoyé, réponse : $resp")
            }
        } catch {
            [System.Windows.MessageBox]::Show("Erreur lors de l'envoi : $_")
        }
    } else {
        [System.Windows.MessageBox]::Show("UPN, Hardware Hash ou Token manquant.")
    }
})

# --- QR code
$btnQR.Add_Click({
    $ticketId = $global:LastTicketId
    if ($ticketId -and $serial) {
        $qrObj = [PSCustomObject]@{
            "Ticket" = $ticketId
            "Serial" = $serial
            "Manufacturer" = $manufacturer
            "Model" = $model
        }
        $qrPayload = $qrObj | ConvertTo-Json -Compress
        $qrUrl = "https://api.qrserver.com/v1/create-qr-code/?size=400x400&data=$([uri]::EscapeDataString($qrPayload))"
        $qrFile = "$env:TEMP\ticketid_qrcode.png"
        try {
            Invoke-WebRequest -Uri $qrUrl -OutFile $qrFile -UseBasicParsing
            $img = [System.Windows.Media.Imaging.BitmapImage]::new()
            $stream = [System.IO.File]::OpenRead($qrFile)
            $img.BeginInit()
            $img.StreamSource = $stream
            $img.CacheOption = [System.Windows.Media.Imaging.BitmapCacheOption]::OnLoad
            $img.EndInit()
            $stream.Close()
            $qrWindow = New-Object Windows.Window
            $qrWindow.Title = "QR Code Ticket"
            $qrWindow.Width = 450
            $qrWindow.Height = 480
            $imgCtrl = New-Object System.Windows.Controls.Image
            $imgCtrl.Source = $img
            $imgCtrl.Margin = '24'
            $qrWindow.Content = $imgCtrl
            $qrWindow.WindowStartupLocation = "CenterScreen"
            $qrWindow.Topmost = $true
            $qrWindow.ShowDialog() | Out-Null
        } catch {
            [System.Windows.MessageBox]::Show("Erreur génération QR code : $_")
        }
    } else {
        [System.Windows.MessageBox]::Show("Aucun ticketId disponible. Envoie d'abord au webhook !")
    }
})

# --- Bouton Quitter (avec confirmation)
$btnQuit.Add_Click({
    $result = [System.Windows.MessageBox]::Show("Voulez-vous vraiment quitter ?", "Confirmer la fermeture", "YesNo", "Question")
    if ($result -eq "Yes") {
        $window.Close()
        Stop-Process -Id $PID
    }
})

# --- Affichage de la fenêtre
$window.Topmost = $true
$window.ShowDialog() | Out-Null
