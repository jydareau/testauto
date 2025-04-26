Add-Type -AssemblyName PresentationFramework
Add-Type -AssemblyName PresentationCore
Add-Type -AssemblyName WindowsBase

# Logo "shield AutoPilot" intégré (miniature 90x90px, fond Peach Fuzz)
$logoBase64 = @"
iVBORw0KGgoAAAANSUhEUgAAAIwAAACMCAYAAAD5H/9jAAAACXBIWXMAAAsSAAALEgHS3X78AAACfElEQVR4nO3dvWpUQRiF4fNpIHRCcCAyZEmzAzvCklAEk6QNAZdoSYJJp2aUwX5KZT2yL8PF7OnMTv4Dczz7vnzAyBgwIABAwYMGDAgQYMGDAgQYMGDAgQYMGDAgQYMGDAgQYMGDAgQYMGDKgQQ5OkXtcdQOhcbpN33QOwB7b37C9zFfBzGrfsl5h7uWAd8CmwF7gu/A4MAL+By8BQ8Cw8Dc8B28XvgW4E5wF1wIzgK3gQk8AnR2Lsj3Ptj9wIakbwF3iU8xHBnysZv4WAWwK3j8WQDdB6wF7hb+E0GfAOuBr4Bb8Kx8QO8BQZ/PYmsFbiN5jDKe2jfAzbBRbAm8AybBk8CSwNmwH9ABzAX8AA8y2Zj/58/Y/gXWwufA28KngOXAHPBcH/K1rb/ANs6p1lPhvFgwYMGDAgQYMGDAgQYMGDAgQYMGDAgQYMGDAgQYMGDAgQYMGDAgQYMGDNihfgQgbOsTDQ2TAAAAAElFTkSuQmCC
"@

function Get-LogoImage {
    $bytes = [Convert]::FromBase64String($logoBase64)
    $stream = New-Object System.IO.MemoryStream(,$bytes)
    $image = New-Object System.Windows.Media.Imaging.PngBitmapDecoder($stream, [System.Windows.Media.Imaging.BitmapCreateOptions]::None, [System.Windows.Media.Imaging.BitmapCacheOption]::OnLoad)
    $frame = $image.Frames[0]
    return $frame
}

[xml]$xaml = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        Title="AutoPilot QR - OOBE" WindowStartupLocation="CenterScreen" 
        WindowState="Maximized" Background="#FFF7F0" FontFamily="Segoe UI" ResizeMode="NoResize" Topmost="True">
    <Grid>
        <StackPanel Margin="0,48,0,0" HorizontalAlignment="Center">
            <Border Background="#FFD1B2" CornerRadius="65" Width="130" Height="130" Margin="0,0,0,24" HorizontalAlignment="Center">
                <Image Name="LogoImage" Width="90" Height="90" VerticalAlignment="Center" HorizontalAlignment="Center" Margin="0"/>
            </Border>
            <Border Background="#fff" CornerRadius="28" Margin="32,0,32,20" Padding="24">
                <StackPanel>
                    <TextBlock Name="LblSerial" Text="Numéro de série : ..." Foreground="#272728" FontWeight="Bold" FontSize="24" Margin="0,0,0,9"/>
                    <TextBlock Name="LblBrand"  Text="Marque : ..." Foreground="#272728" FontWeight="Bold" FontSize="24" Margin="0,0,0,9"/>
                    <TextBlock Name="LblModel"  Text="Modèle : ..." Foreground="#272728" FontWeight="Bold" FontSize="24"/>
                </StackPanel>
            </Border>
            <TextBlock Text="Group Tag" Foreground="#5D90E3" FontWeight="Bold" FontSize="18" Margin="30,2,0,2"/>
            <ComboBox  Name="GroupTagCombo" Margin="32,0,32,12" FontSize="20" Background="#fff" Foreground="#272728" Height="44">
                <ComboBoxItem>Aucun</ComboBoxItem>
                <ComboBoxItem>VIP-Devices</ComboBoxItem>
                <ComboBoxItem>Direction</ComboBoxItem>
                <ComboBoxItem>RH</ComboBoxItem>
                <ComboBoxItem>Comptabilité</ComboBoxItem>
                <ComboBoxItem>IT</ComboBoxItem>
                <ComboBoxItem>Stagiaire</ComboBoxItem>
                <ComboBoxItem>Personnalisé</ComboBoxItem>
            </ComboBox>
            <TextBox Name="CustomGroupTag" Margin="32,0,32,12" Padding="9" FontSize="20" Background="#fff" Foreground="#272728" Visibility="Collapsed" Height="44"/>
            <TextBlock Text="UPN utilisateur" Foreground="#5D90E3" FontWeight="Bold" FontSize="18" Margin="30,0,0,2"/>
            <TextBox Name="UpnBox" Margin="32,0,32,12" Padding="9" FontSize="20" Background="#fff" Foreground="#272728" Height="44"/>
            <TextBlock Text="Token API/Bearer" Foreground="#5D90E3" FontWeight="Bold" FontSize="18" Margin="30,0,0,2"/>
            <PasswordBox Name="TokenBox" Margin="32,0,32,20" Padding="9" FontSize="20" Background="#fff" Foreground="#272728" Height="44"/>
            <StackPanel Orientation="Horizontal" HorizontalAlignment="Center" Margin="0,12,0,22">
                <Button Name="BtnSend" Content="Envoyer" Width="135" Height="56" Margin="7" Padding="9"
                    Background="#FFC6A0" Foreground="#272728" FontWeight="Bold" BorderThickness="0"  Cursor="Hand"/>
                <Button Name="BtnQR" Content="QR Code" Width="110" Height="56" Margin="7" Padding="9"
                    Background="#5D90E3" Foreground="#fff" FontWeight="Bold" BorderThickness="0" Cursor="Hand"/>
                <Button Name="BtnQuit" Content="Quitter" Width="90" Height="56" Margin="7" Padding="9"
                    Background="#FFD1B2" Foreground="#B81E1E" FontWeight="Bold" BorderThickness="0" Cursor="Hand"/>
            </StackPanel>
        </StackPanel>
    </Grid>
</Window>
"@

# --- Script PowerShell pour l'UI ---
$reader = (New-Object System.Xml.XmlNodeReader $xaml)
$window = [Windows.Markup.XamlReader]::Load($reader)

$logoImage    = $window.FindName("LogoImage")
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

# --- Affiche le logo intégré
try {
    $logoImage.Source = Get-LogoImage
} catch {}

# --- Infos device
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
                [System.Windows.MessageBox]::Show("Numéro de ticket : $($resp.ticketId)`nScannez le QR code pour valider/importer.", "Envoyé", "OK", "Info")
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
            [System.Windows.MessageBox]::Show("Erreur génération QR code : $_")
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

# --- Plein écran dès le départ
$window.WindowState = 'Maximized'
$window.Topmost = $true
$window.ShowDialog() | Out-Null
