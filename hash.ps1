Add-Type -AssemblyName PresentationFramework

# -------- LANGUE --------
$Lang = "FR" # Mets "EN" pour anglais

# -------- XAML ---------
[xml]$xaml = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        Title="AutoPilot QR - OOBE"
        WindowStartupLocation="CenterScreen"
        Width="860" Height="520"
        Background="#FFF7F0" FontFamily="Segoe UI" ResizeMode="NoResize" Topmost="True">
    <Grid>
        <Grid.ColumnDefinitions>
            <ColumnDefinition Width="270"/>
            <ColumnDefinition Width="*"/>
        </Grid.ColumnDefinitions>
        <!-- Colonne de gauche -->
        <StackPanel Grid.Column="0" Background="#FFE4D6" Margin="0,0,0,0" VerticalAlignment="Stretch">
            <TextBlock Name="TxtTitre" Margin="24,38,12,12" Foreground="#5D90E3" FontWeight="Bold" FontSize="22"/>
            <TextBlock Name="TxtExplicatif" TextWrapping="Wrap" Margin="24,0,18,0" FontSize="16" Foreground="#282828"/>
            <TextBlock Name="TxtWarning" Margin="24,28,10,0" FontSize="14" Foreground="#B87C46"/>
        </StackPanel>
        <!-- Colonne de droite (formulaire) -->
        <StackPanel Grid.Column="1" Margin="0,20,0,0" HorizontalAlignment="Center">
            <Image Name="LogoImage" Width="90" Height="90" Margin="0,10,0,2"/>
            <TextBlock Text="AutoPilot" HorizontalAlignment="Center" FontSize="18" Foreground="#282828" Margin="0,0,0,16"/>
            <Border Background="#fff" CornerRadius="28" Margin="18,0,18,18" Padding="24">
                <StackPanel>
                    <TextBlock Name="LblSerial" Foreground="#272728" FontWeight="Bold" FontSize="22" Margin="0,0,0,8"/>
                    <TextBlock Name="LblBrand"  Foreground="#272728" FontWeight="Bold" FontSize="22" Margin="0,0,0,8"/>
                    <TextBlock Name="LblModel"  Foreground="#272728" FontWeight="Bold" FontSize="22"/>
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
            <TextBlock Name="LblUpn" Foreground="#5D90E3" FontWeight="Bold" FontSize="16" Margin="18,0,0,2"/>
            <TextBox Name="UpnBox" Margin="18,0,18,10" Padding="9" FontSize="18" Background="#fff" Foreground="#272728" Height="38"/>
            <TextBlock Name="LblToken" Foreground="#5D90E3" FontWeight="Bold" FontSize="16" Margin="18,0,0,2"/>
            <PasswordBox Name="TokenBox" Margin="18,0,18,18" Padding="9" FontSize="18" Background="#fff" Foreground="#272728" Height="38"/>
            <StackPanel Orientation="Horizontal" HorizontalAlignment="Center" Margin="0,12,0,22">
                <Button Name="BtnSend" Content="Envoyer" Width="120" Height="46" Margin="7" Padding="
