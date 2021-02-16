#-------------------------------------------------------------#
#----Initial Declarations-------------------------------------#
#-------------------------------------------------------------#

Add-Type -AssemblyName PresentationCore, PresentationFramework

$Xaml = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation" xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml" Title="INSTALLSON +" Height="400" Width="525" HorizontalAlignment="Left" VerticalAlignment="Top" Margin="0,0,0,0">
    <Grid>
        <Grid HorizontalAlignment="Right" Width="517" Margin="0,-1,0,1" Background="#FF9E8F8F">
            <Grid.ColumnDefinitions>
                <ColumnDefinition Width="259*"/>
                <ColumnDefinition Width="258*"/>
            </Grid.ColumnDefinitions>
            <TextBlock Grid.Column="1" TextWrapping="Wrap" Text="" Name="Retour" Background="#FFC7C4C4"/>
            <TextBox HorizontalAlignment="Left" Height="23" Margin="10,41,0,0" TextWrapping="Wrap" Text="" VerticalAlignment="Top" Width="120" Name="Poste"/>
            <Label Content="Nom du poste" HorizontalAlignment="Left" Margin="10,10,0,0" VerticalAlignment="Top"/>
            <Button Content="Valider le poste" HorizontalAlignment="Left" Margin="135,41,0,0" VerticalAlignment="Top" Width="114" Height="23" Name="kl3uzh3bk3a0w"/>
            <Button Content="Install son" HorizontalAlignment="Left" Margin="10,85,0,0" VerticalAlignment="Top" Width="120" Name="kl3uzh3b2z6ee"/>
            <Button Content="Display Link" HorizontalAlignment="Left" Margin="135,85,0,0" VerticalAlignment="Top" Width="120" Name="kl3uzh3bcgzdr"/>
            <Button Content="Infos sur les versions" HorizontalAlignment="Left" Margin="25,110,0,0" VerticalAlignment="Top" Width="180" Name="Version_b"/>
            <Label Content="Gestion des impressions" HorizontalAlignment="Left" Margin="55,142,0,0" VerticalAlignment="Top" Width="141"/>
            <Label Content="Modele imprimante" HorizontalAlignment="Left" Margin="10,183,0,0" VerticalAlignment="Top" Width="120"/>
            <ComboBox HorizontalAlignment="Left" Margin="10,214,0,0" VerticalAlignment="Top" Width="120" Name="ComboImp" ItemsSource="{Binding IMP}"/>
            <Label Content="Port ou IP" HorizontalAlignment="Left" Margin="10,241,0,0" VerticalAlignment="Top" Width="64"/>
            <ComboBox HorizontalAlignment="Left" Margin="10,272,0,0" VerticalAlignment="Top" Width="120" RenderTransformOrigin="0.457,0.72" Name="ComboPort" ItemsSource="{Binding PORT}"/>
            <Button Content="Infos" HorizontalAlignment="Left" Margin="164,214,0,0" VerticalAlignment="Top" Width="75" Name="kl3uzh3bxspqs"/>
            <Button Content="Installer" HorizontalAlignment="Left" Margin="164,239,0,0" VerticalAlignment="Top" Width="75" Name="kl3uzh3bb5ivb"/>
            <Button Content="Desinstaller" HorizontalAlignment="Left" Margin="164,264,0,0" VerticalAlignment="Top" Width="75" Name="kl3uzh3cjawp9"/>
            <Button Content="PrintMngr" HorizontalAlignment="Left" Margin="164,289,0,0" VerticalAlignment="Top" Width="75" Name="kl3uzh3cf64w2"/>
            <TextBox HorizontalAlignment="Left" Height="23" Margin="10,295,0,0" TextWrapping="Wrap" Text="" VerticalAlignment="Top" Width="120" Name="IpAddr"/>
        </Grid>
</Grid>
</Window>
"@

#-------------------------------------------------------------#
#----Control Event Handlers-----------------------------------#
#-------------------------------------------------------------#


function Valide()
{
    $Error.Clear()
    if ($Poste.Text -eq "")
    {$Retour.Text = "Veuillez saisir un nom de poste"}
    else {
	$Boot = Get-WmiObject -ComputerName $Poste.Text Win32_OperatingSystem|Select-Object LastBootUpTime
    $LastBoot = ($Boot).LastBootUpTime
    $LastBootS =[System.Management.ManagementDateTimeConverter]::ToDateTime($lastboot)
    $Wmi = Get-WmiObject -ComputerName $Poste.Text Win32_ComputerSystem
    $Model=($Wmi).Model
    $InfosModel= $Model|Out-String
    $Drivers = Get-WmiObject Win32_PnPSignedDriver -ComputerName $Poste.Text| Select-Object -Property DriverVersion, Manufacturer
    $DriverVers = $Drivers | Where-Object { $_.Manufacturer -like "*Realtek*" } | Out-String
    $Retour.Text = "Version du pilote audio $DriverVers" , " Model: $InfosModel",
    "Heure du demarrage: $LastBootS"
    $kl3uzh3b2z6ee.IsEnabled = $true
    $kl3uzh3bcgzdr.IsEnabled = $true
    $kl3uzh3bxspqs.IsEnabled = $true
    $kl3uzh3bb5ivb.IsEnabled = $true
    $kl3uzh3cjawp9.IsEnabled = $true
}
}

function InstallSon()
{
    $ComputerName = $Poste.Text
    $Wmi = Get-WmiObject -ComputerName $Poste.Text Win32_ComputerSystem
    $Model=($Wmi).Model

#if (Test-Path "\\$Poste\c$\temp\DUP\DELLMUP.exe")
if (Test-Path "\\$Poste.text\c$\temp\DUP_$Model\DELLMUP.exe")
{
    $Retour.Text = "Un pilote est deja present dans le dossier c:\temp de $Poste.text veuillez le supprimer et recommencer"
    Invoke-Item "\\$Poste.text\c$\temp\"
}
else {
 #   if (Test-Path "\\cw01pnmtst00\IP\Domaines clients\CR NMP\Drivers\DUP")
 if (Test-Path "\\cw01pnmtst00\IP\Domaines clients\CR NMP\Drivers\DUP_$Model")
    {
    $SourceF = "\\cw01pnmtst00\IP\Domaines clients\CR NMP\Drivers\DUP_$Model"
    ROBOCOPY "$SourceF" "\\$ComputerName\c$\temp\Drivers" /E | %{$data = $_.Split([char]9); if("$($data[4])" -ne "") { $file = "$($data[4])"} ;Write-Progress "Percentage $($data[0])" -Activity "Robocopy" -CurrentOperation "$($file)"  -ErrorAction SilentlyContinue; }
    $Retour.Text = "Copie des fichiers terminee installation va commencer"
    $Credential = Get-Credential
    $UserName = $Credential.UserName
    $Password = $Credential.GetNetworkCredential().Password
    .\PsExec.exe \\$ComputerName -u $UserName -p $Password -h cmd /c "c:\temp\Drivers\DELLMUP.exe" /s /v"FORCERESTART=true" /v"LOGFILE=c:\temp\logmup.log" /v"FORCE=true"
    Write-Host "L'INSTALLATION DES PILOTES EST TERMINEE"
    $Retour.Text = "L'INSTALLATION DES PILOTES EST TERMINEE"
}

    else
    { $Retour.Text = "Impossible de recuperer les pilotes sur le serveur IP"
    }

}}


function Versioninfo()
{
    Invoke-Item "Chemin du fichier d'infos des versions"
}

function Installer
{   
    $Error.Clear()
    $ComputerName = $Poste.Text
    $NomImprimante = $ComboImp.Text
    $Driver = $ComboImp.Text
    $PortImp = $ComboPort.Text
    if ($PortImp -eq "" -and $IpAddr.Text -eq "")
    {$Retour.Text = "Veuillez selectionner un port ou une adresse IP:"}
    else{
    Add-Printer -ComputerName $ComputerName -Name $NomImprimante -DriverName $Driver -Port $PortImp
    $Retour.Text = $Error
}
}

function Desinstaller
{
    $Error.Clear()
    $ComputerName = $Poste.Text
    $NomImprimante = $ComboImp.Text
    Remove-Printer -ComputerName $ComputerName -Name $NomImprimante*
    $Retour.Text = $Error
}

function infos
{
$Error.Clear()
$ImpInstall = get-printer -ComputerName $Poste.Text | Select-Object -Property Name | Format-Table | Out-String
$Retour.Text = $Error
$Retour.Text = $ImpInstall
}

function Printmgmt ()
{
 Start-Process printmanagement.msc
}

function DisplayLink()
{   
    $ComputerName = $Poste.Text
    $Error.Clear()
    $Credential = Get-Credential
    $UserName = $Credential.UserName
    $Password = $Credential.GetNetworkCredential().Password
    robocopy "\\cw01pnmtst00\IP\Domaines clients\CR NMP\Drivers\DisplayLink_Win10RS.msi" "\\$ComputerName\c$\temp\DisplayLink_Win10RS.msi"
    .\PsExec.exe \\$ComputerName -u $UserName -p $Password -h cmd /c msiexec /i "\\$ComputerName\c:\temp\DisplayLink_Win10RS.msi" /norestart /quiet
    $Retour.Text = "Display Link installé."
    write-host "Display Link installé."
}
#endregion

#-------------------------------------------------------------#
#----Script Execution-----------------------------------------#
#-------------------------------------------------------------#

$Window = [Windows.Markup.XamlReader]::Parse($Xaml)

[xml]$xml = $Xaml

$Source ="\\CW01pnmtst00\IP\Domaines clients\CR NMP\Drivers\DUP"
$xml.SelectNodes("//*[@Name]") | ForEach-Object { Set-Variable -Name $_.Name -Value $Window.FindName($_.Name) }

$kl3uzh3b2z6ee.IsEnabled = $false
$kl3uzh3bcgzdr.IsEnabled = $false
$kl3uzh3bxspqs.IsEnabled = $false
$kl3uzh3bb5ivb.IsEnabled = $false
$kl3uzh3cjawp9.IsEnabled = $false

$kl3uzh3bcgzdr.Add_Click({DisplayLink $this $_})
$kl3uzh3bk3a0w.Add_Click({Valide $this $_})
$kl3uzh3b2z6ee.Add_Click({InstallSon $this $_})
$Version_b.Add_Click({Versioninfo $this $_})
$kl3uzh3bxspqs.Add_Click({infos $this $_})
$kl3uzh3bb5ivb.Add_Click({Installer $this $_})
$kl3uzh3cjawp9.Add_Click({Desinstaller $this $_})
$kl3uzh3cf64w2.Add_Click({Printmgmt $this $_})
$ComboImp.SelectedItem = 1

$State = [PSCustomObject]@{}


Function Set-Binding {
    Param($Target,$Property,$Index,$Name)
 
    $Binding = New-Object System.Windows.Data.Binding
    $Binding.Path = "["+$Index+"]"
    $Binding.Mode = [System.Windows.Data.BindingMode]::TwoWay
    


    [void]$Target.SetBinding($Property,$Binding)
}

function FillDataContext($props){

    For ($i=0; $i -lt $props.Length; $i++) {
   
   $prop = $props[$i]
   $DataContext.Add($DataObject."$prop")
   
    $getter = [scriptblock]::Create("return `$DataContext['$i']")
    $setter = [scriptblock]::Create("param(`$val) return `$DataContext['$i']=`$val")
    $State | Add-Member -Name $prop -MemberType ScriptProperty -Value  $getter -SecondValue $setter
               
       }
   }



$DataObject =  ConvertFrom-Json @"
{    "DELL" : ["5590","5500","5510"],
    "IMP" : ["Brother HL-L6300DW series","Epson AL-M300 series","Epson AL-C300 series"],
    "PORT" : ["USB001:","LPT1:","LPT2:"]}
"@

$DataContext = New-Object System.Collections.ObjectModel.ObservableCollection[Object]
FillDataContext @("DELL","IMP","PORT") 

$Window.DataContext = $DataContext
Set-Binding -Target $ComboImp -Property $([System.Windows.Controls.ComboBox]::ItemsSourceProperty) -Index 1 -Name "IMP"
Set-Binding -Target $ComboPort -Property $([System.Windows.Controls.ComboBox]::ItemsSourceProperty) -Index 2 -Name "PORT"
$Window.ShowDialog()
