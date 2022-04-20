
Add-Type -AssemblyName PresentationCore, PresentationFramework

$Xaml = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation" Title="Soucis_Son" Height="130" Width="410" Background="#FF2B1D1D" BorderBrush="#FF191919">
    <Window.Effect>
        <DropShadowEffect/>
    </Window.Effect>
    <Grid>
        <TextBox HorizontalAlignment="Left" Margin="10,10,0,0" Text="Nom du poste" TextWrapping="Wrap" VerticalAlignment="Top" Width="141" BorderBrush="Black" Background="#FF593F3F" Height="25" TextAlignment="Center" Name="TPoste">
            <TextBox.Effect>
                <DropShadowEffect/>
            </TextBox.Effect>
        </TextBox>
        <Button Content="OK" HorizontalAlignment="Center" Margin="0,10,0,0" VerticalAlignment="Top" Height="25" Background="#FF9E7B7B" Width="28" FontFamily="MS Gothic" Name="BOk">
            <Button.Effect>
                <DropShadowEffect/>
            </Button.Effect>
        </Button>
        <TextBlock HorizontalAlignment="Left" Margin="276,15,0,0" Text="Connexion.." TextWrapping="Wrap" VerticalAlignment="Top" Width="87" Background="#FFB46464" TextAlignment="Center" Name="TEtat">
            <TextBlock.Effect>
                <DropShadowEffect/>
            </TextBlock.Effect>
        </TextBlock>
        <Button Content="Installer les drivers" HorizontalAlignment="Left" Margin="241,53,0,0" VerticalAlignment="Top" Width="145" Background="#FF533838" FontFamily="MS Gothic" Name="BInstall">
            <Button.Effect>
                <DropShadowEffect/>
            </Button.Effect>
        </Button>

    </Grid>
</Window>
"@

#-------------------------------------------------------------#
#----Control Event Handlers-----------------------------------#
#-------------------------------------------------------------#

function FOk(){
$Poste=$TPoste.Text
$Wmi = Get-WmiObject -ComputerName $Poste Win32_ComputerSystem
$Model=($Wmi).Model
$Date=get-date -UFormat %d%m%Y
mkdir -p logs
$LogFile=".\logs\$Date$Poste.txt"
if (Test-Connection $Poste){
$BInstall.IsEnabled = $true
$TEtat.Background="#FF37C73D"
$TEtat.Text="$Model"
$Drivers = Get-WmiObject Win32_PnPSignedDriver -ComputerName $Poste | Select-Object -Property DriverVersion, Manufacturer
$DriverVers = $Drivers | Where-Object { $_.Manufacturer -like "*Realtek*" } | Out-String
"Versions avant installation $DriverVers" >> $LogFile
}
else {$TEtat.Background="#FFCF2626"
$TEtat.Text="Hors Ligne"}
}



function FInstall(){
$Poste=$TPoste.Text
$Wmi = Get-WmiObject -ComputerName $Poste Win32_ComputerSystem
$Model=($Wmi).Model
$Date=get-date -UFormat %d%m%Y
$LogFile=".\logs\$Date$Poste.txt"
 if (Test-Path "\\cw01pnmtst00\IP\Domaines clients\CR NMP\Drivers\DUP_$Model")
    {
    $SourceF = "\\cw01pnmtst00\IP\Domaines clients\CR NMP\Drivers\DUP_$Model"
    ROBOCOPY "$SourceF" "\\$Poste\c$\temp\Drivers" /E | %{$data = $_.Split([char]9); if("$($data[4])" -ne "") { $file = "$($data[4])"} ;Write-Progress "Pourcentage copié: $($data[0])" -Activity "Copie des drivers" -CurrentOperation "$($file)"  -ErrorAction SilentlyContinue;}
    write-progress -Activity "Terminé " -status "100%" -Completed
    Write-Host "Copie des fichiers terminee installation va commencer"
    $Credential = Get-Credential
    $UserName = $Credential.UserName
    $Password = $Credential.GetNetworkCredential().Password
    .\PsExec.exe \\$Poste -u $UserName -p $Password -h cmd /c "c:\temp\Drivers\DELLMUP.exe" /s /v"FORCERESTART=true" /v"LOGFILE=c:\temp\logmup.log" /v"FORCE=true"
    $Drivers = Get-WmiObject Win32_PnPSignedDriver -ComputerName $Poste | Select-Object -Property DriverVersion, Manufacturer
    $DriverVers = $Drivers | Where-Object { $_.Manufacturer -like "*Realtek*" } | Out-String
    "Versions après installation $DriverVers" >> $LogFile
    Write-Host "L'INSTALLATION DES PILOTES EST TERMINEE"
}

    else
    { Write-Host "Impossible de recuperer les pilotes sur le serveur IP"
    }

}


#endregion

#-------------------------------------------------------------#
#----Script Execution-----------------------------------------#
#-------------------------------------------------------------#

$Window = [Windows.Markup.XamlReader]::Parse($Xaml)

[xml]$xml = $Xaml

$xml.SelectNodes("//*[@Name]") | ForEach-Object { Set-Variable -Name $_.Name -Value $Window.FindName($_.Name) }

$BInstall.IsEnabled = $false

$BOk.Add_Click({FOk $this $_})
$BInstall.Add_Click({FInstall $this $_})


[void]$Window.ShowDialog()