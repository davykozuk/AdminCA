
Add-Type -AssemblyName PresentationCore, PresentationFramework

$Xaml = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation" Title="Soucis_Son" Height="283.181" Width="549.771" Background="#FF2B1D1D" BorderBrush="#FF191919">
    <Window.Effect>
        <DropShadowEffect/>
    </Window.Effect>
    <Grid Background="#FFF1F2F9">

        <Label Content="NOM DU POSTE" HorizontalAlignment="Left" Margin="32.041,30.092,0,0" VerticalAlignment="Top" Height="26.99" Width="146.773" HorizontalContentAlignment="Center" FontWeight="Bold" FontSize="14" Name="LPoste"/>
    
        <TextBox HorizontalAlignment="Left" Height="32.268" Margin="32,62.082,0,0" TextWrapping="Wrap" VerticalAlignment="Top" Width="146.773" TextAlignment="Center" VerticalContentAlignment="Center" Name="TPoste" />
   
        <Button Content="APPLIQUER" HorizontalAlignment="Left" Margin="205.439,62.082,0,0" VerticalAlignment="Top" Width="72.941" Height="32.268" RenderTransformOrigin="0.549,0.602" FontWeight="Bold" Name="BOk" >
               <Button.Effect>
                <DropShadowEffect/>
            </Button.Effect>
        </Button>
        <TextBlock HorizontalAlignment="Left" Margin="32,143,0,0" Name="TEtat" TextWrapping="Wrap" VerticalAlignment="Top" Height="46" Width="225" UseLayoutRounding="True" TextAlignment="Center" Background="#FFD39090"><Run/><LineBreak/><Run Text="Aucun poste selectionné"/><Run Text="..."/></TextBlock>
   
        <Button Content="Installer les pilotes" HorizontalAlignment="Left" Margin="362.52,131.947,0,0" VerticalAlignment="Top" Width="138.844" Height="57.053" Background="#FF4BD63D" BorderBrush="Black" FontWeight="Bold" FontSize="14" VerticalContentAlignment="Center" HorizontalContentAlignment="Center" UseLayoutRounding="False"> 
            <Button.Effect>
                <DropShadowEffect/>
            </Button.Effect>
        </Button>
   
        <Button Content="Vérifier la version &#xD;&#xA;des pilotes installés" HorizontalAlignment="Left" Margin="362.52,49.572,0,0" VerticalAlignment="Top" Width="138.844" Height="57.053" HorizontalContentAlignment="Center" VerticalContentAlignment="Center" Background="#FFC8CF39" BorderBrush="Black" FontSize="14" FontWeight="Bold" Name="BCheck">
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
#$BInstall.IsEnabled = $true
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


function FCheck(){
$Drivers = Get-WmiObject Win32_PnPSignedDriver -ComputerName $Poste | Select-Object -Property DriverVersion, Manufacturer
$DriverVers = $Drivers | Where-Object { $_.Manufacturer -like "*Realtek*" } | Out-String
Write-Host  "Pilote actuellement installé "$DriverVers "version cible" $cible }



#endregion

#-------------------------------------------------------------#
#----Script Execution-----------------------------------------#
#-------------------------------------------------------------#

$Window = [Windows.Markup.XamlReader]::Parse($Xaml)

[xml]$xml = $Xaml

$xml.SelectNodes("//*[@Name]") | ForEach-Object { Set-Variable -Name $_.Name -Value $Window.FindName($_.Name) }

#$BInstall.IsEnabled = $false

$BOk.Add_Click({FOk $this $_})
$BInstall.Add_Click({FInstall $this $_})
$BCheck.Add_Click({FCheck $this $_})


[void]$Window.ShowDialog()
