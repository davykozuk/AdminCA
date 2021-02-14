

#-------------------------------------------------------------#
#----Initial Declarations-------------------------------------#
#-------------------------------------------------------------#

Add-Type -AssemblyName PresentationCore, PresentationFramework

$Xaml = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation" xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml" Title="MainWindow" Height="350" Width="525">

    <Grid RenderTransformOrigin="0.5,0.5">
        <Grid HorizontalAlignment="Left" Height="152" Margin="10,10,0,0" VerticalAlignment="Top" Width="247">
            <TextBox HorizontalAlignment="Left" Height="23" Margin="10,10,0,0" TextWrapping="Wrap" VerticalAlignment="Top" Width="120" Name="Poste"/>
            <TextBlock HorizontalAlignment="Left" Margin="10,50,0,0" TextWrapping="Wrap" VerticalAlignment="Top" Height="92" Width="227"/>
            <Button Content="Infos" HorizontalAlignment="Left" Margin="162,13,0,0" VerticalAlignment="Top" Width="75" Name="Infosb"/>
        </Grid>
        <Grid HorizontalAlignment="Left" Height="87" Margin="262,10,0,0" VerticalAlignment="Top" Width="245">
            <ComboBox HorizontalAlignment="Left" Margin="10,10,0,0" VerticalAlignment="Top" Width="120" Name="ComboDell" ItemsSource="{Binding DELL}"/>
            <Button Content="Copier" HorizontalAlignment="Left" Margin="160,7,0,0" VerticalAlignment="Top" Width="75" Name="Copieb"/>
            <Button Content="Reparer" HorizontalAlignment="Left" Margin="160,32,0,0" VerticalAlignment="Top" Width="75" Name="Reparerb"/>
            <Button Content="Verifier" HorizontalAlignment="Left" Margin="160,57,0,0" VerticalAlignment="Top" Width="75" Name="Verifierb"/>
        </Grid>
        <TextBlock HorizontalAlignment="Left" Margin="298,68,0,0" TextWrapping="Wrap" Text="TextBlock" VerticalAlignment="Top" Height="22" Width="69" RenderTransformOrigin="0.488,0.293"/>
        <Grid HorizontalAlignment="Left" Height="142" Margin="10,167,0,0" VerticalAlignment="Top" Width="247">
            <ComboBox HorizontalAlignment="Left" Margin="10,35,0,0" VerticalAlignment="Top" Width="120" Name="ComboImp" ItemsSource="{Binding IMP}"/>
            <ComboBox HorizontalAlignment="Left" Margin="10,75,0,0" VerticalAlignment="Top" Width="84" Name="CombPort" ItemsSource="{Binding PORT}" Height="27"/>
            <TextBlock HorizontalAlignment="Left" Margin="107,78,0,0" TextWrapping="Wrap" Text="TextBlock" VerticalAlignment="Top" Height="54" Width="130"/>
            <Button Content="Desinstaller" HorizontalAlignment="Left" Margin="152,29,0,0" VerticalAlignment="Top" Width="75" Name="Desinstallerb"/>
            <TextBlock HorizontalAlignment="Left" Margin="10,10,0,0" TextWrapping="Wrap" Text="TextBlock" VerticalAlignment="Top" Width="120" Height="20"/>
            <Button Content="Print Mng" HorizontalAlignment="Left" Margin="12,112,0,0" VerticalAlignment="Top" Width="65" Height="21" Name="Printmgmtb"/>
        </Grid>
        <Button Content="Installer" HorizontalAlignment="Left" Margin="162,173,0,0" VerticalAlignment="Top" Width="75" Name="Installerb"/>
        <Grid HorizontalAlignment="Left" Height="207" VerticalAlignment="Top" Width="250" Margin="257,102,0,0"/>
    </Grid>

</Window>
"@

#-------------------------------------------------------------#
#----Control Event Handlers-----------------------------------#
#-------------------------------------------------------------#


#Write your code here
function Infos()
{
	$UserCo = QUERY SESSION
	$Model = Get-WmiObject Win32_ComputerSystem | select Model
	$Boot = Get-WmiObject Win32_OperatingSystem | select LastBootUpTime
    $Wmi = Get-WmiObject Win32_ComputerSystem
    $Model=($Wmi).Modele
    $InfosModel=$Model|Out-String
    $Retour1.Texte = $InfosModel


}

#endregion

#-------------------------------------------------------------#
#----Script Execution-----------------------------------------#
#-------------------------------------------------------------#

$Window = [Windows.Markup.XamlReader]::Parse($Xaml)

[xml]$xml = $Xaml

$xml.SelectNodes("//*[@Name]") | ForEach-Object { Set-Variable -Name $_.Name -Value $Window.FindName($_.Name) }


$Infosb.Add_Click({Infos $this $_})
$Copieb.Add_Click({Copier $this $_})
$Reparerb.Add_Click({Reparer $this $_})
$Verifierb.Add_Click({Verifier $this $_})
$Desinstallerb.Add_Click({Desinstaller $this $_})
$Printmgmtb.Add_Click({Printmgmt $this $_})
$Installerb.Add_Click({Installer $this $_})

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

{
    "DELL" : ["5590","5500","5510"],
    "IMP" : ["Brother","M300","C300"],
    "PORT" : ["USB001:","LPT1:","LPT2:"]
}

"@

$DataContext = New-Object System.Collections.ObjectModel.ObservableCollection[Object]
FillDataContext @("DELL","IMP","PORT") 

$Window.DataContext = $DataContext
Set-Binding -Target $ComboDell -Property $([System.Windows.Controls.ComboBox]::ItemsSourceProperty) -Index 0 -Name "DELL"
Set-Binding -Target $ComboImp -Property $([System.Windows.Controls.ComboBox]::ItemsSourceProperty) -Index 1 -Name "IMP"
Set-Binding -Target $CombPort -Property $([System.Windows.Controls.ComboBox]::ItemsSourceProperty) -Index 2 -Name "PORT"
$Window.ShowDialog()


