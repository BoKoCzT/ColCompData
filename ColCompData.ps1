# Collect Computer Data
#
# $Cesta = $MyInvocation.MyCommand.Path           # "c:\tmp"
# $Subor = $Cesta + "\CompData.txt"
$Cesta = $PSScriptRoot
$Subor = $Cesta + "\CompData.txt"
# $myFilePath = Split-Path $file -Parent
# $DiskPismeno = gwmi win32_diskdrive | ?{$_.interfacetype -eq "USB"} | %{gwmi -Query "ASSOCIATORS OF {Win32_DiskDrive.DeviceID=`"$($_.DeviceID.replace('\','\\'))`"} WHERE AssocClass = Win32_DiskDriveToDiskPartition"} |  %{gwmi -Query "ASSOCIATORS OF {Win32_DiskPartition.DeviceID=`"$($_.DeviceID)`"} WHERE AssocClass = Win32_LogicalDiskToPartition"} | %{$_.deviceid}
$Oddelovac = "--------------------------------------------------------------------------------------------------"


# Overenie, či existuje a je prístupná cesta pre uloženie súboru:
if (Test-Path $Cesta)
{
    Write-Host "Cesta OK..."
}
else # cesta nebola prístupná
{
    New-Item $Cesta -ItemType Directory   # Vytvoriť danú cestu
    Write-Host "Zložka bola vytvorená..."
}

Clear-Content $Subor   # premazať obsah súboru, než doň začneme zapisovať informácie

    Add-Type -AssemblyName System.Windows.Forms
    Add-Type -AssemblyName System.Drawing

    $form = New-Object System.Windows.Forms.Form
    $form.Text = 'Zadanie mena učiteľa'
    $form.Size = New-Object System.Drawing.Size(300,200)
    $form.StartPosition = 'CenterScreen'

    $okButton = New-Object System.Windows.Forms.Button
    $okButton.Location = New-Object System.Drawing.Point(75,120)
    $okButton.Size = New-Object System.Drawing.Size(75,23)
    $okButton.Text = 'OK'
    $okButton.DialogResult = [System.Windows.Forms.DialogResult]::OK
    $form.AcceptButton = $okButton
    $form.Controls.Add($okButton)

    $cancelButton = New-Object System.Windows.Forms.Button
    $cancelButton.Location = New-Object System.Drawing.Point(150,120)
    $cancelButton.Size = New-Object System.Drawing.Size(75,23)
    $cancelButton.Text = 'Storno'
    $cancelButton.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
    $form.CancelButton = $cancelButton
    $form.Controls.Add($cancelButton)

    $label = New-Object System.Windows.Forms.Label
    $label.Location = New-Object System.Drawing.Point(10,20)
    $label.Size = New-Object System.Drawing.Size(280,20)
    $label.Text = 'Prosím zadajte svoje meno do textového poľa:'
    $form.Controls.Add($label)

    $textBox = New-Object System.Windows.Forms.TextBox
    $textBox.Location = New-Object System.Drawing.Point(10,40)
    $textBox.Size = New-Object System.Drawing.Size(260,20)
    $form.Controls.Add($textBox)

    $form.Topmost = $true

    $form.Add_Shown({$textBox.Select()})
    $result = $form.ShowDialog()

    if ($result -eq [System.Windows.Forms.DialogResult]::OK)
    {
        $Meno = $textBox.Text
    }
    else
    {
        $Meno = "Chyba, nebolo zadané užívateľské meno..."
    }


Write-Host "Čakajte prosím, zapisujem do súboru..."
$Pocitac = Get-CimInstance -ClassName Win32_ComputerSystem

#Vytvoríme cestu súboru z cesty skriptu a mena užívateľa:
$Subor = $Cesta + "\" + $Meno + ".txt"
Write-Host "Vytvorená cesta: " + $Subor

# Meno používateľa:
$Oddelovac | Out-File -FilePath $Subor -Append
"Meno užívateľa: $Meno" | Out-File -FilePath $Subor -Append
" " | Out-File -FilePath $Subor -Append
" " | Out-File -FilePath $Subor -Append
Test-NetConnection | Select-Object SourceAddress | Out-File -FilePath $Subor -Append

$SerialNum = Get-WmiObject win32_bios | select SerialNumber
$Oddelovac | Out-File -FilePath $Subor -Append
"Typ počítača:" | Out-File -FilePath $Subor -Append
" " | Out-File -FilePath $Subor -Append
$OS = Get-WmiObject -class Win32_OperatingSystem
"OS:                      " + $OS.Caption | Out-File -FilePath $Subor -Append
“Užívateľ:                ” + $env:UserName | Out-File -FilePath $Subor -Append
"Názov počítača:          " + $env:ComputerName + "( " + $OS.CSName + " )" | Out-File -FilePath $Subor -Append
"Architektúra počítača:   " + $Pocitac.Description | Out-File -FilePath $Subor -Append
"Doména počítača:         " + $Pocitac.Domain | Out-File -FilePath $Subor -Append
"Výrobca počítača:        " + $Pocitac.Manufacturer | Out-File -FilePath $Subor -Append
"Model počítača:          " + $Pocitac.Model | Out-File -FilePath $Subor -Append
"Sériové číslo počítača:  " + $SerialNum.SerialNumber  | Out-File -FilePath $Subor -Append
"Pracovná skupina počítača:     " + $Pocitac.Workgroup | Out-File -FilePath $Subor -Append
" " | Out-File -FilePath $Subor -Append

$RAM = Get-WmiObject -Query "SELECT TotalVisibleMemorySize, FreePhysicalMemory FROM Win32_OperatingSystem"
$totalRAM = [math]::Round($RAM.TotalVisibleMemorySize/1MB, 2)
$freeRAM = [math]::Round($RAM.FreePhysicalMemory/1MB, 2)
$usedRAM = [math]::Round(($RAM.TotalVisibleMemorySize - $RAM.FreePhysicalMemory)/1MB, 2)
" "  | Out-File -FilePath $Subor -Append
" " | Out-File -FilePath $Subor -Append

$Oddelovac | Out-File -FilePath $Subor -Append
"RAM:" | Out-File -FilePath $Subor -Append
"Veľkosť RAM = $totalRAM" | Out-File -FilePath $Subor -Append
"Voľná RAM = $freeRAM" | Out-File -FilePath $Subor -Append
"Použitá RAM = $usedRAM" | Out-File -FilePath $Subor -Append
" " | Out-File -FilePath $Subor -Append
" " | Out-File -FilePath $Subor -Append

# Operaacny System
$Oddelovac | Out-File -FilePath $Subor -Append
"Operačný systém:    " + $OS.Caption | Out-File -FilePath $Subor -Append
"Číslo zostavy:      " + $OS.BuildNumber | Out-File -FilePath $Subor -Append
"Architektúra OS:    " + $OS.OSArchitecture | Out-File -FilePath $Subor -Append
" " | Out-File -FilePath $Subor -Append

# Sieť:
$Oddelovac | Out-File -FilePath $Subor -Append
"Sieťové rozhrania:" | Out-File -FilePath $Subor -Append
Get-NetAdapter | Select Name, MacAddress, LinkSpeed | Out-File -FilePath $Subor -Append
" " | Out-File -FilePath $Subor -Append
" " | Out-File -FilePath $Subor -Append

# Disky:
$Oddelovac | Out-File -FilePath $Subor -Append
"Disky v systéme:    " | Out-File -FilePath $Subor -Append
Get-CimInstance -ClassName Win32_LogicalDisk | Out-File -FilePath $Subor -Append
$Disk = Get-WmiObject -class Win32_LogicalDisk -Filter "DeviceID='C:'"
"ID disku: " + $Disk.DeviceID | Out-File -FilePath $Subor -Append
$Disk_TotalSpace = [math]::Round($Disk.Size/1GB, 2)
$Disk_FreeSpace = [math]::Round($Disk.FreeSpace/1GB, 2)
$Disk_UsedSpace = [math]::Round(($Disk.Size - $Disk.FreeSpace)/1GB, 2)
"Celkové miesto na systémovom disku: " + $Disk_TotalSpace | Out-File -FilePath $Subor -Append
"Voľné miesto na systémovom disku: " + $Disk_FreeSpace | Out-File -FilePath $Subor -Append
"Použité miesto na systémovom disku: " + $Disk_UsedSpace | Out-File -FilePath $Subor -Append
" " | Out-File -FilePath $Subor -Append
" " | Out-File -FilePath $Subor -Append
# CPU:
$Oddelovac | Out-File -FilePath $Subor -Append
"CPU:" | Out-File -FilePath $Subor -Append
Get-WmiObject -class Win32_Processor | Out-File -FilePath $Subor -Append
" " | Out-File -FilePath $Subor -Append

# GPU:
$Oddelovac | Out-File -FilePath $Subor -Append
"GPU:" | Out-File -FilePath $Subor -Append
Get-WmiObject win32_VideoController | select name, currentH*, currentv* | Out-File -FilePath $Subor -Append

"Grafické výstupy:" | Out-File -FilePath $Subor -Append

$adapterTypes = @{ #https://www.magnumdb.com/search?q=parent:D3DKMDT_VIDEO_OUTPUT_TECHNOLOGY
    '-2' = 'Unknown'
    '-1' = 'Unknown'
    '0' = 'VGA D-Sub'
    '1' = 'S-Video'
    '2' = 'Compozitné'
    '3' = 'Componentné'
    '4' = 'DVI'
    '5' = 'HDMI'
    '6' = 'LVDS'
    '8' = 'D-Jpn'
    '9' = 'SDI'
    '10' = 'DisplayPort (external)'
    '11' = 'DisplayPort (internal)'
    '12' = 'Unified Display Interface'
    '13' = 'Unified Display Interface (embedded)'
    '14' = 'SDTV dongle'
    '15' = 'Miracast'
    '16' = 'Internal'
    '2147483648' = 'Internal'
}

$arrMonitors = @()

$monitors = gwmi WmiMonitorID -Namespace root/wmi
$connections = gwmi WmiMonitorConnectionParams -Namespace root/wmi

foreach ($monitor in $monitors)
{
    $manufacturer = $monitor.ManufacturerName
    $name = $monitor.UserFriendlyName
    $connectionType = ($connections | ? {$_.InstanceName -eq $monitor.InstanceName}).VideoOutputTechnology

    if ($manufacturer -ne $null) {$manufacturer =[System.Text.Encoding]::ASCII.GetString($manufacturer -ne 0)}
	if ($name -ne $null) {$name =[System.Text.Encoding]::ASCII.GetString($name -ne 0)}
    $connectionType = $adapterTypes."$connectionType"
    if ($connectionType -eq $null){$connectionType = 'Unknown'}

    if(($manufacturer -ne $null) -or ($name -ne $null)){$arrMonitors += "$manufacturer $name ($connectionType)"}

}

$i = 0
$strMonitors = ''
if ($arrMonitors.Count -gt 0){
    foreach ($monitor in $arrMonitors){
        if ($i -eq 0){$strMonitors += $arrMonitors[$i]}
        else{$strMonitors += "`n"; $strMonitors += $arrMonitors[$i]}
        $i++
    }
}

if ($strMonitors -eq ''){$strMonitors = 'None Found'}
$strMonitors | Out-File -FilePath $Subor -Append

" " | Out-File -FilePath $Subor -Append
" " | Out-File -FilePath $Subor -Append

$Oddelovac | Out-File -FilePath $Subor -Append
"Tlačiarne:" | Out-File -FilePath $Subor -Append
Get-Printer | Format-Table | Out-File -FilePath $Subor -Append
notepad.exe $Subor