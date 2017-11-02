Add-Type -AssemblyName PresentationFramework

$ZIPUpdate = "\\10.252.70.3\Users\Public\Downloads\AVImark Beta 2016.4.5.zip"
$DestinationPath = "C:\Temp\Avimark"
$AVIMarkZip = "C:\Temp\Avimark\AVImark Beta 2016.4.5.zip"

$TotalSteps = 6

$Step = 1
$Activity = "Running AVIMark Auto Update Tasks"
$Task = "Downloading the AVImark update archive:"

Write-Progress -Id 1 -Activity $Activity -Status $Task -PercentComplete ($Step / $TotalSteps * 100)

Start-Sleep -Seconds 2

if ((Test-Path -Path $Destination) -eq $false) {
    New-Item -Path $Destination -ItemType Container -Force
    $Download = (New-Object -COM Shell.Application).NameSpace($ZIPUpdate)
    $Destination = (New-Object -COM Shell.Application).NameSpace($DestinationPath)
    $Destination.CopyHere($Download.Items(), 16)
}

if ((Test-Path -Path $AVIMarkZip) -eq $false) {
    $Download = (New-Object -COM Shell.Application).NameSpace($ZIPUpdate)
    $Destination = (New-Object -COM Shell.Application).NameSpace($DestinationPath)
    $Destination.CopyHere($Download.Items(), 16)
}

$Step = $Step + 1
$Task = "Killing all AVImark related tasks:"

Write-Progress -Id 1 -Activity $Activity -Status $Task -PercentComplete ($Step / $TotalSteps * 100)

Start-Sleep -Seconds 2

$AVIMarkProcess = Get-Process | ? {$_.ProcessName -like "AVIM*"}

if (($AVIMarkProcess | Select-Object -ExpandProperty Name) -like "AVIM*") {
     
    foreach ($Process in ($AVIMarkProcess | Select-Object -ExpandProperty ProcessName)) {
        Stop-Process -Name $Process -Force
    }
}

$MPSProcess = Get-Process | ? {$_.ProcessName -like "MPS*"}

if (($MPSProcess | Select-Object -ExpandProperty ProcessName) -like "MPS*") {

    foreach ($Process in ($MPSProcess | Select-Object -ExpandProperty Name)) {
        Stop-Process -Name $Process -Force
    }
}

$Step = $Step + 1
$Task = "Stopping all AVImark services and dependencies:"

Write-Progress -Id 1 -Activity $Activity -Status $Task -PercentComplete ($Step / $TotalSteps * 100)

Start-Sleep -Seconds 2

$AVIMarkServer = Get-Service | ? {$_.ServiceName -like "AVIM*"}

if (($AVIMarkServer | Select-Object -ExpandProperty Status) -eq "Running") {
    
    foreach ($Service in ($AVIMarkServer | Select-Object -ExpandProperty Name)) {
        Stop-Service -Name $Service -Force -ErrorAction SilentlyContinue
    }
}

$IDEXXService = Get-Service | ? {$_.DisplayName -like "IDEXX*"}

if (($IDEXXService | Select-Object -ExpandProperty Status) -eq "Running" ) {
    
    foreach ($Service in ($IDEXXService | Select-Object -ExpandProperty Name)) {
        Stop-Service -Name $Service -Force -ErrorAction SilentlyContinue
    }
}

$Vetstoria = Get-Service | ? {$_.ServiceName -like "Vets*"}

if (($Vetstoria | Select-Object -ExpandProperty Status) -eq "Running") {
    
    foreach ($Service in ($Vetstoria | Select-Object -ExpandProperty Name)) {
        Stop-Service -Name $Service -Force -ErrorAction SilentlyContinue
    }
}

$Step = $Step + 1
$Task = "Backing up AVIMark directory:"

Write-Progress -Id 1 -Activity $Activity -Status $Task -PercentComplete ($Step / $TotalSteps * 100)

$BackCheck = "D:\Backup\AVImark"

if ((Test-Path -Path $BackCheck) -eq $false) {
    New-Item -Path $BackCheck -ItemType Container -Force | Out-Null
}

$Source = "C:\users\aowens\desktop\AVImark"
$Backup = (New-Object -COM Shell.Application).NameSpace($Source)
$DestinationPath = "D:\Backup\AVImark"
$Destination = (New-Object -COM Shell.Application).NameSpace($DestinationPath)
$Destination.CopyHere($Backup.Items(), 16)

$Step = $Step + 1
$Task = "Extracting ZIP archive to D:\AVImark:"

Write-Progress -Id 1 -Activity $Activity -Status $Task -PercentComplete ($Step / $TotalSteps * 100)

if ((Test-Path -Path $BackCheck) -eq $true) {

    $ZipFile = (New-Object -COM Shell.Application).NameSpace($AVIMarkZip)
    $DestinationPath = "C:\users\aowens\desktop\AVImark"
    $Destination = (New-Object -COM Shell.Application).NameSpace($DestinationPath)

    if ((Test-Path -Path $AVIMarkZip) -eq $true) {
        $Destination.CopyHere($ZipFile.Items(), 16)
    } else {
        Write-Error "AVIMark update file did not transfer properly!"
    }

} else {
        Write-Error "AVIMark backup did not complete successfully. Please, manually backup to 'D:\Backup\Avimark' and re-run script."
}

$Task = "completed"

Write-Progress -Id 1 -Activity $Activity -Status $Task -Completed

$Continue = [System.Windows.MessageBox]::Show("Please, continue manual update of Avimark, then press 'OK' to start AVImark services","Continue manually","OKCancel","Information")

    switch ($Continue) {

    'OK' {

        foreach ($Service in ($AVIMarkServer | Select-Object -ExpandProperty Name)) {
            Start-Service -Name $Service -Verbose
        }

        foreach ($Service in ($Vetstoria | Select-Object -ExpandProperty Name)) {
            Start-Service -Name $Service -Verbose
        }

        foreach ($Service in ($IDEXXService | Select-Object -ExpandProperty Name)) {
            Start-Service -Name $Service -Verbose
        }
    }

    'Cancel' {

        Exit

    }
    }