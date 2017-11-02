Add-Type -AssemblyName PresentationFramework
$host.ui.rawui.WindowTitle = "AVImark Auto-Updater"

$ZIPUpdate = "\\10.252.70.3\Users\Public\Downloads\BETAOPT.zip"
$DestinationPath = "C:\Temp\Avimark"
$AVIMarkZip = "C:\Temp\Avimark\BETAOPT.zip"

$SourceTable = @()

if ((Test-Path -Path "D:\apps\vss") -eq $true) {
    $SourceTable += "D:\apps\vss"
}

if ((Test-Path -Path "D:\apps\AVImark") -eq $true) {
    $SourceTable += "D:\apps\AVImark"
}

if ((Test-Path -Path "C:\AVImark") -eq $true) {
    $SourceTable += "C:\AVImark"
}

if ((Test-Path -Path "D:\AVImark") -eq $true) {
    $SourceTable += "D:\AVImark"
}

if ((Test-Path -Path "E:\AVImark") -eq $true) {
    $SourceTable += "E:\AVImark"
}

if ((Test-Path -Path "F:\AVImark") -eq $true) {
    $SourceTable += "F:\AVImark"
}

$TotalSteps = 6 + (($SourceTable).Count)

$Step = 1
$Activity = "Running AVImark Auto Update Tasks"
$Task = "Downloading the AVImark update archive: '$ZipUpdate'"

Write-Progress -Id 1 -Activity $Activity -Status $Task -PercentComplete ($Step / $TotalSteps * 100)

Start-Sleep -Seconds 2

if ((Test-Path -Path $DestinationPath) -eq $false) {
    New-Item -Path $DestinationPath -ItemType Container -Force | Out-Null
    $Download = (New-Object -COM Shell.Application).NameSpace($ZIPUpdate)
    $Destination = (New-Object -COM Shell.Application).NameSpace($DestinationPath)
    $Destination.CopyHere($Download, 16)
}

if ((Test-Path -Path $AVIMarkZip) -eq $false) {
    $Download = (New-Object -COM Shell.Application).NameSpace($ZIPUpdate)
    $Destination = (New-Object -COM Shell.Application).NameSpace($DestinationPath)
    $Destination.CopyHere($Download, 16)
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

Clear-Host

foreach ($Source in $SourceTable) {

    $Step = $Step + 1
    $Task = "Backing up AVImark directory: '$Source'"

    Write-Progress -Id 1 -Activity $Activity -Status $Task -PercentComplete ($Step / $TotalSteps * 100)

    $BackInitial = (Get-ChildItem $Source -Recurse).Count

    $Backup = (New-Object -COM Shell.Application).NameSpace($Source)

    $DestinationPath = "$Source\Backup\AVImark"

    if ((Test-Path -Path $DestinationPath) -eq $false) {
        New-Item -Path $DestinationPath -ItemType Container -Force | Out-Null
    }

    $Destination = (New-Object -COM Shell.Application).NameSpace($DestinationPath)
    $Destination.CopyHere($Backup.Items(), 16)

    $Step = $Step + 1
    $Task = "Verifiying backup directory: '$DestinationPath'"

    Write-Progress -Id 1 -Activity $Activity -Status $Task -PercentComplete ($Step / $TotalSteps * 100)

    Start-Sleep -Seconds 2

    $BackCheck = (Get-ChildItem $DestinationPath -Recurse).Count

    if ($BackCheck -eq $BackInitial) {

        $Step = $Step + 1
        $Task = "Extracting ZIP archive to '$Source'"

        Write-Progress -Id 1 -Activity $Activity -Status $Task -PercentComplete ($Step / $TotalSteps * 100)

        $ZipFile = (New-Object -COM Shell.Application).NameSpace($AVIMarkZip)
        $Destination = (New-Object -COM Shell.Application).NameSpace($Source)

        if ((Test-Path -Path $AVIMarkZip) -eq $true) {
            $Destination.CopyHere($ZipFile.Items(), 16)
            
        } else {
            Write-Error "AVImark update file did not transfer properly!"
        }

    } else {
            Write-Error "AVImark backup did not complete successfully. Please, manually backup to '$Source' and re-run script."
    }

    $Task = "completed"

    Write-Progress -Id 1 -Activity $Activity -Status $Task -Completed

    $Manual = [System.Windows.MessageBox]::Show("Please, continue manual update of Avimark, press 'OK' to launch AVImark Server Guardian","Continue Manually","OKCancel","Information")

        switch ($Manual) {

        'OK' {
        
            Set-Location $Source
            .\AVImarkGuardian.exe

        }

        'Cancel' {

            Exit

        }
    }
}

$Continue = [System.Windows.MessageBox]::Show("After manual update, choose 'OK' to turn on all AVImark services","Start Services","OKCancel","Information")

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
            
        Remove-Item -Path $AVIMarkZip -Force
        
        }
    }

    'Cancel' {

        Exit

    }
}

Clear-Host

$End = Read-Host "All processes have completed! Press any key to end script: "

if ($End -ne "") {
    Exit
} else {
    Write-Error "Expression cannot be null value!"
    Return
}