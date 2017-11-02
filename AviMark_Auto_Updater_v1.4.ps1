### Startup Definitions ### 

Add-Type -AssemblyName PresentationFramework
$host.ui.rawui.WindowTitle = "AVIMark Auto-Updater"

### Function Definitions ###
Function Write-Log {
    [CmdletBinding()]
    Param(
    [Parameter(
	    Mandatory=$False)]
    [ValidateSet("INFO","WARN","ERROR","FATAL","DEBUG")]
    [String]
    $Level = "INFO",

    [Parameter(
        Mandatory=$True)]
    [string]$Message,
    
    [Parameter(
        Mandatory=$True,
        ValueFromPipeline=$True,
	    ValueFromPipelineByPropertyName=$True)]
    [string]$Variable,

    [Parameter(
	Mandatory=$True,
	ValueFromPipeline=$True,
	ValueFromPipelineByPropertyName=$True)]
    [string]$Path
    )

    $Stamp = (Get-Date).toString("MM/dd/yyyy HH:mm:ss")
    $Line = "$Stamp $Level - $Variable : $Message"
    If($Path) {
        Add-Content $Path -Value $Line
    }
    Else {
        Write-Output $Line
    }
}

### Script Block ####

$ZIPUpdate = "\\10.252.70.3\Users\Public\Downloads\AVImark 2016.2.7.zip"
$DestinationPath = "C:\Temp\Avimark"
$AVIMarkZip = "C:\Temp\Avimark\AVImark 2016.2.7.zip"
$Log = "C:\Temp\avimarkupdate.log"

if (!(Test-Path -Path $Log)) {
    New-Item -Path $Log -ItemType File | Out-Null
}

$SourceTable = @()

$Sources = @("D:\apps\vss", "D:\apps\AVImark", "C:\AVImark", "D:\AVImark", "E:\AVImark", "F:\AVImark")

foreach ($Path in $Sources) {
   if (Test-Path -Path $Path) {
       Write-Log -Path $Log -Level INFO -Variable $Log -Message "Found '$Path' as valid update path"
       $SourceTable += $Path
   }
}

$TotalSteps = 6 + (($SourceTable).Count)

$Step = 0
$Activity = "Running AVImark Auto Update Tasks"
$Task = "Downloading the AVImark update archive: '$ZipUpdate'"

Write-Progress -Id 1 -Activity $Activity -Status $Task -PercentComplete ($Step / $TotalSteps * 100)

Start-Sleep -Seconds 2

if (Test-Connection -Count 2 -ComputerName "10.252.70.3" -ErrorAction Stop) {

    if (!(Test-Path -Path $DestinationPath)) {
        New-Item -Path $DestinationPath -ItemType Container -Force | Out-Null
        $Download = (New-Object -COM Shell.Application).NameSpace($ZIPUpdate)
        $Destination = (New-Object -COM Shell.Application).NameSpace($DestinationPath)
        $Destination.CopyHere($Download, 16)
    }

    if (!(Test-Path -Path $AVIMarkZip)) {
        $Download = (New-Object -COM Shell.Application).NameSpace($ZIPUpdate)
        $Destination = (New-Object -COM Shell.Application).NameSpace($DestinationPath)
        $Destination.CopyHere($Download, 16)
    }

} else {

    Write-Error -Message "Unable to establish connection to server hosting update files"
    Write-Log -Level ERROR -Path $Log -Message "Unable to establish connection to server hosting update files"
    $End = Read-Host "Press any key to end script"
        if ($End -ne "") {
        Exit
        }
}

$Step ++
$Task = "Killing all AVImark related tasks:"

Write-Progress -Id 1 -Activity $Activity -Status $Task -PercentComplete ($Step / $TotalSteps * 100)

Start-Sleep -Seconds 2

$AVIMarkProcess = Get-Process | ? {$_.ProcessName -like "AVIM*"}

if (($AVIMarkProcess | Select-Object -ExpandProperty Name) -like "AVIM*") {
     
    foreach ($Process in ($AVIMarkProcess | Select-Object -ExpandProperty ProcessName)) {
        Stop-Process -Name $Process -Force
        Write-Log -Level INFO -Path $Log -Variable $Process -Message "Successfully stopped process $Process"
    }
}

$MPSProcess = Get-Process | ? {$_.ProcessName -like "MPS*"}

if (($MPSProcess | Select-Object -ExpandProperty ProcessName) -like "MPS*") {

    foreach ($Process in ($MPSProcess | Select-Object -ExpandProperty Name)) {
        Stop-Process -Name $Process -Force
        Write-Log -Level INFO -Path $Log -Variable $Process -Message "Successfully stopped process $Process"
    }
}

$Step ++
$Task = "Stopping all AVImark services and dependencies:"

Write-Progress -Id 1 -Activity $Activity -Status $Task -PercentComplete ($Step / $TotalSteps * 100)

Start-Sleep -Seconds 2

$VDS = Get-Service | ? {$_.ServiceName -like "VDSD*"}

if (($VDS | Select-Object -ExpandProperty Status) -eq "Running") {
    
    foreach ($Service in ($VDS | Select-Object -ExpandProperty Name)) {
        Stop-Service -Name $Service -Force -ErrorAction SilentlyContinue | Out-Null
        Write-Log -Level INFO -Path $Log -Variable $Service -Message "Successfully stopped service $Service"
    }
}

$Vetstoria = Get-Service | ? {$_.ServiceName -like "Vets*"}

if (($Vetstoria | Select-Object -ExpandProperty Status) -eq "Running") {
    
    foreach ($Service in ($Vetstoria | Select-Object -ExpandProperty Name)) {
        Stop-Service -Name $Service -Force -ErrorAction SilentlyContinue | Out-Null
        Write-Log -Level INFO -Path $Log -Variable $Service -Message "Successfully stopped service $Service"
    }
}

$Vetlogic = Get-Service | ? {$_.ServiceName -like "VetLogic*"}

if (($Vetlogic | Select-Object -ExpandProperty Status) -eq "Running") {
    
    foreach ($Service in ($Vetlogic | Select-Object -ExpandProperty Name)) {
        Stop-Service -Name $Service -Force -ErrorAction SilentlyContinue | Out-Null
        Write-Log -Level INFO -Path $Log -Variable $Service -Message "Successfully stopped service $Service"
    }
}

$IDEXXService = Get-Service | ? {$_.DisplayName -like "IDEXX*"}

if (($IDEXXService | Select-Object -ExpandProperty Status) -eq "Running" ) {
    
    foreach ($Service in ($IDEXXService | Select-Object -ExpandProperty Name)) {
        Stop-Service -Name $Service -Force -ErrorAction SilentlyContinue | Out-Null
        Write-Log -Level INFO -Path $Log -Variable $Service -Message "Successfully stopped service $Service"
    }
}

$AVIMarkServer = Get-Service | ? {$_.ServiceName -like "AVIM*"}

if (($AVIMarkServer | Select-Object -ExpandProperty Status) -eq "Running") {
    
    foreach ($Service in ($AVIMarkServer | Select-Object -ExpandProperty Name)) {
        Stop-Service -Name $Service -Force -ErrorAction SilentlyContinue | Out-Null
        Write-Log -Level INFO -Path $Log -Variable $Service -Message "Successfully stopped service $Service"
    }
}

Clear-Host

foreach ($Source in $SourceTable) {

    $Destination = $Source -replace "([A-Z])\w+"

    if ($Destination -like "D:\\") {
    $Destination = "D:\"
    }

    $Destination = $Destination + "Backup"

    $Step ++
    $Task = "Backing up AVImark directory to '$Destination'"

    Write-Progress -Id 1 -Activity $Activity -Status $Task -PercentComplete ($Step / $TotalSteps * 100)

    $BackInitial = (Get-ChildItem $Source).Count
    
    if(!(Test-Path -Path $Destination)) {
        New-Item -Path $Destination -ItemType Container -Force | Out-Null

            if (Test-Path $Destination) {

                Write-Log -Level INFO -Path $Log -Variable $Destination -Message "Successfully created directory $Destination"

            } else {

                Write-Log -Level ERROR -Path $Log -Variable $Destination -Message "Unable to create directory $Destination. Error Reason: $Error[0]"
                
                $End = Read-Host "Press any key to end script"
                    if ($End -ne "") {
                    Exit
                }

            }
        }

    $Count = (Get-ChildItem $Source).Count

    $Operation = 0

    foreach ($File in (Get-ChildItem $Source)) {

        $File = $Source + "\" + $File
        $Operation ++
        Write-Progress -Id 2 -Activity ('Copying data: ' + $File) -Status 'Progress:' -PercentComplete ($Operation / $Count * 100)
        Copy-Item $File -Destination $Destination -Force

            if (Test-Path $File) {

                Write-Log -Level INFO -Path $Log -Variable $File -Message "Successfully copied $File to $Destination"
                
            } else {

                Write-Log -Level ERROR -Path $Log -Variable $File -Message "Error in copying $File to $Destination. Error Reason: $Error[0]"
                $End = Read-Host "Press any key to end script"
                    if ($End -ne "") {
                    Exit
                    }        
            }
    }

    Write-Progress -Id 2 -Activity "Completed" -Status "Completed" -Completed

    $Step ++
    $Task = "Verifiying backup directory: '$DestinationPath'"

    Write-Progress -Id 1 -Activity $Activity -Status $Task -PercentComplete ($Step / $TotalSteps * 100)

    Start-Sleep -Seconds 2

    $BackCheck = (Get-ChildItem $Destination).Count

    if ($BackCheck -eq $BackInitial) {

        Write-Log -Level INFO -Path $Log -Message "Successfully verified backup now proceeding to unzip archive"

        $Step ++
        $Task = "Extracting ZIP archive to '$Source'"

        Write-Progress -Id 1 -Activity $Activity -Status $Task -PercentComplete ($Step / $TotalSteps * 100)

        $ZipFile = (New-Object -COM Shell.Application).NameSpace($AVIMarkZip)
        $Destination = (New-Object -COM Shell.Application).NameSpace($Source)

        if (Test-Path -Path $AVIMarkZip) {

            $Destination.CopyHere($ZipFile.Items(), 16)
            
        } else {

            Write-Error "AVImark update file did not download from server properly. Please, verify connection to server, and try again."
            Write-Log -Level ERROR -Path $Log -Mesage "AVImark update file did not download from server properly. Please, verify connection to server, and try again."
            $End = Read-Host "Press any key to end script"
                if ($End -ne "") {
                Exit
                }
        }

    } else {

            Write-Error "AVImark backup did not complete successfully. Please, manually backup to '$Source' and re-run script."
            Write-Log -Level ERROR -Path $Log -Variable $Source -Mesage "AVImark backup did not complete successfully. Please, manually backup to '$Source' and re-run script."
            $End = Read-Host "Press any key to end script"
                if ($End -ne "") {
                Exit
                }
    }

    $Task = "Completed"

    Write-Log -Level INFO -Path $Log -Message "All automatic update processes completed. Manual user intervention is required to complete"

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
        }    

        foreach ($Service in ($VDS | Select-Object -ExpandProperty Name)) {
            Start-Service -Name $Service -Verbose
        }

        foreach ($Service in ($Vetlogic | Select-Object -ExpandProperty Name)) {
            Start-Service -Name $Service -Verbose
        }    

        Remove-Item -Path $AVIMarkZip -Force
    }

    'Cancel' {

        Exit

    }
}

Clear-Host

Write-Host "All update processes have been completed!" -ForegroundColor Green

$End = Read-Host "Press any key to end script"

if ($End -ne "") {
    Exit
}