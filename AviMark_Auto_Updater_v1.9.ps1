### Startup Definitions ### 

Add-Type -AssemblyName PresentationFramework
$host.ui.rawui.WindowTitle = "AVIMark Auto-Updater"

### Function Definitions ###

Function Write-Log {
    [CmdletBinding()]
    Param(
    [Parameter(
	    Mandatory=$False)]
    [ValidateSet("INFO","WARNING","ERROR","FATAL","DEBUG")]
    [String]
    $Level = "INFO",

    [Parameter(
        Mandatory=$True)]
    [string]$Message,
    
    [Parameter(
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
    $Line = "$Stamp $Level - $Message"
    If($Path) {
        Add-Content $Path -Value $Line
    }
    Else {
        Write-Output $Line
    }
}

function Get-Formatted-Date {
    $day =  (Get-Date).Day

    $month = (Get-Date).Month

    $year = (Get-Date).Year

$date = "$month" + "_" + "$day" + "_" + "$year"
$date
}

function Start-Service-Progress {
    param(
        $Service,
        $Log
    )
    
    $ScriptBlock = {
    param(
        $Service
    )
    
        Start-Service -Name $Service
            
    }
    
    Start-Job -Name $Service -ScriptBlock $ScriptBlock | Out-Null
    
    Write-Host "Starting service " -NoNewLine
    Write-Host $Service -NoNewline
    Write-Host " { " -NoNewline
    
    while ((Get-Job -Name $Service | Select-Object -ExpandProperty State) -eq "Running") {
    
        Start-Sleep -Seconds 1
        Write-Host "#" -NoNewLine 
    
    }
    
    Write-Host " }" -NoNewLine
    
    if ((Get-Service | ? {$_.ServiceName -like $Service} | Select-Object -ExpandProperty Status) -eq "Running") {
    
        Write-Host " {" -NoNewLine
        Write-Host " OK " -NoNewLine -Foregroundcolor Green
        Write-Host "}" -NoNewLine
        Write-Log -Level INFO -Path $Log -Variable $Service -Message "$Service successfully started"
        
        } else {
        
        Write-Host " {" -NoNewLine
        Write-Host " FAIL " -NoNewLine -Foregroundcolor Red
        Write-Host "}" -NoNewLine
        Write-Log -Level ERROR -Path $Log -Variable $Service -Message "$Service did not successfully start"
        
        }
        
        "`n"
}

### Script Block ####

$ZIPUpdate = "\\10.252.70.3\Users\Public\Downloads\AVImark 2016.2.7.zip"
$DestinationPath = "C:\Temp\Avimark"
$AVIMarkZip = "C:\Temp\Avimark\AVImark 2016.2.7.zip"
$Log = "C:\Temp\avimarkupdate.log"
$Extracted = "C:\Temp\Avimark\AVImark 2016.2.7"

if (!(Test-Path -Path "C:\Temp")) {
    New-Item -Path "C:Temp" -ItemType Container -Force | Out-Null
}

if (!(Test-Path -Path $Log)) {
    New-Item -Path $Log -ItemType File -Force | Out-Null
}

$SourceTable = @()

$Sources = @("D:\apps\vss", "D:\apps\AVImark", "C:\AVImark", "D:\AVImark", "E:\AVImark", "F:\AVImark")

foreach ($Path in $Sources) {

    if (Test-Path -Path $Path) {
   
       Write-Log -Path $Log -Level INFO -Variable $Path -Message "Found '$Path' as valid update path"
       $SourceTable += $Path       
    }
}

if ($SourceTable -eq $null) {

    Write-Log -Path $Log -Level INFO -Variable $Path -Message "No valid AVImark paths found"

    $Exit = [System.Windows.MessageBox]::Show("No valid AVImark update paths were found","Update Error","OK","Error")
    
        switch ($Exit) {
    
            'OK' { Exit }
            
        }
}

$TotalSteps = 7 + (($SourceTable).Count)

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

    Write-Log -Level ERROR -Path $Log -Message "Unable to establish connection to server hosting update files"
    $Exit = [System.Windows.MessageBox]::Show("Unable to establish connection to server hosting update files. Please, check connection to '$ZIPUpdate', and try again.","Update Error","OK","Error")

    switch ($Exit) {

        'OK' { Exit }
        
    }
}

$Step ++
$Task = "Extracting '$AVIMarkZip' archive to '$Extracted'"

Write-Progress -Id 1 -Activity $Activity -Status $Task -PercentComplete ($Step / $TotalSteps * 100)

if (!(Test-Path $Extracted)) {

    Write-Log -Level INFO -Path $Log -Message "Created directory '$Extracted'"
    New-Item -Path $Extracted -ItemType Container -Force | Out-Null

}

$ZipFile = (New-Object -COM Shell.Application).NameSpace($AVIMarkZip)
$Destination = (New-Object -COM Shell.Application).NameSpace($Extracted)

if (Test-Path -Path $AVIMarkZip) {
    
        $Destination.CopyHere($ZipFile.Items(), 16)
        
    } else {

        Write-Log -Level ERROR -Path $Log -Message "AVImark update file did not download from server properly. Please, verify connection to server, and try again."
        $Exit = [System.Windows.MessageBox]::Show("AVImark update file did not download from server properly. Please, verify connection to server, and try again.","Update Error","OK","Error")

        switch ($Exit) {

            'OK' { Exit }
            
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

Clear-Host

foreach ($Source in $SourceTable) {

    $Destination = $Source -replace "([A-Z])\w+"

    if ($Destination -like "D:\\") {
    $Destination = "D:\"
    }

    $Destination = $Destination + "Backup_" + (Get-Formatted-Date)

    $Step ++
    $Task = "Backing up AVImark directory to '$Destination'"

    Write-Progress -Id 1 -Activity $Activity -Status $Task -PercentComplete ($Step / $TotalSteps * 100)
    
    if(!(Test-Path -Path $Destination)) {
    
        New-Item -Path $Destination -ItemType Container -Force | Out-Null

            if (Test-Path $Destination) {

                Write-Log -Level INFO -Path $Log -Variable $Destination -Message "Successfully created directory '$Destination'"

            } else {

                Write-Log -Level ERROR -Path $Log -Variable $Destination -Message ("Unable to create directory '$Destination' - Error Reason: " + $Error[0])
                
                $Exit = [System.Windows.MessageBox]::Show("Unable to create directory '$Destination' - Exact error reason can be found in $Log","Update Error","OK","Error")

                switch ($Exit) {

                    'OK' { Exit }
                    
                }

            }
        }

    $Count = (Get-ChildItem $Source).Count

    $Operation = 0

    foreach ($File in (Get-ChildItem $Source)) {

        $File = $Source + "\" + $File
        $Operation ++
        Write-Progress -Id 2 -Activity ('Copying data: ' + $File) -Status 'Progress:' -PercentComplete ($Operation / $Count * 100)
        Copy-Item -LiteralPath $File -Destination $Destination -Force

            if (Test-Path $File) {

                Write-Log -Level INFO -Path $Log -Variable $File -Message "Successfully copied '$File' to '$Destination'"
                
            } else {

                Write-Log -Level ERROR -Path $Log -Variable $File -Message ("Error in copying '$File' to '$Destination' - Error Reason: " + $Error[0])
                $Exit = [System.Windows.MessageBox]::Show("Error in copying '$File' to '$Destination' - Exact error reason can be found in $Log","Update Error","OK","Error")

                switch ($Exit) {

                    'OK' { Exit }
                    
                }
            }
    }

    $BackCheck = (Get-ChildItem $Destination).Count

    Write-Progress -Id 2 -Activity "Completed" -Status "Completed" -Completed

    $Step ++
    $Task = "Verifiying backup directory: '$Destination'"

    Write-Progress -Id 1 -Activity $Activity -Status $Task -PercentComplete ($Step / $TotalSteps * 100)

    Start-Sleep -Seconds 2

    if ($BackCheck -eq $Count) {

        Write-Log -Level INFO -Path $Log -Message "Successfully verified backup now proceeding to copy update files"

        $Step ++
        $Task = "Copying update files to '$Source'"

        Write-Progress -Id 1 -Activity $Activity -Status $Task -PercentComplete ($Step / $TotalSteps * 100)

        $Count = (Get-ChildItem $Extracted).Count

        $Operation = 0

        if (Test-Path -Path $Extracted) {

            foreach ($File in (Get-ChildItem $Extracted)) {

                $File = $Extracted + "\" + $File
                $Operation ++
                Write-Progress -Id 2 -Activity ('Copying data: ' + $File) -Status 'Progress:' -PercentComplete ($Operation / $Count * 100)
                Copy-Item -LiteralPath $File -Destination $Source -Force

                if (Test-Path $File) {
                    
                        Write-Log -Level INFO -Path $Log -Variable $File -Message "Successfully copied '$File' to '$Source'"
                        
                    } else {
        
                        Write-Log -Level ERROR -Path $Log -Variable $File -Message ("Error in copying '$File' to '$Source' - Error Reason: " + $Error[0])
                        $Exit = [System.Windows.MessageBox]::Show("Error in copying '$File' to '$Source' - Exact error reason can be found in $Log","Update Error","OK","Error")
        
                        switch ($Exit) {
        
                            'OK' { Exit }
                            
                        }
                    }
            } 
            
        } else {
        
            Write-Log -Level ERROR -Path $Log -Message "AVImark update file did not extract properly."
            $Exit = [System.Windows.MessageBox]::Show("AVImark update file did not extract properly.","Update Error","OK","Error")

            switch ($Exit) {

                'OK' { Exit }
                
            }
        }

    } else {
    
            Write-Log -Level ERROR -Path $Log -Message "AVImark backup did not complete successfully. Please, manually backup to '$Source' and re-run script."
            $Exit = [System.Windows.MessageBox]::Show("AVImark backup did not complete successfully. Please, manually backup to '$Source' and re-run script.","Update Error","OK","Error")

            switch ($Exit) {

                'OK' { Exit }
                
            }
    }

    Write-Progress -Id 2 -Activity "Completed" -Status "Completed" -Completed

    $Task = "Completed"

    Write-Log -Level INFO -Path $Log -Message "All automatic update processes completed. Manual user intervention is required to complete"

    Write-Progress -Id 1 -Activity $Activity -Status $Task -Completed

    $AVIMarkServer = Get-Service | ? {$_.ServiceName -eq "AVIMarkServer"}

    if ($AVIMarkServer.DisplayName -eq "AVIMark Server") {

        $Manual = [System.Windows.MessageBox]::Show("Please, continue manual update of Avimark, press 'OK' to launch AVImark Server Guardian","Continue Manually","OKCancel","Information")

            switch ($Manual) {

            'OK' {
            
                Set-Location $Source
                .\AVImarkGuardian.exe
                
                Write-Log -Level INFO -Path $Log -Message "User launched AVIMarkGuardian.exe to start manual update"

            }

            'Cancel' {
                
                Write-Log -Level INFO -Path $Log -Message "User canceled launching AVIMarkGaurdian.exe"
                Exit

            }
        }

    } else {

        $Manual = [System.Windows.MessageBox]::Show("Please, continue manual update of Avimark, press 'OK' to launch AVImark","Continue Manually","OKCancel","Information")
        
            switch ($Manual) {

            'OK' {
            
                Set-Location $Source
                .\AVImark.exe
                
                Write-Log -Level INFO -Path $Log -Message "User launched AVIMark.exe to complete update"

            }

            'Cancel' {
                
                Write-Log -Level INFO -Path $Log -Message "User canceled launching AVIMark.exe"
                Exit

            }        
        }
    }
}

$Continue = [System.Windows.MessageBox]::Show("After manual update, choose 'OK' to turn on all AVImark services","Start Services","OKCancel","Information")

    switch ($Continue) {

    'OK' {

            Clear-Host

            foreach ($Service in ($Vetstoria | Select-Object -ExpandProperty Name)) {
                
                Start-Service-Progress -Service $Service -Log $Log

            }

            foreach ($Service in ($AVIMarkServer | Select-Object -ExpandProperty Name)) {
                
                Start-Service-Progress -Service $Service -Log $Log

            }

            foreach ($Service in ($IDEXXService | Select-Object -ExpandProperty Name)) {

                Start-Service-Progress -Service $Service -Log $Log

            }    

            foreach ($Service in ($VDS | Select-Object -ExpandProperty Name)) {
                
                Start-Service-Progress -Service $Service -Log $Log

            }

            foreach ($Service in ($Vetlogic | Select-Object -ExpandProperty Name)) {
            
                Start-Service-Progress -Service $Service -Log $Log
        
            }

        Remove-Item -Path $AVIMarkZip -Force
        
            if (!(Test-Path -Path $AVIMarkZip)) {
            
            Write-Log -Level INFO -Path $Log -Variable $AVIMarkZip -Message "$AVIMArkZip was successfully deleted"
                        
            } else {
            
            Write-Log -Level WARNING -Path $Log -Variable $AVIMarkZip -Message "$AVIMArkZip was not deleted"
            
            }
    }

    'Cancel' {

                Write-Log -Level WARNING -Path $Log -Message "User opted out of automatically restarting services"
                Exit

             }
}

Clear-Host

Write-Log -Level INFO -Path $Log -Message "AVIMark auto-update processes completed"

$Exit = [System.Windows.MessageBox]::Show("All update processes have been completed. Press 'OK' to end","Update Complete","OK","Information")

switch ($Exit) {

    'OK' { Exit }
    
}