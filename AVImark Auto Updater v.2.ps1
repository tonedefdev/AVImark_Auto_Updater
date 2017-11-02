Clear-Host

Add-Type -AssemblyName PresentationFramework

Function Get-Zip ($initialDirectory)
{
    [System.Reflection.Assembly]::LoadWithPartialName("System.windows.forms") | Out-Null
    
    $OpenFileDialog = New-Object System.Windows.Forms.OpenFileDialog
    $OpenFileDialog.Title = "Select AVIMark Update Zip File"
    $OpenFileDialog.initialDirectory = $initialDirectory
    $OpenFileDialog.filter = "ZIP (*.zip)| *.zip"
    $OpenFileDialog.ShowDialog() | Out-Null
    $OpenFileDialog.filename
}

Function Get-CSV ($initialDirectory)
{
    [System.Reflection.Assembly]::LoadWithPartialName("System.windows.forms") | Out-Null
    
    $OpenFileDialog = New-Object System.Windows.Forms.OpenFileDialog
    $OpenFileDialog.Title = "Navigate to 'AVImark Auto Updater Server List'"
    $OpenFileDialog.initialDirectory = $initialDirectory
    $OpenFileDialog.filter = "CSV (*.csv)| *.csv"
    $OpenFileDialog.ShowDialog() | Out-Null
    $OpenFileDialog.filename
}

$ExtractSB = {
param (
    $Directory,
    $Zip
)

    $AVIMarkZip = "C:\Temp\AVImark\$Zip"
    $UnzipPath = "C:\Temp\AVImark\$Directory"

        if (!(Test-Path $UnzipPath)) {

            New-Item -Path $UnzipPath -ItemType Container -Force | Out-Null

        }

        $ZipFile = (New-Object -COM Shell.Application).NameSpace($AVIMarkZip)
        $Destination = (New-Object -COM Shell.Application).NameSpace($UnzipPath)

        if (Test-Path -Path $AVIMarkZip) {
    
            $Destination.CopyHere($ZipFile.Items(), 16)
        
        }

}

$StopServiceSB = {
param(
    $Log
)

    $Log = "C:\temp\avimarkupdate.log"

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

    $AVIMarkServer = Get-Service | ? {$_.DisplayName -like "AVIm*"}

    if (($AVIMarkServer | Select-Object -ExpandProperty Status) -eq "Running") {
    
        foreach ($Service in ($AVIMarkServer | Select-Object -ExpandProperty Name)) {
            Stop-Service -Name $Service -Force -ErrorAction SilentlyContinue -WarningAction SilentlyContinue | Out-Null
            Write-Log -Level INFO -Path $Log -Variable $Service -Message "Successfully stopped service $Service"
        }
    }

    $VDS = Get-Service | ? {$_.ServiceName -like "VDSD*"}

    if (($VDS | Select-Object -ExpandProperty Status) -eq "Running") {
    
        foreach ($Service in ($VDS | Select-Object -ExpandProperty Name)) {
            Stop-Service -Name $Service -Force -ErrorAction SilentlyContinue -WarningAction SilentlyContinue | Out-Null
            Write-Log -Level INFO -Path $Log -Variable $Service -Message "Successfully stopped service $Service"
        }
    }

    $Vetstoria = Get-Service | ? {$_.ServiceName -like "Vets*"}

    if (($Vetstoria | Select-Object -ExpandProperty Status) -eq "Running") {
    
        foreach ($Service in ($Vetstoria | Select-Object -ExpandProperty Name)) {
            Stop-Service -Name $Service -Force -ErrorAction SilentlyContinue -WarningAction SilentlyContinue | Out-Null
            Write-Log -Level INFO -Path $Log -Variable $Service -Message "Successfully stopped service $Service"
        }
    }

    $IDEXXService = Get-Service | ? {$_.DisplayName -like "IDEXX*"}

    if (($IDEXXService | Select-Object -ExpandProperty Status) -eq "Running" ) {
    
        foreach ($Service in ($IDEXXService | Select-Object -ExpandProperty Name)) {
            Stop-Service -Name $Service -Force -ErrorAction SilentlyContinue -WarningAction SilentlyContinue | Out-Null
            Write-Log -Level INFO -Path $Log -Variable $Service -Message "Successfully stopped service $Service"
        }
    }
}

$StopProcessSB = {

    $Log = "C:\temp\avimarkupdate.log"

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

    $AVIMarkProcess = Get-Process | ? {$_.ProcessName -like "AVIM*"}

    if (($AVIMarkProcess | Select-Object -ExpandProperty Name) -like "AVIM*") {
     
        foreach ($Process in ($AVIMarkProcess | Select-Object -ExpandProperty ProcessName)) {
            Stop-Process -Name $Process -Force -WarningAction SilentlyContinue -ErrorAction SilentlyContinue
            Write-Log -Level INFO -Path $Log -Variable $Process -Message "Successfully stopped process $Process"
        }
    }

    $MPSProcess = Get-Process | ? {$_.ProcessName -like "MPS*"}

    if (($MPSProcess | Select-Object -ExpandProperty ProcessName) -like "MPS*") {

        foreach ($Process in ($MPSProcess | Select-Object -ExpandProperty Name)) {
            Stop-Process -Name $Process -Force -WarningAction SilentlyContinue -ErrorAction SilentlyContinue
            Write-Log -Level INFO -Path $Log -Variable $Process -Message "Successfully stopped process $Process"
        }
    }
}


$ScriptBlock = {
param(
    $ExtractSB,
    $StopServiceSB,
    $StopProcessSB,
    $Computer,
    $ID,
    $Source,
    $Zip
)
    Function Get-Formatted-Date {
        $day =  (Get-Date).Day

        $month = (Get-Date).Month

        $year = (Get-Date).Year

    $date = "$month" + "_" + "$day" + "_" + "$year"
    $date
    } 

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
    
    Import-Module BitsTransfer

    $Directory = $Zip -replace '.zip$' 

    $Destination = "\\$Computer\C$\Temp\AVImark"
    $Log = "\\$Computer\C$\Temp\avimarkupdate.log"
    $Extracted = "\\$Computer\C$\temp\AVImark\$Directory"
    $AVImarkZip = "\\$Computer\C$\temp\AVImark\$Zip"

    if (!(Test-Path -Path $Log)) {
        New-Item -Path $Log -ItemType File -Force | Out-Null
    }

    Write-Log -Path $Log -Level INFO -Variable $Log -Message "Connected to machine and created $Log"

    if (!(Test-Path -Path $Destination)) {
        New-Item -Path $Destination -ItemType Container -Force | Out-Null
        Write-Log -Path $Log -Level INFO -Variable $Destination -Message "Created directory $Destination"
    }

    Write-Log -Path $Log -Level INFO -Variable $Source -Message "Starting transfer of $Source to $Destination"

    if (!(Test-Path $AVImarkZIP)) {
        Start-BitsTransfer -Source $Source -Destination $Destination -Description ($Computer + ' - Copying data: ' + $Source) -DisplayName "Transferring AVImark Zip File to $Destination"
    }

    if (Test-Path $AVImarkZIP) {

        Write-Log -Path $Log -Level INFO -Variable $AVImarkZIP -Message "Successfully uploaded $AVImarkZIP"

    } else {

        Write-Log -Path $Log -Level ERROR -Variable $Source -Message "There was an error uploading $AVImarkZIP. Check connection to server and try again"

    }


    $SourceTable = @()

    $Sources = @("\\$Computer\D$\AVImark", "\\$Computer\D$\apps\vss", "\\$Computer\D$\apps\AVImark", "\\$Computer\C$\AVImark", "\\$Computer\E$\AVImark", "\\$Computer\F$\AVImark", "\\$Computer\G$\AVImark")

    foreach ($Path in $Sources) {

        if (Test-Path -Path $Path) {
           Write-Log -Path $Log -Level INFO -Variable $Path -Message "Found '$Path' as valid update path"
           $SourceTable += $Path       
        }
    }

    if ($SourceTable -eq $null) {
        Write-Log -Path $Log -Level ERROR -Variable $Path -Message "No valid AVImark paths found"
        Send-MailMessage -Subject "AVIMark Update Error" -Body "No valid AVImark paths found" -From "helpdesk@nvanet.com" -To "b6z9m6w3x3z7k2m1@nva-it-team.slack.com" -Attachments $Log -SmtpServer "nvaexch.nva.local"
        Continue
    }

    $TotalSteps = 6 + (($SourceTable).Count)

    $Step = 1
    $Activity = "Running AVImark Auto Update Tasks"
    $Task = "$Computer - Extracting '$AVIMarkZip' archive to '$Extracted'"

    Write-Progress -Id $ID -Activity $Activity -Status $Task -PercentComplete ($Step / $TotalSteps * 100)

    if (!(Test-Path $Extracted)) {
    Invoke-Command -ComputerName $Computer -ScriptBlock $ExtractSB -ArgumentList $Directory, $Zip
    }

    if ((Test-Path $Extracted.Count) -gt 20) {

        Write-Log -Path $Log -Level INFO -Variable $AVImarkZip -Message "$AVImarkZip extracted successfully"

    } else {
        
        Write-Log -Path $Log -Level ERROR -Variable $AVImarkZip -Message "$AVImarkZip had issues extracting. Check the directory on the server and try again"
        Send-MailMessage -Subject "AVIMark Update Error" -Body "$Computer - $AVImarkZip had issues extracting. Check the directory on the server and try again" -From "helpdesk@nvanet.com" -To "b6z9m6w3x3z7k2m1@nva-it-team.slack.com" -Attachments $Log -SmtpServer "nvaexch.nva.local"       
        Write-Progress -Id $ID -Activity $Activity -Status "Error" -Completed
        Continue

    }

    $Step ++
    $Task = "$Computer - Stopping all AVImark services and dependencies:"

    Write-Progress -Id $ID -Activity $Activity -Status $Task -PercentComplete ($Step / $TotalSteps * 100)

    Start-Sleep -Seconds 2

    Invoke-Command -ComputerName $Computer -ScriptBlock $StopServiceSB -ArgumentList $Log

    $Step ++
    $Task = "$Computer - Killing all AVImark related tasks:"

    Write-Progress -Id $ID -Activity $Activity -Status $Task -PercentComplete ($Step / $TotalSteps * 100)

    Start-Sleep -Seconds 2

    Invoke-Command -ComputerName $Computer -ScriptBlock $StopProcessSB -ArgumentList $Log

    foreach ($Source in $SourceTable) {

        $Destination = $Source -replace "([A-Z])\w+$"

        $Destination = $Destination + "Backup_" + (Get-Formatted-Date)

        $Step ++
        $Task = "$Computer - Backing up AVImark directory to '$Destination'"
        Write-Progress -Id $ID -Activity $Activity -Status $Task -PercentComplete ($Step / $TotalSteps * 100)

        if(!(Test-Path -Path $Destination)) {
    
        New-Item -Path $Destination -ItemType Container -Force | Out-Null

            if (Test-Path $Destination) {

                Write-Log -Level INFO -Path $Log -Variable $Destination -Message "Successfully created directory '$Destination'"

            } else {

                Write-Log -Level ERROR -Path $Log -Variable $Destination -Message ("Unable to create directory '$Destination' - Error Reason: " + $Error[0])
                
                    
            }
        }

        $Count = (Get-ChildItem $Source).Count
        $Operation = 0

            foreach ($File in (Get-ChildItem $Source)) {

                $File = $Source + "\" + $File
                $Operation ++
                Write-Progress -Id $ID -Activity ($Computer  + ' - Backing up ' + $File) -Status 'Progress:' -PercentComplete ($Operation / $Count * 100)
                Copy-Item -LiteralPath $File -Destination $Destination -Force

                    if (Test-Path $File) {

                        Write-Log -Level INFO -Path $Log -Variable $File -Message "Successfully backed up '$File' to '$Destination'"
                
                    } else {

                        Write-Log -Level ERROR -Path $Log -Variable $File -Message ("Error in backing up '$File' to '$Destination' - Error Reason: " + $Error[0])

                    }
            }

        $BackCheck = (Get-ChildItem $Destination).Count

        Write-Progress -Id $ID -Activity "Completed" -Status "Completed" -Completed

        $Step ++
        $Task = "$Computer - Verifiying backup directory: '$Destination'"

        Write-Progress -Id $ID -Activity $Activity -Status $Task -PercentComplete ($Step / $TotalSteps * 100)

        if ($BackCheck -eq $Count) {

            Write-Log -Level INFO -Path $Log -Message "Successfully verified backup now proceeding to copy update files"

            $Step ++
            $Task = "$Computer - Copying update files to: '$Source'"

            Write-Progress -Id $ID -Activity $Activity -Status $Task -PercentComplete ($Step / $TotalSteps * 100)

            $Count1 = (Get-ChildItem $Extracted).Count

            $Operation = 0

            if (Test-Path -Path $Extracted) {

                foreach ($File in (Get-ChildItem $Extracted)) {

                    $File = $Extracted + "\" + $File
                    $Operation ++
                    Write-Progress -Id $ID -Activity ($Computer + ' - Copying update: ' + $File) -Status 'Progress:' -PercentComplete ($Operation / $Count1 * 100)
                    Copy-Item -LiteralPath $File -Destination $Source -Force

                    if (Test-Path $File) {
                    
                            Write-Log -Level INFO -Path $Log -Variable $File -Message "Successfully copied '$File' to '$Source'"
                        
                        } else {
                            
                            Write-Log -Level ERROR -Path $Log -Variable $File -Message ("Error in copying '$File' to '$Source' - Error Reason: " + $Error[0])
                        }
                }
                 
            } else {

            Write-Log -Level ERROR -Path $Log -Message "AVImark update file did not extract properly."    
            Send-MailMessage -Subject "AVIMark Update Error" -Body "$Computer : AVImark update file did not extract properly. Ensure '$Extracted' directory was created properly" -From "helpdesk@nvanet.com" -To "b6z9m6w3x3z7k2m1@nva-it-team.slack.com" -Attachments $Log -SmtpServer "nvaexch.nva.local"
            Write-Progress -Id $ID -Activity "Completed" -Completed           
            Continue

            }

        } else {

            Write-Log -Level ERROR -Path $Log -Message "AVImark backup did not complete successfully. Backup Count: $Count - Backup Check Count: $Backcheck"
            Send-MailMessage -Subject "AVImark Update Error" -Body "$Computer : Backup Count: $Count - Backup Check Count: $Backcheck - The backup count/check was off. Make sure the backup is not backing up to a pre-existing directory with additional items that might throw the count off" -From "helpdesk@nvanet.com" -To "b6z9m6w3x3z7k2m1@nva-it-team.slack.com" -Attachments $Log -SmtpServer "nvaexch.nva.local"            
            Continue
        }

        Write-Log -Level INFO -Path $Log -Message "All automatic update processes completed. Manual user intervention is required to complete" 
        Send-MailMessage -Subject "AVImark Auto-Updater" -Body "$Computer - All automatic update processes completed. Remote into server to finish updates" -From "helpdesk@nvanet.com" -To "b6z9m6w3x3z7k2m1@nva-it-team.slack.com" -Attachments $Log -SmtpServer "nvaexch.nva.local"               
        Write-Progress -Id $ID -Activity $Activity -Status "$Computer - All automatic update processes completed. Manual user intervention is required to complete" -Completed

    }
}

$Throttle = 5
$initialSessionState = [System.Management.Automation.Runspaces.InitialSessionState]::CreateDefault()
$RunspacePool = [RunspaceFactory]::CreateRunspacePool(1, $Throttle,$initialSessionState,$Host)
$RunspacePool.Open()
$Jobs = @()
$Computers = Import-CSV -Path (Get-CSV)

$Connected = @()
$Computers | foreach {

    $Computer = $_.Server
    
    $TestConnect = [bool](Test-Connection -Count 1 -ComputerName $Computer -ErrorAction SilentlyContinue)
    if ($TestConnect -eq $true) {
        $Connect = New-Object System.Object
        $Connect | Add-Member -Type NoteProperty -Name Server -Value $Computer
        $Connect | Add-Member -Type NoteProperty -Name Connection -Value $TestConnect
        $Connected += $Connect
    }
}

$Connected | Out-GridView -Title "Servers Currently Online & Available to be Updated"

$IDRange = @(1..$Connected.Count)
$i = 0
$Source = Get-Zip
$Zip = $Source -match '(([A-Z])\w+ ([1-9])\d\d\d.\d.\d.zip)$'
$Zip = $Matches.Values | Select-Object -Last 1

if ($Zip -eq $null) {

    $Exit = [System.Windows.MessageBox]::Show("'$Source' is not properly formatted. Make sure the .ZIP filename adheres to 'FILENAME xxxx.x.x' format. Rename the ZIP file and try again.","ZIP Filename Error","OK","Error")

    switch ($Exit) {

        'OK' { Exit }
        
    }
}

foreach ($Computer in $Connected) {

   $Computer = $Computer.Server

   $ID = $IDRange[$i]

   $Job = [powershell]::Create().AddScript($ScriptBlock).AddArgument($ExtractSB).AddArgument($StopServiceSB).AddArgument($StopProcessSB).AddArgument($Computer).AddArgument($ID).AddArgument($Source).AddArgument($Zip)
   $Job.RunspacePool = $RunspacePool
   $Jobs += New-Object PSObject -Property @{
      Pipe = $Job
      Result = $Job.BeginInvoke()
   }

   $i++
}

$Results = @()

foreach ($Job in $Jobs) {
   
    $Results += $Job.Pipe.EndInvoke($Job.Result)

}