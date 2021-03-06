function Start-Service-Progress {
    param(
        $Service
    )
    
    $ScriptBlock = {
    param(
        $Service
    )
    
        Stop-Service -Name $Service
            
    }
    
    Start-Job -Name $Service -ScriptBlock $ScriptBlock | Out-Null
    
    Write-Host "Stopping service " -NoNewLine
    Write-Host $Service -NoNewline
    Write-Host " { " -NoNewline
    
    while ((Get-Job -Name $Service | Select-Object -ExpandProperty State) -eq "Running") {
    
        Start-Sleep -Seconds 1
        Write-Host "#" -NoNewLine 
    
    }
    
    Write-Host " }" -NoNewLine
    
    if ((Get-Service | ? {$_.ServiceName -like $Service} | Select-Object -ExpandProperty Status) -eq "Stopped") {
    
        Write-Host " {" -NoNewLine
        Write-Host " OK " -NoNewLine -Foregroundcolor Green
        Write-Host "}" -NoNewLine
        
        } else {
        
        Write-Host " {" -NoNewLine
        Write-Host " FAIL " -NoNewLine -Foregroundcolor Red
        Write-Host "}" -NoNewLine
        
        }

    Remove-Job -Name $Service
        
        "`n"
}

$Vetstoria = Get-Service | ? {$_.ServiceName -like "Vets*"}
$IDEXXService = Get-Service | ? {$_.DisplayName -like "IDEXX*"}
$Vetlogic = Get-Service | ? {$_.ServiceName -like "VetLogic*"}
$VDS = Get-Service | ? {$_.ServiceName -like "VDSD*"}
$AVImarkServer = Get-Service | ? {$_.ServiceName -eq "AVImarkServer"}

foreach ($Service in $Vetstoria) {
    Stop-Service-Progress -Service $Service
}

foreach ($Service in $IDEXXService) {
    Stop-Service-Progress -Service $Service
}

foreach ($Service in $Vetlogic) {
    Stop-Service-Progress -Service $Service
}

foreach ($Service in $VDS) {
    Stop-Service-Progress -Service $Service
}

foreach ($Service in $AVImarkServer) {
    Stop-Service-Progress -Service $Service
}