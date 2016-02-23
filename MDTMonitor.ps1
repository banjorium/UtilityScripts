$target = "VM1"
$deploymentShare = "E:\Deployment"
 
# Connect to the MDT share
Write-Host "Connecting to MDT share." -ForegroundColor Green
Add-PSSnapin "Microsoft.BDD.PSSNAPIN"
If (!(Test-Path MDT:)) { New-PSDrive -Name MDT -Root $deploymentShare -PSProvider MDTPROVIDER }
 
# Clean up the MDT monitor data for the target VM if it exists
Write-Host "Clearing MDT monitor data." -ForegroundColor Green
Get-MDTMonitorData -Path MDT: | Where-Object { $_.Name -eq $target } | Remove-MDTMonitorData -Path MDT:
 
# Wait for the OS deployment to start before monitoring
# This may require user intervention to boot the VM from the MDT ISO if an OS exists on the vDisk
If ((Test-Path variable:InProgress) -eq $True) { Remove-Variable -Name InProgress }
Do {
    $InProgress = Get-MDTMonitorData -Path MDT: | Where-Object { $_.Name -eq $target }
    If ($InProgress) {
        If ($InProgress.PercentComplete -eq 100) {
            $Seconds = 30
            $tsStarted = $False
            Write-Host "Waiting for task sequence to begin..." -ForegroundColor Green
        } Else {
            $Seconds = 0
            $tsStarted = $True
            Write-Host "Task sequence has begun. Moving to monitoring phase." -ForegroundColor Green
        }
    } Else {
        $Seconds = 30
        $tsStarted = $False
        Write-Host "Waiting for task sequence to begin..." -ForegroundColor Green
    }
    Start-Sleep -Seconds $Seconds
} Until ($tsStarted -eq $True)
 
# Monitor the MDT OS deployment once started
Write-Host "Waiting for task sequence to complete." -ForegroundColor Green
If ((Test-Path variable:InProgress) -eq $True) { Remove-Variable -Name InProgress }
Do {
    $InProgress = Get-MDTMonitorData -Path MDT: | Where-Object { $_.Name -eq $target }
    If ( $InProgress.PercentComplete -lt 100 ) {
        If ( $InProgress.StepName.Length -eq 0 ) { $StatusText = "Waiting for update" } Else { $StatusText = $InProgress.StepName }
        Write-Progress -Activity "Task sequence in progress" -Status $StatusText -PercentComplete $InProgress.PercentComplete
        Switch ($InProgress.PercentComplete) {
            {$_ -lt 25}{$Seconds = 35; break}
            {$_ -lt 50}{$Seconds = 30; break}
            {$_ -lt 75}{$Seconds = 10; break}
            {$_ -lt 100}{$Seconds = 5; break}
        }
        Start-Sleep -Seconds $Seconds
    } Else {
        Write-Progress -Activity "Task sequence complete" -Status $StatusText -PercentComplete 100
    }
} Until ($InProgress.CurrentStep -eq $InProgress.TotalSteps)
Write-Host "Task sequence complete." -ForegroundColor Green