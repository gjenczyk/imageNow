<#############################################################################
# NAME: StopINowServices.ps1
# 
# AUTHOR:  Gregg Jenczyk, UMass (UITS)
# DATE:  2014/05/07
# EMAIL: gjenczyk@umassp.edu
# 
# COMMENT:  This script checks the status of the ImageNow related services
#           and stops them if they are currently running.
#
# THIS MUST BE RUN AS ADMINISTRATOR!!!!
#
# VERSION HISTORY
# 1.0 2014.05.07 Initial Version.
# 1.1 2014.05208 Added handling for situations where services fail to stop.
#
# TO ADD:
# 
# USEFUL SNIPPETS:
# "$(Get-Date) " | Out-File -Append $runLog
#############################################################################>

. "\\ssisnas215c2.umasscs.net\diimages67prd\script\PowerShell\sendmail.ps1"

$env = hostname
$env = $env.ToUpper().substring(2,5)
$serviceArray = @()
$failureArray = @()

$env = hostname
$env = $env.ToUpper().substring(2,5)
$serviceArray = @()
$failureArray = @()
$taskArray = @()
$taskFailArray = @()
$mailingList = @("gjenczyk@umassp.edu","cmatera@umassp.edu")
$stoppedServices = $false
$stoppedTasks = $false
$failedServices = $false
$failedTasks = $false
$successMessage = $false

Get-Service | Where-Object {$_.Name -match "ImageNow*" -and $_.Name -notmatch "ImageNow Se*"} | ForEach-Object {

        if ($_.Status -eq 'Running')
        {
            Stop-Service $_.Name
            if ($?)
            {
                $stoppedService = $_.Name
                "Stopped $startedService at $(Get-Date)" | Out-File -Append -FilePath "d:\inserver6\log\INowServiceMonitor.log"
                $serviceArray += "$stoppedService`n"
                $stoppedServices = $true
            }
            else
            {
                $failedStop = $_.Name
                "Failed to stop $failedStop at $(Get-Date)" | Out-File -Append -FilePath "d:\inserver6\log\INowServiceMonitor.log"
                $failureArray += "$failedStop"
                $stoppedServices = $true
                $failedServices = $true
            }
        }
    
}

Get-Service | Where-Object {$_.Name -match "ImageNow Se*"} | ForEach-Object {

        if ($_.Status -eq 'Running')
        {
            Stop-Service $_.Name
            if ($?)
            {
                $stoppedService = $_.Name
                "Stopped $startedService at $(Get-Date)" | Out-File -Append -FilePath "d:\inserver6\log\INowServiceMonitor.log"
                $serviceArray += "$stoppedService`n"
                $stoppedServices = $true
            }
            else
            {
                $failedStop = $_.Name
                "Failed to stop $failedStop at $(Get-Date)" | Out-File -Append -FilePath "d:\inserver6\log\INowServiceMonitor.log"
                $failureArray += "$failedStop"
                $stoppedServices = $true
                $failedServices = $true
            }
        }
    
}

$tasks = & SCHTASKS /Query /V /FO CSV | ConvertFrom-Csv
foreach ($task in $tasks)
{
    if($task.Author -match "UMASSP*")
    {
        $tony = $task.TaskName
        & SCHTASKS /Change /DISABLE /TN "$tony"
        if ($?)
        {
            "Stopped $tony at $(Get-Date)" | Out-File -Append -FilePath "d:\inserver6\log\INowServiceMonitor.log"
            $taskArray += "$tony`n"
            $stoppedTasks = $true
        }
        else
        {
            "Failed to stop $tony at $(Get-Date)" | Out-File -Append -FilePath "d:\inserver6\log\INowServiceMonitor.log"
             $taskFailArray += "$tony`n"
             $stoppedTasks = $true
             $failedTasks = $true
        }
    }
}


if ($stoppedServices)
{
    $message ="";
    
    if ($serviceArray.Length -gt 1)
    {
        $subject = "[DI $env Notice] Services have been stopped on $(hostname)"
        $message = "The following services on $(hostname) were successfully stopped:`n`n${serviceArray}`n"
        $successMessage = $true
    }
        else
    {
        $subject = "[DI $env Notice] $startedService has been stopped"
        $message = "The instance of $startedService on $(hostname) was successfully stopped."
        $successMessage = $true
    }
    if($stoppedTasks)
    {
        $message += "`nThe following tasks were set to disabled: `n`n${taskArray}`n"
        if($failedTasks)
        {
            $message += "`nThe following tasks could not be disabled: `n`n${taskFailArray}`n"
        }
    }
    if ($failedServices)
    {
        if ($failureArray -ne 0)
        {
            $subject = "[DI $env Notice] Service Notification for $(hostname)"
            $failMessage = "The following service(s) could not be stopped:`n`n${failureArray}"
            
            if ($successMessage)
            {
                $message = "$message`n`n$failMessage"
            }
            else
            {
                $message = $failMessage
            }
        }
    }

    sendmail -t $mailingList -s $subject -m $message
}
