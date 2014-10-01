<##############################################################################
# NAME: INowServiceMonitor.ps1
# 
# AUTHOR:  Gregg Jenczyk, UMass (UITS)
# DATE:  2014/05/05
# EMAIL: gjenczyk@umassp.edu
# 
# COMMENT:  This script checks the status of the ImageNow related services
#           and starts them if they aren't currently running.
#
# VERSION HISTORY
# 1.0 2014.05.05 Initial Version.
# 1.1 2014.05.08 Finished handling for situations where services fail to start.
# 1.2 2014.10.01 Added a section to start services on a foreign computer
#
# TO ADD:
# 
# USEFUL SNIPPETS:
# "$(Get-Date) " | Out-File -Append $runLog
##############################################################################>

. "\\ssisnas215c2.umasscs.net\diimages67prd\script\PowerShell\sendmail.ps1"

$logPath = "\\ssisnas215c2.umasscs.net\diimages67prd\log\INowServiceMonitor.log"

$hostnm = hostname
$env = $hostnm.ToUpper().substring(2,5)
$serviceArray = @()
$failureArray = @()

$env = hostname
$env = $env.ToUpper().substring(2,5)
$serviceArray = @()
$failureArray = @()
$restartedServices = $false
$failedServices = $false
$successMessage = $false
$foreignComputer = @("DIPRD67APPRAS01")

Get-Service | Where-Object {$_.Name -match "ImageNow Se*" -and $_.Status -eq "Stopped"} | ForEach-Object {
    
    Start-Service $_.Name
    if ($?)
    {
        $startedService = $_.Name
        "Started $startedService at $(Get-Date)" | Out-File -Append -FilePath $logPath
        $serviceArray += "$startedService`n"
        $restartedServices = $true
    }
    else
    {
        $failedStart = $_.Name
        "Failed to start $failedStart at $(Get-Date)" | Out-File -Append -FilePath $logPath
        $failureArray += "$failedStart"
        $restartedServices = $true
        $failedServices = $true
    } 
}

Get-Service | Where-Object {$_.Name -match "ImageNow*" -and $_.Status -eq "Stopped"} | ForEach-Object {

    if ($_.Name -notmatch "ImageNow Retention*" -and $_.Name -notmatch "ImageNow Import*") 
    {
        Start-Service $_.Name
        if ($?)
        {
            $startedService = $_.Name
            "Started $startedService at $(Get-Date)" | Out-File -Append -FilePath $logPath
            $serviceArray += "$startedService`n"
            $restartedServices = $true
        }
        else
        {
            $failedStart = $_.Name
            "Failed to start $failedStart at $(Get-Date)" | Out-File -Append -FilePath $logPath
            $failureArray += "$failedStart"
            $restartedServices = $true
            $failedServices = $true
        }
    } 

    if ($_.Name -match "ImageNow Import*")# -and $_.Status -eq 'Stopped')
    {
        $importService = $_.Name
        $appServers = @("DIPRD67WEBIMG01", "DIPRD67WEBIMG02")
        [System.Collections.ArrayList]$appArray = $appServers

        if ((Get-WmiObject -Class Win32_Service -Property StartMode -Filter "Name='$importService'") | Where-Object {$_.StartMode -match "Manual"})
        {
            $appArray.Remove($hostnm)

            if (Get-Service -ComputerName $appArray | Where-Object {$_.Name -match "ImageNow Import*" -and $_.Status -match "Stopped"})
            {
               Start-Service -Name $importService
               if ($?)
                {
                    $startedService = $importService
                    "Started $startedService at $(Get-Date)" | Out-File -Append -FilePath $logPath
                    $serviceArray += "$startedService`n"
                    $restartedServices = $true
                }
                else
                {
                    $failedStart = $importService
                    "Failed to start $failedStart at $(Get-Date)" | Out-File -Append -FilePath $logPath
                    $failureArray += "$failedStart"
                    $restartedServices = $true
                    $failedServices = $true
                }
            }
            elseif ((Get-Service -ComputerName $appArray | Where-Object {$_.Name -match "ImageNow Import*" -and $_.Status -match "Running"}) -and (Get-Service | Where-Object {$_.Name -match "ImageNow Import*" -and $_.Status -match "Running"}))
            {
                Stop-Service -Name $importService -Force
                if ($?)
                {
                    "Stopped $importService at $(Get-Date) because it was running on $appArray" | Out-File -Append -FilePath $logPath
                }
                else 
                {
                    "Failed to stop $importService at $(Get-Date)" | Out-File -Append -FilePath $logPath
                    $failStopSub = "[DI $env Notice] Failed to stop $importService on $(hostname)"
                    $failStopMessage = "INowServiceMonitor has detected that the Import Agent service is running on two servers`nAn attempt to stop the second instance of the service was not successful.`nPlease stop the service on $(hostname) as soon as possible.`nFailure to stop the service may result in problems with imports."
                    sendmail -t gjenczyk@umassp.edu, cmatera@umassp.edu -s 
                }
            }

        }
        else
        {
            
            Start-Service $importService
            if ($?)
            {
                $startedService = $importService
                "Started $startedService at $(Get-Date)" | Out-File -Append -FilePath $logPath
                $serviceArray += "$startedService`n"
                $restartedServices = $true
            }
            else
            {
                $failedStart = $importService
                "Failed to start $failedStart at $(Get-Date)" | Out-File -Append -FilePath $logPath
                $failureArray += "$failedStart"
                $restartedServices = $true
                $failedServices = $true
            }
        }

    }
}

foreach ($computer in $foreignComputer)
{
    "$computer"

    Get-Service -ComputerName $computer | Where-Object {$_.Name -match 'Image*' -and $_.Status -match 'Stopped*'} | ForEach-Object {

        $foreignService = $_.Name
    
        Set-Service -ComputerName $computer $_.Name -Status Running
    
        if ($?)
        {
            $startedService = $foreignService
            "Started $startedService at $(Get-Date)" | Out-File -Append -FilePath $logPath
            $serviceArray += "$startedService`n"
            $restartedServices = $true
        }
        else
        {
            $failedStart = $foreignService
            "Failed to start $failedStart at $(Get-Date)" | Out-File -Append -FilePath $logPath
            $failureArray += "$failedStart"
            $restartedServices = $true
            $failedServices = $true
        }
    }
}

if ($restartedServices)
{
    if ($serviceArray.Length -gt 1)
    {
        $subject = "[DI $env Notice] Services have been restarted on $(hostname)"
        $message = "The following services on $(hostname) were stopped:`n`n${serviceArray}`nThey were restarted at $(Get-Date)"
        $successMessage = $true
    }
    else
    {
        $subject = "[DI $env Notice] $startedService has been restarted"
        $message = "The instance of $startedService on $(hostname) was stopped.`nIt has been restarted at $(Get-Date)"
        $successMessage = $true
    }

    if ($failedServices)
    {
        if ($failureArray -ne 0)
        {
            $subject = "[DI $env Notice] Service Notification for $(hostname)"
            $failMessage = "The following service(s) could not be started:`n`n${failureArray}"
            
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

    sendmail -t gjenczyk@umassp.edu, cmatera@umassp.edu, lprudden@umassp.edu -s $subject -m $message
}