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
# 1.3 2015.02.13 This is wicked ugly and needs a thorough cleaning.  re: import agent part
#
# TO ADD:
# 
# USEFUL SNIPPETS:
# "$(Get-Date) " | Out-File -Append $runLog
##############################################################################>

. "\\ssisnas215c2.umasscs.net\diimages67prd\script\PowerShell\sendmail.ps1"

$logPath = "\\ssisnas215c2.umasscs.net\diimages67prd\log\INowServiceMonitor.log"

$env = hostname
$hostnm = hostname
$env = $env.ToUpper().substring(2,5)
$serviceArray = @()
$failureArray = @()
$foreignServiceArray = @()
$foreignFailureArray = @()
$restartedServices = $false
$restartedForeignServices = $false
$failedServices = $false
$failedForeignServices
$successMessage = $false
$foreignSuccess = $false
$foreignComputer = @("DI${env}APPRAS01")
$serviceEmail = @("gjenczyk@umassp.edu", "cmatera@umassp.edu")

Get-Service | Where-Object {$_.Name -match "ImageNow Se*" -and $_.Status -eq "Stopped"} | ForEach-Object {
    
    Start-Service $_.Name
    if ($?)
    {
        $startedService = $_.Name
        "SERVER BLOCK 1: Started $startedService at $(Get-Date)" | Out-File -Append -FilePath $logPath
        $serviceArray += "$startedService`n"
        $restartedServices = $true
    }
    else
    {
        $error[0] | Format-List -Force | Out-File -Append -FilePath $logPath
        $failedStart = $_.Name
        "SERVER BLOCK 2: Failed to start $failedStart at $(Get-Date)" | Out-File -Append -FilePath $logPath
        $failureArray += "$failedStart"
        $restartedServices = $true
        $failedServices = $true
    } 
}

Get-Service | Where-Object {$_.Name -match "ImageNow*" -and $_.Status -eq "Stopped"} | ForEach-Object {

    if ($_.Name -notmatch "ImageNow Retention*" -and $_.Name -notmatch "ImageNow Import*" -and $_.Name -notmatch "ImageNow Monitor*") 
    {
        Start-Service $_.Name
        if ($?)
        {
            $startedService = $_.Name
            "SERVICE BLOCK 1: Started $startedService at $(Get-Date)" | Out-File -Append -FilePath $logPath
            $serviceArray += "$startedService`n"
            $restartedServices = $true
        }
        else
        {
            $error[0] | Format-List -Force | Out-File -Append -FilePath $logPath
            $failedStart = $_.Name
            "SERVICE BLOCK 2: Failed to start $failedStart at $(Get-Date)" | Out-File -Append -FilePath $logPath
            $failureArray += "$failedStart"
            $restartedServices = $true
            $failedServices = $true
        }
    } 

    if ($_.Name -match "ImageNow Import*")
    {
        $importService = $_.Name
        $appServers = @("DI${env}WEBIMG01", "DI${env}WEBIMG02")
        [System.Collections.ArrayList]$appArray = $appServers

        if ((Get-WmiObject -Class Win32_Service -Property StartMode,Name -Filter "Name LIKE '$importService'") | Where-Object {$_.StartMode -match "Manual"})
        {
            $appArray.Remove($hostnm)
            if (Get-Service -ComputerName $appArray | Where-Object {$_.Name -match "ImageNow Import*" -and $_.Status -match "Stopped"})
            {
               Start-Service -Name $importService
               if ($?)
                {
                    $startedService = $importService
                    "INPORT SERVICE 1: Started $startedService at $(Get-Date)" | Out-File -Append -FilePath $logPath
                    $serviceArray += "$startedService`n"
                    $restartedServices = $true
                }
                else
                {
                    $error[0] | Format-List -Force | Out-File -Append -FilePath $logPath
                    $failedStart = $importService
                    "INPORT SERVICE 2: Failed to start $failedStart at $(Get-Date)" | Out-File -Append -FilePath $logPath
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
                    $error[0] | Format-List -Force | Out-File -Append -FilePath $logPath
                    "INPORT SERVICE 3: Failed to stop $importService at $(Get-Date)" | Out-File -Append -FilePath $logPath
                    $failStopSub = "[DI $env Notice] Failed to stop $importService on $(hostname)"
                    $failStopMessage = "INowServiceMonitor has detected that the Import Agent service is running on two servers`nAn attempt to stop the second instance of the service was not successful.`nPlease stop the service on $(hostname) as soon as possible.`nFailure to stop the service may result in problems with imports."
                    sendmail -t $serviceEmail -s 
                }
            }

        }
        elseif ((Get-WmiObject -Class Win32_Service -Property StartMode,Name -Filter "Name LIKE '$importService'") | Where-Object {$_.StartMode -match "Auto"})
        {
            $appArray.Remove($hostnm)
            if ((Get-Service -ComputerName $appArray | Where-Object {$_.Name -match "ImageNow Import*" -and $_.Status -match "Running"}))
            {
                (Get-Service -ComputerName $appArray | Where-Object {$_.Name -match "ImageNow Import*" -and $_.Status -match "Running"}).Stop()
                if ($?)
                {
                    "Stopped $importService at $(Get-Date) because it was running on $appArray" | Out-File -Append -FilePath $logPath
                }
                else 
                {
                    $error[0] | Format-List -Force | Out-File -Append -FilePath $logPath
                    "INPORT SERVICE 4: Failed to stop $importService at $(Get-Date)" | Out-File -Append -FilePath $logPath
                    $failStopSub = "[DI $env Notice] Failed to stop $importService on $(hostname)"
                    $failStopMessage = "INowServiceMonitor has detected that the Import Agent service is running on two servers`nAn attempt to stop the second instance of the service was not successful.`nPlease stop the service on $(hostname) as soon as possible.`nFailure to stop the service may result in problems with imports."
                    sendmail -t $serviceEmail -s $failStopSub -m $failStopMessage
                }

                Start-Service $importService
                if ($?)
                {
                    $startedService = $importService
                    "IMPORT SERVICE 5: Started $startedService at $(Get-Date)" | Out-File -Append -FilePath $logPath
                    $serviceArray += "$startedService`n"
                    $restartedServices = $true
                }
                else
                {
                    $error[0] | Format-List -Force | Out-File -Append -FilePath $logPath
                    $failedStart = $importService
                    "INPORT SERVICE 6: Failed to start $failedStart at $(Get-Date)" | Out-File -Append -FilePath $logPath
                    $failureArray += "$failedStart"
                    $restartedServices = $true
                    $failedServices = $true
                }

            }
            else
            {
            
                Start-Service $importService
                if ($?)
                {
                    $startedService = $importService
                    "IMPORT SERVICE 7: Started $startedService at $(Get-Date)" | Out-File -Append -FilePath $logPath
                    $serviceArray += "$startedService`n"
                    $restartedServices = $true
                }
                else
                {
                    $error[0] | Format-List -Force | Out-File -Append -FilePath $logPath
                    $failedStart = $importService
                    "INPORT SERVICE 8: Failed to start $failedStart at $(Get-Date)" | Out-File -Append -FilePath $logPath
                    $failureArray += "$failedStart"
                    $restartedServices = $true
                    $failedServices = $true
                }
            }
        }
        else
        {
            Start-Service $importService
            if ($?)
            {
                $startedService = $importService
                "IMPORT SERVICE 9: Started $startedService at $(Get-Date)" | Out-File -Append -FilePath $logPath
                $serviceArray += "$startedService`n"
                $restartedServices = $true
            }
            else
            {
                $error[0] | Format-List -Force | Out-File -Append -FilePath $logPath
                $failedStart = $importService
                "INPORT SERVICE 10: Failed to start $failedStart at $(Get-Date)" | Out-File -Append -FilePath $logPath
                $failureArray += "$failedStart"
                $restartedServices = $true
                $failedServices = $true
            }
        }

    }
}

foreach ($computer in $foreignComputer)
{

    Get-Service -ComputerName $computer | Where-Object {$_.Name -match 'Image*' -and $_.Status -match 'Stopped*'} | ForEach-Object {

        $foreignService = $_.Name
        if(Get-WmiObject -Class Win32_Service -ComputerName $computer -Property StartMode,Name -Filter "Name LIKE '$foreignService' AND StartMode LIKE 'Auto'")
        {
            Set-Service -ComputerName $computer $foreignService -Status Running
    
            if ($?)
            {
                "FOREIGN SERVICE 1: Started $foreignService at $(Get-Date)" | Out-File -Append -FilePath $logPath
                $foreignServiceArray += "$foreignService`n"
                $restartedForeignServices = $true
                $foreignSuccses = $true
            }
            else
            {
                $error[0] | Format-List -Force | Out-File -Append -FilePath $logPath
                "FOREIGN SERVICE 2: Failed to start $foreignService at $(Get-Date)" | Out-File -Append -FilePath $logPath
                $foreignFailureArray += "$foreignService`n"
                $failedForeignServices = $true
                $restartedForeignServices = $true
            }
        }
    }

    if ($restartedForeignServices)
    {
        $foreignSub = "[DI $env Notice] Services have been restarted on $computer"
        $foreignMessage = "The following services on $computer were stopped:`n`n${foreignServiceArray}`nThey were restarted at $(Get-Date)"
        if ($failedForeignServices)
        {
            $foreignSub = "[DI $env Notice] Service Notification for $computer"
            $foreignFailMessage = "The following service(s) could not be started on ${foreignServiceArray}:`n`n${foreignFailureArray}"
            if ($foreignSuccess)
            {
                $foreignMessage += "`n$foreignFailMessage"
            }
            else
            {
                $foreignMessage = $foreignFailMessage
            }
        }
      #sendmail -t $serviceEmail -s $foreignSub -m $foreignMessage
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
            $failMessage = "The following service(s) could not be started on $(hostname):`n`n${failureArray}"
            
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

    sendmail -t $serviceEmail -s $subject -m $message
}