<##############################################################################
# NAME: timeCop.ps1
# 
# AUTHOR:  Gregg Jenczyk, UMass (UITS)
# DATE:  2015/06/17
# EMAIL: gjenczyk@umassp.edu
# 
# COMMENT:  This script monitors the time on a collection of servers, and logs 
#           the difference. Maybe it will send an email, too.
#          
#
# VERSION HISTORY
# 1.0 2015.06.17 Initial Version.
#
# TO ADD:
# 
# USEFUL SNIPPETS:
# "$(Get-Date) " | Out-File -Append $runLog
##############################################################################>

## -- INCLUDES -- ##

. "\\ssisnas215c2.umasscs.net\diimages67prd\script\PowerShell\sendmail.ps1"

## --- CONFIG ---- ##

$env = ([environment]::MachineName).Substring(2)
$env = $env -replace "W.*",""
$thisBox = [environment]::MachineName
$logDate = $(get-date -format 'yyyyMMdd')
$timeLog = "\\ssisnas215c2.umasscs.net\diimages67prd\log\timeCop_${logDate}.log"
$logHeader = "----------------------------------------------------------------------------------"

$serverList = @("DI${env}WEBIMG01", "DI${env}WEBIMG02", "DI${env}WEBRPT01", "DI${env}WEB01", "DI${env}WEB02", "DI${env}APPRAS01", "DI${env}APPDCS01")
$domainControllers = @("COPRDDC01") #, "COPRDDC02")
$timeHash = $null
$timeHash = @{}
$timeStore = ""

## --- FUNCTIONS ---##

# -- MAIN -- #

try
{
    $logHeader | Out-File -Append $timeLog

    $dcCatch = $null
    foreach ($dc in $domainControllers)
    {
        $dcNetTime = & net time \\$dc
        "$dcNetTime" | Out-File -Append $timeLog
        #$dcTime = $dcNetTime -split '([0-9]\/[0-9].\/[0-9].+)'
        $dcTime = ([regex]::matches($dcNetTime, "([0-9]\/[0-9].\/[0-9].+M)") | %{$_.value})
        $dcTime = $(Get-Date $dcTime)
        "Domain Controller: ${dc} Time: $dcTime" | Out-File -Append $timeLog
        if ($dcCatch -eq $null)
        {
            $dcCatch = $dcTime
            $dcBench = $dc
        }
        else
        {
            $dcDiff = $dcCatch - $dcTime
        }
    }

    "Total difference in seconds: $([math]::abs($dcDiff.TotalSeconds))" | Out-File -Append $timeLog

    foreach ($server in $serverList)
    {
        $timeStore = Get-WmiObject Win32_OperatingSystem -ComputerName $server
        $timeString = $timeStore.LocalDateTime.Substring(0, $timeStore.LocalDateTime.length-4)
        $y = $timeString.Substring(0,4)
        $m = $timeString.Substring(4,2)
        $d = $timeString.Substring(6,2)
        $h = $timeString.Substring(8,2)
        $mi = $timeString.Substring(10,2)
        $s = $timeString.Substring(12,2)
        $ms = $timeString.Substring(15,6)
        $timeToUse = [datetime] "${m}/${d}/${y} ${h}:${mi}:${s}.${ms}"
        $timeHash.Add($server, $timeToUse) #[DateTime]$timeStore.LocalDateTime.Substring(0, $timeStore.LocalDateTime.length-4))
    }

    $previous = $null 
    $current = $null
    $curTime = $null
    $prevTime = $null
    foreach ($item in $timeHash.GetEnumerator() | Sort -Property Name)
    {
        $difference = $dcCatch - $item.Value
        #caluclate difference between sever and dc
        "Server: $($item.Name) Time: $($item.Value) Difference from ${dcBench}: $([math]::abs($difference.TotalSeconds))" | Out-File -Append $timeLog

        if([math]::abs($difference.TotalSeconds) -gt 10)
        {
            if([math]::abs($difference.TotalSeconds) -gt 60)
            {
                "CRITICAL: Time difference is $([math]::abs($difference.TotalSeconds))" | Out-File -Append $timeLog
            }
            else
            {
                "WARNING: Time difference is $([math]::abs($difference.TotalSeconds))" | Out-File -Append $timeLog
            } 
        }

        $current = $item.Name
        $curTime = $item.Value
        #"CURRENT SERVER: $currentServer"
        #"PREVIOUS SERVER: $previousServer"
        if ($previous -ne $null)
        {
            if($current.Substring(0, $current.Length-2) -eq $previous.Substring(0, $previous.Length-2))
            {
                $famDif = $curTime - $prevTime
                "------ Total difference in seconds between ${current} and ${previous}: $([math]::abs($famDif.TotalSeconds))" | Out-File -Append $timeLog
                if([math]::abs($famDif.TotalSeconds) -gt 10)
                {
                    if([math]::abs($famDif.TotalSeconds) -gt 60)
                    {
                        "----- CRITICAL: Time difference is $([math]::abs($famDif.TotalSeconds))" | Out-File -Append $timeLog
                    }
                    else
                    {
                        "----- WARNING: Time difference is $([math]::abs($famDif.TotalSeconds))" | Out-File -Append $timeLog
                    } 
                }
            }
        }
        
        $previous = $current
        $prevTime = $item.Value
    }
    
}
catch [system.exception]{
    $error[0] | Out-File -Append $timeLog
}
finally
{
    "done"
}
 #$timeHash