<##############################################################################
# NAME: INowTaskMonitor.ps1
# 
# AUTHOR:  Gregg Jenczyk, UMass (UITS)
# DATE:  2014/06/18
# EMAIL: gjenczyk@umassp.edu
# 
# COMMENT:  This script checks the status of the ImageNow related tasks and
#           starts them if they aren't currently running and they should be.
#
# VERSION HISTORY
# 1.0 2014.06.18 Initial Version.
#
# TO ADD:
# 
# USEFUL SNIPPETS:
# "$(Get-Date) " | Out-File -Append $runLog
##############################################################################>

. "\\boisnas215c1.umasscs.net\diimages67tst\script\PowerShell\sendmail.ps1"

$logPath = "\\boisnas215c1.umasscs.net\diimages67tst\log\INowTaskMonitor.log"

$hostnm = hostname
$env = $hostnm.ToUpper().substring(2,5)


$appServers  = @("DITST67WEBIMG01","DITST67WEBIMG02")
[System.Collections.ArrayList]$appArray = $appServers
$thisServer = $hostnm
$appArray.Remove($thisServer)
$thatServer = $appArray

echo $thisServer
echo $thatServer

\\boisnas215c1.umasscs.net\diimages67tst\script\PowerShell\Get-ScheduledTask.ps1 | ForEach-Object {

    if ($_.State -eq "Disabled") {
        \\boisnas215c1.umasscs.net\diimages67tst\script\PowerShell\Get-ScheduledTask.ps1 -ComputerName $thatServer -TaskName $_.TaskName
    }

}
