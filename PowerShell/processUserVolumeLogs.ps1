<# #############################################################################
# NAME: processUserVolumeLogs.ps1
# 
# AUTHOR:  Gregg Jenczyk, UMass (UITS)
# DATE:  2014/12/26
# EMAIL: gjenczyk@umassp.edu
# 
# COMMENT:  This script packages up and emails the userVolumeLoggger output
#           into a pivot table and ...
#
# VERSION HISTORY
# 1.0 2014.12.26 Initial Version.
#
# TO ADD:
# 
# USEFUL SNIPPETS:
# "$(Get-Date) " | Out-File -Append $runLog
# #############################################################################>

Param([string]$passedDate)

#-- INCLUDES --#
. "\\boisnas215c1.umasscs.net\diimages67tst\script\PowerShell\sendmail.ps1"
. "\\boisnas215c1.umasscs.net\diimages67tst\script\PowerShell\enVar.ps1"

$passedDate = 1262015


#-- CONFIG --#

$csvDir = "${root}script\log\userVolumeLogger"

#- LOGGING -#
$runLog = "${root}log\running_log-processUserVolumeLogs.log"
$scriptLog = "${root}log\processUserVolumeLogs_${logDate}.log"

#-- MAIN --#

$a = "<style>"
$a = $a + "TABLE{border-width: 1px;border-style: solid;border-color: black;border-collapse: collapse;}"
$a = $a + "TH{border-width: 1px;padding: 5px;border-style: solid;border-color: black;background-color: gainsboro;}"
$a = $a + "TD{border-width: 1px;padding: 5px;border-style: solid;border-color: black;text-align: center;}"
$a = $a + "</style>"

Import-Csv ${csvDir}_R_${passedDate}.csv -Header "Hour","Total","Boston","Dartmouth","Lowell","UITS","Other" | ConvertTo-HTML -head $a | Out-File -FilePath ${csvDir}_F_${passedDate}.csv

$body = (Get-Content ${csvDir}_F_${passedDate}.csv)

$attachment = @("${csvDir}_R_${passedDate}.csv")
if (Test-Path "${csvDir}_D_${passedDate}.csv")
{
    $attachment += "${csvDir}_D_${passedDate}.csv"
}

sendmail -s "[DI $env Notice] User Report for $(Get-Date -Format MM/dd/yyyy)" -a $attachment -m $body -to gjenczyk@umassp.edu -flag "BodyAsHtml"

Remove-Item -Path "${csvDir}_D_${passedDate}.csv" -ErrorAction SilentlyContinue
Remove-Item -Path "${csvDir}_R_${passedDate}.csv"