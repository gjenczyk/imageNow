# #############################################################################
# NAME: DI_ResetDocsToIdle.ps1
# 
# AUTHOR:  Gregg Jenczyk, UMass (UITS)
# DATE:  2014/06/18
# EMAIL: gjenczyk@umassp.edu
# 
# COMMENT:  This script will reset all docs in DI in a working state to
#           idle.  If a user is logged in to the queue the script is
#           checking, the script will not process docs in that qeueu
#
# VERSION HISTORY
# 1.0 2014.06.18 Initial Version.
#
# TO ADD:
# 
# USEFUL SNIPPETS:
# "$(Get-Date) " | Out-File -Append $runLog
# #############################################################################

#-- INCLUDES --#
. "\\ssisnas215c2.umasscs.net\diimages67prd\script\PowerShell\sendmail.ps1"

#-- CONFIG --#

$localRoot = "D:\"
$root = "\\ssisnas215c2.umasscs.net\diimages67prd\"
$env = ([environment]::MachineName).Substring(2)
$env = $env -replace "W.*",""
$logDate = $(get-date -format 'yyyyMMdd')

#- LOGGING -#
$runLog = "${root}log\run_log-DI_ResetDocsToIdle_${logDate}.log"
#$scriptLog = "${root}log\ScriptName_${logDate}.log"

#-- MAIN --#
"$(get-date) - Starting DI_ResetDocsToIdle Script" | Out-File $runLog -Append

#D:\inserver6\bin64\intool --cmd reset-item-status

<#
Use this if you want to be notified when the script finishes running

sendmail -t "gjenczyk@umassp.edu" -s "[DI ${env} Notice] ScriptName.ps1 has finished running" -m ${message}
#>

$error[0] | Format-List -Force | Out-File $runLog -Append
"$(get-date) - Finishing DI_ResetDocsToIdle Script" | Out-File $runLog -Append