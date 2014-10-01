# #############################################################################
# NAME: enVar.ps1
# 
# AUTHOR:  Gregg Jenczyk, UMass (UITS)
# DATE:  2014/07/09
# EMAIL: gjenczyk@umassp.edu
# 
# COMMENT:  This file contains config variables commonly used by other scripts 
#
# VERSION HISTORY
# 1.0 2014.07.09 Initial Version.
#
# TO ADD:
# 
# USEFUL SNIPPETS:
# "$(Get-Date) " | Out-File -Append $runLog
# #############################################################################


#-- CONFIG --#

$localRoot = "D:\"
$shareRoot = "Y:\"
$root = "\\ssisnas215c2.umasscs.net\diimages67prd\"
$env = ([environment]::MachineName).Substring(2)
$env = $env -replace "W.*",""
$logDate = $(get-date -format 'yyyyMMdd')


#- LOGGING -#
$runLog = "${root}log\running_log-${scriptName}.log"
$scriptLog = "${root}log\ScriptName_${logDate}.log"

