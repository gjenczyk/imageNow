# #############################################################################
# NAME: GenericIntoolRunner.ps1
# 
# AUTHOR:  Gregg Jenczyk, UMass (UITS)
# DATE:  2014/04/18
# EMAIL: gjenczyk@umassp.edu
# 
# COMMENT:  This script is a template for running intool scripts via 
#           PowerShell.  You can use this on demand or to set up scheduled
#           jobs.  Just be sure to save this as something else.
#
# VERSION HISTORY
# 1.0 2014.04.18 Initial Version.
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
$root = "Z:\"
$env = ([environment]::MachineName).Substring(2)
$env = $env -replace "W.*",""
$logDate = $(get-date -format 'yyyyMMdd')

#- LOGGING -#
$runLog = "${root}log\running_log-scriptName.log"
$scriptLog = "${root}log\ScriptName_${logDate}.log"

#-- MAIN --#
"$(get-date) - Starting ScriptName Script" | Out-File $runLog -Append

D:\inserver6\bin64\intool --cmd run-iscript --file ${root}script\IMPORT_DSR_Archives.js

<#
Use this if you want to be notified when the script finishes running
#>
sendmail -t "gjenczyk@umassp.edu" -s "[DI ${env} Notice] IMPORT_runner.ps1 has finished running" -m ${message}


$error[0] | Out-File $runLog -Append
"$(get-date) - Finishing ScriptName Script" | Out-File $runLog -Append