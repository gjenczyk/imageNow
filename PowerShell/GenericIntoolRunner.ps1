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
$root = "\\boisnas215c1.umasscs.net\diimages67tst\"
$env = ([environment]::MachineName).Substring(2)
$env = $env -replace "W.*",""
$logDate = $(get-date -format 'yyyyMMdd')
$scriptName = "CompleteTaskAndMove" #change this here - no extension

#- LOGGING -#
$runLog = "${root}log\running_log-${scriptName}.log"
$scriptLog = "${root}log\${scriptName}_${logDate}.log"

#-- MAIN --#
"$(get-date) - Starting ${scriptName} Script" | Out-File $runLog -Append

D:\inserver6\bin64\intool --cmd run-iscript --file ${root}script\${scriptName}.js >> $runLog

<#
Use this if you want to be notified when the script finishes running

sendmail -t "gjenczyk@umassp.edu" -s "[DI ${env} Notice] ${scriptName}.ps1 has finished running" -m ${message}
#>

$error[0] | Format-List -Force | Out-File $runLog -Append
"$(get-date) - Finishing ${scriptName} Script" | Out-File $runLog -Append