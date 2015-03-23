##############################################################################
# NAME: CommonAppSearch.ps1
# 
# AUTHOR:  Gregg Jenczyk, UMass (UITS)
# DATE:  2014/04/11
# EMAIL: gjenczyk@umassp.edu
# 
# COMMENT:  This script unzips files delivered from Common App, checks the pdfs
#           contained within for errors, renames the valid pdfs according to DI 
#           standards, and moves them to the imports directory.
#
# VERSION HISTORY
# 1.0 2014.04.11 Initial Version.
#
# TO ADD:
# 
# USEFUL SNIPPETS:
# "$(Get-Date) " | Out-File -Append $runLog
##############################################################################

#-- INCLUDES --#
. "\\ssisnas215c2.umasscs.net\diimages67prd\script\PowerShell\sendmail.ps1"

#-- CONFIG --#

$root = "D:\"
$env = ([environment]::MachineName).Substring(2)
$env = $env -replace "W.*",""
$logDate = $(get-date -format 'yyyyMMdd')

#- LOGGING -#
$runLog = "${root}inserver6\log\run_log-CommonAppSearch.log"
$scriptLog = "${root}inserver6\log\CommonAppSearch_${logDate}.log"

#-- MAIN --#
"$(get-date) - Starting CommonAppSearch Script" | Out-File $runLog -Append
$sw = [System.Diagnostics.Stopwatch]::StartNew()

D:\inserver6\bin64\intool --cmd run-iscript --file \\ssisnas215c2.umasscs.net\diimages67prd\script\CommonAppSearch.js >> $runLog

$sw.Stop()
$totalRunTime = $sw.ElapsedMilliseconds/1000

#
if ($totalRunTime -lt 2){
    if($(Get-Content $scriptLog | where { $_ -match "First row of cursor not found:"} | Measure-Object).count -eq 3){
    $message="There are currently no documents in the Common App Docs queues."
    } else {
    $message="The matching script finished in $totalRunTime seconds - that seems kind of quick.`nThere may have been an error logging in with intool.`nPlease verify that the script actually ran."
    }
} else {
 $numberMatch = $(Get-Content $scriptLog | where { $_ -match "Successfully reindexed"} | Measure-Object).count
 $docLines = $(Get-Content $scriptLog | where { $_ -match "document id:"} | Measure-Object).count
 #$totalDocs = $($docLines - 3)
 $message="The script finished running at $(get-date -Format 'h:mm \o\n dddd, MMMM dd').
 It took $('{0:N2}' -f $($totalRunTime/60)) minutes to make ${numberMatch} matches out of ${docLines} documents across all Common App queues.
 Average search time: $('{0:N2}' -f $($totalRunTime/($docLines))) seconds per document."
}

sendmail -t UITS.DI.CORE@umassp.edu -s "[DI ${env} Notice] CommonAppSearch.ps1 has finished running" -m ${message}

Start-Sleep -s 2
Get-Item -path $scriptLog | Rename-Item -NewName {$_.Name -replace ".log","-$(Get-Date -format 'HHmm').log"}

$error[0] | Out-File $runLog -Append
"$(get-date) - Finishing CommonAppSearch Script" | Out-File $runLog -Append