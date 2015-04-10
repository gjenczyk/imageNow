<# #############################################################################
# NAME: emailPDFs.ps1
# 
# AUTHOR:  Gregg Jenczyk, UMass (UITS)
# DATE:  2015/04/06
# EMAIL: gjenczyk@umassp.edu
# 
# COMMENT:  This script sucks up tiffs put out by an iScript, merges them,
            converts them to PDF and passes them back to 

    INSTRUCTIONS: NEED TO TUNR ON WINRM AND ADD HOSTA TO HOSTB's TRUSTED HOSTS
                  ON EACH SERVER!!!
                  
#
# VERSION HISTORY
# 1.0 2015.04.06 Initial Version.
#
# TO ADD:
# 
# USEFUL SNIPPETS:
# "$(Get-Date) " | Out-File -Append $runLog
# #############################################################################>

Param(
    [string] $address,
    [string[]] $attachments,
    [string] $emplid,
    [string] $name
    )

#-- INCLUDES --#
. "\\boisnas215c1.umasscs.net\diimages67tst\script\PowerShell\sendmail.ps1"

#-- CONFIG --#
$localRoot = "D:\"
$root = "D:\inserver6\"
$env = ([environment]::MachineName).Substring(2)
$env = $env -replace "W.*",""
$convServ = "DI${env}WEBRPT01"
$logDate = $(get-date -format 'yyyyMMdd')
$scriptName = "emailPDFs" #change this here - no extension
$returnCode = 0;

#- LOGGING -#
$runLog = "${root}log\run_log-${scriptName}_${logDate}.log"

#-- MAIN --#
"$(get-date) - Starting ${scriptName} Script" | Out-File $runLog -Append
try
{
    $subject = "[DI ${env} NOTICE] PDF EXPORT FOR ${emplid}"
    $message = "Attached is the file generated for ${name}"
    $message += "`n`n`n`n`n`n`n`n"
    $message += "------------------------`nThis is an automated message.`nPLEASE DO NOT REPLY TO THIS MESSAGE.`nQuestions can be sent to UITS.DI.CORE@umassp.edu"
    sendmail -t $address -s $subject -m $message -a $attachments #-flag BodyAsHtml
    
}
catch [system.exception]{
    $error[0] | Out-File $runLog -Append
    $returnCode = 1;
}
finally
{
    $error[0] | Out-File $runLog -Append
    "$(get-date) - Finishing ${scriptName} Script`n Returned: ${returnCode}" | Out-File $runLog -Append
    [Environment]::Exit($returnCode)
}


