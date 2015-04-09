<# #############################################################################
# NAME: mergeTiffs.ps1
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
    [string[]] $path,
    [string] $name
    )

#-- INCLUDES --#
. "D:\inserver6\script\PowerShell\sendmail.ps1"

#-- CONFIG --#
$localRoot = "D:\"
$root = "D:\inserver6\"
$env = ([environment]::MachineName).Substring(2)
$env = $env -replace "W.*",""
$convServ = "DI${env}WEBRPT01"
$logDate = $(get-date -format 'yyyyMMdd')
$scriptName = "mergePDFs" #change this here - no extension
$returnCode = 0;

#-- WORKING PATHS --#
$webBaseDir = "D:\inserver6\output\complete\"
$webimgDir = "${webBaseDir}${name}\"
$gs = "\\DI${env}WEBRPT01\D$\\Program Files\gs\gs9.16\bin\gswin64c.exe"

#- LOGGING -#
$runLog = "${root}log\run_log-${scriptName}_${logDate}.log"

#-- MAIN --#
"$(get-date) - Starting ${scriptName} Script" | Out-File $runLog -Append
"WEBRPT SERVER = ${convServ}" | Out-File $runLog -Append
#Enter-PSSession -ComputerName $convServ -Authentication NegotiateWithImplicitCredential
$session = New-PSSession $convServ -Authentication NegotiateWithImplicitCredential
try {
    $inputString = ""
    ForEach ($item in $path)
    {
        $inputString += "`"${item}`" "
    }
    $mergedDoc = "${webimgDir}${name}.pdf"
    & $gs -dBATCH -dNOPAUSE -q -sDEVICE=pdfwrite -o ${mergedDoc} $inputString
    #cleanup
    ForEach ($item in $path)
    {
        Remove-Item -Path $item
    }
}
catch [system.exception]{
    $error[0] | Out-File $runLog -Append
    $returnCode = 1;
}
finally
{
    Remove-PSSession -Session $session
}

$error[0] | Out-File $runLog -Append
"$(get-date) - Finishing ${scriptName} Script`n Returned: ${returnCode}" | Out-File $runLog -Append
[Environment]::Exit($returnCode)
