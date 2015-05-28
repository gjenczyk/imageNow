<# #############################################################################
# NAME: mergePDFs.ps1
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
    [string] $name,
    [string] $appno,
    [string] $outFile,
    [string] $protect
    )

#-- INCLUDES --#
. "\\ssisnas215c2.umasscs.net\diimages67prd\script\PowerShell\sendmail.ps1"

#-- CONFIG --#
$localRoot = "D:\"
$root = "D:\inserver6\"
$env = ([environment]::MachineName).Substring(2)
$env = $env -replace "W.*",""
$convServ = "DI${env}WEBRPT01"
$logDate = $(get-date -format 'yyyyMMdd')
$scriptName = "DocumentExport_files" #change this here - no extension
$returnCode = 0;

#-- WORKING PATHS --#
$webBaseDir = "D:\inserver6\output\complete\"
$webimgDir = "${webBaseDir}${name}_${appno}\"
$gs = "\\DI${env}WEBRPT01\D$\\Program Files\gs\gs9.14\bin\gswin64c.exe"
$IMC = "D:\ImageMagick\convert.exe"

#- LOGGING -#
$runLog = "${root}log\run_log-${scriptName}_${logDate}.log"

#-- MAIN --#
"$(get-date) - Starting ${scriptName} Script" | Out-File $runLog -Append
# determine if file is protected
$opts = ""
if($protect -eq "true")
{
    $opts = "-dPermissions=-4096"
    #manuallly: -4096
}
"WEBRPT SERVER = ${convServ}" | Out-File $runLog -Append
#Enter-PSSession -ComputerName $convServ -Authentication NegotiateWithImplicitCredential
$session = New-PSSession $convServ -Authentication NegotiateWithImplicitCredential
try {
    if($path.Length -lt 2)
    {
        "No need to merge: ${path}" | Out-File $runLog -Append
    }
    else
    {
        $inputString = ""
        ForEach ($item in $path)
        {
            $inputString += "`"${item}`" "
        }
        $mergedDoc = "${webimgDir}${name}_${appno}.${outFile}"
        if ($outFile -eq "pdf")
        {
            & $gs -dBATCH -dNOPAUSE -q -sDEVICE=pdfwrite <#-dDownsampleColorImages=true -dDownsampleGrayImages=true -dDownsampleMonoImages=true -dGrayImageResolution=200#> -sOwnerPassword=gregg -sUserPassword= -dEncryptionR=3 ${opts} -o ${mergedDoc} $inputString >> ${runLog}
            "$gs -dBATCH -dNOPAUSE -q -sDEVICE=pdfwrite <#-dDownsampleColorImages=true -dDownsampleGrayImages=true -dDownsampleMonoImages=true -dGrayImageResolution=200#> -sOwnerPassword=gregg -sUserPassword= -dEncryptionR=3 ${opts} -o ${mergedDoc} $inputString" | Out-File $runLog -Append
        }
        elseif ($outFile -eq "tif")
        {
            & $IMC $inputString -compress lzw -type bilevel -adjoin ${mergedDoc} >> ${runLog}
            "$IMC $inputString -compress lzw -adjoin ${mergedDoc}" | Out-File $runLog -Append
        }
        else
        {
            "Invalid format ${outFile}" | Out-File $runLog -Append       
        }
        #cleanup
        ForEach ($item in $path)
        {
            Remove-Item -Path $item
        }
    }
}
catch [system.exception]{
    $error[0] | Format-List -Force | Out-File $runLog -Append
    $returnCode = 1;
}
finally
{
    Remove-PSSession -Session $session
    $error[0] | Format-List -Force | Out-File $runLog -Append
    "$(get-date) - Finishing ${scriptName} Script`n Returned: ${returnCode}" | Out-File $runLog -Append
    [Environment]::Exit($returnCode)
}