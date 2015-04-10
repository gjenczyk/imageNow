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
    [string] $emplid,
    [string] $docId,
    [string] $docType,
    [string] $seqNum)

#-- INCLUDES --#
. "\\boisnas215c1.umasscs.net\diimages67tst\script\PowerShell\sendmail.ps1"

#-- CONFIG --#
$localRoot = "D:\"
$root = "D:\inserver6\"
$env = ([environment]::MachineName).Substring(2)
$env = $env -replace "W.*",""
$convServ = "DI${env}WEBRPT01"
$logDate = $(get-date -format 'yyyyMMdd')
$scriptName = "mergeTiff" #change this here - no extension
$returnCode = 0;

#-- WORKING PATHS --#
$webBaseDir = "D:\inserver6\output\"
$webimgDir = "${webBaseDir}${emplid}\${docId}\"
$completeDir = "${webBaseDir}complete\${emplid}"
$expDir = "\\DI${env}WEBRPT01\D$\INMAC\tiffExport\${emplid}\"
$mergeDir = "\\DI${env}WEBRPT01\D$\INMAC\tiffExport\${emplid}\${docId}\"
$tiffcp = "\\DI${env}WEBRPT01\D$\Program Files (x86)\GnuWin32\bin\tiffcp.exe"
$tiff2pdf = "\\DI${env}WEBRPT01\D$\Program Files (x86)\GnuWin32\bin\tiff2pdf.exe"

#- LOGGING -#
$runLog = "${root}log\run_log-${scriptName}_${logDate}.log"

#-- MAIN --#
"$(get-date) - Starting ${scriptName} Script" | Out-File $runLog -Append
"WEBRPT SERVER = ${convServ}" | Out-File $runLog -Append
#Enter-PSSession -ComputerName $convServ -Authentication NegotiateWithImplicitCredential
$session = New-PSSession $convServ -Authentication NegotiateWithImplicitCredential
try {
    if(-Not $(Test-Path ${mergeDir}))
    {
        New-Item ${mergeDir} -ItemType directory
    }
    Get-ChildItem ${webimgDir}* -File | ForEach-Object {
        #"INCOMING TIFF: $($_.FullName)" | Out-File $runLog -Append
        Copy-Item $_ -Destination ${mergeDir}
        $error[0] | Out-File $runLog -Append
    }

    $mergeFile = ""
    $pdfOut = ""
    Get-ChildItem ${mergeDir}* -File | ForEach-Object {
        $baseName = $_.BaseName
        $extension = $_.Extension
        $renArr = [regex]::split($baseName,'_') 
        $uniqueId = "$($renArr[0])_$($renArr[1])"
        $pageNo = $renArr[2]
        #handle single page exports
        if($pageNo -eq $null)
        {
            $pageNo = "1"
        }
        $mergeFile = "${mergeDir}${uniqueId}${extension}"
        #"MERGE FILE: ${mergeFile}" | Out-File $runLog -Append
        $pdfOut = "${mergeDir}$($renArr[0])_$($renArr[1]).PDF"
        #"PDF: $pdfOut" | Out-File $runLog -Append
        Rename-Item $_ -NewName "${pageNo}_${uniqueId}${extension}"
    }

    $tiffArr = @()
    Get-ChildItem ${mergeDir}* | ForEach-Object {
        $tiffArr += $_.Name
    }

    $arrSize = $tiffArr.Length
    $sortedArr = @()
    $highestVal = 0
    while ($sortedArr.Length -lt $arrSize)
    {
        for ($i = 0; $i -lt $tiffArr.Length; $i++)
        {
            $curValue = [regex]::Split($tiffArr[$i],'_')
            $thisOne = [int]$curValue[0]
            if ($thisOne -eq ($highestVal + 1))
            {
                $sortedArr += $tiffArr[$i]
                $highestVal += 1
            }
        }
    }

    $sourceFiles = ""
    for ($j = 0; $j -lt $sortedArr.Length; $j++)
    {
       $sourceFiles += "`"${mergeDir}$($sortedArr[$j])`" "
    }

    & $tiffcp -c none $sourceFiles $mergeFile
    for ($k = 0; $k -lt $sortedArr.Length; $k++)
    {
       Remove-Item -Path "${mergeDir}$($sortedArr[$k])"
    }

    & $tiff2pdf $mergeFile -o $pdfOut
    if(-Not $(Test-Path ${completeDir}))
    {
        New-Item -Path ${completeDir} -ItemType directory
    }
    Move-Item -Path $pdfOut -Destination "${completeDir}\${docType}_${seqNum}.pdf"
    Remove-Item -Path ${expDir} -Recurse
}
catch [system.exception]{
    $error[0] | Out-File $runLog -Append
    $returnCode = 1;
}
finally
{
    Remove-PSSession -Session $session
}
#cleanup
Get-ChildItem $webimgDir -File | ForEach-Object {
     Remove-Item $_.FullName
}
Remove-Item "${webBaseDir}${emplid}\" -Recurse

$error[0] | Out-File $runLog -Append
"$(get-date) - Finishing ${scriptName} Script`n Returned: ${returnCode}" | Out-File $runLog -Append
[Environment]::Exit($returnCode)
