<# #############################################################################
# NAME: GenericIntoolRunner.ps1
# 
# AUTHOR:  Gregg Jenczyk, UMass (UITS)
# DATE:  2014/04/18
# EMAIL: gjenczyk@umassp.edu
# 
# COMMENT:  This script is a template for running intool scripts via 
#           PowerShell.  You can use this on demand or to set up scheduled
#           jobs.  Just be sure to save this as something else.

    INSTRUCTIONS: NEED TO TUNR ON WINRM AND ADD HOSTA TO HOSTB's TRUSTED HOSTS
                  ON EACH SERVER!!!
                  
#
# VERSION HISTORY
# 1.0 2014.04.18 Initial Version.
#
# TO ADD:
# 
# USEFUL SNIPPETS:
# "$(Get-Date) " | Out-File -Append $runLog
# #############################################################################>

Param(
    [string] $emplid,
    [string] $docId
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
$scriptName = "mergeTiff" #change this here - no extension

#-- WORKING PATHS --#
$webBaseDir = "D:\inserver6\output\"
$webimgDir = "${webBaseDir}${emplid}\${docId}\"
$completeDir = "${webBaseDir}complete\${emplid}"
$expDir = "\\DIDEV67WEBRPT01\D$\INMAC\tiffExport\${emplid}\"
$mergeDir = "\\DIDEV67WEBRPT01\D$\INMAC\tiffExport\${emplid}\${docId}\"
$tiffcp = "\\DIDEV67WEBRPT01\D$\Program Files (x86)\GnuWin32\bin\tiffcp.exe"
$tiff2pdf = "\\DIDEV67WEBRPT01\D$\Program Files (x86)\GnuWin32\bin\tiff2pdf.exe"

#- LOGGING -#
$runLog = "${root}log\run_log-${scriptName}.log"

#-- MAIN --#
"$(get-date) - Starting ${scriptName} Script" | Out-File $runLog -Append
"WEBRPT SERVER = ${convServ}" | Out-File $runLog -Append
#D:\inserver6\bin64\intool --cmd run-iscript --file ${root}script\${scriptName}.js
Enter-PSSession -ComputerName $convServ -Authentication NegotiateWithImplicitCredential

New-Item ${mergeDir} -ItemType directory

Get-ChildItem ${webimgDir}* -File | ForEach-Object {
    "$($_.FullName)" | Out-File $runLog -Append
    Copy-Item $_ -Destination ${mergeDir}
    $error[0] | Out-File $runLog -Append
}

    $mergeFile = ""
    $pdfOut = ""
    Get-ChildItem ${mergeDir}* -File | ForEach-Object {
        $renArr = [regex]::split($_.Name,'_|\.')
        $mergeFile = "${mergeDir}$($renArr[0])_$($renArr[1]).$($renArr[3])"
        $pdfOut = "${mergeDir}$($renArr[0])_$($renArr[1]).PDF"
        "$pdfOut" | Out-File $runLog -Append
        Rename-Item $_ -NewName "$($renArr[2])_$($renArr[0])_$($renArr[1]).$($renArr[3])"
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

    #Remove-Item -Path $mergeFile
    New-Item -Path ${completeDir} -ItemType directory
    Move-Item -Path $pdfOut -Destination ${completeDir}
    Remove-Item -Path ${expDir} -Recurse
Exit-PSSession

#cleanup
Get-ChildItem $webimgDir -File | ForEach-Object {
     Remove-Item $_.FullName
}
Remove-Item "${webBaseDir}${emplid}\" -Recurse

<#
Use this if you want to be notified when the script finishes running#>

<#$attachments = @()
Get-ChildItem -Path ${webimgDir}\complete\* | ForEach-Object {
$attachments += $_.FullName
}
$message = "I guess you want this..."
sendmail -t "gjenczyk@umassp.edu","cmatera@umassp.edu" -s "[DI ${env} Notice]Here is the pdf you need" -m ${message} -a $attachments
Get-ChildItem -Path ${webimgDir}\complete\* | ForEach-Object {
Remove-Item $_.FullName
}#>

$error[0] | Out-File $runLog -Append
"$(get-date) - Finishing ${scriptName} Script" | Out-File $runLog -Append