# #############################################################################
# NAME: UMBUA_Databank_Import.ps1
# 
# AUTHOR:  Gregg Jenczyk, UMass (UITS)
# DATE:  2013/12/31
# EMAIL: gjenczyk@umassp.edu
# 
# COMMENT:  Databank...
#            
# VERSION HISTORY
# 1.0 2013.12.31 Initial Version.
#
# TO ADD:
#
# USEFUL SNIPPETS:
# "$(Get-Date) " | Out-File -Append $runLog
# #############################################################################

#INCLUDES
. "\\ssisnas215c2.umasscs.net\diimages67prd\script\PowerShell\unZip.ps1"
. "\\ssisnas215c2.umasscs.net\diimages67prd\script\PowerShell\sendmail.ps1"

# CONFIG #

$startDir = pwd

$root = "Y:\"
$env = ([environment]::MachineName).Substring(2)
$env = $env -replace "W.*",""
$DIemail = @("UITS.DI.CORE@umassp.edu", "Terence.phalen@umb.edu")  #@("gjenczyk@umassp.edu")
#- Date formatting -#
$myDate = (get-date -format 'yyyy-MM-dd')
$shortDate = (get-date -format 'yyyyMMdd')
$la = (Get-Date -Format yyyy.MM.dd.hh.mm.ss)
$l = [regex]::Split($la, '\.')
$y = $l[0]
$m = $l[1]
$d = $l[2]
$h = $l[3]
$mi = $l[4]
$s = $l[5]


#build paths 
$inputDirectory = "${root}DI_${env}_DATABANK_AD_INBOUND\"
$errorDirectory = "${inputDirectory}error\"
$archiveDirectory = "${inputDirectory}archive\"
$dataBankLog_workingPath = "${inputDirectory}logs\"

#runlog setup
$runLogPath = "\\ssisnas215c2.umasscs.net\diimages67prd\log\"
$runLogFileName = "run_log-UMBUA_Databank_Import_${shortDate}.log"
$fullRunLog = "${runLogPath}${runLogFileName}"
$scriptLog = "D:\inserver6\log\UMBUA_Databank_Import_${y}${m}${d}.log"
$doneLog = "D:\inserver6\log\UMBUA_Databank_Import_DONE_${y}${m}${d}${h}${mi}.log"

cd $inputDirectory

# FUNCTIONS #

function dataBankLog ([string] $subProcessName, [string] $headerString, [string] $infoToLog) {

    #build log filename+path
    $logFileName = "${myDate}-DataBank Imports-${subProcessName}.csv"
    $fullLogPath = "${dataBankLog_workingPath}${logFileName}"

    #checks to see if the log already exists, if not it creates it
    if (!(Test-Path ${fullLogPath})) {
        echo ${headerString} > ${fullLogPath}
    }
    
    #adds the info to the log
    echo ${infoToLog} >> ${fullLogPath}
    return ,$fullLogPath

} #end DIImportLogging.log

# MAIN #
$startTime = Get-Date
"UMBUA_Databank_Import.ps1 starting at $startTime" | Out-File -Append $fullRunLog

Get-ChildItem -Path *.zip | ForEach-Object {
    #declare locations
    $fileName = $_.BaseName
    $unzipDir = "1_unzip\${fileName}\"
    $logHeader = "Date\Time, File Name"
    $fileDate = $_.LastWriteTime.ToString('yyyy-MM-dd\\hh:mm:ss tt')

    $logName = dataBankLog "1_unzip" "${logHeader}" "${fileDate},${fileName}"

    #if the destination folder for the unzipped files doesn't exist, create it
   if (!(Test-Path ${unzipDir})) {
       mkdir $unzipDir
       } # end if Test-Path


  $status = $(unZip "$_" "${inputDirectory}${unzipDir}")

    #if unzipping the file fails, move failed file to an error directory and send an email 
   if ($status -eq $false) {
    "Could not unzip $_.  It has been moved to the error directory." | Out-File -Append $fullRunLog
    $errorFileName = [regex]::split($_, '\\')
    $errorFile = $errorFileName[2]
    $errorText = "There was a problem unzipping $errorFile.`nIt has been moved to ${errorDirectory}"
    sendmail -to $DIemail -s "[DI $env Error] Databank Import Error" -m $errorText
    Move-Item $_ ${errorDirectory}
    } else {
        Move-Item $_ ${archiveDirectory}
        }

} # end Get-ChildItem
try
{
    D:\inserver6\bin64\intool --cmd run-iscript --file \\ssisnas215c2.umasscs.net\diimages67prd\script\UMBUA_Databank_Import\UMBUA_Databank_Import.js >> $fullRunLog
}
catch [system.exception]
{
    $error[0] | Format-List -Force | Out-File -Append $fullRunLog
    if($error[0] -match "The specified network name is no longer available")
    {
        $attempt = 0
        do
        {
            Start-Sleep -Seconds 3
            if((Test-Path -Path D:\inserver6\bin64\intool.exe) -and (Test-Path -Path \\boisnas215c1.umasscs.net\diimages67tst\script\UMBUA_Databank_Import\UMBUA_Databank_Import.js))
            {
                D:\inserver6\bin64\intool --cmd run-iscript --file \\ssisnas215c2.umasscs.net\diimages67prd\script\UMBUA_Databank_Import\UMBUA_Databank_Import.js >> $fullRunLog
                if($?)
                {
                    break
                }
            }
            $error[0] | Format-List -Force | Out-File -Append $fullRunLog
            $attempt++
        }while ($attempt -lt 5)
        if($attempt -ge 4)
        {
            sendmail -to @("UITS.DI.CORE@umassp.edu") -s "There might have been a problem loading databank files" -m "We tried 5 times and couldn't reach the share" -a ${fullRunLog}
        }
    }
}

$endTime = Get-Date
"UMBUA_Databank_Import.ps1 finished running at $endTime" | Out-File -Append $fullRunLog
$totalTime = $endTime-$startTime
"Total run time = $totalTime" | Out-File -Append $fullRunLog

#wrap up notification email
$subject = "[DI $env Notice] DataBank files Received for ${myDate}"

if ((${logName}) -eq $null ){
    $subject = "[DI $env Notice] No DataBank files Received on ${myDate}"        
    $message = "No files were received on ${myDate} for the ${h}:${mi} load."
    sendmail -to $DIemail -s $subject -m $message
} 
elseif (Test-Path ${logName}) 
{
    $alertLvl = "Notice"
    #get number of images loaded
    $fileCount = 0
    Get-Content $scriptLog | ForEach-Object {
        if ($_ -match 'createOrRouteDoc')
        {
            $fileCount++
        }
    }

    if ($fileCount -eq 0)
    {
        $alertLvl = "Error"
    }

    #change log name so counts are not picked up in the next load
    Rename-Item -Path $scriptLog -NewName $doneLog
    #
    $subject = "[DI $env $alertLvl] DataBank files Received on ${myDate}"    
    $humanTime = [regex]::split($totalTime, '\:|\.')
    $minutes = $humanTime[1]
    $seconds = $humanTime[2]
    $message = "Attached is a list of zip files received from DataBank on ${myDate} for the ${h}:${mi} load.`n
    The imports process took ${minutes} minutes and ${seconds} seconds to load $fileCount images."
    sendmail -to $DIemail -s $subject -a $logName -m $message
} 

Get-Item -path $logName | Rename-Item -NewName {$_.Name -replace ".csv","-$(Get-Date -format 'HHmm').csv"}

cd $startDir