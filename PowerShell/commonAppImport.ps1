# #############################################################################
# NAME: commonAppImport.ps1
# 
# AUTHOR:  Gregg Jenczyk, UMass (UITS)
# DATE:  2013/08/14
# EMAIL: gjenczyk@umassp.edu
# 
# COMMENT:  This script unzips files delivered from Common App, checks the pdfs
#           contained within for errors, renames the valid pdfs according to DI 
#           standards, and moves them to the imports directory.
#
# VERSION HISTORY
# 1.0 2013.08.14 Initial Version.
# 1.1 2014.04.11 Ported all functionality from .sh version
#
# TO ADD
# Nothing, at this time.
#
# USEFUL SNIPPETS
# "$(Get-Date) " | Out-File -Append $runLog
# #############################################################################

#-- INCLUDES --#
. "\\boisnas215c1.umasscs.net\diimages67tst\script\PowerShell\sendmail.ps1"
. "\\boisnas215c1.umasscs.net\diimages67tst\script\PowerShell\unZip.ps1"

#-- CONFIG --#

$root = "D:\"
$shareRoot = "\\boisnas215c1.umasscs.net\di_interfaces\"
$env = ([environment]::MachineName).Substring(2)
$env = $env -replace "W.*",""
$myDate = (get-date -format 'yyyy-MM-dd')
$shortDate = (get-date -format 'yyMMdd')

    #-- EMAIL AND LOADING CONFIGS --#
    # change the value of sneakLoading if you are going to do an unscheduled import so
    # emails do not get sent to the campuses (values: $true or $false)
    $sneakLoading = $true
    # flag to change contents of email based on admissions cycle
    $peakCyle = $false


    if (!$sneakLoading) {
       $alertEmail = ("gjenczyk@umassp.edu") #("UITS.DI.CORE@umassp.edu")
       $errorEmail = ("gjenczyk@umassp.edu") #("UITS.DI.CORE@umassp.edu, caohelp@commonapp.net")
       $bostonEmail = ("gjenczyk@umassp.edu") #("UITS.DI.CORE@umassp.edu, john.drew@umb.edu, lisa.williams@umb.edu, krystal.burgos@umb.edu")
       $dartmouthEmail = ("gjenczyk@umassp.edu") #("UITS.DI.CORE@umassp.edu, athompson@umassd.edu, kmagnusson@umassd.edu, j1mello@umassd.edu, kvasconcelos@umassd.edu, mortiz@umassd.edu")
       $lowellEmail = ("gjenczyk@umassp.edu") #("UITS.DI.CORE@umassp.edu, christine_bryan@uml.edu, kathleen_shannon@uml.edu")
       $bostonError = ("gjenczyk@umassp.edu") #("UITS.DI.CORE@umassp.edu, caohelp@commonapp.net, john.drew@umb.edu, lisa.williams@umb.edu, krystal.burgos@umb.edu")
       $dartmouthError = ("gjenczyk@umassp.edu") #("UITS.DI.CORE@umassp.edu, caohelp@commonapp.net, athompson@umassd.edu, kmagnusson@umassd.edu, j1mello@umassd.edu, kvasconcelos@umassd.edu, mortiz@umassd.edu")
       $lowellError = ("gjenczyk@umassp.edu") #("UITS.DI.CORE@umassp.edu, caohelp@commonapp.net, christine_bryan@uml.edu, kathleen_shannon@uml.edu")
    } else {
       $divertedEmail = ("noone@nowhere.com")
       $comApSupport = $divertedEmail
       $errorEmail = $divertedEmail
       $bostonEmail = $divertedEmail
       $dartmouthEmail = $divertedEmail
       $lowellEmail = $divertedEmail
       $bostonError = $divertedEmail
       $dartmouthError = $divertedEmail
       $lowellError = $divertedEmail
    }
    
    $errorBody = "Common App Support: Please regenerate and resend the the attached records via SDS.
    `nContact the UMass Document Imaging team (UITS.DI.CORE@umassp.edu) with any questions.
    `n-
    `nDO NOT REPLY TO THIS EMAIL."

#FUNCTIONS#

function DIImportLogging.log ([string] $subProcessName, [string] $campusAbbreviation, [string] $headerString, [string] $infoToLog) {
    #build log filename+path
    $logFileName = "${myDate}-${campusAbbreviation}-Common App-${subProcessName}.csv" 
    $fullLogPath = "${DIImportLogging_workingPath}${logFileName}"

    #checks to see if the log already exists, if not it creates it
    if (!(Test-Path ${fullLogPath})) {
        echo ${headerString} > ${fullLogPath}
    }
    
    #adds the info to the log
    echo ${infoToLog} >> ${fullLogPath}
} #end DIImportLogging.log


function translateCampus ([string] $rawName ) {
    #converts the campus names to format used by our scripts
    switch ($rawName)
        {
            UMBOS {"UMBUA"}
            UMDAR {"UMDUA"}
            UMLOW {"UMLUA"}
            default {"error"}
        }
}

function shortCampus ([string] $rawName ) {
    #converts the campus names to format used by our scripts
    switch ($rawName)
        {
            UMBOS {"B"}
            UMDAR {"D"}
            UMLOW {"L"}
            default {"error"}
        }
}

function longCampus ([string] $rawName ) {
    #converts the campus names to format used by our scripts
    switch ($rawName)
        {
            UMBOS {"UMass Boston"}
            UMDAR {"UMass Dartmouth"}
            UMLOW {"UMass Lowell"}
            default {"error"}
        }
}

function emailCampus ([string] $rawName ) {
    #converts the campus names to format used by our scripts
    switch ($rawName)
        {
            UMBOS {$bostonEmail}
            UMDAR {$dartmouthEmail}
            UMLOW {$lowellEmail}
            default {$alertEmail}
        }
}

function errorCampus ([string] $rawName ) {
    #converts the campus names to format used by our scripts
    switch ($rawName)
        {
            UMBOS {$bostonError}
            UMDAR {$dartmouthError}
            UMLOW {$lowellError}
            default {$alertEmail}
        }
}

function commonAppUnzip ([string] $unzipCampus ) {
   #variables
   $logHeader = ("Date\Time, File Name")
   $unzipPath = ("${inputDirectory}1_unzipFiles\${unzipCampus}")
   $unzipDate = (Get-date -format "yyyy-MM-dd")
   $campusAlert = $(emailCampus($unzipCampus))
   $emailFile = ("${inputDirectory}ca_logs\${unzipDate}-$(translateCampus($unzipCampus))-Common App-1_unzip.csv")

   "$(Get-Date) Unzipping files for $unzipCampus" | Out-File -Append $runLog

   #moves zips locally for faster processing.

   Get-ChildItem -Path ${shareRoot}DI_TST67_COMMONAPP_AD_INBOUND -File ${unzipCampus}_*zip | ForEach-Object {
    Move-Item -Path $_.FullName -Destination $inputDirectory
   }

   $zipsToProcess = (Get-ChildItem $inputDirectory -Filter "${unzipCampus}_*.zip").count

   #finds all zips for a campus and writes their name to a log
   Get-ChildItem -File ${unzipCampus}_*zip | ForEach-Object {  
   $fileDate = ($_.LastWriteTime.ToString('yyyy-MM-dd\\ hh:mm:ss tt'))
   DIImportLogging.log "1_unzip" "$(translateCampus ${unzipCampus})" "${logHeader}" "${fileDate},$_"

   #if the destination folder for the unzipped files doesn't exist, create it
   if (!(Test-Path ${unzipPath})) {
       mkdir $unzipPath
       }
   #check for 0 byte zips
   if ($_.length -eq 0) {
        echo $_.FullName "is 0 bytes.  It has been moved to the error directory."
        Move-Item $_ ${errorDirectory}\${unzipCampus}
      } else { #unzip the files 
    
   $unzipFile = $_.FullName     
   $status = $(unZip -file $unzipFile -destination $unzipPath)  

   #if unzipping the file fails, move failed file to an error directory and send an email 
   if ($status -eq $false) {
   $unzipErrSub = "[DI $env Error] Common App Import Error"
   $unzipErrMes = "Could not unzip ${unzipFile}.`nIt has been moved to the error directory"
    sendmail -to $alertEmail -s $unzipErrSub -m $unzipErrMes
    Move-Item $_ ${errorDirectory}
    } else {
        Move-Item $_ ${archiveDirectory}
        }
   }
   }
   "$(Get-Date) Finished unzipping files for $unzipCampus" | Out-File -Append $runLog
   
   $fileCount = (Get-ChildItem $unzipPath -filter "*.pdf").count

   #send email based on number of zips rec'd per campus
    #if $peakCyle is false, send non-threatening emails.
   if ($peakCyle -eq $false){
      if ($zipsToProcess -eq 0){
          sendmail -to $campusAlert -s "[DI $env Notice] ${unzipCampus} No Common App Files Received for ${unzipDate}" -m "No zips were received for $(longCampus($unzipCampus)) on ${unzipDate}."
        } else {
          sendmail -to $campusAlert -s "[DI $env Notice] ${zipsToProcess} ${unzipCampus} Common App Files Received for ${unzipDate}" -m "Attached is a list of zip files received from Common App for $(longCampus($unzipCampus)) on ${unzipDate}.`nThe zips contained ${fileCount} pdfs." -a $emailFile
        }
    } else {
      if ($zipsToProcess -eq 0){
        sendmail -to $campusAlert -s "[DI $env Warning] ${unzipCampus} No Common App Files Received for ${unzipDate}" -m "No zips received for $(longCampus($unzipCampus)) on ${unzipDate}.`nYou may want to verify this on the Common App Control Center."
      } elseif ($unzipCampus -ne "UMDAR" -and $zipsToProcess -eq 1) {
        sendmail -to $campusAlert -s "[DI $env Warning] ${unzipCampus} Unexpected Number of Common App Files Received for ${unzipDate}" -m "Two zips were expected for $(longCampus($unzipCampus)) on ${unzipDate}.`nWe have only received ${zipsToProcess}.`nYou may want to verify this is correct on the Common App Control Center. `nThe zip(s) contained ${fileCount} pdfs." -a $emailFile
      } elseif ($unzipCampus -eq "UMDAR" -and $zipsToProcess -eq 2) {
        sendmail -to $campusAlert -s "[DI $env Warning] ${unzipCampus} Unexpected Number of Common App Files Received for ${unzipDate}" -m "Three zips were expected for $(longCampus($unzipCampus)) on ${unzipDate}.`nWe have only received ${zipsToProcess}.`nYou may want to verify this is correct on the Common App Control Center. `nThe zip(s) contained ${fileCount} pdfs." -a $emailFile
      } else {
        sendmail -to $campusAlert -s "[DI $env Notice] ${unzipCampus} Common App Files Received for ${unzipDate}" -m "Attached is a list of zip files received from Common App for $(longCampus($unzipCampus)) on ${unzipDate}.`nThe zips contained ${fileCount} pdfs." -a $emailFile
      }
    }

} # END commonAppUnzip

function commonAppErrors ([string] $processCampus){
    "$(Get-Date) Checking for errors for $processCampus" | Out-File -Append $runLog
    #setup vars in func
    $logDate = $myDate #(Get-Date -format yyyy-MM-dd) 
    $currentCampus = $processCampus 
    $dirToCheck = ("${inputDirectory}1_unzipFiles\$currentCampus") 
    $todaysLog = ("ca_logs\${currentCampus}_errorPDF_${logDate}.csv")
    $errorCSV = ("ca_logs\${currentCampus}_REPRINTS_${logDate}.csv")
    $filesToCheck = ((Get-ChildItem -Path "$dirToCheck" -exclude "*.xml").count) 
    $targetEmail = errorCampus($processCampus)
    $logHeader = "CAMPUS,CAID,CODE,STUDENT NAME,CEEB,RECOMMENDER ID"
    $emptyPDF
   
    #creates a list of error pdfs to process
    if ($filesToCheck -ne 0) {
        #initializes log file to get rid of whitespace weirdness (probably not necessary when live, but it's handy for testing). 
        New-Item -force -ItemType file $todaysLog | Out-Null
        #create a list of pdfs with problems and remove blank lines from list.
        Select-String -Path $dirToCheck\*.pdf -Simple -Pattern "Error retreiving document for" | Select-Object -unique path | Format-Table -HideTableHeaders >> ${todaysLog}
        Get-childItem $dirToCheck\*.pdf | ? {$_.Length -eq 753} | ForEach-Object -Process {$_.FullName} | Format-Table -HideTableHeaders >> ${todaysLog}
        (gc ${todaysLog}) | ? {$_.trim() -ne "" } | ForEach {$_.TrimEnd()} |set-content ${todaysLog} 

        #gets rid of error files with no content for cleanliness
        if ((get-item $todaysLog).length -eq 0) {
            Remove-Item $todaysLog
        }
    } #End of error csv creation
    
    #move the error files  
    if (Test-Path -Path $todaysLog){
        echo $logHeader > ${errorCSV}
        Get-Content $todaysLog | ForEach-Object {
            Move-Item -LiteralPath $_ -Destination ${errorDirectory}${currentCampus} -Force
            $errorSplit = [regex]::split($_, '\\|_|\(|\)|\.')
            
             #create csv for error email
            if ($errorSplit[11] -eq "cao") {
                $tempName = Write-Output ($errorSplit[8]+','+$errorSplit[12]+',CAO,'+$errorSplit[14]+' '+$errorSplit[13]+',,')
            } elseif ($errorSplit[11] -eq "WS") {
                $tempName = Write-Output ($errorSplit[8]+','+$errorSplit[12]+','+$errorSplit[11]+','+$errorSplit[14]+' '+$errorSplit[13]+',,')   
            }else{
                $tempName = Write-Output ($errorSplit[8]+','+$errorSplit[13]+','+$errorSplit[11]+','+$errorSplit[15]+' '+$errorSplit[14]+','+$errorSplit[12]+','+$errorSplit[17])
            }

            echo $tempName >> ${errorCSV}
        }
        $errSub = "[DI $env Notice] ${processCampus} Common App Reprints for ${logDate}"
        sendmail -to $targetEmail -a ${errorCSV} -s $errSub -m $errorBody
    }
    "$(Get-Date) Finished error check for $processCampus" | Out-File -Append $runLog
}#end commonAppErrors

function commonAppProcess ([string] $processCampus){
    "$(Get-Date) Processing pdfs for $processCampus" | Out-File -Append $runLog
    #setup variables
    $drawerName = $(shortCampus ${processCampus})
    $dateTime = (Get-Date -Format 'yyyy-MM-dd\\ hh:mm:ss tt')
    $logHeader = ("Date\Time, Original File Name, Renamed File Name")
    $incomingFilesRegex = ("^.*\(([0-9]+)\)*([a-zA-Z]{2,3})_*([0-9]*)_([0-9]+)_.*_([A-Z]+)_*([0-9]*).pdf$")
    #
    Get-ChildItem -recurse ${inputDirectory}\1_unzipFiles\${processCampus} | Where-Object {$_.Name -match "${incomingFilesRegex}"} | ForEach-Object {
       #I think these 2 steps are necessary because I can't find a sed equivalent for powershell.  
       $strippedName = [regex]::split($_, '!|_|\.|\(|\)')
        if ($strippedName[2] -eq "cao") {
            $tempName = Write-Output ('CA_'+${shortDate}+'_'+${drawerName}+'_'+$strippedName[3]+'_CAO_'+$strippedName[7]+'_'+$strippedName[0]+'_'+$strippedName[1])
        } elseif ($strippedName[2] -eq "WS") {
            $tempName = Write-Output ('CA_'+${shortDate}+'_'+${drawerName}+'_'+$strippedName[3]+'_'+$strippedName[2]+'_'+$strippedName[6]+'_'+$strippedName[0]+'_'+$strippedName[1])    
        }else{
           $tempName = Write-Output ('CA_'+${shortDate}+'_'+${drawerName}+'_'+$strippedName[4]+'_'+$strippedName[2]+'_'+$strippedName[7]+'_'+$strippedName[3]+'_'+$strippedName[1])
        }
        do {#add to log and move files

        $newName = ($tempName+'.pdf')
        Rename-Item $_.FullName $newName
        DIImportLogging.log "2_rename" "$(translateCampus $processCampus)" "${logHeader}" "${dateTime},$_,$newName"
        
        #Move-Item ${inputDirectory}\1_unzipFiles\${processCampus}\$newName -Destination ${outputDirectory}
        #& ${root}convertFiletype.ps1
        } while ($? -eq 0)

<#        do {
        if (Move-Item .\1_unzipFiles\${processCampus}\$newName -Destination ${outputDirectory} -ne 0) {
            #$newName = ($tempName+'.pdf')
            Move-Item .\1_unzipFiles\${processCampus}\$newName -Destination ${outputDirectory}}
        } while ($? -ne 0)
#>    }
    "$(Get-Date) Finished processing pdfs for $processCampus" | Out-File -Append $runLog
}# end commonAppProcess

function commonAppClean ([string] $cleanCampus) {

    "$(Get-Date) Cleaning up the server for $cleanCampus" | Out-File -Append $runLog
    $altName = $(translateCampus($cleanCampus))

    #move the renamed pdfs
    Get-Item "${inputDirectory}1_unzipFiles\${cleanCampus}\*pdf" | ForEach-Object {
      $pdfName = $_.name
      if ((Get-Item $_.FullName).length -eq 0) {
        "$pdfName is empty :(" | Out-File -Append $runLog
      }
      elseif ((Get-Item $_.FullName) -match "_SR_")
      {
        Move-Item -Path ${inputDirectory}1_unzipFiles\${cleanCampus}\$pdfName -Destination ${srDirectory}${pdfName}

        "${srDirectory}${pdfName}" | Out-File -Append $runLog
      }
      else 
      {
        Move-Item -Path ${inputDirectory}1_unzipFiles\${cleanCampus}\$pdfName -Destination ${outputDirectory}${pdfName} 
      }
       
    }

    #move the logs
    Move-Item -Path "${DIImportLogging_workingPath}*" -Destination "${shareBase}ca_logs\"

    #move the errors
    Move-Item -Path "${errorDirectory}${cleanCampus}\*" -Destination "${shareBase}error\${cleanCampus}\"

    #move xml files from the unzip directories
    Move-Item -Path "${inputDirectory}1_unzipFiles\${cleanCampus}\*xml" -Destination "${shareBase}1_unzipFiles\${cleanCampus}\"

    #finally, move the zips back
    Move-Item -Path "${archiveDirectory}${cleanCampus}_*zip" -Destination "${shareBase}archive\"
} #end commonAppClean


# building the working path - establishing directory locations

$inputDirectory = "${root}DI_${env}_COMMONAPP_AD_INBOUND\"
$srDirectory = "${root}DI_${env}_COMMONAPP_AD_OUTBOUND\"
$outputDirectory = "${shareRoot}import_agent\DI_${env}_SA_AD_INBOUND\"
$shareBase = "${shareRoot}DI_${env}_COMMONAPP_AD_INBOUND\"
$archiveDirectory = "${inputDirectory}archive\"
$errorDirectory = "${inputDirectory}error\"
cd $inputDirectory

# Logging setup

$DIImportLogging_workingPath = "${inputDirectory}ca_logs\"
$DIImportLogging_completePath = "${root}import_agent\${env}_import_logs\"
$runLogDir = "\\boisnas215c1.umasscs.net\diimages67tst\log\"
$runLogDelim = "*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*"
$runLogName = "running_log-commonAppImport.log"
$runLog = "${runLogDir}${runLogName}"


#Begin
$startTime = (Get-Date)

"$runLogDelim`n$startTime commonAppImport.ps1 starting" | Out-File -Append $runLog

#unzippping the files

commonAppUnzip UMBOS
commonAppUnzip UMDAR
commonAppUnzip UMLOW

#extract the error pdfs

commonAppErrors UMBOS
commonAppErrors UMDAR
commonAppErrors UMLOW

#process the files

commonAppProcess UMBOS
commonAppProcess UMDAR
commonAppProcess UMLOW

#clean up the WEBRPT Server

commonAppClean UMBOS
commonAppClean UMDAR
commonAppClean UMLOW


#wrapping it up

$endTime = (Get-Date)
$totalTime = ($endTime-$startTime)

"$endTime commonAppImport.ps1 finished`ntotal run time = $totalTime`n$runLogDelim" | Out-File -Append $runLog

return