<# #############################################################################
# NAME: cleanUp.ps1
# 
# AUTHOR:  Gregg Jenczyk, UMass (UITS)
# DATE:  2014/01/14
# EMAIL: gjenczyk@umassp.edu
# 
# COMMENT:  This script deleted files older than a given date in a given
#           folder on the server.
#            
# VERSION HISTORY
# 1.0 2014.01.14 Initial Version.
# 2.0 2015.03.27 Added handling for permission errors 
# TO ADD:
#
# USEFUL SNIPPETS:
# "$(Get-Date) " | Out-File -Append $runLog
# #############################################################################>
#### CONFIG ####

# Environment #

$root = "Y:\"
$shareRoot = "\\ssisnas215c2.umasscs.net\diimages67prd\"
$env = ([environment]::MachineName).Substring(2)
$logDate = $(get-date -format 'yyyyMMdd')
$env = $env -replace "W.*",""
$inserver6 = "D:\inserver6\"
$dump = "D:\PurgeOldFiles\"

# Log config #

$runLog = "${inserver6}log\run_log-cleanUp_${logDate}.log"
$runLogDelim = "*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*"
${runLogDelim} | Out-File -Append ${runLog}
"$(date) - Cleaning up old files." | Out-File -Append ${runLog}

function cleanUp () {
    Param(
        [string]$dir,
        [string]$age,
        [string[]]$extension,
        [switch]$subfolder,
        [switch]$removeSubs
    )

    if (!$subfolder){

       if (Test-Path $dir) {

        $files =[IO.Directory]::GetFiles($dir)

        foreach($file in $files) {
            $fileName = [System.IO.Path]::GetFileName($file)
            $ageToCheck = New-TimeSpan -days $age
            $fileAge = (get-item $file).LastWriteTime

            if (((get-date) - $fileAge) -gt $ageToCheck) {
                "BLOCK 1: Deleting $file" | Out-File ${runLog} -Append
                Remove-Item -path $file
                if (!$?)
                {
                    "$($error[0])" |  Out-File ${runLog} -Append
                    if ($error[0] -match "Not Enough Permission")
                    {
                        "$file -> $dump"
                        Move-Item -Path $file -Destination $dump
                        if ($?)
                        {
                            "BLOCK 2: Deleting ${dump}${fileName}" | Out-File ${runLog} -Append
                            Remove-Item -path ${dump}${fileName} -Force
                        }
                    }
                } 
            } # end of age check & deletion
        } # end of for each
         
       } #end of if dir exists

    } else {
        
        if (Test-Path $dir) {

            Get-ChildItem $dir | ForEach-Object {
                $ageToCheck = New-TimeSpan -days $age
                if ($_.PSIsContainer)
                {
                    $workingSub = "${dir}$_\"

                    $files =[IO.Directory]::GetFiles($workingSub)
            
                    foreach($file in $files) {

                        $fileAge = (get-item $file).LastWriteTime

                        if (((get-date) - $fileAge) -gt $ageToCheck) {
                            "BLOCK 3: Deleting $file from" | Out-File ${runLog} -Append
                             Remove-Item -path $file
                        } #end of delete
                    } #end of for each in files

                    if($removeSubs)
                    {
                        Get-ChildItem ${workingSub} | ForEach-Object {
                            $subSub = "${workingSub}$_"
                            "subSub $subSub"
                            $folderAge = (get-item $subSub).LastWriteTime
                            if (((get-date) - $folderAge) -gt $ageToCheck)
                            {
                                "BLOCK 4: Removing $subSub || $folderAge" | Out-File -Append ${runLog}
                                Remove-Item -path $subSub -recurse -force
                                if(!$?)
                                {
                                    $subSub = $subSub.Substring(0,$subSub.length-1)
                                    "BLOCK 5: Removing $subSub || $folderAge" | Out-File -Append ${runLog}
                                    Remove-Item -path $subSub -force
                                }
                            } #end of folder cleanUp
                        } #end of for each in sub
                    } #end of removeSubs
                } #end of for directories
                elseif (!$_.PsIsContainer)
                {
                    $subFile = "${dir}$_"
                    $subfileAge = (get-item $file).LastWriteTime
                    if (((get-date) - $subfileAge) -gt $ageToCheck) 
                    {
                        "BLOCK 6: Removing $subFile || $subfileAge" | Out-File -Append ${runLog}
#                        Remove-Item -path $subFile -force
                    } #end of subFile cleanUp
                } #end of subfile cleanup
            } #end of foreach-object
        } #end of if dir exists
    } # end of else

}

# USAGE -d: Directory -a: age of file in days -e: file specifier -s: flag if subfolders need to be removed #

cleanUp -d ${root}import_agent\DI_${env}_SA_AD_INBOUND\ -a 14 -e "*"
cleanUp -d ${root}import_agent\DI_${env}_SA_AD_INBOUND\failure\ -a 14
cleanUp -d ${root}import_agent\DI_${env}_SA_AD_INBOUND\success\ -a 14
cleanUp -d ${root}DI_${env}_COMMONAPP_AD_INBOUND\archive\ -a 14
cleanUp -d ${root}DI_${env}_COMMONAPP_AD_INBOUND\ca_logs\ -a 14
cleanUp -d ${root}DI_${env}_COMMONAPP_AD_INBOUND\error\ -a 14 -s
cleanUp -d ${root}DI_${env}_COMMONAPP_AD_INBOUND\1_unzipFiles\ -a 14 -s
cleanUp -d ${root}DI_${env}_DATABANK_AD_INBOUND\archive\ -a 14
cleanUp -d ${root}DI_${env}_DATABANK_AD_INBOUND\logs\ -a 14
cleanUp -d ${shareRoot}INMAC\out\ -a 5 -s -r
cleanUp -d ${shareRoot}log\ -a 60 -s -r
#cleanUp -d ${root}DI_${env}_DATABANK_AD_INBOUND\error\ -a 7 -s


"$(date) - Clean up complete." | Out-File -Append ${runLog}
