# #############################################################################
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
#
# TO ADD:
#
# USEFUL SNIPPETS:
# "$(Get-Date) " | Out-File -Append $runLog
# #############################################################################
#### CONFIG ####

# Environment #

$root = "Y:\"
$env = ([environment]::MachineName).Substring(2)
$env = $env -replace "W.*",""
$inserver6 = "D:\inserver6\"

# Log config #

$runLog = "${inserver6}log\run_log-cleanUp.log"
$runLogDelim = "*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*"
${runLogDelim} | Out-File -Append ${runLog}
"$(date) - Cleaning up old files." | Out-File -Append ${runLog}

function cleanUp () {
    Param(
        [string]$dir,
        [string]$age,
        [string[]]$extension,
        [switch]$subfolder
    )

    if (!$subfolder){

       if (Test-Path $dir) {

        $files =[IO.Directory]::GetFiles($dir)

        foreach($file in $files) {

            $ageToCheck = New-TimeSpan -days $age
            $fileAge = (get-item $file).LastWriteTime

            if (((get-date) - $fileAge) -gt $ageToCheck) {
                "Deleting $file" | Out-File ${runLog} -Append
                Remove-Item -path $file 
            } # end of age check & deletion
        } # end of for each
         
       } #end of if dir exists

    } else {
        
        if (Test-Path $dir) {

         Get-ChildItem $dir | ForEach-Object {

            $workingSub = "${dir}$_\"
            
            $files =[IO.Directory]::GetFiles($workingSub)
            
            foreach($file in $files) {

                $ageToCheck = New-TimeSpan -days $age
                $fileAge = (get-item $file).LastWriteTime

                if (((get-date) - $fileAge) -gt $ageToCheck) {
                    "Deleting $file" | Out-File ${runLog} -Append
                     Remove-Item -path $file 
                } #end of delete
            } #end of for each in files
            } #end of foreach-object
        } #end of if dir exists
    } # end of else

}

# USAGE -d: Directory -a: age of file in days -e: file specifier -s: flag if subfolders need to be removed #

cleanUp -d ${root}import_agent\DI_${env}_SA_AD_INBOUND\ -a 14 -e "*"
cleanUp -d ${root}import_agent\DI_${env}_SA_AD_INBOUND\failure\ -a 7
cleanUp -d ${root}import_agent\DI_${env}_SA_AD_INBOUND\success\ -a 14
cleanUp -d ${root}DI_${env}_COMMONAPP_AD_INBOUND\archive\ -a 14
cleanUp -d ${root}DI_${env}_COMMONAPP_AD_INBOUND\ca_logs\ -a 14
cleanUp -d ${root}DI_${env}_COMMONAPP_AD_INBOUND\error\ -a 7 -s
cleanUp -d ${root}DI_${env}_COMMONAPP_AD_INBOUND\1_unzipFiles\ -a 7 -s
cleanUp -d ${root}DI_${env}_DATABANK_AD_INBOUND\archive\ -a 14
cleanUp -d ${root}DI_${env}_DATABANK_AD_INBOUND\logs\ -a 14
#cleanUp -d ${root}DI_${env}_DATABANK_AD_INBOUND\error\ -a 7 -s


"$(date) - Clean up complete." | Out-File -Append ${runLog}
