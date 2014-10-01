#!/bin/bash
set -o errexit

errorCode="0"

# get environment abreviation (DEV|TST\PRD)
hostAbrev="$(hostname)"
hostAbrev="${hostAbrev:2:3}"
hostAbrev="$(echo ${hostAbrev} | tr '[:lower:]' '[:upper:]')"
lockTo=/export/$(hostname -s)/inserver6/script/lock/cleanUp.lock

# log initialization
runLog="/export/$(hostname -s)/inserver6/log/running_log-cleanUp.log"
runLogDelim="*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*"
echo ${runLogDelim} >> ${runLog}
echo "$(date) - Cleaning up old files." >> ${runLog}

if (mkdir $lockTo); then >> ${runLog} 2>&1


#usage: cleanupOldFiles "folder & subfolders to clean out" "age (days) of files to remove" "campus OR "
function cleanUp {

	local subfolders=false
	local multiCampus=false
	local extensions=( "*.doc" "*.docx" "*.PDF" "*.pdf" "*.txt" "*.zip" "*.docm" )
	local offices=( "UMBOS" "UMDAR" "UMLOW" )

	while [ $# -gt 0 ]
	do
		case $1 in
			-f) subfolders=true;;
			-m) multiCampus=true;;
			*) break;;
		esac
		shift
	done
	if [ $multiCampus == true ]; then
		for office in "${offices[@]}"
		do
			local curcam=""
			curcam="${1}$office"
			if [ "$(ls -A ${curcam})" ]; then
					echo "$(date) - Checking ${curcam}" >> $runLog 2>&1
					find ${curcam} -maxdepth 1 -type f -name "${3}" -mtime +${2} -exec rm -f -v {} \; >> $runLog 2>&1
			fi
		done
	elif [ $subfolders == false ]; then

		if [ "${3}" == "multi" ]; then
			for type in "${extensions[@]}"
			do
				if [ "$(ls -A ${1})" ]; then
					echo "$(date) - Checking ${1} for $type" >> $runLog 2>&1
					find ${1} -maxdepth 1 -type f -name "$type" -mtime +${2} -exec rm -f -v {} \; >> $runLog 2>&1
				fi
			done
		else
			if [ "$(ls -A ${1})" ]; then
				echo "$(date) - Checking ${1}" >> $runLog 2>&1
				find ${1} -maxdepth 1 -type f -name "${3}" -mtime +${2} -exec rm -f -v {} \; >> $runLog 2>&1
			fi
		fi
	else
		if [ "$(ls -A ${1})" ]; then
			echo "$(date) - Checking ${1}" >> $runLog 2>&1
			find ${1} -type f -name "${3}" -mtime +${2} -exec rm -f -v {} \; >> $runLog 2>&1
			
			for D in `find ${1} -mindepth 1 -type d` 
			do
				if [ "$(ls -A ${D})" ]; then
					continue
				else
					rmdir --verbose ${D}  >> $runLog 2>&1
				fi
			done
		fi
	fi
}

# remove files older than x days.
#USAGE: "-f if subfolders need to be removed" "directory" "age of file in days" "filetypes to remove (multi if there could be multiple file types)"
cleanUp "/di_interfaces/import_agent/DI_${hostAbrev}_SA_AD_INBOUND/" "5" "*.csv"
cleanUp "/di_interfaces/import_agent/DI_${hostAbrev}_SA_AD_INBOUND/failure/" "5" "multi" 
cleanUp "/di_interfaces/import_agent/DI_${hostAbrev}_SA_AD_INBOUND/success/" "5" "multi" 
cleanUp "/di_interfaces/DI_${hostAbrev}_COMMONAPP_AD_INBOUND/archive/" "7" "*.zip"
cleanUp "/di_interfaces/DI_${hostAbrev}_COMMONAPP_AD_INBOUND/ca_logs/" "30" "*.csv"
cleanUp -m "/di_interfaces/DI_${hostAbrev}_COMMONAPP_AD_INBOUND/1_unzipFiles/" "7" "*.xml"
cleanUp "/di_interfaces/DI_${hostAbrev}_COMMONAPP_AD_INBOUND/error/" "7" "multi"
#cleanUp "/di_interfaces/DI_${hostAbrev}_DATABANK_AD_INBOUND/archive/" "14" "*.zip"
#cleanUp -f "/di_interfaces/DI_${hostAbrev}_DATABANK_AD_INBOUND/error/" "7" "*.csv"

echo "$(date) - Clean up complete." >> ${runLog}


rm -rf "$lockTo" >> ${runLog} 2>&1
exit ${errorCode}
else

echo "$(date) - Script already running" >> ${runLog} 

fi