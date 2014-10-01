#!/bin/bash
set -o errexit

### --- Configuration --- ###
# Set this to 1 if you need to load some files and don't want emails to go to everybody.
sneakLoading="0"
# Set this to 1 if we're at the low point in the cycle and may expect fewer than the full amount of zips.
cycleMessage="1"

if [ $sneakLoading == 1 ]; then 
	# Enter the email addresss to send notifications to here:
	sneakEmail="gjenczyk@umassp.edu"
	alertEmail=$sneakEmail
	errorReportEmail=$sneakEmail
	bostonAlert=$sneakEmail
	dartmouthAlert=$sneakEmail
	lowellAlert=$sneakEmail
	bostonErrors=$sneakEmail
	dartmouthErrors=$sneakEmail
	lowellErrors=$sneakEmail
	runType="_MANUAL"
else
	alertEmail="UITS.DI.CORE@umassp.edu"
	errorReportEmail="UITS.DI.CORE@umassp.edu caohelp@commonapp.net"
	bostonAlert="UITS.DI.CORE@umassp.edu john.drew@umb.edu lisa.williams@umb.edu krystal.burgos@umb.edu"
	dartmouthAlert="UITS.DI.CORE@umassp.edu athompson@umassd.edu kmagnusson@umassd.edu j1mello@umassd.edu kvasconcelos@umassd.edu mortiz@umassd.edu"
	lowellAlert="UITS.DI.CORE@umassp.edu christine_bryan@uml.edu kathleen_shannon@uml.edu"
	bostonErrors="UITS.DI.CORE@umassp.edu caohelp@commonapp.net john.drew@umb.edu lisa.williams@umb.edu krystal.burgos@umb.edu"
	dartmouthErrors="UITS.DI.CORE@umassp.edu caohelp@commonapp.net athompson@umassd.edu kmagnusson@umassd.edu j1mello@umassd.edu kvasconcelos@umassd.edu mortiz@umassd.edu"
	lowellErrors="UITS.DI.CORE@umassp.edu caohelp@commonapp.net christine_bryan@uml.edu kathleen_shannon@uml.edu"
	runtype=""
fi

errorBody="Common App Support: Please regenerate and resend the the attached records via SDS.
Contact the UMass Document Imaging team (UITS.DI.CORE@umassp.edu) with any questions.
-
DO NOT REPLY TO THIS EMAIL."

### --- Libraries and functions --- ###

# Logging library for ImageNow
source "/export/$(hostname -s)/inserver6/script/lib/DILogger-Library.sh"

# Locking library (NFS safe)
source "/export/$(hostname -s)/inserver6/script/lib/MutexLock-Library.sh"

# BEGIN RUNNING LOG
errorCode="0"
runLogDelim="*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*"
runLog="/export/$(hostname -s)/inserver6/log/${runType}commonAppImport_run_log.log"
runLock="/export/$(hostname -s)/inserver6/script/lock/commonAppImport.lock"

echo ${runLogDelim} >> ${runLog}
echo "$(date) - BEGINING commonAppImport.sh" >> ${runLog}

# usage: DIImportLogging.log "sub-process name" "campus abreviation" "header string" "info to log"
# globals: DIImportLogging_workingPath,
function DIImportLogging.log {
	# build log filename + path
	local logFileName="$(date +%Y-%m-%d)|${2}|Common_App|${1}${runType}.csv"
	local fullLogPath="${DIImportLogging_workingPath}${logFileName}"

	# check if log file already exists

	# if the file hasn't been created yet, create it and write out the header
	if [[ ! -e ${fullLogPath} ]]; then
		echo "${3}" > "${fullLogPath}"
	fi

	# append log line to log
	echo "${4}" >> "${fullLogPath}"
}

function translateCampusName {
	case "$1" in

	UMBOS)  echo "UMBUA"
	    ;;
	UMDAR)  echo  "UMDUA"
	    ;;
	UMLOW)  echo  "UMLUA"
	    ;;
	esac
}

function shortCampusName {
	case "$1" in

	UMBOS)  echo "B"
	    ;;
	UMDAR)  echo  "D"
	    ;;
	UMLOW)  echo  "L"
	    ;;
	esac
}

function longCampusName {
	case "$1" in

	UMBOS)  echo "UMass Boston"
	    ;;
	UMDAR)  echo  "UMass Dartmouth"
	    ;;
	UMLOW)  echo  "UMass Lowell"
	    ;;
	esac
}

function campusEmail {
	case "$1" in

	UMBOS)  echo ${bostonAlert}
	    ;;
	UMDAR)  echo ${dartmouthAlert}
	    ;;
	UMLOW)  echo ${lowellAlert}
	    ;;
	esac
}

function campusErrorEmail {
	case "$1" in

	UMBOS)  echo ${bostonErrors}
	    ;;
	UMDAR)  echo ${dartmouthErrors}
	    ;;
	UMLOW)  echo ${lowellErrors}
	    ;;
	esac
}

# find all the zip files in the directory and itterate through each one
function commonAppUnzip {
	echo "$(date) - Unzipping files for ${1}..." >> ${runLog}
	local currentCampus="${1}"
	local logHeader="Date/Time,File Name"
	local unzipDate="$(date +%Y-%m-%d)"
	local zipsToProcess=`ls -1 ${inputDirectory}${currentCampus}_*.zip 2>/dev/null | wc -l`
	local campusAlert=$(campusEmail $currentCampus)
	local emailFile="${inputDirectory}ca_logs/${unzipDate}|$(translateCampusName $currentCampus)|Common_App|1_unzip${runType}.csv"
	local unzipPath="1_unzipFiles/${currentCampus}"
	local fileCount
	# odd syntax is to protect against whitespace and other characters in the filenames
	find . -maxdepth 1 -type f -name "${currentCampus}_*.zip" -print0 | while IFS= read -rd $'\0' f ; do >> $runLog 2>&1
		# log out current file
		local fileDate="$(date +%Y-%m-%d\ %I:%M:%S\ %p -r "$f")"
		DIImportLogging.log "1_unzip" "$(translateCampusName $currentCampus)" "${logHeader}" "${fileDate},${f}"

		# perform the unzip
		
		if ! unzip -t -q -d "${unzipPath}" "${f}"; then 
			echo "could not unzip file [${f}]. It has been moved to the error directory" |
				mailx -s "[DI ${hostAbrev} Error] Common App Import Error" ${alertEmail}
			mv "${f}" "error"
			errorCode="1"
		else
			unzip -q -d "${unzipPath}" "${f}" >> $runLog 2>&1
			mv "${f}" "archive" >> $runLog 2>&1
		fi
	done

	echo "$(date) - Finished unzipping files for ${1}" >> ${runLog}
	fileCount=`ls ${unzipPath} | grep pdf$ | wc -l`

	#Sends an email, may alert if there is a problem with the number of zips rec'd depending on the cycleMessage flag
	if [ $cycleMessage -eq 1 ]; then
		if [ $zipsToProcess -eq 0 ]; then
			(echo -e "No files were received for $(longCampusName $currentCampus) on ${unzipDate}") | mailx -s "[DI $hostAbrev Notice] ${1} No Common App Files Received for ${unzipDate}" ${campusAlert} >> $runLog 2>&1
		else
			(echo -e "Attached is a list of zip files received from Common App for $(longCampusName $currentCampus) on ${unzipDate} \nThe zips contained ${fileCount} pdfs."; uuencode "${emailFile}" "${unzipDate}_$(longCampusName $currentCampus)_Common_App_zips.csv") | mailx -s "[DI $hostAbrev Notice] ${zipsToProcess} ${1} Common App File(s) Received for ${unzipDate}" ${campusAlert} >> $runLog 2>&1
		fi
	else
		if [ $zipsToProcess -eq 0 ]; then
			(echo -e "No files were received for $(longCampusName $currentCampus) on ${unzipDate}.\nYou may want to verify this on the Common App Control Center") | mailx -s "[DI $hostAbrev Warning] ${1} No Common App Files Received for ${unzipDate}" ${campusAlert} >> $runLog 2>&1
		elif [ $currentCampus != "UMDAR" -a $zipsToProcess -eq 1 ]; then
			(echo -e "Two zips were expected for $(longCampusName $currentCampus) on ${unzipDate}.\nWe have only received ${zipsToProcess}.\nYou may want to verify this is correct on the Common App Control Center. \nThe zip(s) contained ${fileCount} pdfs."; uuencode "${emailFile}" "${unzipDate}_$(longCampusName $currentCampus)_Common_App_zips.csv") | mailx -s "[DI $hostAbrev Warning] ${1} Unexpected Number of Common App Files Received for ${unzipDate}" ${campusAlert} >> $runLog 2>&1
		elif [ $currentCampus == "UMDAR" -a $zipsToProcess -lt 3 ]; then
			(echo -e "Three zips were expected for $(longCampusName $currentCampus) on ${unzipDate}.\nWe have only received ${zipsToProcess}.\nYou may want to verify this is correct on the Common App Control Center. \nThe zip(s) contained ${fileCount} pdfs."; uuencode "${emailFile}" "${unzipDate}_$(longCampusName $currentCampus)_Common_App_zips.csv") | mailx -s "[DI $hostAbrev Warning] ${1} Unexpected Number of Common App Files Received for ${unzipDate}" ${campusAlert} >> $runLog 2>&1
		else
			(echo -e "Attached is a list of zip files received from Common App for $(longCampusName $currentCampus) on ${unzipDate} \nThe zips contained ${fileCount} pdfs."; uuencode "${emailFile}" "${unzipDate}_$(longCampusName $currentCampus)_Common_App_zips.csv") | mailx -s "[DI $hostAbrev Notice] ${1} Common App Files Received for ${unzipDate}" ${campusAlert} >> $runLog 2>&1
		fi
	fi
}

function checkForErrorPDFs {
	echo "$(date) - Checking for errors for ${1}" >> ${runLog}
	local currentCampus="${1}"
	local logDate="$(date +%Y-%m-%d)"
	local dirToCheck="${inputDirectory}1_unzipFiles/${currentCampus}"
	local todaysLog="ca_logs/${1}_errorPDF_${logDate}${runType}.csv"
	local filesToCheck=`ls -1 ${dirToCheck}/*.pdf 2>/dev/null | wc -l`
	local numberOfErrors=0
	local logHeader="CAMPUS,CAID,CODE,STUDENT NAME,CEEB,RECOMMENDER ID"
	local targetEmail=$(campusErrorEmail $currentCampus)
	local emailFile="${DIImportLogging_workingPath}${logDate}|$(translateCampusName $currentCampus)|Common_App|3_errors${runType}.csv"

	if [ ${filesToCheck} != 0 ]; then
		grep -r -i "Error retreiving document for" $dirToCheck | awk ' {for (i=3; i<NF; i++) printf $i " "; $NF=""; print $NF}' |  sed 's/[ \t]*$//' >> `echo $todaysLog`
		find $dirToCheck -size 753c >> `echo $todaysLog`

		while IFS= read line; do 
			mv "${line}" "${inputDirectory}error/${1}" >> $runLog 2>&1
			numberOfErrors=$((numberOfErrors+1))
		done < "${inputDirectory}$todaysLog"
		if [ ${numberOfErrors} != 0 ]; then
			set +H
			emailFormatRegex="^.*\(([0-9]+)\)*([a-zA-Z]{2,3})_*([0-9]*)_([0-9]+)_+([\'[:space:]A-Za-z-]+_[\'[:space:]A-Za-z-]+)_.*([a-zA-Z]{2,3}+)_*([0-9]*).pdf$"

			while IFS= read f; do
				local userFriendly="$(echo ${f} | sed -re "s#${emailFormatRegex}#${1},\4,\2,\5,\3,\7 #g")"
				local caseFixer="$(echo $userFriendly | tr '[:lower:]' '[:upper:]')"
				DIImportLogging.log "3_errors" "$(translateCampusName $currentCampus)" "${logHeader}" "${caseFixer}"
			done <   "${inputDirectory}$todaysLog"

			rm  "${inputDirectory}$todaysLog" >> $runLog 2>&1

			(echo -e "${errorBody}"; uuencode "${emailFile}" "${logDate}_$(longCampusName $currentCampus)_Common_App_errors.csv") | mailx -s "[DI $hostAbrev Notice] ${1} Common App Reprints for ${logDate}" ${targetEmail} >> $runLog 2>&1
		fi
	fi
	echo "$(date) - Error check complete for ${1}" >> ${runLog}
}

function commonAppProcessFiles {
	echo "$(date) - Processing pdfs for ${1}" >> ${runLog}
	local currentCampus="${1}"
	local currDate="$(date +%y%m%d)"
	local drawerName=$(shortCampusName $1)
	local dateTime="$(date +%Y-%m-%d\ %I:%M:%S\ %p)"
	# setup logging
	local logHeader="Date/Time,Original File Name,Renamed File Name"
	# turn off history substitution to allow for correct regex parsing
	set +H
	incomingFilesRegex="^.*\)([a-zA-Z]{2,3})_([0-9]+)_.*_([A-Z]+).pdf$"
	
	uberRegex="^.*\(([0-9]+)\)*([a-zA-Z]{2,3})_*([0-9]*)_([0-9]+)_.*_([A-Z]+)_*([0-9]*).pdf$"

	find "1_unzipFiles/${1}" -regextype posix-extended -regex "${uberRegex}" -print0 | while IFS= read -rd $'\0' f ; do >> $runLog 2>&1
		# do some stuff to the files here
		local newName="$(echo ${f} | sed -re "s#${uberRegex}#CA_${currDate}_${drawerName}_\4_\2_\5_\3_\1.pdf#g")"
		local standardizedName="$(echo $newName | tr '[:lower:]' '[:upper:]')"
		DIImportLogging.log "2_rename" "$(translateCampusName $currentCampus)" "${logHeader}" "${dateTime},${f},${newName}"
		mv "${f}" "${outputDirectory}${standardizedName}"
	done

	echo "$(date) - Finished processing pdfs for ${1}" >> ${runLog}
}

### --- main script --- ###

# attempt to get a lock on the process
if  (mkdir $runLock) >> ${runLog} 2>&1; then
	
	# get environment abreviation (DEV|TST\PRD)
	hostAbrev="$(hostname)"
	hostAbrev="${hostAbrev:2:3}"
	hostAbrev="$(echo ${hostAbrev} | tr '[:lower:]' '[:upper:]')"

	# build working path
	inputDirectory="/di_interfaces/DI_${hostAbrev}_COMMONAPP_AD_INBOUND/"
	cd "${inputDirectory}" >> $runLog 2>&1

	outputDirectory="/di_interfaces/import_agent/DI_${hostAbrev}_SA_AD_INBOUND/"

	# build log paths
	DIImportLogging_workingPath="ca_logs/"


	# unzip the files
	commonAppUnzip "UMBOS"
	commonAppUnzip "UMDAR"
	commonAppUnzip "UMLOW"

	# check for the Contact Support messagge
	checkForErrorPDFs "UMBOS"
	checkForErrorPDFs "UMDAR"
	checkForErrorPDFs "UMLOW"

	# process the files
	commonAppProcessFiles "UMBOS"
	commonAppProcessFiles "UMDAR"
	commonAppProcessFiles "UMLOW"


	rm -rf "$runLock" >> ${runLog} 2>&1

else

	echo "`date` - Script already running" >> ${runLog} 

fi

echo ${runLogDelim} >> ${runLog}
exit ${errorCode} >> $runLog 2>&1