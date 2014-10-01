#!/bin/bash

#source /export/$(hostname -s)/inserver6/bin/setenv.sh

#cd /export/$(hostname -s)/inserver6/script
#cd /export/$(hostname -s)/inserver6/script

#echo `Date` >> /export/$(hostname -s)/inserver6/log/running_log-CommonAppSearch
#intool --cmd run-iscript --file /export/$(hostname -s)/inserver6/script/CommonAppSearch.js

logTo=/export/$(hostname -s)/inserver6/log/running_log-CommonAppSearch.log
lockTo=/export/$(hostname -s)/inserver6/script/lock/CommonAppSearch.lock
logName=/export/$(hostname -s)/inserver6/log/CommonAppSearch_$(date +"%Y%m%d").log

cd /export/$(hostname -s)/inserver6/script  >> ${logTo} 2>&1

if  (mkdir $lockTo); then >> ${logTo} 2>&1

MACH_OS=`uname -s`
IMAGENOWDIR6=/export/$(hostname -s)/inserver6
ODBCINI=$IMAGENOWDIR6/etc/odbc.ini
LD_LIBRARY_PATH=$IMAGENOWDIR6/odbc/lib:$IMAGENOWDIR6/bin:$IMAGENOWDIR6/fulltext/k2/_ilnx21/bin:/usr/local/waspc6.5/lib:/usr/lib
PATH=$PATH:$IMAGENOWDIR6/fulltext/k2/_ilnx21/bin
IMAGE_GEAR_PDF_RESOURCE_PATH=./Resource/PDF/
IMAGE_GEAR_PS_RESOURCE_PATH=./Resource/PS/
IMAGE_GEAR_HOST_FONT_PATH=./Resource/PS/Fonts/

export IMAGENOWDIR6 ODBCINI LD_LIBRARY_PATH PATH IMAGE_GEAR_PDF_RESOURCE_PATH IMAGE_GEAR_PS_RESOURCE_PATH IMAGE_GEAR_HOST_FONT_PATH  >> ${logTo} 2>&1

echo `date` >> ${logTo}
begin=$(date +%s)
/export/$(hostname -s)/inserver6/bin/intool1 --cmd run-iscript --file /export/$(hostname -s)/inserver6/script/CommonAppSearch.js >> ${logTo} 2>&1
end=$(date +%s)
total=$((end - begin))
if [ $total -lt 30 ]; then
	message="The matching script finished in $total seconds - that seems kind of quick.  
	There may have been an error logging in with intool.  Please verify that the 
	script actually ran."
else
	numberMatch=`grep 'Successfully reindexed' ${logName} | wc -l`
	docLines=`grep 'document id:' ${logName} | wc -l`
	totalDocs=`expr $docLines - 3`
	message="The script finished running at $(date). \nIt took $((total/60)) minutes to make ${numberMatch} matches out of ${totalDocs} documents across all Common App queues."
fi
echo -e ${message} | mailx -s "MyCommonAppSearch.sh has finished running" UITS.DI.CORE@umassp.edu >> ${logTo} 2>&1

rm -rf "$lockTo" >> ${logTo} 2>&1

else

echo "`date` - Script already running" >> ${logTo} 

fi