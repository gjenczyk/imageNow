/********************************************************************************
        Name:           SA_LoadMonitor
        Author:        Gregg Jenczyk
        Created:        08/01/14
        Last Updated:   
        For Version:    
---------------------------------------------------------------------------------
        Summary:
               Checks to see what (if anything) has been sent to us from SA.
               
        Mod Summary:
               Date-Initials: Modification description.
               
********************************************************************************/

// ********************* Include additional libraries *******************
//#link "inxml"    //XML parser
//#link "sedbc"    //Database object
//#link "secomobj" //COM object
#include "$IMAGENOWDIR6$\\script\\lib\\iScriptDebug.jsh"
#include "$IMAGENOWDIR6$\\script\\lib\\HostDBLookupInfo.jsh"
#include "$IMAGENOWDIR6$\\script\\lib\\envVariable.jsh"

// *********************         Configuration        *******************

// logging
#define LOG_TO_FILE         true    // false - log to stdout if ran by intool, true - log to inserverXX/log/ directory
#define DEBUG_LEVEL         5       // 0 - 5.  0 least output, 5 most verbose
#define SPLIT_LOG_BY_THREAD false   // set to true in high volume scripts when multiple worker threads are used (workflow, external message agent, etc)
#define MAX_LOG_FILE_SIZE   100     // Maximum size of log file (in MB) before a new one will be created

// *********************       End  Configuration     *******************

// ********************* Initialize global variables ********************
var debug;
var DELIVERY_DIR = "\\\\boisnas215c1.umasscs.net\\di_interfaces\\import_agent\\DI_"+ENV_U3+"_SA_AD_INBOUND\\";
var FLAG_FILE_DIR = "Z:\\script\\lock\\SA\\";
var FILE_DIR_LOCK = "Z:\\script\\lock\\SA\\SA_LoadMonitor_lock.txt";
var DRAWER_NAME;
var reportTypes = ["WADOC", "zadr0257-1", "zadr0257-2", "zadr011", "zadr012", "zadr0244", "zadr257a-1", "zadr257a-2"];
var EMAIL_DISTRIBUTION = "gjenczyk@umassp.edu";

/**
* Main body of script.
* @method main
* @return {Boolean} True on success, false on error.
*/
function main ()
{
        try
        { 
            debug = new iScriptDebug("SA_LoadMonitor", LOG_TO_FILE, DEBUG_LEVEL);
            debug.log("WARNING", "SA_LoadMonitor.js starting.\n");

            Date.prototype.addHours= function(h){
              this.setHours(this.getHours()+h);
              return this;
            }

            var now = new Date().addHours(4);
            var hours = now.getHours()
            var minutes = now.getMinutes();
            var weekDay = now.getDay();
            var day = now.getDate();
            var month = now.getMonth() + 1;
            var year = now.getFullYear();

            minutes = padIfNot2(minutes);
            hours = padIfNot2(hours);
            month = padIfNot2(month);
            day = padIfNot2(day);

            var currentTime = hours+":"+minutes;
            var dbDate = ""+month+day+year;

            debug.log("DEBUG","The current DB time is %s on %s, %s.\n",currentTime, weekDay, dbDate);

            //remove lock files from the previous day
            
            var lockFiles = SElib.directory(FLAG_FILE_DIR+"*", false, ~FATTR_SUBDIR);
            if (lockFiles == null)
            {
              debug.log("DEBUG","No lock file found.\n");
            }
            else
            {
              var FileCount = getArrayLength(lockFiles);
              for (var g = 0; g < FileCount; g++)
              {
                //printf("%s\tsize = %d\twrite date/time = %s\n",lockFiles[i].name, lockFiles[i].size,lockFiles[i].write);
                var checktime = new Date()
                checktime = checktime.getTime();
                checktime = checktime.toString();
                checktime = checktime.substring(0, checktime.length-3);
                var fileAgeInHours = (checktime - lockFiles[g].write)/3600;
                //printf("checktime is %s lockwrite is %s Lock file age is: [%s]\n", checktime, lockFiles[i].write, fileAgeInHours);
                if(fileAgeInHours > 20.7)
                {
                  debug.log("WARNING","Deleting old locks [%s] for new load cylce.\n", lockFiles[g].name);
                  Clib.remove(lockFiles[g].name)
                  //delete old locks here.  do txt files and folders
                }
                else
                {
                  debug.log("DEBUG","%s is only [%s] hours old. It has [%s] hours to live!!\n", lockFiles[g].name, fileAgeInHours, 24-fileAgeInHours);
                }
              }
            }

            if (!Clib.fopen(FILE_DIR_LOCK,"r"))
            {
              //blow away old locks
              for (var h = 0; h<reportTypes.length; h++){
                var lockRemoval = reportTypes[h];
                var isLocked = SElib.directory(FLAG_FILE_DIR+lockRemoval)
                if (isLocked != null)
                {
                  Clib.rmdir(FLAG_FILE_DIR+lockRemoval);
                  debug.log("INFO","Removing lock for [%s]\n", lockRemoval);
                }
              }
              Clib.fopen(FILE_DIR_LOCK,"a");
              debug.log("INFO","Established new lock for [%s]\n", dbDate);
              var dbsaver = Clib.fopen(FILE_DIR_LOCK,"w");
              debug.log("DEBUG","DBSAVER: [%s] & DBDATE: [%s]\n",dbsaver, dbDate);
              Clib.fputs(dbDate,dbsaver);
              Clib.fclose(dbsaver);
            } 
            
            
            for (var i = 0; i<reportTypes.length; i++)
            {  
              var reportType = reportTypes[i];
              var importLoads = false;
              var inmacLoads = false;
              
              //Check here for existence of load complete reports
              if (SElib.directory(FLAG_FILE_DIR+reportType))
              {
                continue;
              }

              var sqlfp = Clib.fopen(FILE_DIR_LOCK,"r");
              var sqlDate = Clib.fgets(8,sqlfp);
              Clib.fclose(sqlfp);
              //printf("%s\n",sqlDate);

              importLoads = loadStatusCheck(reportType,sqlDate,1);
              if (importLoads)
              {
                inmacLoads = loadStatusCheck(reportType,sqlDate,2);
              }
              else
              {
                debug.log("DEBUG","No docs for report type [%s] in Import drawer, skipping INMAC check\n",reportType);
                //inmacLoads = false;
              }

              debug.log("INFO", "For [%s]: importLoads is [%s] and inmacLoads is [%s]\n", reportType, importLoads, inmacLoads); 

              if (inmacLoads != false && importLoads == inmacLoads)
              {
                debug.log("INFO","Loading of [%s] is complete.\n", reportToEnglish(reportType));
                
                generateFileList(reportType, checktime);

                var emailCounts = "";

                emailCounts = generateCountArray(reportType);

                var emailMessage = "Succesfully imported " + inmacLoads + " " + reportType + " file(s) from SA."+
                                   "\n\n" + emailCounts + "\n\nINMAC\n------------------------------------------\n"+
                                   //"\n\n\n\nINMAC\n------------------------------------------\n"+
                                   "This is an automated message. PLEASE DO NOT REPLY TO THIS MESSAGE.\n"+ 
                                   "Questions can be sent to UITS.DI.CORE@umassp.edu";
                
                sendMail(reportType,emailMessage);
                Clib.mkdir(FLAG_FILE_DIR+reportType);
              }

              if (inmacLoads != false && importLoads > inmacLoads)
              {
                debug.log("INFO","importLoads = [%s], inmacLoads = [%s].  Checking to see if there are any documents in INMAC SA AD Error\n", importLoads, inmacLoads);
                var errorLoads = loadStatusCheck(reportType,sqlDate,3);
                if (inmacLoads + errorLoads == importLoads)
                {
                  
                  generateFileList(reportType, checktime);

                  var emailCounts = "";

                  emailCounts = generateCountArray(reportType);

                  debug.log("INFO","Loading of [%s] is complete.\n", reportToEnglish(reportType));
                  var emailMessage = "Successfully imported " + inmacLoads + " " + reportType + " file(s) from SA.\n" +
                                    errorLoads + " " + reportType + " file(s) did not load due to conversion errors.\n" +
                                   "\n\n" + emailCounts + "\n\nINMAC\n------------------------------------------\n"+
                                   //"\n\n\n\n\nINMAC\n------------------------------------------\n"+
                                   "This is an automated message. PLEASE DO NOT REPLY TO THIS MESSAGE.\n"+ 
                                   "Questions can be sent to UITS.DI.CORE@umassp.edu";
                  sendMail(reportType,emailMessage);
                  Clib.mkdir(FLAG_FILE_DIR+reportType);
                }
              }

            }

        }
        
        catch(e)
        {
               if(!debug)
               {
                       printf("\n\nFATAL iSCRIPT ERROR: %s\n\n", e.toString());
               }
               else
               {
                       debug.setIndent(0);
                       debug.log("CRITICAL", "***********************************************\n");
                       debug.log("CRITICAL", "***********************************************\n");
                       debug.log("CRITICAL", "**                                           **\n");
                       debug.log("CRITICAL", "**    ***    Fatal iScript Error!     ***    **\n");
                       debug.log("CRITICAL", "**                                           **\n");
                       debug.log("CRITICAL", "***********************************************\n");
                       debug.log("CRITICAL", "***********************************************\n");
                       debug.log("CRITICAL", "\n\n\n%s\n\n\n", e.toString());
                       debug.log("CRITICAL", "\n\nThis script has failed in an unexpected way.  Please\ncontact the original author of this script within\nyour organization.  For additiona support,\n contact Perceptive Software Customer Support at 800-941-7460 ext. 2\nAlternatively, you may wish to email support@perceptivesoftware.com\nPlease attach:\n - This log file\n - The associated script [%s]\n - Any supporting files that might be specific to this script\n\n", _argv[0]);
                       debug.log("CRITICAL", "***********************************************\n");
                       debug.log("CRITICAL", "***********************************************\n");
                       if (DEBUG_LEVEL < 3 && typeof(debug.getLogHistory) === "function")
                       {
                               debug.popLogHistory(11);
                               debug.log("CRITICAL", "Log History:\n\n%s\n\n", debug.getLogHistory());
                       }
               }
        }
        
        finally
        {
               if (debug) debug.finish();
               return;
        }
}

// ********************* Function Definitions **********************************

function sendMail(report,body) {
    debug.log("DEBUG","Sending email report for [%s] laod\n", report);

    var em = new INMail;
    em.from = "Document.Imaging.Support@umassp.edu";
    em.to= EMAIL_DISTRIBUTION;
    em.subject="[DI " + ENV_U3 + " Notice] " + reportToEnglish(report) + " Load Report";
    em.body=body;
    em.smtpFrom="192.168.254.71"
    em.smtp="69.16.78.38";
    em.send();
}

function padIfNot2(timeVal) {
  if (timeVal.toString().length != 2) {
      timeVal = "0" + timeVal;
  }
  return timeVal;
}

function loadStatusCheck(report,checkDate,stage) {

  //Variable set up
  var ERROR_WF_INFO = "";
  var OTHER_TABLES = "";
  var OTHER_ANDS = "";

  if (stage == 1)
  {
    DRAWER_NAME = "Import"
    OTHER_TABLES = "LEFT OUTER JOIN INUSER.IN_WF_ITEM " + 
                   "ON INUSER.IN_DOC.DOC_ID = INUSER.IN_WF_ITEM.OBJ_ID "
  }
  else if (stage == 2)
  {
    DRAWER_NAME = "INMAC"
  }
  else if (stage == 3)
  {
    DRAWER_NAME = "IMPORT"
    ERROR_WF_INFO = "INUSER.IN_WF_QUEUE.QUEUE_NAME, " +
                    "INUSER.IN_DOC_KW.FREE_FIELD, "
    OTHER_TABLES = "INNER JOIN INUSER.IN_WF_ITEM " +
                   "ON INUSER.IN_DOC.DOC_ID = INUSER.IN_WF_ITEM.OBJ_ID " +
                   "INNER JOIN INUSER.IN_WF_QUEUE " +
                   "ON INUSER.IN_WF_QUEUE.QUEUE_ID = INUSER.IN_WF_ITEM.QUEUE_ID " +
                   "FULL OUTER JOIN INUSER.IN_DOC_KW " +
                   "ON INUSER.IN_DOC.DOC_ID = INUSER.IN_DOC_KW.DOC_ID "
    OTHER_ANDS = "AND INUSER.IN_WF_QUEUE.QUEUE_NAME IN ('INMAC SA AD Error','INMAC SA AD Holding') "
  }

  //filter out the on demand profiles
  if (report == "zadr011" || report == "zadr012")
  {
    report = report + "%RABATCH";
  }

  debug.log("DEBUG","Checking status for %s %s documents.\n", report, DRAWER_NAME);

  var sql = "SELECT INUSER.IN_DRAWER.DRAWER_NAME, " +
            "INUSER.IN_DOC.FOLDER, " +
            ERROR_WF_INFO + 
            "INUSER.IN_VERSION.CREATION_TIME " +
            "FROM INUSER.IN_DOC " +
            "INNER JOIN INUSER.IN_DRAWER " +
            "ON INUSER.IN_DRAWER.DRAWER_ID = INUSER.IN_DOC.DRAWER_ID " +
            "INNER JOIN INUSER.IN_VERSION " +
            "ON INUSER.IN_DOC.DOC_ID = INUSER.IN_VERSION.DOC_ID " +
            "INNER JOIN INUSER.IN_INSTANCE " +
            "ON INUSER.IN_DOC.INSTANCE_ID = INUSER.IN_INSTANCE.INSTANCE_ID " +
            OTHER_TABLES + 
            "WHERE INUSER.IN_DRAWER.DRAWER_NAME = '"+DRAWER_NAME+"' " +
            "AND INUSER.IN_INSTANCE.DELETION_STATUS <> 1 " +
            "AND INUSER.IN_DOC.FOLDER like '" + report + "%' " +
            OTHER_ANDS + 
            "AND INUSER.IN_VERSION.CREATION_TIME > TO_DATE('"+checkDate+" 06:00:00AM','MMDDYYYY HH:MI:SSAM')";
  var returnVal; 
  var cur = getHostDBLookupInfo_cur(sql,returnVal);
            
  if(!cur || cur == null)
  {
    debug.log("WARNING","no results returned for query.\n");
    return false;
  } 
  
  var rowCount = 0;
  while(cur.next())
  {
    rowCount++;
  }  

  return rowCount;

/*  while(cur.next())
  {     
    //do stuff
  }
*/
}

function reportToEnglish(badName) {
  var goodName;
  debug.log("DEBUG","Inside report name converter.\n");
  switch (badName){
    case "WADOC":
      goodName = "Web Attachments";
      break;
    case "zadr0257-1":
      goodName = "Web Applications";
      break;  
    case "zadr0257-2":
      goodName = "Disclosure Statements";
      break;
    case "zadr011":
      goodName = "Undergraduate Profile Sheet";
      break; 
    case "zadr012":
      goodName = "Graduate Profile Sheet";
      break;
    case "zadr0244":
      goodName = "New Test Scores";
      break; 
    case "zadr257a-1":
      goodName = "Reprinted Web Applications";
      break;
    case "zadr257a-2":
      goodName = "Reprinted Disclosure Statements";
      break; 
    default:
      goodName = "Some report name"; 
  }  
  debug.log("DEBUG","report name is %s\n",goodName);
  return goodName;
}

function generateFileList(fileType, currentTime){

  debug.log("DEBUG","Generating list for [%s] at [%s]\n", fileType, currentTime);

  var allCsvs = SElib.directory(DELIVERY_DIR+fileType+"*", false, ~FATTR_SUBDIR)
  if (allCsvs == null)
    {
        debug.log("ERROR","No files found for search spec \"%s\".\n",DELIVERY_DIR+fileType+"*");
    }
    else
    {
        var csvCount = getArrayLength(allCsvs);
        var fileCatch;
        var filePtr;
        var csvPtr;
        
        fileCatch = FLAG_FILE_DIR+fileType+".txt";


        for (var j = 0; j < csvCount; j++)
        {
          var fileContents = "";
          //printf("%s\n",(currentTime - allCsvs[i].write)/3600);
          //printf("%s\n%s\n------------\n",Clib.ctime(currentTime),Clib.ctime(allCsvs[i].write));
          //printf("[%s]%s Create date/time = %s\n", currentTime, allCsvs[i].name, allCsvs[i].write);
          if ((currentTime - allCsvs[j].write)/3600 <= 8 && allCsvs[j].name.indexOf(".csv") > -1) //makes sure we're not looking at older files or docms
          {            

            filePtr = Clib.fopen(fileCatch,"a");
            csvPtr = Clib.fopen(allCsvs[j].name,"r");

            while (fileContents != null)
            {
              fileContents = Clib.fgets(csvPtr);
              //printf("%s\n",allCsvs[i].name);
              if(fileContents)
              {
                Clib.fputs(fileContents,filePtr);
              }
            }

            Clib.fclose(csvPtr);
            Clib.fclose(filePtr);         
          }
        }    
    }

}

function generateCountArray(reportToCount){

  debug.log("DEBUG","Generating [%s] array count\n", reportToCount);

  var officeCount = new Array();
  var officeSort = new Array();
  var trimFile = FLAG_FILE_DIR+reportToCount+".txt";
  debug.log("DEBUG","trimFile is [%s]\n", trimFile);
  var trimPtr = Clib.fopen(trimFile,"r");

  var pattern = new RegExp(/(.*?)\^([\^UM].*?)[\^]([A-Z].*?)[\^]([A-Z].*?)[\^].*/);
  var fileLine = "";
  var lineSplit;

  while (fileLine != null)
  {
    fileLine = Clib.fgets(trimPtr);
    if (fileLine != null)
    {
      var arrayPrep = "";
      lineSplit = pattern.exec(fileLine);
      
      arrayPrep = lineSplit[2] + " " + lineSplit[3] + " " + lineSplit[4];
      //printf("%s\n", arrayPrep);
      officeCount.push(arrayPrep);
    }

  }

  Clib.fclose(trimPtr);

  officeCount.sort();

  var a = [];
  var b = [];
  var prev;

  for (var j = 0; j < officeCount.length; j++)
  {
    if (officeCount[j] !== prev) {
      a.push(officeCount[j]);
      b.push(1);
    }
    else
    {
      b[b.length-1]++;
    }
    prev = officeCount[j];
  }

  officeSort = [a, b]

  var campusCounts = "";
  for(var z = 0; z < officeSort[0].length; z++)
  {
      campusCounts += officeSort[0][z] + " --- " + officeSort[1][z] + "\n";
  }

  return campusCounts;

}

//