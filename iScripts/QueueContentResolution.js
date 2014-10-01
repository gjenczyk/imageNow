/********************************************************************************
        Name:           Queue_Content_Resolution
        Author:        Gregg Jenczyk
        Created:        03/10/14
        Last Updated:   
        For Version:    
---------------------------------------------------------------------------------
        Summary:
               Finds all documents in an import error queue, reports on them via email, and archives the bad files.
               
        Mod Summary:
               Date-Initials: Modification description.
               
********************************************************************************/

// ********************* Include additional libraries *******************
//#link "inxml"    //XML parser
//#link "sedbc"    //Database object
//#link "secomobj" //COM object
#include "$IMAGENOWDIR6$\\script\\lib\\iScriptDebug.jsh"
#include "$IMAGENOWDIR6$\\script\\lib\\HostDBLookupInfo.jsh"

// *********************         Configuration        *******************

// logging
#define LOG_TO_FILE         true    // false - log to stdout if ran by intool, true - log to inserverXX/log/ directory
#define DEBUG_LEVEL         5       // 0 - 5.  0 least output, 5 most verbose
#define SPLIT_LOG_BY_THREAD false   // set to true in high volume scripts when multiple worker threads are used (workflow, external message agent, etc)
#define MAX_LOG_FILE_SIZE   100     // Maximum size of log file (in MB) before a new one will be created

// *********************       End  Configuration     *******************

// ********************* Initialize global variables ********************
var debug;
var QUEUE_TO_CHECK = "INMAC SA AD Error";
var COUNTER_COUNT = "3";
var CSV_DIR = "Y:\\import_agent\\DI_PRD67_SA_AD_INBOUND\\";
var DOC_DIR = "Y:\\import_agent\\DI_PRD67_SA_AD_INBOUND\\success\\";
var POWERSHELL_ROOT = "C:\\Windows\\System32\\WindowsPowerShell\\v1.0\\powershell.exe";
var POWERSHELL_EMAIL = "\\\\ssisnas215c2.umasscs.net\\diimages67prd\\script\\PowerShell\\QueueContentEmail.ps1"

/**
* Main body of script.
* @method main
* @return {Boolean} True on success, false on error.
*/
function main ()
{
        try
        { 
            debug = new iScriptDebug("QueueContentResolution", LOG_TO_FILE, DEBUG_LEVEL);
            debug.log("WARNING", "QueueContentResolution.js starting.\n");

            var sql="SELECT INUSER.IN_DOC.FOLDER, "+
                    "INUSER.IN_DOC.DOC_ID, "+
                    "INUSER.IN_WF_ITEM.ITEM_ID, "+
                    "INUSER.IN_PHSOB.WORKING_NAME, "+
                    "INUSER.IN_DOC_KW.FREE_FIELD "+
                    "FROM INUSER.IN_WF_ITEM "+
                    "INNER JOIN INUSER.IN_DOC "+
                    "ON INUSER.IN_WF_ITEM.OBJ_ID = INUSER.IN_DOC.DOC_ID "+
                    "INNER JOIN INUSER.IN_WF_QUEUE "+
                    "ON INUSER.IN_WF_ITEM.QUEUE_ID = INUSER.IN_WF_QUEUE.QUEUE_ID "+
                    "INNER JOIN INUSER.IN_VERSION "+
                    "ON INUSER.IN_VERSION.DOC_ID = INUSER.IN_DOC.DOC_ID "+
                    "INNER JOIN INUSER.IN_LOGOB "+
                    "ON INUSER.IN_VERSION.VERSION_ID = INUSER.IN_LOGOB.VERSION_ID "+
                    "INNER JOIN INUSER.IN_PHSOB "+
                    "ON INUSER.IN_PHSOB.PHSOB_ID = INUSER.IN_LOGOB.PHSOB_ID "+
                    "INNER JOIN INUSER.IN_DOC_KW "+
                    "ON INUSER.IN_DOC_KW.DOC_ID = INUSER.IN_DOC.DOC_ID "+
                    "INNER JOIN INUSER.IN_INSTANCE_PROP "+
                    "ON INUSER.IN_DOC.INSTANCE_ID = INUSER.IN_INSTANCE_PROP.INSTANCE_ID "+
                    "INNER JOIN INUSER.IN_PROP "+
                    "ON INUSER.IN_INSTANCE_PROP.PROP_ID = INUSER.IN_PROP.PROP_ID "+
                    "WHERE INUSER.IN_WF_QUEUE.QUEUE_NAME LIKE '"+QUEUE_TO_CHECK+"' "+
                    "AND INUSER.IN_PROP.PROP_NAME LIKE 'Counter' "+
                    "AND INUSER.IN_INSTANCE_PROP.NUMBER_VAL = "+COUNTER_COUNT+" "+
                    "AND INUSER.IN_PHSOB.WORKING_NAME NOT LIKE '%tif' ";
            //debug.log("INFO","sql is [%s]\n",sql);
            var returnVal; 
            var cur = getHostDBLookupInfo_cur(sql,returnVal);
            
            if(!cur || cur == null)
            {
              debug.log("WARNING","no results returned for query.\n");
              var noErrsHead = " Notice] No errors in ";
              var noErrsBody = "There are currently no errored documents in "+QUEUE_TO_CHECK+".\n"+
                               "\n\n\n\nINMAC\n------------------------------------------\n"+
                               "This is an automated message. PLEASE DO NOT REPLY TO THIS MESSAGE.\n"+ 
                               "Questions can be sent to UITS.DI.CORE@umassp.edu";
              sendMail(noErrsHead,noErrsBody);
              return false;
            } 

            //stopgap to prevent masss emailing of tons of documents if something goes wrong
            var rowCount = 0;
            while(cur.next())
            {
              debug.log("DEBUG","Found [%s]\n",cur[1]);
              rowCount++;
            } 
            debug.log("DEBUG","Number of rows returned by query: [%s]\n",rowCount);
            //printf("rowCount = [%s]\n", rowCount)
            if (rowCount > 10)
            {
              debug.log(" WARNING","Unexpected volume of documents [%s].  Script aborting.\n", QUEUE_TO_CHECK);
              //printf("Large volume of documents [%s].  Script aborting.\n", QUEUE_TO_CHECK);
              var impErrorHead = "Error] Unexpected volume of documents in ";
              var impErrorBody = "There are currently "+rowCount+" documents in "+QUEUE_TO_CHECK+".\n"+
                                 "This may indicate that there was a serious problem that affected the imports process.\n"+
                                 "Please review the documents in "+QUEUE_TO_CHECK+" and take the appropriate corrective action.\n"+
                                 "\n\n\nThe standard import error emails have not been sent to the campuses.\n";
              sendMail(impErrorHead,impErrorBody);
              return false;
            }

            cur.reload();

            while(cur.next())
            {

              /*for (i=0;i<cur.columns();i++)
                {
                    printf("cur[%s] = %s\n",i,cur[i]);    
                }
              */

                var cmd = "";
                var folder = cur[0];
                var docId = cur[1];
                var wfId = cur[2];
                var fileName = cur[3];
                var fullNotes =cur[4];

                //prepare info to pass to *shell script
                var CSV_Name = CSV_DIR+folder+".csv";
                var DOC_Name = DOC_DIR+fileName;

                //trimming errors so we only see the last one  
                var notes = "'"+fullNotes.substring(fullNotes.lastIndexOf("INMAC ERR: "),fullNotes.lastIndexOf(' ('))+"'";
                //printf("%s\n",notes);
                //debug.log("ERROR","Here we are [%s\n\t %s\n\t %s\n\t %s\n\t %s\n\t]\n", POWERSHELL_ROOT, POWERSHELL_EMAIL, CSV_Name, DOC_Name, notes);
                Clib.sprintf(cmd, '%s %s %s %s %s', POWERSHELL_ROOT, POWERSHELL_EMAIL, CSV_Name, DOC_Name, notes);
                var rtn = exec(cmd, 0);

                var wfItem = INWfItem.get(wfId);

                if (wfItem == null || !wfItem)
                { 
                 debug.log("ERROR","Failed to get wfItem. Error: %s\n", getErrMsg());
                 continue;
                }
                else
                {
                  if (!wfItem.archive())
                  {
                    debug.log("ERROR","Failed to archive [%s], ID: [%s], Reason: [%s]\n", cur[0], docId, getErrMsg());
                    continue;
                  }
                  else
                  {
                    debug.log("INFO","Archived [%s], ID: [%s]\n",cur[0], docId);
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

function exec(cmd, expected_return)
{
  //debug.log("INFO", "Exec cmd: *%s*\n", cmd);
  var rtn;
  rtn = Clib.system(cmd);
  debug.log("DEBUG", "exec returned %s\n", rtn);
  // tiffcp doesn't return 0 on success
  if(rtn != expected_return)
  {
    debug.log("ERROR", "Couldn't call system cmd: %s\n", cmd);
    return false;
  }
  else
  {
    return true;
  }
}

function sendMail(header,errInfo) {
    debug.log("INFO","Sending mail with sendMail\n");

    var em = new INMail;
    em.from = "Document.Imaging.Support@umassp.edu";
    em.to="UITS.DI.CORE@umassp.edu";
    em.subject="[DI PRD67" + header + " " + QUEUE_TO_CHECK;
    em.body=errInfo;
    em.smtpFrom="192.168.254.71"
    em.smtp="69.16.78.38";
    em.send();
}

//