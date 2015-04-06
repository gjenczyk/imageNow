/********************************************************************************
        Name:          GA_DocumentExport
        Author:        Gregg Jenczyk
        Created:        04/03/201
        Last Updated:   
        For Version:    
---------------------------------------------------------------------------------
        Summary:
               This script will email a selection of docs in the system to a user
               in the form of a pdf.  Requires tiif lib for gnuwin32.
               
        Mod Summary:
               Date-Initials: Modification description.



********************************************************************************/

// ********************* Include additional libraries *******************
//#link "inxml"    //XML parser
//#link "sedbc"    //Database object
//#link "secomobj" //COM object
#include "$IMAGENOWDIR6$\\script\\lib\\iScriptDebug.jsh"
#include "$IMAGENOWDIR6$\\script\\lib\\HostDBLookupInfo.jsh"
#include "$IMAGENOWDIR6$\\script\\STL\\packages\\Document\\exportDocPhsOb.js"

// *********************         Configuration        *******************

// logging
#define LOG_TO_FILE         true    // false - log to stdout if ran by intool, true - log to inserverXX/log/ directory
#define DEBUG_LEVEL         5       // 0 - 5.  0 least output, 5 most verbose
#define SPLIT_LOG_BY_THREAD false   // set to true in high volume scripts when multiple worker threads are used (workflow, external message agent, etc)
#define MAX_LOG_FILE_SIZE   100     // Maximum size of log file (in MB) before a new one will be created


// *********************       End  Configuration     *******************

// ********************* Initialize global variables ********************

sql = "";
var POWERSHELL_ROOT = "C:\\Windows\\System32\\WindowsPowerShell\\v1.0\\powershell.exe";
var POWERSHELL_MERGE = imagenowDir6+"\\script\\PowerShell\\mergeTiffs.ps1"

/**
* Main body of script.
* @method main
* @return {Boolean} True on success, false on error.
*/
  function main ()
{
    try
    {
      debug = new iScriptDebug("GA_DocumentExport", LOG_TO_FILE, DEBUG_LEVEL);
      debug.log("WARNING", "GA_DocumentExport script started.\n");
      var wfItem = new INWfItem("301YY4P_04VZ7RN4N0142Y5");//currentWfItem.id);//
      if(!wfItem.id || !wfItem.getInfo())
      {
        debug.log("CRITICAL", " Couldn't get info for wfItem: %s\n", getErrMsg());
        return false;
      }
      //collect information about who routed the document into the queue
      printf(wfItem.queueStartUserName);
      var router = new INUser(wfItem.queueStartUserName);
       //get information about inbound doc
       //determine docs to send
       //convert matching docs to pdfs
       //email to person who routed doc in
       //forward doc to complete queue
       //get doc
    	var doc = new INDocument(wfItem.objectId);//"301YY4P_04VZ7RN4N0142XW");  
    	if(!doc.getInfo())
    	{
    	    printf("Couldn't get doc info: %s.\n",
    	    getErrMsg()); return false;
    	}

      //exportDoc(doc);

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
  debug.log("INFO", "Exec cmd: *%s*\n", cmd);
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

function exportDoc(doc)
{
  //make an output dir for the applicant's info because for some damn reason perceptive won't create a dir 2 deep
  var exportDir = imagenowDir6+"\\output\\"+doc.field1+"\\";//+doc.id+"\\";
  Clib.mkdir(exportDir);
  if(!exportDocPhsOb(doc,exportDir + "\\"+doc.id+"\\","ALL","ALL",true))
  {
    printf(getErrMsg());
  }
  else
  {
    var cmd = "";
    Clib.sprintf(cmd, '%s %s %s %s', POWERSHELL_ROOT, POWERSHELL_MERGE, doc.field1, doc.id);
    var rtn = exec(cmd, 0);

  }
}

function emailDocs()
{

}


//