/********************************************************************************
        Name:          FA_ReadyToMove
        Author:        Gregg Jenczyk
        Created:        12/15/2014
        Last Updated:   
        For Version:    
---------------------------------------------------------------------------------
        Summary:
               Create a users.txt file - keep appending the output of the 
               attached query to the txt file every hour, send a .csv file every 
               day @ midnight - this will keep track of usage.   
               
        Mod Summary:
               Date-Initials: Modification description.



********************************************************************************/

// ********************* Include additional libraries *******************
//#link "inxml"    //XML parser
//#link "sedbc"    //Database object
//#link "secomobj" //COM object
#include "$IMAGENOWDIR6$\\script\\lib\\iScriptDebug.jsh"
#include "$IMAGENOWDIR6$\\script\\lib\\HostDBLookupInfo.jsh"
#include "$IMAGENOWDIR6$\\script\\lib\\yaml_loader.jsh"

// *********************         Configuration        *******************

// logging
#define LOG_TO_FILE         true    // false - log to stdout if ran by intool, true - log to inserverXX/log/ directory
#define DEBUG_LEVEL         5       // 0 - 5.  0 least output, 5 most verbose
#define SPLIT_LOG_BY_THREAD false   // set to true in high volume scripts when multiple worker threads are used (workflow, external message agent, etc)
#define MAX_LOG_FILE_SIZE   100     // Maximum size of log file (in MB) before a new one will be created


// *********************       End  Configuration     *******************

// ********************* Initialize global variables ********************
CALLED_BY_EM = false;
sql = "";

/**
* Main body of script.
* @method main
* @return {Boolean} True on success, false on error.
*/
  function main ()
{
    try
    {
      debug = new iScriptDebug("FA_ReadyToMove", LOG_TO_FILE, DEBUG_LEVEL);
      debug.log("WARNING", "FA_ReadyToMove script started.\n");

      RTM_processByTask();

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

//This function will receive informaiton sufficient to find a folder.  Return true or 
//false to calling script. 
function RTM_processByTrigger()
{
  //set the flag on a folder for the id passed
  if(!setRTMFlag())
  {
    //could not set the flag
    debug.log("ERROR","Could not set RTM flag for [folder]\n");
    return false;
  }
  //check to see if the folder is ready to move
  if (!checkRouteReady())
  {
    //found documents in the NDE queues
    debug.log("INFO","[folder] is not ready to move.\n");
    return true;
  }
  else
  {
    debug.log("INFO","[folder] is ready to move.\n");
    //try to route folder to correct review process
    if(!routeToReview())
    {
      debug.log("ERROR","Could not route [folder] to review: [%s]\n",getErrMsg());
      return false;
    }
    else
    {
      debug.log("INFO","Routed [folder]\n");
      return true;
    }
  }
}//end RTM_processByTrigger

//This function will get a list of all folders in the Waiting for Required docs queue and will evaluate them 
//to see if they are ready to move to the decision ready process
function RTM_processByTask()
{
  debug.log("INFO","Starting FA_ReadyToMove_intool\n");
  //get a list of folders in all FA Waiting for required docs queues with the flag set
  var intool_sql = "SELECT INUSER.IN_WF_ITEM.ITEM_ID, " +
                   "  INUSER.IN_WF_ITEM.OBJ_ID, " +
                   "  INUSER.IN_WF_QUEUE.QUEUE_NAME, " +
                   "  INUSER.IN_PROJ.PROJ_NAME, " +
                   "  INUSER.IN_PROJ_TYPE.PROJ_TYPE_NAME " +
                   "FROM INUSER.IN_WF_ITEM " +
                   "INNER JOIN INUSER.IN_WF_QUEUE " +
                   "ON INUSER.IN_WF_ITEM.QUEUE_ID = INUSER.IN_WF_QUEUE.QUEUE_ID " +
                   "INNER JOIN INUSER.IN_PROJ " +
                   "ON INUSER.IN_WF_ITEM.OBJ_ID = INUSER.IN_PROJ.PROJ_ID " +
                   "INNER JOIN INUSER.IN_INSTANCE " +
                   "ON INUSER.IN_INSTANCE.INSTANCE_ID = INUSER.IN_PROJ.INSTANCE_ID " +
                   "INNER JOIN INUSER.IN_INSTANCE_PROP " +
                   "ON INUSER.IN_INSTANCE.INSTANCE_ID = INUSER.IN_INSTANCE_PROP.INSTANCE_ID " +
                   "INNER JOIN INUSER.IN_PROP " +
                   "ON INUSER.IN_PROP.PROP_ID = INUSER.IN_INSTANCE_PROP.PROP_ID " +
                   "INNER JOIN INUSER.IN_PROJ_TYPE " +
                   "ON INUSER.IN_PROJ.PROJ_TYPE_ID = INUSER.IN_PROJ_TYPE.PROJ_TYPE_ID " +
                   "WHERE INUSER.IN_WF_QUEUE.QUEUE_NAME LIKE '%Waiting for Required Docs)' " +
                   "AND INUSER.IN_PROP.PROP_NAME = 'Ready to Move' " +
                   "AND INUSER.IN_INSTANCE_PROP.NUMBER_VAL = 1;";
  var returnVal; 
  var cur = getHostDBLookupInfo_cur(intool_sql,returnVal);

  if(!cur || cur == null)
  {
    debug.log("Error","Can't get list of CPs.\n");
    return false;
  }

  var wfID = "";
  var objID = "";
  var queueName = "";
  var folderName = "";
  var folderType = "";

  while (cur.next()) {
    wfID = cur[0];
    objID = cur[1];
    queueName = cur[2];
    folderName = cur[3];
    folderType = cur[4];
  }

}// end RTM_processByTask

function setRTMFlag()
{

}// end setRTMFlag

function checkRouteReady()
{
  /*for a given folder


  */
}// end checkRouteReady

function routeToReview()
{
  debug.log("INFO","Attempting to load YAML\n");
  loadYAMLConfig(imagenowDir6+"\\script\\config_scripts\\FA_ReadyToMove\\");
}// End routeToReview

function failureEmail()
{

}// end failureEmail

//