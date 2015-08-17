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
#include "$IMAGENOWDIR6$\\script\\lib\\RouteItem.jsh"
#include "$IMAGENOWDIR6$\\script\\lib\\UnifiedPropertyManager.jsh"
#include "$IMAGENOWDIR6$\\script\\lib\\GetProp.jsh"
#include "$IMAGENOWDIR6$\\script\\lib\\SetProp.jsh"

// *********************         Configuration        *******************

// logging
#define LOG_TO_FILE         true    // false - log to stdout if ran by intool, true - log to inserverXX/log/ directory
#define DEBUG_LEVEL         5       // 0 - 5.  0 least output, 5 most verbose
#define SPLIT_LOG_BY_THREAD false   // set to true in high volume scripts when multiple worker threads are used (workflow, external message agent, etc)
#define MAX_LOG_FILE_SIZE   100     // Maximum size of log file (in MB) before a new one will be created


// *********************       End  Configuration     *******************

// ********************* Initialize global variables ********************
var CALLED_BY_EM = false;
var sql = "";
var TEST = "321YZ78_07B9M06W0000017";
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

      var externalMsgObj;
      var params = getInputParams();
      
      if(null != params[0]) 
      {
          debug.log("INFO","Working with folder: [%s]\n",params[0]);
          RTM_processByTrigger(params[0]);
      }
      else if((externalMsgObj = getInputPairs()) != undefined) 
      {

        setSuccess(false);
        var oneOff = source["FolderId"];
        if(RTM_processByTrigger(oneOff))
        {
          setSuccess(true);
        }
      }
      else
      {
         RTM_processByTask()
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

//This function will receive informaiton sufficient to find a folder.  Return true or 
//false to calling script. 
function RTM_processByTrigger(fldrID)
{
  debug.log("WARNING","Starting RTM_processByTrigger for [%s].\n", fldrID);
  var successStatus = true;
  //make sure there is a folder for the trigger passed
  try
  {
    //make a folder object
    var fldr = new INFolder(fldrID);
    //make sure we can get info about the folder
    if(!fldr.getInfo())
    {
      debug.log("ERROR","Could not get info about [folder] - [%s]\n",getErrMsg());
      throw "Could not get folder info.";
    }
    //get workflow information about the folder
    var wfFldr = fldr.getWfInfo();
    if(!wfFldr || wfFldr == null)
    {
      debug.log("ERROR","Folder is not currently in workflow.");
      throw "Folder is not in workflow.";
    } 
    //set the flag on a folder for the id passed
    if(!setRTMFlag(fldr))
    {
      //could not set the flag
      debug.log("ERROR","Could not set RTM flag for [%s]\n", fldr);
      throw "Could not set RTM flag";
    }
    //check to see if the folder is ready to move
    var folder = {wfID:wfFldr[0].id, objID:fldr.id, queue:wfFldr[0].queueName, name:fldr.name, type:fldr.folderTypeName};
    printf(folder.wfID+" - "+folder.objID+" - "+folder.queue+" - "+folder.name+" - "+folder.type+" - \n")
    if (!checkRouteReady(folder))
    {
      //found documents in the NDE queues
      debug.log("INFO","[%s] is not ready to move.\n", fldr);
      throw "Folder is not ready to move"
    }
    else
    {
      //folder is ready to move
      debug.log("INFO","[%s] is ready to move.\n", fldr);
      //try to route folder to correct review process
      if(!routeToReview(folder))
      {
        //couldn't route the folder to a review queue
        debug.log("ERROR","Could not route [%s] to review queue - [%s]\n",fldr, getErrMsg());
        throw "Could not route [" + folder.id + "] to review process";
      }
      else
      {
        //routed folder to the review queue
        debug.log("INFO","Routed [%s]\n", fldr);
      }
    }
  }//end try
  catch(miss)
  {
    //handle any errors and send out an alert email
    debug.log("WARNING","Did not move folder: [%s] - [%s]\n",fldr, miss);
    failureEmail();
    successStatus = false;
  }
  finally
  {
    //return the results to the calling script
    return successStatus;
  }
}//end RTM_processByTrigger

//This function will get a list of all folders in the Waiting for Required docs queue and will evaluate them 
//to see if they are ready to move to the decision ready process
function RTM_processByTask()
{
  debug.log("WARNING","Starting RTM_processByTask.\n");
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
  //if there are no folders with the RTM flag or if we couldn't get a list
  if(!cur || cur == null)
  {
    debug.log("ERROR","Could not retrieve a list of folders to work.\n");
    return false;
  }

  var folders = []
  while (cur.next())
  {

    var folder = {wfID:"", objID:"", queue:"", name:"", type:"", drawer:""};
    folder.wfID = cur[0];
    folder.objID = cur[1];
    folder.queue = cur[2];
    folder.name = cur[3];
    folder.type = cur[4];
    folder.drawer = cur[5];

    folders.push(folder);
  }

  //need to do it this way because it doesn't like multiple connections to the db
  for (var a = 0; a < folders.length; a++)
  {
    //check to see if folder is ready to move
    debug.log("DEBUG","Checking to see if [%s] is ready to move to review process.\n", folders[a].name);

    if(!checkRouteReady(folders[a]))
    {
      debug.log("INFO","[%s] is not ready for review.\n", folders[a].name);
      continue;
    }

    debug.log("INFO","No documents found for [%s] in review queue - preparing to route to working process.\n", folders[a].name);
    if(!routeToReview(folders[a]))
    {
      debug.log("ERROR","Unable to route [%s]\n", folders[a].name);
      failureEmail();
      continue;
    }
  }

}//end RTM_processByTask

function setRTMFlag(folder)
{
  //get info about the folder
  debug.log("DEBUG","Setting Ready to Move to true for [%s]\n", folder.id);
  var rtmCP = "Ready to Move";
  var rtm = GetProp(folder, rtmCP);
  if(rtm != "true")
  {
    if(!SetProp(folder, rtmCP, true))
    {
      debug.log("ERROR","Unable to set Ready to Move flag for [%s]: [%s]\n", folder.id, getErrMsg());
      return false;
    }
    else
    {
      debug.log("DEBUG","Successfully set RtM flag to true for [%s]\n", folder.id);
      return true;
    }
  }
  else
  {
    debug.log("DEBUG","Ready to Move flag has already been set for [%s]\n", folder.id);
    return true;
  }
  
}//end setRTMFlag

function checkRouteReady(folder)
{
  var fldr = new INFolder(folder.objID)
  var inProps = ["Student ID"]; // "Lifetime Doc","Process"
  var upm = new UnifiedPropertyManager();
  var opPropValues = upm.GetAllProps(fldr, inProps);
  var studentID = opPropValues[0];
  var campusCode = folder.type.slice(-3);
  var process = "";

  var prc_sql = "SELECT ISCRIPTUSER.PROCESS_DETAILS_X.PROC_CODE " +
  " FROM ISCRIPTUSER.PROCESS_DETAILS_X " +
  "WHERE ISCRIPTUSER.PROCESS_DETAILS_X.FOLDERTYPE = '" + folder.type.substring(0,folder.type.length-4) + "';";
  var returnVal;
  var cur = getHostDBLookupInfo_cur(prc_sql,returnVal);
  if(!cur || cur == null)
  {
    debug.log("ERROR","Could not find process configuration for [%s]\n", folder.type);
    return false;
  }
  process = cur[0];

  debug.log("DEBUG","[%s] Campus Code: [%s] Student ID: [%s] Process Code: [%s]\n", folder.type, campusCode, studentID, process);

  var rr_sql = "SELECT INUSER.IN_WF_QUEUE.QUEUE_NAME, " +
  "INUSER.IN_DOC.DOC_ID, " +
  "INUSER.IN_DOC_TYPE.DOC_TYPE_NAME " +
  "FROM INUSER.IN_WF_ITEM " +
  "INNER JOIN INUSER.IN_WF_QUEUE " +
  "ON INUSER.IN_WF_QUEUE.QUEUE_ID = INUSER.IN_WF_ITEM.QUEUE_ID " +
  "INNER JOIN INUSER.IN_DOC " +
  "ON INUSER.IN_WF_ITEM.OBJ_ID = INUSER.IN_DOC.DOC_ID " +
  "INNER JOIN INUSER.IN_INSTANCE " +
  "ON INUSER.IN_INSTANCE.INSTANCE_ID = INUSER.IN_DOC.INSTANCE_ID " +
  "INNER JOIN INUSER.IN_DOC_TYPE " +
  "ON INUSER.IN_DOC.DOC_TYPE_ID = INUSER.IN_DOC_TYPE.DOC_TYPE_ID " +
  "INNER JOIN ISCRIPTUSER.DOC_TYPE_X " +
  "ON INUSER.IN_DOC_TYPE.DOC_TYPE_NAME = ISCRIPTUSER.DOC_TYPE_X.DOC_TYPE_NAME " +
  "INNER JOIN ISCRIPTUSER.DI_DOCT_PROCESS_X " +
  "ON ISCRIPTUSER.DOC_TYPE_X.DOC_TYPE_CODE = ISCRIPTUSER.DI_DOCT_PROCESS_X.DOC_TYPE_CODE " +
  "WHERE INUSER.IN_WF_QUEUE.QUEUE_NAME LIKE '%" + campusCode + " Review for Completeness)' " +
  "AND INUSER.IN_DOC.FOLDER = '" + studentID + "' " +
  "AND PROC_CODE = '" + process + "';";

  returnVal = ""; 
  cur = getHostDBLookupInfo_cur(rr_sql,returnVal);
  //if there are no folders with the RTM flag or if we couldn't get a list
  if(!cur || cur == null)
  {
    debug.log("WARNING","No matching documents found for [%s].  Cleared for take-off!\n", folder.name);
    return true;
  }
  else
  {
    debug.log("INFO","Found documents belonging to [%s] for process code [%s]. Holding for now...\n", folder.name, process);
    while(cur.next())
    {
      debug.log("DEBUG","Queue: [%s], docID: [%s]\n", cur[0], cur[1]);
    }
    
    //let the other function know we can't route yet
    return false;
  }
}//end checkRouteReady

function routeToReview(folder)
{
  debug.log("INFO","Attempting to load YAML\n");
  loadYAMLConfig(imagenowDir6+"\\script\\config_scripts\\FA_ReadyToMove\\");
  var foundConfig = false;
  var desQ = '';
  var RTM_CONFIG;

  for (var process_config in CFG.FA_ReadyToMove)
  {
    if(foundConfig)
    {
      break;
    }
    RTM_CONFIG = CFG.FA_ReadyToMove[process_config].PROCESS_CONFIG;
    debug.log("DEBUG","Begin processing RTM configuration for [%s]\n", folder.type);
    
    for (var b = 0; b < RTM_CONFIG.length; b++)
    {
      var match = folder.type.indexOf(RTM_CONFIG[b].FOLDER_TYPE)
      if (match == 0)
      {
        debug.log("INFO","Found matching folder type. Applying configuration [%s] [%s]...\n", RTM_CONFIG[b].FOLDER_TYPE, RTM_CONFIG[b].PROCESS_QUEUE)
        desQ = folder.type.slice(-3) + " " + RTM_CONFIG[b].PROCESS_QUEUE; 
        foundConfig = true;
        break;
      }
    }
  } // end of for each process_config

  if(desQ == '')
  {
    debug.log("ERROR","Could not find routing configuraiton for [some bs]\n");
    return false;
  }

  debug.log("INFO","Found target queue [%s]  for [%s].  Preparing to route...\n", desQ, folder.name);
  var fldr = new INWfItem(folder.wfID);
  if(!fldr.getInfo())
  {
    debug.log("ERROR","Could not get info for folder: [%s] [%s]", folder.id, getErrMsg());
    return false;
  }

//check to see if it's in wf
  //printf(fldr + " " + fldr.id + " " + fldr.getInfo())
  var rsn = "FA_ReadyToMove: Routing to " + desQ;
  if(!RouteItem(fldr,desQ,rsn))
  {
    debug.log("ERROR","Unable to route folder!\n");
    return false;
  }
  else
  {
    debug.log("INFO","Routed the folder!\n");
    return true;
  }
  //return true;
}//end routeToReview

function failureEmail()
{
  //send an email with debugging info if we couldn't set the flag
}//end failureEmail

//