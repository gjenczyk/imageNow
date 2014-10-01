/********************************************************************************
        Name:          UA_PriotityRouting
        Author:        Gregg Jenczyk
        Created:        7/2/2014
        Last Updated:   
        For Version:    
---------------------------------------------------------------------------------
        Summary:
               This script sends folders to the appropriate decision ready queue
               from the pre-decision process without the need for 
               
        Mod Summary:
               Date-Initials: Modification description.



********************************************************************************/

// ********************* Include additional libraries *******************
//#link "inxml"    //XML parser
//#link "sedbc"    //Database object
//#link "secomobj" //COM object
#include "$IMAGENOWDIR6$\\script\\lib\\iScriptDebug.jsh"
#include "$IMAGENOWDIR6$\\script\\lib\\HostDBLookupInfo.jsh"
#include "$IMAGENOWDIR6$\\script\\lib\\RouteItem.jsh"
#include "$IMAGENOWDIR6$\\script\\lib\\UA_Profile_Config.jsh"
#include "$IMAGENOWDIR6$\\script\\lib\\UA_DetermineQueue.jsh"
#include "$IMAGENOWDIR6$\\script\\lib\\GetProp.jsh"
#include "$IMAGENOWDIR6$\\script\\lib\\commonSharedFunction.jsh"

// *********************         Configuration        *******************

// logging
#define LOG_TO_FILE         true    // false - log to stdout if ran by intool, true - log to inserverXX/log/ directory
#define DEBUG_LEVEL         5       // 0 - 5.  0 least output, 5 most verbose
#define SPLIT_LOG_BY_THREAD false   // set to true in high volume scripts when multiple worker threads are used (workflow, external message agent, etc)
#define MAX_LOG_FILE_SIZE   100     // Maximum size of log file (in MB) before a new one will be created


// *********************       End  Configuration     *******************

// ********************* Initialize global variables ********************

/**
* Main body of script.
* @method main
* @return {Boolean} True on success, false on error.
*/
function main ()
{
        try
        { 
           debug = new iScriptDebug("USE SCRIPT FILE NAME", LOG_TO_FILE, DEBUG_LEVEL);
           debug.log("WARNING", "UA_PriotityRouting script started.\n");

           var wfItem = new INWfItem(currentWfItem.id);

           if(!wfItem.getInfo())
           {
              debug.log("ERROR", "Couldn't get info for wfItem: %s Reason: %s\n", wfItem, getErrMsg());
              return false;
           }

           var priorityQueue = wfItem.queueName
           var campus = priorityQueue.substring(0,3);

           debug.log("INFO","Processing [%s]\n", wfItem.objectId);

           if (wfItem.type != "2")
           {
              debug.log("ERROR", "Inbound item is not a folder.\n");
              wfItem.setState(WfItemState.Idle, "Item is not a folder");
              return false;
           }

           var folder = new INFolder(wfItem.objectId); 

           if (!folder.getInfo())
           {
            debug.log("ERROR", "Couldn't get info for folder [%s], Reason: [%s]\n", folder.name, getErrMsg());
            wfItem.setState(WfItemState.Idle, "Unable to get folder info");
            return false;
           }

/*           var folderCPs = folder.getCustomProperties();

           for (i=0; i < folderCPs.length; i++)
           {
              debug.log("INFO", "[%s] [%s]\n", folderCPs[i].name, folderCPs[i].getValue())
           }
*/
          debug.log("INFO","Determing queue for [%s]\n", folder.name);

          FindQueue(folder, campus, 1)
        
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


//