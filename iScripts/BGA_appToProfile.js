/********************************************************************************
        Name:          BGA_appToProfile
        Author:        Gregg Jenczyk
        Created:        7/16/2015
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
#include "$IMAGENOWDIR6$\\script\\lib\\GetProp.jsh"
#include "$IMAGENOWDIR6$\\script\\lib\\RouteItem.jsh"

// *********************         Configuration        *******************

// logging
#define LOG_TO_FILE         true    // false - log to stdout if ran by intool, true - log to inserverXX/log/ directory
#define DEBUG_LEVEL         5       // 0 - 5.  0 least output, 5 most verbose
#define SPLIT_LOG_BY_THREAD false   // set to true in high volume scripts when multiple worker threads are used (workflow, external message agent, etc)
#define MAX_LOG_FILE_SIZE   100     // Maximum size of log file (in MB) before a new one will be created


// *********************       End  Configuration     *******************

// ********************* Initialize global variables ********************
var INPUT_DOC = "Application Graduate Admissions";
var SWAP_DOC = "Profile Sheet";
var WF_QUEUE = "UMBGA PostReview Evaluator";
var REMOVE_APP = false;
/**
* Main body of script.
* @method main
* @return {Boolean} True on success, false on error.
*/
  function main ()
{
    try
    {
      debug = new iScriptDebug("BGA_appToProfile", LOG_TO_FILE, DEBUG_LEVEL);
      debug.log("WARNING", "BGA_appToProfile script started.\n");

      var wfItem = new INWfItem(currentWfItem.id);//"301YY4X_04VZ80N4N020G6T");//
      if(!wfItem.id || !wfItem.getInfo())
      {
        debug.log("ERROR", "Couldn't get info for wfItem: %s\n", getErrMsg());
        return false;
      }

      var wfQ = currentWfQueue.name;//"UMBGA PostReview Evaluator"; //
      if(wfQ != WF_QUEUE)
      {
        debug.log("INFO","[%s] is not the correct queue for switching.\n", wfQ);
        return false;
      }

      var wfDoc = new INDocument(wfItem.objectId);
      if(!wfDoc.getInfo())
      {
        debug.log("ERROR", "Couldn't get info for wfDoc: %s\n", getErrMsg());
        return false;
      }
      if(wfDoc.docTypeName != INPUT_DOC)
      {
        debug.log("INFO","[%s] is not the correct doctype for switching.\n", wfDoc.docTypeName);
        return false;
      }

      debug.log("INFO","Working [%s] in [%s]\n", wfDoc.id, wfQ);

      var appNo = GetProp(wfDoc, "SA Application Nbr");

      sql = "SELECT INUSER.IN_DOC.DOC_ID " +
      "FROM INUSER.IN_DOC " +
      "INNER JOIN INUSER.IN_DOC_TYPE " +
      "ON INUSER.IN_DOC_TYPE.DOC_TYPE_ID = INUSER.IN_DOC.DOC_TYPE_ID " +
      "INNER JOIN INUSER.IN_INSTANCE " +
      "ON INUSER.IN_INSTANCE.INSTANCE_ID = INUSER.IN_DOC.INSTANCE_ID " +
      "INNER JOIN INUSER.IN_INSTANCE_PROP " +
      "ON INUSER.IN_INSTANCE.INSTANCE_ID = INUSER.IN_INSTANCE_PROP.INSTANCE_ID " +
      "INNER JOIN INUSER.IN_PROP " +
      "ON INUSER.IN_PROP.PROP_ID = INUSER.IN_INSTANCE_PROP.PROP_ID " +
      "INNER JOIN INUSER.IN_DRAWER " +
      "ON INUSER.IN_DRAWER.DRAWER_ID = INUSER.IN_DOC.DRAWER_ID " +
      "WHERE INUSER.IN_DOC.FOLDER = '" + wfDoc.field1 + "' " +
      "AND INUSER.IN_INSTANCE_PROP.STRING_VAL = '" + appNo + "' " +
      "AND INUSER.IN_DRAWER.DRAWER_NAME = '" + wfDoc.drawer + "' " +
      "AND INUSER.IN_DOC_TYPE.DOC_TYPE_NAME = '" + SWAP_DOC + "' " +
      "AND INUSER.IN_INSTANCE.DELETION_STATUS = '0';";

      var returnVal; 
      var cur = getHostDBLookupInfo_cur(sql,returnVal);
      //if there are no folders with the RTM flag or if we couldn't get a list
      if(!cur || cur == null)
      {
        debug.log("ERROR","Could not retrieve a list of folders to work.\n");
        return false;
      }

      var results = 0;
      while (cur.next())
      {
        results+=1;
        debug.log("DEBUG","Profile Sheet [%s]\n", cur[0]);
      }
      if(results > 1)
      {
        debug.log("ERROR","More than one profile sheet found. Not switching documents.\n");
        return false;
      }

      cur.reload();
      var profileID = cur[0];

      var profile = new INDocument(profileID);
      if(!profile.getInfo())
      {
        debug.log("ERROR", "Couldn't get info for profile: %s\n", getErrMsg());
        return false;
      }

      var profileWF = profile.getWfInfo();
      if(profileWF)
      {
        if(profileWF.length > 0)
        {
          var item = profileWF[0];
          if(item.state == 2)
          {
            debug.log("WARNING","Profile Sheet [%s] is currently open in workflow. Cannot route to [%s]\n", profile.id, WF_QUEUE);
            return false;
          }
          else
          {
            debug.log("INFO","Routing profile sheet [%s] to [%s]\n", profile.id, WF_QUEUE);
            if(!RouteItem(item.id, WF_QUEUE, "BGA_appToProfile"))
            {
              debug.log("ERROR","Cannot add profile to [%s]\n", WF_QUEUE);
              return false;
            }
            else
            {
              debug.log("INFO","Successfully routed [%s] to [%s].\n", profile.id, WF_QUEUE);
              REMOVE_APP = true;
            }
          }
        }
        else
        {
          debug.log("DEBUG","[%s] is not currently in workflow.\n",profile.id);
          var revQ = new INWfQueue("",WF_QUEUE);
          if(!revQ.createItem(WfItemType.Document, profile.id, WfItemPriority.Medium))
          {
            debug.log("ERROR","Cannot add profile to [%s]\n", WF_QUEUE);
            return false;
          }
          else
          {
            debug.log("INFO","Successfully added [%s] to [%s].\n", profile.id, WF_QUEUE);
            REMOVE_APP = true;
          }
        }

      }
      else
      {
        debug.log("ERROR","No workflow items found: %s.\n", getErrMsg());
        return false;
      }

      if(REMOVE_APP)
      {
        if(!wfItem.archive())
        {
            debug.log("ERROR","Failed to archive application [%s]: %s.\n", wfDoc.id, getErrMsg());
        }
        else
        {
          debug.log("INFO","Successfully archived application [%s].\n", wfDoc.id);
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


//