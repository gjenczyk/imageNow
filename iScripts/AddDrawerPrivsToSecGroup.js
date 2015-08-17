/********************************************************************************
        Name:           AddDrawerPrivsToSecGroup.js
        Author:         Gregg Jenczyk
        Created:        05/02/2014
        Last Updated:
        For Version:
---------------------------------------------------------------------------------
        Summary:
               This script adds drawer permissisons to security groups.

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
#define TEST_RUN            false    // Setting this to true returns the list of security groups that will be changed.
// *********************       End  Configuration     *******************

// ********************* Initialize global variables ********************
var debug;

var powerGroups = [
//      "DRW_UMLFA_All",
//      "DRW_UMDFA_Basic",
//      "DRW_UMBFA All",
//      "DRW_UMBFA_Basic",
//      "DRW_UMLFA_All",
//      "DRW_UMLFA_Basic"
//        "DRW_UMDFA_Delete"
//        "DRW_UMLFA_Delete"
//        "DRW_UMBFA_Delete"
];

var regGroups =[
      "DRW_UMLFA_Opn"
];

var globalPrivs = [
     "SEARCH_CONTENTS=1"
];
var drawerPrivs = [
  "DOC_VIEW=1",
  "CUSTOM_PROP_MODIFY=1",
  "DOC_MODIFY_KEYS=1",
  "DOC_MODIFY_NOTES=1"
];


/**
* Main body of script.
* @method main
* @return {Boolean} True on success, false on error.
*/
function main ()
{
        try
        {
            debug = new iScriptDebug("AddDrawerPrivsToSecGroup.js", LOG_TO_FILE, DEBUG_LEVEL);
            debug.log("WARNING", "AddDrawerPrivsToSecGroup.js starting.\n");

            //securty groups, privs, priv type (currently just GLOBAL or DRAWER)
            processPrivs(regGroups,globalPrivs, "GLOBAL");
            processPrivs(regGroups,drawerPrivs, "DRAWER");

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

function processPrivs(grouping, privArr, type)
{
  //define drawer here, but this could be dynamic via regex
  var privDrawer = "UMLFA";
  if(!type || type == null)
  {
    debug.log("ERROR","Priv type is not defined.\n");
    return false;
  }
  if(privArr.length < 1)
  {
    debug.log("ERROR","No privs have been passed for [%s]\n", type);
    return false;
  }
  var privs = new Array()
  for (var i = 0; i < privArr.length; i++)
  {
    privs.push(privArr[i]);
  }

  debug.log("DEBUG","privs: [%s]\n", privs);

  if(type == "GLOBAL")
  {
    privDrawer = null;
  }
  for (i=0; i < grouping.length; i++)
  {
  	//Enter Drawer below
  	debug.log("DEBUG","Adding [%s] to [%s]\n", privs, grouping[i]);
    setPrivs(grouping[i],privDrawer,privs);
  }

} // end processPrivs


function setPrivs(secGrp,drawer,privArray)
{
  if(!drawer || drawer == null)
  {
    debug.log("INFO","Setting global privs.\n")
    if (!INPriv.setGlobalPrivs(secGrp,privArray))
    {
      debug.log("ERROR","Failed to set global privs: [%s]\n", getErrMsg());
    }
    else
    {
      debug.log("INFO","Privlige set for [%s]\n", privArray, secGrp);
    }
  }
  else
  {
    debug.log("INFO","Setting drawer privs.\n")
    if (!INPriv.setDrawerPrivs(secGrp,drawer,privArray))
    {
      debug.log("ERROR","Failed to set drawer privs: [%s]\n", getErrMsg());
    }
    else
    {
      debug.log("INFO","Privliges set for [%s]\n", secGrp);
    }

  }
}


//
