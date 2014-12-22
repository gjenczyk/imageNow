/********************************************************************************
        Name:          updateFolderAndDocCPs
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

// *********************         Configuration        *******************

// logging
#define LOG_TO_FILE         true    // false - log to stdout if ran by intool, true - log to inserverXX/log/ directory
#define DEBUG_LEVEL         5       // 0 - 5.  0 least output, 5 most verbose
#define SPLIT_LOG_BY_THREAD false   // set to true in high volume scripts when multiple worker threads are used (workflow, external message agent, etc)
#define MAX_LOG_FILE_SIZE   100     // Maximum size of log file (in MB) before a new one will be created


// *********************       End  Configuration     *******************

// ********************* Initialize global variables ********************

workingCsv = imagenowDir6 + "\\script\\dartb.csv"

cpsToUpdate = {
  "Notification Plan",
  "Program Action",
  "Program Action Reason"
};

folderType = "Application Undergraduate Admissions DUA";

/**
* Main body of script.
* @method main
* @return {Boolean} True on success, false on error.
*/
  function main ()
{
    try
    {
       debug = new iScriptDebug("updateFolderAndDocCPs", LOG_TO_FILE, DEBUG_LEVEL);
       debug.log("WARNING", "updateFolderAndDocCPs script started.\n");

      var csvList = Clib.fopen(workingCsv,"r");
      /* check to  */
      if ( csvList == null )
      {
        debug.log("ERROR","csv is either missing or empty.\n");
        return false;
      } /* end if ( csvList == null ) */

      //update based on csv line
      var line;
      var splitLine;
      var emplid;
      var appNo;
      var folder;
      var notificPlan;
      var progAct;
      var progActReason;
      var workingVal;

      while ( null != (line=Clib.fgets(csvList)) )
      {
        line = line.replace(/\r?\n|\r/, "");
        splitLine = line.split(",");

        emplid = splitLine[2];
        appNo = splitLine[3];
        notificPlan = splitLine[9];
        progAct = splitLine[7];
        progActReason = splitLine[8];

        folderName = emplid + " APP: " + appNo;
        folder = new INFolder(folderName, folderType);

        if (!folder.getInfo())
        {
          debug.log("ERROR", "Failed to get information for folder.  Error: %s\n", getErrMsg());
          continue;
        }
        else
        {
          debug.log("INFO", "Updating contents of [%s]\n", folderName);
          //update folder CPs
          var folderCPs = folder.getCustomProperties();

          if (!folderCPs || folderCPs == null)
          {
            debug.log("ERROR","Could not retrieve custom props for [%s]. Folder not processed.\n", folderName);
            continue;
          }

          //go through folder CPs and find the ones we're looking for
          for (var i=0; i<folderCPs.length; i++)
          {
            if (contains(cpsToUpdate, folderCPs[i].name))
            {  
              workingVal = customSelector(folderCPs[i].name, notificPlan, progAct, progActReason);

              debug.log("INFO","Property: [%s] Current Value: [%s] Future Value: [%s]\n", folderCPs[i].name, folderCPs[i].getValue(), workingVal);

              if (!folderCPs[i].setValue(workingVal))
              {
                debug.log("ERROR","Can't store prop value! [%s] [%s]\n", folderCPs[i].name, workingVal);
              }
            }
          } // end for (var i=0; i<folderCPs.length; i++)

          if (!folder.setCustomProperties(folderCPs))
          {
            debug.log("ERROR", "Can't set prop values on [%s]!\n", folderName);
          }

          //get docs in folder
          var folderDocs = folder.getDocList();

          for (var j=0; j<folderDocs.length; j++)
          {
            var doc = new INDocument(folderDocs[j].id);
            if (!doc.getInfo())
            { 
              debug.log("ERROR", "Failed to get information for [%s].  Error: %s\n", doc.id, getErrMsg());
              continue;
            }
            else
            {
              //update the documents in the folder
              var docCPs = doc.getCustomProperties();
              for (var k=0; k<docCPs.length; k++)
              {
                if (contains(cpsToUpdate, docCPs[k].name))
                {  
                  workingVal = customSelector(docCPs[k].name, notificPlan, progAct, progActReason);
                  debug.log("INFO","Property: [%s] Current Value: [%s] Future Value: [%s]\n", docCPs[k].name, docCPs[k].getValue(), workingVal);

                  if (!docCPs[k].setValue(workingVal))
                  {
                    debug.log("ERROR","Can't store prop value! [%s] [%s]\n", docCPs[k].name, workingVal);
                  }
                }
              } // end for (var k=0; k<docCPs.length; k++)

              if (!doc.setCustomProperties(docCPs))
              {
                debug.log("ERROR", "Can't set prop values on [%s]!\n", doc.id);
              }
            } //end working the document
          } // end iterating through the folder contents
        } // end working the folder and contents
      } // end while ( null != (line=Clib.fgets(csvList)) )
      Clib.fclose(csvList);
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
/* check to see if an array contains a string*/
function contains(a, obj) {
    for (var i = 0; i < a.length; i++) {
        if (a[i] === obj) {
            return true;
        }
    }
    return false;
}

/* there should be a better way of figuring this out, but this will let you 
know which CP you're working with */
function customSelector(name, val1, val2, val3) {

  if (name == "Notification Plan")
  {
    return val1;
  }
  else if (name == "Program Action")
  {
    return val2;
  }
  else if (name == "Program Action Reason")
  {
    return val3;
  }
  else
  {
    return false;
  }
}

//