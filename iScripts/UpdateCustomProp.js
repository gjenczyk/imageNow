/********************************************************************************
        Name:          UpdateCustomProp.js
        Author:        Gregg Jenczyk
        Created:        10/16/14
        Last Updated:   
        For Version:    
---------------------------------------------------------------------------------
        Summary:
               This script will update a custom property based on 
               
        Mod Summary:
               Date-Initials: Modification description.
               
********************************************************************************/

// ********************* Include additional libraries *******************
//#link "inxml"    //XML parser
//#link "sedbc"    //Database object
//#link "secomobj" //COM object
#include "$IMAGENOWDIR6$\\script\\lib\\iScriptDebug.jsh"
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
var PROP_TO_CHANGE = "Program Code";


/**
* Main body of script.
* @method main
* @return {Boolean} True on success, false on error.
*/
function main ()
{
        try
        { 
            debug = new iScriptDebug("UpdateCustomProp", LOG_TO_FILE, DEBUG_LEVEL);
            debug.log("WARNING", "UpdateCustomProp.js starting.\n");

            /* open the csv with the doc ids */
            var csvList = Clib.fopen(imagenowDir6+"\\script\\bga_docs.csv","r");

            /* check to  */
            if ( csvList == null )
            {
              debug.log("ERROR","csv is either missing or empty.\n");
              return false;
            } /* end if ( csvList == null ) */

            while ( null != (line=Clib.fgets(csvList)) )
            {
              /* remove newline character from line */
              var curDoc = line.substring(0,line.length-1);

              var doc = new INDocument(curDoc);

              if(!doc.getInfo())
              {
                debug.log("ERROR","Couldn't get info for doc id: [%s], [%s]\n", doc.id, getErrMsg());
                continue;
              } /* if(!doc.getInfo()) */

              debug.log("INFO", "Processing document id: [%s]\n", doc.id);
              /* make sure document is not currently being processed in workflow */ 
              var wfCheck = doc.getWfInfo();
              if (wfCheck.length > 0)
              {
                var wfItem = wfCheck[0];
                
                if(wfItem.state == 2)
                {
                  /* move to next on list if someone has the document open right now */
                  debug.log("ERROR","Document [%s] is currently being processed in workflow. Custom properties will not be updated\n", doc.id);
                  continue;
                }/* end if(wfItem.state == 2) */

              }/* end if (wfCheck.length > 0) */
                
              var props = doc.getCustomProperties();

              if (props)
              { 
                for (var i = 0; i < props.length; i++)
                {
                  /* only look at the desired custom property */
                  if (props[i].name == PROP_TO_CHANGE)
                  {
                    var curVal = props[i].getValue();
                    var newVal = customTranslator(curVal);
                    
                    if (newVal == false)
                    {
                      debug.log("ERROR","Configuration doesn't exist for [%s].  Moving on to next document\n", curVal);
                      continue;
                    } /* end if (newVal == false) */

                    debug.log("INFO","Value of [%s] before change is [%s], attempting to change to [%s]\n", PROP_TO_CHANGE, curVal, newVal);

                    /* stage the new value of the custom property */
                    if(!props[i].setValue(newVal))
                    {
                      debug.log("ERROR","Could not stage [%s] to [%s] for [%s].  Error: [%s]\n", newVal, PROP_TO_CHANGE, doc.id, getErrMsg());
                      continue
                    }

                    /* attempt to set the custom property to the new value */
                    if (!doc.setCustomProperties(props))
                    {
                      debug.log("ERROR","Could not set custom properties for [%s] Error: [%s]\n", doc.id, getErrMsg());
                      continue;
                    }
                    else 
                    {
                      debug.log("INFO","Successfully updated [%s] to [%s] for [%s]\n", PROP_TO_CHANGE, newVal, doc.id);
                    }
                    /* end setting of custom porpterty */

                  } /* end if (props[i].name == PROP_TO_CHANGE) */

                } /* end for (var i = 0; i < props.length; i++) */ 

              } /* end if (props) */
                
            } /* end  while ( null != (line=Clib.fgets(csvList)) ) */

          Clib.fclose(csvList);

        } /* end of try */
        
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
} // end of main

// ********************* Function Definitions **********************************
/************************************************************************
  Name: customTranslator

  Purpose: This function will return a desired value based on the input 
            fed to it.

  Parameter: inValue - the raw value to check for

  Output: outValue - the switch value if found, else false

************************************************************************/
function customTranslator(inValue)
{
  var outValue;

  switch(inValue)
  {
    case "Science & Mathematics - Grad":
      outValue = "CSM-G";
      break;
    case "Education and Human Devlopment":
      outValue = "GCE-G";
      break;
    case "Nursing & Health Sciences":
      outValue = "NUR-G";
      break;
    case "Advancing&Professional Studies":
      outValue = "CAPSG";
      break;
    case "Grad Sch Policy&Global Studies":
      outValue = "SPS-G";
      break;
    case "Management - Graduate":
      outValue = "MGT-G";
      break;
    case "Liberal Arts - Graduate":
      outValue = "LA-G";
      break;
    case "School of Global Incl&Soc Dev":
      outValue = "GISDG";
      break;
    case "CPCS - Graduate":
      outValue = "CPCSG";
      break;      
    default:
      outValue = false;
  }

  return outValue;
}

//