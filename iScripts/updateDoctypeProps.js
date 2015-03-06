/********************************************************************************
        Name:          updateDoctypeProps
        Author:        Gregg Jenczyk
        Created:        03/06/2015
        Last Updated:   
        For Version:    
---------------------------------------------------------------------------------
        Summary:
               This script will change the custom properties on a doctype.  It can
               be used to add, remove, or reorder all CPs on a given doctype.
               
        Mod Summary:
               Date-Initials: Modification description.



********************************************************************************/

// ********************* Include additional libraries *******************
//#link "inxml"    //XML parser
//#link "sedbc"    //Database object
//#link "secomobj" //COM object
#include "$IMAGENOWDIR6$\\script\\lib\\iScriptDebug.jsh"
#include "$IMAGENOWDIR6$\\script\\lib\\HostDBLookupInfo.jsh"
#include "$IMAGENOWDIR6$\\script\\lib\\PropertyManager.jsh"

// *********************         Configuration        *******************

// logging
#define LOG_TO_FILE         true    // false - log to stdout if ran by intool, true - log to inserverXX/log/ directory
#define DEBUG_LEVEL         5       // 0 - 5.  0 least output, 5 most verbose
#define SPLIT_LOG_BY_THREAD false   // set to true in high volume scripts when multiple worker threads are used (workflow, external message agent, etc)
#define MAX_LOG_FILE_SIZE   100     // Maximum size of log file (in MB) before a new one will be created


// *********************       End  Configuration     *******************

// ********************* Initialize global variables ********************

DOCTYPES_TO_UPDATE = ["test"]; //all doctypes to work
CPS_TO_ADD = [{name:"",position:""}]; //all CPs to add
CPS_TO_REMOVE = [{name:"Alias"},{name:"Alias 1"}]; //all CPs to remove
DOCTYPE_OK = {name:"DOCTYPE_OK",value:false};
CP_ADD_OK = {name:"CP_ADD_OK",value:false};
CP_REMOVE_OK = {name:"CP_REMOVE_OK",value:false};

/**
* Main body of script.
* @method main
* @return {Boolean} True on success, false on error.
*/
  function main ()
{
    try
    {
      debug = new iScriptDebug("updateDoctypeProps", LOG_TO_FILE, DEBUG_LEVEL);
      debug.log("WARNING", "updateDoctypeProps script started.\n");

      //check the configuration for validity.  Don't do anything if the configured information is invalid
      if (!checkConfig())
      {
        debug.log("ERROR","One or more configured items are not valid.\n");
        return false;
      }// end if (!checkCPConfig())

      debug.log("DEBUG","CP_ADD_OK is [%s] and CP_REMOVE_OK [%s] and DOCTYPE_OK is [%s]\n", CP_ADD_OK.value, CP_REMOVE_OK.value, DOCTYPE_OK.value);

      if((!CP_ADD_OK.value && !CP_REMOVE_OK.value)|| !DOCTYPE_OK.value)
      {
        debug.log("ERROR","There is nothing to update!\n");
        return false;
      }

      //If we're adding or removing
      if ((CP_ADD_OK.value || CP_REMOVE_OK.value) && DOCTYPE_OK.value)
      {
        for (var i = 0; i < DOCTYPES_TO_UPDATE.length; i ++)
        {
          debug.log("INFO","Working the [%s] doctype.\n",DOCTYPES_TO_UPDATE[i]);
          //get info aobut the doctype
          var docTypeName = DOCTYPES_TO_UPDATE[i];
          var docType = new INDocType();
          var propArr = [];
          docType.name = docTypeName;
          if (docType.getInfo())
          {
            debug.log("DEBUG","Doctype id:%s, name:%s, desc:%s, isActive:%s\n", docType.id, docType.name, docType.desc, docType.isActive);
            var props = docType.props;
            debug.log("DEBUG","Number of custom properties: %d\n",props.length);
            for (i=0; i<props.length; i++)
            {
                debug.log("INFO","Doctype CPs: id:%s, name:%s, isRequired:%s \n",props[i].id, props[i].name, props[i].isRequired);
                propArr.push(props[i]);
            }//end for (i=0; i<props.length; i++)
          }
          else
          {
            debug.log("ERROR","Failed to retrieve info for document type - Error: %s\n.", getErrMsg());
            return false;
          }//end if (docType.getInfo()) else

          //if adding values
          if(CP_ADD_OK.value)
          {
            debug.log("INFO","Adding CPs to [%s]\n",docType.name);
            //do something here to update the document with new valuss
          }//end if adding values

          //if removing values
          if(CP_REMOVE_OK.value)
          {
            debug.log("INFO","Removing CPs from [%s]\n",docType.name);
            for (var l = 0; l < propArr.length; l++)
            {
              for (var m = 0; m < CPS_TO_REMOVE.length; m++)
              {
                if(propArr[l].name == CPS_TO_REMOVE[m].name)
                {
                  propArr.splice(l,1);
                }
              }
            }//end for (var l = 0; l < propArr.length; l++)

            for (i=0; i<propArr.length; i++)
            {
              debug.log("DEBUG","NEW DOCTYPE PROPS: id:%s, name:%s, isRequired:%s \n",propArr[i].id, propArr[i].name, propArr[i].isRequired);
            } //end for (i=0; i<propArr.length; i++)

            if(docType.update(docType.name, docType.desc, docType.isActive, 0, propArr))
            {
              debug.log("INFO","Successfully updated document type [%s].\n", docType.name);
            } 
            else
            {
              debug.log("ERROR","Aborting - Failed to update document type - %s\n.", getErrMsg());
              return false;
            }
          }//end if removing values
        }//end of for (var i = 0; i < DOCTYPES_TO_UPDATE.length; i ++)
      }//if (CP_ADD_OK || CP_REMOVE_OK) 
    }// end try
        
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
    }//end catch
    
    finally
    {
           if (debug) debug.finish();
           return;
    }
}//end main

// ********************* Function Definitions **********************************
//function that checks to see if the values in the CP config actually exist
function checkConfig()
{
  // working with some global vars here
  debug.log("INFO","Checking to see if provided configuration is valid.\n");
  
  //check custom property info
  var sql = "SELECT PROP_NAME FROM IN_PROP;";
  var returnVal; 
  var cur = getHostDBLookupInfo_cur(sql,returnVal);
  var allProps = [];

  if(!cur || cur == null)
  {
    debug.log("Error","Can't get list of CPs.\n");
    return false;
  }
  
  while (cur.next())
  {
    allProps.push(cur[0]);
  }

  //check the cps to add
  var cpAdd = checkCPs(allProps,CPS_TO_ADD,CP_ADD_OK);
  //end check the cps to add
  
  //check the cps to remove
  var cpRem = checkCPs(allProps,CPS_TO_REMOVE,CP_REMOVE_OK);
  //end check the cps to remove

  //check the document types
  var dtToWork = checkDocTypes(DOCTYPES_TO_UPDATE, DOCTYPE_OK);
  //end check the document types

  //return true if it all is OK to proceed
  if (!cpAdd || !cpRem || !dtToWork)
  {  
    return false;
  }
  else
  {
    return true;
  }
}//end checkCPConfig()

// function to check to see if an item is in array since iScript doesn't have indexOf
function arrCheck(arr, item)
{
  var success = false;
  for (var i = 0; i < arr.length; i++)
  {
    if (item == arr[i])
    {
      success = true;
    }
  }
  return success;
}//end arrCheck

//function to check if CPs exist
function checkCPs(allCpArr, workCpArr, cpFlag)
{
  for(var j = 0; j < workCpArr.length; j++)
  {
    if (workCpArr[j].name == "")
    {
      debug.log("WARNING","No value was passed for item [%s] in config.  Not changing any CPs.\n",j)
      cpFlag.value = false;
      break;
    }
    else if (arrCheck(allCpArr,workCpArr[j].name))
    {
      debug.log("INFO","Found a match for [%s].  Setting [%s] to true\n",workCpArr[j].name, cpFlag.name);
      cpFlag.value = true;
    }
    else
    {
      debug.log("ERROR","CP [%s] was not found.  Are you sure it exists?\n", workCpArr[j].name);
      cpFlag.value = false;
      return false;
    }
  }

  return true;
}//end checkCPs

//function to verify doctypes exist
function checkDocTypes(dtArr, dtFlag)
{
  for (var k = 0; k < dtArr.length; k++)
  {
    var curDT = INDocType.get(dtArr[k]);
    if(!curDT || curDT == null)
    {
      debug.log("ERROR","Couldn't get info for [%s] - [%s]\n", dtArr[k], getErrMsg());
      return false;
    }
  }
  dtFlag.value = true;
  debug.log("INFO","Setting [%s] to [%s].\n",dtFlag.name, dtFlag.value);
  return true;
}//end checkDocTypes

//