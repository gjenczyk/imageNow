/********************************************************************************
        Name:          updateDoctypeProps
        Author:        Gregg Jenczyk
        Created:        03/06/2015
        Last Updated:   
        For Version:    
---------------------------------------------------------------------------------
        Summary:
               This script will change the custom properties on a doctype.  It can
               be used to add or remove all CPs on a given doctype.
               
        Mod Summary:
               Date-Initials: Modification description.
               
EXAMPLE YAML CONFIG:
--------------------
DOCTYPE_CONFIG:
 - DOCTYPE_TO_UPDATE:
   DOC_INFO:
    - name: Doctype1
      create: false
  CPD_TO_ADD:
   - name: CustProp1
     position: 2
   - name: CustProp2
     position:
   - name: CustProp3
  CPS_TO_REMOVE:
   - name: CustProp4
 - DOCTYPE_TO_UPDATE:
   DOC_INFO:
    - name: Doctype2
      create: true
      list: DocTypeList1
  CPS_TO_ADD:
   - name: CustProp1
   - name: CustProp4
 - DOCTYPE_TO_UPDATE:
   DOC_INFO:
    - name: Doctype3
   CPS_TO_REMOVE:
    - name: CustProp9

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

DOCTYPE_OK = {name:"DOCTYPE_OK",value:false,create:false};
CP_ADD_OK = {name:"CP_ADD_OK",value:false};
CP_REMOVE_OK = {name:"CP_REMOVE_OK",value:false};
FLAGS = [DOCTYPE_OK,CP_ADD_OK,CP_REMOVE_OK];

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

      //load the yaml with doc config information
      debug.log("INFO","Attempting to load YAML\n");
      loadYAMLConfig(imagenowDir6+"\\script\\config_scripts\\updateDoctypeProps\\");

      for (var docTypeConfig in CFG.updateDoctypeProps)
      { 
        debug.log("DEBUG","Begin processing UDTP configuration\n");
        var UDTP_CONFIG = CFG.updateDoctypeProps[docTypeConfig].DOCTYPE_CONFIG;

        debug.log("DEBUG","Doctypes in YAML File [%s]\n",UDTP_CONFIG.length);
        //update each doctype passed
        for(var i=0; i<UDTP_CONFIG.length; i++)
        {
          //reset flags to false
          resetFlags();

          //values from the yaml
          var workingDocType = UDTP_CONFIG[i].DOC_INFO;
          workingDocType = workingDocType[0];
          var cpsToAdd = UDTP_CONFIG[i].CPS_TO_ADD;
          var cpsToRemove = UDTP_CONFIG[i].CPS_TO_REMOVE;

          var yamlArr = [workingDocType, cpsToAdd, cpsToRemove];

          for (var d = 0; d < yamlArr.length; d++)
          {
            if(!yamlArr[d])
            {
              yamlArr[d] == null;
            }
          }

          debug.log("DEBUG","UDTP_CONFIG Config: DOCTYPE_TO_UPDATE[%s], DOCTYPE_CREATE[%s], DOCTYPE_LIST[%s], CPS_TO_ADD[%s], CPS_TO_REMOVE[%s]\n",workingDocType.name, workingDocType.create, workingDocType.list, cpsToAdd, cpsToRemove);

          //check the configuration for validity.  Don't do anything if the configured information is invalid
          if (!checkConfig(workingDocType, cpsToAdd, cpsToRemove))
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
            debug.log("INFO","*--------------------Working the [%s] doctype--------------------*\n",workingDocType.name);
            //get info aobut the doctype
            var docType = new INDocType();
            var propArr = [];
            docType.name = workingDocType.name;
            if (docType.getInfo())
            {
              debug.log("DEBUG","Doctype id:%s, name:%s, desc:%s, isActive:%s\n", docType.id, docType.name, docType.desc, docType.isActive);
              var props = docType.props;
              debug.log("DEBUG","Number of custom properties: %d\n",props.length);
              for (b=0; b<props.length; b++)
              {
                debug.log("INFO","EXISTING DOCTYPE PROPS: id:%s, name:%s, isRequired:%s \n",props[b].id, props[b].name, props[b].isRequired);
                propArr.push(props[b]);
              }//end for (b=0; b<props.length; b++)
            }
            else
            {
              //create it here if we want to do that?
              debug.log("ERROR","Failed to retrieve info for document type - Error: [%s]\n", getErrMsg());
              return false;
            }//end if (docType.getInfo()) else

              //add document type to list if the list is set
              if(workingDocType.list)
              {
                debug.log("DEBUG","Attempting to add [%s] to [%s]\n",workingDocType.name, workingDocType.list);
                if(!addToList(workingDocType))
                {
                  debug.log("ERROR","Could not add [%s] to [%s]. Check config and rerun!\n",workingDocType.name, workingDocType.list);
                  return false;
                }
                else
                {
                  debug.log("INFO","Added [%s] to [%s]\n",workingDocType.name, workingDocType.list);
                }
              }//end adding to list

            //if adding values
            if(CP_ADD_OK.value)
            {
              debug.log("INFO","Adding CPs to [%s]\n",docType.name);

              //get the number of props to be inserted, sort them, and increment the postion appropriately.  
              //verify that we have a valid positional parameter.  If not, insert it at the end of the list.
              var propLen = propArr.length;
              var insLen = cpsToAdd.length;
              var posMod = 0;

              for (var f = 0; f < insLen; f++)
              {
                //set position if we have a bad value.
                if(!parseInt(cpsToAdd[f].position) || !cpsToAdd[f].position || cpsToAdd[f].position == null)
                {
                  debug.log("WARNING","Passed bad value for position: [%s] position: [%s]\n",cpsToAdd[f].name, cpsToAdd[f].position);
                  cpsToAdd[f].position = propLen + posMod;
                  debug.log("INFO","New position for [%s] is [%s]\n",cpsToAdd[f].name, cpsToAdd[f].position);
                  posMod++;
                }
                //catch it if people enter int as a string
                if (typeof cpsToAdd[f].position == "string")
                {
                 cpsToAdd[f].position = Number(cpsToAdd[f].position);
                }
              }//end for (var f = 0; f < insLen; f++)

              //now that that's done, sort cpsToAdd by position and increment values so they go where you want them to!
              cpsToAdd.sort(function(a, b){return a.position-b.position});

              //now, adjust postion so props go where you want them to
              for (var e = 0; e < insLen; e++)
              {
                if (cpsToAdd[e].position <= propLen)
                {
                  cpsToAdd[e].position +=e;
                }
                else
                {
                  cpsToAdd[e].position = propLen + e;
                }
              }//end for (var e = 0; e < insLen; e++)

              //insert the values in the array
              for (var g = 0; g < insLen; g++)
              {
                debug.log("DEBUG","Adding: [%s] [%s]\n",cpsToAdd[g].position,cpsToAdd[g].name)
                propb = new INProperty();
                propb.name = cpsToAdd[g].name;
                propb.getInfo();
                var propa = [];
                propa[0] = new INClassProp();
                propa[0].id = propb.id;
                propa[0].name = cpsToAdd[g].name;
                propa[0].isRequired  = false;
                propArr.splice(cpsToAdd[g].position,0,propa[0]);
              }//end for (var g = 0; g < insLen; g++)

              //lastly, make sure we're not inserting a value that is already there
        
              var arrToCheck = propArr.sort(function compare(a,b) {if (a.id < b.id)return -1;if (a.id > b.id)return 1;return 0;});
              var dupCheck = arrToCheck.length;

              var dupErr = false;
              while(dupCheck > 1)
              {
                dupCheck--;
                if(arrToCheck[dupCheck].name == arrToCheck[dupCheck-1].name)
                {
                  debug.log("ERROR","Attempted to insert same CP twice! CP: [%s].\n",arrToCheck[dupCheck].name);
                  dupErr = true;
                }
              } //end while(dupCheck > 1)

              if(dupErr)
              {
                debug.log("ERROR","Duplicate CPs were detected.  Please fix configuration and re-run. Be sure to remove configuration for documents that have already been processed.\n")
                retrun false;
              }

              if (!updateDoc(docType,propArr))
              {
                debug.log("ERROR","Could not update doctype [%s]\n",docType.name);
                return false;
              }

            }//end if adding values

            //if removing values
            if(CP_REMOVE_OK.value)
            {
              debug.log("INFO","Removing CPs from [%s]\n",docType.name);
              
              // go back out and re-run with new array
              for (var m = 0; m < cpsToRemove.length; m++)
              {
                propArr = removeFromArray(propArr, cpsToRemove);
              } // proparr will now have all configured CPs removed

              if (!updateDoc(docType,propArr))
              {
                debug.log("ERROR","Could not update doctype [%s]\n",docType.name);
                return false;
              }
            }//end if removing values
          }//if (CP_ADD_OK || CP_REMOVE_OK) 
        }//update each doctype passed
      }//end for (var docTypeConfig in CFG.updateDoctypeProps)
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
function checkConfig(docType,addCusProps,remCusProps)
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
  var cpAdd = checkCPs(allProps,addCusProps,CP_ADD_OK);
  //end check the cps to add
  
  //check the cps to remove
  var cpRem = checkCPs(allProps,remCusProps,CP_REMOVE_OK);
  //end check the cps to remove

  //check the document types
  var dtToWork = checkDocTypes(docType, DOCTYPE_OK);
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
  for (var c = 0; c < arr.length; c++)
  {
    if (item == arr[c])
    {
      success = true;
    }
  }
  return success;
}//end arrCheck

//function to check if CPs exist
function checkCPs(allCpArr, workCpArr, cpFlag)
{
  if(workCpArr == null || !workCpArr)
  {
    debug.log("WARNING","No CP config was available to check. [%s] is false, not changing CPs.\n",cpFlag.name);
    cpFlag.value = false;
    return true;
  }

  for(var j = 0; j < workCpArr.length; j++)
  {
    if (workCpArr[j].name == "" || workCpArr[j].name == null)
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
  var curDT = INDocType.get(dtArr.name);
  if(!curDT || curDT == null)
  {   
    if (dtArr.create)
    {
      if(!createDocType(dtArr))
      {
        debug.log("ERROR","Could not create docType [%s]\n",dtArr.name);
        return false;    
      }
      else
      {
        dtFlag.create = true;
      }
    }
    else
    {
      debug.log("ERROR","Couldn't get info for [%s] - [%s]\n", dtArr.name, getErrMsg());
      return false;
    }
  }

  dtFlag.value = true;
  debug.log("INFO","Setting [%s] to [%s].\n",dtFlag.name, dtFlag.value);
  return true;
}//end checkDocTypes

//function to update doctype
function updateDoc(docType1, propArr1)
{
  for (m=0; m<propArr1.length; m++)
  {
    debug.log("DEBUG","NEW DOCTYPE PROPS: id:%s, name:%s, isRequired:%s \n",propArr1[m].id, propArr1[m].name, propArr1[m].isRequired);
  }

  /*debug.log("DEBUG","docType1.name = [%s]\n", docType1.name);
  debug.log("DEBUG","docType1.desc = [%s]\n", docType1.desc);
  debug.log("DEBUG","docType1.isActive = [%s]\n", docType1.isActive);
  debug.log("DEBUG","propArr1 = [%s]\n", propArr1);
  for(var p = 0; var p < propArr1.length; p++)
  {
    debug.log("DEBUG","ID:             %s\n", propArr1[p].id);
    debug.log("DEBUG","name:           %s\n", propArr1[p].name);
    debug.log("DEBUG","type:           %s\n", propArr1[p].type);
    debug.log("DEBUG","defaultValue:   %s\n", propArr1[p].defaultValue);
    debug.log("DEBUG","displayFormat   %s\n", propArr1[p].displayFormat);
    debug.log("DEBUG","isActive:       %s\n", propArr1[p].isActive);
  }*/
  debug.log("DEBUG","Attempting to update [%s]. This may take some time....\n", docType1.name);
  if(docType1.update(docType1.name, docType1.desc, docType1.isActive, 0, propArr1))
  {
    debug.log("INFO","Successfully updated document type [%s].\n", docType1.name);
  } 
  else
  {
    debug.log("ERROR","Aborting - Failed to update document type - %s.\n", getErrMsg());
    return false;
  }
  return true;
} //end function to update doctype

//function to make sure flags are false for a new document type
function resetFlags () {
// is this crazy?
  for (var flag in FLAGS)
  {
    for (var subFlag in FLAGS[flag])
    {
      if(FLAGS[flag][subFlag] == true)
      {
        FLAGS[flag][subFlag] = false;
      }
    }
  }
}//end function to make sure flags are false for a new document type

//function to create a doctype
function createDocType(makeDoc)
{
  var newDocType = INDocType.add(makeDoc.name);
  if(!newDocType || newDocType == null)
  {
    debug.log("ERROR","Error creating newDocType [%s] [%s]\n",makeDoc.name, getErrMsg());
    return false;
  }
  else
  {
    return true;
  }
}//end function to create a doctype

//function to add a doctype to a list
function addToList(dtInfo)
{
  var addList = new INDocTypeList("",dtInfo.list)
  var targetList = addList.getInfo();
  if(!targetList || targetList == null)
  {
    debug.log("ERROR","DocTypeList [%s] doesn't exist!!\n",dtInfo.list);
    return false;
  }
  else
  {
    debug.log("INFO","DocTypeList [%s] exists! Attempting to add [%s]\n",dtInfo.list, dtInfo.name);
    var typeToAdd = new INDocType();
    typeToAdd.name = dtInfo.name;
    var addedType = typeToAdd.getInfo();

    if (!addedType || addedType == null)
    {
      debug.log("ERROR","Couldn't find [%s]\n",dtInfo.name);
      return false;
    }
    
    if(!addList.updateMembers(typeToAdd.id))
    {
      debug.log("ERROR","Couldn't update doctype list [%s] - [%s]\n", addList.name, getErrMsg());
      return false;
    }
  }
  return true;
}//end function to add a doctype to a list

function removeFromArray(docCPArray, removeCPArray)
{
  debug.log("DEBUG","Inside removeFromArray.\n");
  for (var l = 0; l < docCPArray.length; l++)
  {
    for (var m = 0; m < removeCPArray.length; m++)
    {
      //debug.log("DEBUG","propArr[%s] is [%s and cpsToRemove[%s] is [%s]\n", l, docCPArray[l].name, m, removeCPArray[m].name);
      if(docCPArray[l].name == removeCPArray[m].name)
      {
        docCPArray.splice(l,1);
        return docCPArray;
      }
    }
  }
  return docCPArray;
}//end removeFromArray

//