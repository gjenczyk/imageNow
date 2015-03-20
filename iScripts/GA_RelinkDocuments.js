/********************************************************************************
        Name:          GA_RelinkDocuments.js
        Author:        Gregg Jenczyk
        Created:        10/16/14
        Last Updated:   
        For Version:    6.7
---------------------------------------------------------------------------------
        Summary:
               This script will re-key all documents matching the index values
               of a driver document routed in to a relinking queue.
               
        Mod Summary:
               Date-Initials: Modification description.

               QUERY FOR EASY MODS:

SELECT val1.doc_id, 
val2.prop_name,
val2.string_val,
val1.string_val
FROM 
(SELECT INUSER.IN_DOC.DOC_ID,
  INUSER.IN_DRAWER.DRAWER_NAME,
  INUSER.IN_PROP.PROP_NAME,
  INUSER.IN_INSTANCE_PROP.STRING_VAL
FROM INUSER.IN_DOC
INNER JOIN INUSER.IN_DRAWER
ON INUSER.IN_DRAWER.DRAWER_ID = INUSER.IN_DOC.DRAWER_ID
INNER JOIN INUSER.IN_INSTANCE
ON INUSER.IN_DOC.INSTANCE_ID = INUSER.IN_INSTANCE.INSTANCE_ID
INNER JOIN INUSER.IN_INSTANCE_PROP
ON INUSER.IN_INSTANCE.INSTANCE_ID = INUSER.IN_INSTANCE_PROP.INSTANCE_ID
INNER JOIN INUSER.IN_PROP
ON INUSER.IN_PROP.PROP_ID   = INUSER.IN_INSTANCE_PROP.PROP_ID
WHERE INUSER.IN_DOC.DOC_ID <> '321YYC6_062WSRGVE00001P'
AND INUSER.IN_DRAWER.DRAWER_NAME IN ('UMBGA','UMBGA Lockbox')
AND INUSER.IN_DOC.FOLDER                = '01559762'
AND INUSER.IN_INSTANCE.DELETION_STATUS <> 1
AND PROP_NAME = 'SA Application Nbr') val1
full outer JOIN 
(SELECT INUSER.IN_DOC.DOC_ID,
INUSER.IN_DRAWER.DRAWER_NAME,
INUSER.IN_PROP.PROP_NAME,
INUSER.IN_INSTANCE_PROP.STRING_VAL
FROM INUSER.IN_DOC
INNER JOIN INUSER.IN_DRAWER
ON INUSER.IN_DRAWER.DRAWER_ID = INUSER.IN_DOC.DRAWER_ID
INNER JOIN INUSER.IN_INSTANCE
ON INUSER.IN_INSTANCE.INSTANCE_ID = INUSER.IN_DOC.INSTANCE_ID 
INNER JOIN INUSER.IN_INSTANCE_PROP
ON INUSER.IN_INSTANCE.INSTANCE_ID = INUSER.IN_INSTANCE_PROP.INSTANCE_ID 
INNER JOIN INUSER.IN_PROP 
ON INUSER.IN_INSTANCE_PROP.PROP_ID = INUSER.IN_PROP.PROP_ID
WHERE INUSER.IN_DRAWER.DRAWER_NAME LIKE 'UMBGA%'
AND INUSER.IN_DOC.FOLDER = '01559762'  
AND PROP_NAME = 'Shared') val2 
on val1.doc_id = val2.doc_id;
               
********************************************************************************/

// ********************* Include additional libraries *******************
//#link "inxml"    //XML parser
//#link "sedbc"    //Database object
//#link "secomobj" //COM object
#include "$IMAGENOWDIR6$\\script\\lib\\iScriptDebug.jsh"
#include "$IMAGENOWDIR6$\\script\\lib\\RouteItem.jsh"
#include "$IMAGENOWDIR6$\\script\\lib\\envVariable.jsh"
#include "$IMAGENOWDIR6$\\script\\lib\\HostDBLookupInfo.jsh"
#include "$IMAGENOWDIR6$\\script\\lib\\UnifiedPropertyManager.jsh"
#include "$IMAGENOWDIR6$\\script\\STL\\packages\\Document\\reindexDocument.js"

// *********************         Configuration        *******************

// logging
#define LOG_TO_FILE         true    // false - log to stdout if ran by intool, true - log to inserverXX/log/ directory
#define DEBUG_LEVEL         5       // 0 - 5.  0 least output, 5 most verbose
#define SPLIT_LOG_BY_THREAD false   // set to true in high volume scripts when multiple worker threads are used (workflow, external message agent, etc)
#define MAX_LOG_FILE_SIZE   100     // Maximum size of log file (in MB) before a new one will be created

// *********************       End  Configuration     *******************

// ********************* Initialize global variables ********************
var debug;
var ERROR_FLAG = false; // presumption is that everything will succeed!
var DOC_TYPE = "Relink Documents ";
var ERR_QUEUE = " Relink Error"; 
var COMPLETE_QUEUE = " Recycle Bin";
var SHARED_Y_STRING = "301YT7N_000CFJ25Y0000NX";
var UM = "UM";

var WHITE_LIST = {

    "BGA":["Document App Plan", "Signature Reason", "Source"],
    "DGA":["Document App Plan", "Signature Reason", "Source"],
    "LGA":["Document App Plan", "Signature Reason", "Source"]

};

var SHARED_UPDATE_CPS = {"DOB", "SSN#, Legacy"};
var SHARED_BLANK_CPS = {"Plan", "Plan Code", "Program", "Program Code", "SA Application Nbr", "Sub-Plan", "Sub-Plan Code", "Term Code"};


/**
* Main body of script.
* @method main
* @return {Boolean} True on success, false on error.
*/
function main ()
{
        try
        { 
            debug = new iScriptDebug("GA_RelinkDocuments", LOG_TO_FILE, DEBUG_LEVEL);
            debug.log("WARNING", "GA_RelinkDocuments.js starting.\n");

            var wfItem = INWfItem.get(currentWfItem.id);

			      if (!wfItem.getInfo())
			      {
				      debug.log("ERROR","Could not get info for wf ID: [%s]. Error: [%s]\n", currentWfItem.id, getErrMsg());
				      return false;
			      }

            // extract campus name from current workflow queue
			      var campus = wfItem.queueName.substring(2,5);
            // build trigger doc type name & determine error queue
            var relinkTrigger = DOC_TYPE + campus;
            var errorQueue = UM + campus + ERR_QUEUE;
            var triggerCreator = wfItem.creationUserName;
            var otherThing = wfItem.stateUserName;
            debug.log("DEBUG","stateUserName = [%s]\n", otherThing);
            var recycleQueue = UM + campus + COMPLETE_QUEUE;

			      var wfDoc = new INDocument(wfItem.objectId);
			      if (!wfDoc.getInfo())
            {
              debug.log("ERROR","Could not get info for doc ID: [%s]. Error: [%s]\n", wfDoc.name, getErrMsg());
              RouteItem(wfItem.id, errorQueue, "Could not get doc Info!");
              return false;
            }

            //check to make sure document is the right driver type to kick off relinking //verify DOC_TYPE + drawer = current wfItem doctype
            if (wfDoc.docTypeName != relinkTrigger)
            {
              debug.log("WARNING","Document [%s][%s] is not type: [%s]. Routing to [%s]\n", wfDoc.id, wfDoc.docTypeName, relinkTrigger, errorQueue);
              RouteItem(wfItem.id, errorQueue, "Incorrect doc type for relinking");
              return false;
            }

            //get all relevant info from trigger doc
            var driverDrawer = wfDoc.drawer;
            var driverField1 = wfDoc.field1;

            var driverCPs = new UnifiedPropertyManager();
            var custProps = driverCPs.GetAllCustomProps(wfDoc);
            var driverPropArrray = [];
            var driverAppNumber; 

            for (cp = 0; cp<custProps.length; cp++)
            {
              debug.log("DEBUG","Trigger CP Name/Value [%s] [%s]\n", custProps[cp].name, custProps[cp].getValue());
              driverPropArrray[cp] = new driverCustProp(custProps[cp].name, custProps[cp].getValue());
              if(custProps[cp].name == "SA Application Nbr")
              {
                driverAppNumber = custProps[cp].getValue();
              }
            } // end for (cp = 0; cp<custProps.length; cp++)

            // get the whitelist for the current campus
            var currentWhiteList;
            for (var campusName in WHITE_LIST)
            {
              if (campusName == campus)
              {
                currentWhiteList = eval("WHITE_LIST." + campusName);
                debug.log("DEBUG","currentWhiteList = %s %s\n", currentWhiteList[0], currentWhiteList[1]);
                break;
              }
            }

            var driverDrawer = wfDoc.drawer;
            var driverField1 = wfDoc.field1;
            var driverField2 = wfDoc.field2;
            var driverField3 = wfDoc.field3;
            var driverField4 = wfDoc.field4;
            var driverId = wfDoc.id;

            if (!driverField1 || driverField1 == null)
            {
              debug.log("ERROR","Driver document has missing value in Field1.\n");
              RouteItem(wfItem.id, errorQueue, "Please update student ID");
              return false;
            }

            //printf("%s %s %s %s\n", driverField1, driverField2, driverField3, driverField4);

            //search for all documents matching the driver
            //sql
            var sql = "SELECT val0.doc_id, " +
                      "val2.prop_name, " +
                      "val2.string_val, " +
                      "val1.string_val " +
                      "FROM " +
                      "(SELECT INUSER.IN_DOC.DOC_ID " + 
                      " FROM INUSER.IN_DOC " + 
                      " INNER JOIN INUSER.IN_DRAWER " + 
                      " ON INUSER.IN_DRAWER.DRAWER_ID = INUSER.IN_DOC.DRAWER_ID " + 
                      " INNER JOIN INUSER.IN_INSTANCE " + 
                      " ON INUSER.IN_DOC.INSTANCE_ID = INUSER.IN_INSTANCE.INSTANCE_ID " +
                      " WHERE INUSER.IN_DRAWER.DRAWER_NAME IN ('" + driverDrawer + "','" + driverDrawer + " Lockbox') " + 
                      " AND INUSER.IN_INSTANCE.DELETION_STATUS <> 1 " +
                      " AND INUSER.IN_DOC.DOC_ID<> '"+ driverId +"' " +           
                      " AND INUSER.IN_DOC.FOLDER = '" + driverField1 + "') val0 " + 
                      "LEFT JOIN " + 
                      "(SELECT INUSER.IN_DOC.DOC_ID, " +
                      " INUSER.IN_DRAWER.DRAWER_NAME, " +
                      " INUSER.IN_PROP.PROP_NAME, " +
                      " INUSER.IN_INSTANCE_PROP.STRING_VAL " +
                      "FROM INUSER.IN_DOC " +
                      "INNER JOIN INUSER.IN_DRAWER " +
                      "ON INUSER.IN_DRAWER.DRAWER_ID = INUSER.IN_DOC.DRAWER_ID " +
                      "INNER JOIN INUSER.IN_INSTANCE " + 
                      "ON INUSER.IN_DOC.INSTANCE_ID = INUSER.IN_INSTANCE.INSTANCE_ID " +
                      "INNER JOIN INUSER.IN_INSTANCE_PROP " +
                      "ON INUSER.IN_INSTANCE.INSTANCE_ID = INUSER.IN_INSTANCE_PROP.INSTANCE_ID " +
                      "INNER JOIN INUSER.IN_PROP " +
                      "ON INUSER.IN_PROP.PROP_ID   = INUSER.IN_INSTANCE_PROP.PROP_ID " +
                      "WHERE INUSER.IN_DOC.FOLDER = '" + driverField1 + "'" +
                      "AND PROP_NAME = 'SA Application Nbr') val1 ON val0.doc_id = val1.doc_id " +
                      "LEFT JOIN " +
                      "(SELECT INUSER.IN_DOC.DOC_ID, " +
                      "INUSER.IN_DRAWER.DRAWER_NAME, " +
                      "INUSER.IN_PROP.PROP_NAME, " +
                      "INUSER.IN_INSTANCE_PROP.STRING_VAL " +
                      "FROM INUSER.IN_DOC " +
                      "INNER JOIN INUSER.IN_DRAWER " +
                      "ON INUSER.IN_DRAWER.DRAWER_ID = INUSER.IN_DOC.DRAWER_ID " +
                      "INNER JOIN INUSER.IN_INSTANCE " +
                      "ON INUSER.IN_INSTANCE.INSTANCE_ID = INUSER.IN_DOC.INSTANCE_ID " +
                      "INNER JOIN INUSER.IN_INSTANCE_PROP " +
                      "ON INUSER.IN_INSTANCE.INSTANCE_ID = INUSER.IN_INSTANCE_PROP.INSTANCE_ID " +
                      "INNER JOIN INUSER.IN_PROP " +
                      "ON INUSER.IN_INSTANCE_PROP.PROP_ID = INUSER.IN_PROP.PROP_ID " +
                      "WHERE INUSER.IN_DRAWER.DRAWER_NAME LIKE '" + driverDrawer + "%' " +
                      "AND INUSER.IN_DOC.FOLDER = '" + driverField1 + "' " +
                      "AND PROP_NAME = 'Shared') val2 " +
                      "on val0.doc_id = val2.doc_id;"; 

            var returnVal;
            var cur = getHostDBLookupInfo_cur(sql,returnVal);
            
            if(!cur || cur == null)
            {
              debug.log("WARNING","no results returned for query.\n");
              RouteItem(wfItem.id, errorQueue, "No matching docs found!");
              return false;
            } 

            var docCount = 0;
            var sharedCount = 0;
            var updateCount = 0;

            while(cur.next())
            {     
              docCount++;
              var relinkID = cur[0];
              var sharedProp = cur[1];
              var sharedVal = cur[2];
              var targetAppNumber = cur[3];

              var docToUpdate = new INDocument(relinkID);
              if (!docToUpdate.getInfo())
              {
                debug.log("ERROR","Could not get info for doc ID: [%s]. Error: [%s]\n", docToUpdate.name, getErrMsg());
                ERROR_FLAG = true;
                continue;
              }
              //printf("[%s] [%s] [%s]\n", relinkID, sharedProp, sharedVal);

              // account for items that are currently ope in workflow
              var wfStatus = docToUpdate.getWfInfo();
              if (!wfStatus || wfStatus == null)
              {
                debug.log("ERROR","Couldn't get WFInfo for %s.  Error: %s\n", docToUpdate.id, getErrMsg());
                return false;
              } 

              if(wfStatus.length > 0)
              {
                var wfToCheck = new INWfItem(wfStatus[0].id);
                if (wfToCheck.state == 2)  
                {
                  debug.log("ERROR","Item [%s]is currently being processed in workflow.  Cannot relink. [%s]\n", docToUpdate.id, wfToCheck.state);
                  ERROR_FLAG = true;
                  continue;
                }
              }

              if ((!sharedProp || !sharedVal) && docToUpdate.field4.toUpperCase() != "SHARED")
              {
                
                //don't process if the app #'s don't match
                if (driverAppNumber != targetAppNumber)
                {
                  debug.log("INFO","Not processing [%s] because SA Application Nbrs don't match: D:[%s]T:[%s]\n", relinkID, driverAppNumber, targetAppNumber);
                  continue;
                }

                var targetDrawer = docToUpdate.drawer;
                var targetField1 = docToUpdate.field1;
                var targetField2 = docToUpdate.field2;
                var targetField3 = docToUpdate.field3;
                var targetField4 = docToUpdate.field4;
                var targetField5 = docToUpdate.field5;
                var targetType = docToUpdate.docTypeName;

                var keys = new INKeys(targetDrawer, driverField1, driverField2, driverField3, driverField4, targetField5, targetType);
                //printf("%s\n", keys.toString());

                var targetCPName;
                var targetCPVal;

                debug.log("DEBUG","Target info: %s %s %s %s %s %s\n", targetDrawer, targetField1, targetField2, targetField3, targetField4, targetType);
                //printf("Target info: %s %s %s %s %s %s\n", targetDrawer, targetField1, targetField2, targetField3, targetField4, targetType);

                var targetCPs = new UnifiedPropertyManager();
                var targetCustProps = targetCPs.GetAllCustomProps(docToUpdate); 
                var setProp;
                var propsToUpdate = [];

                for (tp = 0; tp<targetCustProps.length; tp++)
                {
                  targetCPName = targetCustProps[tp].name;
                  targetCPVal = targetCustProps[tp].getValue();

                  //debug.log("INFO","Target CP Name & Value [%s] [%s]\n", targetCustProps[tp].name, targetCustProps[tp].getValue());
                  //printf(""Target CP Name & Value [%s] [%s]\n", targetCustProps[tp].name, targetCustProps[tp].getValue());
                  

                  for (var a in driverPropArrray)
                  {
                    if (driverPropArrray[a].name == targetCPName && !include(currentWhiteList, driverPropArrray[a].name))
                    {
                      //printf("#:%s dpan:%s tcpn:%s dpav:%s tcpv:%s\n", a, driverPropArrray[a].name, targetCPName, driverPropArrray[a].value, targetCPVal);

                      setProp = new INInstanceProp();
                      setProp.name = targetCPName;
                      setProp.setValue(driverPropArrray[a].value);

                      propsToUpdate.push(setProp);
                      break;
                    }  // end if (driverPropArrray[a].name == targetCPName)
                  } // end for (var a in driverPropArrray)
                } // end for (tp = 0; tp<targetCustProps.length; tp++)

              } // end if (!sharedProp || !sharedVal)

              //debug.log("INFO","\n\n\n\n%s\n\n\n\n\n\n", docToUpdate.field4.toUpperCase());
              else if ((sharedProp && sharedVal) ||  (docToUpdate.field4.toUpperCase() == "SHARED"))
              {

                sharedCount++;
                var targetDrawer = docToUpdate.drawer;
                var targetField1 = docToUpdate.field1;
                var targetField2 = docToUpdate.field2;
                var targetField3 = "";
                var targetField4 = "Shared";
                var targetField5 = docToUpdate.field5;
                var targetType = docToUpdate.docTypeName;

                var keys = new INKeys(targetDrawer, driverField1, driverField2, targetField3, targetField4, targetField5, targetType);
                //printf("%s\n", keys.toString());

                var targetCPName;
                var targetCPVal;

                debug.log("DEBUG","Target info: %s %s %s %s %s %s\n", targetDrawer, targetField1, targetField2, targetField3, targetField4, targetType);
                //printf("Target info: %s %s %s %s %s %s\n", targetDrawer, targetField1, targetField2, targetField3, targetField4, targetType);

                var targetCPs = new UnifiedPropertyManager();
                var targetCustProps = targetCPs.GetAllCustomProps(docToUpdate); 
                var setProp;
                var propsToUpdate = [];

                for (tp = 0; tp<targetCustProps.length; tp++)
                {
                  targetCPName = targetCustProps[tp].name;
                  targetCPVal = targetCustProps[tp].getValue();

                  //debug.log("INFO","Target CP Name & Value [%s] [%s]\n", targetCustProps[tp].name, targetCustProps[tp].getValue());

                  for (var a in driverPropArrray)
                  {
                    if ((driverPropArrray[a].name == targetCPName) && include(SHARED_UPDATE_CPS, targetCPName) && !include(currentWhiteList, driverPropArrray[a].name))
                    {
                      //printf("#:%s dpan:%s tcpn:%s \n", a, driverPropArrray[a].name, targetCPName);
                      //printf("#:%s dpav:%s tcpv:%s\n", a, driverPropArrray[a].value, targetCPVal);
                      setProp = new INInstanceProp();
                      setProp.name = targetCPName;
                      setProp.setValue(driverPropArrray[a].value);
                      propsToUpdate.push(setProp);
                      break;
                    }  // end if (driverPropArrray[a].name == targetCPName)
                    if ((driverPropArrray[a].name == targetCPName) && include(SHARED_BLANK_CPS, targetCPName) && !include(currentWhiteList, driverPropArrray[a].name))
                    {
                      //printf("#:%s dpan:%s tcpn:%s \n", a, driverPropArrray[a].name, targetCPName);
                      //printf("#:%s dpav:%s tcpv:%s\n", a, driverPropArrray[a].value, targetCPVal);
                      setProp = new INInstanceProp();
                      setProp.name = targetCPName;
                      setProp.setValue("");
                      propsToUpdate.push(setProp);
                      break;
                    }  // end if ((driverPropArrray[a].name == targetCPName) && include(SHARED_BLANK_CPS, targetCPName))
                  } // end for (var a in driverPropArrray)
                } // end for (tp = 0; tp<targetCustProps.length; tp++)

              }// end if (sharedProp && sharedVal)
              // if we can't determine what type of document it is
              else
              {
                debug.log("ERROR","Unable to determine if [%s] is shared (VALS:[%s][%s]).\n", relinkID, sharedProp, sharedVal);
                ERROR_FLAG = true;
                continue;
              }

              // set the index values / custom properties on the documents
              if(!docToUpdate.setCustomProperties(propsToUpdate))
                {
                  debug.log("ERROR","Couldn't update CPs for [%s]. Error: %s\n", docToUpdate.id, getErrMsg());
                  ERROR_FLAG = true;
                }
                else
                {
                  debug.log("INFO","Custom Property values been updated for [%s]\n", relinkID);
                }
                if (!reindexDocument(relinkID, keys, "APPEND"))
                {
                  debug.log("ERROR","Couldn't update keys for [%s]. Error: %s\n", docToUpdate.id, getErrMsg());
                  ERROR_FLAG = true;
                }
                else
                {
                  debug.log("INFO","Index values been updated for [%s] [%s]\n", relinkID, keys);
                }

                if (!writeToWFHistory(docToUpdate, triggerCreator))
                {
                  ERROR_FLAG = true;
                }

                if(!ERROR_FLAG)
                {
                  updateCount++;
                }

            } // end while(Cur.next())
            debug.log("DEBUG","Evaluated [%s] documents ([%s] shared) and updated [%s] belonging to [%s]\n",docCount,sharedCount,updateCount,driverField1);
            //clean up the driver now that we're done relinking (maybe?)
            if (ERROR_FLAG) // if relinking fails, add note to driver and route to error queue
            {
              debug.log("ERROR","The script encountered errors in relinking.  Routing [%s] to [%s]\n", wfItem.id, errorQueue);
              RouteItem(wfItem.id, errorQueue, "Unable to relink document(s).  Please check GA_RelinkDocuments log.");
            }
            else
            {
              debug.log("INFO","Relinking complete - routing [%s] to [%s]\n", wfDoc.id, recycleQueue);
              RouteItem(wfItem.id, recycleQueue, "Relinking successful");
            }

        } // end of try
        
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

function driverCustProp (name, value)
{
   this.name = name;
   this.value = value;
} //end driverCustProp

function writeToWFHistory (targetDoc, creator)
{
  var itemWFInfo = targetDoc.getWfInfo();

  if (!itemWFInfo || itemWFInfo == null)
  {
    debug.log("ERROR","Couldn't get WFInfo for %s.  Error: %s\n", targetDoc.id, getErrMsg());
    return false;
  } 

  if (itemWFInfo.length < 1) // add relinking info to notes field for docs not in WF
  {
    debug.log("INFO","No workflow history exists to write to for [%s].  Writing to notes.\n", targetDoc.id);
    var notes = targetDoc.getNotes();
    var curTime = new Date();
    notes += "\n"+ curTime +" This item was relinked via iScript by " + creator;

    targetDoc.setNotes(notes);

    return true;
  }
  else
  {
    var itemToNote = new INWfItem(itemWFInfo[0].id);

    if (!itemToNote.getInfo())
    {
      debug.log("ERROR","Failed to get item. Error: %s.\n", getErrMsg());
      return false;
    }

    if (itemToNote.state == 2)  //GJ - is this step necessary?
    {
      debug.log("ERROR","Item [%s]is currently being processed in workflow.  Cannot relink. [%s]\n", targetDoc.id, itemToNote.state);
      return false;
    }

    if (!itemToNote.setState(WfItemState.Idle, "This item was relinked via iScript by " + creator))
    {
      debug.log("ERROR","Could not set state: %s.\n", getErrMsg());
      return false;
    }

  return true;
  }
} // end writeToWFHistory

function include(arr, obj) {
    for(var i=0; i<arr.length; i++) {
        if (arr[i] == obj) return true;
    }
} // end include

//