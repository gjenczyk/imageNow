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
val2.string_val 
FROM 
(SELECT INUSER.IN_DOC.DOC_ID, 
INUSER.IN_DRAWER.DRAWER_NAME
FROM INUSER.IN_DOC
INNER JOIN INUSER.IN_DRAWER
ON INUSER.IN_DRAWER.DRAWER_ID = INUSER.IN_DOC.DRAWER_ID
WHERE INUSER.IN_DRAWER.DRAWER_NAME LIKE 'UMBGA%' 
AND INUSER.IN_DOC.FOLDER = '01559762'
AND INUSER.IN_DOC.DOC_ID <> '321YYC6_062WSRGVE00001P') val1
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
var DOC_TYPE = "Relink Documents ";
var ERR_QUEUE = " Relink Error"; 
var SHARED_Y_STRING = "301YT7N_000CFJ25Y0000NX";

var WHITE_LIST = {

    "BGA":["Document App Plan", "Source"],
    "DGA":["Document App Plan", "Source"],
    "LGA":["Document App Plan", "Source"]

};

var SHARED_UPDATE_CPS = {"DOB", "SSN#, Legacy"};
var SHARED_BLANK_CPS = {"Plan", "Plan Code"};

//probable global variables
//
//var QUEUE_COMPLETE = Recycle Bin;
//var QUEUE_ERROR = Relinking Error;

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

            var wfItem = INWfItem.get("321YYC6_062WSHGVE00002E");
            //"321YYC6_062WSHGVE00002N");//currentWfItem.id);
			      if (!wfItem.getInfo())
			      {
				      debug.log("ERROR","Could not get info for wf ID: [%s]. Error: [%s]\n", currentWfItem.id, gettErrMsg());
				      return false;
			      }

            // extract campus name from current workflow queue
			      var campus = wfItem.queueName.substring(0,3);
            // build trigger doc type name & determine error queue
            var relinkTrigger = DOC_TYPE + campus;
            var errorQueue = campus + ERR_QUEUE;
            var triggerCreator = wfItem.creationUserName;

			      var wfDoc = new INDocument(wfItem.objectId);
			      if (!wfDoc.getInfo())
            {
              debug.log("ERROR","Could not get info for doc ID: [%s]. Error: [%s]\n", wfDoc.name, gettErrMsg());
              return false;
            }

            //check to make sure document is the right driver type to kick off relinking //verify DOC_TYPE + drawer = current wfItem doctype
            if (wfDoc.docTypeName != relinkTrigger)
            {
              debug.log("WARNING","Document [%s][%s] is not type: [%s]. Routing to [%s]\n", wfDoc.id, wfDoc.docTypeName, relinkTrigger, errorQueue);
              //RouteItem(wfItem.id, errorQueue, "Incorrect doc type for relinking");
              //return false;
            }

            //get all relevant info from trigger doc
            var driverDrawer = wfDoc.drawer;
            var driverField1 = wfDoc.field1;

            var driverCPs = new UnifiedPropertyManager();
            var custProps = driverCPs.GetAllCustomProps(wfDoc);
            var driverPropArrray = []; 

            for (cp = 0; cp<custProps.length; cp++)
            {
              debug.log("INFO","Trigger CP Name/Value [%s] [%s]\n", custProps[cp].name, custProps[cp].getValue());
              driverPropArrray[cp] = new driverCustProp(custProps[cp].name, custProps[cp].getValue());
            } // end for (cp = 0; cp<custProps.length; cp++)

            // get the whitelist for the current campus
            var currentWhiteList;
            for (var campusName in WHITE_LIST)
            {
              if (campusName == campus)
              {
                currentWhiteList = eval("WHITE_LIST." + campusName);
                printf("currentWhiteList = %s %s\n", currentWhiteList[0], currentWhiteList[1]);
                break;
              }
            }

            var driverDrawer = wfDoc.drawer;
            var driverField1 = wfDoc.field1;
            var driverField2 = wfDoc.field2;
            var driverField3 = wfDoc.field3;
            var driverField4 = wfDoc.field4;
            var driverId = wfDoc.id;

            //printf("%s %s %s %s\n", driverField1, driverField2, driverField3, driverField4);

            //search for all documents matching the driver
            //sql
            var sql = "SELECT val1.doc_id, " +
                      "val2.prop_name, " +
                      "val2.string_val " +
                      "  FROM " +
                      "(SELECT INUSER.IN_DOC.DOC_ID, " +
                      " INUSER.IN_DRAWER.DRAWER_NAME " +
                      "FROM INUSER.IN_DOC " +
                      "INNER JOIN INUSER.IN_DRAWER " +
                      "ON INUSER.IN_DRAWER.DRAWER_ID = INUSER.IN_DOC.DRAWER_ID " +
                      "WHERE INUSER.IN_DRAWER.DRAWER_NAME LIKE '" + driverDrawer + "%' " +
                      "AND INUSER.IN_DOC.DOC_ID <> '"+ driverId +"'" + 
                      "AND INUSER.IN_DOC.FOLDER = '" + driverField1 + "') val1 " +
                      "full outer JOIN " +
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
                      "on val1.doc_id = val2.doc_id;" 

            var returnVal;
            var cur = getHostDBLookupInfo_cur(sql,returnVal);
            
            if(!cur || cur == null)
            {
              debug.log("WARNING","no results returned for query.\n");
              return false;
            } 


            while(cur.next())
            {     
              var relinkID = cur[0];
              var sharedProp = cur[1];
              var sharedVal = cur[2];

              var docToUpdate = new INDocument(relinkID);
              if (!docToUpdate.getInfo())
              {
                debug.log("ERROR","Could not get info for doc ID: [%s]. Error: [%s]\n", docToUpdate.name, gettErrMsg());
                return false;
              }
              //printf("[%s] [%s] [%s]\n", relinkID, sharedProp, sharedVal);
              if (!sharedProp || !sharedVal)
              {
                
                var targetDrawer = docToUpdate.drawer;
                var targetField1 = docToUpdate.field1;
                var targetField2 = docToUpdate.field2;
                var targetField3 = docToUpdate.field3;
                var targetField4 = docToUpdate.field4;
                var targetField5 = docToUpdate.field5;
                var targetType = docToUpdate.docTypeName;

                var keys = new INKeys(targetDrawer, driverField1, driverField2, driverField3, driverField4, targetField5, targetType, "WithCustProps");
                printf("%s\n", keys.toString());

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
                    if (driverPropArrray[a].name == targetCPName)
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
                
                if(!docToUpdate.setCustomProperties(propsToUpdate))
                {
                  printf("%s\n", getErrMsg());
                }
                else
                {
                  printf("Custom Property values been updated\n");
                }
                if (!reindexDocument(relinkID, keys, "APPEND"))
                {
                  printf("%s\n", getErrMsg());
                }
                else
                {
                  printf("Index values been updated\n");
                }

                //writeToWFHistory(docToUpdate, triggerCreator);

              } // end if (!sharedProp || !sharedVal)
              else if (sharedProp && sharedVal)
              {
                debug.log("INFO","Processing \n");

                var targetDrawer = docToUpdate.drawer;
                var targetField1 = docToUpdate.field1;
                var targetField2 = docToUpdate.field2;
                var targetField3 = docToUpdate.field3;
                var targetField4 = "SHARED";
                var targetField5 = docToUpdate.field5;
                var targetType = docToUpdate.docTypeName;

                var keys = new INKeys(targetDrawer, driverField1, driverField2, driverField3, targetField4, targetField5, targetType, "WithCustProps");
                printf("%s\n", keys.toString());

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
                    if ((driverPropArrray[a].name == targetCPName) && include(SHARED_UPDATE_CPS, targetCPName))
                    {
                      //printf("#:%s dpan:%s tcpn:%s \n", a, driverPropArrray[a].name, targetCPName);
                      //printf("#:%s dpav:%s tcpv:%s\n", a, driverPropArrray[a].value, targetCPVal);
                      setProp = new INInstanceProp();
                      setProp.name = targetCPName;
                      setProp.setValue(driverPropArrray[a].value);
                      propsToUpdate.push(setProp);
                      break;
                    }  // end if (driverPropArrray[a].name == targetCPName)
                    if ((driverPropArrray[a].name == targetCPName) && include(SHARED_BLANK_CPS, targetCPName))
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

                if(!docToUpdate.setCustomProperties(propsToUpdate))
                {
                  debug.log("Couldn't set cus%s\n", getErrMsg());
                }
                else
                {
                  printf("Custom Property values been updated\n");
                }
                if (!reindexDocument(relinkID, keys, "APPEND"))
                {
                  printf("%s\n", getErrMsg());
                }
                else
                {
                  printf("Index values been updated\n");
                }   

                //writeToWFHistory(docToUpdate, triggerCreator);

              }// end if (sharedProp && sharedVal)
              else
              {
                printf("Unable to determine if this document is shared.\n");
                continue;
              }

            } // end while(Cur.next())


            //route driver to recycle bin
              // if relinking fails, add note to driver and route to error queue?
              // will we need an error flag?           

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
    printf("Couldn't get WFInfo for %s.  Error: %s\n", targetDoc.id, getErrMsg());
    return false;
  } // end writeToWFHistory

  var itemToNote = new INWfItem(itemWFInfo[0].id);

  if (!itemToNote.getInfo())
  {
    printf("Failed to get item. Error: %s.\n", getErrMsg());
    return false;
  }

  if (itemToNote.state != 1)  //GJ - is this step necessary?
  {
    printf("Item [%s]is currently being processed in workflow.  Cannot relink.\n", targetDoc.id);
    return false;
  }

  if (!itemToNote.setState(WfItemState.Idle, "This item was relinked via iScript by " + creator))
  {
    printf("Could not set state: %s.\n", getErrMsg());
    return false;
  }

  return true;
} // end writeToWFHistory

function include(arr, obj) {
    for(var i=0; i<arr.length; i++) {
        if (arr[i] == obj) return true;
    }
} // end include

//