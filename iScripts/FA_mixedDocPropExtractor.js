/****************************************************************** @fileoverview<pre>
 Name:		mixedDocPropExtractor.js
 Author(s):		Rajive - Umass
 Created:     	
 Last Updated:	
 For Version(s):	6.x
 Summary:		Extracts the mixed (index and custom) properties from work flow item.

 Business Use:	Designed to be an inscript on a workflow queue.

 Mod Summary:

</pre>*******************************************************************************/
// ********************* Open additional libraries ***********************************
#include "$IMAGENOWDIR6$\\script\\lib\\iScriptDebug.jsh"
#include "$IMAGENOWDIR6$\\script\\lib\\StoreDoc.jsh"
#include "$IMAGENOWDIR6$\\script\\lib\\spoolFileToBuffer.jsh"
#include "$IMAGENOWDIR6$\\script\\lib\\CreateOrRouteDoc.jsh"
#include "$IMAGENOWDIR6$\\script\\lib\\RouteItem.jsh"
#include "$IMAGENOWDIR6$\\script\\lib\\envVariable.jsh"
#include "$IMAGENOWDIR6$\\script\\lib\\UnifiedPropertyManager.jsh"
#include "$IMAGENOWDIR6$\\script\\lib\\Locking.jsh"
#include "$IMAGENOWDIR6$\\script\\lib\\csvObject.jsh"
#include "$IMAGENOWDIR6$\\script\\lib\\yaml_loader.jsh"
#include "$IMAGENOWDIR6$\\script\\lib\\HostDBLookupInfo.jsh"

// ********************* Initialize global variables *********************************

#define ERROR_Q "Checklist Update Error"
#define DATA_ENTRY_Q "Review for Completeness"
#define CAMPUS_ERR_Q " Data Entry Sorter Error"

//index Properties
#define Drawer_PROP "drawer"
#define Student_ID_PROP "folder"
//#define Student_Name_PROP "tab"
#define AidYear_PROP "f3"
#define Career_PROP "f4"
//#define Link_DateTime_PROP "f5"
#define Doc_Typ_PROP "docTypeName"

//Custom Properties
#define TaxYear_PROP "Tax Year"
#define Term_PROP "Term Code"
#define Gen_PROP "GEN"
#define Checklist_Update "Checklist Update"

var debug = null;
//Chris for 67 Upgrade
var DIRE = "D:\\inserver6\\script\\lockforcsv\\";
var FILEB = "Y:\\DI_"+ ENV_U3 +"_SA_FA_OUTBOUND\\BFA_Checklist_Updt.csv";
var FILEL = "Y:\\DI_"+ ENV_U3 +"_SA_FA_OUTBOUND\\LFA_Checklist_Updt.csv";
var FILED = "Y:\\DI_"+ ENV_U3 +"_SA_FA_OUTBOUND\\DFA_Checklist_Updt.csv";

#define LOG_TO_FILE true       //Set to true will log everything to log file
#define DEBUG_LEVEL 5           //Sets the log level - 5 is the highest, will display everything

// ********************* Body of Program *********************************************
/** **********************************************************************************
  *		Main body of script.
  *
  * @param {none} None
  * @returns {void} None
  ***********************************************************************************/
function main()
{
	try
	{
		debug = new iScriptDebug("FA_mixedDocPropExtractor", LOG_TO_FILE, DEBUG_LEVEL);
		debug.log("INFO", "FA_mixedDocPropExtractor script started.\n");

//		var wfItem = new INWfItem("301YX7Y_04JN75QBN0000GH");
		//var wfItem = new INWfItem("321YZ6T_077Y797L4000095");
		var wfItem = new INWfItem(currentWfItem.id);//"321YZ79_079E0NHQ20005FW");//"321YZ6T_077Y797L4000095");//
		if(!wfItem.id || !wfItem.getInfo())
		{
			debug.log("CRITICAL", " Couldn't get info for wfItem: %s\n", getErrMsg());
//			RouteItem(wfItem,SYSTEM_ERROR_QUEUE," Couldn't get info for wfItem");	
			return false;
		}

		var ScriptQName = wfItem.queueName;  // collection of info to build error queue
		var campus = ScriptQName.substring(0, 3);
 		debug.log("DEBUG","The campus code is: [%s] [%s]\n", campus, ScriptQName.indexOf("FA"));

		var errorQueueDE = ERROR_Q+" ("+campus+" "+DATA_ENTRY_Q+")";
        var errorQueueOTH = campus + CAMPUS_ERR_Q;
		debug.log("INFO", "Processing wfItem.type: [%s]\n",wfItem.type);
		foundQueue = false; 
		if(wfItem.type == 1) // doc type
		{
		    var doc = new INDocument(wfItem.objectId);
// 			var doc = new INDocument("301YX2W_03N29HSF50001TL");
//			var doc = new INDocument("301YWBJ_02P72NCV40001LR");
            if(!doc.id || !doc.getInfo())
            {
				debug.log("ERROR", "unable to get Document info for workflow item with id [%s]\n", wfItem.objectId);
				RouteItem(wfItem, errorQueueOTH, "unable to get info for Document");
				return false;
			}
			debug.log("INFO","Document [%s] ID: [%s]\n",doc.docTypeName, doc.id);
			var inProps = [Drawer_PROP,AidYear_PROP,Student_ID_PROP,Doc_Typ_PROP,TaxYear_PROP,Term_PROP,Gen_PROP,Checklist_Update];
			//var inProps = [Drawer_PROP - 0,AidYear_PROP - 1,Student_ID_PROP - 2,Doc_Typ_PROP - 3,TaxYear_PROP - 4,Term_PROP - 5,Gen_PROP - 6,Checklist_Update - 7];
			var upm = new UnifiedPropertyManager();
			var opPropValues = upm.GetAllProps(doc, inProps);
			
			if (opPropValues[7] && opPropValues[7] == 'Y') // Only perform check on Checklist Update = Y
			{
				debug.log("INFO", "Institution: [%s] Aid Year: [%s] Student ID: [%s] Doc Type: [%s] Checklist Update: [%s]\n",opPropValues[0],opPropValues[1],opPropValues[2],opPropValues[3],opPropValues[7]);	
				if (!(opPropValues[0] && opPropValues[1] && opPropValues[2] && opPropValues[3]))
           		{
					debug.log("ERROR", "Institution, Aid Year, Student ID, and Document Type: [%s],[%s],[%s],[%s],[%s].\n", opPropValues[0],opPropValues[1],opPropValues[2],opPropValues[3]);
					RouteItem(wfItem, errorQueueDE, "Data Entry Error: Application Number, Student ID, Document Type, App Center, and Career are mandatory fields.");
					return false;
				}
				//not sure if something like this is needed
				/*if(!validateData(opPropValues))
				{
					var errorReason ='';
					debug.log("ERROR", "Student ID(%s), SA Org ID(%s) or CEEB Code(%s) is invalid: \n",opPropValues[4].length,opPropValues[6].length,opPropValues[7].length);
				
					while (errorReason == '') // Assign the error message based on condition
					{
						if (opPropValues[4].length != 8)
						{
							errorReason = "Student ID is less than 8 characters";
							break;
						}
						if ((opPropValues[6] != '' && opPropValues[6].length != 8) && (opPropValues[7] != '' && (opPropValues[7].length !=4 && opPropValues[7].length != 6)))
						{
							errorReason = "Invalid SA Org ID AND CEEB character count";
							break;
						}
						if (opPropValues[6] != '' && opPropValues[6].length != 8) 
						{
							errorReason = "SA Org ID entered is less than 8 characters";
							break;
						}
						if (opPropValues[7] != '' && (opPropValues[7].length !=4 || opPropValues[7].length != 6))
						{
							errorReason = "CEEB code entered is not 4 or 6 characters";
							break;
						}
						else
						{
							errorReason = "Student Id, SA Org ID or CEEB Code value are invalid";
							break;
						}
					} // end while error assignment

					RouteItem(wfItem, errorQueueDE, "Data Entry Error: "+errorReason+".\n");
					return false;				
				}*/
			} // end if (opPropValues[9])
/*			if (!(opPropValues[7] == "Y" || opPropValues[7] == true || opPropValues[7])) // not needed anymore
            {
				debug.log("INFO", "The Checklist Update field was ‘no’ – document bypassed.\n");
//				RouteItem(wfItem, SYSTEM_ERROR_QUEUE.ERROR_QUEUE, "Application Number, Student ID and Document Type are mandatory fields:");
				return false;
			}			*/
//			var opPropValues[0].substring(2,3);
			switch (opPropValues[0].substring(2,3)) {
     			case "L":
         			opPropValues[0] = "UMLOW";
         			break;
     			case "B":
         			opPropValues[0] = "UMBOS";
     				break;
     			case "D":
         			opPropValues[0] = "UMDAR";
     				break;        			
     			default:
         			debug.log("INFO", "Switch over:\n");
 			}

 			if (opPropValues[7] == "Y")
 			{
				var writeSuccessful = writeToCSV(opPropValues);

				if(!writeSuccessful)
				{
					debug.log("DEBUG","csv would have been: %s,%s,%s,%s,%s,%s,%s,%s,%s\n",opPropValues[0],opPropValues[1],opPropValues[2],opPropValues[3],opPropValues[4],opPropValues[5],opPropValues[6],opPropValues[7]);
					debug.log("CRITICAL", "Writing to CSV unsuccessful: %s\n", getErrMsg());

					RouteItem(wfItem,errorQueueOTH,"Writing to CSV unsuccessful:");	
					//return false;
				}
			}
			else
			{
				debug.log("INFO", "Value of Checklist Update property is not 'Y': %s\n",opPropValues[7]);
				routeToReview(wfItem, doc, campus)	
			}				

//			var acv = {appNum, stuId, docTyp};
			var csv = opPropValues.join(",");
			debug.log("INFO", "csv: [%s]\n", csv);
			routeToReview(wfItem, doc, campus);
		}
		else 
		{
			debug.log("CRITICAL", "wfItem is not doc type: %s\n", getErrMsg());
			RouteItem(wfItem,errorQueueOTH," Couldn't get info for wfItem");	
			return false;
		}
	}

	catch(e)
	{
		debug.log("CRITICAL", "***********************************************\n");
		debug.log("CRITICAL", "***********************************************\n");
		debug.log("CRITICAL", "**                                           **\n");
		debug.log("CRITICAL", "**    ***    Fatal iScript ERROR!     ***    **\n");
		debug.log("CRITICAL", "**                                           **\n");
		debug.log("CRITICAL", "***********************************************\n");
		debug.log("CRITICAL", "***********************************************\n");
		debug.log("CRITICAL", "\n\n\n%s\n\n\n", e.toString());
		debug.log("CRITICAL", "\n\nThis script has failed in an unexpected way.  Please\ncontact Perceptive Software Customer Support at 800-941-7460 ext. 2\nAlternatively, you may wish to email support@imagenow.com\nPlease attach:\n - This log file\n - The associated script [%s]\n - Any supporting files that might be specific to this script\n\n", _argv[0]);
		debug.log("CRITICAL", "***********************************************\n");
		debug.log("CRITICAL", "***********************************************\n");
	}

	finally
	{
		debug.finish();
		return 0;
  	}
}

function validateData(dataToValidate)
{
	//will there be anything to validate?
	/*if(!(dataToValidate[4].length == 8  && (dataToValidate[6] == '' || dataToValidate[6].length == 8) && (dataToValidate[7] == '' || dataToValidate[7].length ==4 || dataToValidate[7].length == 6)))
	{
		return false;
	}*/
	return true;
}

function writeToCSV(opPropValues)
{ 
	var arrCSVConfig =
	[
		{name:'field1', multiple:false, massageFunction:csvObject.simpleTextFormat, massageConfig:{intWidth: 0, intAlign:-1, rgxInitial:/^.*$/, rgxFinal:/^.*$/}},
		{name:'field2', multiple:false, massageFunction:csvObject.simpleTextFormat, massageConfig:{intWidth: 0, intAlign:-1, rgxInitial:/^.*$/, rgxFinal:/^.*$/}},
		{name:'field3', multiple:false, massageFunction:csvObject.simpleTextFormat, massageConfig:{intWidth: 0, intAlign:-1, rgxInitial:/^.*$/, rgxFinal:/^.*$/}},
		{name:'field4', multiple:false, massageFunction:csvObject.simpleTextFormat, massageConfig:{intWidth: 0, intAlign:-1, rgxInitial:/^.*$/, rgxFinal:/^.*$/}},
		{name:'field5', multiple:false, massageFunction:csvObject.simpleTextFormat, massageConfig:{intWidth: 0, intAlign:-1, rgxInitial:/^.*$/, rgxFinal:/^.*$/}},
		{name:'field6', multiple:false, massageFunction:csvObject.simpleTextFormat, massageConfig:{intWidth: 0, intAlign:-1, rgxInitial:/^.*$/, rgxFinal:/^.*$/}},
		{name:'field7', multiple:false, massageFunction:csvObject.simpleTextFormat, massageConfig:{intWidth: 0, intAlign:-1, rgxInitial:/^.*$/, rgxFinal:/^.*$/}},
//		{name:'field8', multiple:false, massageFunction:csvObject.simpleTextFormat, massageConfig:{intWidth: 0, intAlign:-1, rgxInitial:/^.*$/, rgxFinal:/^.*$/}},
//		{name:'field9', multiple:false, massageFunction:csvObject.simpleTextFormat, massageConfig:{intWidth: 0, intAlign:-1, rgxInitial:/^.*$/, rgxFinal:/^.*$/}}
	];
	switch (opPropValues[0]){

		case "UMBOS":
			var objCSV = new csvObject(FILEB, arrCSVConfig, {intHeaderLen:0, delim:',', innerDelim:' '});
			break;
		case "UMLOW":
			var objCSV = new csvObject(FILEL, arrCSVConfig, {intHeaderLen:0, delim:',', innerDelim:' '});
			break;
		case "UMDAR":
			var objCSV = new csvObject(FILED, arrCSVConfig, {intHeaderLen:0, delim:',', innerDelim:' '});
			break;
	}


	if(!objCSV.openFile('a'))
	{
		debug.log('ERROR', "exampleWrite: unable to open CSV file for: [%s]\n", opPropValues[0]);
		return false;
	}
	var strFullofvalues = "";

	for(var i=0; i<opPropValues.length; i++)
	{
		if(i != 0)
			strFullofvalues = strFullofvalues + ", field" + (i+1) + ": '\"" + opPropValues[i] + "\"'";
		else
			strFullofvalues = strFullofvalues + "field" + (i+1) + ": '\"" + opPropValues[i] + "\"'";
	}

	debug.log('INFO', "object creation [%s]\n", strFullofvalues);

	var valueObj = eval( "({"+strFullofvalues+"})" );
	debug.log('INFO', "object creation [%s]\n", valueObj.field1);
	//var DIRE = "../script/lockforcsv";

	//end of that
	debug.log('INFO', "directory path to open [%s]\n", DIRE);
	
	var Lock = new Locking();
	if(!Lock.getLock(DIRE))
	{
		debug.log('ERROR', "Unable to obtain directory: [%s]\n", DIRE);
		return false;
	}

	//write csv object
	if(!objCSV.writeRowObject(valueObj))
	{
		debug.log('ERROR', "exampleWrite: unable to write to CSV file for: [%s]\n", opPropValues[0]);
	}
	if(!objCSV.closeFile())
	{
		debug.log('ERROR', "exampleWrite: unable to close CSV file for: [%s]\n", opPropValues[0]);
	}

	if(!Lock.Unlock(DIRE))
	{
		debug.log('ERROR', "Unable to Unlock directory: [%s]\n", DIRE);
		return false;		
	}
	
	return true;
}
/*

*/

//function to find to where a doc should be routed
function routeToReview(wfDoc, docObj, campus, errQ)
{
	debug.log("DEBUG","Preparing to route: [%s] from: [%s]\n", docObj, wfDoc.queueName);
	var desQ = "";

	var sql = "SELECT DP.PROCESS_CODE " +
	"FROM ISCRIPTUSER.DI_DOCT_PROCESS DP " +
	"INNER JOIN ISCRIPTUSER.PROCESS_DETAILS PD " +
	"ON DP.PROCESS_CODE = PD.PROCESS " +
	"WHERE DP.DOCTYPE_CODE = '"+docObj.docTypeName+"' " +
	"AND DP.IS_PRIMARY = 'Y'";

	var  returnVal = new Array();
	var cur = getHostDBLookupInfo_cur(sql,returnVal);

	if(!cur)
	{
		debug.log("WARNING","Unable to determine process code associated with document.  Routing to catch-all queue\n");
	}
	else
	{
		loadYAMLConfig(imagenowDir6+"\\script\\config_scripts\\FA_mixedDocPropExtractor\\");
		var procCode = cur[0];
		for (var campusConfig in CFG.FA_mixedDocPropExtractor)
      	{
      		var thisConfig = CFG.FA_mixedDocPropExtractor[campusConfig].CAMPUS_CONFIG;
      		for (var i = 0; i < thisConfig.length; i++)
      		{
      			var sourceQ = thisConfig[i].SOURCE_QUEUE;
      			var routingConfig = thisConfig[i].ROUTING_CONFIG;
      			
      			for (var j = 0; j < sourceQ.length; j++)
      			{
      				if(sourceQ[j].name == wfDoc.queueName)
      				{
      					for (var k = 0; k < routingConfig.length; k++)
      					{
      						if(procCode == routingConfig[k].process_code)
      						{
      							if(routingConfig[k].multi_subqueues)
      							{
      								debug.log("DEBUG","Process [%s] is configured for multiple subqueues.\n", procCode);
      								var docMatch = false;
      								for (var l = 0; l < routingConfig[k].subqueue_mapping.length; l++)
      								{
      									if(routingConfig[k].subqueue_mapping[l].doc_types.length == 0)
      									{
      										desQ = routingConfig[k].subqueue_mapping[l].destination_queue;
      										debug.log("DEBUG","Found destination queue: [%s]\n", desQ);
      										break;
      									}
      									else
      									{
	      									for (var m = 0; m < routingConfig[k].subqueue_mapping[l].doc_types.length; m++)
	      									{
	      										if(routingConfig[k].subqueue_mapping[l].doc_types[m] == docObj.docTypeName)
	      										{
	      											desQ = routingConfig[k].subqueue_mapping[l].destination_queue;
      												debug.log("DEBUG","Found destination queue: [%s]\n", desQ);
      												docMatch = true;
      												break;
	      										}// end if(routingConfig[k].subqueue_mapping[l].doc_types[m] == docObj.docTypeName)

	      									}// end for (var m = 0; m < routingConfig[k].subqueue_mapping[l].doc_types.length; m++)
      									}
      									if(docMatch)
	      								{
	      									break;
	      								}
      								}// end for (var l = 0; l < routingConfig[k].subqueue_mapping.length; l++)
      							}// end if(routingConfig[k].multi_subqueues)
      							else
      							{
      								desQ = routingConfig[k].destination_queue;
      								debug.log("DEBUG","Found destination queue: [%s]\n", desQ);
      								break;
      							}
      							
      						}// end if(procCode == routingConfig[k].process_code)
      					}// end for (var k = 0; k < routingConfig.length; k++)
      				}// end if(sourceQ[j].name == wfDoc.queueName)
      			}// end for (var j = 0; j < sourceQ.length; j++)
      		}// end for (var i = 0; i < thisConfig.length; i++)
      	}// end for (var campusConfig in CFG.FA_mixedDocPropExtractor)
	}// end else
	if(!desQ || desQ == null)
	{
		var otherQ = "Other (" + campus + " Review for Completeness)";
		debug.log("WARNING","Unable to find data entry queue for [%s].  Routing to [%s].\n", procCode, otherQ);
		RouteItem(wfDoc, otherQ, "No review to process mapping");
	}
	else
	{
		debug.log("DEBUG","Routing [%s] to [%s]\n", docObj, desQ);
		RouteItem(wfDoc, desQ, "Found " + desQ + " for " + procCode);
	}

}// end routeToReview
//

