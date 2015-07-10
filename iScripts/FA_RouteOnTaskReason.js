/********************************************************************************
	Name:			FA_RouteOnTaskReason.js
	Author:			Rajive Rishi & Gregg Jenczyk
	Created:
	Last Updated:
	For Version:	6.7
	Script Version:
---------------------------------------------------------------------------------
    Summary:

	Mod Summary:

    Business Use:
		This script is designed to be executed via inbound action 
		and Route on the basis of the reason defined.

********************************************************************************/

#include "$IMAGENOWDIR6$\\script\\lib\\iScriptDebug.jsh"
#include "$IMAGENOWDIR6$\\script\\lib\\commonSharedFunction.jsh"
#include "$IMAGENOWDIR6$\\script\\lib\\envVariable.jsh"
#include "$IMAGENOWDIR6$\\script\\lib\\RouteItem.jsh"
#include "$IMAGENOWDIR6$\\script\\lib\\CreateOrRouteDoc.jsh"
#include "$IMAGENOWDIR6$\\script\\lib\\HostDBLookupInfo.jsh"
#include "$IMAGENOWDIR6$\\script\\lib\\yaml_loader.jsh"
#include "$IMAGENOWDIR6$\\script\\lib\\UnifiedPropertyManager.jsh"
#include "$IMAGENOWDIR6$\\script\\lib\\GetProp.jsh"

// *********************         Configuration        *******************

#define CONFIG_VERIFIED		true	// set to true when configuration values have been verified
#define LOG_TO_FILE 		true	// false - log to stdout if ran by intool, true - log to inserverXX/log/ directory
#define DEBUG_LEVEL 		5		// 0 - 5.  0 least output, 5 most verbose

// *********************       End  Configuration     *******************

// ********************* Initialize global variables ********************
/* Use this to Overide the Script such that you can run from intool */

//currentWfItem = new INWfItem('301YWBD_02P71ZCV300006T');

//End of that 

//This is the directory where the plan files will be created
//var SYSTEM_ERROR_QUEUE = "Error Queue";

//
var debug = "";
var templateIDs = [];
// ********************* Include additional libraries *******************


/** ****************************************************************************
  *		Main body of script.
  *
  * @param {none} None
  * @returns {void} None
  *****************************************************************************/
function main ()
{
	try
	{

	  	debug = new iScriptDebug("FA_RouteOnTaskReason", LOG_TO_FILE, DEBUG_LEVEL);

	  	loadYAMLConfig(imagenowDir6+"\\script\\config_scripts\\FA_RouteOnTaskReason\\");

	   	if(typeof(currentWfItem) == "undefined")  //workflow
		{
			debug.log("CRITICAL", " This script is designed to run from workflow.\n");  //workflow
			return false;
		}

		//get wfItem info
		var wfItem = new INWfItem(currentWfItem.id);//"321YZ65_071KRS1460002H1");//

		if(!wfItem.id || !wfItem.getInfo())
		{
			debug.log("CRITICAL", "Couldn't get info for wfItem: %s\n", getErrMsg());
			RouteItem(wfItem,SYSTEM_ERROR_QUEUE," Couldn't get info for wfItem");
			return false;
		}
		debug.log("INFO", "Processing wfItem.type: [%s]\n",wfItem.type);
		foundQueue = false;

		if(wfItem.type == 2)
		{
		    var proj = new INFolder(wfItem.objectId);//"301YWCH_02YVKE07B000043");

		    debug.log("INFO", "Project ID is: [%s] \n",proj.id);
		    var sourceQ = wfItem.queueName;
            var errorQ = sourceToErrorQ(wfItem.queueName);
            debug.log("INFO", "errorQ [%s] \n", errorQ);
			debug.log("INFO", "templateIDs [%s] \n", templateIDs.length);

           if(!proj.id || !proj.getInfo())
            {
				debug.log("ERROR", "unable to get Project info for workflow item with id [%s]\n", wfItem.objectId);
				RouteItem(wfItem, errorQ, "unable to get info for Project");
				return false;
			}
			var taskArray = new Array();
			if (!INTask.getTasks(proj.id, "", "", taskArray))
			{
				debug.log("INFO", "getTasks failed for Project ID: [%s]\n", proj.id);
				RouteItem(wfItem, errorQ, "unable to get info for document");
				return false;
			}
			else
			{
				debug.log("INFO", "Tasks Found: " + taskArray.length +"\n");
				if(taskArray.length==0)
				{
				 debug.log("INFO", "getTasks failed:  [%s]\n", getErrMsg());
				 RouteItem(wfItem, errorQ, "unable to get info for document");
				 return false;
				}
				else
				{
					var highestTimeIndex = -1;
					var tempIdForName = "";
					highestTime = -1;
					for (var x = taskArray.length - 1; x >= 0; x--) 
					{
						debug.log("INFO", "Task["+x+"]: \n");
						debug.log("INFO","templateIDs [%s] \n", templateIDs.toString());
						debug.log("INFO", "taskArray[x].TaskTemplateID := [%s]\n", taskArray[x].taskTemplateID);
						debug.log("INFO", "taskArray[x].State := [%s]\n",taskArray[x].state);
						//printf("yyy"+taskArray[x].assignmentDate+":"+taskArray[x].taskTemplateID+":"+"\n");
						if(templateIDs.contains(taskArray[x].taskTemplateID))
						{ //printf("xxxxxxxxxxxxxxxx"+taskArray[x].assignmentDate);
								 if (!(highestTime >= taskArray[x].assignmentDate) && checkTemplateId(taskArray[x].taskTemplateID))
								 {
								  highestTimeIndex = x;
								  highestTime = taskArray[x].assignmentDate;
								  debug.log("INFO","second TIME IS [%s]\n",taskArray[x].assignmentDate.toString());
								  debug.log("INFO", "second Highest Time = %f \n", highestTime);
								  tempIdForName = taskArray[x].taskTemplateID;
								 }
								 
								 debug.log("INFO", " COMPLETE highestTime =%f \n taskArray[x].completionDate =%f \n  state:= [%s] %d\n",highestTime,taskArray[x].completionDate,taskArray[x].state,highestTimeIndex);
						}
						debug.log("INFO", "Highest time: [%s] and highest index: [%s]\n",highestTime, highestTimeIndex);											
					};

					if(	highestTime == Math.round(-1) || highestTimeIndex == Math.round(-1) )
					{
						debug.log("CRITICAL","[%s]\n", getErrMsg());
						RouteItem(wfItem,errorQ,"Unable to find a completed task");
						return false;
					}
					else if(taskArray[highestTimeIndex].state != "Complete")
					{
						debug.log("CRITICAL","[%s]\n", getErrMsg());						
						RouteItem(wfItem,errorQ,"Latest assigned task is not complete");
						return false;
					} 

					debug.log("INFO", "---highestTimeIndex %d [%s] -- \n",highestTimeIndex,taskArray[highestTimeIndex].state);
					taskID = taskArray[highestTimeIndex].id;

				    debug.log("DEBUG", "Check reason code for the task ID [%s]\n",taskID);
					var sql = "SELECT ITEM_NAME FROM (SELECT T.task_id , T.TASK_TEMPLATE_ID, T.TASK_STATE, T.COMPLETION_TIME, I.ITEM_NAME, ROW_NUMBER() OVER (PARTITION BY T.task_id ORDER BY T.task_id) AS Rank FROM inuser.in_task T INNER JOIN (SELECT inuser.in_task_hist.task_id, MAX(inuser.in_task_hist.MOD_SEQ_NUM) AS MOD_SEQ_NUM FROM inuser.in_task_hist GROUP BY inuser.in_task_hist.task_id) H ON T.task_id = H.task_id INNER JOIN inuser.in_task_hist PH ON H.task_id = PH.task_id AND H.MOD_SEQ_NUM = PH.MOD_SEQ_NUM LEFT OUTER JOIN inuser.in_list_item I ON I.list_item_id = PH.reason_id) AllRows WHERE Rank =1 AND task_id IN ('"+taskID+"')";

					var  returnVal = new Array(1);
					getHostDBLookupInfo(sql,returnVal);

					if(!returnVal)
					{
						debug.log("ERROR", "Failed to get check number from dblookup with query: [%s].\n", sql);
						RouteItem(wfItem,errorQ,"Cannot find a matching reason in the query");
						return false;
					}
					debug.log("DEBUG", "RETURN VALUE := [%s] \n", returnVal);
					debug.log("DEBUG", "RETURN VALUE 0  := [%s]\n", returnVal[0]);
					var reasonCode = returnVal[0];
					var foundReason = true;
					ConvTempName = getTaskNamefromID(tempIdForName);
					var routingQ = sourceToReason(sourceQ, reasonCode, ConvTempName);
					debug.log("DEBUG", "RETURN VALUE  := [%s]\n", routingQ);
					if(!routingQ)
					{
						foundReason = false;
						debug.log("ERROR", "reason code not found: [%s].\n", sql);
						RouteItem(wfItem,errorQ,"reason code not found");
						return false;							
					}
					if(foundReason==false)
					{
					  debug.log("ERROR", "Unable to find the reason code in the config .\n");
					  RouteItem(wfItem,errorQ,"Cannot find a template/Reason");
					  return false;
					}
					if(foundReason==true)
					{
					  //gregg did this ---->
					  var upmProj = new UnifiedPropertyManager();
					  var projCPs = upmProj.GetAllCustomProps(proj);
					  for (var k = projCPs.length-1;k>=0;k--)
					  {
					 	if (projCPs[k].name === "Task Reason")
						{
							debug.log('INFO','projCPs.name is [%s]\n',projCPs[k].name);
						 	debug.log('INFO','projCPs.value is [%s]\n',projCPs[k].getValue());
							
							if(!projCPs[k].setValue(reasonCode))
							{
								debug.log('INFO','Unable to update task reason\n');
								RouteItem(wfItem,errorQ,"Could not update task reason");

							} else {
								debug.log('INFO','Proposed new Task Reason is [%s]\n',projCPs[k].getValue());
								if (proj.setCustomProperties(projCPs))
								{
									debug.log("INFO", "Updated Task Reason to [%s]\n",reasonCode);
								}
								else
								{
									debug.log("INFO", "Setting custom Properties on project failed\n");
									RouteItem(wfItem,errorQ,"Could not update task reason");
								}
							}
						}

					  }//<----- to this!

					  debug.log("INFO", "Found Reason routing to appropriate queue.\n");
					  RouteItem(wfItem,routingQ,reasonCode);
					  replace_document(taskArray,proj);
					  return false;
					}
				}
			}
		}
		else
	  	{
			debug.log("CRITICAL", " The Document routed is of wrong type [%s]\n", getErrMsg());
			RouteItem(wfItem,SYSTEM_ERROR_QUEUE,"wfItem is not a project");
			return false;
	  	}
	  	foundQueue = true;

		if(foundQueue==false)
		{
		   	debug.log("CRITICAL", " The queueName does not match [%s]\n", getErrMsg());
			RouteItem(wfItem,SYSTEM_ERROR_QUEUE," Couldn't get info for wfItem");
			return false;
		}
	}
	catch(e)
	{
		if(!debug)
		{
			printf("\n\nFATAL iSCRIPT ERROR: [%s]\n\n", e.toString());
		}
		else
		{
			debug.log("CRITICAL", "***********************************************\n");
			debug.log("CRITICAL", "***********************************************\n");
			debug.log("CRITICAL", "**                                           **\n");
			debug.log("CRITICAL", "**    ***    Fatal iScript Error!     ***    **\n");
			debug.log("CRITICAL", "**                                           **\n");
			debug.log("CRITICAL", "***********************************************\n");
			debug.log("CRITICAL", "***********************************************\n");
			debug.log("CRITICAL", "\n\n\n%s\n\n\n", e.toString());
			debug.log("CRITICAL", "\n\nThis script has failed in an unexpected way.  Please\ncontact Perceptive Software Customer Support at 800-941-7460 ext. 2\nAlternatively, you may wish to email support@imagenow.com\nPlease attach:\n - This log file\n - The associated script [%s]\n - Any supporting files that might be specific to this script\n\n", _argv[0]);
			debug.log("CRITICAL", "***********************************************\n");
			debug.log("CRITICAL", "***********************************************\n");
		}
	}
	finally
	{
		if(debug)
		{
			debug.finish();
		}
	}
}


function replace_document(taskarray,proj)
{
	
	var tasklist = [];

	for (var x = 0; x < taskarray.length; x++) 
	{
		tasklist[x] = taskarray[x].id;
	}
	
	var params = "'"+tasklist.join("','")+"'";

	//debug.log("INFO","params"+params+"\n");
    //This is where you specify an image//

    var nam4file = proj.folderTypeName;
	var name4filepath = nam4file.substr(-3,3);
    var filePath = imagenowDir6+"\\script\\"+name4filepath+"_ProjectTasksReport.txt";
    //debug.log("DEBUG","filePath is [%s]\n",filePath);

    ptr = Clib.fopen(filePath, "w");
    var sql = "SELECT task_id, TASK_TEMPLATE_ID,TASK_TEMPLATE_NAME,TASK_STATE,USR_FIRST_NAME,"+
    "USR_LAST_NAME,TO_CHAR(new_time(COMPLETION_TIME,'GMT','EST'),'DD-MON-YY'),"+
    "TO_CHAR(new_time(ASSIGN_TIME,'GMT','EST'),'DD-MON-YY'),COMPLETION_USR_ID,CREATION_USR_ID,"+
    "comment_text,SEQ_NUM,ITEM_NAME FROM (SELECT T.task_id , T.TASK_TEMPLATE_ID,TTMPL.TASK_TEMPLATE_NAME,"+
    "T.TASK_STATE, USR.USR_FIRST_NAME, USR.USR_LAST_NAME,T.COMPLETION_TIME, T.ASSIGN_TIME,"+ 
    "T.COMPLETION_USR_ID,TCMT.CREATION_USR_ID,TCMT.comment_text, TCMT.SEQ_NUM,I.ITEM_NAME "+
    "FROM inuser.in_task T inner join inuser.in_task_template TTMPL on T.TASK_TEMPLATE_ID = TTMPL.TASK_TEMPLATE_ID "+
    " AND T.completion_usr_id IS NOT NULL left outer join inuser.IN_TASK_COMMENT TCMT on TCMT.TASK_ID = t.task_id and "+
    " T.completion_usr_id is not null left outer join inuser.in_sc_usr USR on USR.usr_id = TCMT.CREATION_USR_ID "+
    " OR USR.usr_id = T.COMPLETION_USR_ID INNER JOIN (SELECT inuser.in_task_hist.task_id, MAX(inuser.in_task_hist.MOD_SEQ_NUM) AS MOD_SEQ_NUM "+
    "FROM inuser.in_task_hist GROUP BY inuser.in_task_hist.task_id) H ON T.task_id = H.task_id "+
	"INNER JOIN inuser.in_task_hist PH ON H.task_id = PH.task_id AND H.MOD_SEQ_NUM = PH.MOD_SEQ_NUM "+
	"LEFT OUTER JOIN inuser.in_list_item I ON I.list_item_id = PH.reason_id) where task_id IN "+
	"("+params+") and "+
//		'301YX3M_02WQD63EY000YP2','301YX2M_02QGJ9EYM0008R8','301YX7R_03NP4NXNL007ES5'
	"TASK_TEMPLATE_ID not in ('301YX1B_02K9RF9M200012W','301YX1W_02NJFMGSC00006Z','301YX1B_02K9RG9M20000KE',"+
	"'301YX1B_02K9RB9M2000351') order by COMPLETION_TIME, TASK_ID, SEQ_NUM";

	var  returnVal = new Array();
	var cur = getHostDBLookupInfo_cur(sql,returnVal);

	if(!cur)
	{
		debug.log("ERROR", "Failed to get check number from dblookup with query: [%s].\n", sql);
		return false;
	}
	else
	{
		var taskId_orig = "";
		var comments = [];
		var comment = "";
		var commentText = "";
		var completedBy = ""
		var commentedUser =""
		var text2write = "";
		var templateName = "";
		var dateTime ="";
		var reason = "";
		var k = 0;
		var initial = 0;
		var pos = 0;
		while(cur.next())
		{
			var taskId = cur[0].toString();
			if (taskId_orig != "" && taskId_orig != taskId)
			{	
				if(initial == 0)
				{
					line = templateName + "\nCompleted By: "+ completedBy + "\nOn: "+ dateTime + "\nReason: "+ reason;
					initial++;
				}
				else
				{
					line = "\n\n"+templateName + "\nCompleted By: "+ completedBy + "\nOn: "+ dateTime + "\nReason: "+ reason;
					debug.log("INFO","line is [%s]\n",line);
				}
				//Clib.fsetpos(ptr, pos)
//				printf("00000000000000"+Clib.feof(ptr));				
				Clib.fputs(line,ptr);
				//Clib.fgetpos(ptr, pos);
//				printf("99999999999999"+Clib.feof(ptr));
//				var eofPtr = Clib.fseek(ptr, -1, "SEEK_END");
//				printf("------"+comments.length);
				for(var itr=0;itr<comments.length;itr++)
				{
					line = "\n" + comments[itr];
					//Clib.fsetpos(ptr, pos)
					Clib.fputs(line,ptr);
					//Clib.fgetpos(ptr, pos);
				}				
				k = 0;
				comments = [];
			}			
			taskId_orig = taskId;
			templateName = cur[2]+"";
			dateTime = cur[6]+"";
			reason = cur[12]+"";
			completedBy = cur[4]+" "+cur[5];
			commentedUser = cur[4]+" "+cur[5];
			commentText = cur[10];
			comment = commentedUser + " commented: " + commentText+"";
			if(commentText !== null && commentText !== "")
			{
				comments[k] = comment;
				k++;
			}
			pos++;
		}
		if (pos === 0)
		{
			debug.log("INFO", "No data returned - Double check the Template Id's.\n");
		}
		if (initial == 0)
		{
			line = templateName + "\nCompleted By: "+ completedBy + "\nOn: "+ dateTime + "\nReason: "+ reason;			
		}
		else
		{
			line = "\n\n"+templateName + "\nCompleted By: "+ completedBy + "\nOn: "+ dateTime + "\nReason: "+ reason;
		}
		Clib.fputs(line,ptr);
//		var eofPtr = Clib.fseek(ptr, 0, "SEEK_END");
//		printf("------"+comments.length);
		for(var itr=0;itr<comments.length;itr++)
		{
			line = "\n" + comments[itr];
			//Clib.fsetpos(ptr, pos)
			Clib.fputs(line,ptr);
			//Clib.fgetpos(ptr, pos);
		}
		Clib.fclose(ptr);


		var projTypeNam = proj.folderTypeName;
		var drawer ="UM"+projTypeNam.substr(-3,3);
		var docType = "Task Sheet " + projTypeNam.substr(-3,3);
//		printf("proj.name"+proj.name+"<-------->"+"proj.projTypeName"+proj.projTypeName+"--------------");
		var projName = proj.name; //studentId App: F4 
		var folder = projName.substring(0,8);
		var f3 = projName.substr(-8,5);

		var inProps = ["Student ID", "Student Name", "Career"];
		//var inProps = [Drawer_PROP,Career,App_Center,Student_ID_PROP,Application_Num_PROP,Doc_Typ_PROP,Recmdr_Name_PROP,Ceeb_Code_PROP,Org_Id_PROP];
		var upm = new UnifiedPropertyManager();
		var opPropValues = upm.GetAllProps(proj, inProps);
//		printf(">>>>>>>>>>>"+opPropValues[2]+"<<<<<<"+opPropValues[1]+">>>>>>>>>>>>>>>>>>>>>>>>");
		var d = new Date();
		var hours = "";
		var dnight = "";
		if (d.getHours() > 12)
		{
			hours = d.getHours()-12;
			dnight = "pm";
		}
		else
		{
			hours = d.getHours();
			dnight = "am";
		}

		var f5 = (make2digit(d.getMonth()+1))+"/"+make2digit(d.getDate())+"/"+d.getFullYear()+" "+make2digit(hours)+":"+make2digit(d.getMinutes())+":"+make2digit(d.getSeconds())+" "+dnight;
//check if doc already exist
		var docs = proj.getDocList();
		for (var docitr=0;docitr<docs.length;docitr++)
		{
			var docid2rem = docs[docitr].id;
//			printf(docid2rem+"|||||||||||||||"+docs[docitr].docTypeName+"--"+docType+"<<<<<|||")
			if (docs[docitr].docTypeName === docType)
			{
				proj.removeDocument(docid2rem);

				var remfromsystem = INDocument(docid2rem);
			 	remfromsystem.getInfo();
			 	if(remfromsystem.remove())
		        {
		            debug.log("INFO","Removed document: [%s]\n", remfromsystem.id);
		        }
		        else
		        {
		            debug.log("INFO","could not remove document: [%s]\n", getErrMsg());
		        }
			}
		}
//check over
		var doc = new INDocument(drawer, folder, opPropValues[1], f3, opPropValues[2], f5, docType);
	    var attr = new Array();
	    attr["phsob.file.type"] = "txt";
	    attr["phsob.working.name"] = "TaskSummary.txt";
	    attr["phsob.source"]="PhsobSource.Iscript";
	    var logob = doc.storeObject(filePath, attr);
	    if(!doc.getInfo())
	    {
	    	debug.log("ERROR","errMsg= [%s]\n", getErrMsg());
	    }
//		printf("docid ++++++"+doc.id);
	    if(!proj.addDocument(doc.id))
		{
		     debug.log("ERROR","Failed to add document [%s] to project [%s].\n", doc.id, proj.name);
		     debug.log("ERROR","errMsg= [%s]\n", getErrMsg());
		}
	}   
}

function make2digit(num)
{
	if(num<10)
	{
		num = "0"+num.toString();
	}
	return num;
}

function getTaskNamefromID(tempIdForName)
{
	for (var i=templateIDs.length-1 ; i>=0; i--)
	{
		if(templateIDs[i][0].toString() === tempIdForName)
		{
			return templateIDs[i][1];
		}
	}
}

function templateIdFromName(templateName)
{		debug.log("INFO", "new templateName name: [%s]\n", templateName);
		var newtaskTemp = new INTaskTemplate();
		newtaskTemp.name = templateName;
		if (!newtaskTemp.getInfo())
		{
			debug.log("ERROR", "Couldn't get task template info: [%s]\n", getErrMsg());
		}
		debug.log("INFO", "new task name: [%s]\n", newtaskTemp.id);
		return newtaskTemp.id;
}


function sourceToErrorQ(sourceQ)
{
		for (var sourceQConfig in CFG.FA_RouteOnTaskReason)
		{ debug.log("INFO","sourceQConfig: [%s] and  sourceQ: [%s] and CFG.QUEUE = [%s]\n", sourceQConfig, sourceQ,CFG.FA_RouteOnTaskReason[sourceQConfig].QUEUE);
			if(CFG.FA_RouteOnTaskReason.hasOwnProperty(sourceQConfig) && CFG.FA_RouteOnTaskReason[sourceQConfig].QUEUE ===  sourceQ)
			{
				
				debug.log("INFO","***************sourceQConfig: [%s] and  sourceQ: [%s] and CFG.QUEUE = [%s]\n", sourceQConfig, sourceQ,CFG.FA_RouteOnTaskReason[sourceQConfig].QUEUE);
				debug.log("INFO","CFG.FA_RouteOnTaskReason[sourceQConfig].ROUTING.length: [%s]\n", CFG.FA_RouteOnTaskReason[sourceQConfig].ROUTING.length);
				for(var i = 0; i < CFG.FA_RouteOnTaskReason[sourceQConfig].ROUTING.length; i++)
				{
					var temptemplateID = [];
					var templateName = CFG.FA_RouteOnTaskReason[sourceQConfig].ROUTING[i].TASK_TEMPLATE;
					debug.log("DEBUG","templateName = [%s]\n", templateName);
					temptemplateID.push(templateIdFromName(templateName));
					//templateIDs[i][0] = templateIdFromName(templateName);
					temptemplateID.push(templateName);
					templateIDs.push(temptemplateID);
					debug.log("INFO","templateName: [%s] and  templateID: [%s]\n", templateName, templateIDs[i])
				}	 
				return CFG.FA_RouteOnTaskReason[sourceQConfig].ERROR_QUEUE;	
			}
		}
}

function sourceToReason(sourceQ, reasonCode, ConvTempName)
{
	for (var sourceQConfig in CFG.FA_RouteOnTaskReason)
	{debug.log("INFO", " queue:[%s] sourceQ: [%s]\n",CFG.FA_RouteOnTaskReason[sourceQConfig].QUEUE,sourceQ);
		if(CFG.FA_RouteOnTaskReason.hasOwnProperty(sourceQConfig) && CFG.FA_RouteOnTaskReason[sourceQConfig].QUEUE ===  sourceQ)
		{debug.log("INFO", " inside if CFG.FA_RouteOnTaskReason[sourceQConfig].ROUTING.length: [%s] \n",CFG.FA_RouteOnTaskReason[sourceQConfig].ROUTING.length);
			for(var i = 0; i < CFG.FA_RouteOnTaskReason[sourceQConfig].ROUTING.length; i++)
			{	debug.log("INFO", " inside if CFG.FA_RouteOnTaskReason[sourceQConfig].ROUTING queue.length: [%s]\n",CFG.FA_RouteOnTaskReason[sourceQConfig].ROUTING[i].ROUTE_QUEUE.length);
				debug.log("INFO","ConvTempName: [%s]\n",ConvTempName);
				if(CFG.FA_RouteOnTaskReason[sourceQConfig].ROUTING[i].TASK_TEMPLATE === ConvTempName)
				{
					for(var j = 0; j < CFG.FA_RouteOnTaskReason[sourceQConfig].ROUTING[i].ROUTE_QUEUE.length; j++)
					{debug.log("INFO", " value of i: [%s] our route reason: [%s] reasonCode: [%s] \n",i, CFG.FA_RouteOnTaskReason[sourceQConfig].ROUTING[i].ROUTE_QUEUE[j].ROUTE_REASON,reasonCode);
						if(CFG.FA_RouteOnTaskReason[sourceQConfig].ROUTING[i].ROUTE_QUEUE[j].ROUTE_REASON == reasonCode)
						{
							return CFG.FA_RouteOnTaskReason[sourceQConfig].ROUTING[i].ROUTE_QUEUE[j].ROUTE_QUEUE_NAME; 
						}	
					}
				}
			} 
		}
	}
}

function checkTemplateId(tempId)
{
	return templateIDs.contains(tempId);
}

Array.prototype.contains = function(obj) {
	var i = this.length;
    while (i--)
    {
    	if (this[i][0] === obj)
    	{
        	return true;
        }
    }     
    return false;
} 



