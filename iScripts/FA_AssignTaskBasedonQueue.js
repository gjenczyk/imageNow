	/********************************************************************************
	Name:			FA_AssignTaskBasedOnQueue.js
	Author:			Rajive Rishi & Gregg Jenczyk
	Created:		06/09/15
	Last Updated:	
	For Version:	6.7
	Script Version:
---------------------------------------------------------------------------------
    Summary:

		
	Mod Summary:
		
    Business Use:  
		This script is designed to be executed via intool
		
********************************************************************************/
#include "$IMAGENOWDIR6$\\script\\lib\\iScriptDebug.jsh"
#include "$IMAGENOWDIR6$\\script\\lib\\commonSharedFunction.jsh"
#include "$IMAGENOWDIR6$\\script\\lib\\envVariable.jsh"
#include "$IMAGENOWDIR6$\\script\\lib\\RouteItem.jsh"
#include "$IMAGENOWDIR6$\\script\\lib\\CreateOrRouteDoc.jsh"
#include "$IMAGENOWDIR6$\\script\\lib\\AssignTaskToProject.jsh"
#include "$IMAGENOWDIR6$\\script\\lib\\yaml_loader.jsh"


// *********************         Configuration        *******************

#define CONFIG_VERIFIED		true	// set to true when configuration values have been verified
#define LOG_TO_FILE 		true	// false - log to stdout if ran by intool, true - log to inserverXX/log/ directory
#define DEBUG_LEVEL 		4		// 0 - 5.  0 least output, 5 most verbose   
     

// *********************       End  Configuration     *******************

// ********************* Initialize global variables ********************
//* Use this to Overide the Script such that you can run from intool *//

//currentWfItem = new INWfItem('301YW9L_02MSEZ7XS000009');

//End of that line


//This is the directory where the plan files will be created
var debug = "";
var intVslChunkSize = 3000;

// ********************* Include additional libraries *******************


/** ****************************************************************************
  *		Main body of script.
  *
  * @param {none} None
  * @returns {void} None
  *****************************************************************************/
function main ()
{

try{

  		debug = new iScriptDebug("FA_AssignTaskBasedonQueue", LOG_TO_FILE, DEBUG_LEVEL);
  		debug.log("INFO","Attempting to load YAML\n");
 
  		loadYAMLConfig(imagenowDir6+"\\script\\config_scripts\\FA_AssignTaskBasedonQueue\\");
		debug.log("INFO","My name is: FA_AssignTaskBasedonQueue.js\n");
		//		var wfQ = new INWfQueue(currentWfQueue.id);

/*   		if(typeof(currentWfItem) == "undefined")  //workflow
		{
			debug.log("CRITICAL", " This script is designed to run from workflow.\n");  //workflow
			return false;
		}
*/		
		//get wfItem info
		var wfItem = new INWfItem(currentWfItem.id);//"321YZ65_071KRS1460002H1");//
		if(!wfItem.id || !wfItem.getInfo())
		{
			debug.log("CRITICAL", "Couldn't get info for wfItem: %s\n", getErrMsg());
			RouteItem(wfItem,ERROR_QUEUE,"Couldn't get info for wfItem");	
			return false;
		}
		debug.log("INFO", "Processing wfItem.type: [%s]\n",wfItem.type);
		if (wfItem.type == 2)
		{
			debug.log("INFO","Project ID = [%s]\n",wfItem.objectId);
		}	
		debug.log("DEBUG", "Got Workflow Item Info Processing Items\n");
		foundQueue = false;
	
	for (var sourceQConfig in CFG.FA_AssignTaskBasedonQueue)
	{	
		debug.log("DEBUG","I'm now in the First For Loop\n");
		var CAMPUS_CONFIG = CFG.FA_AssignTaskBasedonQueue[sourceQConfig].CAMPUS_CONFIG;
		//printf("CAMPUS_CONFIG.length" + CAMPUS_CONFIG.length);
		debug.log("DEBUG","Elements in YAML File [%s]\n",CAMPUS_CONFIG.length);
		for(var len = 0; len < CAMPUS_CONFIG.length; len ++)
		{ 
			 debug.log("DEBUG","I am now in the second for loop\n");
			 debug.log("DEBUG","Comparing WFQName [%s] to Config Source QName [%s]\n", wfItem.queueName, CAMPUS_CONFIG[len].SOURCE_QUEUE);
			 //printf("CAMPUS_CONFIG.length" + wfItem.queueName + "---" + CAMPUS_CONFIG[len].SOURCE_QUEUE);
		     if(wfItem.queueName == CAMPUS_CONFIG[len].SOURCE_QUEUE)
			 { 			debug.log("INFO", "WFQName:  %s and  Config Source QName is: %s\n", wfItem.queueName,CAMPUS_CONFIG[len].SOURCE_QUEUE);
						if(wfItem.type == 2)
						{
						    var proj = new INFolder(wfItem.objectId);
                            if(!proj.id || !proj.getInfo())
		                    {
								debug.log("ERROR", "unable to get Project info for workflow item with id [%s]\n", wfItem.objectId);
								RouteItem(wfItem, CAMPUS_CONFIG[len].ERROR_QUEUE, "unable to get info for PROJECT");
								return false;
							}	
							//If CHECK_FOR_EXISTING_TEMPLATE = False = the script will just assign the task
							if(CAMPUS_CONFIG[len].CHECK_FOR_EXISTING_TEMPLATE==false)
							{
							  debug.log("DEBUG", "Check for Existing Template Config is false\n");
							  debug.log("INFO", "Adding the following task: [%s] to the following project: [%s]\n",CAMPUS_CONFIG[len].TASK_TEMPLATE, proj.id);	
							  AssignTaskToProject(proj,CAMPUS_CONFIG[len].TASK_TEMPLATE,CAMPUS_CONFIG[len].ERROR_QUEUE,CAMPUS_CONFIG[len].ERROR_REASON_FAIL_TO_ADD_TASK_TEMPLATE)
							}
							else
							{
									var taskArray = new Array();
									if (!INTask.getTasks(proj.id, "", "", taskArray))
									{
										debug.log("ERROR", "getTasks failed:  %s\n", getErrMsg());
										RouteItem(wfItem, CAMPUS_CONFIG[len].ERROR_QUEUE, "unable to get info for TASK");
										return false;
									}
									else
									{
										debug.log("INFO", "Tasks Found: " + taskArray.length +"\n");
										if(taskArray.length==0)
										{
										  // add the task
										  debug.log("INFO", "Adding the following task: [%s] to the following WfItem(Project): [%s]\n",CAMPUS_CONFIG[len].TASK_TEMPLATE, wfItem.id);
										  AssignTaskToProject(proj,CAMPUS_CONFIG[len].TASK_TEMPLATE,CAMPUS_CONFIG[len].ERROR_QUEUE,CAMPUS_CONFIG[len].ERROR_REASON_FAIL_TO_ADD_TASK_TEMPLATE)
										}
										else
										{
										   
												var taskTemp = new INTaskTemplate();
												taskTemp.name = CAMPUS_CONFIG[len].TASK_TEMPLATE;
												
												if(!taskTemp.getInfo())
												{
												  debug.log("ERROR","taskTemp.getInfo failed: %s\n", getErrMsg());
												  RouteItem(wfItem,CAMPUS_CONFIG[len].ERROR_QUEUE,CAMPUS_CONFIG[len].ERROR_REASON_FAIL_TO_ADD_TASK_TEMPLATE);
												  return false;
												}
												templateID = taskTemp.id;							   
												foundCompleteTask = 0;
												foundTaskNotComplete = 0;
												debug.log("INFO", "%s\n", templateID);
												for(var x = 0; x < taskArray.length; x++)
												{
													debug.log("INFO", "\nTask["+x+"]: \n");
													debug.log("INFO","\n\nTask: %s \n", taskArray[x].toString());
													debug.log("INFO", "taskArray[x].TaskTemplateID := %s\n", taskArray[x].taskTemplateID);
													debug.log("INFO", "taskArray[x].State := %s\n",taskArray[x].state);
													
													if(templateID === taskArray[x].taskTemplateID)
													{
															if(taskArray[x].state === "Complete")
															{
															  debug.log("INFO", " COMPLETE ");
															  foundCompleteTask = Math.round(foundCompleteTask) + 1;
															}
															else
															{
															 debug.log("INFO", "NOT COMPLETE ");
															 foundTaskNotComplete = Math.round(foundTaskNotComplete) + 1;
															}
													}
													debug.log("INFO", "---\n");
												}
												debug.log("INFO", "\nfoundCompleteTask:= %s: \t foundTaskNotComplete:= %s \n",foundCompleteTask,foundTaskNotComplete);
												if(foundTaskNotComplete >0)
												{
												  continue;
												}
												
												 debug.log("INFO", "Check for Existing Template Config is false,adding the following task: [%s] to the following project: [%s]\n",CAMPUS_CONFIG[len].TASK_TEMPLATE, proj.id);	
												 AssignTaskToProject(proj,CAMPUS_CONFIG[len].TASK_TEMPLATE,CAMPUS_CONFIG[len].ERROR_QUEUE,CAMPUS_CONFIG[len].ERROR_REASON_FAIL_TO_ADD_TASK_TEMPLATE)
										}
							
							
							
									}						
							
							}	           
							
						}
						else
						  {
								 debug.log("ERROR", "THE WFITEM is not a Project\n");
								 RouteItem(wfItem,CAMPUS_CONFIG[len].ERROR_QUEUE, "THE WFITEM is not a Project");
								return false;
						  }
						  foundQueue = true;
						  break;
						  
			}
				
	 	}// end of for
	}
	 if(foundQueue == false)
	 {
	 	debug.log("INFO","Could Not find a match for [%s] --> to YAML Source QName Check YAML - this may or may not be an issue\n",wfItem.queueName);
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

  
