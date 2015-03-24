/********************************************************************************
        Name:          CompleteTaskAndMove.js
        Author:        Gregg Jenczyk
        Created:        08/22/14
        Last Updated:   
        For Version:    
---------------------------------------------------------------------------------
        Summary:
               This will take every folder in a workflow queue, complete a specified
               task with a specified reason, and route the folder to a destination 
               queue.
               
        Mod Summary:
               Date-Initials: Modification description.
               
********************************************************************************/

// ********************* Include additional libraries *******************
//#link "inxml"    //XML parser
//#link "sedbc"    //Database object
//#link "secomobj" //COM object
#include "$IMAGENOWDIR6$\\script\\lib\\iScriptDebug.jsh"
#include "$IMAGENOWDIR6$\\script\\lib\\envVariable.jsh"
#include "$IMAGENOWDIR6$\\script\\lib\\yaml_loader.jsh"
#include "$IMAGENOWDIR6$\\script\\lib\\RouteItem.jsh"

// *********************         Configuration        *******************

// logging
#define LOG_TO_FILE         true    // false - log to stdout if ran by intool, true - log to inserverXX/log/ directory
#define DEBUG_LEVEL         5       // 0 - 5.  0 least output, 5 most verbose
#define SPLIT_LOG_BY_THREAD false   // set to true in high volume scripts when multiple worker threads are used (workflow, external message agent, etc)
#define MAX_LOG_FILE_SIZE   100     // Maximum size of log file (in MB) before a new one will be created
#define USE_CSV             false    //true if using a csv to complete projects, false if just doing everythign in a queue

// *********************       End  Configuration     *******************

// ********************* Initialize global variables ********************
var debug;

/**
* Main body of script.
* @method main
* @return {Boolean} True on success, false on error.
*/
function main ()
{
        try
        { 
            debug = new iScriptDebug("CompleteTaskAndMove", LOG_TO_FILE, DEBUG_LEVEL);
            debug.log("WARNING", "CompleteTaskAndMove.js starting.\n");

            if (USE_CSV)
            {
              debug.log("INFO", "Processing based on csv\n");
              processByCSV();
            }
            else
            {
              debug.log("INFO","Processing contents of WF queues\n")
              processByQueue();
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

function processByQueue()
{
  debug.log("INFO","Attempting to load YAML\n");
  loadYAMLConfig(imagenowDir6+"\\script\\CompleteTaskAndMove\\queue_config\\");

  for (var sourceQConfig in CFG.queue_config)
  { 
    debug.log("DEBUG","Begin processing CTaM configuration\n");
    var CTaM_CONFIG = CFG.queue_config[sourceQConfig].CTaM_CONFIG;

    debug.log("DEBUG","Elements in YAML File [%s]\n",CTaM_CONFIG.length);
    for(var i=0; i<CTaM_CONFIG.length; i++)
    {
      debug.log("DEBUG","CTaM Config: [%s], [%s], [%s], [%s], [%s]\n",CTaM_CONFIG[i].SOURCE_QUEUE, CTaM_CONFIG[i].ERROR_QUEUE, CTaM_CONFIG[i].TASK_TEMPLATE, CTaM_CONFIG[i].TASK_REASON, CTaM_CONFIG[i].DESTINATION_QUEUE);

      var wfQueue = new INWfQueue();
      wfQueue.name = CTaM_CONFIG[i].SOURCE_QUEUE;

      var queueStartTime = new Date(1970, 1, 1);

      var wfItems = wfQueue.getItemList(WfItemState.Any, WfItemQueryDirection.AfterTimestamp, 2, queueStartTime);

      if (wfItems == null || !wfItems)
      {
        debug.log("ERROR","Unable to get items from [%s]: [%s]\n", CTaM_CONFIG[i].SOURCE_QUEUE, getErrMsg());
        continue;
      }

      printf("Found [%s] items in [%s]\n", wfItems.length, CTaM_CONFIG[i].SOURCE_QUEUE);
      var curCount = wfItems.length;
      debug.log("INFO","Found [%s] items in [%s]\n", wfItems.length, CTaM_CONFIG[i].SOURCE_QUEUE);

      var taskTemplate = new INTaskTemplate();
      taskTemplate.name = CTaM_CONFIG[i].TASK_TEMPLATE;
      if (!taskTemplate.getInfo())
      {
        debug.log("ERROR","Could not get info for task template [%s]: [%s]", CTaM_CONFIG[i].TASK_TEMPLATE, getErrMsg());
        continue;
      }

      //printf("%s\n", taskTemplate.id)
      var targetTemplate = taskTemplate.id;
      var targetQueue = CTaM_CONFIG[i].DESTINATION_QUEUE;

      var taskReasonList = taskTemplate.actionReasonListID;
      var reasonList = INBizList.get(taskReasonList);
      
      if (!reasonList || reasonList == null)
      {
        debug.log("ERROR","Could not get task reason list for [%s]: [%s]\n", CTaM_CONFIG[i].TASK_TEMPLATE, getErrMsg());
        continue;
      }

      var taskReason = reasonList.getMembers();

      if (!taskReason || taskReason == null)
      {
        debug.log("ERROR","Could not get task reason list for [%s]: [%s]\n", reasonList.name, getErrMsg());
        continue;
      }

      //validate the task reason is in the list and get it's ID
      var taskInfo = {flag:false,id:null,text:null};
      taskInfo = validateElement(taskInfo,taskReason,CTaM_CONFIG[i].TASK_REASON);

      if (!taskInfo.flag)
      {
        debug.log("ERROR","Could not find the task reason [%s] in [%s]\n", CTaM_CONFIG[i].TASK_REASON, taskTemplate.name);
        continue;
      }

      for (var k=0; k<wfItems.length; k++)
      {
        var folder = new INFolder(wfItems[k].objectId);
        if (!folder.getInfo())
        {
          debug.log("ERROR","Could not retrieve info for folder ID [%s]: [%s]\n", wfItems[k].objectId, getErrMsg());
          continue;
        }

        var wfFolder = folder.getWfInfo();
        if (!wfFolder || wfFolder == null)
        {
          debug.log("ERROR","Could not retrieve workflow info for folder ID [%s]: [%s]\n", folder.id, getErrMsg());
          continue;
        }              
        debug.log("DEBUG","Working folder: [%s]\n",folder.id);
        var taskList = new Array();
        if(!INTask.getTasks(folder.id,"","",taskList))
        {
          debug.log("ERROR","Failed to get tasks for [%s]: [%s]\n", folder.id, getErrMsg());
          continue;
        }
        else
        {
          var readyToRoute = false;

          var taskToUse = {id:null,taskTemplateID:null,creationTime:0};
          taskToUse = getLatestTask(taskToUse, taskList, targetTemplate);

          if (taskToUse.id == null)
          {
            debug.log("WARNING","No available tasks to complete for [%s]\n",folder.id);
            continue;
          }

          var thisTask = new INTask(taskToUse.taskTemplateID, taskToUse.id);

          if (!thisTask.getInfo())
          {
            debug.log("ERROR", "Could not get info for task [%s]: [%s]\n", taskToUse.id, getErrMsg());
            continue;
          }

          var taskToWork;
          if(!(taskToWork = INTask.workByID(thisTask.id)))
          {
              debug.log("ERROR","Unable to get task [%s] for processing: [%s]\n", thisTask.id, getErrMsg());
              continue;
          }
          else
          {
            debug.log("INFO","Working task [%s] for [%s]\n", thisTask.id, folder.id);

            if(!thisTask.complete(taskInfo.id))
            {
              debug.log("ERROR","Could not complete task for [%s]: [%s]\n", folder.id, getErrMsg());
              continue;
            }
            else
            {
              if(!thisTask.addComment("Task completed by CompleteTaskAndMove.js: " + taskInfo.text))
              {
                debug.log("ERROR", "Could not add comment to task with completion\n");
              }

              printf("[%s]ID: [%s] - Completed [%s] with a reason of [%s]\n", curCount, folder.id, thisTask.id, taskInfo.text);
              debug.log("INFO","[%s]ID: [%s] - Completed [%s] with a reason of [%s]\n", curCount, folder.id, thisTask.id, taskInfo.text);
              curCount--;
              readyToRoute = true;
            }
          }// end of working current assigned task

          if (readyToRoute)
          {          
            if(!RouteItem(wfFolder[0].id, targetQueue, "CTaM: " + taskInfo.text + " to " + targetQueue))
            {
              debug.log("ERROR","Failed to route [%s]: [%s]\n", folder.id, getErrMsg());
              continue;
            }
            else
            {
              debug.log("INFO","Successfully routed [%s] to [%s].\n", folder.id, targetQueue);
            }
          }

        }// end of task processing for current folder

      }// end of processing for items in source queue

    } // end of for individual configs in yaml

  } // end of for config in yaml

} // END OF PROCESS BY QUEUE 

function processByCSV()
{
  // place to check for csv
  //csv format EMPLID,APP NO,Task Template,Task Reason,Destination queue
  csvPath = imagenowDir6+"\\script\\CompleteTaskAndMove\\";
  workingCsv = "";

  var csvResult = SElib.directory(csvPath+"*.csv", false, ~FATTR_SUBDIR);

  if (csvResult == null)
  {
    debug.log("WARNING","No csvs detected in [%s]. Exiting processByCSV\n", csvPath);
    return false;
  }
  else
  {
    if(csvResult.length > 1) //get out if there are multiple csvs in the target drawer
    {
      debug.log("WARNING","[%s] csvs detected in [%s].  There can be only 1.\n", csvResult.length, csvPath);
      return false;
    }
    debug.log("INFO","Working with the csv [%s]\n", csvResult[0].name);
    workingCsv = csvResult[0].name;
  }  // end of csvResult != null

  var csvLineCount = countLines(workingCsv);

  if (!csvLineCount || csvLineCount == null)
  {
    debug.log("ERROR","Couldn't get the number of rows in [%s].\n",workingCsv);
    return false;
  }

  //open & read & close csv
  var fp = Clib.fopen(workingCsv, "r");
  if ( fp == null )
  {
    debug.log("ERROR","Could not open [%s] for reading\n", workingCsv);
    return false;
  } // end if fp = null

   //for each row in the csv
   while ( null != (line=Clib.fgets(fp)) )
   {     
      var curLineNum = csvLineCount;
      columns = line.replace(/\r?\n|\r/,'').split(",");
      if(columns.length != 5)
      {
        debug.log("ERROR","Line has incorrect number [%s] of columns: [%s]\n", columns.length, line);
        csvLineCount--;
        continue;
      }
      else
      {
        csvLineCount--;
      }

      var emplid = columns[0];
      var appNo = columns[1];
      var tskTemp = columns[2];
      var tskRsn = columns[3];
      var desQueue = columns[4];
      var csvCampus = tskTemp.substr(tskTemp.length - 3);

      debug.log("INFO","[%s]emplid[%s],appNo[%s],tskTemp[%s],tskRsn[%s],desQueue[%s],campus[%s]\n", curLineNum, emplid, appNo , tskTemp , tskRsn , desQueue, csvCampus);

      

      // make sure folder exists
      var folder = new INFolder(emplid + " APP: " + appNo, "Application Undergraduate Admissions " + csvCampus);
      if (!folder.getInfo())
      {
        debug.log("ERROR","Folder: [%s APP: %s] [Applications Undergraduate Admissions %s] doesn't exist\n", emplid, appNo, csvCampus);
        continue;
      }

      // make sure task template exists
      var csvTemplate = new INTaskTemplate();
      csvTemplate.name = tskTemp;
      if (!csvTemplate.getInfo())
      {
        debug.log("ERROR","Task template [%s] doesn't exist\n", tskTemp);
        continue;
      }

      //make sure we can see the reason list
      var csvTargetTemplate = csvTemplate.id;
      var csvReasonList = csvTemplate.actionReasonListID;
      var csvRsnLst = INBizList.get(csvReasonList);
                
      if (!csvRsnLst || csvRsnLst == null)
      {
        debug.log("ERROR","Could not get task reason list for [%s]: [%s]\n", tskTemp, getErrMsg());
        continue;
      }

       var csvTskRsn = csvRsnLst.getMembers();

      if (!csvTskRsn || csvTskRsn == null)
      {
        debug.log("ERROR","Could not get task reason list for [%s]: [%s]\n", csvRsnLst.name, getErrMsg());
        continue;
      }

      //validate the task reason is in the list and get its ID
      var taskInfo = {flag:false,id:null,text:null};
      taskInfo = validateElement(taskInfo,csvTskRsn,tskRsn);

      if (!taskInfo.flag)
      {
        debug.log("ERROR","Could not find the task reason [%s] in [%s]\n", csvTskRsn, csvTemplate.name);
        continue;
      }

      //check to see if the workflow queue exists.

      var queueExists = INWfAdmin.queueSearch(desQueue);

      if (queueExists.length > 1)
      {
        debug.log("WARNING","Multiple results found for [%s]. Going to next row.\n", desQueue);
        continue;
      }
      if (!queueExists[0] || queueExists[0] == null)
      {
        debug.log("WARNING","Could not find workflow queue with the name: [%s]\n", desQueue);
        continue;
      } 
      // end of check for wf queue

      var wfFolder = folder.getWfInfo();
      if (!wfFolder || wfFolder == null)
      {
        debug.log("ERROR","Could not retrieve workflow info for folder ID [%s]: [%s]\n", folder.id, getErrMsg());
        continue;
      }

      var csvTaskList = new Array();
      if(!INTask.getTasks(folder.id,"","",csvTaskList))
      {
       debug.log("ERROR","Failed to get tasks for [%s]: [%s]\n", folder.id, getErrMsg());
       continue;
      }
      else
      {
        var readyToRoute = false;
        var useThisTask = {id:null,taskTemplateID:null,creationTime:0};

        useThisTask = getLatestTask(uesThisTask, csvTaskList, csvTargetTemplate);

        if(useThisTask.id == null)
        {
          debug.log("WARNING","No available tasks to complete for [%s]\n",folder.id);
          continue;
        }

        var thisTask = new INTask(useThisTask.taskTemplateID, useThisTask.id);

        if (!thisTask.getInfo())
        {
          debug.log("ERROR", "Could not get info for task [%s]: [%s]\n", csvTaskList[l].id, getErrMsg());
          continue;
        }

        var taskToWork;
        if(!(taskToWork = INTask.workByID(thisTask.id)))
        {
          debug.log("ERROR","Unable to get task [%s] for processing: [%s]\n", thisTask.id, getErrMsg());
          continue;
        }
        else
        {
          debug.log("INFO","Working task [%s] for [%s]\n", thisTask.id, folder.id);

          if(!thisTask.complete(taskInfo.id))
          {
            debug.log("ERROR","Could not complete task for [%s]: [%s]\n", folder.id, getErrMsg());
            continue;
          }
          else
          {
            if(!thisTask.addComment("Task completed by CompleteTaskAndMove.js: " + taskInfo.text))
            {
              debug.log("ERROR", "Could not add comment to task with completion: [%s]\n", getErrMsg());
            }

            printf("[%s] ID: [%s] - Completed [%s] with a reason of [%s]\n", curLineNum, folder.id, thisTask.id, taskInfo.text);
            debug.log("INFO","[%s] ID: [%s] - Completed [%s] with a reason of [%s]\n", curLineNum, folder.id, thisTask.id, taskInfo.text);
            readyToRoute = true;
          }
        }// end of working current assigned task

        if (readyToRoute)
        {          
          if(!wfFolder[0].id)
          {
            debug.log("ERROR","Failed to add [%s] to destination queue: [%s]\n", folder.id, getErrMsg());
            continue;
          }
          if(!RouteItem(wfFolder[0].id, desQueue, "CTaM: " + taskInfo.text + " to " + desQueue))
          {
            debug.log("ERROR","Failed to route [%s]: [%s]\n", folder.id, getErrMsg());
            continue;
          }
          else
          {
            debug.log("INFO","Successfully routed [%s] to [%s].\n", folder.id, desQueue);
          }
        } // end of if ready to route
        else
        {
          debug.log("WARNING", "[%s]  - [%s APP: %s] was not routed.\n", folder.id, emplid, appNo);
        }
      }// end of task processing for current folder
   } // ens of while
  Clib.fclose(fp);
} // END OF PROCESS BY CSV

//this funciton gets the number of lines in the csv
function countLines(csvName)
{
  var lineNum = 0;
  var lc = Clib.fopen(csvName, "r");
  if ( lc == null )
  {
    debug.log("ERROR","Error opening [%s] for reading.\n",csvName);
    Clib.fclose(lc);
    return false;
  }
  else
  {
    while ( null != (line=Clib.fgets(lc)) )
    {
      lineNum++;
    }
  }
  Clib.fclose(lc);
  return lineNum;
}//end countLines

//this function will return the latest task
function getLatestTask(taskHash, taskArray, tmpltChk)
{
  for (var m = 0; m<taskArray.length; m++)
  {
    //if we've got the right template and the task is assigned
    if((taskArray[m].taskTemplateID == tmpltChk) && (taskArray[m].state == "Assigned"))
    {
      //if this task's creation time is greater than the previous greatest time
      if(taskArray[m].creationTime > taskHash.creationTime)
      {
        taskHash.id = taskArray[m].id;
        taskHash.taskTemplateID = taskArray[m].taskTemplateID;
        taskHash.creationTime = taskArray[m].creationTime;
      }//end if this task's creation time is greater than the previous greatest time
    }//end if we've got the right template and the task is assigned
  }//end for each task in array

  return taskHash;
}//end function to return the latest task

//this function will validate that an element exists in an array
function validateElement(itemHash,itemArr,item)
{
  for (var i=0; i<itemArr.length; i++)
  {
    if(itemArr[i].text == item)
    {
      itemHash.id = itemArr[i].id;
      itemHash.text = itemArr[i].text;
      itemHash.flag = true;
      break;
    }
  }// end of getting the task reason ID
  return itemHash;
}
//end function to check to see if item is in array
//