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
#define USE_CSV             true    //true if using a csv to complete projects, false if just doing everythign in a queue

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

                var wfItems = wfQueue.getItemList(WfItemState.Any, WfItemQueryDirection.AfterTimestamp, 5000, queueStartTime);

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


                var reasonExists = false;
                var taskCompletionReason;
                var textTaskCompletionReason;

                for (var j=0; j<taskReason.length; j++)
                {
                  if(taskReason[j].text == CTaM_CONFIG[i].TASK_REASON)
                  {
                    taskCompletionReason = taskReason[j].id;
                    textTaskCompletionReason = taskReason[j].text;
                    reasonExists = true;
                    break;
                  }
                }// end of getting the task reason ID

                if (!reasonExists)
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
              
                  var taskList = new Array();
                  if(!INTask.getTasks(folder.id,"","",taskList))
                  {
                    debug.log("ERROR","Failed to get tasks for [%s]: [%s]\n", folder.id, getErrMsg());
                    continue;
                  }
                  else
                  {
                    var readyToRoute = false;

                    for(var l=0; l<taskList.length; l++)
                    {
                      if (targetTemplate == taskList[l].taskTemplateID)
                      {
                        var thisTask = new INTask(taskList[l].taskTemplateID, taskList[l].id);

                        if (!thisTask.getInfo())
                        {
                          debug.log("ERROR", "Could not get info for task [%s]: [%s]\n", taskList[l].id, getErrMsg());
                          continue;
                        }

                        if (thisTask.state == "Assigned")
                        {
                          var taskToWork;
                          if(!(taskToWork = INTask.workByID(thisTask.id)))
                          {
                              debug.log("ERROR","Unable to get task [%s] for processing: [%s]\n", thisTask.id, getErrMsg());
                              continue;
                          }
                          else
                          {
                            debug.log("INFO","Working task [%s] for [%s]\n", thisTask.id, folder.id);

                            if(!thisTask.complete(taskCompletionReason))
                            {
                              debug.log("ERROR","Could not complete task for [%s]: [%s]\n", folder.id, getErrMsg());
                              continue;
                            }
                            else
                            {
                              if(!thisTask.addComment("Task completed by CompleteTaskAndMove.js: " + textTaskCompletionReason))
                              {
                                debug.log("ERROR", "Could not add comment to task with completion\n");
                              }

                              printf("[%s]ID: [%s] - Completed [%s] with a reason of [%s]\n", curCount, folder.id, thisTask.id, textTaskCompletionReason);
                              debug.log("INFO","[%s]ID: [%s] - Completed [%s] with a reason of [%s]\n", curCount, folder.id, thisTask.id, textTaskCompletionReason);
                              curCount--;
                              readyToRoute = true;
                              //break to only complete one task if there are multiple!
                              break;
                            }

                          }// end of working current assigned task

                        }// end of processing for assigned tasks

                      }// end of processing for target task template

                    }// end of for loop checking for target task template

                    if (readyToRoute)
                    {          
                      if(!RouteItem(wfFolder[0].id, targetQueue, "CTaM: " + textTaskCompletionReason + " to " + targetQueue))
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

  //open & read & close csv
  var fp = Clib.fopen(workingCsv, "r");
  if ( fp == null )
  {
    debug.log("ERROR","Could not open [%s] for reading\n", workingCsv);
    return false;
  } // end if fp = null

   while ( null != (line=Clib.fgets(fp)) )
   {     
      columns = line.replace(/\r?\n|\r/,'').split(",");
      if(columns.length != 5)
      {
        debug.log("ERROR","Line has incorrect number [%s] of columns: [%s]\n", columns.length, line);
        continue;
      }

      var emplid = columns[0];
      var appNo = columns[1];
      var tskTemp = columns[2];
      var tskRsn = columns[3];
      var desQueue = columns[4];
      var csvCampus = tskTemp.substr(tskTemp.length - 3);

      debug.log("INFO","emplid[%s],appNo[%s],tskTemp[%s],tskRsn[%s],desQueue[%s],campus[%s]\n", emplid, appNo , tskTemp , tskRsn , desQueue, csvCampus);

      

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


      var csvReasonExists = false;
      var csvtaskCmpltnReason;
      var csvTextTaskCmpltnReason;

      for (var j=0; j<csvTskRsn.length; j++)
      {
        if(csvTskRsn[j].text == tskRsn)
        {
          csvtaskCmpltnReason = csvTskRsn[j].id;
          csvTextTaskCmpltnReason = csvTskRsn[j].text;
          csvReasonExists = true;
          break;
        }
      }// end of getting the task reason ID

      if (!csvReasonExists)
      {
        debug.log("ERROR","Could not find the task reason [%s] in [%s]\n", tsk, csvTemplate.name);
        continue;
      }

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

        for(var l=0; l<csvTaskList.length; l++)
          {
            if (csvTargetTemplate == csvTaskList[l].taskTemplateID)
            {
              var thisTask = new INTask(csvTaskList[l].taskTemplateID, csvTaskList[l].id);

              if (!thisTask.getInfo())
              {
                debug.log("ERROR", "Could not get info for task [%s]: [%s]\n", csvTaskList[l].id, getErrMsg());
                continue;
              }

              if (thisTask.state == "Assigned")
              {
                var taskToWork;
                if(!(taskToWork = INTask.workByID(thisTask.id)))
                {
                  debug.log("ERROR","Unable to get task [%s] for processing: [%s]\n", thisTask.id, getErrMsg());
                  continue;
                }
                else
                {
                  debug.log("INFO","Working task [%s] for [%s]\n", thisTask.id, folder.id);

                  if(!thisTask.complete(csvtaskCmpltnReason))
                  {
                    debug.log("ERROR","Could not complete task for [%s]: [%s]\n", folder.id, getErrMsg());
                    continue;
                  }
                  else
                  {
                    if(!thisTask.addComment("Task completed by CompleteTaskAndMove.js: " + csvTextTaskCmpltnReason))
                    {
                      debug.log("ERROR", "Could not add comment to task with completion: [%s]\n", getErrMsg());
                    }

                    printf("ID: [%s] - Completed [%s] with a reason of [%s]\n", folder.id, thisTask.id, csvTextTaskCmpltnReason);
                    debug.log("INFO","ID: [%s] - Completed [%s] with a reason of [%s]\n", folder.id, thisTask.id, csvTextTaskCmpltnReason);
                    readyToRoute = true;
                    //break to only complete one task if there are multiple!
                    break;
                  }

                }// end of working current assigned task

              }// end of processing for assigned tasks

            }// end of processing for target task template

          }// end of for loop checking for target task template

          if (readyToRoute)
          {          
            if(!RouteItem(wfFolder[0].id, desQueue, "CTaM: " + csvTextTaskCmpltnReason + " to " + desQueue))
            {
              debug.log("ERROR","Failed to route [%s]: [%s]\n", folder.id, getErrMsg());
              continue;
            }
            else
            {
              debug.log("INFO","Successfully routed [%s] to [%s].\n", folder.id, desQueue);
            }
          } // end of if ready to route

        }// end of task processing for current folder

   } // ens of while
  Clib.fclose(fp);

} // END OF PROCESS BY CSV


//