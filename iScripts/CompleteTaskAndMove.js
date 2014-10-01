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


//