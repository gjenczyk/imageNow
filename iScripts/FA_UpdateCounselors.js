/********************************************************************************
	Name:			FA_UpdateCounselors.js
	Author:			UITS & Gregg Jenczyk
	Created:		06/11/2015
	Last Updated:	
	For Version:	6.7
	Script Version:
---------------------------------------------------------------------------------
    Summary:
		Routes a Folder from the decision ready queues to the appropriate counselor superq/subq combination based on the workflow history and admit type criteria	
	
	Mod Summary:
		
    Business Use:  
		This script will be run via intool  // not anymore
	Intool command: 
		intool --cmd run-iscript --file FA_UpdateCounselors.js   // This is NOT an intool script though
		
********************************************************************************/

#include "$IMAGENOWDIR6$\\script\\lib\\iScriptDebug.jsh"
#include "$IMAGENOWDIR6$\\script\\lib\\commonSharedFunction.jsh"
#include "$IMAGENOWDIR6$\\script\\lib\\GetDocsByVsl.jsh"
#include "$IMAGENOWDIR6$\\script\\lib\\RouteItem.jsh"
#include "$IMAGENOWDIR6$\\script\\lib\\GetProp.jsh"
#include "$IMAGENOWDIR6$\\script\\lib\\PropertyManager.jsh"
#include "$IMAGENOWDIR6$\\script\\lib\\yaml_loader.jsh"

// *********************         Configuration        *******************

#define CONFIG_VERIFIED		true	// set to true when configuration values have been verified
#define LOG_TO_FILE 		true	// false - log to stdout if ran by intool, true - log to inserverXX/log/ directory
#define DEBUG_LEVEL 		5		// 0 - 5.  0 least output, 5 most verbose   
     

// *********************       End  Configuration     *******************

// ********************* Initialize global variables ********************
//* Use this to Overide the Script such that you can run from intool *//

var debug = "";

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
		debug = new iScriptDebug("FA_UpdateCounselors", LOG_TO_FILE, DEBUG_LEVEL);
	    debug.log("WARNING", "FA_UpdateCounselors script started.\n");



  	    var wfItem = new INWfItem(currentWfItem.id);//"321YZ78_07B9LS6W000001T");//
		if(!wfItem.id || !wfItem.getInfo())
		{
			debug.log("CRITICAL", "Couldn't get info for wfItem: %s\n", getErrMsg());
			return false;
		}

/*		if (!wfItem.setState(WfItemState.Working, "FA_UpdateCounselors"))
		{
    		debug.log("ERROR","Could not set state: %s.\n", getErrMsg());
    		return false;
		}
*/
		var wfQueue = wfItem.queueName;
		var wfPrefix = wfQueue.substring(0,3);

		debug.log("INFO","Attempting to load YAML\n");
 
  	    loadYAMLConfig(imagenowDir6+"\\script\\config_scripts\\FA_UpdateCounselors\\");

		for (var sourceQConfig in CFG.FA_UpdateCounselors)
		{
			//debug.log("DEBUG","it's this = [%s] - -[%s]\n", CFG.FA_UpdateCounselors[sourceQConfig].CAMPUS_CONFIG, sourceQConfig);
			
			if (sourceQConfig.substring(0,3) != wfPrefix)
			{
				continue;
			}

			var CAMPUS_CONFIG = CFG.FA_UpdateCounselors[sourceQConfig].CAMPUS_CONFIG
			for (var i = 0; i < CAMPUS_CONFIG.length; i++)
			{
				debug.log("DEBUG","CAMPUS_CONFIG[i].SOURCE_QUEUE is " + CAMPUS_CONFIG[i].SOURCE_QUEUE + "\n")
				if (CAMPUS_CONFIG[i].SOURCE_QUEUE == wfQueue)
				{
					debug.log("DEBUG","Found configuration for [%s]\n", wfQueue);
					var wfFldr = new INFolder(wfItem.objectId);
					var reason = "";
					var userId = "";

					if (!wfFldr.getInfo())
					{
						debug.log("ERROR", "Getting Info for the Folder failed: [%s].\n", getErrMsg());
						return false;
					}

					debug.log("INFO", "Loading Folder : [%s] Folder type : [%s].\n", wfFldr.name, wfFldr.projTypeName);

					if (wfFldr.projTypeName != CAMPUS_CONFIG[i].FOLDER_TYPE)
					{
						debug.log("WARNING", "Unexpected Folder Type : [%s] Folder name: [%s].\n", wfFldr.projTypeName, wfFldr.name);
						return false;
					}

					debug.log("INFO", "Looping through Workflow History: [%s] Folder type : [%s].\n", wfFldr.name, wfFldr.projTypeName);
					var history = wfItem.getHistory();
					if(history == null)
					{
						debug.log("WARNING", " No Workflow History Exists: [%s]\n", getErrMsg());
						return false;
					}

					debug.log("DEBUG", "Number of history logs:  [%s]\n", history.length);
					for (var j = history.length - 1; j >= 0; j--)
					{
						reason = trim(history[j].reasonText);
						var RoutedInFound = false;

						if (reason == "Routed In")
						{
							userId = trim(history[j].stateUserName);
							RoutedInFound = true; 
							debug.log("INFO", "Reason : [%s], User ID is :[%s]\n", reason, userId);

							//assign the value to the prop
							if (userId == "")
							{
								debug.log("WARNING", "User Id is blank in latest 'Routed In' history log. [%s]\n", reason);
								var moveReason = "User ID in WF History is null.";
								RouteItem(wfItem, CAMPUS_CONFIG[i].ERROR_QUEUE, "FA_UpdateCounselor: No user found.", true);
								break;
							}
							else
							{
								prop = new INInstanceProp();
								prop.name = CAMPUS_CONFIG[i].COUNSELOR_PROP;
								prop.setValue(userId);
								
								wfFldr.getInfo();
								if (wfFldr.setCustomProperties(prop))
								{
									var pm = new PropertyManager();
									counselorId = pm.get(wfFldr, CAMPUS_CONFIG[i].COUNSELOR_PROP);
									debug.log("INFO", " Success! Setting Property. Assigned Counselor is now: [%s]. Routing Folder : [%s]\n", counselorId, wfFldr.name );
									if(!sendToCounselor(wfItem, CAMPUS_CONFIG[i].DESTINATION_QUEUE, counselorId, CAMPUS_CONFIG[i].SECURITY_GROUPS))
									{
										debug.log("ERROR","Could not send to counselor queue - routing to error.\n");
										RouteItem(wfItem, CAMPUS_CONFIG[i].ERROR_QUEUE, "FA_UpdateCounselor: Could not find counselor queue.", true);
										return false;
									}
									break;
								}
								else
								{
									debug.log("INFO", "Setting Assigned Counselor Property failed. [%s]\n", getErrMsg() );
									var moveReason = "Setting Assigned Counselor Property failed";
									RouteItem(wfItem, CAMPUS_CONFIG[i].ERROR_QUEUE, "FA_UpdateCounselor: Could not assign counselor.", true);
									break;
								}
							} //end of assigning value

						} // end if (reason == "Routed In")

					} // end for (var j = history.length - 1; j >= 0; j--)

					if (!RoutedInFound)
					{
						debug.log("INFO", "Routed In log not found in WF history. Folder Name:  [%s]\n", wfFldr.name );
						RouteItem(wfItem, CAMPUS_CONFIG[i].ERROR_QUEUE, "FA_UpdateCounselor: Folder wasn't routed in?", true);
					}

				} // end if (CAMPUS_CONFIG[i].SOURCE_QUEUE == wfQueue)

			} //end for (var i = 0; i < CAMPUS_CONFIG.length; i++)

		} // end for (var sourceQConfig in CFG.FA_UpdateCounselors) 

	} // end try
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
			debug.log("INFO", "FA_UpdateCounselor script finished.\n");
			debug.finish();
		}
	}
}

//this function will send a folder to the correct queue
function sendToCounselor(wfItem, qName, counsName, qSecurity)
{
	debug.log("DEBUG","inside sendToCounselor\n");
	debug.log("DEBUG","qName is [%s] and counsName is [%s]\n", qName, counsName);

	var queue = new INWfQueue("",qName);
	var sQtest = queue.getSubQueueMembers();

	if(!sQtest)
	{
		debug.log("ERROR", "Couldn't get queue type for [%s]\n", qName);
		return false;
	}
	
	debug.log("DEBUG","Number of subQ in [%s]: %d\n", qName, sQtest.length);
	if(sQtest.length == 0)
	{
		debug.log("INFO","Preparing to route [%s] to [%s]\n", wfItem.objectId, qName);
	}
	else
	{
		debug.log("INFO","Checking to see if target subqueue exists for [%s]\n", counsName);
		var delim = "";
		var counsParts = counsName.split(delim);
		var firstName = counsParts[1];
		var lastName = counsParts[2];
		var counsQ = firstName + " " + lastName + " (" + qName + ")";
		var foundQ = false;

		for (var i=0; i<sQtest.length; i++)
	    {
	        debug.log("DEBUG","%d: sub queue name: [%s] looking for [%s]\n", i, sQtest[i], counsQ);
	        if(sQtest[i] == counsQ)
	        {
	        	debug.log("INFO","Found [%s] - routing project!\n",counsQ);
	        	foundQ = true;
	        	break;
	        }
	    }

	    if(!foundQ)
	    {
	    	debug.log("WARNING","Couldn't find a matching queue!\n");
	    	if(qSecurity == null || qSecurity == undefined)
	    	{
	    		debug.log("INFO","No security groups configured for [%s].  No subqueue will be created for [%s].\n", qName, firstName + " " + lastName);
	    		return false;
	    	}
	    	if(!createSubQueue(qName, counsParts, qSecurity))
	    	{
	    		debug.log("WARNING","User does not have permission to act as a counselor.\n");
	    		return false;	
	    	}
	    	
	    }

	    if(!RouteItem(wfItem, counsQ, "FA_UpdateCounselors: " + counsName, true))
	    {
	    	debug.log("INFO","Unable to route folder!\n")
	    	return false;
	    }
	   	return true;
	}
} // end sendToCounselor

//function to create a counselor subq if the user is in the correct security group
function createSubQueue(cQueue, cInfo, accesGrps)
{
	debug.log("DEBUG","Attempting to suqueue in [%s] for [%s]\n", cQueue, cInfo);
	var makeNewQueue = false;

	for(var s = 0; s < accesGrps.length; s++)
	{
		var group = new INGroup(accesGrps[s])
		if(!group.getInfo())
		{
			debug.log("ERROR","[%s] is not a valid security group.\n", accesGrps[s]);
			//don't want to continue with bad config
			return false;
		}

		var users = INGroup.getMembers(accesGrps[s])
		for (var t = 0; t < users.length; t++)
		{
			if(users[t] == cInfo[0])
			{
				debug.log("INFO","[%s] is a member of [%s]. Setting makeNewQueue to true.\n", users[t], accesGrps[s]);
				makeNewQueue = true;
				break;
			}
		}
		if(makeNewQueue)
		{
			break;
		}
	}// end for each in accesGrps

	if(makeNewQueue)
	{
		debug.log("DEBUG","Making subqueue for [%s] in [%s]\n", cInfo, cQueue);

		var superQ=new INWfQueue("", cQueue);
		var q = new INWfQueue();
		q.name = cInfo[1] + " " + cInfo[2];
		var newQ = q.name + " (" + cQueue +")";
		if (!superQ.addSubQueue(q.name))
		{
		    debug.log("ERROR","Failed to add subqueue [%s]: [%s]\n", q.name, getErrMsg());
		    return false;
		}
		 else
		{
		    debug.log("INFO","Created [%s] in [%s]\n", q.name, cQueue);
		    var blankQ = new INWfQueue("",newQ);
		    for (var p = 0; p < accesGrps.length; p++)
		    {
		    	debug.log("INFO","Adding group [%s] to [%s]\n", accesGrps[p], q.name)
		    	if(!blankQ.addUser(accesGrps[p]))
		    	{
		    		debug.log("ERROR","Could not add [%s] to [%s]: [%s]\n", accesGrps[p], blankQ, getErrMsg())
		    		continue;
		    	}		    	
		    }
		}
		return true;
	}
	else
	{
		debug.log("INFO","User [%s] is not in the correct security groups. No subqueue will be made.\n", cInfo);
		return false;
	}
	
} // end createSubQueue
//