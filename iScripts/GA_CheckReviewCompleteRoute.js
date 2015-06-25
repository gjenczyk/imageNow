/********************************************************************************
        Name:          GA_CheckReviewCompleteRoute.js
        Author:        Gregg Jenczyk
        Created:        06/12/2015
        Last Updated:   
        For Version:    
---------------------------------------------------------------------------------
        Summary:
               This script will make sure that a graduate decision has been stamped
               and routed correctly, and that the person who did the routing had the 
               authority to do so.   
               
        Mod Summary:
               Date-Initials: Modification description.



********************************************************************************/

// ********************* Include additional libraries *******************
//#link "inxml"    //XML parser
//#link "sedbc"    //Database object
//#link "secomobj" //COM object
#include "$IMAGENOWDIR6$\\script\\lib\\iScriptDebug.jsh"
#include "$IMAGENOWDIR6$\\script\\lib\\HostDBLookupInfo.jsh"
#include "$IMAGENOWDIR6$\\script\\lib\\GetProp.jsh"
#include "$IMAGENOWDIR6$\\script\\lib\\SetProp.jsh"
#include "$IMAGENOWDIR6$\\script\\lib\\RouteItem.jsh"
#include "$IMAGENOWDIR6$\\script\\STL\\packages\\Document\\getDocLogobArray.js"

// *********************         Configuration        *******************

// logging
#define LOG_TO_FILE         true    // false - log to stdout if ran by intool, true - log to inserverXX/log/ directory
#define DEBUG_LEVEL         5       // 0 - 5.  0 least output, 5 most verbose
#define SPLIT_LOG_BY_THREAD false   // set to true in high volume scripts when multiple worker threads are used (workflow, external message agent, etc)
#define MAX_LOG_FILE_SIZE   100     // Maximum size of log file (in MB) before a new one will be created


// *********************       End  Configuration     *******************

// ********************* Initialize global variables ********************

var VALIDATION_QUEUE = "Review Complete";
var SOURCE_QUEUE = "Ready for Review";
var MISSING_PROFILE_Q = " App Missing Profile";
var ERROR_Q = " Review Error";
var REASON_CP = "Signature Reason";
var VALIDATION_STAMP = "GA Decision Validation"
var NON_GPD_GROUPS = ["GA Directors", " Campus Administrators"];
var STAMP_TYPES = ["Admit", "Admit with Conditions", "Deny", "Waitlist"];
var PERMISSION_DENIED = "001-User does not have permission to route item";
var MISSING_PROFILE = "002-Matching Profile Sheet missing";
var MISSING_DECISION = "003-Digital Signature missing";
/**
* Main body of script.
* @method main
* @return {Boolean} True on success, false on error.
*/
  function main ()
{
    try
    {
      debug = new iScriptDebug("GA_CheckReviewCompleteRoute", LOG_TO_FILE, DEBUG_LEVEL);
      debug.log("WARNING", "GA_CheckReviewCompleteRoute script started.\n");

      var queuePatt = /(UM.GA)/;
      var queuePrefix = currentWfQueue.name.match(queuePatt);
      var errorQueue = queuePrefix[0] + ERROR_Q;

      //get wfItem info
      var wfItem = new INWfItem(currentWfItem.id);//"321YZ6K_076VZQREF000026");//"321YZ6G_071KRY1460001YN");//
      if(!wfItem.id || !wfItem.getInfo())
      {
        debug.log("ERROR", "Couldn't get info for wfItem: %s\n", getErrMsg());
        //route to error
        RouteItem(wfItem, errorQueue, getErrMsg());
        return false;
      }

      //get doc info
      var wfDoc = new INDocument(wfItem.objectId);
      if(!wfDoc.getInfo())
      {
        debug.log("ERROR", "Couldn't get info for doc: %s\n", getErrMsg());
        //route to error
        RouteItem(wfItem, errorQueue, getErrMsg());
        return false;
      }

      debug.log("INFO", "Checking review status for: [%s]\n", wfDoc);

      //get history associated with the document
      var history = wfItem.getHistory();
      if (!history || history == null || history.length < 2)
      {
        debug.log("ERROR","Insufficient history found for [%s]\n", wfDoc);
        //route to error
        RouteItem(wfItem, errorQueue, "Insufficient history found for " + wfDoc.id);
        return false;
      }

      //populate vars w/ info about the current and previous queue
      var validationQueue = routingHistory(history, VALIDATION_QUEUE, "Routed In");
      var sourceQueue = routingHistory(history, SOURCE_QUEUE, "Routed Out");

      //checks to make sure we're looking at the right history
      if((!validationQueue || validationQueue == null) || (!sourceQueue || sourceQueue == null))
      {
        debug.log("ERROR","Document is missing the required workflow history.\n")
        //route to error
        RouteItem(wfItem, errorQueue, "Document is missing the required workflow history.");
        return false;
      }

      //check to make sure the user is correct
      var userID = "";

      if (((!validationQueue.stateUserName || validationQueue.stateUserName== null) || (!sourceQueue.stateUserName || sourceQueue.stateUserName == null)) || (validationQueue.stateUserName != sourceQueue.stateUserName))
      {
        debug.log("ERROR","User mismatch for validationQueue [%s] and sourceQueue [%s]\n", validationQueue.stateUserName, sourceQueue.stateUserName);
        //route to error
        RouteItem(wfItem, errorQueue, "Routing user mismatch: " + validationQueue.stateUserName + " & " + sourceQueue.stateUserName);
        return false;
      }
      else
      {
        userID = validationQueue.stateUserName;
        debug.log("INFO","Found user: [%s]\n", userID);
      }

      //get list of groups assigned to the source queue
      var queueGroups = queueMembership(sourceQueue.queueName);

      if (!queueGroups || queueGroups == null)
      {
        debug.log("ERROR","Unable to determine if security groups are applied to [%s]\n", sourceQueue.queueName)
        //route to error
        RouteItem(wfItem, errorQueue, "Unable to find security for: " + sourceQueue.queueName);
        return false;
      }

      //make sure the router is able to do so
      var routerGroups = clearedForRouting(userID, queueGroups, NON_GPD_GROUPS, sourceQueue.queueName);
      if (!routerGroups || routerGroups == null)
      {
        debug.log("ERROR","[%s] doesn't have configured decision authority in [%s].\n", userID, sourceQueue.queueName);
        //route to error
        RouteItem(wfItem, sourceQueue.queueName, PERMISSION_DENIED);
        return false;
      }

      debug.log("INFO","The document is coming from: [%s] and is currently in [%s].\n", sourceQueue.queueName, validationQueue.queueName);
      
      //find the annotation
      var annotationTemplate = "";
      var annotationLocation = "";

      var profileSheet = findProfile(wfDoc);

      if(!profileSheet || profileSheet == null)
      {
        debug.log("ERROR","No profile sheet found for [%s]\n", wfDoc);
        var missingProfileQ = "UM" + wfDoc.drawer.substring(2,5) + MISSING_PROFILE_Q;
        RouteItem(wfItem, missingProfileQ, MISSING_PROFILE);
        return false;
      }

      var stamp = checkForStamps(wfDoc, profileSheet, validationQueue.queueName, annotationTemplate, annotationLocation);

      if(!stamp || stamp == null)
      {
        debug.log("ERROR","No decision stamps found.\n");
        RouteItem(wfItem, sourceQueue.queueName, MISSING_DECISION);
        return false;
      }

      debug.log("INFO","[%s] has permission to apply [%s] in [%s]\n", userID, annotationTemplate, sourceQueue.queueName);
      //update signature reason on applicaiton and apply validation stamp
      var signatureReason = annotationTemplate.substring(4);
      if(!SetProp(wfDoc, REASON_CP, signatureReason))
      {
        debug.log("ERROR","Could not set [%s] to [%s]\n", REASON_CP, signatureReason);
      }
      
      debug.log("INFO","Successfully set [%s] to [%s]\n", REASON_CP, signatureReason);
      
      applyValidationStamp(stamp, annotationTemplate, sourceQueue, routerGroups);

    } // end try
        
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
}

// ********************* Function Definitions **********************************

function routingHistory(wfHist, matchQ, reason)
{
  //get the current and previous queues from the WF History
  var qMatch = "";

  for (var h = wfHist.length-1; h >= 0; h--)
  {
    if (wfHist[h].reasonText == reason && wfHist[h].queueName.indexOf(matchQ) >= 0)
    {
      qMatch = wfHist[h];
      debug.log("INFO","Document came from queue: [%s]\n", qMatch.queueName);
      break;
    }
  }

  if(qMatch)
  {
    return qMatch;
  }
  else 
  {
    debug.log("ERROR","No matching queues found in wf history containing [%s] and [%s]\n", matchQ, reason);
    return false;
  }

} // end routingHistory

function queueMembership(qName)
{
  var queue = new INWfQueue();
  var qMembership = [];
  queue.name = qName;
  debug.log("INFO","Checking queue membership for [%s]\n", qName);
  var members = queue.getAccessUserList();
  if (members == null || members.length<1)
  {
      debug.log("ERROR","No members for queue [%s]\n", queue.name);
      return false;
  }
  else
  {
      for (var i=0; i<members.length; i++)
      {
          //exclude the individual users , workflow groups, and revs from decisions
          if (members[i].userType == 1 && !(members[i].userName.slice(-3) == "Rev" || members[i].userName.slice(0,1) == "U"))
          {
            qMembership.push(members[i].userName);
          }
      }

      return qMembership;
  }

} // end queueMembership

//check to see if this guy can route it
function clearedForRouting(uID, qGroup, oGroup, sourceQ)
{
  //get the name of the source subQ
  var subQPatt = / (\(U.*Ready for Review\))/;
  var campusPatt = /\(UM(.*)Ready for Review\)/;
  var subQName = sourceQ.replace(subQPatt, '');
  var campusName = sourceQ.match(campusPatt);
  var groupsToCheck = oGroup;
  var validGroups = [];

  //apply campus name to the standard groups
  if ((campusName[1].length != 4) && (campusName[1].indexOf("GA") < 0))
  {
    debug.log("ERROR","Unable to determine campus.  Found: [ %s]\n", campusName[1]);
    return false;
  }
  else
  {
    debug.log("DEBUG","Campus prefix is [ %s]\n",campusName[1]);
    var campusGrpPrefix = "UM" + campusName[1].substring(0,1);
    for(var g = 0; g < groupsToCheck.length; g++)
    {
      debug.log("DEBUG","Campus group is [%s]\n",campusGrpPrefix + groupsToCheck[g]);
      groupsToCheck[g] = campusGrpPrefix + groupsToCheck[g];
    }
  } // end of applying campus names to groups

  //and subq to groups to check
  for(var c = 0; c < qGroup.length; c ++)
  {
    if (qGroup[c].indexOf(subQName) > 0)
    {
      debug.log("INFO", "Found security group matching queue name convention: [%s] [%s]\n", qGroup[c], sourceQ);
      groupsToCheck.unshift(qGroup[c]);
      break;
    }
  } // end of adding group to check

  //make sure user is in one of those groups
  var permission = false;
  for (var s = 0; s < groupsToCheck.length; s++)
  {
    var members = INUser.getGroupMembers(groupsToCheck[s]);
    if (!members || members == null)
    {
      debug.log("ERROR","[%s] No member found or not a group: %s.\n",groupsToCheck[s], getErrMsg());
      continue;
    }
    else
    {
      for (var i=0;i<members.length;i++)
      {
        debug.log("DEBUG","User: [%s] Member: %s: %s.\n", uID, i, members[i]);
        if(uID == members[i])
        {
          debug.log("INFO","[%s] is a member of [%s]. Permission to route granted.\n", uID, groupsToCheck[s]);
          permission = true;
          validGroups.push(groupsToCheck[s]);
          continue;
        }
      }  
    }
  } // end of check for group membership

  if(!permission)
  {
    debug.log("WARNING","User [%s] does not have permission to make decisions in [%s].\n", uID, sourceQ);
    return false;
  }
  else
  {
    debug.log("INFO","User [%s] has decision privliges in [%s]\n", uID, sourceQ);
    return validGroups;
  }

} // end clearedForRouting

//function to find profile sheet
function findProfile(application)
{
  //find the profile sheet
  var appNo = GetProp(application, "SA Application Nbr");
  var profile_sql = "SELECT INUSER.IN_DOC.DOC_ID " +
    "FROM INUSER.IN_DOC " +
    "INNER JOIN INUSER.IN_DOC_TYPE " +
    "ON INUSER.IN_DOC_TYPE.DOC_TYPE_ID = INUSER.IN_DOC.DOC_TYPE_ID " +
    "INNER JOIN INUSER.IN_INSTANCE " +
    "ON INUSER.IN_INSTANCE.INSTANCE_ID = INUSER.IN_DOC.INSTANCE_ID " +
    "INNER JOIN INUSER.IN_INSTANCE_PROP " +
    "ON INUSER.IN_INSTANCE.INSTANCE_ID = INUSER.IN_INSTANCE_PROP.INSTANCE_ID " +
    "INNER JOIN INUSER.IN_DRAWER " +
    "ON INUSER.IN_DRAWER.DRAWER_ID = INUSER.IN_DOC.DRAWER_ID " +
    "WHERE INUSER.IN_DOC.FOLDER = '" + application.folder + "' " +
    "AND INUSER.IN_DRAWER.DRAWER_NAME = '" + application.drawer + "' " +
    "AND INUSER.IN_DOC_TYPE.DOC_TYPE_NAME LIKE 'Profile Sheet%' " +
    "AND INUSER.IN_INSTANCE_PROP.STRING_VAL = '" + appNo + "' " +
    "AND INUSER.IN_INSTANCE.DELETION_STATUS = '0' " +
    "ORDER BY INUSER.IN_INSTANCE.CREATION_TIME DESC;";

  var returnVal; 
  var cur = getHostDBLookupInfo_cur(profile_sql,returnVal);
  //if there are no folders with the RTM flag or if we couldn't get a list
  if(!cur || cur == null)
  {
    //missing profile, send to error
    debug.log("ERROR","Couldn't find a matching profile sheet\n");
    return false;
  }

  var profile = new INDocument(cur[0]);
  if(!profile.getInfo())
  {
    debug.log("ERROR","Could not get info for [%s] - [%s]\n", cur[0], getErrMsg());
    return false;
  }

  debug.log("INFO","Found profile sheet [%s]\n", profile);
  return profile;
} // find profile

//function to see if we can find stamps on a document
function checkForStamps(application, profile, decision, &template, &page)
{
  debug.log("DEBUG","Looking for stamps belonging to [%s]\n", application);
  //get the expected stamp template
  var decTempPatt = /(.*) \(UM([BDL]GA)/;
  var decTempArr = decision.match(decTempPatt);
  var decTemp = decTempArr[2] + " " + decTempArr[1];
  template = decTemp;
  debug.log("INFO", "Expecting a stamp template of [%s] for most recent stamp\n", decTemp);
  //get stamps from the application
  var appStamp = lookForStamps(application, decTemp);
  
  var profileStamp = lookForStamps(profile, decTemp);

  //return the latest stamp 
  debug.log("DEBUG","profileStamp: [%s] appStamp: [%s]\n", profileStamp, appStamp);
  var foundStamp;
  if(!appStamp && !profileStamp)
  {
    debug.log("WARNING", "No decsion stamps found.\n")
    return false;
  }
  else if (appStamp && !profileStamp)
  {
    debug.log("DEBUG","Found latest stamp on application: [%s] [%s]\n", appStamp.id, appStamp.text);
    foundStamp = appStamp;
    return foundStamp;
  }
  else if (!appStamp && profileStamp)
  {
    debug.log("DEBUG","Found latest stamp on profile sheet: [%s] [%s]\n", profileStamp.id, profileStamp.text);
    foundStamp = profileStamp;
    return foundStamp;
  }
  else
  {
    foundStamp =  profileStamp.creationTime > appStamp.creationTime ? profileStamp : appStamp;
    debug.log("DEBUG","Found lastest stamp from both application and profile sheet: [%s] [%s]\n", foundStamp.id, foundStamp.text);
    return foundStamp;
  }
  
} // end checkForStamps

//function to check for a stamp on a document
function lookForStamps(doc, decTemp)
{
  debug.log("DEBUG","Searching for stamp belonging to [%s]\n", doc);
  //set up current time and variables to store the newest stamp
  var time = new Date();
  var now = time.getTime();
  var newestStamp = null;
  var currentStamp = null;
  var stampDiff = null;
  var pages = getDocLogobArray(doc);

  //for each page in document
  for (var p = 0; p < pages.length; p++)
  {
    var logObj = new INLogicalObject(pages[p].id);

    if(!logObj || logObj == null)
    {
      debug.log("ERROR","Unable to create logical object for [%s]: [%s]\n", doc.id, getErrMsg());
      return false;
    }

    var subObjs = logObj.getSubObject(SubobType.Stamp);

    if (!subObjs || subObjs == null)
    {
      debug.log("DEBUG","Unable to find stamps for [%s]: [%s]\n", doc.id, getErrMsg());
      continue;
      //
    }

    //get the most recently applied stamp
    for (var i = 0; i < subObjs.length; i++)
    {
      currentStamp = subObjs[i];
      debug.log("DEBUG", "Subob ID: [%s] Type: [%s] Stamp Text: [%s]\n", currentStamp.id, currentStamp.type, currentStamp.text);
      var stampTimeDiff = now-currentStamp.creationTime;

      if(newestStamp == null || stampTimeDiff < stampDiff)
      {
        newestStamp = currentStamp;
        stampDiff = stampTimeDiff;
      }
    } // end for each subobject on page
  }// end for each page in document

  if(!newestStamp || newestStamp == null)
  {
    debug.log("WARNING","Unable to find newest stamp on [%s].\n", doc);
    return false;
  }
  else
  {
    debug.log("INFO","Found newest stamp: [%s] [%s] [%s]\n", newestStamp.id, newestStamp.text, newestStamp.creationTime);
    var newestStampTempl = new INSubobTemplate(newestStamp.templId);
    newestStampTempl.getInfo();
    if(newestStampTempl.name.indexOf(decTemp) < 0)
    {
      debug.log("WARNING","Newest stamp does not match queue.  Expecting [%s], we have [%s]\n", decTemp, newestStampTempl.name);
      return false;
    }
    else
    {
      debug.log("INFO","Newest stamp matches queue.  Expecting [%s], we have [%s]\n", decTemp, newestStampTempl.name);
      return newestStamp;
    }
  }
  
} // end of lookForStamps

//function to apply validation stamps
function applyValidationStamp(stamp, stamptmpl, queue, groups)
{
  //determine the stamp's location via the logob id
  var stamp_sql = "SELECT INUSER.IN_DOC.DOC_ID, " +
  "INUSER.IN_LOGOB.SEQ_NUM " +
  "FROM INUSER.IN_SUBOB " +
  "INNER JOIN INUSER.IN_LOGOB_SUBOB " +
  "ON INUSER.IN_SUBOB.SUBOB_ID = INUSER.IN_LOGOB_SUBOB.SUBOB_ID " +
  "INNER JOIN INUSER.IN_LOGOB " +
  "ON INUSER.IN_LOGOB.LOGOB_ID = INUSER.IN_LOGOB_SUBOB.LOGOB_ID " +
  "INNER JOIN INUSER.IN_VERSION " +
  "ON INUSER.IN_VERSION.VERSION_ID = INUSER.IN_LOGOB.VERSION_ID " +
  "INNER JOIN INUSER.IN_DOC " +
  "ON INUSER.IN_DOC.DOC_ID = INUSER.IN_VERSION.DOC_ID " +
  "WHERE INUSER.IN_SUBOB.SUBOB_ID = '" + stamp.id + "';";

  var returnVal; 
  var cur = getHostDBLookupInfo_cur(stamp_sql,returnVal);
  //if there are no folders with the RTM flag or if we couldn't get a list
  if(!cur || cur == null)
  {
    debug.log("ERROR","Couldn't find where the stamp has been applied!  Applying validation to application\n");
  }

  var subobjTempl = new INSubobTemplate(VALIDATION_STAMP, SubobType.Stamp);
  subobjTempl.getInfo();
  if (!subobjTempl.id)
  {
      debug.log("ERROR","Template doesn't exist: %s.\n", stamp.templId);
      return false;
  }
  else
  {
      var page = parseInt(cur[1]);
      var logob = new INLogicalObject(cur[0], -1, page);
      var valSubob = new INSubObject();

      valSubob.templId = subobjTempl.id;
      var f = new INFont("Calibri", 11, 0xff0000);
      valSubob.font = f;
      valSubob.text = "USER ID: " + stamp.creationUserName + "\nMember Of: " + groups + "\nDecision Stamp: " + stamptmpl + "\nStamp Text: " + stamp.text+ "\nSource Queue:" + queue.queueName;
      valSubob.location = "-5,-5";
      valSubob.color = "0xFF0000";
      valSubob.justify = 1;
      valSubob.frameStyle = 1;
      valSubob.fillColor = 0xffff;
      for(var key in valSubob)
      {
        printf(key + " " + valSubob[key] + "\n")
      }
      if (logob.storeSubObject(SubobType.Stamp, valSubob))
      {
          printf("Text subobj applied for doc: %s.\n", logob.docId);
      }
      else
      {
          printf("Error: %s.\n", getErrMsg());
      }
  }
} // end applyValidationStamp

//