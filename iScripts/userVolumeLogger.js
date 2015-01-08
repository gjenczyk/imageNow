/********************************************************************************
        Name:          userVolumeLogger
        Author:        Gregg Jenczyk
        Created:        12/15/2014
        Last Updated:   
        For Version:    
---------------------------------------------------------------------------------
        Summary:
               Create a users.txt file - keep appending the output of the 
               attached query to the txt file every hour, send a .csv file every 
               day @ midnight - this will keep track of usage.   
               
        Mod Summary:
               Date-Initials: Modification description.



********************************************************************************/

// ********************* Include additional libraries *******************
//#link "inxml"    //XML parser
//#link "sedbc"    //Database object
//#link "secomobj" //COM object
#include "$IMAGENOWDIR6$\\script\\lib\\iScriptDebug.jsh"
#include "$IMAGENOWDIR6$\\script\\lib\\HostDBLookupInfo.jsh"

// *********************         Configuration        *******************

// logging
#define LOG_TO_FILE         true    // false - log to stdout if ran by intool, true - log to inserverXX/log/ directory
#define DEBUG_LEVEL         5       // 0 - 5.  0 least output, 5 most verbose
#define SPLIT_LOG_BY_THREAD false   // set to true in high volume scripts when multiple worker threads are used (workflow, external message agent, etc)
#define MAX_LOG_FILE_SIZE   100     // Maximum size of log file (in MB) before a new one will be created


// *********************       End  Configuration     *******************

// ********************* Initialize global variables ********************

var logDir = imagenowDir6 + "\\script\\log\\";
var POWERSHELL_EMAIL = "\\\\boisnas215c1.umasscs.net\\diimages67tst\\script\\PowerShell\\processUserVolumeLogs.ps1"
var POWERSHELL_ROOT = "C:\\Windows\\System32\\WindowsPowerShell\\v1.0\\powershell.exe";

sql = "SELECT IN_SC_USR.USR_ID, "+
        "IN_SC_USR.USR_NAME, "+
        "IN_SC_USR.USR_FIRST_NAME, "+
        "IN_SC_USR.USR_LAST_NAME, "+
        "IN_LIC_MON.IS_ACTIVE, "+
        "IN_SC_USR.USR_ORG, "+
        "IN_LIC_MON.LOGIN_TIME, "+
        "IN_SC_USR.USR_ORG_UNIT "+
      "FROM IN_SC_USR "+
      "INNER JOIN IN_LIC_MON "+
      "ON IN_SC_USR.USR_ID = IN_LIC_MON.USR_ID "+
      "WHERE IN_SC_USR.USR_NAME NOT LIKE '%agent%' "+
      "AND IN_SC_USR.USR_NAME NOT LIKE '%inserver%' "+
      "AND IN_LIC_MON.IS_ACTIVE  = 1";

      var userTotal = 0;
      var bostonCount = 0;
      var dartmouthCount = 0;
      var lowellCount = 0;
      var uitsCount = 0;
      var otherCount = 0;

/**
* Main body of script.
* @method main
* @return {Boolean} True on success, false on error.
*/
  function main ()
{
    try
    {
      debug = new iScriptDebug("userVolumeLogger", LOG_TO_FILE, DEBUG_LEVEL);
      debug.log("WARNING", "userVolumeLogger script started.\n");

      var date = new Date();
      var currentDate = "" + (date.getMonth()+1) + date.getDate() + date.getFullYear();
      var currentTime = date.toString().split(" ");
      currentTime = currentTime[3].split(":");
      var currentHour = currentTime[0];

      // packages up the time object nicely
      var time = { curDate : currentDate,
                   curTime : currentTime,
                   curHour : currentHour};

      //get the current user count broken down by location
      hourlyUserCount(time);
      //get the users who have logged in since the last running
      currentUsers(time);
      


      //if the time is X, call the powershell script to process the previous day's output
      if (currentHour == 23)
      {
        //kick off powershell script
          //read the file
          //convert the data to a pivot table
          //email the pivot table
          //delete the old log after it's been processed!
          var cmd = "";
          Clib.sprintf(cmd, '%s %s', POWERSHELL_ROOT, POWERSHELL_EMAIL);
          var rtn = exec(cmd, 0);
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

/****************************************************************

  This function increments the campus count global variables

****************************************************************/
function incrementCampusCount(campusOffice){
  var co = campusOffice.toUpperCase()
  if (co == "BOSTON")
  {
    bostonCount++;
  }
  else if (co == "DARTMOUTH")
  {
    dartmouthCount++;
  }
  else if (co == "LOWELL")
  {
    lowellCount++;
  }
  else if (co == "UITS")
  {
    uitsCount++;
  }
  else
  {
    otherCount++;
  }
} // end incrementCampusCount

/****************************************************************

  This function gets the total number of users logged in at the time the script ran

****************************************************************/

function hourlyUserCount(timeObj){
 var returnVal;
      var cur = getHostDBLookupInfo_cur(sql,returnVal); 
            
      if(!cur || cur == null)
      {
        debug.log("WARNING","no results returned for query.\n");
        return false;
      }

      while(cur.next())
      {
        incrementCampusCount(cur[5]);
        userTotal++;
      } 

      debug.log("INFO","Hour [%s] USERTOTAL: %s | bostonCount = %s; dartmouthCount = %s; lowellCount = %s; uitsCount = %s; otherCount = %s;\n", timeObj.curHour ,userTotal, bostonCount, dartmouthCount, lowellCount, uitsCount, otherCount);
      //build the string to write to the csv
      var csvLine = timeObj.curHour + "," + userTotal + "," + bostonCount + "," + dartmouthCount + "," + lowellCount + "," + uitsCount + "," + otherCount + "\n";

      var userLog = logDir + "userVolumeLogger_R_" + timeObj.curDate + ".csv";

      var uL = Clib.fopen(userLog, "a");
      Clib.fputs(csvLine, uL);
      Clib.fclose(uL);
} // end hourlyUserCount

/****************************************************************

  This function gets the users who have logged in since the last time the script ran

****************************************************************/

function currentUsers(timeObj){

  sql += " AND IN_LIC_MON.LOGIN_TIME = TO_CHAR(IN_LIC_MON.LOGIN_TIME,'YYYY-MON-DD HH24:MI:SS')"

  var returnVal;
  var cur = getHostDBLookupInfo_cur(sql,returnVal); 
            
  if(!cur || cur == null)
  {
    debug.log("WARNING","no results returned for query.\n");
    return false;
  }

  var numCol = cur.columns();
  var detailLog = logDir + "userVolumeLogger_D_" + timeObj.curDate + ".csv";
  var dL = Clib.fopen(detailLog, "a");
  var detailedLine = "";

  while(cur.next())
  {
    for (var i = 0; i < numCol-1; i ++)
    {
      detailedLine += cur[i] + ",";
    }

      detailedLine += cur[numCol-1] + "\n";

      Clib.fputs(detailedLine, dL);
      incrementCampusCount(cur[5]);
      userTotal++;
  } 

    Clib.fclose(dL);
} // end current users count

function exec(cmd, expected_return)
{
  //debug.log("INFO", "Exec cmd: *%s*\n", cmd);
  var rtn;
  rtn = Clib.system(cmd);
  debug.log("DEBUG", "exec returned %s\n", rtn);
  // tiffcp doesn't return 0 on success
  if(rtn != expected_return)
  {
    debug.log("ERROR", "Couldn't call system cmd: %s\n", cmd);
    return false;
  }
  else
  {
    return true;
  }
} // end exec

//