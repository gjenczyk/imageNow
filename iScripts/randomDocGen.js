/********************************************************************************
        Name:          randomDocGen
        Author:        Gregg Jenczyk
        Created:        07/08/2015
        Last Updated:   
        For Version:    
---------------------------------------------------------------------------------
        Summary:
               Makes a bunch of random docs using a csv   
               
        Mod Summary:
               Date-Initials: Modification description.



********************************************************************************/

// ********************* Include additional libraries *******************
//#link "inxml"    //XML parser
//#link "sedbc"    //Database object
//#link "secomobj" //COM object
#include "$IMAGENOWDIR6$\\script\\lib\\iScriptDebug.jsh"
#include "$IMAGENOWDIR6$\\script\\lib\\HostDBLookupInfo.jsh"
#include "$IMAGENOWDIR6$\\script\\STL\\packages\\System\\generateUniqueID.js"


// *********************         Configuration        *******************

// logging
#define LOG_TO_FILE         true    // false - log to stdout if ran by intool, true - log to inserverXX/log/ directory
#define DEBUG_LEVEL         5       // 0 - 5.  0 least output, 5 most verbose
#define SPLIT_LOG_BY_THREAD false   // set to true in high volume scripts when multiple worker threads are used (workflow, external message agent, etc)
#define MAX_LOG_FILE_SIZE   100     // Maximum size of log file (in MB) before a new one will be created


// *********************       End  Configuration     *******************

// ********************* Initialize global variables ********************

var CSV_PATH = imagenowDir6+"\\script\\MOCK_DATA.csv";
var TXT_PATH = imagenowDir6+"\\script\\MOCK_DATA.txt";
var QUEUE = "DFA Link Documents";//Router"; //
var DRAWER = "UMDFA";
var DOC_TYPE_LIST = "UMD Financial Aid";
var CP_SOURCE = "iScript";
var CP_APP_PLAN = "COIN FA Custom Page";


/**
* Main body of script.
* @method main
* @return {Boolean} True on success, false on error.
*/
  function main ()
{
    try
    {
      debug = new iScriptDebug("randomDocGen", LOG_TO_FILE, DEBUG_LEVEL);
      debug.log("WARNING", "randomDocGen script started.\n");


      debug.log("DEBUG","CSV_PATH is [%s]\n", CSV_PATH);
      var dtList = new INDocTypeList("", DOC_TYPE_LIST);
            if(!dtList.getInfo())
            {
              return false;
            }
            
            var count = 0;

      var fp = Clib.fopen(CSV_PATH, "r");
      if ( fp == null )
         printf(
              "\aError opening file for reading.\n")
      else
         while ( null != (line=Clib.fgets(fp)) )
         {
            count+= 1;
            printf(count+"\n");
            var u = Math.floor(Math.random() * dtList.members.length);
            //Clib.fputs(line, stdout)
            line = line.split(',');
            var emplid = '0'+line[0];
            //printf("emplid: "+emplid+"\n")
            var name = line[2] + ',' + line[1];
            //printf("name: "+name+"\n")
            var aidYear = line[3];
            //printf("aidYear: "+aidYear+"\n")
            var career = line[4];
            //printf("career: "+career+"\n")
            var linkDate = line[5].substring(0,line[5].length-1);
            //printf("linkDate: "+linkDate+"\n")
            var docType = dtList.members[u].name
            //printf(dtList.members[u].name+"\n")
            var doc = new INDocument(DRAWER, emplid, name, aidYear, career, linkDate, docType);
            var props = new Array();
            var prop1 = new INInstanceProp();
            prop1.name = "Source";
            prop1.setValue(CP_SOURCE);
            props.push(prop1);
            var prop2 = new INInstanceProp();
            prop2.name = "Document App Plan";
            prop2.setValue(CP_APP_PLAN);
            props.push(prop2);
            if (!doc.create(props))
            {

                  debug.log("ERROR","Failed to create new document: %s.\n", getErrMsg());

            }
              var returnID = "";
              var newName = generateUniqueID();
              //path[1] = generateUniqueID();
              if (!doc.setName(newName, returnID))
              {
                printf("Fail = %s.\n", getErrMsg());
              }
              else
              {
/*var attr = new Array;
attr["phsob.file.type"] = "txt";
attr["phsob.working.name"] = "MOCK_DATA.txt";
attr["phsob.source"]="PhsobSource.Iscript";
var logob = doc.storeObject(TXT_PATH, attr);
if (logob == null)
{
    printf("Failed. Error:%s.\n", getErrMsg());
}
else
{
    printf("Successful. logobID: %s.\n", logob.id);
}*/
                var wfQueue = new INWfQueue();
                wfQueue.name = QUEUE;
                wfQueue.createItem(WfItemType.Document, doc.id, WfItemPriority.Medium);
              }

         }
      Clib.fclose(fp);

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


//