/********************************************************************************
        Name:          PurgeMultiContents
        Author:        Gregg Jenczyk
        Created:        3/11/2014
        Last Updated:   
        For Version:    
---------------------------------------------------------------------------------
        Summary:
               This script will blow away the folders created in 65 that were
               migrated to 67.  It removes both the folders and the contents.
               To convert the script to handle other versions:
                          65  <---> 67      
                   INProject  <---> INFolder
                .getDocList() <---> .getShortcuts()
               
        Mod Summary:
               Date-Initials: Modification description.


Cut and paste SQL:
SELECT count(*)
FROM
  (SELECT INUSER.IN_INSTANCE.INSTANCE_NAME,
    INUSER.IN_INSTANCE.CLASS_TYPE,
    INUSER.IN_INSTANCE.DELETION_STATUS
  FROM INUSER.IN_INSTANCE
  WHERE INUSER.IN_INSTANCE.CLASS_TYPE    = '2'
  AND INUSER.IN_INSTANCE.DELETION_STATUS = '0'
  ORDER BY INUSER.IN_INSTANCE.INSTANCE_NAME
  )
INNER JOIN INUSER.IN_PROJ
ON INUSER.IN_PROJ.PROJ_NAME = INSTANCE_NAME
INNER JOIN INUSER.IN_PROJ_TYPE
ON INUSER.IN_PROJ.PROJ_TYPE_ID = INUSER.IN_PROJ_TYPE.PROJ_TYPE_ID
--WHERE RowNum                  <= '500'
WHERE INUSER.IN_PROJ_TYPE.PROJ_TYPE_NAME like '%LUA'

SELECT count(*)
FROM INUSER.IN_DOC
INNER JOIN INUSER.IN_DRAWER
ON INUSER.IN_DOC.DRAWER_ID = INUSER.IN_DRAWER.DRAWER_ID
INNER JOIN INUSER.IN_INSTANCE
ON INUSER.IN_DOC.INSTANCE_ID      = INUSER.IN_INSTANCE.INSTANCE_ID
WHERE INUSER.IN_DOC.IS_IN_PROJECT = '0'
AND INUSER.IN_DRAWER.DRAWER_NAME LIKE '%LUA%'
AND INUSER.IN_INSTANCE.DELETION_STATUS = '0'
               
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
#define TEST                false    //IF true, see what will be deleted when script runs.
#define PURGE_FOLDERS       false   //If true, run purgeFolders function
#define PURGE_DOCUMENTS     true    //If true, run purgeDocuments function

// *********************       End  Configuration     *******************

// ********************* Initialize global variables ********************
var NUM_RESULTS = 4150;
var TARGET_DRAWER_FOLDER = ["DFA"];
var TARGET_DRAWER_DOC = ["DFA"];


/**
* Main body of script.
* @method main
* @return {Boolean} True on success, false on error.
*/
function main ()
{
        try
        { 
            debug = new iScriptDebug("USE SCRIPT FILE NAME", LOG_TO_FILE, DEBUG_LEVEL);
            debug.log("WARNING", "PurgeMultiContents.js starting.\n");

            if(NUM_RESULTS == null || isNaN(NUM_RESULTS))
            {
              printf("Invalid value for NUM_RESULTS: [%s]\n", NUM_RESULTS);
            }

            
            if (PURGE_FOLDERS) {
              for (i=0; i < TARGET_DRAWER_FOLDER.length; i++) {
                purgeFolders(TARGET_DRAWER_FOLDER[i]);
              }
            }

            if (PURGE_DOCUMENTS) {
              for (i=0; i < TARGET_DRAWER_DOC.length; i++) {
                purgeDocuments(TARGET_DRAWER_DOC[i]);
              }
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

function purgeFolders(folderType) 
{
      //get a list of all folders of a certaintype
    var sql="SELECT * " +
          "FROM " +
          "(SELECT INUSER.IN_INSTANCE.INSTANCE_NAME, " +
          "INUSER.IN_INSTANCE.CLASS_TYPE, " +
          "INUSER.IN_INSTANCE.DELETION_STATUS " +
          "FROM INUSER.IN_INSTANCE " +
          "WHERE INUSER.IN_INSTANCE.CLASS_TYPE    = '2' " +
          "AND INUSER.IN_INSTANCE.DELETION_STATUS = '0' " +
          "ORDER BY INUSER.IN_INSTANCE.INSTANCE_NAME " +
          ") " +
          "INNER JOIN INUSER.IN_PROJ " +
          "ON INUSER.IN_PROJ.PROJ_NAME = INSTANCE_NAME " +
          "INNER JOIN INUSER.IN_PROJ_TYPE " +
          "ON INUSER.IN_PROJ.PROJ_TYPE_ID = INUSER.IN_PROJ_TYPE.PROJ_TYPE_ID " +
          "WHERE RowNum <= '"+ NUM_RESULTS +"'";

    if (!folderType || folderType == null)
    {
      debug.log("DEBUG","purgeFolders: No type specified\n");
    }
    else 
    {
      debug.log("DEBUG", "Target Drawer = [%s]\n", folderType);
      var sqlRestrict = " AND INUSER.IN_PROJ_TYPE.PROJ_TYPE_NAME like '%" + folderType + "'"
      sql += sqlRestrict;
    }
    printf("%s\n", folderType);
    var returnVal; 
    var cur = getHostDBLookupInfo_cur(sql,returnVal);
            
    if(!cur || cur == null)
    {
      debug.log("WARNING","no results returned for query.\n");
      return false;
    } 
    var recordsProcessed = 0;
    while(cur.next())
    {

      var fldName = cur[0];
      var itemType = cur[1];
      var deletionStatus = cur[2];
      var fldId = cur[3];

/*              debug.log("DEBUG","fldId is [%s] and fldName is [%s]\n", fldId, fldName);

                for (i=0;i<cur.columns();i++)
                {
                  printf("cur[%s] = %s\n",i,cur[i]);    
                }
*/
      if (deletionStatus == 1)
      {
        debug.log("ERROR","[%s] has already been deleted.\n", fldId);
        continue;
      }  
      
      var folder = INProject(fldId);

      if (folder == null || !folder)
      { 
        debug.log("ERROR","Failed to get folder. Error: %s\n", getErrMsg());
        continue;
      }

      else
      {
        debug.log("DEBUG", "Found [%s] [%s], checking for contents....\n", fldName, fldId);

        var folderContents = folder.getDocList();

        if (folderContents == null || !folderContents)
        { 
          debug.log("ERROR","Failed to get folderContents. Error: %s\n", getErrMsg());
          continue;
        }

        debug.log("DEBUG","folderContents.length = [%s]\n", folderContents.length);
          
        for(var i=0; i<folderContents.length; i++)
        {
          debug.log("DEBUG","Found Document ID: [%s]\n",folderContents[i].id);

          var docToDelete = new INDocument(folderContents[i].id)
                    
          docToDelete.getInfo();

          if (docToDelete == null || !docToDelete)
          { 
            debug.log("ERROR","Failed to get docToDelete. Error: %s\n", getErrMsg());
            continue;
          }

          if (!TEST) {

            if (!docToDelete.remove())
            {
              debug.log("ERROR","Failed to delete ID: [%s],  Reason: [%s]\n", docToDelete.id, getErrMsg());
              continue;
            }
            else
            {
              debug.log("INFO","Deleted [%s]\n",docToDelete.id);
            }
          } //end of if TEST 
        }

        if (!TEST) { 
                  
          if (!folder.remove())
          {
            debug.log("ERROR","Failed to delete ID: [%s] NAME: [%s],  Reason: [%s]\n", fldId, fldName, getErrMsg());
            continue;
          }
          else
          {
            debug.log("INFO","[%s of %s]Deleted [%s], Name: [%s]\n", recordsProcessed+1, NUM_RESULTS, fldId, fldName);
            printf("[%s of %s]Deleted [%s], Name: [%s]\n", recordsProcessed+1, NUM_RESULTS, fldId, fldName);
            recordsProcessed++;
          }
        }
      }
        debug.log("DEBUG", "Done with [%s] [%s]\n", fldId, fldName);  
    }
      debug.log("INFO", "Deleted [%s] folders\n", recordsProcessed);

} //end purgeFolders

function purgeDocuments (drawerToPurge)
{
  var sql= "SELECT * " +  
           "FROM INUSER.IN_DOC " +
           "INNER JOIN INUSER.IN_DRAWER " +
           "ON INUSER.IN_DOC.DRAWER_ID = INUSER.IN_DRAWER.DRAWER_ID " +
           "INNER JOIN INUSER.IN_INSTANCE " +
           "ON INUSER.IN_DOC.INSTANCE_ID = INUSER.IN_INSTANCE.INSTANCE_ID " +
           "AND INUSER.IN_INSTANCE.DELETION_STATUS = '0' " +
           "WHERE IS_IN_PROJECT = '0'" +
           "AND RowNum <= '"+ NUM_RESULTS +"'";

  if (!drawerToPurge || drawerToPurge == null)
  {
    debug.log("DEBUG","purgeDocuments: No Drawer specified\n");
  }
  else 
  {
    debug.log("DEBUG", "Target Drawer = [%s]\n", drawerToPurge);
    var sqlRestrict = " AND INUSER.IN_DRAWER.DRAWER_NAME like '%" + drawerToPurge + "%'"
    sql += sqlRestrict;
  }

  var returnVal; 
  var cur = getHostDBLookupInfo_cur(sql,returnVal);
            
  if(!cur || cur == null)
  {
    debug.log("WARNING","no results returned for query.\n");
    return false;
  }

  var recordsProcessed = 0;

  while(cur.next())
  {
     var docId = cur[0];

     var doc = new INDocument(docId);

     doc.getInfo();

     debug.log("INFO","Found [%s] [%s] [%s]\n", doc.id, doc.drawer, doc.docTypeName);

     if (!TEST)
     {
        if (!doc.remove())
        {
          debug.log("ERROR","Could not delete [%s]. Error: [%s]\n", doc.id, getErrMsg());
          continue;
        }
        else
        {
          debug.log("INFO","Succesfully deleted [%s] [%s of %s]\n", doc.id, recordsProcessed+1, NUM_RESULTS);
          printf("Succesfully deleted [%s] [%s] [%s of %s]\n", doc.id, doc.drawer, recordsProcessed+1, NUM_RESULTS);
          recordsProcessed++;
        }
     }
  }
  debug.log("INFO","Deleted [%s] documents\n", recordsProcessed)
}
//