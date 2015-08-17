/********************************************************************************
        Name:         GetPageFileSize.js
        Author:        Gregg Jenczyk
        Created:        10/21/14
        Last Updated:   
        For Version:    
---------------------------------------------------------------------------------
        Summary:
               This script will update a custom property based on 
               
        Mod Summary:
               Date-Initials: Modification description.
               
SQL FOR DOCLIST:
SELECT count(INUSER.IN_DOC.DOC_ID)
FROM INUSER.IN_DOC
INNER JOIN INUSER.IN_INSTANCE
ON INUSER.IN_INSTANCE.INSTANCE_ID = INUSER.IN_DOC.INSTANCE_ID
INNER JOIN INUSER.IN_INSTANCE_PROP
ON INUSER.IN_INSTANCE.INSTANCE_ID        = INUSER.IN_INSTANCE_PROP.INSTANCE_ID
WHERE INUSER.IN_INSTANCE_PROP.STRING_VAL = 'CommonApp'
AND TRUNC(INUSER.IN_INSTANCE.CREATION_TIME) > TO_DATE('07/07/2014', 'dd-mm-yy');              


********************************************************************************/

// ********************* Include additional libraries *******************
//#link "inxml"    //XML parser
//#link "sedbc"    //Database object
//#link "secomobj" //COM object
#include "$IMAGENOWDIR6$\\script\\lib\\iScriptDebug.jsh"
#include "$IMAGENOWDIR6$\\script\\lib\\envVariable.jsh"

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
            debug = new iScriptDebug("GetPageFileSize", LOG_TO_FILE, DEBUG_LEVEL);
            debug.log("WARNING", "GetPageFileSize.js starting.\n");

            /* open the csv with the doc ids */
            var csvList = Clib.fopen(imagenowDir6+"\\script\\docs_to_check.csv","r");

            /* check to  */
            if ( csvList == null )
            {
              debug.log("ERROR","csv is either missing or empty.\n");
              return false;
            } /* end if ( csvList == null ) */

            while ( null != (line=Clib.fgets(csvList)) )
            {
              /* remove newline character from line */
              var curDoc = line.substring(0,line.length-1);

              var doc = new INDocument(curDoc);

              if(!doc.getInfo())
              {
                debug.log("ERROR","Couldn't get info for doc id: [%s], [%s]\n", doc.id, getErrMsg());
                continue;
              } /* if(!doc.getInfo()) */

                var ver = new INVersion(doc.id, -1);
                if (!ver.getInfo())
                {
                  debug.log("ERROR", "getPageSizes: Failed to get version info for document [%s]\n", doc);
                  return false;
                }
                var numPages = ver.logobCount;
                
                //get page sizes from document
                var returnSize = new Array();
                for (var i=1; i<=numPages; i++)
                {
                  //get the logob
                  var logob = new INLogicalObject(doc.id, -1, i);
                  if (!logob.getInfo())
                  {
                    debug.log("ERROR", "getPageSizes: Failed to get logob for page [%s] of document [%s]\n", i, doc);
                    return false;
                  }
                  else if (!logob.retrieveObject())
                  {
                    debug.log("ERROR", "getPageSizes: Failed to retrieve page [%s] of document\n", i, doc);
                    return false;
                  }
                  
                  //find the file
                  var files = SElib.directory(logob.filePath, false);
                  if (!files || files.length != 1)
                  { 
                    debug.log("ERROR", "getPageSizes: Failed to find the osm file for page [%s] of document [%s]\n", i, doc);
                    return false;
                  }
                  returnSize.push(files[0].size);
                }

                for (var j = 0; j < returnSize.length; j++)
                {
                  if (returnSize[j] < 10000)
                  {
                    printf("got one!!\n");
                    debug.log("WARNING","Small tiff size detected [%s] [%s bytes]\n", doc.id, returnSize[j]);
                  }
                }

            } /* end while ( null != (line=Clib.fgets(csvList)) ) */

          Clib.fclose(csvList);

        } /* end of try */
        
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