/********************************************************************************
        Name:          GA_DocumentExport
        Author:        Gregg Jenczyk
        Created:        04/03/201
        Last Updated:   
        For Version:    
---------------------------------------------------------------------------------
        Summary:
               This script will convert documents for an applicant to one or more 
               pdfs.
               
        Mod Summary:
               Date-Initials: Modification description.



********************************************************************************/

// ********************* Include additional libraries *******************
//#link "inxml"    //XML parser
//#link "sedbc"    //Database object
//#link "secomobj" //COM object
#include "$IMAGENOWDIR6$\\script\\lib\\iScriptDebug.jsh"
#include "$IMAGENOWDIR6$\\script\\lib\\HostDBLookupInfo.jsh"
#include "$IMAGENOWDIR6$\\script\\STL\\packages\\Document\\exportDocPhsOb.js"
#include "$IMAGENOWDIR6$\\script\\lib\\yaml_loader.jsh"
#include "$IMAGENOWDIR6$\\script\\lib\\csvObject.jsh"
#include "$IMAGENOWDIR6$\\script\\lib\\GetUniqueF5DateTime.js"
#include "$IMAGENOWDIR6$\\script\\lib\\GetProp.jsh"

// *********************         Configuration        *******************

// logging
#define LOG_TO_FILE         true    // false - log to stdout if ran by intool, true - log to inserverXX/log/ directory
#define DEBUG_LEVEL         5       // 0 - 5.  0 least output, 5 most verbose
#define SPLIT_LOG_BY_THREAD false   // set to true in high volume scripts when multiple worker threads are used (workflow, external message agent, etc)
#define MAX_LOG_FILE_SIZE   100     // Maximum size of log file (in MB) before a new one will be created


// *********************       End  Configuration     *******************

// ********************* Initialize global variables ********************

var POWERSHELL_ROOT = "C:\\Windows\\System32\\WindowsPowerShell\\v1.0\\powershell.exe";
var POWERSHELL_MERGE_TIFF = imagenowDir6+"\\script\\PowerShell\\mergeTiffs.ps1"
var POWERSHELL_MERGE_PDF = imagenowDir6+"\\script\\PowerShell\\mergePDFs.ps1"

/**
* Main body of script.
* @method main
* @return {Boolean} True on success, false on error.
*/
  function main ()
{
    try
    {
      debug = new iScriptDebug("GA_DocumentExport", LOG_TO_FILE, DEBUG_LEVEL);
      debug.log("WARNING", "GA_DocumentExport script started.\n");

      //get informaiton about the current wf item
      var wfItem = new INWfItem(currentWfItem.id);//"321YZ49_06Y84KJ6000000M");//
      if(!wfItem.id || !wfItem.getInfo())
      {
        debug.log("CRITICAL", " Couldn't get info for wfItem: %s\n", getErrMsg());
        return false;
      }

      //get queue name
      var queueName = wfItem.queueName;

      //store the results from the yaml here
      var docTypesUsed = [];
      var docTypesNotUsed = [];
      var docTypeOrder = "";
      var pdfDocType = "";
      var trigger = "";
      var emailFlag = false;
      var importFlag = false;
      var singleDoc = false;

      //identify the correct doctype list to use for exporting in the yaml
      loadYAMLConfig(imagenowDir6+"\\script\\config_scripts\\GA_DocumentExport\\");
      for (var office in CFG.GA_DocumentExport)
      {
        var exportConfig = CFG.GA_DocumentExport[office].EXPORT_CONFIG;
        for (var i = 0; i < exportConfig.length ; i++)
        {
          if (exportConfig[i].SCRIPT_Q == queueName)
          {
            docTypesUsed = exportConfig[i].TYPES_TO_INCLUDE;
            docTypesNotUsed = exportConfig[i].TYPES_TO_EXCLUDE;
            docTypeOrder = exportConfig[i].DOCTYPE_ORDER;
            pdfDocType = exportConfig[i].IN_PDF_DOC;
            trigger = exportConfig[i].TRIGGER;
            emailFlag = exportConfig[i].EMAIL_OUTPUT;
            importFlag = exportConfig[i].IMPORT_OUTPUT;
            singleDoc = exportConfig[i].SINGLE_FILE;
            break;
          }
        }
      }//end checking the yaml for the list

      //get out if we can't find a list in the config
      if((!docTypesUsed || docTypesUsed == null) || (!docTypesNotUsed || docTypesNotUsed == null) || (!pdfDocType || pdfDocType == null) || (!trigger || trigger == null))
      {
        debug.log("ERROR","Invalid config for [%s] - docTypesUsed:[%s] pdfDocType:[%s] pdfDocQCfg:[%s]\n",queueName, docTypesUsed, pdfDocType, trigger);
        return false;
      }

      //get information about inbound doc
      var doc = new INDocument(wfItem.objectId);//"301YY4P_04VZ7RN4N0142XW");  
      if(!doc.getInfo())
      {
          debug.log("ERROR","Couldn't get doc info: [%s]\n",getErrMsg()); 
          return false;
      }

      //check to make sure that the document is the correct type for pdf generation
      if(!(doc.docTypeName == trigger))
      {
        debug.log("INFO","Current doc: [%s] is not the correct type for pdf generation: [%s]\n", doc.docTypeName, trigger);
        return false;
      }

      //get information for the query
      var useDocs = prepareTypeLists(docTypesUsed);
      var skipDocs = prepareTypeLists(docTypesNotUsed);
      var saAppNo = GetProp(doc, "SA Application Nbr");
   
      //get a list of documents to send
      var docList = findDocsToSend(doc.field1, doc.drawer, saAppNo, useDocs, skipDocs);
      if (!docList || docList == null)
      {
        debug.log("ERROR","Couln't find exportable documents for [%s] in [%s] and [%s]\n",doc.field1, doc.drawer, docTypeList);
        return false;
      }

      //export all documents returned by the query
      if(!processMatchingDocs(docList))
      {
        debug.log("ERROR","Unable to export all documents.\n");
        return false;
      }

      //get the documents that need to go at the front of the pdf
      var sortOrder = retrieveSortOrder(docTypeOrder);
      if(!sortOrder || sortOrder == null)
      {
        debug.log("INFO","No configured order for pdf.\n");
      }

      //get the list of exported files
      var exportList = getExportedFiles(doc, sortOrder, singleDoc);
      if(!exportList || exportList == null)
      {
        debug.log("ERROR","Unable to get list of exported files!\n");
        return false;
      }

/*      if(emailFlag) // ------------------- This isn't going to work if we're going by profile sheet
      {
        //use this if we're going to be emailing stuff to someone...
        var routerEmail = getRouterEmail(wfItem.queueStartUserName);
      }
*/ 
      if(importFlag)
      {
        var revDoc = createNewDoc(doc, exportList, pdfDocType);
        if(!revDoc || revDoc == null)
        {
          debug.log("ERROR","Unable to create the pdf document.\n");
          return false;
        }
      }//end if(importFlag)

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

//this function turns an array into a string with each element enclosed in single
//quotes and seperated by a comma
function prepareTypeLists(typeArray)
{
  var sqlString = "";
  for (var t = 0; t < typeArray.length; t++)
  {
     var tempPrep = "'"+typeArray[t]+"'";
     sqlString += tempPrep;
     if(t < (typeArray.length - 1))
     {
      sqlString += ",";
     }
  }
  //return the formatted string
  return sqlString;
}//end of prepareTypeLists

//function to get list of documents that can be added to the pdf
function findDocsToSend(emplid, drawer, appNo, usedType, skipType)
{
  sql = "SELECT DISTINCT(INUSER.IN_DOC.DOC_ID), " +
        "INUSER.IN_DOC_TYPE.DOC_TYPE_NAME " +
        "FROM INUSER.IN_DOC " +
        "INNER JOIN INUSER.IN_DOC_TYPE " +
        "ON INUSER.IN_DOC_TYPE.DOC_TYPE_ID = INUSER.IN_DOC.DOC_TYPE_ID " +
        "INNER JOIN INUSER.IN_DOC_TYPE_LIST_MEMBER " +
        "ON INUSER.IN_DOC_TYPE.DOC_TYPE_ID = INUSER.IN_DOC_TYPE_LIST_MEMBER.DOC_TYPE_ID " +
        "INNER JOIN INUSER.IN_DOC_TYPE_LIST " +
        "ON INUSER.IN_DOC_TYPE_LIST.DOC_TYPE_LIST_ID = INUSER.IN_DOC_TYPE_LIST_MEMBER.DOC_TYPE_LIST_ID " +
        "INNER JOIN INUSER.IN_DRAWER " +
        "ON INUSER.IN_DRAWER.DRAWER_ID = INUSER.IN_DOC.DRAWER_ID " +
        "INNER JOIN INUSER.IN_INSTANCE " +
        "ON INUSER.IN_DOC.INSTANCE_ID = INUSER.IN_INSTANCE.INSTANCE_ID " +
        "INNER JOIN INUSER.IN_INSTANCE_PROP " +
        "ON INUSER.IN_INSTANCE.INSTANCE_ID = INUSER.IN_INSTANCE_PROP.INSTANCE_ID " +
        "INNER JOIN INUSER.IN_PROP " +
        "ON INUSER.IN_PROP.PROP_ID  = INUSER.IN_INSTANCE_PROP.PROP_ID " +
        "WHERE INUSER.IN_DOC.FOLDER = '" + emplid + "' " +
        "AND INUSER.IN_DRAWER.DRAWER_NAME LIKE '" + drawer + "%' " +
        "AND ((INUSER.IN_PROP.PROP_NAME         = 'SA Application Nbr' " +
        "AND INUSER.IN_INSTANCE_PROP.STRING_VAL = '" + appNo + "') " +
        "OR (INUSER.IN_PROP.PROP_NAME           = 'Shared' " +
        "AND INUSER.IN_INSTANCE_PROP.STRING_VAL = '301YT7N_000CFJ25Y0000NX')) " +
        "AND INUSER.IN_DOC_TYPE_LIST.LIST_NAME IN (" + usedType + ") " +
        "MINUS " +
        "SELECT INUSER.IN_DOC.DOC_ID, " +
        "INUSER.IN_DOC_TYPE.DOC_TYPE_NAME " +
        "FROM INUSER.IN_DOC " +
        "INNER JOIN INUSER.IN_DOC_TYPE " +
        "ON INUSER.IN_DOC_TYPE.DOC_TYPE_ID = INUSER.IN_DOC.DOC_TYPE_ID " +
        "INNER JOIN INUSER.IN_DOC_TYPE_LIST_MEMBER " +
        "ON INUSER.IN_DOC_TYPE.DOC_TYPE_ID = INUSER.IN_DOC_TYPE_LIST_MEMBER.DOC_TYPE_ID " +
        "INNER JOIN INUSER.IN_DOC_TYPE_LIST " +
        "ON INUSER.IN_DOC_TYPE_LIST.DOC_TYPE_LIST_ID = INUSER.IN_DOC_TYPE_LIST_MEMBER.DOC_TYPE_LIST_ID " +
        "INNER JOIN INUSER.IN_DRAWER " +
        "ON INUSER.IN_DRAWER.DRAWER_ID = INUSER.IN_DOC.DRAWER_ID " +
        "WHERE INUSER.IN_DRAWER.DRAWER_NAME LIKE '" + drawer + "%' " +
        "AND INUSER.IN_DOC.FOLDER = '" + emplid + "' " +
        "AND INUSER.IN_DOC_TYPE_LIST.LIST_NAME IN (" + skipType + ");";

  var returnVal;
  var cur = getHostDBLookupInfo_cur(sql,returnVal);

  if(!cur || cur == null)
  {
    debug.log("WARNING","no results returned for query.\n");
    return false;
  }
  //return the cursor object
  return cur; 
}// end of findDocsToSend

//function that sends each row from the query results to the export function
function processMatchingDocs(curObj)
{
  var count = 0;
  while(curObj.next())
  {
    count++;
    var tiffDoc = new INDocument(curObj[0]);
    if(!tiffDoc.getInfo())
    {
        debug.log("ERROR","Couldn't get doc info: [%s]\n",getErrMsg()); 
        return false;
    }

    if(!exportDoc(tiffDoc, count))
    {
      debug.log("ERROR","Could not export [%s]\n",tiffDoc.id);
      return false;
    }
  }//end of the exporting
  debug.log("INFO","Finished exporting documents.\n");
  return true;
}//end processMatchingDocs

//function that extracts the tiffs from the doc and converts them to a single pdf
function exportDoc(doc, seq)
{
  //make an output dir for the applicant's info because for some damn reason perceptive won't create a dir 2 deep
  //sorry, it's lexmark now
  var exportDir = imagenowDir6+"\\output\\"+doc.field1+"\\";//+doc.id+"\\";
  Clib.mkdir(exportDir);
  if(!exportDocPhsOb(doc,exportDir + "\\"+doc.id+"\\","ALL","ALL",true))
  {
    debug.log("ERROR","Could not export [%s] - [%s]\n",doc.id,getErrMsg());
    return false;
  }
  else
  {
    var cmd = "";
    //windows' command line doesn't like spaces
    var safeSpace = "'" + doc.docTypeName + "'";
    Clib.sprintf(cmd, '%s %s %s %s %s %s', POWERSHELL_ROOT, POWERSHELL_MERGE_TIFF, doc.field1, doc.id, safeSpace, seq);
    var rtn = exec(cmd, 0);
    //retrun the exit code
    return rtn;
  }
}//end exportDoc

//function to get the sort order for the documents defined in the sort order list
function retrieveSortOrder(listType)
{
  var dtList = new INDocTypeList("",listType);
  dtList.getInfo();
  var dtArray = dtList.members;
  var strTypeArray = [];
  for(var l = 0; l < dtArray.length; l++)
  {
      strTypeArray.push(dtArray[l].name);
  }
  //return the sort order, if there is one
  return strTypeArray;
}//end retrieveSortOrder

//function that gets a list of the exported files and sorts them, if there is a sorting list
//this funciton will also merge the pdfs if only one pdf is desired
function getExportedFiles(docObj, docSortArr, singleDoc)
{
  var filePath = "D:\\inserver6\\output\\complete\\" + docObj.field1 + "\\";
  var pdfImports = SElib.directory(filePath + "*");

  //sort the files if there is a sort list
  if (docSortArr.length > 0)
  {
    pdfImports = orderExports(pdfImports, docSortArr);
  }
  //merge the files if there should be only one
  if(singleDoc)
  {
    var pArray = [];
    for(var i = 0; i < pdfImports.length; i++)
    {
     pArray.push("'"+pdfImports[i].name+"'");
    }
    var merged = mergePdfs(pArray, docObj.field1);
    pdfImports = SElib.directory(filePath + "*");
  }
  //return the list of files
  return pdfImports;
}//end getExportedFiles

//function to sort the files according to an doctype list.
//if doctypes are not on the list, they are added at the end of the array FCFS
function orderExports(dirObj, order)
{
  var orderedArray = [];
  //add the docs on the list to the the array according to the order
  for(var p = 0; p < order.length; p++)
  {
    for (var o = 0; o < dirObj.length; o++)
    {
      var parts = SElib.splitFilename(dirObj[o].name);
      if(order[p].indexOf(parts.name.slice(0, -2)) == 0)
      {
        orderedArray.push(dirObj[o]);
      }
    }
  }//end of adding the ordered docs

  //remove the already-added docs from the remaining docs
  for (var q = 0; q < orderedArray.length; q++)
  {
    dirObj = removeFromArray(orderedArray[q], dirObj);
  }

  //add the leftover docs to the array
  for (var i = 0; i < dirObj.length; i++)
  {
    orderedArray.push(dirObj[i]);
  }

  //return the sorted array
  return orderedArray;
}//end of orderExports

//function to remove items from an array, since iScrip doesn't have array.indexOf()
function removeFromArray(element, arr)
{
  for (var q = 0; q < arr.length; q++)
  {
    if(element.name == arr[q].name)
    {
      arr.splice(q,1);
      break;
    }
  }
  //return array with one less element
  return arr;
}//end removeFromArray

//function to merge multiple pdfs into one
function mergePdfs(pdfPath, pdfName)
{
  var cmd = "";
  Clib.sprintf(cmd, '%s %s %s %s', POWERSHELL_ROOT, POWERSHELL_MERGE_PDF, pdfPath, pdfName);
  var rtn = exec(cmd, 0);
  return rtn;
}//end mergePdfs

//function to create the document back in imageNow
function createNewDoc(docObj, expFiles, docObjType)
{
  //where the pdfs are
  var filePath = "D:\\inserver6\\output\\complete\\" + docObj.field1 + "\\";
  var pdfKeys = popKeys(docObj, docObjType);
  var props = docObj.getCustomProperties();
  var pdfDoc = new INDocument(pdfKeys);

  //make the document with the properties
  if(!pdfDoc.create(props))
  {
    debug.log("ERROR","Could not create a document to store the pdf!\n");
    return false;
  }

  debug.log("INFO","Created [%s]\n", pdfDoc);

  //for each file that's in the export dir
  for (var j = 0; j < expFiles.length; j++)
  {
    debug.log("INFO","Working file [%s] = [%s]\n",j, expFiles[j].name);
    var parts = SElib.splitFilename(expFiles[j].name);
    printf(parts.name+"\n");
    var workingPdf = parts.name + parts.ext;
    var attr = new Array;
    attr["phsob.file.type"] = parts.ext;
    attr["phsob.working.name"] = workingPdf;
    attr["phsob.source"]="GA_DocumentExport";
    //add the pdf to the document
    var logob = pdfDoc.storeObject(expFiles[j].name, attr);
    if (logob == null)
    {
        debug.log("ERROR","Could not import document:%s.\n", getErrMsg());
        return false;
    }
    else
    {
      debug.log("INFO","Successfully imported logobID: %s.\n", logob.id);
      //delete the pdf from the server
      Clib.remove(expFiles[j].name);
    }
  }//end of for each pdf

  //remove the export directory
  Clib.rmdir(filePath)
  debug.log("INFO","Removed [%s]\n",filePath);

  //return the newly created doc object
  return pdfDoc;
}//end createNewDoc

//make the keys for the new doc with a new linkDate/Time
function popKeys(docObj, docObjType)
{
  //set create time
  var linkDate_obj = new GetUniqueF5DateTime(true);

  //docKeys
  var keys = new INKeys(docObj.drawer,docObj.field1,docObj.field2,docObj.field3,docObj.field4,linkDate_obj.tellTheTime(),docObjType);
  debug.log("INFO","Keys: [%s]\n",keys.toString());
  //return the keys for the new document
  return keys;
}//end popKeys

//get the email address of the person who routed the document into the queue
//this isn't the way to do it because this is driven by the profile sheet.
function getRouterEmail(lastUser)
{
  //collect information about who routed the document into the queue
  var router = new INUser(lastUser);
  if(!router.getInfo())
  {
  debug.log("ERROR","Could not get user information for [%s] - [%s]\n", lastUser, getErrMsg());
  return false;
  }
  //make sure user has an email address on file
  if(!router.email || router.email == null)
  {
  debug.log("ERROR","No email address found for [%s - %s %s]\n",router.id, router.lastName, router.firstName);
  return false;
  }
  var emailAddr = router.email;
  //fun fact - manager users don't have email addresses in the system.
  return emailAddr;
}

//this function passes a command to the command line and returns the exit code (if it's defined!)
function exec(cmd, expected_return)
{
  debug.log("INFO", "Exec cmd: *%s*\n", cmd);
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
}//end exec

//