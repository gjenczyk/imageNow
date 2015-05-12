/********************************************************************************
        Name:          GA_DocumentExport
        Author:        Gregg Jenczyk
        Created:        04/03/201
        Last Updated:   
        For Version:    
---------------------------------------------------------------------------------
        Summary:
               This script will convert documents for an applicant a single file.
               2.0 - Security has been applied so the pdfs can't be printed
               3.0 - Output can be tiff for extra security               
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
#include "$IMAGENOWDIR6$\\script\\lib\\envVariable.jsh"

// *********************         Configuration        *******************

// logging
#define LOG_TO_FILE         true    // false - log to stdout if ran by intool, true - log to inserverXX/log/ directory
#define DEBUG_LEVEL         5       // 0 - 5.  0 least output, 5 most verbose
#define SPLIT_LOG_BY_THREAD false   // set to true in high volume scripts when multiple worker threads are used (workflow, external message agent, etc)
#define MAX_LOG_FILE_SIZE   100     // Maximum size of log file (in MB) before a new one will be created

// *********************       End  Configuration     *******************

// ********************* Initialize global variables ********************

var POWERSHELL_ROOT = "C:\\Windows\\System32\\WindowsPowerShell\\v1.0\\powershell.exe";
var POWERSHELL_MERGE_TIFF = imagenowDir6+"\\script\\PowerShell\\DocumentExport_tiffs.ps1"
var POWERSHELL_MERGE_DOC = imagenowDir6+"\\script\\PowerShell\\DocumentExport_files.ps1"
var POWERSHELL_EMAIL_DOC = imagenowDir6+"\\script\\PowerShell\\DocumentExport_email.ps1"
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

      if(typeof(currentWfItem) == "undefined")
      {
        debug.log("NOTIFY","Running GA_DocumentExport as an intool action\n");
        var docListPath = "D:\\inserver6\\script\\docs.csv"
        var fp = Clib.fopen(docListPath, "r");
        if (fp == null)
        {
          debug.log("ERROR","Could not read [%s] - [%s]\n", docListPath, getErrMsg());
        }
        else
        {
           while ( null != (line=Clib.fgets(fp)) )
           {
              Clib.fputs(line, stdout)
              var tstID = line.substring(0,line.length-1);
              printf(tstID);
              var curDocObj = new INDocument(tstID);
              //get informaiton about the current object
              if(!curDocObj.getInfo())
              {
                debug.log("CRITICAL","Failed to get info for [%s] : [%s]\n", tstID, getErrMsg());
                return false;
              }
              else
              {
                documentExport_intool(curDocObj);
              }      
          }

          Clib.fclose(fp);
        }
      }
      else
      {
        debug.log("NOTIFY","Running GA_DocumentExport as an inbound action\n");
        //get informaiton about the current wf item
        var wfItem = new INWfItem(currentWfItem.id);//"321YZ48_06Y23JJFR00003S");//
        if(!wfItem.id || !wfItem.getInfo())
        {
          debug.log("CRITICAL", " Couldn't get info for wfItem: %s\n", getErrMsg());
          return false;
        }

        documentExportWF(wfItem);
      }
    }//end main
        
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
//this function will do the conversion for inbound docs
function documentExportWF(wfObj)
{
  //get queue name
      var queueName = wfObj.queueName;

      //store the results from the yaml here
      var docTypesUsed = [];
      var docTypesNotUsed = [];
      var docTypeOrder = "";
      var singleDocType = "";
      var trigger = [];
      var emailFlag = false;
      var importFlag = false;
      var fileFormat = null;
      var protectPdf = true;

      //identify the correct doctype list to use for exporting in the yaml
      loadYAMLConfig(imagenowDir6+"\\script\\config_scripts\\GA_DocumentExport\\");
      for (var office in CFG.GA_DocumentExport)
      {
        var exportConfig = CFG.GA_DocumentExport[office].EXPORT_CONFIG;
        for (var i = 0; i < exportConfig.length ; i++)
        {
          debug.log("DEBUG","exportConfig[i].SCRIPT_Q == [%s]\n",exportConfig[i].SCRIPT_Q)
          if (exportConfig[i].SCRIPT_Q == queueName)
          {
            docTypesUsed = exportConfig[i].TYPES_TO_INCLUDE;
            docTypesNotUsed = exportConfig[i].TYPES_TO_EXCLUDE;
            docTypeOrder = exportConfig[i].DOCTYPE_ORDER;
            singleDocType = exportConfig[i].IN_SINGLE_DOC;
            trigger = exportConfig[i].TRIGGER;
            emailFlag = exportConfig[i].EMAIL_OUTPUT;
            importFlag = exportConfig[i].IMPORT_OUTPUT;
            fileFormat = exportConfig[i].FILE_FORMAT;
            protectPdf = exportConfig[i].PROTECT_PDF;
            break;
          }
        }
      }//end checking the yaml for the list

      //get out if we can't find a list in the config
      if((!docTypesUsed || docTypesUsed == null) || (!docTypesNotUsed || docTypesNotUsed == null) || (!singleDocType || singleDocType == null) || (!trigger || trigger == null))
      {
        debug.log("ERROR","Invalid config for [%s] - docTypesUsed:[%s] singleDocType:[%s] trigger:[%s]\n",queueName, docTypesUsed, singleDocType, trigger);
        return false;
      }

      //get information about inbound doc
      var doc = new INDocument(wfObj.objectId);//"301YY4P_04VZ7RN4N0142XW");  
      if(!doc.getInfo())
      {
          debug.log("ERROR","Couldn't get doc info: [%s]\n",getErrMsg()); 
          return false;
      }

      //check to make sure that the document is the correct type for generation
      var foundTrigger = false;
      for (var t = 0; t < trigger.length; t++)
      {
        if(doc.docTypeName == trigger[t])
        {
          debug.log("DEBUG","Found a valid trigger doctype [%s]\n", doc.docTypeName);
          foundTrigger = true;
        }
      }
      if(!foundTrigger)
      {
        debug.log("INFO","Current doc: [%s] [%s] is not the correct type for generation: [%s]\n", doc.id, doc.docTypeName, trigger);
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
      if(!processMatchingDocs(docList, saAppNo, fileFormat))
      {
        debug.log("ERROR","Unable to export all documents.\n");
        return false;
      }

      //get the documents that need to go at the front of the export
      var sortOrder = retrieveSortOrder(docTypeOrder);
      if(!sortOrder || sortOrder == null)
      {
        debug.log("INFO","No configured order for export.\n");
      }

      //get the list of exported files
      var exportList = getExportedFiles(doc, sortOrder, fileFormat, protectPdf, saAppNo);
      if(!exportList || exportList == null)
      {
        debug.log("ERROR","Unable to get list of exported files!\n");
        return false;
      }

      //if we are configured to send an email
      if(emailFlag) 
      {
        //use this if we're going to be emailing stuff to someone...
        var routerEmail = getRouterEmail(wfObj.queueStartUserName);
        if(!routerEmail || routerEmail == null)
        {
          debug.log("ERROR","Could not send email - no address found.\n");
        }
        else
        {
          //send the email
          sendFile(routerEmail, exportList, doc);
        }
      }//end if(emailFlag)

      //if we are configured to import a document
      if(importFlag)
      {
        var revDoc = createNewDoc(doc, exportList, singleDocType, saAppNo);
        if(!revDoc || revDoc == null)
        {
          debug.log("ERROR","Unable to create the new document.\n");
          return false;
        }
      }//end if(importFlag)

      //clean up files
      cleanUpServer(exportList);
}//end documentExportWF

//this function will convert docs via intool
function documentExport_intool(doc)
{
  //get queue name
      var queueName = doc.drawer + " intool";

      //store the results from the yaml here
      var docTypesUsed = [];
      var docTypesNotUsed = [];
      var docTypeOrder = "";
      var singleDocType = "";
      var trigger = [];
      var emailFlag = false;
      var importFlag = false;
      var fileFormat = null;
      var protectPdf = true;

      //identify the correct doctype list to use for exporting in the yaml
      loadYAMLConfig(imagenowDir6+"\\script\\config_scripts\\GA_DocumentExport\\");
      for (var office in CFG.GA_DocumentExport)
      {
        var exportConfig = CFG.GA_DocumentExport[office].EXPORT_CONFIG;
        for (var i = 0; i < exportConfig.length ; i++)
        {
          debug.log("DEBUG","exportConfig[i].SCRIPT_Q == [%s]\n",exportConfig[i].SCRIPT_Q)
          if (exportConfig[i].SCRIPT_Q == queueName)
          {
            docTypesUsed = exportConfig[i].TYPES_TO_INCLUDE;
            docTypesNotUsed = exportConfig[i].TYPES_TO_EXCLUDE;
            docTypeOrder = exportConfig[i].DOCTYPE_ORDER;
            singleDocType = exportConfig[i].IN_SINGLE_DOC;
            trigger = exportConfig[i].TRIGGER;
            emailFlag = exportConfig[i].EMAIL_OUTPUT;
            importFlag = exportConfig[i].IMPORT_OUTPUT;
            fileFormat = exportConfig[i].FILE_FORMAT;
            protectPdf = exportConfig[i].PROTECT_PDF;
            break;
          }
        }
      }//end checking the yaml for the list

      //get out if we can't find a list in the config
      if((!docTypesUsed || docTypesUsed == null) || (!docTypesNotUsed || docTypesNotUsed == null) || (!singleDocType || singleDocType == null) || (!trigger || trigger == null))
      {
        debug.log("ERROR","Invalid config for [%s] - docTypesUsed:[%s] singleDocType:[%s] trigger:[%s]\n",queueName, docTypesUsed, singleDocType, trigger);
        return false;
      }

      //check to make sure that the document is the correct type for generation
      var foundTrigger = false;
      for (var t = 0; t < trigger.length; t++)
      {
        if(doc.docTypeName == trigger[t])
        {
          debug.log("DEBUG","Found a valid trigger doctype [%s]\n", doc.docTypeName);
          foundTrigger = true;
        }
      }
      if(!foundTrigger)
      {
        debug.log("INFO","Current doc: [%s] [%s] is not the correct type for generation: [%s]\n", doc.id, doc.docTypeName, trigger);
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
      if(!processMatchingDocs(docList, saAppNo, fileFormat))
      {
        debug.log("ERROR","Unable to export all documents.\n");
        return false;
      }

      //get the documents that need to go at the front of the export
      var sortOrder = retrieveSortOrder(docTypeOrder);
      if(!sortOrder || sortOrder == null)
      {
        debug.log("INFO","No configured order for export.\n");
      }

      //get the list of exported files
      var exportList = getExportedFiles(doc, sortOrder, fileFormat, protectPdf, saAppNo);
      if(!exportList || exportList == null)
      {
        debug.log("ERROR","Unable to get list of exported files!\n");
        return false;
      }

      //if we are configured to import a document
      if(importFlag)
      {
        var revDoc = createNewDoc(doc, exportList, singleDocType, saAppNo);
        if(!revDoc || revDoc == null)
        {
          debug.log("ERROR","Unable to create the new document.\n");
          return false;
        }
      }//end if(importFlag)

      //clean up files
      cleanUpServer(exportList);
}//end documentExport_intool

//this function turns an array into a string with each element enclosed in single
//quotes and seperated by a comma
function prepareTypeLists(typeArray)
{
  debug.log("DEBUG","Inside prepareTypeLists.\n");
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

//function to get list of documents that can be added to the export
function findDocsToSend(emplid, drawer, appNo, usedType, skipType)
{
  debug.log("DEBUG","Inside findDocsToSend.\n");
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
        "AND INUSER.IN_INSTANCE.DELETION_STATUS <> '1' " +
        "AND INUSER.IN_DRAWER.DRAWER_NAME LIKE '" + drawer + "%' " +
        "AND ((INUSER.IN_PROP.PROP_NAME = 'SA Application Nbr' " +
        "AND INUSER.IN_INSTANCE_PROP.STRING_VAL = '" + appNo + "') " +
        "OR (INUSER.IN_PROP.PROP_NAME = 'Shared' " +
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
function processMatchingDocs(curObj, appNo, fileFormat)
{
  debug.log("DEBUG","Inside processMatchingDocs.\n");
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

    if(!exportDoc(tiffDoc, count, appNo, fileFormat))
    {
      debug.log("ERROR","Could not export [%s]\n",tiffDoc.id);
      return false;
    }
  }//end of the exporting
  debug.log("INFO","Finished exporting documents.\n");
  return true;
}//end processMatchingDocs

//function that extracts the tiffs from the doc and converts them to a single export
function exportDoc(doc, seq, appNo, fileFormat)
{
  debug.log("DEBUG","Inside exportDoc.\n");
  //make an output dir for the applicant's info because for some damn reason perceptive won't create a dir 2 deep
  //sorry, it's lexmark now
  var exportDir = "D:\\inserver6\\output\\"+doc.field1+"_"+appNo+"\\";//+doc.id+"\\";
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
    debug.log("DEBUG","safeSpace is [%s]\n", safeSpace);
    if(safeSpace.indexOf("&") > 0)
    {
      var pat = /(&)/g;
      safeSpace = safeSpace.replace(pat, '"&"');
      debug.log("DEBUG","safeSpace is now [%s]\n", safeSpace);
    }
    Clib.sprintf(cmd, '%s %s %s %s %s %s %s %s', POWERSHELL_ROOT, POWERSHELL_MERGE_TIFF, doc.field1, appNo, doc.id, safeSpace, seq, fileFormat);
    var rtn = exec(cmd, 0);
    //retrun the exit code
    return rtn;
  }
}//end exportDoc

//function to get the sort order for the documents defined in the sort order list
function retrieveSortOrder(listType)
{
  debug.log("DEBUG","Inside retrieveSortOrder.\n");
  var dtList = new INDocTypeList("",listType);
  dtList.getInfo();
  debug.log("DEBUG","Checking list: [%s]\n",dtList.name);
  var dtArray = dtList.members;
  var strTypeArray = [];
  debug.log("DEBUG","dtArray.length = [%s]\n", dtArray.length);
  for(var l = 0; l < dtArray.length; l++)
  {
      strTypeArray.push(dtArray[l].name);
  }
  for(var l = 0; l < strTypeArray.length; l++)
  {
      debug.log("DEBUG","strTypeArray = [%s]\n", strTypeArray[l]);
  }
  //return the sort order, if there is one
  return strTypeArray;
}//end retrieveSortOrder

//function that gets a list of the exported files and sorts them, if there is a sorting list
function getExportedFiles(docObj, docSortArr, fileFormat, protect, appNo)
{
  debug.log("DEBUG","Inside getExportedFiles.\n");
  var filePath = "D:\\inserver6\\output\\complete\\" + docObj.field1 + "_" + appNo + "\\";
  var newImports = SElib.directory(filePath + "*");
  debug.log("DEBUG","docSortArr.length = [%s]\n", docSortArr.length);
  //sort the files if there is a sort list
  if (docSortArr.length > 0)
  {
    newImports = orderExports(newImports, docSortArr);
  }
  //merge the files if there should be only one

    var pArray = [];
    for(var i = 0; i < newImports.length; i++)
    {
     debug.log("DEBUG","newImports[i].name is [%s]\n", newImports[i].name);
     if(newImports[i].name.indexOf("&") > 0)
     {
       var pat = /(&)/g;
       newImports[i].name = newImports[i].name.replace(pat, '"&"');
       debug.log("DEBUG","newImports[i].name is now [%s]\n", newImports[i].name);
     }
     pArray.push("'"+newImports[i].name+"'");
    }
    var merged = mergeDocs(pArray, docObj.field1, appNo, fileFormat, protect);
    newImports = SElib.directory(filePath + "*");
  
  //return the list of files
  return newImports;
}//end getExportedFiles

//function to sort the files according to an doctype list.
//if doctypes are not on the list, they are added at the end of the array FCFS
function orderExports(dirObj, order)
{
  debug.log("DEBUG","Inside orderExports.\n");
  var orderedArray = [];
  //add the docs on the list to the the array according to the order
  debug.log("DEBUG","order.length = [%s]\n", order.length);
  for(var p = 0; p < order.length; p++)
  {
    debug.log("DEBUG","dirObj.length = [%s]\n", dirObj.length);
    for (var o = 0; o < dirObj.length; o++)
    {      
      var parts = SElib.splitFilename(dirObj[o].name);
      var trimPt = parts.name.lastIndexOf("_");
      //debug.log("DEBUG","trimPt = [%s]\n", trimPt);
      //debug.log("DEBUG","parts.name.substring(0, trimPt) = [%s]\n", parts.name.substring(0, trimPt));
      //debug.log("DEBUG","order[p] = [%s]\n", order[p]);
      if(order[p].indexOf(parts.name.substring(0, trimPt)) == 0)
      {
        orderedArray.push(dirObj[o]);
      }
    }
  }//end of adding the ordered docs

  //remove the already-added docs from the remaining docs
  debug.log("DEBUG","orderedArray.length = [%s]\n", orderedArray.length);
  for (var q = 0; q < orderedArray.length; q++)
  {
    debug.log("DEBUG","orderedArray[q] = [%s]\n", orderedArray[q]);
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
  debug.log("DEBUG","Inside removeFromArray.\n");
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

//function to merge multiple files into one
function mergeDocs(docPath, docName, appNo, fileFormat, protect)
{
  debug.log("DEBUG","Inside mergeDocs.\n");
  var cmd = "";
  Clib.sprintf(cmd, '%s %s %s %s %s %s %s', POWERSHELL_ROOT, POWERSHELL_MERGE_DOC, docPath, docName, appNo, fileFormat, protect);
  var rtn = exec(cmd, 0);
  return rtn;
}//end mergeDocs

//function to create the document back in imageNow
function createNewDoc(docObj, expFiles, docObjType, appNo)
{
  debug.log("DEBUG","Inside createNewDoc.\n");
  //where the files are
  var filePath = "D:\\inserver6\\output\\complete\\" + docObj.field1 + "_" +appNo+ "\\";
  var newKeys = popKeys(docObj, docObjType);
  var props = docObj.getCustomProperties();
  var newDoc = new INDocument(newKeys);

  //make the document with the properties
  if(!newDoc.create(props))
  {
    debug.log("ERROR","Could not create a document to store the file!\n");
    return false;
  }

  debug.log("INFO","Created [%s]\n", newDoc);

  //for each file that's in the export dir
  for (var j = 0; j < expFiles.length; j++)
  {
    debug.log("INFO","Working file [%s] = [%s]\n",j, expFiles[j].name);
    var parts = SElib.splitFilename(expFiles[j].name);
    var workingDoc = parts.name + parts.ext;
    var attr = new Array;
    attr["phsob.file.type"] = parts.ext;
    attr["phsob.working.name"] = workingDoc;
    attr["phsob.source"]="GA_DocumentExport";
    //add the file to the document
    var logob = newDoc.storeObject(expFiles[j].name, attr);
    if (logob == null)
    {
        debug.log("ERROR","Could not import document:%s.\n", getErrMsg());
        return false;
    }
    else
    {
      debug.log("INFO","Successfully imported logobID: %s.\n", logob.id);
    }
  }//end of for each file

  //return the newly created doc object
  return newDoc;
}//end createNewDoc

//make the keys for the new doc with a new linkDate/Time
function popKeys(docObj, docObjType)
{
  debug.log("DEBUG","Inside popKeys.\n");
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
  debug.log("INFO","Looking for email address for [%s].\n", lastUser);
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
}//end getRouterEmail

//function to send the file via email
function sendFile(address, attachment, docObj)
{
  debug.log("DEBUG","Sending email to [%s]\n", address);
  var attch = [];
  for(var e = 0; e < attachment.length; e++)
  {
    attch.push("'" + attachment[e].name + "'");
  }
  var cmd = "";
  Clib.sprintf(cmd, '%s %s %s %s %s %s', POWERSHELL_ROOT, POWERSHELL_EMAIL_DOC, address, attch, docObj.field1, "'" + docObj.field2 + "'");
  var rtn = exec(cmd, 0);
  return rtn;
}//end sendFile

//function to clean up leftover files
function cleanUpServer(filesObj)
{
  debug.log("INFO","Cleaning up files after process.\n");
  var baseDir = "";
  for (var b = 0; b < filesObj.length; b ++)
  {
    var parts = SElib.splitFilename(filesObj[b].name);
    baseDir = parts.dir;
    //delete the file from the server
    Clib.remove(filesObj[b].name);
    debug.log("INFO","Removed [%s]\n",filesObj[b].name);
  }
  //remove the export directory
  Clib.rmdir(baseDir)
  debug.log("INFO","Removed [%s]\n",baseDir);
}//end cleanUpServer


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