/********************************************************************************
        Name:          GA_DocumentExport
        Author:        Gregg Jenczyk
        Created:        04/03/201
        Last Updated:   
        For Version:    
---------------------------------------------------------------------------------
        Summary:
               This script will email a selection of docs in the system to a user
               in the form of a pdf.  Requires tiif lib for gnuwin32.
               
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
var POWERSHELL_MERGE = imagenowDir6+"\\script\\PowerShell\\mergeTiffs.ps1"

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
      var wfItem = new INWfItem("301YY4P_04VZ7RN4N0142Y5");//currentWfItem.id);//
      if(!wfItem.id || !wfItem.getInfo())
      {
        debug.log("CRITICAL", " Couldn't get info for wfItem: %s\n", getErrMsg());
        return false;
      }

      //get queue name
      var queueName = wfItem.queueName;

      //store the results from the yaml here
      var docTypeList = "";
      var pdfDocType = "";
      var pdfDocQCfg = imagenowDir6 + "\\script\\";

      //collect information about who routed the document into the queue
      var router = new INUser(wfItem.queueStartUserName);
      if(!router.getInfo())
      {
        debug.log("ERROR","Could not get user information for [%s] - [%s]\n",wfItem.queueStartUserName,getErrMsg());
        return false;
      }
      //make sure user has an email address on file
/*      if(!router.email || router.email == null)
      {
        debug.log("ERROR","No email address found for [%s - %s %s]\n",router.id, router.lastName, router.firstName);
        return false;
      }
*/
      //get information about inbound doc
      var doc = new INDocument(wfItem.objectId);//"301YY4P_04VZ7RN4N0142XW");  
      if(!doc.getInfo())
      {
          debug.log("ERROR","Couldn't get doc info: [%s]\n",getErrMsg()); 
          return false;
      }
      //get doc keys
      var drawer = doc.drawer;
      var emplid = doc.field1;
      var name = doc.field2;
      var term = doc.field3;
      var plan = doc.field4;

      //identify the correct doctype list to use for exporting in the yaml
      loadYAMLConfig(imagenowDir6+"\\script\\config_scripts\\GA_DocumentExport\\");
      for (var office in CFG.GA_DocumentExport)
      {
        var exportConfig = CFG.GA_DocumentExport[office].EXPORT_CONFIG;
        for (var i = 0; i < exportConfig.length ; i++)
        {
          if (exportConfig[i].SCRIPT_Q == queueName)
          {
            docTypeList = exportConfig[i].DOCTYPE_LIST;
            pdfDocType = exportConfig[i].IN_PDF_DOC;
            pdfDocQCfg += exportConfig[i].CREATE_IN;
            break;
          }
        }
      }//end checking the yaml for the list

      //get out if we can't find a list in the config
      if((!docTypeList || docTypeList == null) || (!pdfDocType || pdfDocType == null) || (!pdfDocQCfg || pdfDocQCfg == null))
      {
        debug.log("ERROR","Invalid config for [%s] - docTypeList:[%s] pdfDocType:[%s] pdfDocQCfg:[%s]\n",queueName,docTypeList,pdfDocType,pdfDocQCfg);
        return false;
      }

      //get a list of documents to send
      var docList = findDocsToSend(emplid, drawer, docTypeList);

      if (!docList || docList == null)
      {
        debug.log("ERROR","Couln't find exportable documents for [%s] in [%s] and [%s]\n",emplid, drawer, docTypeList);
        return false;
      }

      //prepare the documents for emailing
/*      while(docList.next())
      {
        var tiffDoc = new INDocument(docList[0]);
        if(!tiffDoc.getInfo())
        {
            debug.log("ERROR","Couldn't get doc info: [%s]\n",getErrMsg()); 
            //what to do here? return false;
        }

        if(!exportDoc(tiffDoc))
        {
          debug.log("ERROR","Could not export [%s]\n",tiffDoc.id);
          //what to do here?
        }
      }//end of the exporting
*/
      var revDoc = createNewDoc(drawer,emplid,name,term,plan,pdfDocType);


      //prpbably should have some kind of check here...
      addDocToQueue(revDoc, pdfDocQCfg);
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
function findDocsToSend(emplid, drawer, list)
{
  sql = "SELECT INUSER.IN_DOC.DOC_ID, " +
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
        "WHERE INUSER.IN_DOC.FOLDER    = '" + emplid + "' " +
        "AND INUSER.IN_DRAWER.DRAWER_NAME like '" + drawer + "%' " +
        "AND LIST_NAME = '" + list + "';"

        var returnVal;
        var cur = getHostDBLookupInfo_cur(sql,returnVal);

        if(!cur || cur == null)
        {
          debug.log("WARNING","no results returned for query.\n");
          return false;
        }

        return cur; 
}

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
}

function exportDoc(doc)
{
  //make an output dir for the applicant's info because for some damn reason perceptive won't create a dir 2 deep
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
    Clib.sprintf(cmd, '%s %s %s %s', POWERSHELL_ROOT, POWERSHELL_MERGE, doc.field1, doc.id);
    var rtn = exec(cmd, 0);
    return rtn;
  }
}//end export doc

function createNewDoc(drawer,field1,field2,field3,field4,type)
{
  //get unique dateTime
  var linkDate_obj = new GetUniqueF5DateTime(true);
  var linkDate = linkDate_obj.tellTheTime();

  var keys = new INKeys(drawer,field1,field2,field3,field4,linkDate,type);
  debug.log("INFO","Creating new document: [%s]\n",keys.toString());
  var pdfDoc = new INDocument(keys);
return pdfDoc;
  //or use doc Id to create a document instance.
  var filePath = "D:\\inserver6\\output\\complete\\" + field1 + "\\";
  var pdfImports = SElib.directory(filePath + "*");
  for (var j = 0; j < pdfImports.length; j++)
  {
    debug.log("INFO","pdfImports[%s] = [%s]\n",j, pdfImports[j].name);
    var parts = SElib.splitFilename(pdfImports[j].name);
    var workingPdf = parts.name + parts.ext;
    var attr = new Array;
    attr["phsob.file.type"] = parts.ext;
    attr["phsob.working.name"] = workingPdf;
    attr["phsob.source"]="GA_DocumentExport";
    var logob = pdfDoc.storeObject(pdfImports[j].name, attr);
    if (logob == null)
    {
        debug.log("ERROR","Could not import document:%s.\n", getErrMsg());
        return false;
    }
    else
    {
      debug.log("INFO","Successfully imported logobID: %s.\n", logob.id);
      if(!Clib.remove(pdfImports[j].name))
      {
        debug.log("ERROR","Could not delete: [%s] - [%s]\n",pdfImports[j].name, getErrMsg());
      }
      else
      {
        debug.log("INFO","Successfully deleted [%s]\n",pdfImports[j].name);
      }
    }
  }//end of for each pdf
  if(!Clib.rmdir(filePath))
  {
    debug.log("ERROR","Could not remove [%s] - [%s]\n",filePath, getErrMsg());
  }
  else
  {
    debug.log("INFO","Removed [%s]\n",filePath);
  }

  return pdfDoc;

}//end create new doc

function addDocToQueue(doc, routeCfg)
{
  doc.getInfo();
  var docPlan = doc.field4;
  var subPlan = GetProp(doc,"Sub-Plan",false);

  var columnConfig =
  [
    {name:'plan'},
    {name:'subplan'},
    {name:'subqueue'}
  ];
  var objCSV = new csvObject(routeCfg, columnConfig, {intHeaderLen:1, delim:',', innerDelim:''});
  if(!objCSV.openFile('r'))
  {
    debug.log("ERROR", "loadPlanSubPlanQueues: unable to open file at [%s]\n", routeCfg);
    return false;
  }

  var rtnObject = {};
  var row =0;
  while(true)
  {
    row++;
    var line = objCSV.getNextRowObject();
    if(line === false)
    {
      debug.log("CRITICAL", "loadPlanSubPlanQueues: unable to read data on row [%s]\n", row);
      break;
    }
    else if(line === null)
    {
      //debug.log("DEBUG", "loadPlanSubPlanQueues: end of file\n");
      break;
    }

    //if (docPlan === line['plan'])
    //{
      printf("%s %s\n",line['subplan'],line['plan']);
    //}

  }
  objCSV.closeFile();
}

//