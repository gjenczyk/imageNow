/*********************************************************************************************
	name:			FA_Restore_Keys.js
	Author:         Dave Snigier (UMass) & Gregg Jenczyk
	Created:		06/11/2015
	Last Updated:
	For Version:	6.7+
----------------------------------------------------------------------------------------------
	Summary:
		Runs as an inbound action on the preapp queue, only runs for items routed from specific queue(s)
		Will remap keys between index values and custom properties within a single document.
		Replace the index values with the temp values from the the CPs. Not the other way around.

	Mod Summary:

	Business Use:
		Resets the key structure of documents that were false positives in the preapp match section

	TODO:
		Currently only works with string type custom properties

*********************************************************************************************/

//********************* Include additional libraries *******************
#include "$IMAGENOWDIR6$\\script\\lib\\iScriptDebug.jsh"
#include "$IMAGENOWDIR6$\\script\\lib\\yaml_loader.jsh"
#include "$IMAGENOWDIR6$\\script\\lib\\PropertyManager.jsh"
#include "$IMAGENOWDIR6$\\script\\lib\\RouteItem.jsh"
#include "$IMAGENOWDIR6$\\script\\lib\\envVariable.jsh"

//*********************         Configuration        *******************

// Logging
// false - log to stdout(intool), wfUser(inscript), true - log to inserverXX/log/ directory
#define LOG_TO_FILE true

// constants seem to have variable scoping issues with complex scripts, so we're just going to use a global variable
// slower is better than broken after all
var DEBUG_LEVEL = 5;

// where to send alerts when critical errors are encountered
var EMAIL_ERRORS_NOTIFY = "UITS.DI.CORE@umassp.edu"//"UITS.DI.ADMIN.APP@umassp.edu";

// default error queue if one is not set in the config
var QUEUE_ERROR = "";


// for testing only:

// currentWfItem = new INWfItem('301YWD3_031X3ZVE6000080');
// currentWfItem.getInfo();
// currentWfQueue = new INWfQueue('301YWD3_031X3ZVE6000080');
// currentWfQueue.name = currentWfItem.queueName;


// JSLint configuration:
/*global
iScriptDebug, LOG_TO_FILE, DEBUG_LEVEL, CONFIG_VERIFIED, Buffer, Clib, printf,
INAutoForm, INBizList, INBizListItem, INClassProp,
INDocManager, INDocPriv, INDocType, INDocTypeList, INDocument, INDrawer, INExternMsg,
INFont, INGroup, INInstanceProp, INKeys, INLib, INLogicalObject, INMail, INOsm, INPriv,
INPrivEnum, INFolder, InProjManager, INProjectType, INProjTypeList, INProperty,
INRemoteService, INRetentionHold, INSequence, INSubObject, INSubobPriv, INSubobTemplate,
INTask, INTaskTemplate, INUser, INVersion, INView, INWfAdmin, INWfItem, INWfQueue,
INWfQueuePriv, INWorksheet, INWsDataDef, INWsPresentation, currentTask, currentWfItem,
currentWfQueue, currentDestinationWfQueue, _argv*/


// *********************       End  Configuration     *******************

// ********************* Initialize global variables ********************

var debug = "";
// INKeys object used to update the index keys
var STAGE_KEYS = new INKeys();
// array of ininstance props used in the update process
var STAGE_PROPS = [];
// config for the currentWfQueue
var CUR_CFG = false;


// ********************* Include additional libraries *******************

// ********************* Function definitions ***************************


// this will take care of staging valuesin prep for being set
// uses global arrays STAGE_KEYS for index values and STAGE_PROPS for the custom propertys respectively
// STAGE_PROPS should be constructed from INKeys and STAGE_KEYS from an array literal
// @param {string property|index} outputType cooresponds to what type of value is being mapped
// @param {string} name is either the name of the custom property or the name of the index key
// @param {string} value is what the custom property or index key should be set to
// @returns {Boolean} true if successful, false if errors were encountered (e.g. custom property doesn't exist in the system)
function mapValueToOutput(outputType, name, value) {
	debug.log('DEBUG', 'called: mapValueToOutput(%s, %s, %s)\n', outputType, name, value);
			if (value === null)
			{
				value = "";
			}
	switch (outputType.toUpperCase()) {
		case "PROPERTY":
			STAGE_PROPS.push({name:name, value:value});
			// STAGE_PROPS[name] = value;
			break;
		case "INDEX":
			STAGE_KEYS[name] = value;
			break;
		default:
			debug.log('CRITICAL', 'mapValueToOutput: Config is incorrect for remap values\n');
			return false;
	}
	return true;
}


// routes document to error queue, sends error email
function throwError(errorReason, sourceDocWfItem) {
	debug.log('DEBUG', 'Called: throwError(%s, %s)\n', errorReason, sourceDocWfItem);
	// route to specific error queue if it exists, otherwise route to the default
	var routeToQueue;
	var tmpItem;
	if (typeof CUR_CFG !== 'undefined' && CUR_CFG.errorQueue) {
		routeToQueue = CUR_CFG.errorQueue;
	} else {
		routeToQueue = QUEUE_ERROR;
	}

	// log out basic error (level defaults to ERROR)
	debug.log('ERROR', errorReason + '\n');
	debug.log('ERROR', 'Routing all docs to error queue\n');

	// route source document to error queue
	RouteItem(sourceDocWfItem, routeToQueue, errorReason, true);

	// send email with error message
	// EMAIL_ERRORS_NOTIFY = dsnigier@umassp.edu
	var wf = sourceDocWfItem;
	var body = "Restore Keys script failed with the following error message: \n" +
		errorReason + "\n" +
		"Workflow Item ID: " + wf.id + "\n" +
		"Document ID: " + wf.objectId + "\n" +
		WEBNOW_ENV_URL + "docid=" + wf.objectId + "&workflow=1";
	var subject = "[DI " + ENV_U3 + " Error] Restore Keys failure";
	var cmd = 'echo \"' + body + '\" | mailx -s \"' + subject + '\" ' + EMAIL_ERRORS_NOTIFY;
	debug.log('DEBUG', 'Email string:\n');
	debug.log('DEBUG', '%s\n', cmd);
	Clib.system(cmd);
}

/** ****************************************************************************
  *		Main body of script.
  *
  * @param {none} None
  * @returns {void} None
  *****************************************************************************/
function main() {
	try {
		var i = 0;
		var i2 = 0;

		var strH = "----------------------------------------------------------------------------------------------------\n";
		var strF = "____________________________________________________________________________________________________\n";
		debug = new iScriptDebug("USE SCRIPT FILE NAME", LOG_TO_FILE, DEBUG_LEVEL, undefined, {strHeader:strH, strFooter:strF});
		debug.showINowInfo("INFO");
		
		if (typeof currentWfItem === 'undefined') {
			debug.log("CRITICAL", "This script is not designed to run from intool.\n");  //intool
			return false;
		}

		// --- load report config files --- //
		loadYAMLConfig(imagenowDir6+"\\script\\config_scripts\\FA_Restore_Keys\\config\\");

		// find proper configuration for this queue
		for (var queueConfig in CFG.config) {
			debug.log("DEBUG","queueConfig = [%s]\n", queueConfig);
			if (CFG.config.hasOwnProperty(queueConfig) &&
				queueConfig == currentWfQueue.name) {
				// found match
				CUR_CFG = CFG.config[queueConfig];
			}
		}

		if (!CUR_CFG) {
			throwError('cannot find config for queue: ' + currentWfQueue.name, currentWfItem);
			return false;
		}

		currentWfItem.getInfo();


		// --- load document and determine if script should be run --- //
		var doc = new INDocument(currentWfItem.objectId);
		if (!doc || !doc.getInfo()) {
			throwError('Cannot create a new document object from the current workflow item', currentWfItem);
			return false;
		}

		// see if the script should continue running and restore the keys
		// only will occur if it was routed in from one of the previous queues in config

		var history = currentWfItem.getHistory();
		if (history === null) {
			throwError('Cannot retreive history for wfitem: ' + getErrMsg(), currentWfItem);
			return false;
		}

		// debug.log('DEBUG', 'History:\n');
		// debug.logObject('DEBUG', history, 10);

		var runRemap = false;
		for (i = history.length - 1; i >= 0; i--) {
/*			for (var key in history[i])
			{
				debug.log("DEBUG", "history[%s][%s] is [%s]\n", i, key, history[i][key]);
			}
*/			if (history[i].stateDetail === 2) {
				//debug.log("DEBUG","history[i].stateDetail = [%s]\n", history[i].stateDetail);
				var toCheck = CUR_CFG.previousQueues.length - 3;
				for (i2 = CUR_CFG.previousQueues.length - 1; i2 >= toCheck; i2--) {
					//debug.log("DEBUG","CUR_CFG.previousQueues[i2] = [%s] & history[i].queueName = [%s]\n", CUR_CFG.previousQueues[i2], history[i].queueName);
					if (CUR_CFG.previousQueues[i2] === history[i].queueName) {
						runRemap = true;
						debug.log('DEBUG', 'Document came from an approved queue - remapping will now commence.\n');
						break;
					}
					else
					{
						debug.log('DEBUG', 'Document come from a banned queue, remapping will not happen.\n');
						break;
					}
				}
				break;//if(runRemap = false)
			}
		}

		if (!runRemap) {
			debug.log('DEBUG', 'Item routed from a queue included in the configuration. Will not attempt remapping\n');
			return false;
		}

		// --- wipe out all index keys except those in the whitelist --- //

		for (i = CUR_CFG.saveIndexes.length - 1; i >= 0; i--) {
			debug.log('DEBUG', 'Save Index[i]: [%s] and doc[CUR_CFG.saveIndexes[i]: [%s] .\n', CUR_CFG.saveIndexes[i], doc[CUR_CFG.saveIndexes[i]] );
			mapValueToOutput("index", CUR_CFG.saveIndexes[i], doc[CUR_CFG.saveIndexes[i]]);
		}

		// keys are working properly
		debug.log('DEBUG', 'document object from wf\n');
		debug.logObject('DEBUG', doc, 1000);
		debug.log('DEBUG', 'Whitelist index keys:\n');
		debug.logObject('DEBUG', STAGE_KEYS, 1000);

		// --- wipe out all custom properties except for those on the whitelist --- //
		var originalProps = doc.getCustomProperties();
		
		var pm = new PropertyManager();
		var foundMatch = false;

		for (i2 = originalProps.length - 1; i2 >= 0; i2--) {
			foundMatch = false;
			for (i = CUR_CFG.saveProperties.length - 1; i >= 0; i--) {
				if (originalProps[i2].name === CUR_CFG.saveProperties[i]) {
					foundMatch = true;
				}
			}
			if (!foundMatch) {
				mapValueToOutput("property", originalProps[i2].name, "");
			}
		}

		debug.log('DEBUG', 'finished mapping props\n');
		debug.log('DEBUG', 'STAGE_PROPS\n');
		debug.logObject('DEBUG', STAGE_PROPS, 100);
		// --- Process remapped values --- //
		// remap values from one property/key to another
		debug.logObject('DEBUG', CUR_CFG.remapValues, 100);
		for (i = 0; i < CUR_CFG.remapValues.length; i++) {
			var value = false;
			var outputType = CUR_CFG.remapValues[i].destType;
			var outputName = CUR_CFG.remapValues[i].destName;
			var inputName = CUR_CFG.remapValues[i].sourceName;
			var inputType = CUR_CFG.remapValues[i].sourceType;
			switch (inputType.toUpperCase()) {
				case "PROPERTY":
					for (i2 = originalProps.length - 1; i2 >= 0; i2--) {
						if (originalProps[i2].name === inputName) {
							value = pm.get(doc, inputName);
							break;
						}
					}
					if (value === false || !mapValueToOutput(outputType, outputName, value)) {
						throwError('Cannot map value: ' + inputName, currentWfItem);
						return false;
					}
					break;
				case "INDEX":
					value = doc[inputName];
					if (value === false || !mapValueToOutput(outputType, outputName, value)) {
						throwError('Cannot map value: ' + inputName, currentWfItem);
						return false;
					}
					break;
				default:
					throwError('Config is incorrect for remap values: ' + inputType, currentWfItem);
					return false;
			}
		}

		for (i = 0; i < CUR_CFG.assignValues.length; i++) {
			var value = false;
			var outputType = CUR_CFG.assignValues[i].destType;
			var outputName = CUR_CFG.assignValues[i].destName;
			var value = CUR_CFG.assignValues[i].destValue;
			switch (outputType.toUpperCase()) {
				case "PROPERTY":
					if (value === false || !mapValueToOutput(outputType, outputName, value)) {
						throwError('Cannot map assign value: ' + outputName, currentWfItem);
						return false;
					}
					break;
				case "INDEX":
					value = doc[outputName];
					if (value === false || !mapValueToOutput(outputType, outputName, value)) {
						throwError('Cannot map assign value: ' + outputName, currentWfItem);
						return false;
					}
					break;
				default:
					throwError('Config is incorrect for remap values: ' + inputType, currentWfItem);
					return false;
			}
		}

		debug.log('DEBUG', 'Stage keys:\n');
		debug.logObject('DEBUG', STAGE_KEYS, 100);

		debug.log('DEBUG', 'stage props:\n');
		debug.logObject('DEBUG', STAGE_PROPS, 10);



		// --- persist the changes --- //

		if (!pm.set(doc, STAGE_PROPS)) {
			throwError('Unable to set custom properties: ' + getErrMsg(), currentWfItem);
			return false;
		}

		//if (INDocManager.moveDocument(doc.id, STAGE_KEYS, "APPEND")) {
		if (doc.setProperties(STAGE_KEYS)) {
				debug.log('DEBUG', 'keys successfully restored\n');
		} else {
			throwError('Unable to relink document: ' + getErrMsg(), currentWfItem);
			return false;
		}

	} catch (e) {
		if (!debug) {
			printf("\n\nFATAL iSCRIPT ERROR: %s\n\n", e.toString());
		}
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
	} finally {
		debug.finish();
		return;
	}
}
