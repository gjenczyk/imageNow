' *********************************************************** @fileoverview<pre>
'	Name:           INMAC_MailAgentConvert.vbs
'	Author:         Perceptive Software, Inc
'	Created:        08/28/2008-LMS
'	Last Updated:   03/22/2011-CAL
'	Script Version: $Id: INMAC_MailAgentConvert.vbs 28453 2013-11-13 21:04:37Z clingor $
'	For Version:    [ any ]
' ------------------------------------------------------------------------------
'	Summary:
'		This is responsible for driving the Mail Agent Convert to Tiff process
'
'	Mod Summary:
'		03/16/2009-RTH: Added IrfanView as an option for converting files.
'		03/26/2009-ADD: Adding ini settings: LogDir, StatsDir and DeleteLogFilesAfterXDays
'		04/30/2009-CAL: Added global variable for array size of "filesToMove" when archiving files.
'		05/14/2009-HAA: Added logic to execute PDFToText when setup in PDF section of .ini
'		06/03/2009-CAL: Changed page and doc timeouts from CInt to CDbl to allow longer timeout values.
'		07/13/2009-CAD: Updated to correct path issue with MSPF (and potentially other command-line execution)
'		08/06/2009-CAL: Fixed bug on process username check.  Now checks for multiple processes on exact username and domain.
'		09/09/2009-JWM: Fixed bug where text directory not being created when using an application for conversion
'		02/08/2010-CAL: Added DisplayAlerts=False before closing Excel workbooks to avoid popup to save on close.
'		05/25/2010-CAD: Use literal date value instead of string in CDate to make this locale-independant
'		06/03/2010-CAL: Updated to replace case insensitive ".pdf" with ".txt".
'		06/25/2010-BSW: Added new conversion method for Word documents for command-line printing instead of using COM.
'		06/30/2010-JWM: Updated whitespace formatting
'		                Added TextExtractionMethod=EXCEL which uses COM to save the worksheet content as tab delimited text
'		07/08/2010-BSW: Added logic to automatically set the default printer to the printer defined in the INMAC.ini file.
'						Set the INMAC.ini setting 'ForceDefault' to true/false to toggle forcefully setting default printer.
'		03/09/2011-CAL: Updated to automatically create log and stats directories if they don't already exist.
'		03/10/2011-CAL: Check file size before converting and consider a 0 KB file size to be a conversion error.
'		03/22/2011-CAL: Updates to allow IE to quit properly via COM.
'						Updates to allow Excel to print every non-blank worksheet.
'						Updates to allow Excel to extract text from every non-blank worksheet.
'
'	Business Use:
'		INMAC - ImageNow Mail Agent Convert process
' </pre>***********************************************************************/

' **************************** Global Variables ********************************
Dim gObjDebugFile:Set gObjDebugFile         = Nothing      'File we write to
Dim gStrDebugFilePath:gStrDebugFilePath     = ""           'Path to file we write to
Dim gStrDebugHeader:gStrDebugHeader         = ""&vbCrLf    'Header string
Dim gStrDebugFooter:gStrDebugFooter         = ""&vbCrLf    'Footer string
Dim gBlnDebugHasWritten:gBlnDebugHasWritten = False        'Has output been written
Dim gBlnDebugIsOpen:gBlnDebugIsOpen         = False        'Is output file open
Dim gIntDebugStartTime:gIntDebugStartTime   = GetEpochMS() 'MS since 1/1/1970 till script started
Dim gStrDebugScriptName:gStrDebugScriptName = ""           'Base name of script
Dim gStrDebugLineHeader:gStrDebugLineHeader = ""           '<scriptname>:<offset>
Dim gIntDebugLevel
Dim OLECMDID_PRINT:OLECMDID_PRINT           = 6
Dim OLECMDEXECOPT_DONTPROMPTUSER:OLECMDEXECOPT_DONTPROMPTUSER = 2
Dim gLastPrintJobId:gLastPrintJobId         = 0

' Array size of nativePages and filesToMove
Dim gMaxArraySize:gMaxArraySize             = 50000

Dim gFsObj:Set gFsObj                       = CreateObject("Scripting.FileSystemObject")
Dim gLastErrorMsg: gLastErrorMsg            = ""
Dim gMaxDirToProcess
Dim gMyPid
Dim gAdoConnection
Dim gPreProcess
Dim gNetworkUser 'User running the scheduled task

Dim INMAC_PRINTER_STATUS_INI
Dim INMAC_INI
Dim STATS_INI
Dim IMAGENOW_PRINTER_NAME
Dim FORCE_DEFAULT
Dim INMAC_INSTANCE
Dim INPRINTER_PAGE_TIMEOUT_SECS
Dim INPRINTER_JOB_TIMEOUT_SECS
Dim MAGIC_NUM_SEC:MAGIC_NUM_SEC = "INMAC Magic Number Extraction"

Dim INMAC_USE_EXTERN_MSG
Dim INMAC_USE_ORACLE_DB
Dim INMAC_INPUT_BASE_DIR
Dim INMAC_OUTPUT_BASE_DIR
Dim INMAC_COLOR_REDUCTION
Dim INMAC_LOG_BASE_DIR
Dim INMAC_STATS_BASE_DIR
Dim DELETE_LOG_FILES_AFTER_X_DAYS
Dim gStrLogDir
Dim AppInitialized
Dim AppFinish
Dim AllProcessNames 'Read all CommandExe processes from ini and terminate them on startup
Dim gStats
Dim gTempDirBase

' ***************************** Body of Program ********************************
On Error Resume Next
Call Init()

DebugLog "INFO", "Started\n"
LogPrinterAndUserInfo()

If IsPrinterReady() = false Then
	' DebugLog "CRITICAL", "\n\n\nProblem with ImageNow Printer: [" & IMAGENOW_PRINTER_NAME & "]  Abort\n" & _
	' "ImageNow Printer must be installed, set as the default, no spooling, and ImageNow Printer\n" & _
	' "Mail Agent Convert must have exclusive access to the ImageNow Printer\n\n"
	SendAdminEmail("The ImageNow Virtual printer was not ready")
	Finish
End If

' find files that are ready for conversion
Set inFolders = gFsObj.GetFolder(INMAC_INPUT_BASE_DIR)

For each inFolder in inFolders.subFolders
	DebugLog "DEBUG", "Checking if ready for conversion: " & inFolder.Name & "\n"
	
	' INMAC_Submit.js iScript will create an file called EXPORT_COMPLETE when all pages
	' have been copied to the filesystem
	Err.Clear
	Set isDone = gFsObj.GetFile(infolder.Path & "\\EXPORT_COMPLETE")
	
	If (Err.number <> 0) Then
		Call gStatsInc("Total documents not fully exported yet")
		DebugLog "DEBUG", "Files not ready yet: " & inFolder.Name & "\n"
		Err.Clear
	Else
		Call gStatsInc(" Total incoming documents started")
		ConvertDocument(inFolder)
		
		Err.Clear
		
		'Move Original File to out folder
		For Each file in inFolder.Files
			If file.Name <> "EXPORT_COMPLETE" Then
				temp = INMAC_OUTPUT_BASE_DIR & inFolder.Name & "\" & SplitFilename(file.Path)(2) & "_ORIG" & SplitFilename(file.Path)(3)
				Call DebugLog("DEBUG", "Moving: [" & file.Path & "] to [" & temp & "]\n")
				rtn = gFsObj.MoveFile( file.Path, temp)
				If Err.number <> 0 Then
					DebugLog "WARNING", "Could not move native file [" & Err.Description & "]: " & file.Path & "\n"
					Err.Clear
				End If
			End If
		Next
		
		If gFsObj.DeleteFolder(inFolder.Path, true) = true Then
			Call DebugLog("ERROR", "Last Err.Description: " & Err.Description & "\n")
			Call DebugLog("ERROR", "Could not remove directory: " & inFolder.Path & "\n")
		End IF
		gMaxDirToProcess = gMaxDirToProcess - 1
		If gMaxDirToProcess < 1 Then
			Call DebugLog("INFO", "Finished due to MaxDirsPerRun\n")
			Call Finish()
		End If
	End If
Next
Finish
' *************************** End Body of Program ******************************

' ******************************************************************************
Function Init()
	On Error Resume Next
	
	Dim strStatsDir
	arrSplitFileName = SplitFileName(WScript.ScriptFullName)
	If IsNull(arrSplitFileName) Then
		MsgBox "Unable to split filename: ["&WScript.ScriptFullName&"]","Error - ["&WScript.ScriptFullName&"]"
		WScript.Quit(1)
	End If
	
	INMAC_INI = arrSplitFileName(1) & "INMAC.ini"
	
	'backup INMAC.ini to INMAC.ini.bak if it doesn't already exist
	INMAC_BACKUP_INI = arrSplitFileName(1) & "INMAC.ini.bak"
	Dim inmacIniFile:Set inmacIniFile = gFsObj.GetFile(INMAC_INI)
	If (Not(gFsObj.FileExists(INMAC_BACKUP_INI)) OR inmacIniFile.Size > 2048) Then
		'if backup doesn't exist or INMAC.ini file size is greater than 2KB, copy INMAC_INI to INMAC_BACKUP_INI
		gFsObj.CopyFile INMAC_INI, INMAC_BACKUP_INI, True
	End If
	
	INMAC_STATS_BASE_DIR = GetINIString("INMAC", "StatsDir", arrSplitFileName(1)&"stats\", INMAC_INI)
	strStatsDir = INMAC_STATS_BASE_DIR
	
	'automatically create stats directory if necessary
	If gFsObj.FolderExists(strStatsDir) = False Then
		gFsObj.CreateFolder(strStatsDir)
	End If
	
	dailyStatsFile = GetINIString("INMAC", "DailyStatsFile", "INMAC_Stats", INMAC_INI)
	STATS_INI = strStatsDir & dailyStatsFile & "_" & GetDateYYYYMMDD & ".ini"
	
	INMAC_PRINTER_STATUS_INI = arrSplitFileName(1) & "INMAC_Status.ini"
	IMAGENOW_PRINTER_NAME = GetINIString("INMAC", "ImageNowPrinterName", "", INMAC_INI)
	FORCE_DEFAULT = GetINIString("INMAC", "ForceDefault", "", INMAC_INI)
	INMAC_INSTANCE = GetINIString("INMAC", "InmacInstance", "", INMAC_INI)
	
	INMAC_INPUT_BASE_DIR = GetINIString("INMAC", "InputDir", "", INMAC_INI)
	INMAC_OUTPUT_BASE_DIR = GetINIString("INMAC", "OutputDir", "", INMAC_INI)
	INMAC_LOG_BASE_DIR = GetINIString("INMAC", "LogDir", "", INMAC_INI)
	INMAC_USE_EXTERN_MSG = GetINIString("INMAC External Messaging", "Enable", "", INMAC_INI)
	INMAC_USE_ORACLE_DB = GetINIString("INMAC External Messaging", "Oracle", "", INMAC_INI)
	INMAC_COLOR_REDUCTION = GetINIString("Compression", "Color reduction", "", INMAC_INI)
	
	INPRINTER_PAGE_TIMEOUT_SECS = CDbl(GetINIString("INMAC", "PageTimeoutSecs", "15", INMAC_INI))
	INPRINTER_JOB_TIMEOUT_SECS = CDbl(GetINIString("INMAC", "JobTimeoutSecs", "60", INMAC_INI))
	DELETE_LOG_FILES_AFTER_X_DAYS = CInt(GetINIString("INMAC", "DeleteLogFilesAfterXDays", "60", INMAC_INI))
	
	gPreProcess = CBool(GetINIString("INMAC Magic Number Extraction", "EnableMagicNumber", "false", INMAC_INI) or GetINIString("INMAC", "EnableZipExtraction", "false", INMAC_INI))
	
	Set AppInitialized = CreateObject("Scripting.Dictionary")
	Set AppFinish = CreateObject("Scripting.Dictionary")
	Set AllProcessNames = CreateObject("Scripting.Dictionary")
	Set gStats = CreateObject("Scripting.Dictionary")
	AllProcessNames.CompareMode = 1 'case insensitive matching
	gMaxDirToProcess = CLng(GetINIString("INMAC", "MaxDirsPerRun", "100", INMAC_INI))
	gIntDebugLevel = CInt(GetINIString("INMAC","DebugLevel","3",INMAC_INI))
	
	' Set initial script timeout
	WScript.Timeout = 2 * INPRINTER_JOB_TIMEOUT_SECS
	
	DebugNew()
	LogFilesCleanUp()
	
	' setup temp dir
	Err.Clear
	gTempDirBase = GetINIString("INMAC", "TempDir", arrSplitFileName(1) & "temp\", INMAC_INI)
	gFsObj.GetFolder(gTempDirBase)
	If Err.Number <> 0 Then
		' Dir doesn't exist, try to create
		Call DebugLog("INFO", "Creating Temp Dir: " & gTempDirBase & "\n")
		Err.Clear
		gFsObj.CreateFolder(gTempDirBase)
		If Err.Number <> 0 Then
			Call DebugLog("ERROR", "Could not create temp dir\n")
			Finish
		End If
	End If
	
	numWscripts=0
	Call DebugLog("DEBUG", "numWscripts: " & numWscripts & "\n")
	' Find own PID
	Set wshNetwork = CreateObject("WScript.Network")
	gNetworkUser = wshNetwork.Username
	strDomain = wshNetwork.UserDomain
	
	strComputer = "."
	Set colProcesses = GetObject("winmgmts:" & _
		"{impersonationLevel=impersonate}!\\" & strComputer & _
		"\root\cimv2").ExecQuery("Select * from Win32_Process")
		
	For Each objProcess in colProcesses
		If objProcess.Name = "wscript.exe" OR objProcess.Name = "cscript.exe" Then
			Call DebugLog("INFO", "Found wscript PID " & objProcess.ProcessId & " belonging to " & objProcess.Name & "\n")
			
			strNameOfUser = ""
			strDomainOfUser = ""
			Return = objProcess.GetOwner(strNameOfUser, strDomainOfUser)
			Call DebugLog("DEBUG", "-- User: [" & gNetworkUser & "], Domain: [" & strDomain & "], Process User: [" & strNameOfUser & "], Process Domain: [" & strDomainOfUser & "]\n")
			
			If Return <> 0 Then
				DebugLog "ERROR", "Could not get owner info for process " & _
					objProcess.Name & VBNewLine _
					& "Error = " & Return & "\n"
			Else
				'If InStr(UCase(strUser), UCase(strNameOfUser)) Then
				If (StrComp(gNetworkUser,strNameOfUser) = 0 And StrComp(strDomain,strDomainOfUser) = 0) Then
					numWscripts=numWscripts+1
				End If
			End If
		End If
	Next
	
	If numWscripts <> 1 Then
		Call DebugLog("CRITICAL", numWscripts & " instances of wscript found!  Configuration problem\n")
		SendAdminEmail("More than one instance of wscript running")
		Finish
	End If
	
	' Setup INMAC Stats sections
	lastRunDate = GetINIString("INMAC Stats", "lastRunDate", "", STATS_INI)
	If lastRunDate <> GetDateYYYYMMDD() Then
		Call ClearINICounters("INMAC Stats", STATS_INI)
		Call WriteINIString("INMAC Stats", "lastRunDate", GetDateYYYYMMDD(), STATS_INI)
	End If
	Call WriteINIString("INMAC Stats", "lastRunTime", GetTimeStamp(), STATS_INI)
	
	'1. Set WScript.Timeout
	'2. Kill existing processes
	'KillAllProc( GetINIString(configSection, "CommandExe", "", INMAC_INI) )
	
	' Find and terminate existing INMAC processes for this instance
	GetAllProcs(INMAC_INI)
	TerminateProcs()
	
	
	'3. Reset WScript.Timeout after every successful status.ini read
	
End Function

' ******************************************************************************
' Read all CommandExe processes from configuration ini
Function GetAllProcs(FileName)
	On Error Resume Next
	Dim INIContents, PosSection, PosEndSection, sContents, key
	
	'Get contents of the INI file As a string
	INIContents = GetFile(FileName)
	
	' Read all CommandExe processes from configuration ini
	PosPrevious = 1
	PosSection = 1
	Do While PosSection > 0
		PosSection = InStr(PosPrevious, INIContents, vbCrLf & "CommandExe=", vbTextCompare)
		If PosSection = 0 Then Exit Do
		PosEndSection = InStr(PosSection, INIContents, vbCrLf & "[")
		sContents = Mid(INIContents, PosSection, PosEndSection - PosSection)
		key = SeparateField(sContents, vbCrLf & "CommandExe=", vbCrLf)
		If Not(isEmpty(key)) AND Not(AllProcessNames.Exists(key)) Then
			AllProcessNames.Add key, True
		End If
		PosPrevious = PosEndSection
	Loop
End Function

' ******************************************************************************
' Terminate all running processes matching process names found in AllProcessNames
Function TerminateProcs()
	On Error Resume Next
	DebugLog "INFO", "TerminateProcs: Killing all INMAC processes running by user " & gNetworkUser & " (if still running)\n"
	
	strComputer = "."
	Set colProcesses = GetObject("winmgmts:" & _
		"{impersonationLevel=impersonate}!\\" & strComputer & _
		"\root\cimv2").ExecQuery("Select * from Win32_Process")
	
	For Each objProcess in colProcesses
		If AllProcessNames.Exists(Trim(UCase(objProcess.Name))) Then
			Call DebugLog("DEBUG", "TerminateProcs: Found existing process " & objProcess.Name & " " & objProcess.ProcessId & "\n")
			Return = objProcess.GetOwner(strNameOfUser)
			If Return <> 0 Then
				DebugLog "ERROR", "TerminateProcs: Could not get owner info for process " & objProcess.Name & VBNewLine & "Error = " & Return & "\n"
			Else
				If InStr(UCase(gNetworkUser), UCase(strNameOfUser)) Then
					rtn = objProcess.Terminate(0)
					If rtn <> 0 Then
						Call DebugLog("WARNING", "TerminateProcs: Unable to terminate process via win32_process, using taskkill\n")
						ExecNoWait( GetINIString("INMAC", "KillCommand", "TASKKILL /F /PID ", INMAC_INI) & CStr(objProcess.ProcessId))
					End If
					DebugLog "NOTIFY", "TerminateProcs: Killed process " & objProcess.ProcessID & " " & objProcess.Name & "\n"
				Else
					DebugLog "INFO", "TerminateProcs: NOT killing Process " & objProcess.Name & " It is owned by " & "\" & strNameOfUser & ".\n"
				End If
			End If
		End If
	Next
End Function


' ******************************************************************************
Function Finish()
	On Error Resume Next
	
	Call LogStats()
	
	' Cleanup any apps if necessary
	For Each key in AppFinish.keys
		Call DebugLog("INFO", "AppFinish: " & key & "\n")
		KillAllProc(AppFinish.Item(key))
	Next
	
	' Cleanup temp directory
	Err.Clear
	Set f = gFsObj.GetFolder(gTempDirBase)
	Call DebugLog("DEBUG", "Removing temp folders\n")
	If Err.Number <> 0 Then
		Call DebugLog("ERROR", "Could not get temp folder\n")
	Else
		Set fc = f.SubFolders
		Err.Clear
		For Each fldr in fc
			Call DebugLog("DEBUG", "Removing " & fldr.Path & "\n")
			Call gFsObj.DeleteFolder(fldr.Path, true)
			If Err.Number <> 0 Then
				Call DebugLog("ERROR", "Could not remove dir [" & Err.Description & "] " & fldr.Path & "\n")
				Err.Clear
			End If
		Next
	End If
	
	Call DebugFinish()
	WScript.Quit
End Function

' ******************************************************************************
Function LogFilesCleanUp()
	Set inFolders = gFsObj.GetFolder(gStrLogDir)
	
	For Each file in inFolders.Files
		If DateDiff("d", file.DateCreated, Now) > DELETE_LOG_FILES_AFTER_X_DAYS And Right(file.Name, 4) = ".log" And Left(file.Name, 6) = "INMAC_" Then
			Call DebugLog("INFO", "LogFilesCleanUp: Deleting log file " & file.Name & "\n")
			gFsObj.DeleteFile(file)
		End If
	Next
End Function

' ******************************************************************************
Function SendAdminEmail(msg)
	On Error Resume Next
	
	' Debug may not be instantiated at the point as an FYI
	
	If GetINIString("INMAC Alerts", "SendAlertEmail", "", INMAC_INI) <> "true" Then
		Call DebugLog("WARNING", "Email Alerts not enabled.\n")
		Exit Function
	End IF
	
	' check to see if we've sent an email recently
	lastEmail = CDbl(GetINIString("INMAC Alerts", "LastEmailSentTime", "0", INMAC_INI))
	If (GetEpochMS() < ((15*60000) + lastEmail)) Then
		Call DebugLog("WARNING", "Suppressing email due to recent email\n")
		Exit Function
	End If
	Call WriteINIString("INMAC Alerts", "LastEmailSentTime", CStr(GetEpochMS()), INMAC_INI)
	
	Set mailObj = CreateObject("Persits.MailSender")
	If Err.Number <> 0 Then
		Call DebugLog("CRITICAL", "Could not send email (is AspEmail registered?): " & Err.Description & "\n")
		Exit Function
	End IF
	
	' OLE: http://www.aspemail.com/
	' mailObj.Username =
	' mailObj.Password =
	
	mailObj.Host = GetINIString("INMAC Alerts", "MailServer", "", INMAC_INI)
	mailObj.From = GetINIString("INMAC Alerts", "MailFrom", "", INMAC_INI)
	mailObj.Subject = GetINIString("INMAC Alerts", "MailSubject", "ImageNow INMAC - Critical Conversion Error", INMAC_INI) & " [" & IMAGENOW_PRINTER_NAME & "]"
	mailObj.IsHTML = true
	mailObj.Body = msg
	mailObj.AddAddress( GetINIString("INMAC Alerts", "MailTo", "", INMAC_INI) )
	'mailObj.AddCC()
	'mailObj.Priority = ""
	
	mailObj.Send()
	If Err.Number <> 0 Then
		Call DebugLog("CRITICAL", "Could not send email: " & Err.Description & "\n")
	End IF
	
	Call DebugLog("INFO", "Email Sent\n")
	mailObj = nothing
End Function

' ******************************************************************************
Function ConvertDocument(folder)
	On Error Resume Next
	Dim outFolderStr
	Dim nativePages()
	ReDim nativePages(gMaxArraySize)
	Dim numNativePages
	Dim regFileFormat
	Dim thisFileType
	Dim Matches
	Dim convertResult
	ConvertDocument = false
	
	If gPreProcess Then
		PreProcess(folder)
	End If
	
	gLastErrorMsg = ""
	Set regFileFormat = New RegExp
	regFileFormat.Pattern = "^(.*)-(.*)-([^\.]*)\.?(.*)$"
	
	numNativePages = 0
	
	DebugLog "DEBUG", "Enter ConvertDocument: " & folder.Path & "\n"
	
	' Create output directory
	outFolderStr = INMAC_OUTPUT_BASE_DIR & folder.Name
	
	' see if output directory exists
	Err.Clear
	Set temp = gFsObj.GetFolder( outFolderStr )
	If Err.Number = 0 Then
		Call DebugLog("WARNING", "Removing existing job directory! " & outFolderStr & "\n")
		Call gFsObj.DeleteFolder(outFolderStr, true)
		If Err.Number <> 0 Then
			Call DebugLog("ERROR", "Could not delete existing job directory: " & Err.Description & "\n")
		End If
	End If
	Err.Clear
	DebugLog "DEBUG", "Creating output folder: " & outFolderStr & "\n"
	
	gFsObj.CreateFolder outFolderStr
	If (Err.number <> 0) Then
		DebugLog "ERROR", "ConvertDocument: Could not create output dir: " & Err.Description & "\n"
		' FIXME Exit Function
	End If
	
	For Each file in folder.Files
		If file.Name <> "EXPORT_COMPLETE" Then
			If regFileFormat.Test(file.Name) <> true Then
				Call gStatsInc("ERROR Unknown input files")
				DebugLog "ERROR", "ConvertDocument: Ignoring unknown file: " & file.Name & "\n"
			Else
				DebugLog "DEBUG", file.Name & " size: " & file.size & "\n"
				If file.size = 0 Then
					Call DebugLog("ERROR", "File [" & file.Name & "] is empty, cannot convert document\n")
					ConvertDocument = false
					Set Matches = regFileFormat.Execute(file.Name)
					thisFileType = Matches.Item(0).SubMatches.Item(3)
					Call NotifyFailed(outFolderStr, CStr(numNativePages+1), thisFileType)
					Exit Function
				End If
				call gStatsAdd("Total Incoming KB", Int( file.size / 1000 ))
				call gStatsInc("Total Output Pages")
				nativePages(numNativePages) = file.Name
				numNativePages = numNativePages + 1
				DebugLog "DEBUG", "ConvertDocument: Found file: " & file.Name & "\n"
			End If
		End If
	Next
	
	If numNativePages = 0 Then
		Call DebugLog("ERROR", "No files to convert\n")
		ConvertDocument = false
		Exit Function
	End If
	
	' ensure we are sorted alphanumerically
	Call QuickSort(nativePages,0, numNativePages-1)
	
	For pc = 0 to numNativePages
		CurrentPrintingExe = false
		pageStartTime = GetEpochMS()
		If nativePages(pc) = "" Then
			Exit For
		End If
		Call gStatsInc(" Total native input pages (files)")
		Set Matches = regFileFormat.Execute(nativePages(pc))
		
		thisFileType = Matches.Item(0).SubMatches.Item(3)
		If thisFileType = "" Then
			thisFileType = "NOEXTENSION"
		End If
		
		onlyIfFirstPage = GetINIString("INMAC " & thisFileType, "OnlyIfFirstPage", "undefined", INMAC_INI)
		cType = GetINIString("INMAC " & thisFileType, "ConversionType", "undefined", INMAC_INI)
		If cType = "undefined" Then
			DebugLog "ERROR", "No Conversion Specified for file type '" & thisFileType & "'\n"
			If GetINIString("INMAC", "UnknownFiletypeAction", "COPY", INMAC_INI) = "ERROR" Then
				convertResult = false
				gLastErrorMsg = "No Conversion for " & thisFileType
			Else
				convertResult = CopyNoConvert(folder.Path, nativePages(pc), outFolderStr)
			End If
		ElseIf cType = "false" Or (onlyIfFirstPage = "true" And pc > 0) Then
			DebugLog "DEBUG", "Not Converting " & thisFileType & "\n"
			convertResult = CopyNoConvert(folder.Path, nativePages(pc), outFolderStr)
		Else
			If UCase(thisFileType) = "PDF" Or UCase(thisFileType) = "XLS" Or UCase(thisFileType) = "XLSX" Then
				cExtractMethod = GetINIString("INMAC " & thisFileType, "TextExtractionMethod", "undefined", INMAC_INI)
				If cExtractMethod = "undefined" Then
					DebugLog "INFO", "No Text Extraction Method Specified for " & thisFileType & "\n"
				Else
					DebugLog "INFO", "Extracting Text from " & thisFileType & " using " & cExtractMethod & "\n"
				
					textExtractResult = ExtractTextWithApplication("INMAC " & thisFileType, cExtractMethod, folder.Path, nativePages(pc), outFolderStr)
					If textExtractResult = false Then
						DebugLog "ERROR", "Problem extracting text from file\n"
						gLastErrorMsg = "Problem extracting text from file\n"
						ConvertDocument = false
						Call NotifyFailed(outFolderStr, CStr(pc+1), thisFileType)
						Exit Function
					End If
				End If
			End If
			
			cMethod = GetINIString("INMAC " & thisFileType, "ConversionMethod", "undefined", INMAC_INI)
			If cMethod = "undefined" Then
				DebugLog "ERROR", "No Conversion Method Specified for " & thisFileType & "\n"
				gLastErrorMsg = "No Method for " & thisFileType
			Else
				DebugLog "INFO", "Converting " & thisFileType & " with " & cType & " using " & cMethod & "\n"
				If UCase(cType) = "INPRINTER" Then
					convertResult = ConvertWithINPrinter("INMAC " & thisFileType, cMethod, folder.Path, nativePages(pc), outFolderStr)
				ElseIf UCase(cType) = "APPLICATION" Then
					convertResult = ConvertWithApplication("INMAC " & thisFileType, cMethod, folder.Path, nativePages(pc), outFolderStr)
				End If
			End If
		End If
		
		If convertResult = false Then
			DebugLog "ERROR", "Problem converting File\n"
			ConvertDocument = false
			Call NotifyFailed(outFolderStr, CStr(pc+1), thisFileType)
			Exit Function
		End If
		
		Call gStatsInc(" Total output pages Converted")
		pageConvertTime=FormatNumber(((GetEpochMS())-pageStartTime)/1000,2)
		Call gStatsAdd(cType & ":" & cMethod & " Total Time", pageConvertTime)
		Call gStatsInc(cType & ":" & cMethod & " Total Converts")
		CAll gStatsAdd("Grand Total Image Conversion Time", pageConvertTime)
		
		DebugLog "INFO", "File Conversion Time: ["&pageConvertTime&"] secs\n"
	Next
	
	ConvertDocument = true
	
	gStatsInc("Total ImageNow Documents Converted")
	
	'Move Original Files to out folder BEFORE inserting to EMA
	For Each file in folder.Files
		If file.Name <> "EXPORT_COMPLETE" Then
			temp = INMAC_OUTPUT_BASE_DIR & inFolder.Name & "\" & SplitFilename(file.Path)(2) & "_ORIG" & SplitFilename(file.Path)(3)
			Call DebugLog("DEBUG", "Moving: [" & file.Path & "] to [" & temp & "]\n")
			rtn = gFsObj.MoveFile( file.Path, temp)
			If Err.number <> 0 Then
				DebugLog "WARNING", "Could not move native file [" & Err.Description & "]: " & file.Path & "\n"
				Err.Clear
			End If
		End If
	Next
	
	' Tell ImageNow workflow scripts that this one is done
	NotifyDone(outFolderStr)
	DebugLog "DEBUG", "Exit ConvertDocument: " & ConvertDocument & "\n"
End Function

' ******************************************************************************
Function PreProcess(folder)
	On Error Resume Next
	PreProcess = false
	
	Dim Matches
	Dim regFileFormat
	Dim thisValue
	
	' For each file, see if it's a zip or a magic number
	Set regFileFormat = New RegExp
	regFileFormat.Pattern = "(.*)\.(.*)$"
	
	For Each file in folder.Files
		Err.Clear
		
		If file.Name <> "EXPORT_COMPLETE" Then
			If regFileFormat.Test(file.Name) <> true Then
				thisFileType = ""
				thisFileName = file.Name
			Else
				Set Matches = regFileFormat.Execute(file.Name)
				thisFileType = Matches.Item(0).SubMatches.Item(1)
				thisFileName = Matches.Item(0).SubMatches.Item(0)
			End If
			
			If UCase(GetINIString(MAGIC_NUM_SEC, "EnableMagicNumber", "", INMAC_INI)) = "TRUE" AND UCase(thisFileType) <> "ZIP" Then
				' If we are unable to identify it via filetype
				If GetINIString("INMAC " & thisFileType, "ConversionType", "", INMAC_INI) = "" Then
					Call DebugLog("DEBUG", "PreProcess: Attempting to determine filetype for [" & thisFileType & "] based on magic number\n")
					keys = split(GetINIKeysInSection(MAGIC_NUM_SEC, INMAC_INI), "^")
					fileType = CheckMagicNumbers(file, keys, thisFileType)
					If fileType <> false Then
						newFileName = file.Path & "." & fileType
						Call DebugLog("INFO", "Renaming [" & file.Path & "] to [" & newFileName  & "]\n")
						file = gFsObj.MoveFile(file.Path, newFileName)
						
						' rename variable for additionaly processing of .zips determined by magic number
						' file.Name = file.Name & "." & fileType
						Set file = gFsObj.getFile(newFileName)
						thisFileType = fileType
						Call gStatsInc("Total filetypes determined from Magic Number")
						If Err.Number <> 0 Then
							Call DebugLog("ERROR", "Could not rename file: " & Err.Description & "\n")
							Exit Function
						End If
					End If ' End found filetype
				End If ' End no filetype specified
			End If ' End Magic Number enabled
			
			If ( UCase(GetINIString("INMAC", "EnableZipExtraction", "", INMAC_INI)) = "TRUE" _
				AND UCase(thisFileType) = "ZIP" ) Then
					Call DebugLog("INFO", "Extracting zip for [" & file.Name & "]\n")
					arrSplitFileName = SplitFileName(WScript.ScriptFullName)
					cmd = """" & arrSplitFileName(1) & "unzip"" -o -j -q -d " & GetINIString("SAVE", "Output Directory", "", INMAC_INI) & " """ & file.Path & """"
					rtn = ExecWaitWithTimeout(cmd, 2000)
					If rtn Then
						Set printerOutputDir = gFsObj.GetFolder( GetINIString("SAVE", "Output Directory", "", INMAC_INI) )
						fc = 0
						For Each f in printerOutputDir.Files
							fc = fc + 1
							Call DebugLog("DEBUG", "Found output file: " & f.Name & "\n")
							Set Matches = regFileFormat.Execute(f.Name)
							If Matches Then
								zippedFileType = Matches.Item(0).SubMatches.Item(1)
							Else
								zippedFileType = ""
							End If
							Call gFsObj.MoveFile(f.Path, folder.Path & "\" & thisFileName & "_" & fc & "." & zippedFileType)
						Next
						gFsObj.DeleteFile(file)
					End If
					
					Call gStatsInc("Total Zip files processed")
			End If ' End if zip extraction
		End If  ' End !EXPORT_COMPLETE
	Next
End Function

' ******************************************************************************
Function CheckMagicNumbers(file, keys, thisFileType)
	On Error Resume Next
	CheckMagicNumbers = false
	
	' 11/20/08-LMS: This is extremely inefficient and will load the entire contents into memory!
	
	' Ensure that the file isn't too big
	If file.size < 1 OR file.size > 500000 Then
		Call DebugLog("WARNING", "CheckMagicNumbers: File size is " & file.size & " will not check magic number\n")
		Exit Function
	End IF
	
	dim inStream,outStream
	const adTypeText=2
	const adTypeBinary=1
	
	set inStream=WScript.CreateObject("ADODB.Stream")
	
	inStream.Open
	inStream.type=adTypeBinary
	
	inStream.LoadFromFile(file)
	
	' Copy the dat over to a stream for outputting
	set outStream=WScript.CreateObject("ADODB.Stream")
	outStream.Open
	outStream.type=adTypeBinary
	
	dim buff
	buff=inStream.Read(50)
	outStream.Close
	
	For x = 0 to UBound(keys) - 1
		If keys(x) <> "EnableMagicNumber" Then
			thisValue = GetINIString(MAGIC_NUM_SEC, keys(x), "", INMAC_INI)
			Call DebugLog("DEBUG", "CheckMagicNumbers: Checking for magic number [" & keys(x) & "] (" & thisValue & ")\n")
			searchBytes = split(thisValue, "|")
			for bc = 0 to UBound(searchBytes)
				If (ascB(midb(buff, bc+1, 1)) <> CInt(searchBytes(bc))) Then
					bc = -1
					Exit For
				End If
			Next
			
			If bc-1 = UBound(searchBytes) Then
				Call DebugLog("DEBUG", file.Name & " is a " & keys(x) & "\n")
				CheckMagicNumbers = keys(x)
				Exit Function
			End If
		End If
	Next
End Function

' ******************************************************************************
Function ConvertWithINPrinter(configSection, method, pathIn, file, outputFolder)
	On Error Resume Next
	Call DebugLog("DEBUG", "ConvertWithINPrinter: Entered " & configSection & ", " & method & ", " & path & ", " & file & "\n")
	Dim jobId, numExpectedPages, fc
	Dim filesToMove()
	ReDim filesToMove(gMaxArraySize)
	Dim cmd, comObj
	
	' Copy fileIn to temp directory and process
	path = CopyToTempDir(pathIn & "\" & file)
	
	' Make sure INPrinter is ready
	If Not IsPrinterReady Then
		' DebugLog "CRITICAL", "\n\n\nProblem with ImageNow Printer: [" & IMAGENOW_PRINTER_NAME & "]  Abort\n" & _
		' "ImageNow Printer must be installed, set as the default, no spooling, and ImageNow Printer\n" & _
		' "Mail Agent Convert must have exclusive access to the ImageNow Printer\n\n"
		SendAdminEmail("The ImageNow Virtual printer was not ready")
		Finish
	End If
	
	' Make sure directory is empty and ready
	Set printerOutputDir = gFsObj.GetFolder( GetINIString("SAVE", "Output Directory", "", INMAC_INI) )
	For Each f in printerOutputDir.Files
		Call DebugLog("WARNING", "Existing output file: " & f.Name & "\n")
	Next
	
	' Update Color reduction to the value of ColorMode for this file type (if there is one)
	ChangedMode = false
	colorMode = GetINIString(configSection, "ColorMode", "BW", INMAC_INI)
	If colorMode <> "" Then
		Call SetPrinterColorReduction(colorMode)
		ChangedMode = true
	End If
	
	' Setup Status file
	Call WriteINIString("PrintStatus", "CurrentJobFinished", "0", INMAC_PRINTER_STATUS_INI)
	Call WriteINIString("PrintStatus", "LastPageNumber", "0", INMAC_PRINTER_STATUS_INI)
	
	If Err.Number <> 0 Then
		Call DebugLog("ERROR", "Clearing Unexpected error: (Error Number = " & Err.Number & ") " & Err.Description & "\n")
		Err.Clear
	End If
	
	command = GetINIString(configSection, "Command", "", INMAC_INI)
	If gFsObj.FileExists(command) Then
		command = gFsObj.GetFile(command).ShortPath
	End If
	
	' Set initial script timeout
	WScript.Timeout = 2 * INPRINTER_JOB_TIMEOUT_SECS
	
	' These are stub functions for different ways to convert
	Select Case UCase(Method)
		Case "ADOBE"
			' Adobe is much faster if you have a copy of it running, so we launch it once on demand if needed
			' This might not be the case for Adobe 9, it seems faster
			If Not AppInitialized.Exists("ADOBE") Then
				Call DebugLog("INFO", "Initializing Adobe\n")
				rtn = AdobeInitialize(command)
				If Not rtn Then
					gLastErrorMsg = "Could not initialize adobe"
					ConvertWithINPrinter = true
					Exit Function
				End If
				Call AppInitialized.Add("ADOBE", "true")
				' Schedule Adobe to be killed
				Call AppFinish.Add("ADOBE", GetINIString(configSection, "CommandExe", "", INMAC_INI))
				WScript.Sleep(CLng(GetINIString(configSection, "LaunchWaitTimeMS", "0", INMAC_INI)))
			End If
			
			cmd = command & " /t " & path & "\" & file
			ExecNoWait(cmd)
			DebugLog "INFO", "Dispatched to Adobe: " & cmd & "\n"
			
		Case "WORDPAD"
			If Not AppInitialized.Exists(Method) Then
				Call DebugLog("INFO", "Initializing " & Method & "\n")
				Call AppFinish.Add( Method, GetINIString(configSection, "CommandExe", "", INMAC_INI) )
			End If
			
			cmd = command & " /p " & Chr(34) & path & "\" & file & chr(34)
			Call ExecNoWait(cmd)
			DebugLog "INFO", "Dispatched to [" & Method & "]: " & cmd & "\n"
			
		Case "FOXIT"
			If Not AppInitialized.Exists(Method) Then
				Call DebugLog("INFO", "Initializing FOXIT\n")
				' Schedule it to be killed
				Call AppFinish.Add(Method, GetINIString(configSection, "CommandExe", "", INMAC_INI))
				Call AppInitialized.Add(Method, "true")
			End If
			
			cmd = command & " /p " & Chr(34) & path & "\" & file & chr(34)
			Call ExecNoWait(cmd)
			DebugLog "INFO", "Dispatched to FoxIt: " & cmd & "\n"
			
		Case "MSPF"
			' Microsoft Picture and Fax Viewer
			If Not AppInitialized.Exists(Method) Then
				Call DebugLog("INFO", "Initializing MSPF\n")
				' Schedule it to be killed
				Call AppFinish.Add( Method, "RUNDLL32.exe" )
				Call AppInitialized.Add(Method, "true")
			End If
			
			' Can split tiffs too
			cmd = "RunDll32.exe shimgvw.dll,ImageView_PrintTo /pt """ & path & "\" & file  & """ """ & IMAGENOW_PRINTER_NAME & """"
			Call ExecNoWait(cmd)
			DebugLog "INFO", "Dispatched to Microsoft Picture and Fax Viewer: " & cmd & "\n"
			
		Case "IRFANVIEW"
			If Not AppInitialized.Exists(Method) Then
				Call DebugLog("INFO", "Initializing " & Method & "\n")
				Call AppFinish.Add( Method, GetINIString(configSection, "CommandExe", "", INMAC_INI) )
			End If
			
			cmd = command & " " & Chr(34) & path & "\" & file & chr(34) & " /print "
			Call ExecNoWait(cmd)
			DebugLog "INFO", "Dispatched to [" & Method & "]: " & cmd & "\n"
			
		Case "IEXPLORER"
			' Internet Explorer
			If AppInitialized.Exists("IEXPLORER") Then
				' IExplorer sometimes seems to have issues automating - to avoid those issues, we simply start with a
				' fresh process each time - IE actually loads and dies quickly, so this isn't a problem
				KillAllProc(GetINIString(configSection, "CommandExe", "", INMAC_INI))
			End If
			
			If Not AppInitialized.Exists("IEXPLORER") Then
				' If we've printed once from, IE, schedule it for exit
				Call DebugLog("INFO", "Initializing IExplorer\n")
				Call AppInitialized.Add("IEXPLORER", "true")
				' Schedule it to be killed
				Call AppFinish.Add("IEXPLORER", GetINIString(configSection, "CommandExe", "", INMAC_INI))
			End If
			
			IEbail = False ' Flag set for bailing out because of timeouts or errors
			DebugLog "DEBUG", "Open IExplore by CreateObject call to create a COM object\n"
			Set comObj = CreateObject("InternetExplorer.Application")
			comObj.Visible = True
			comObj.Silent = True ' prevent dialog boxes from opening
			
			IEBailTime = getEpochMS() + (10000) ' Allow 10 seconds maximum for IE to load
			Do While comObj.busy ' Wait for IE to load before navigating
				If GetEpochMS() > IEBailTime Then
					DebugLog "ERROR", "IExplore took too long to load\n"
					IEbail = True
					Exit Do
				End If
				wscript.sleep(100) ' Give it 100 ms (1/10 s) then check again if IE is loaded
			Loop
			If Not IEbail Then
				DebugLog "DEBUG", "Navigating IE to " & path & "\" & file & "\n"
				Call comObj.Navigate(path & "\" & file)
				
				' Don't continue until IE has fully loaded the page
				IEBailTime = getEpochMS() + (20000) ' Allow another 20 seconds for IE to not be busy and to load this file
				Do While comObj.busy
					If GetEpochMS() > IEBailTime Then
						DebugLog "ERROR", "IExplore busy for 20 seconds after navigation\n"
						IEbail = True
						Exit Do
					End If
					wscript.sleep(100) ' Give it 100 ms (1/10 s) then check again if IE is loaded
				Loop
			End If
			If Not IEbail Then
				Do While comObj.readyState <> 4 ' 4 indicates the page is fully loaded
					If GetEpochMS() > IEBailTime Then
						DebugLog "ERROR", "IExplore took more than 20 seconds to navigate and load page\n"
						IEbail = True
						Exit Do
					End If
					wscript.sleep(100) ' Give it 100 ms (1/10 s) then check again if the page is fully loaded
				Loop
			End If
			
			If Not IEbail Then
				pStatus = comObj.QueryStatusWB(OLECMDID_PRINT)
				If pStatus <> 3 Then ' 3 indicates print command is supported and enabled
					DebugLog "ERROR", "IExplore not allowing Print command. Status value is " & pStatus & "\n"
					IEbail = True
				End If
			End If
			If Not IEbail Then
				' Print
				Call comObj.ExecWB(OLECMDID_PRINT, OLECMDEXECOPT_DONTPROMPTUSER)
				'Set comObj = Nothing
				DebugLog "INFO", "Dispatched to Internet Explorer: " & path & "\" & file & "\n"
			Else
				DebugLog "ERROR", "Stopping errant iexplore process\n"
				comObj.Quit ' Tell it to stop itself
				KillAllProc(GetINIString(configSection, "CommandExe", "", INMAC_INI)) 
			End If
			
		Case "WORDCMD"
			' Microsoft Word from command line
			If Not AppInitialized.Exists(Method) Then
				Call DebugLog("INFO", "Initializing " & Method & "\n")
				Call AppFinish.Add( Method, GetINIString(configSection, "CommandExe", "", INMAC_INI) )
			End If
			
			cmd = command & " " & Chr(34) & path & "\" & file & chr(34) & " /q /n /mFilePrintDefault /mFileExit"
			Call ExecNoWait(cmd)
			DebugLog "INFO", "Dispatched to [" & Method & "]: " & cmd & "\n"
			
		Case "WORD"
			' Microsoft Word
			If Not AppInitialized.Exists(Method) Then
				Call DebugLog("INFO", "Initializing Word\n")
				Call AppInitialized.Add(Method, "true")
				' Schedule it to be killed
				Call AppFinish.Add(Method, GetINIString(configSection, "CommandExe", "", INMAC_INI))
				Set comObj = CreateObject("Word.Application" )
				comObj.Visible = True ' Added 05/28/2010 by BSW
				WScript.Sleep(CLng(GetINIString(configSection, "LaunchWaitTimeMS", "0", INMAC_INI)))
			End If
			
			wordLaunched = false
			
			' this should be cleaner
			jobExpiredAt = GetEpochMS() + (10000)
			Do While True
				Set comObj = GetObject( , "Word.Application" )
				comObj.Visible = True
				
				If Err.Number <> 0 Then
					Call DebugLog("ERROR", "Could not open Word Automation Object: (Error Number = " & Err.Number & ") " & Err.Description & "\n")
					Err.Clear
				Else
					wordLaunched = true
					Exit Do
				End If
				
				If GetEpochMS() > jobExpiredAt Then
					Call DebugLog("ERROR", "Cancelling Job due to JobTimeout, failed to launch Microsoft Word\n")
					Exit Do
				End If
				WScript.sleep(500)
			Loop
			
			If wordLaunched Then ' Else
				Err.Clear
				Call comObj.Documents.Open(path & "\" & file, False, True, False, "abc") 'open with dummy password - this won't affect documents without a password, but will cause password-protected documents to route to the error queue
				If Err.Number <> 0 Then
					Call DebugLog("ERROR", "Could not open file: (Error Number = " & Err.Number & ") " & Err.Description & "\n")
					Err.Clear
				Else
					Err.Clear
					comObj.Options.PrintBackground = false
					If Err.Number <> 0 Then
						Call DebugLog("ERROR", "Could not turn off Print Background in Word: (Error Number = " & Err.Number & ") " & Err.Description & "\n")
						Err.Clear
					Else
						comObj.PrintOut()
						If Err.Number <> 0 Then
							Call DebugLog("ERROR", "Could not print file: (Error Number = " & Err.Number & ") " & Err.Description & "\n")
							Err.Clear
						End If
					End If
				End If
			End If
			
		Case "EXCEL"
			' Microsoft Excel
			If Not AppInitialized.Exists(Method) Then
				Call DebugLog("INFO", "Initializing Excel\n")
				Call AppInitialized.Add(Method, "true")
				' Schedule it to be killed
				Call AppFinish.Add(Method, GetINIString(configSection, "CommandExe", "", INMAC_INI))
				Set comObj = CreateObject("Excel.Application" )
				comObj.Visible = True ' Added 05/28/2010 by BSW
				WScript.Sleep(CLng(GetINIString(configSection, "LaunchWaitTimeMS", "0", INMAC_INI)))
			End If
			
			excellaunched = false
			
			jobExpiredAt = GetEpochMS() + (10000)
			Do While True
				Set comObj = GetObject( , "Excel.Application" )
				
				If Err.Number <> 0 Then
					Call DebugLog("WARNING", "Could not open existing Excel Automation Object: (Error Number = " & Err.Number & ") " & Err.Description & "\n")
					Err.Clear
				Else
					excellaunched = true
					Exit Do
				End If
				
				' might need to launch again
				Set comObj = CreateObject("Excel.Application" )
				If GetEpochMS() > jobExpiredAt Then
					Call DebugLog("ERROR", "Cancelling Job due to JobTimeout, failed to launch Microsoft Excel\n")
					Exit Do
				End If
				WScript.sleep(500)
			Loop
			If excellaunched Then ' Else
				comObj.Visible = True
				comObj.DisplayAlerts = False
				comObj.Application.DisplayAlerts = False
				Err.Clear
				Call comObj.Workbooks.Open(path & "\" & file, 0, True, 2, "abc") 'open with dummy password - this won't affect worksheets without a password, but will cause password-protected worksheets to route to the error queue
				If Err.Number <> 0 Then
					Call DebugLog("ERROR", "Could not open file: (Error Number = " & Err.Number & ") " & Err.Description & "\n")
					Err.Clear
				Else
					If GetINIString(configSection, "FirstTabOnly", "false", INMAC_INI) = "true" Then
						' Print first worksheet only
						Set objSheet = comObj.ActiveWorkbook.Worksheets(1)
						objSheet.PageSetup.PrintQuality = 300
						objSheet.PrintOut()
					Else
						' Print every non-hidden non-empty worksheet
						For Each objSheet in comObj.ActiveWorkbook.Worksheets
							If (objSheet.UsedRange.Cells.Count > 1 Or (objSheet.UsedRange.Cells.Count = 1 And objSheet.Range("a1").Cells.Value <> False)) Then
								objSheet.Activate
								objSheet.PageSetup.PrintQuality = 300
								'objSheet.PrintOut() 'fails if a tab is hidden
							End If
						Next
						comObj.ActiveWorkbook.PrintOut() 'print all non-hidden non-empty worksheets
					End If
					
					If Err.Number <> 0 Then
						Call DebugLog("ERROR", "Could not print file: (Error Number = " & Err.Number & ") " & Err.Description & "\n")
						Err.Clear
					End If
				End If
			End If
			
		Case "POWERPOINT"
			' Microsoft PPT
			If Not AppInitialized.Exists(Method) Then
				Call DebugLog("INFO", "Initializing Excel\n")
				Call AppInitialized.Add(Method, "true")
				' Schedule it to be killed
				Call AppFinish.Add(Method, GetINIString(configSection, "CommandExe", "", INMAC_INI))
			End If
			
			options = GetINIString(configSection, "PrintOptions", "/pt " & IMAGENOW_PRINTER_NAME, INMAC_INI)
			cmd = command & " " & options & " " & path & "\" & file
			ExecNoWait(cmd)
			DebugLog "INFO", "Dispatched to Powerpoint: " & cmd & "\n"
			
		Case "XPS"
			' Microsoft XML Paper Specification (XPS)
			If Not AppInitialized.Exists(Method) Then
				Call DebugLog("INFO", "Initializing XPS\n")
				Call AppInitialized.Add(Method, "true")
				' Schedule it to be killed
				Call AppFinish.Add(Method, GetINIString(configSection, "CommandExe", "", INMAC_INI))
			End If
			
			cmd = command & " """ & IMAGENOW_PRINTER_NAME & """ """ & path & "\" & file & """"
			Call ExecNoWait(cmd)
			DebugLog "INFO", "Dispatched to XPS: " & cmd & "\n"
			
		Case Else
			Call gStatsInc("ERROR Unknown Methods")
			Call DebugLog("ERROR", "ConvertWithINPrinter: Unknown method " & Method & "\n")
			gLastErrorMsg = "Unknown Method"
			ConvertWithINPrinter = false
			Exit Function
			
	End Select
	Call gStatsInc(" Total " & Method & " Converts")
	
	' Wait for results
	rtn = CheckPrintStatus(numExpectedPages, jobId)
	
	' Put Color reduction back to what it was originally
	If ChangedMode = true Then
		Call SetPrinterColorReduction(INMAC_COLOR_REDUCTION)
	End If	

	' If there was an error
	If Not rtn Then
		Call DebugLog("ERROR", "There was an error - waiting " & GetINIString(configSection, "RecoveryWaitSeconds", "60", INMAC_INI) & " seconds for app/printer/jobs to stablize\n")
		WScript.Sleep(CLng(GetINIString(configSection, "RecoveryWaitSeconds", "60", INMAC_INI)) * 1000)
		Call gStatsInc("ERROR Total Conversion Errors")
		Call gStatsInc("ERROR Conversion Errors for method: " & Method)
		ConvertWithINPrinter = false
		
		' Cleanup process if necessary
		Select Case UCase(Method)
			Case "ADOBE"
				AppInitialized.Remove("ADOBE")
				AppFinish.Remove("ADOBE")
				KillAllProc( GetINIString(configSection, "CommandExe", "", INMAC_INI) )
				
			Case "WORD"
				AppInitialized.Remove("WORD")
				AppFinish.Remove("WORD")
				KillAllProc( GetINIString(configSection, "CommandExe", "", INMAC_INI) )
				
			Case "EXCEL"
				AppInitialized.Remove("EXCEL")
				AppFinish.Remove("EXCEL")
				KillAllProc( GetINIString(configSection, "CommandExe", "", INMAC_INI) )
				
			Case "IEXPLORER"
				AppInitialized.Remove("IEXPLORER")
				AppFinish.Remove("IEXPLORER")
				comObj.Quit()
		End Select
		
		' Give the job time to clear after app kill
		WScript.sleep(5000)
		
		' Clean up any files that might have been printed
		CleanDirectory(printerOutputDir)
		
		Exit Function
	Else
		' Per print teardown - no error
		Err.Clear
		Select Case UCase(Method)
			Case "WORD"
				comObj.Documents.Close(wdDoNotSaveChanges) 'don't prompt to save changes
				If Err.Number <> 0 Then
					Call DebugLog("ERROR", "Could not quit word, killing\n")
					AppInitialized.Remove("WORD")
					AppFinish.Remove("WORD")
					KillAllProc( GetINIString(configSection, "CommandExe", "", INMAC_INI) )
					Err.Clear
				End If
				
				Set comObj = Nothing
				
			Case "EXCEL"
				'Dim xwkbk
				'For Each xwkbk In comObj.Workbooks
					'    Could close each workbook by hand
				'Next
				' 08/18/09-LMS:  Do not need to exit excel, just close workbook if no error
				comObj.DisplayAlerts = False
				comObj.Workbooks.Close(False) 'don't prompt to save changes
				comObj.DisplayAlerts = True
				Dim xwkbk
				For Each xwkbk In comObj.Workbooks
					Call DebugLog("WARNING", "Not closed: " & xwkbk.Name & "\n")
				Next
				'comObj.Application.Quit()
				If Err.Number <> 0 Then
					Call DebugLog("ERROR", "Could not quit excel, killing\n")
					AppInitialized.Remove("EXCEL")
					AppFinish.Remove("EXCEL")
					KillAllProc( GetINIString(configSection, "CommandExe", "", INMAC_INI) )
					Err.Clear
				End If
				
				Set comObj = Nothing
				
			Case "IEXPLORER"
				AppInitialized.Remove("IEXPLORER")
				AppFinish.Remove("IEXPLORER")
				comObj.Quit()
				
				Set comObj = Nothing
		End Select
	End If
	
	tiffCount = 0
	fc = 0
	For Each f in printerOutputDir.Files
		call gStatsAdd("Total Output KB", Int( f.size / 1000 ))
		Call DebugLog("DEBUG", "Found output file: " & f.Name & " size: " & f.size & "\n")
		If InStr( f.Name, CStr(jobId) & "_" ) Then
			filesToMove(fc) = f.Path
			fc = fc + 1
			If InStr( f.Name, ".tif" ) Then
				tiffCount = tiffCount + 1
			End If
		Else
			Call DebugLog("WARNING", "Skipping unknown file : " & f.Name & "\n")
		End If
	Next
	
	If fc > 0 Then
		' Give the job time to clear after app kill
		WScript.sleep(5000)
	End If
	
	If tiffCount <> numExpectedPages Then
		Call DebugLog("WARNING", "Page number mismatch Found " & fc & " but expected " & numExpectedPages & "\n")
		' FIXME
	End If
	
	' Create text directory for text files =================================
	textFolderStr = outputFolder & "\text"
	
	Call DebugLog("DEBUG", "Creating text directory " & textFolderStr & "\n")
	gFsObj.CreateFolder textFolderStr
	If (Err.number <> 0) Then
		DebugLog "WARNING", "ConvertWithINPrinter: Could not create text dir: " & Err.Description & "\n"
		' FIXME Exit Function
	End If
	' ======================================================================
	
	Do While fc > 0
		fc = fc - 1
		ext = SplitFilename(filesToMove(fc))(3)
		tempOutDir = outputFolder
		If ext = ".tif" or ext = ".txt" Then
			If ext = ".txt" Then
				tempOutDir = textFolderStr
			End If
			temp = tempOutDir & "\" & SplitFilename("\"&file)(2) & "_" & SplitFilename(filesToMove(fc))(2) & ext
			Call DebugLog("DEBUG", "Moving: [" & filesToMove(fc) & "] to [" & temp & "]\n")
			rtn = gFsObj.MoveFile( filesToMove(fc), temp)
		End If
	Loop
	
	ConvertWithINPrinter = true
	Call DebugLog("DEBUG", "ConvertWithINPrinter: Exit " & ConvertWithINPrinter & " (" & Err.Description& ") \n")
	Err.Clear
End Function

' ******************************************************************************
Function ConvertWithApplication(configSection, method, path, file, outputFolder)
	On Error Resume Next
	
	Call DebugLog("DEBUG", "ConvertWithApplication: Entered " & configSection & ", " & method & ", " & path & ", " & file & "\n")
	Dim jobId, numExpectedPages, fc
	Dim filesToMove(5000)
	Dim cmd, comObj
	
	' Make sure temp directory is empty and ready
	Set tempDir = gFsObj.GetFolder( GetINIString(configSection, "TempDirectory", GetINIString("SAVE", "Output Directory", "", INMAC_INI) , INMAC_INI))
	
	For Each f in tempDir.Files
		Call DebugLog("WARNING", "Existing output file: " & f.Name & "\n")
	Next
	
	command = GetINIString(configSection, "Command", "", INMAC_INI)
	If gFsObj.FileExists(command) Then
		command = gFsObj.GetFile(command).ShortPath
	End If
	
	' Set initial script timeout
	WScript.Timeout = 2 * INPRINTER_JOB_TIMEOUT_SECS
	
	Select Case UCase(Method)
		Case "INIMAGETOOL"
			If Not AppInitialized.Exists("INIMAGETOOL") Then
				Call DebugLog("INFO", "Initializing InImageTool\n")
				Call AppInitialized.Add("INIMAGETOOL", "true")
				' Schedule it to be killed
				Call AppFinish.Add("INIMAGETOOL", GetINIString(configSection, "CommandExe", "", INMAC_INI))
			End If
			
			cmd = command & " """ &  path & "\" & file _
				& """  --mode split --dir """ & tempDir.Path & """ --bw --dpi " _
				& GetINIString(configSection, "DPI", "240", INMAC_INI)
				
			convertResult = ExecWaitWithTimeout(cmd, INPRINTER_JOB_TIMEOUT_SECS)
			DebugLog "INFO", "InImagetool completed: " & cmd & "\n"
'BEGIN GREGG	
		Case "POWERGHOST"
			dim codeCatch
			If Not AppInitialized.Exists("POWERGHOST") Then
				Call DebugLog("INFO", "Initializing POWERGHOST\n")
				Call AppInitialized.Add("POWERGHOST", "true")
				' Schedule it to be killed
				Call AppFinish.Add("POWERGHOST", GetINIString(configSection, "CommandExe", "", INMAC_INI))
			End If
			
			cmd = GetINIString(configSection, "Command", "", INMAC_INI) & " " & path & "\ " & outputFolder & "\ " & file & " " & INMAC_INSTANCE
			Set WshShell = WScript.CreateObject("WScript.Shell")
			codeCatch = WshShell.Run(cmd, 8, true)
			Call DebugLog("INFO","codeCatch is " & codeCatch & "\n")
			if codeCatch=0 Then 
			convertResult=true
			ElseIf codeCatch=1 Then
			convertResult=false
			gLastErrorMsg = "PDF to tif conversion timed out"
			ElseIf codeCatch=2 Then
			convertResult=false
			gLastErrorMsg = "Page size out of printing range"
			ElseIf codeCatch=3 Then
			convertResult=false
			gLastErrorMsg = "Some pages could not be converted"
			Else
			convertResult=false
			gLastErrorMsg = "PDF to tif conversion was not successful Reason: " & codeCatch & "."
			End If
			Call DebugLog("INFO","convertResult is " & convertResult & "\n")			
			DebugLog "INFO", "GhostScript completed: " & cmd & "\n"

		Case "POWERPRINT"
			dim successCase
			If Not AppInitialized.Exists("POWERPRINT") Then
				Call DebugLog("INFO", "Initializing POWERPRINT\n")
				Call AppInitialized.Add("POWERPRINT", "true")
				' Schedule it to be killed
				Call AppFinish.Add("POWERPRINT", GetINIString(configSection, "CommandExe", "", INMAC_INI))
			End If
			

			cmd = GetINIString(configSection, "Command", "", INMAC_INI) & " " & path & "\ " & outputFolder & "\ " & file & " " & INMAC_INSTANCE
			Call DebugLog("INFO","cmd is " & cmd & "\n")
			Set WshShell = WScript.CreateObject("WScript.Shell")
			successCase = WshShell.Run(cmd, 8, true)
			Call DebugLog("INFO"," successCase is " & successCase & "\n")
			if successCase=0 Then 
			convertResult=true
			ElseIf successCase=1 Then
			convertResult=false
			gLastErrorMsg = "Document is password-protected"
			ElseIf successCase=2 Then
			convertResult=false
			gLastErrorMsg = "File type is blocked by File Block settings on the server."
			ElseIf successCase=3 Then
			convertResult=false
			gLastErrorMsg = "The file appears to be corrupted."
			ElseIf successCase=4 Then
			convertResult=false
			gLastErrorMsg = "Office has determined there may be a probelm with this file and it should be opened with caution."
			ElseIf successCase=5 Then
			convertResult=false
			gLastErrorMsg = "This file contains a macro that cannot be used in an automated process. Open with caution."
			ElseIf successCase=6 Then
			convertResult=false
			gLastErrorMsg = "The file format does not match the file extension. You should not attempt to open the file."
			ElseIf successCase=7 Then
			convertResult=false
			gLastErrorMsg = "This file has a one or more issues that prevent automated processing."
			Else
			convertResult=false
			gLastErrorMsg = "Doc conversion failed for an unspecified reason"
			End If
			Call DebugLog("INFO","convertResult is " & convertResult & "\n")
			DebugLog "INFO", " completed: " & cmd & "\n"

		'END GREGG
			
		Case Else
			Call gStatsInc("ERROR Unknown Methods")
			Call DebugLog("ERROR", "ConvertWithApplication: Unknown method " & Method & "\n")
			gLastErrorMsg = "Unknown Method"
			ConvertWithApplication = false
			Exit Function
	End Select
	Call gStatsInc(" Total " & Method & " Converts")
	
	' If there was an error
	If Not convertResult Then
		ConvertWithApplication = false
		Call gStatsInc("ERROR Total Conversion Errors")
		
		' Cleanup if necessary
		ConvertWithApplication = false
		
		' Give the job time to clear after app kill
		WScript.sleep(5000)
		
		' Clean up any files that might have been printed
		CleanDirectory(tempDir)
		
		Exit Function
	End If
	
	fc = 0
	For Each f in tempDir.Files
		Call DebugLog("DEBUG", "Found output file: " & f.Name & "\n")
		If InStr( f.Name, CStr(jobId) & "_" ) Then
			filesToMove(fc) = f.Path
			fc = fc + 1
		Else
			Call DebugLog("WARNING", "Skipping unknown file : " & f.Name & "\n")
		End If
	Next
	
	' Create text directory for text files =================================
	textFolderStr = outputFolder & "\text"
	
	Call DebugLog("DEBUG", "Creating text directory " & textFolderStr & "\n")
	gFsObj.CreateFolder textFolderStr
	If (Err.number <> 0) Then
		DebugLog "WARNING", "ConvertWithApplication: Could not create text dir: " & Err.Description & "\n"
		' FIXME Exit Function
	End If
	' ======================================================================
	
	Do While fc > 0
		fc = fc - 1
		temp = outputFolder & "\" & SplitFilename(filesToMove(fc))(2) & ".tif"
		Call DebugLog("DEBUG", "Moving: [" & filesToMove(fc) & "] to [" & temp & "]\n")
		rtn = gFsObj.MoveFile( filesToMove(fc), temp)
	Loop
	
	ConvertWithApplication = true
	Call DebugLog("DEBUG", "ConvertWithApplication: Exit " & ConvertWithApplication & " (" & Err.Description& ") \n")
	Err.Clear
End Function

' ******************************************************************************
Function CopyToTempDir(file)
	On Error Resume Next
	Err.Clear
	ext = SplitFilename(file)(3)
	tempDir = GetTempDir()
	temp = tempDir & "\" & SplitFilename(file)(2) & ext
	Call DebugLog("DEBUG", "Copy: [" & file & "] to [" & temp & "]\n")
	Call gFsObj.CopyFile( file, temp)
	CopyToTempDir = tempDir
End Function

' ******************************************************************************
Function GetTempDir()
	On Error Resume Next
	Err.Clear
	Set f = gFsObj.GetFolder(gTempDirBase)
	If Err.Number <> 0 Then
		Call DebugLog("ERROR", "Could not get temp folder\n")
		Finish
	End If
	
	nextNum = 0
	Set fc = f.SubFolders
	Err.Clear
	Do While Err.Number = 0
		nextNum = nextNum + 1
		temp = fc.Item("" & nextNum).Name
	Loop
	
	gFsObj.CreateFolder(gTempDirBase & nextNum)
	GetTempDir = gTempDirBase & nextNum
End Function

' ******************************************************************************
Function ExtractTextWithApplication(configSection, method, path, file, outputFolder)
	On Error Resume Next
	
	Call DebugLog("DEBUG", "ExtractTextWithApplication: Entered " & configSection & ", " & method & ", " & path & ", " & file & "\n")
	Dim jobId, numExpectedPages, fc, tempFilePath, tempFileName
	Dim filesToMove(5000)
	Dim cmd, comObj
	
	' Make sure temp directory is empty and ready
	Set tempDir = gFsObj.GetFolder( GetINIString(configSection, "TempDirectory", GetINIString("SAVE", "Output Directory", "", INMAC_INI) , INMAC_INI))
	
	For Each f in tempDir.Files
		Call DebugLog("WARNING", "Existing output file: " & f.Name & "\n")
	Next
	
	' reset script timeout
	WScript.Timeout = 2 * INPRINTER_JOB_TIMEOUT_SECS
	
	Select Case UCase(Method)
		Case "PDFTOTEXT"
			If Not AppInitialized.Exists("PDFTOTEXT") Then
				Call DebugLog("INFO", "Initializing PDFTOTEXT\n")
				Call AppInitialized.Add("PDFTOTEXT", "true")
				' Schedule it to be killed
				Call AppFinish.Add("PDFTOTEXT", GetINIString(configSection, "TextExtractionCommandExe", "", INMAC_INI))
			End If
			
			tempFileName = file
			tempFileName = Replace(tempFileName, ".pdf", ".txt", 1, -1, 1)
			tempFilePath = tempDir.Path & "\" & tempFileName
			cmd = GetINIString(configSection, "TextExtractionCommand", "", INMAC_INI) & " -layout """ & path & "\" & file & """  " & tempFilePath
			
			txtExtractionResult = ExecWaitWithTimeout(cmd, INPRINTER_JOB_TIMEOUT_SECS)
			DebugLog "INFO", "PDFTOTEXT completed: " & cmd & "\n"
		Case "EXCEL"
			If Not AppInitialized.Exists("EXCEL") Then
				Call DebugLog("INFO", "Initializing EXCEL\n")
				Call AppInitialized.Add("EXCEL", "true")
				' Schedule it to be killed
				Call AppFinish.Add("EXCEL", GetINIString(configSection, "TextExtractionCommandExe", "", INMAC_INI))
				Set comObj = CreateObject("Excel.Application")
				WScript.Sleep(CLng(GetINIString(configSection, "TextExtractionLaunchWaitTimeMS", "0", INMAC_INI)))
			End If
			
			tempFileName = file
			tempFileName = Replace(tempFileName, ".xlsx", ".txt")
			tempFileName = Replace(tempFileName, ".xls", ".txt")
			tempFilePath = tempDir.Path & "\" & tempFileName
			
			excellaunched = false
			jobExpiredAt = GetEpochMS() + (10000)
			Do While True
				Set comObj = GetObject( , "Excel.Application" )
				
				If Err.Number <> 0 Then
					Call DebugLog("WARNING", "Could not open Excel Automation Object: (Error Number = " & Err.Number & ") " & Err.Description & "\n")
					Err.Clear
				Else
					excellaunched = true
					Exit Do
				End If
				
				' might need to launch again
				Set comObj = CreateObject("Excel.Application" )
				If GetEpochMS() > jobExpiredAt Then
					Call DebugLog("ERROR", "Cancelling Job due to JobTimeout, failed to launch Microsoft Excel\n")
					Exit Do
				End If
				WScript.sleep(500)
			Loop
			If excellaunched Then ' Else
				comObj.Visible = True
				comObj.DisplayAlerts = False
				comObj.Application.DisplayAlerts = False
				Err.Clear
				Call comObj.Workbooks.Open(path & "\" & file, 0, True, 2, "abc")
				
				If Err.Number <> 0 Then
					Call DebugLog("ERROR", "Could not open file: (Error Number = " & Err.Number & ") " & Err.Description & "\n")
					Err.Clear
				Else
					If GetINIString(configSection, "FirstTabOnly", "false", INMAC_INI) = "true" Then
						' Print first worksheet only
						set objSheet = comObj.ActiveWorkbook.Worksheets(1)
						objSheet.SaveAs tempFilePath, 21 'creates a tab delimited file
					Else
						' Extract text from all non-empty worksheets
						sheetNum=0
						For Each objSheet In comObj.ActiveWorkbook.Worksheets
							If (objSheet.UsedRange.Cells.Count > 1 Or (objSheet.UsedRange.Cells.Count = 1 And objSheet.Range("a1").Cells.Value <> False)) Then
								objSheet.Activate
								objSheet.SaveAs tempFilePath & sheetNum & ".txt", 21 'creates a tab delimited file
								sheetNum=sheetNum+1
							End If
						Next
					End If
					If Err.Number <> 0 Then
						Call DebugLog("ERROR", "Could not save file as csv (txt): (Error Number = " & Err.Number & ") " & Err.Description & "\n")
						txtExtractionResult = false
						Err.Clear
					Else
						txtExtractionResult = true
						DebugLog "INFO", "EXCEL completed text extraction\n"
					End If
				End If
				
				'Dim xwkbk
				'For Each xwkbk In comObj.Workbooks
					'    Could close each workbook by hand
				'Next
				' 08/18/09-LMS:  Do not need to exit excel, just close workbook if no error
				comObj.Workbooks.Close()
				comObj.DisplayAlerts = True
				Dim xwkbk
				For Each xwkbk In comObj.Workbooks
					Call DebugLog("WARNING", "Not closed: " & xwkbk.Name & "\n")
				Next
				'comObj.Application.Quit()
				If Err.Number <> 0 Then
					Call DebugLog("ERROR", "Could not quit excel, killing\n")
					AppInitialized.Remove("EXCEL")
					AppFinish.Remove("EXCEL")
					KillAllProc( GetINIString(configSection, "TextExtractionCommandExe", "", INMAC_INI) )
					Err.Clear
				End If
				
				Set comObj = Nothing
			End If
		Case Else
			Call gStatsInc("ERROR Unknown Methods")
			Call DebugLog("ERROR", "ExtractTextWithApplication: Unknown method " & Method & "\n")
			gLastErrorMsg = "Unknown Method"
			ExtractTextWithApplication = false
			Exit Function
	End Select
	Call gStatsInc(" Total " & Method & " Text Extractions")
	
	' If there was an error
	If Not txtExtractionResult Then
		ExtractTextWithApplication = false
		Call gStatsInc("ERROR Total Text Extraction Errors")
		
		' Cleanup if necessary
		ExtractTextWithApplication = false
		
		' Give the job time to clear after app kill
		WScript.sleep(5000)
		
		' Clean up any files that might have been printed
		CleanDirectory(tempDir)
		
		Exit Function
	End If
	
	fc = 0
	For Each f in tempDir.Files
		Call DebugLog("DEBUG", "Found output file: " & f.Name & "\n")
		If InStr( f.Name, CStr(tempFileName) ) Then
			filesToMove(fc) = f.Path
			fc = fc + 1
		Else
			Call DebugLog("WARNING", "Skipping unknown file : " & f.Name & "\n")
		End If
	Next
	
	' Create text directory for text files =================================
	textFolderStr = outputFolder & "\extractedText"
	
	gFsObj.CreateFolder textFolderStr
	If (Err.number <> 0) Then
		DebugLog "WARNING", "ConvertDocument: Could not create text dir: " & Err.Description & "\n"
		' FIXME Exit Function
	End If
	' ======================================================================
	
	Do While fc > 0
		fc = fc - 1
		ext = SplitFilename(filesToMove(fc))(3)
		tempOutDir = textFolderStr
		temp = tempOutDir & "\" & SplitFilename("\"&file)(2) & "_" & SplitFilename(filesToMove(fc))(2) & ext
		Call DebugLog("DEBUG", "Moving: [" & filesToMove(fc) & "] to [" & temp & "]\n")
		rtn = gFsObj.MoveFile( filesToMove(fc), temp)
	Loop
	
	ExtractTextWithApplication = true
	Call DebugLog("DEBUG", "ExtractTextWithApplication: Exit " & ExtractTextWithApplication & " (" & Err.Description& ") \n")
	Err.Clear
End Function

' ******************************************************************************
Function NotifyFailed(outFolderStr, pc, thisFileType)
	On Error Resume Next
	Call WriteFile(outFolderStr & "\CONVERSION_ERROR", gLastErrorMsg & " (page " & CStr(pc) & " " & thisFileType & ")")
	If UCase(INMAC_USE_EXTERN_MSG) = "TRUE" Then
		Call InsertExternMsg("failed", ParseDocId(outFolderStr), gLastErrorMsg & " (page " & CStr(pc) & " " & thisFileType & ")")
	End if
End Function

' ******************************************************************************
Function SetPrinterColorReduction(iniValue)
	On Error Resume Next
	
	Call DebugLog("DEBUG", "Setting up color settings for " & iniValue & "\n")
	
	If(UCase(iniValue) = "COLOR") Then
		Call WriteINIString("Compression", "Color reduction", "Optimal", INMAC_INI)
	ElseIf (UCase(iniValue) = "GREY") Then
		Call WriteINIString("Compression", "Color reduction", "GREY", INMAC_INI)
	Else
		Call WriteINIString("Compression", "Color reduction", "BW", INMAC_INI)
	End if
End Function

' ******************************************************************************
Function NotifyDone(outFolderStr)
	Call WriteFile(outFolderStr & "\CONVERSION_COMPLETE", CStr(GetTimeStamp()))
	WScript.sleep(1000)
	Call DebugLog("DEBUG", "Created CONVERSION_COMPLETE file here: " & outFolderStr & "\n")
	dim isItReall
	isItReall = ReportFileStatus(outFolderStr&"\CONVERSION_COMPLETE")
	'WScript.sleep(1000)
	Call DebugLog("DEBUG", isItReall & "\n")
	If UCase(INMAC_USE_EXTERN_MSG) = "TRUE" Then
		'WScript.sleep(5000) 'EMA may run the process script before the files exist if "Offline Files" is enabled
		Call InsertExternMsg("success", ParseDocId(outFolderStr), "")
	End if
End Function

' ******************************************************************************
Function ParseDocId(outFolderStr)
	ParseDocId = "Unknown"
	Dim regFileFormat
	Dim Matches
	Set regFileFormat = New RegExp
	regFileFormat.Pattern = ".*\\.*?_.*?_(.*_.*)"
	Set Matches = regFileFormat.Execute(outFolderStr)
	docId = Matches.Item(0).SubMatches.Item(0)
	If docId = "" Then
		Call DebugLog("ERROR", "Could not find docid\n")
		Exit Function
	End If
	ParseDocId = docId
End Function

' ******************************************************************************
Function InsertExternMsg(result, docId, msg)
	On Error Resume Next
	' 11/18/08-LMS: This is broken out to emulate 'continue' control functionality not available in vbs
	dbUdated = false
	
	For numTries = 1 to CInt( GetINIString("INMAC External Messaging", "maxTries", "5", INMAC_INI) )
		Call DebugLog("DEBUG", "InsertExternMsg: Try " & numTries & "\n")
		
		dbUdated = InsertExternMsgLoop(result, docId, msg)
		If dbUdated = true Then
			InsertExternMsg = true
			Exit For
		End If
	Next
	
	If dbUdated = false Then
		Call DebugLog("CRITICAL", gLastErrorMsg & "\n")
		SendAdminEmail(gLastErrorMsg)
	End If
End Function

' ******************************************************************************
Function InsertExternMsgLoop(result, docId, msg)
	On Error Resume Next
	Err.Clear
	Dim START_TIME
	'START_TIME = "CURRENT_TIMESTAMP"
	START_TIME = "01-01-1970 00:00:00"
	
	InsertExternMsgLoop = false
	
	' 11/18/08-LMS: This is broken out to emulate 'continue' control functionality not available in vbs
	msgId = IMAGENOW_PRINTER_NAME & "_" & GetEpochMS()
	
	Dim sqlCmd
	
	' reset script timeout
	WScript.Timeout = 2 * INPRINTER_JOB_TIMEOUT_SECS
	
	If gAdoConnection.state <> 1 Then
		If Not OpenImageNowDb() Then
			Exit Function
		End If
	End If
	
	If Err.number <> 0 Then
		gLastErrorMsg = "InsertExternMsg: DB ERROR: " & Err.Description
		Exit Function
	End If
	
	' Wrap transaction
	gAdoConnection.BeginTrans
	
	If UCase(INMAC_USE_ORACLE_DB) = "TRUE" Then
		' Set Oracle session datetime format
		sessionParms = "ALTER SESSION SET NLS_TIMESTAMP_FORMAT = 'RRRR-MM-DD HH24:MI:SS.FF'"
		Call DebugLog("DEBUG", "ORACLE SessionParms: Executing: " & sessionParms & "\n")
		gAdoConnection.Execute(sessionParms)
		If Err.number <> 0 Then
			gLastErrorMsg = "SessionParms: DB ERROR: " & Err.Description
			Exit Function
		End If
	End If
	
	sqlCmd = "insert into in_extern_msg (EXTERN_MSG_ID, MSG_TYPE, MSG_NAME, MSG_DIRECTION, MSG_STATUS, START_TIME) values ( '" & msgId & "', 'inmac', 'inmac-done', 1, 1, CURRENT_TIMESTAMP)"
	Call DebugLog("DEBUG", "InsertExternMsgLoop: Executing: " & sqlCmd & "\n")
	gAdoConnection.Execute(sqlCmd)
	If Err.number <> 0 Then
		gLastErrorMsg = "InsertExternMsg: DB ERROR: " & Err.Description
		Exit Function
	End If
	
	sqlCmd = "insert into in_extern_msg_prop (EXTERN_MSG_ID, PROP_NAME, PROP_TYPE, PROP_VALUE) values ('" & msgId & "', 'docId', 0,'" & docId &"')"
	Call DebugLog("DEBUG", "InsertExternMsgLoop: Executing: " & sqlCmd & "\n")
	gAdoConnection.Execute(sqlCmd)
	If Err.number <> 0 Then
		gLastErrorMsg = "InsertExternMsg: DB ERROR: " & Err.Description
		Exit Function
	End If
	
	sqlCmd = "insert into in_extern_msg_prop (EXTERN_MSG_ID, PROP_NAME, PROP_TYPE, PROP_VALUE) values ('" & msgId & "', 'result', 0,'" & result &"')"
	Call DebugLog("DEBUG", "InsertExternMsgLoop: Executing: " & sqlCmd & "\n")
	gAdoConnection.Execute(sqlCmd)
	If Err.number <> 0 Then
		gLastErrorMsg = "InsertExternMsg: DB ERROR: " & Err.Description
		Exit Function
	End If
	
	sqlCmd = "insert into in_extern_msg_prop (EXTERN_MSG_ID, PROP_NAME, PROP_TYPE, PROP_VALUE) values ('" & msgId & "', 'message', 0,'" & msg &"')"
	Call DebugLog("DEBUG", "InsertExternMsgLoop: Executing: " & sqlCmd & "\n")
	gAdoConnection.Execute(sqlCmd)
	If Err.number <> 0 Then
		gLastErrorMsg = "InsertExternMsg: DB ERROR: " & Err.Description
		Exit Function
	End If
	
	gAdoConnection.CommitTrans
	
	InsertExternMsgLoop = true
	
'	gAdoConnection.BeginTrans
'	gAdoConnection.Execute( "insert into in_extern_msg (EXTERN_MSG_ID, MSG_TYPE, MSG_NAME, MSG_DIRECTION, MSG_STATUS, START_TIME) values ( '" & msgId & "', 'inmac', 'inmac-done', 1, 1, CURRENT_TIMESTAMP)")
'	gAdoConnection.Execute( "insert into in_extern_msg_prop (EXTERN_MSG_ID, PROP_NAME, PROP_TYPE, PROP_VALUE) values ('" & msgId & "', 'result', 0,'" & result &"')")
'	If Err.number <> 0 Then
'		msgbox err.Description
'		Exit Function
'	End If
'	gAdoConnection.Execute( "insert into in_extern_msg_prop (EXTERN_MSG_ID, PROP_NAME, PROP_TYPE, PROP_VALUE) values ('" & msgId & "', 'message', 0,'" &  msg &"')")
'	If Err.number <> 0 Then
'		msgbox err.Description
'		Exit Function
'	End If
'	gAdoConnection.CommitTrans
'	If Err.number <> 0 Then
'		msgbox err.Description
'		Exit Function
'	End If
'
'	msgbox Err.Description
End Function

' ******************************************************************************
Function OpenImageNowDb()
	On Error Resume Next
	
	Err.Clear
	
	Set gAdoConnection = CreateObject("ADODB.Connection")
	
	gAdoConnection.Open("Data Source=" & _
		GetINIString("INMAC External Messaging", "dsn", "", INMAC_INI) & "; Uid=" & _
		GetINIString("INMAC External Messaging", "userid", "", INMAC_INI) & "; Pwd=" & _
		GetINIString("INMAC External Messaging", "password", "", INMAC_INI))
		
	If Err.Number <> 0 Then
		gLastErrorMsg = Err.Description
		Exit Function
	End If
	
	Call DebugLog("INFO", "Imagenow DB opened.\n")
	OpenImageNowDb = true
End Function

' ******************************************************************************
Function ExecNoWait(cmd)
	' reset script timeout
	WScript.Timeout = 2 * INPRINTER_JOB_TIMEOUT_SECS
	
	strComputer = "."
	Set objWMIService = GetObject("winmgmts:" _
		& "{impersonationLevel=impersonate}!\\" & strComputer _
		& "\root\cimv2:Win32_Process")
		
	Call DebugLog("INFO", "ExecNoWait: [" & cmd & "]\n")
	rtn = objWMIService.Create(cmd , null, null, pid)
	If Err.Number <> 0 Then
		Call DebugLog("INFO", "Error Number: " & Err.Number & ", Error: " & Err.Description & "\n")
		Err.Clear
	End If
	Call DebugLog("DEBUG", "ExecNoWait: returned: [" & rtn & "] (pid: " & pid & ")\n")
	ExecNoWait = pid
End Function

' ******************************************************************************
Function ExecWaitWithTimeout(cmd, timeout)
	Dim pid, jobExpiredAt, foundPid
	ExecWaitWithTimeout = false
	
	foundPid = false
	foundOnce = true
	
	Call DebugLog("DEBUG", "ExecWaitWithTimeout Enter " & cmd & ", " & timeout & "\n")
	pid = ExecNoWait(cmd)
	
	strComputer = "."
	jobExpiredAt = GetEpochMS() + (CInt(timeout) * 1000)
	Do While True
		Set colProcesses = GetObject("winmgmts:" & _
			"{impersonationLevel=impersonate}!\\" & strComputer & _
			"\root\cimv2").ExecQuery("Select * from Win32_Process")
		
		For Each objProcess in colProcesses
			If CDbl(objProcess.ProcessId) = CDbl(pid) Then
				foundPid = true
				foundOnce = true
				
				Call DebugLog("DEBUG", "Found PID " & pid & " belonging to " & objProcess.Name & "\n")
			End If
		Next
		
		If Not foundOnce Then
			Call DebugLog("WARNING", "Could not find pid " & pid & "\n")
			gLastErrorMsg = "ExecWaitWithTimeout: could not find pid"
			ExecWaitWithTimeout = false
			Exit Function
		End If
		
		If foundOnce AND Not foundPid Then
			' Must have exited normally
			ExecWaitWithTimeout = true
			Exit Function
		End If
		foundPid = false
		
		If GetEpochMS() > jobExpiredAt Then
			Call gStatsInc("ERROR Job Timeouts (App)")
			DebugLog "ERROR", "Cancelling Job due to JobTimeout, failed to run " & objProcess.Name & "\n"
			objProcess.Terminate
			DebugLog "NOTIFY", "Killed process " & objProcess.ProcessID & " " & objProcess.Name & "\n"
			gLastErrorMsg = "ExecWaitWithTimeout: Job timeout"
			ExecWaitWithTimeout = false
			Exit Function
		End If
		WScript.Sleep 500
	Loop
End Function

' ******************************************************************************
Function CleanDirectory(printerOutputDir)
	For Each f in printerOutputDir.Files
		If InStr( f.Name, ".tif") Then
			Call DebugLog("INFO", "Removing output file: " & f.Name & "\n")
			f.Delete()
		Else
			Call DebugLog("WARNING", "Skipping unknown file : " & f.Name & "\n")
		End If
	Next
End Function

' ******************************************************************************
Function AdobeInitialize(cmd)
	AdobeInitialize = false
	
	Dim pid, rtn
	pid = ""
	rtn = false
	
	pid = ExecNoWait(cmd)
	Call DebugLog("DEBUG", "AdobeInitialize: PID: " & pid & "\n")
	
	' Give Adobe Time to start up
	Call WScript.Sleep(1000)
	
	If Not Error Then
		AdobeInitialize = true
	End If
End Function

' ******************************************************************************
Function CopyNoConvert(path, file, outputFolder)
	CopyNoConvert = false
	DebugLog "DEBUG", "CopyNoConvert: Enter " & path & ", " & file & ", " & outputFolder & "\n"
	
	temp = outputFolder & "\" & file
	Call DebugLog("DEBUG", "Copying: [" & path&"\"&file & "] to [" & temp & "]\n")
	rtn = gFsObj.CopyFile( path&"\"&file, temp)
	
	CopyNoConvert = true
	DebugLog "DEBUG", "CopyNoConvert: Exit " & CopyNoConvert & "\n"
End Function

' ******************************************************************************
Function CheckPrintStatus( ByRef TotalPages, ByRef jobId )
	CheckPrintStatus = false
	
	PageTimeout = INPRINTER_PAGE_TIMEOUT_SECS * 1000
	pageExpiredAt = GetEpochMS() + PageTimeout
	LastPage = 0
	
	'set initial script timeout to stop after job timeout
	WScript.Timeout = 2 * INPRINTER_JOB_TIMEOUT_SECS
	
	Do while true
		curPage = CInt(GetINIString("PrintStatus", "LastPageNumber", "0", INMAC_PRINTER_STATUS_INI))
		
		isDone = GetINIString("PrintStatus", "CurrentJobFinished", "0", INMAC_PRINTER_STATUS_INI)
		If isDone = 1 Then
			jobId = CLng(GetINIString("PrintStatus", "LastJobPrintId", "-1", INMAC_PRINTER_STATUS_INI))
			If (jobId = gLastPrintJobId) OR (jobId = -1) Then
				Call DebugLog("DEBUG", "No new job yet\n")
				curPage = 0
				isDone = 0
			Else
				TotalPages = CInt(GetINIString("PrintStatus", "LastJobTotalPages", "0", INMAC_PRINTER_STATUS_INI))
				Call DebugLog("INFO", "Print job successful [" & jobId & "]  Total Pages: " & TotalPages & "\n")
				gLastPrintJobId = CLng(jobId)
				Exit Do
			End If
		End If
		
		If GetEpochMS() > pageExpiredAt Then
			Call gStatsInc("ERROR Page Timeouts")
			DebugLog "ERROR", "Cancelling Job due to PageTimeout\n"
			gLastErrorMsg = "Printer Page Timeout"
			Exit Do
		End If
		
		If curPage > LastPage Then
			pageExpiredAt = GetEpochMS() + PageTimeout 'a new page was found, reset the page timeout
			LastPage = curPage
			DebugLog "DEBUG", "Found " & CStr(LastPage) & " Pages\n"
		End If
		
		' reset script timeout
		WScript.Timeout = 2 * INPRINTER_JOB_TIMEOUT_SECS
		WScript.Sleep 500
	Loop
	
	' reset script timeout
	WScript.Timeout = 2 * INPRINTER_JOB_TIMEOUT_SECS
	
	If isDone Then
		printStatus = GetINIString("PrintStatus", "LastJobStatus", "0", INMAC_PRINTER_STATUS_INI)
		If printStatus = "1" Then
			CheckPrintStatus = true
			Exit Function
		End If
	Else
		If CurrentPrintingExe <> false Then
			KillAllProc(CurrentPrintingExe)
		End If
	End If
End Function

' ******************************************************************************
Function IsPrinterReady
	Dim objWMIService, objItem, colItems, strComputer, defaultPrinter
	IsPrinterReady = false
	
	strComputer ="."
	
	For x = 1 to 3
		' --------------------------------------------
		' Pure WMI Section
		Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\CIMV2")
		Set colItems = objWMIService.ExecQuery("SELECT * FROM Win32_Printer where DeviceID = '" & IMAGENOW_PRINTER_NAME & "'")
		
		' Has to be the default printer with a status of zero
		For Each objItem In colItems
		
			If objItem.Attributes And 4 Then
				defaultPrinter = "True"
			Else
				defaultPrinter = "False"
			End If
		
			If defaultPrinter = "True" AND objItem.PrinterState = "0" AND objItem.SpoolEnabled = "False" Then
				IsPrinterReady = true
				Exit Function
			Else
				If objItem.SpoolEnabled = "True" Then
					Call DebugLog("CRITICAL", "Spooling has not been turned off for ImageNow printer " & IMAGENOW_PRINTER_NAME & ".\n")
				End If
				
				If objItem.PrinterState <> "0" Then
					Call DebugLog("CRITICAL", "The PrinterState of ImageNow printer " & IMAGENOW_PRINTER_NAME & " is " & objItem.PrinterState & " and it should be 0.\n")
				End If

				If defaultPrinter = "False" Then
					Call DebugLog("CRITICAL", "ImageNow printer " & IMAGENOW_PRINTER_NAME & " is not the default printer for user " & gNetworkUser & " and it needs to be.\n")
					'Added 07/08/2010 by BSW to set default printer to printer from INMAC.ini file
					If FORCE_DEFAULT Then
						Call DebugLog("INFO", "Forcefully setting default printer to " & IMAGENOW_PRINTER_NAME & "\n")
						Dim objPrinter
						Set objPrinter = CreateObject("WScript.Network") 
						objPrinter.SetDefaultPrinter IMAGENOW_PRINTER_NAME
						IsPrinterReady = true
					End If
				End If
			
				'Call DebugLog("WARNING", "Printer was not ready - default=" & defaultPrinter & " state=" & objItem.PrinterState & " " & objItem.SpoolEnabled & "\n")
			End If
		Next
		
		WScript.Sleep 500
	Next
End Function

' ******************************************************************************
Function LogPrinterAndUserInfo
	Dim objWMIService, objItem, colItems, strComputer, defaultPrinter
	
	strComputer ="."
	
	' --------------------------------------------
	' Pure WMI Section
	Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\CIMV2")
	Set colItems = objWMIService.ExecQuery("SELECT * FROM Win32_Printer where DeviceID = '" & IMAGENOW_PRINTER_NAME & "'")
	
	' On Error Resume Next
	For Each objItem In colItems
		'If objItem.DeviceID = "inserver6_savepage" Then
		
		If objItem.Attributes And 4 Then
			defaultPrinter = "True"
		Else
			defaultPrinter = "False"
		End If
		
		DebugLog "INFO", "ImageNow Printer Info\n\n" & _
		objItem.name & _
		"\n====================================" & "\n" & _
		"Default : " & defaultPrinter & "\n" & _
		"Availability: " & objItem.Availability & "\n" & _
		"Description: " & objItem.Description & "\n" & _
		"Printer: " & objItem.DeviceID & "\n" & _
		"Driver Name: " & objItem.DriverName & "\n" & _
		"Port Name: " & objItem.PortName & "\n" & _
		"Printer State: " & objItem.PrinterState & "\n" & _
		"Printer Status: " & objItem.PrinterStatus & "\n" & _
		"PrintJobDataType: " & objItem.PrintJobDataType & "\n" & _
		"Print Processor: " & objItem.PrintProcessor & "\n" & _
		"Spool Enabled: " & objItem.SpoolEnabled & "\n" & _
		"Separator File: " & objItem.SeparatorFile & "\n" & _
		"Queued: " & objItem.Queued & "\n" & _
		"Status: " & objItem.Status & "\n" & _
		"StatusInfo: " & objItem.StatusInfo & "\n" & _
		"Published: " & objItem.Published & "\n" & _
		"Shared: " & objItem.Shared & "\n" & _
		"ShareName: " & objItem.ShareName & "\n" & _
		"Direct: " & objItem.Direct & "\n" & _
		"Location: " & objItem.Location & "\n" & _
		"Priority: " & objItem.Priority & "\n" & _
		"Work Offline: " & objItem.WorkOffline & "\n" & _
		"Horizontal Res: " & objItem.HorizontalResolution & "\n" & _
		"Vertical Res: " & objItem.VerticalResolution & "\n" & _
		"User: " & gNetworkUser & "\n\n"
	Next
End Function

' ******************************************************************************
Function KillAllProc(procExe)
	Dim rtn
	
	DebugLog "INFO", "KillAllProc: Killing " & procExe & " Running by " & gNetworkUser & " (if still running)\n"
	
	strComputer = "."
	Set colProcesses = GetObject("winmgmts:" & _
		"{impersonationLevel=impersonate}!\\" & strComputer & _
		"\root\cimv2").ExecQuery("Select * from Win32_Process")
		
	For Each objProcess in colProcesses
		If Trim(UCase(objProcess.Name)) = Trim(UCase(procExe)) Then
			Call DebugLog("DEBUG", "Found " & objProcess.Name & " " & objProcess.ProcessId & "\n")
			Return = objProcess.GetOwner(strNameOfUser)
			If Return <> 0 Then
				DebugLog "ERROR", "Could not get owner info for process " & _
					objProcess.Name & VBNewLine _
					& "Error = " & Return & "\n"
			Else
				If InStr(UCase(gNetworkUser), UCase(strNameOfUser)) Then
					rtn = objProcess.Terminate(0)
					If rtn <> 0 Then
						Call DebugLog("WARNING", "Unable to terminate process via win32_process ("&rtn&"), using taskkill: "&Err.Number&" - "&Err.Description&"\n")
						ExecNoWait( GetINIString("INMAC", "KillCommand", "TASKKILL /F /PID ", INMAC_INI) & CStr(objProcess.ProcessId))
					End If
					
					DebugLog "NOTIFY", "Killed process " & objProcess.ProcessID & " " & objProcess.Name & "\n"
					
				Else
					DebugLog "INFO", "NOT killing Process " _
					& objProcess.Name & " It is owned by " _
					& "\" & strNameOfUser & ".\n"
				End If
			End If
		End If
	Next
End Function

' ******************************************************************************
Function gStatsInc(counter)
	If Not gStats.Exists(counter) Then
		Call gStats.Add(counter, 0)
	End If
	
	gStats.Item(counter) = gStats.Item(counter) + 1
End Function

' ******************************************************************************
Function gStatsAdd(counter, amount)
	If Not gStats.Exists(counter) Then
		Call gStats.Add(counter, 0)
	End If
	
	gStats.Item(counter) = gStats.Item(counter) + amount
End Function

' ******************************************************************************
Function LogStats
	Dim cc
	Dim keys(512)
	cc = 0
	For each key in gStats.keys
		keys(cc) = key
		cc = cc + 1
	Next
	
	On Error Resume Next
	Call QuickSort(keys,0, cc-1)
	
	Call DebugLog("NOTIFY", "\n\n")
	For x = 0 to cc-1
		Call DebugLog("NOTIFY",  keys(x) & ": " & gStats.Item(keys(x)) & "\n")
		Call WriteINIString("INMAC Stats", keys(x), CStr(CLng(GetINIString("INMAC Stats", keys(x), "0", STATS_INI) + CLng(gStats.Item(keys(x))))), STATS_INI)
	Next
	Call DebugLog("NOTIFY", "\n\n")
End Function

' ******************************************************************************
Sub QuickSort(vec,loBound,hiBound)
	Dim pivot,loSwap,hiSwap,temp
	
	'== This procedure is adapted from the algorithm given in:
	'==    Data Abstractions & Structures using C++ by
	'==    Mark Headington and David Riley, pg. 586
	'== Quicksort is the fastest array sorting routine for
	'== unordered arrays.  Its big O is  n log n
	
	'== Two items to sort
	if hiBound - loBound = 1 then
		if vec(loBound) > vec(hiBound) then
			temp=vec(loBound)
			vec(loBound) = vec(hiBound)
			vec(hiBound) = temp
		End If
	End If
	
	'== Three or more items to sort
	pivot = vec(int((loBound + hiBound) / 2))
	vec(int((loBound + hiBound) / 2)) = vec(loBound)
	vec(loBound) = pivot
	loSwap = loBound + 1
	hiSwap = hiBound
	
	do
		'== Find the right loSwap
		while loSwap < hiSwap and vec(loSwap) <= pivot
			loSwap = loSwap + 1
		wend
		
		'== Find the right hiSwap
		while vec(hiSwap) > pivot
			hiSwap = hiSwap - 1
		wend
		
		'== Swap values if loSwap is less then hiSwap
		if loSwap < hiSwap then
			temp = vec(loSwap)
			vec(loSwap) = vec(hiSwap)
			vec(hiSwap) = temp
		End If
	loop while loSwap < hiSwap
	
	vec(loBound) = vec(hiSwap)
	vec(hiSwap) = pivot
	
	'== Recursively call function .. the beauty of Quicksort
	'== 2 or more items in first section
	if loBound < (hiSwap - 1) then Call QuickSort(vec,loBound,hiSwap-1)
	
	'== 2 or more items in second section
	if hiSwap + 1 < hibound then Call QuickSort(vec,hiSwap+1,hiBound)
End Sub  'QuickSort

' ******************************************************************************
Function VerifyPath(fileP)
	If InStrRev(fileP,"\") <> Len(fileP) And InStrRev(fileP,"/") <> Len(fileP) Then
		fileP = fileP & "\"
	End If
	VerifyPath = fileP
End Function

' ******************************************************************************
Function GetText(row, col, Len)
	txtFile = GetFile(f1)
	readFile = split(txtFile,vbCrLf)
	GetText = Mid(readFile(row-1),col,Len)
End Function

' ******************************************************************************
Sub WriteAppendFile(fileDirectory, iValues)
	Set filesys = CreateObject("Scripting.FileSystemObject")
	Set filetxt = filesys.CreateTextFile(fileDirectory, True)
	filetxt.WriteLine (iValues)
	filetxt.Close
End Sub

' ******************************************************************************
Function WriteTextFile(iValues,outputPath,file,iniFile,fileTag)
	difPath = GetINIString("ADVANCED SETTINGS","format.file.path","",iniPath)
	If difPath <> "" Then
		pathToWrite = VerifyPath(difPath)
	Else
		pathToWrite = VerifyPath(outputPath)
	End If
	fileSeparator = GetINIString("SAVE","index.value.separator","-",iniFile)
	workFlowQ = GetINIString(fileTag,"workflow.queue","",iniFile)
	tifFileName = GetUniqueID()
	dim filesys, filetxt, getname, path
	MoveTifFile file , tifFileName,outputPath
	Set filesys = CreateObject("Scripting.FileSystemObject")
	Dim strNewFileName:strNewFileName = outputPath & GetUniqueID()&".txt"
	Set filetxt = filesys.CreateTextFile(strNewFileName, True)
	DebugLog "DEBUG", "WriteTextFile: Created file: [" & strNewFileName & "]" & vbCrLf
	path = filesys.GetAbsolutePathName(outputPath & GetUniqueID() &".txt")
	getname = filesys.GetFileName(path)
	filetxt.WriteLine(iValues & fileSeparator & "1" & fileSeparator & pathToWrite & tifFileName & fileSeparator & workFlowQ)
	filetxt.Close
	WriteTextFile = true
End Function

' ******************************************************************************
Function FindINI()
	Set iniFile = CreateObject("Scripting.FileSystemObject")
	getINIFolderName = iniFile.GetParentFolderName(WScript.ScriptFullName)
	If iniFile.FileExists(getINIFolderName & "\" & iniFileName) Then
		FindINI = getINIFolderName
	Else
		MsgBox "FindINI: Unable to locate ["&iniFileName&"] file.  File must be located in same directory as ["&WScript.ScriptName&"].",,"Error - ["&WScript.ScriptFullName&"]"
	End If
End Function

' ******************************************************************************
Function GetUniqueID()
	milliTime = Timer()
	milliTime = milliTime/60
	milliTime = milliTime/90
	milliTime = cstr(milliTime*100000000000420)
	milliTime = "100"& replace(milliTime,".","")
	milliTime = replace(milliTime,"+","")
	GetUniqueID = milliTime
End Function

' ******************************************************************************
Sub WriteINIStringVirtual(Section, KeyName, Value, FileName)
	WriteINIString Section, KeyName, Value, Server.MapPath(FileName)
End Sub

' ******************************************************************************
Function GetINIStringVirtual(Section, KeyName, Default, FileName)
	GetINIStringVirtual = GetINIString(Section, KeyName, Default, Server.MapPath(FileName))
End Function

' ******************************************************************************
Sub ClearINICounters(Section, FileName)
	Dim INIContents, PosSection, PosEndSection
	
	'Get contents of the INI file As a string
	INIContents = GetFile(FileName)
	
	'Find section
	PosSection = InStr(1, INIContents, "[" & Section & "]", vbTextCompare)
	If PosSection>0 Then
		'Section exists. Find End of section
		PosEndSection = InStr(PosSection, INIContents, vbCrLf & "[")
		'?Is this last section?
		If PosEndSection = 0 Then PosEndSection = Len(INIContents)+1
			'Separate section contents
			Dim OldsContents, NewsContents, Line
			Dim sKeyName, Found
			OldsContents = Mid(INIContents, PosSection, PosEndSection - PosSection)
			OldsContents = split(OldsContents, vbCrLf)
			
			'Temp variable To find a Key
			sKeyName = LCase(KeyName & "=")
			
			'Enumerate section lines
			For Each Line In OldsContents
				If InStr(1, Line, "=", vbTextCompare) > 0 Then
					Line = Left(Line, InStr(1, Line, "=", vbTextCompare)) & "0"
				End If
				NewsContents = NewsContents & Line & vbCrLf
			Next
			
			'Combine pre-section, new section And post-section data.
			NewsContents = Left(NewsContents, Len(NewsContents) - 2)
			INIContents = Left(INIContents, PosSection-1) & _
			NewsContents & Mid(INIContents, PosEndSection)
		Else'If PosSection>0 Then
			'Section Not found. Add section data at the End of file contents.
			If Right(INIContents, 2) <> vbCrLf And Len(INIContents)>0 Then
				INIContents = INIContents & vbCrLf
			End If
			INIContents = INIContents & "[" & Section & "]" & vbCrLf & _
			KeyName & "=" & Value
		End If'If PosSection=0 Then
	Dim FS: Set FS = CreateObject("Scripting.FileSystemObject")
	
	'Go To windows folder If full path Not specified.
	If InStr(FileName, ":\") = 0 And Left (FileName,2)<>"\\" Then
		FileName = FS.GetSpecialFolder(0) & "\" & FileName
	End If
	
	Dim OutStream: Set OutStream = FS.OpenTextFile(FileName, 2, True)
	OutStream.Write INIContents
End Sub

' ******************************************************************************
Sub WriteINIString(Section, KeyName, Value, FileName)
	Dim INIContents, PosSection, PosEndSection
	
	'Get contents of the INI file As a string
	INIContents = GetFile(FileName)
	
	'Find section
	PosSection = InStr(1, INIContents, "[" & Section & "]", vbTextCompare)
	If PosSection>0 Then
		'Section exists. Find End of section
		PosEndSection = InStr(PosSection, INIContents, vbCrLf & "[")
		'?Is this last section?
		If PosEndSection = 0 Then PosEndSection = Len(INIContents)+1
		
		'Separate section contents
		Dim OldsContents, NewsContents, Line
		Dim sKeyName, Found
		OldsContents = Mid(INIContents, PosSection, PosEndSection - PosSection)
		OldsContents = split(OldsContents, vbCrLf)
		
		'Temp variable To find a Key
		sKeyName = LCase(KeyName & "=")
		
		'Enumerate section lines
		For Each Line In OldsContents
			If LCase(Left(Line, Len(sKeyName))) = sKeyName Then
				Line = KeyName & "=" & Value
				Found = True
			End If
			NewsContents = NewsContents & Line & vbCrLf
		Next
		If isempty(Found) Then
			'key Not found - add it at the End of section
			NewsContents = NewsContents & KeyName & "=" & Value
		Else
			'remove last vbCrLf - the vbCrLf is at PosEndSection
			NewsContents = Left(NewsContents, Len(NewsContents) - 2)
		End If
		
		'Combine pre-section, new section And post-section data.
		INIContents = Left(INIContents, PosSection-1) & _
		NewsContents & Mid(INIContents, PosEndSection)
		Else'If PosSection>0 Then
		'Section Not found. Add section data at the End of file contents.
		If Right(INIContents, 2) <> vbCrLf And Len(INIContents)>0 Then
			INIContents = INIContents & vbCrLf
		End If
		INIContents = INIContents & "[" & Section & "]" & vbCrLf & _
		KeyName & "=" & Value
	End If'If PosSection>0 Then
	Dim FS: Set FS = CreateObject("Scripting.FileSystemObject")
	
	'Go To windows folder If full path Not specified.
	If InStr(FileName, ":\") = 0 And Left (FileName,2)<>"\\" Then
		FileName = FS.GetSpecialFolder(0) & "\" & FileName
	End If
	
	Dim OutStream: Set OutStream = FS.OpenTextFile(FileName, 2, True)
	OutStream.Write INIContents
End Sub

' ******************************************************************************
Function GetINIString(Section, KeyName, Default, FileName)
	Dim INIContents, PosSection, PosEndSection, sContents, Value, Found
	
	'Get contents of the INI file As a string
	INIContents = GetFile(FileName)
	
	'Find section
	PosSection = InStr(1, INIContents, "[" & Section & "]", vbTextCompare)
	If PosSection>0 Then
		'Section exists. Find End of section
		PosEndSection = InStr(PosSection, INIContents, vbCrLf & "[")
		'?Is this last section?
		If PosEndSection = 0 Then PosEndSection = Len(INIContents)+1
		'Separate section contents
		sContents = Mid(INIContents, PosSection, PosEndSection - PosSection)
		
		If InStr(1, sContents, vbCrLf & KeyName & "=", vbTextCompare)>0 Then
			Found = True
			'Separate value of a key.
			Value = SeparateField(sContents, vbCrLf & KeyName & "=", vbCrLf)
		End If
	End If
	If isempty(Found) Then Value = Default
	GetINIString = Value
End Function

' ******************************************************************************
' Separates one field between sStart And sEnd
Function SeparateField(ByVal sFrom, ByVal sStart, ByVal sEnd)
	Dim PosB: PosB = InStr(1, sFrom, sStart, 1)
	If PosB > 0 Then
		PosB = PosB + Len(sStart)
		Dim PosE: PosE = InStr(PosB, sFrom, sEnd, 1)
		If PosE = 0 Then PosE = InStr(PosB, sFrom, vbCrLf, 1)
		If PosE = 0 Then PosE = Len(sFrom) + 1
		SeparateField = Mid(sFrom, PosB, PosE - PosB)
	End If
End Function

' ******************************************************************************
Function GetFile(ByVal FileName)
	Dim FS: Set FS = CreateObject("Scripting.FileSystemObject")
	'Go To windows folder If full path Not specified.
	If InStr(FileName, ":\") = 0 And Left (FileName,2)<>"\\" Then
		FileName = FS.GetSpecialFolder(0) & "\" & FileName
	End If
	On Error Resume Next
	GetFile = FS.OpenTextFile(FileName).ReadAll
	'FIXME: Just ignore this error?
	On Error Goto 0
End Function

' ******************************************************************************
Function GetFile(ByVal FileName)
	Dim FS: Set FS = CreateObject("Scripting.FileSystemObject")
	'Go To windows folder If full path Not specified.
	If InStr(FileName, ":\") = 0 And Left (FileName,2)<>"\\" Then
		FileName = FS.GetSpecialFolder(0) & "\" & FileName
	End If
	On Error Resume Next
	GetFile = FS.OpenTextFile(FileName).ReadAll
	'FIXME: Just ignore this error?
	On Error Goto 0
End Function

' ******************************************************************************
Function WriteFile(ByVal FileName, ByVal Contents)
	Dim FS: Set FS = CreateObject("Scripting.FileSystemObject")
	
	'Go To windows folder If full path Not specified.
	If InStr(FileName, ":\") = 0 And Left (FileName,2)<>"\\" Then
		FileName = FS.GetSpecialFolder(0) & "\" & FileName
	End If
	
	Dim OutStream: Set OutStream = FS.OpenTextFile(FileName, 8, True)
	OutStream.Write Contents
	OutStream.Close
End Function

' ******************************************************************************
Function GetDateYYYYMMDD()   ' added 11/23 - TP
	GetDateYYYYMMDD = GetYear4Digit & GetMonth2Digit & GetDay2Digit
End Function

' ******************************************************************************
Function GetDay2Digit()   ' added 11/23 - TP
	todayDayNum = Day(now)
	If (todayDayNum < 10) Then
		todayDayStr = "0" & CStr(todayDayNum)
	Else
		todayDayStr = CStr(todayDayNum)
	End If
	
	GetDay2Digit = todayDayStr
End Function

' ******************************************************************************
Function GetMonth2Digit()   ' added 11/23 - TP
	todayMonthNum = Month(now)
	If (todayMonthNum < 10) Then
		todayMonthStr = "0" & CStr(todayMonthNum)
	Else
		todayMonthStr = CStr(todayMonthNum)
	End If
	
	GetMonth2Digit = todayMonthStr
End Function

' ******************************************************************************
Function GetYear4Digit()   ' added 11/23 - TP
	GetYear4Digit = CStr(Year(now))
End Function

' ******************************************************************************
' Splits a filename into three pieces: dir,name,ext
' @param {String} strFilePath
' @return {String[]} null on error
' ******************************************************************************
Function SplitFilename(strFilePath)
	Dim arrPieces(3) '0=dir, 1=name, 2=ext
	Dim rgxSplitFileName, colMatches
	Set rgxSplitFileName = New RegExp
	rgxSplitFileName.Pattern = "^(.*\\)([^.]+)(.*)$"
	Set colMatches = rgxSplitFileName.Execute(strFilePath)
	If colMatches.Count = 0 Then
		DebugLog "ERROR", "SplitFileName: Unable to split file: ["&strFilepath&"]" & vbCrLf
		SplitFilename = Null
	Else
		arrPieces(1) = colMatches.Item(0).SubMatches.Item(0)
		arrPieces(2) = colMatches.Item(0).SubMatches.Item(1)
		arrPieces(3) = colMatches.Item(0).SubMatches.Item(2)
		SplitFilename = arrPieces
	End If
End Function

' ******************************************************************************
' @return {int}
' ******************************************************************************
Function GetEpochMS()
	Dim dblSecondsSinceMidnight, intDaysSince2007, dte2007
	Const intMilliSecondsSinceEpochTill2007 = 1167631200000
	Const intMilliSecondsPerDay = 86400000
	dte2007 = CDate("January 1, 2007")
	dblSecondsSinceMidnight = Timer() 'Seconds since midnight
	intDaysSince2007 = DateDiff("d", dte2007, Now)
	GetEpochMS = intMilliSecondsSinceEpochTill2007 + (intDaysSince2007*intMilliSecondsPerDay) + Fix(dblSecondsSinceMidnight*1000)
End Function

' ******************************************************************************
' Open output file
' ******************************************************************************
Sub DebugNew()
	Dim arrSplitFileName, strUniqueTag
	If gBlnDebugIsOpen = False Then
		gStrDebugScriptName = WScript.ScriptName
		arrSplitFileName = SplitFileName(WScript.ScriptFullName)
		If IsNull(arrSplitFileName) Then
			MsgBox "DebugNew: Unable to split filename: ["&WScript.ScriptFullName&"]","Error - ["&WScript.ScriptFullName&"]"
			WScript.Quit(1)
		End If
		gStrDebugScriptName = arrSplitFileName(2)
		gStrLogDir = arrSplitFileName(1)
		If Len(INMAC_LOG_BASE_DIR) > 0 Then
			gStrLogDir = INMAC_LOG_BASE_DIR
			
			'automatically create log directory if necessary
			If gFsObj.FolderExists(gStrLogDir) = False Then
				gFsObj.CreateFolder(gStrLogDir)
			End If
		End If
		gStrDebugFilePath = gStrLogDir & "INMAC_" & IMAGENOW_PRINTER_NAME & "_" & GetDateYYYYMMDD() & ".log"
		On Error Resume Next
		Set gObjDebugFile = gFsObj.OpenTextFile(gStrDebugFilePath, 8, True, 0)
		If Err.Number <> 0 Then
			'MsgBox "DebugNew: Unable to open log file: ["&gStrDebugFilePath&"]" & vbCrLf & " Error: ["&Err.Number&"] ["&Err.Description&"]"
			SendAdminEmail("Could not open log file.")
		End If
		On Error Goto 0
		strUniqueTag = IntToString(Modulo(gIntDebugStartTime, 36*36), 36)
		If Len(strUniqueTag) = 1 Then
			strUniqueTag = "0" & strUniqueTag
		End If
		gStrDebugLineHeader = " " & gStrDebugScriptName & ":" & strUniqueTag & " "
		gBlnDebugIsOpen = True
	End If
End Sub

' ******************************************************************************
Function Modulo(intDividend, intDivisor)
	Dim intRemainingValue, intQuotient, intRemainder, strResult
	intQuotient = Fix(Fix(intDividend) / Fix(intDivisor))
	Modulo = intDividend - (intQuotient * intDivisor)
End Function

' ******************************************************************************
' Convert an Integer to a String
' @param {int} intValue
' @param {int} intRadix
' @return {String}
' ******************************************************************************
Function IntToString(intValue, intRadix)
	IntToString = "Error"
	Dim char(36), i, intRemainder
	For i=0 To 9
		char(i) = Chr(i+48)
	Next
	For i=0 To 26
		char(i+10) = Chr(i+65)
	Next
	If intRadix > 36 OR intRadix < 2 Then
		LogConsole "ERROR", "IntToString: Invalid radix: ["&intRadix&"]"
		Exit Function
	ElseIf intRadix = 10 Then
		IntToString = CStr(intValue)
		Exit Function
	End If
	IntToString = ""
	intValue = Fix(intValue)
	Do
		intRemainder = Modulo(intValue, intRadix)
		intValue = (intValue - intRemainder) / intRadix
		IntToString = char(intRemainder) & IntToString
	Loop While CStr(intValue) <> "0"
End Function

' ******************************************************************************
' Close output file
' ******************************************************************************
Sub DebugFinish()
	DebugLog "INFO", "DebugFinish: Duration: ["&FormatNumber(((GetEpochMS())-gIntDebugStartTime)/1000,2)&"] secs" & vbCrLf
	DebugLog "NOTIFY", "DebugFinish: Return code: ["&gScriptReturnCode&"]" & vbCrLf
	If gBlnDebugIsOpen = True Then
		If gBlnDebugHasWritten Then
			gObjDebugFile.Write gStrDebugFooter
		End If
		gObjDebugFile.Close()
		gBlnDebugIsOpen = False
	End If
End Sub

' ******************************************************************************
' Similar to iScriptDebug log (only that it logs to console) gIntDebugLevel controls output level
' @param {String|int} Level
' @param {String} Contents
' ******************************************************************************
Sub DebugLog(ByRef Level, ByRef Contents)
	Dim strOutput, intLevel
	If gIntDebugLevel < 0 OR NOT gBlnDebugIsOpen Then
		Exit Sub
	End If
	strOutput = "" & Contents
	
	' replace \n with vbCrLf
	strOutput = Replace(strOutput, "\n", "" & vbCrLf)
	
	If NOT IsNumeric(Level) Then
		Select Case Level
			Case "RAW" 'Special instruction
				gObjDebugFile.Write Contents
				Exit Sub
			Case "DEBUG"
				intLevel = 5
			Case "INFO"
				intLevel = 4
			Case "NOTIFY"
				intLevel = 3
			Case "WARNING"
				intLevel = 2
			Case "ERROR"
				intLevel = 1
			Case "CRITICAL"
				intLevel = 0
			Case Else
				intLevel = -1
		End Select
	Else
		intLevel = CInt(Level)
	End If
	If gIntDebugLevel >= intLevel Then
		If gBlnDebugHasWritten = False Then
			gObjDebugFile.Write gStrDebugHeader
			gBlnDebugHasWritten = True
		End If
		Select Case intLevel
			Case 5
				strOutput = "[  DEBUG ] " & strOutput
			Case 4
				strOutput = "[  INFO  ] " & strOutput
			Case 3
				strOutput = "[ NOTIFY ] " & strOutput
			Case 2
				strOutput = "[ WARNING] " & strOutput
			Case 1
				strOutput = "[  ERROR ] " & strOutput
			Case 0
				strOutput = "[CRITICAL] " & strOutput
			Case Else
				strOutput = "[ ?????? ] " & strOutput
		End Select
		gObjDebugFile.Write GetTimeStamp() & gStrDebugLineHeader& strOutput
	End If
End Sub

' ******************************************************************************
' Return the current time in the format MM/dd HH:mm:ss.SSS
' @return {String}
' ******************************************************************************
Function GetTimeStamp()
	Dim intYear,intMonth,intDay,intHour,intMinute,intSecond,intMS,dte,tmr
	dte = Date
	tmr = Time
	intYear   = Year(dte)
	intMonth  = Month(dte)
	intDay    = Day(dte)
	intHour   = Hour(tmr)
	intMinute = Minute(tmr)
	intSecond = Second(tmr)
	intMS     = (Timer()*1000) mod 1000
	GetTimeStamp = ""
	If intMonth < 10 Then
		GetTimeStamp = GetTimeStamp & "0"
	End If
	GetTimeStamp = GetTimeStamp & intMonth & "/"
	If intDay < 10 Then
		GetTimeStamp = GetTimeStamp & "0"
	End If
	GetTimeStamp = GetTimeStamp & intDay & " "
	If intHour < 10 Then
		GetTimeStamp = GetTimeStamp & "0"
	End If
	GetTimeStamp = GetTimeStamp & intHour & ":"
	If intMinute < 10 Then
		GetTimeStamp = GetTimeStamp & "0"
	End If
	GetTimeStamp = GetTimeStamp & intMinute & ":"
	If intSecond < 10 Then
		GetTimeStamp = GetTimeStamp & "0"
	End If
	GetTimeStamp = GetTimeStamp & intSecond & "."
	If intMS < 10 Then
		GetTimeStamp = GetTimeStamp & "00"
	ElseIf intMS < 100 Then
		GetTimeStamp = GetTimeStamp & "0"
	End If
	GetTimeStamp = GetTimeStamp & intMS
End Function

' ******************************************************************************
Function GetINIString(Section, KeyName, Default, FileName)
	Dim INIContents, PosSection, PosEndSection, sContents, Value, Found
	
	'Get contents of the INI file As a string
	INIContents = GetFile(FileName)
	
	'Find section
	PosSection = InStr(1, INIContents, "[" & Section & "]", vbTextCompare)
	If PosSection>0 Then
		'Section exists. Find End of section
		PosEndSection = InStr(PosSection, INIContents, vbCrLf & "[")
		'?Is this last section?
		If PosEndSection = 0 Then PosEndSection = Len(INIContents)+1
		'Separate section contents
		sContents = Mid(INIContents, PosSection, PosEndSection - PosSection)

		If InStr(1, sContents, vbCrLf & KeyName & "=", vbTextCompare)>0 Then
			Found = True
			'Separate value of a key.
			Value = SeparateField(sContents, vbCrLf & KeyName & "=", vbCrLf)
		End If
	End If
	If isempty(Found) Then Value = Default
	GetINIString = Value
End Function

' ******************************************************************************
Function GetINIKeysInSection(Section, FileName)
	GetINIKeysInSection = ""
	'Get contents of the INI file As a string
	INIContents = GetFile(FileName)
	
	'Find section
	PosSection = InStr(1, INIContents, "[" & Section & "]", vbTextCompare)
	If PosSection>0 Then
		'Section exists. Find End of section
		PosEndSection = InStr(PosSection, INIContents, vbCrLf & "[")
		'?Is this last section?
		If PosEndSection = 0 Then PosEndSection = Len(INIContents)+1
		'Separate section contents
		sContents = Mid(INIContents, PosSection, PosEndSection - PosSection)
		
		aLines = split(sContents,vbCrLf)
		For x = 0 to UBound(aLines)
			If Not ( _
				(InStr(1, aLines(x), ";") = 1) OR _
				(InStr(1, aLines(x), " ") = 1) OR _
				(Len(aLines(x)) = 0) OR _
				(InStr(1, aLines(x), "=") < 1) ) Then
					GetINIKeysInSection = GetINIKeysInSection & Left(aLines(x), InStr(1, aLines(x), "=")-1) & "^"
			End If
		Next
	End If
End Function
' *********************************************************************************
Function ReportFileStatus(filespec)
   Dim fso, msg
   Set fso = CreateObject("Scripting.FileSystemObject")
   If (fso.FileExists(filespec)) Then
      msg = filespec & " exists."
   Else
      msg = filespec & " doesn't exist."
   End If
   ReportFileStatus = msg
End Function
'