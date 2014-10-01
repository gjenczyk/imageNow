<# #############################################################################
# NAME: QueueContentEmail.ps1
# 
# AUTHOR:  Gregg Jenczyk, UMass (UITS)
# DATE:  2014/04/11
# EMAIL: gjenczyk@umassp.edu
# 
# COMMENT:  This script sends emails w/ attachments for QueueContentResolution.js
#            
# VERSION HISTORY
# 1.0 2014.04.11 Initial Version.
# 2.0 2014.08.04 Added campus specific email functionality
#
# TO ADD:
#
# USEFUL SNIPPETS:
# "$(Get-Date) " | Out-File -Append $runLog
# #############################################################################>

#-- CONFIGURATION --#

#- PARAMETERS -#

Param([string]$csvPath, 
      [string]$docPath, 
      [string]$errorReason
      )

#- INCLUDES -#
. "\\ssisnas215c2.umasscs.net\diimages67prd\script\PowerShell\sendmail.ps1"

#- VARIABLES -#
$log = '\\ssisnas215c2.umasscs.net\diimages67prd\log\email.log'
$csvPath | Out-File  $log -Append
$docPath | Out-File $log -Append
$errorReason | Out-File $log -Append

$env = ([environment]::MachineName).Substring(2)
$env = $env -replace "W.*",""
$root = "Y:\"
$csvName = $csvPath -replace ".*\\",""
$docName = $docPath -replace ".*\\",""
$importType = $docName -replace "_.*",""
$successDir = $docPath -replace "[^\\]*$",""

#- Email Selector -#
$bostonGA = ("UITS.DI.CORE@umassp.edu","Peggy.Patel@umb.edu", "mary.ryan@umb.edu")
$bostonUA = ("UITS.DI.CORE@umassp.edu","john.drew@umb.edu", "Lisa.Williams@umb.edu", "mary.ryan@umb.edu")
$dartmouthGA = ("UITS.DI.CORE@umassp.edu","graduate@umassd.edu","j1mello@umassd.edu")
$dartmouthUA = ("UITS.DI.CORE@umassp.edu","Athompson@umassd.edu", "cgoodine@umassd.edu", "j1mello@umassd.edu")
$lowellGA = ("UITS.DI.CORE@umassp.edu", "linda_southworth@uml.edu", "Daniel_Bedard@uml.edu", "Barbara_Dougherty@uml.edu")
$lowellUA = ("UITS.DI.CORE@umassp.edu", "Kathleen_Shannon@uml.edu", "Dania_Valdes@uml.edu", "Daniel_Bedard@uml.edu")
$DISupport = ("UITS.DI.CORE@umassp.edu")

function campusSwitch ([string] $emailOffice) {
    #selects the correct target emails based on campus and career
    switch ($emailOffice)
    {
        "UMBOS GRAD" {$bostonGA}
        "UMBOS UGRD" {$bostonUA}
        "UMDAR GRAD" {$dartmouthGA}
        "UMDAR UGRD" {$dartmouthUA}
        "UMLOW GRAD" {$lowellGA}
        "UMLOW UGRD" {$lowellUA}
        "UMLOW CSCE" {$lowellUA}
        default {$DISupport}
    }
}


#-- MAIN --#

#- Extract information from file -#
if ($importType -eq "WADOC") {
    Get-Content $csvPath | ForEach-Object {
        $fileInfo = [regex]::split($_, '\^')
    }

    $campus = $fileInfo[1]
    $career = $fileInfo[2]
    $appCenter = $fileInfo[3]
    $emplid = $fileInfo[4]
    $name = $fileInfo[5]

	$customMessage = "The attached file belonging to $name ($emplid) failed to be imported into DI for the following reason:
	-
	${errorReason}"
} elseif ($importType -eq "CA") {
    $fileInfo = [regex]::split($docName, '_')
    $campus = $fileInfo[2]
    switch ($campus) {
        B {$campus = "UMBOS"; break}
        D {$campus = "UMDAR"; break}
        L {$campus = "UMLOW"; break}
        Default {"ERROR"; break}
    }
    $appCenter = "UGRD"
    $customMessage="The attached file from Common App for $campus failed to be imported into DI for the following reason:
	-
	${errorReason}"
} elseif ($importType.Substring(0,4) -eq "zadr") {
    $fileInfo = [regex]::Split($csvName,'_|\.')
    $campus = $fileInfo[3]
    $appCenter = $fileInfo[4]
    $customMessage="The attached $importType file belonging to $campus ($appCenter) failed to be imported into DI for the following reason:
	-
	${errorReason}"
} else {
    $campus="Unknown"
	$customMessage="The attached document failed to be imported for the following reason:
	-
	${errorReason}
	-
	Furthermore, this document does not comply with the naming conventions necessary for 
	indexing to be successful."
}

#- Build mail message -#
$messageTitle = "[DI $env ERROR] $campus $appCenter $importType Import Failure" 

$defaultMessage="

The original failed import has been archived.

Best Regards,

INMAC
------------------------
This is an automated message. PLEASE DO NOT REPLY TO THIS MESSAGE. 
Questions can be sent to UITS.DI.CORE@umassp.edu"

$messageBody = "$customMessage
$defaultMessage"

#- Send email w/ attachments -#
[string[]]$attachments
if ($docPath -ne ''){
    $attachments+=@("$docPath")
}
if ($csvPath -ne ''){
    $attachments+=@("$csvPath")
}

$combo = $campus + " " + $appCenter

sendmail -t  $(campusSwitch($combo)) -s "$messageTitle" -a @($attachments) -m $messageBody

$error[0] | Out-File $log -Append