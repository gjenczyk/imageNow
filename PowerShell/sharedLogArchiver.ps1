<# #############################################################################
# NAME: sharedLogArchiver.ps1
# 
# AUTHOR:  Gregg Jenczyk, UMass (UITS)
# DATE:  2014/05/16
# EMAIL: gjenczyk@umassp.edu
# 
# COMMENT:  This script is used to zip up the previous day's logs until we get the
#           monitor agent sorted out.  This version handles the shared log dir
#
# VERSION HISTORY
# 1.0 2014.05.16 Initial Version.
#
# TO ADD:
# 
# USEFUL SNIPPETS:
# "$(Get-Date) " | Out-File -Append $runLog
# #############################################################################>

#-- INCLUDES --#
. "\\ssisnas215c2.umasscs.net\diimages67prd\script\PowerShell\sendmail.ps1"

#-- CONFIGURATION --#

$shareRoot = "\\ssisnas215c2.umasscs.net\diimages67prd\"
$shareLog = "${shareRoot}log\"
$shareAudit = "${shareRoot}audit\"
$runLog = "${shareLog}run_log-logArchiver.log"
$yesterday = (Get-Date).AddDays(0).Date

#- Date formatting -#
$la = (Get-Date -Format yyyy.MM.dd.hh.mm.ss)
$l = [regex]::Split($la, '\.')
$y = $l[0]
$m = $l[1]
$d = $l[2]
$h = $l[3]
$mi = $l[4]
$s = $l[5]

#-- FUNCTION --#

function ZipFile( $zipfilename, $sourcedir )
{
    [Reflection.Assembly]::LoadWithPartialName( "System.IO.Compression.FileSystem" )
    $compressionLevel = [System.IO.Compression.CompressionLevel]::Optimal
    [System.IO.Compression.ZipFile]::CreateFromDirectory( $sourcedir, $zipfilename, $compressionLevel, $false )
}


$zipBase = "${shareLog}Archive_log_$y$m$d`_$h.$mi.$s"
$zip = "${zipBase}.zip"

New-Item -Path $zipBase -ItemType Directory

$files = Get-ChildItem $shareLog, $shareAudit -Exclude "Archive*" -Recurse -af | Where-Object {$_.LastWriteTime.Date -lt $yesterday}

foreach($file in $files){

   Move-Item $file -Destination $zipBase
}

ZipFile $zip $zipBase

Remove-Item -Path $zipBase -Recurse -Force