<# #############################################################################
# NAME: logArchiver.ps1
# 
# AUTHOR:  Gregg Jenczyk, UMass (UITS)
# DATE:  2014/05/16
# EMAIL: gjenczyk@umassp.edu
# 
# COMMENT:  This script is used to zip up the previous day's logs until we get the
#           monitor agent sorted out.
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
. "\\boisnas215c1.umasscs.net\diimages67tst\script\PowerShell\sendmail.ps1"
. "\\boisnas215c1.umasscs.net\diimages67tst\script\PowerShell\enVar.ps1"

#-- CONFIGURATION --#

$localRoot = "D:\"
$inserver6 = "${localRoot}inserver6\"
$localLog = "${inserver6}log\"
$runLog = "${localLog}run_log-logArchiver.log"
$yesterday = (Get-Date).AddDays(0).Date

$fatalFlag = $false
$fatalFileArray = @()


#- Date formatting -#
$la = (Get-Date -Format yyyy.MM.dd.hh.mm.ss)
$l = [regex]::Split($la, '\.')
$y = $l[0]
$m = $l[1]
$d = $l[2]
$h = $l[3]
$mi = $l[4]
$s = $l[5]

$zipBase = "${localLog}Archive_log_$y$m$d`_$h.$mi.$s"
$zip = "${zipBase}.zip"

#-- FUNCTION --#

function ZipFile( $zipfilename, $sourcedir )
{
    [Reflection.Assembly]::LoadWithPartialName( "System.IO.Compression.FileSystem" )
    $compressionLevel = [System.IO.Compression.CompressionLevel]::Optimal
    [System.IO.Compression.ZipFile]::CreateFromDirectory( $sourcedir, $zipfilename, $compressionLevel, $false )
}

function ErrorCheck( $file )
{
    if(Select-String -Path $file -Pattern "Fatal iScript Error!")
    {
        $script:fatalFlag = $true  
        $script:fatalFileArray += $file.Name+"`n"
    }
}

function EmailNotify
{
    $body = "The following log files on $machine contained fatal errors:`n $fatalFileArray The logs have been archived in $zip.  `nPlease review to see if corrective action is required."
    sendmail -to "gjenczyk@umassp.edu" -s "[DI $env Warning] Fatal iScript Errors Detected on $machine" -m $body
}


#-- MAIN --#

New-Item -Path $zipBase -ItemType Directory

$files = Get-ChildItem $localLog -Exclude "run_log*","Archive*" -Recurse -af | Where-Object {$_.LastWriteTime.Date -lt $yesterday}

foreach($file in $files){

   ErrorCheck($file)

   Move-Item $file -Destination $zipBase
}

ZipFile $zip $zipBase

Remove-Item -Path $zipBase -Recurse -Force

"fatal flag = $fatalFlag" | Out-File -Append $runLog

if ($fatalFlag)
{
   EmailNotify
}
