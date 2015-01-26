# #############################################################################
# NAME: MoveFromOldShare.ps1
# 
# AUTHOR:  Gregg Jenczyk, UMass (UITS)
# DATE:  2014/07/09
# EMAIL: gjenczyk@umassp.edu
# 
# COMMENT:  This script is a template for running intool scripts via 
#           PowerShell.  You can use this on demand or to set up scheduled
#           jobs.  Just be sure to save this as something else.
#
# VERSION HISTORY
# 1.0 2014.07.09 Initial Version.
#
# TO ADD:
# 
# USEFUL SNIPPETS:
# "$(Get-Date) " | Out-File -Append $runLog
# #############################################################################

#-- INCLUDES --#
. "\\boisnas215c1.umasscs.net\diimages67tst\script\PowerShell\sendmail.ps1"
. "\\boisnas215c1.umasscs.net\diimages67tst\script\PowerShell\enVar.ps1"

#-- CONFIG --#

$scriptName = 'MoveFromOldShare'
$oldShare = $env.Substring(0,$env.Length-2)
$filesTransferred = $false

#-- MAIN --#

if (Test-Path ${shareRoot}import_agent\DI_${oldShare}_SA_AD_INBOUND\*.*) {

    $filesTransferred = $true

    Get-ChildItem -path "${shareRoot}import_agent\DI_${oldShare}_SA_AD_INBOUND\" -File | ForEach-Object {

        echo $_.FullName

        Move-Item $_.FullName -Destination "${shareRoot}import_agent\DI_${env}_SA_AD_INBOUND\"
 
    } 
}

if ($filesTransferred) {

    $message = "Files have been moved from $oldShare to $env"

    sendmail -t "gjenczyk@umassp.edu" -s "[DI ${env} Notice] ${scriptName}.ps1 has moved files" -m ${message}

}

$error[0] | Out-File $runLog -Append