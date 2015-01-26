# #############################################################################
# NAME: unZip.ps1
# 
# AUTHOR:  Gregg Jenczyk, UMass (UITS)
# DATE:  2013/12/11
# EMAIL: gjenczyk@umassp.edu
# 
# COMMENT:  This scriptlet unzips files. 
#
# VERSION HISTORY
# 1.0 2013.12.11 Initial Version.
#
# TO ADD:
#
# USEFUL SNIPPETS:
# "$(Get-Date) " | Out-File -Append $runLog
# #############################################################################

function unZip ($file, $destination) {

    Add-Type -As System.IO.Compression.FileSystem

    $ZipFile = Get-Item $file

    [IO.Compression.ZipFile]::ExtractToDirectory( $ZipFile, $destination )
}



<#function unZip ($file, $destination) {

    if ((get-item $file).Length -eq 0){

    return $false

    } else {

    $shell = (new-object -com shell.application)
    $zip = ($shell.NameSpace($file))
    foreach($item in $zip.items())
    {
        $shell.Namespace($destination).copyhere($item)
        #echo "Extracting $file to $destination..."
    } 

    return $true

    }
}#>