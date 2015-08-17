# #############################################################################
# NAME: powerprint.ps1
# 
# AUTHOR:  Gregg Jenczyk, UMass (UITS)
# DATE:  2014/01/14
# EMAIL: gjenczyk@umassp.edu
# 
# COMMENT:  This script will convert Word docs to tiff via the ImageNow Printer
#
# VERSION HISTORY
# 1.0 2014.01.14 Initial Version.
# 1.1 2014.03.14 Added handling for corrupted filesfdsa
#
# TO ADD
# -Additional error handling based on powerprint.log output
# #############################################################################
## CONFIG ##

$inDir = $args[0] #just input directory
$outDir = $args[1]  #just output directory
$docName = $args[2] #just the input file name
$inmacVersion = $args[3] #which inmac this came from
$doc = $inDir+$docName
$uName = $env:username
$conDir = 'D:\INMAC\'+$inmacVersion+'\printer_output\'
$conDoc = $conDir+$docName

## FUNCTIONS ##
function convert ([string]$docPath, [string]$workDir, [string]$workDoc){
    "Converting Word to tif" | write-host

    copyItem $docPath $workDir
    
    $passObj=new-object -comobject word.application

    $passDoc = $passObj.Documents.Open($workDoc,$false,$true,$false,"groovy")
    if ($?)
    {
        $id = (gwmi -query "select * from win32_process where name='WINWORD.EXE'" | %{if($_.GetOwner().User -eq "$uName"){echo $_.ProcessId}})

        $passObj.PrintOut()
        try
        {
            $passObj.ActiveDocument.Close([ref]0)
        }
        catch 
        {
           [System.Runtime.InteropServices.ExternalException]
                
           Stop-Process $id
           remove-item $workDoc
           [Environment]::Exit(7)
        }
        $passObj.quit()
        $error[0] | Out-File D:\INMAC\${inmacVersion}_powerprint.log -Append
        remove-item $workDoc
    } else {
        $error[0] | Out-File D:\INMAC\${inmacVersion}_powerprint.log -Append 
        if (($error[0] -like '*The password is incorrect*') -eq $true)
        {
            $passObj.quit()
            remove-item $workDoc
            [Environment]::Exit(1)  
        } elseif  (($error[0] -like '*You are attempting to open a f*') -eq $true)
        {
            $passObj.quit()
            remove-item $workDoc
            [Environment]::Exit(2)  
        } elseif  (($error[0] -like '*The file appears to be corrupted*') -eq $true)
        {
            $passObj.quit()
            remove-item $workDoc
            [Environment]::Exit(3)  
        } elseif  (($error[0] -like '*To help protect your computer this file cannot be opened*') -eq $true)
        {
            $passObj.quit()
            remove-item $workDoc
            [Environment]::Exit(4)  
        } elseif  (($error[0] -like '*Exception from HRESULT: 0x800A175D*') -eq $true)
        {
            $passObj.quit()
            remove-item $workDoc
            [Environment]::Exit(5)  
        } elseif  (($error[0] -like '*the file format does not match the file extension*') -eq $true)
        {
            $passObj.quit()
            remove-item $workDoc
            [Environment]::Exit(6)  
        } else {
            $passObj.quit()
            remove-item $workDoc
            [Environment]::Exit(8)  
        } #Add more errors here as you think of them. 

    }

}

function copyItem ([string]$wdDoc, [string]$prntDir){
    "Copying doc" | write-host
    copy-item $wdDoc  $prntDir
    $error[0] | Out-File D:\INMAC\${inmacVersion}_powerprint.log -Append
}

## BEGIN ##

convert $doc $conDir $conDoc 

$error[0] | Out-File D:\INMAC\${inmacVersion}_powerprint.log -Append 
exit 0;
