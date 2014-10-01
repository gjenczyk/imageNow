# #############################################################################
# NAME: ghost.ps1
# 
# AUTHOR:  Gregg Jenczyk, UMass (UITS)
# DATE:  2013/11/29
# EMAIL: gjenczyk@umassp.edu
# 
# COMMENT:  This script will convert pdfs to tiff using GhostScript
#
# VERSION HISTORY
# 1.0 2013.11.29 Initial Version.
# 1.1 2013.12.17 Added cleanup for orphaned gswin processes
#
# TO ADD:
#?
# #############################################################################
# CONFIG #

$pdf = $args[0] 
#"pdf is $pdf"| Out-File D:\INMAC\ne.txt -Append 
$outDir = $args[1] 
#"outDir is $outDir"| Out-File D:\INMAC\ne.txt -Append 
$pdfName = $args[2] 
#"pdfName is $pdfName"| Out-File D:\INMAC\ne.txt -Append 
$inmacVersion = $args[3]
#"inmac version is $inmacVersion" | Out-File D:\INMAC\ne.txt -Append
$tool = 'D:\Program Files\gs\gs9.14\bin\gswin64c.exe'
$timeout = new-timespan -Minutes 5
$sw = [diagnostics.stopwatch]::StartNew()
$pdfPath = $pdf+$pdfName  
#"pdfPath is $pdfPath"| Out-File D:\INMAC\ne.txt -Append 
# MAIN #

    Get-Item $pdfPath |   ForEach-Object {   
  
        $tiff = $pdfPath -replace 'pdf$','tif'
        $tiffName = $pdfName -replace 'pdf$','tif'
        $outTif = $outDir+$tiffName
           
        #'Processing ' + $pdf | Out-File -Append "D:\INMAC\logs\${inmacVersion}_ghost.log"
        $workingTif = $outTif -replace '.tif$',''
        $param = "-sOutputFile=$workingTif-%06d.tif"

        #$convert = Start-Process $tool -ArgumentList "-q -dNOPAUSE -sDEVICE=tiffg4 $param -r600 $pdfPath -c quit" -passthru
        $convert = Start-Process $tool -ArgumentList "-q -dNumRenderingThreads=4 -dBandBufferSpace=500000000 -sBandListStorage=memory -dBufferSpace=1000000000 -dNOPAUSE -sDEVICE=tiffg4 $param -r600 -c 30000000 setvmthreshold -f $pdfPath -c quit" -passthru
        
        $error[0] | Out-File D:\INMAC\${inmacVersion}_powerghost.log -Append
        
        do
        {
            #this makes sure the vbscript doesn't do any additional processing until the conversion is finished.
            start-sleep -m 250

            if ($sw.elapsed -gt $timeout){
                #kill any GhostScript instances that may be left over after conversion times out
                gwmi -query "select * from win32_process where ProcessId='$convert.Id'" | %{if($_.GetOwner().User -eq "$userName"){$_.terminate()}}                
                $error[0] | Out-File D:\INMAC\${inmacVersion}_ghost.log -Append 
                "$userName - $pdfName timed out @ $(get-date)"| Out-File  D:\INMAC\${inmacVersion}_powerghost.log -Append
                [Environment]::Exit(1)
            }
        } until ((get-process -id $convert.Id -erroraction "silentlycontinue") -eq $null )        

        
 }
