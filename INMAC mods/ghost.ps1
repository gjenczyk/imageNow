<# #############################################################################
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
# 2.0 2013.10.24 Improved error handling
#
# TO ADD:
#?
# #############################################################################>
# CONFIG #

$pdf = $args[0] 
#"pdf is $pdf"| Out-File D:\INMAC\ne.txt -Append 
$outDir = $args[1] 
#"outDir is $outDir"| Out-File D:\INMAC\ne.txt -Append 
$pdfName = $args[2] 
#"pdfName is $pdfName"| Out-File D:\INMAC\ne.txt -Append 
$inmacVersion = $args[3]
#"inmac version is $inmacVersion" | Out-File D:\INMAC\ne.txt -Append

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

        #Convert the pdf
        $pinfo = New-Object System.Diagnostics.ProcessStartInfo
        $pinfo.FileName = "D:\Program Files\gs\gs9.14\bin\gswin64c.exe"
        $pinfo.RedirectStandardError = $true
        $pinfo.RedirectStandardOutput = $true
        $pinfo.UseShellExecute = $false
        $pinfo.Arguments = "-q -dNumRenderingThreads=4 -dBandBufferSpace=500000000 -sBandListStorage=memory -dBufferSpace=1000000000 -dNOPAUSE -sDEVICE=tiffg4 $param -r600 -c 30000000 setvmthreshold -f $pdfPath -c quit"
        $p = New-Object System.Diagnostics.Process
        $p.StartInfo = $pinfo
        $p.Start() #| Out-Null
        $timer = $p.WaitForExit(300000);

        if (!$timer)
        {
            "$userName - $pdfName timed out @ $(get-date)"| Out-File  D:\INMAC\${inmacVersion}_powerghost.log -Append
            [Environment]::Exit(1)
        }

        $stdout = $p.StandardOutput.ReadToEnd()
        $stderr = $p.StandardError.ReadToEnd()
        #Write-Host "stdout: $stdout"
        #Write-Host "stderr: $stderr"
        #Write-Host "exit code: " + $convert.ExitCode

        if ($stderr)
        {
            if ($stderr.Contains("GPL Ghostscript 9.14: Unrecoverable error, exit code 1"))
            {
                if($stderr.Contains("Error updating TIFF header"))
                {
                    "------------------------------------------------------------------" | Out-File D:\INMAC\${inmacVersion}_powerghost.log -Append
                    "$stderr" | Out-File D:\INMAC\${inmacVersion}_powerghost.log -Append
                    [Environment]::Exit(2)
                }
                else 
                {
                    "------------------------------------------------------------------" | Out-File D:\INMAC\${inmacVersion}_powerghost.log -Append
                    "$stderr" | Out-File D:\INMAC\${inmacVersion}_powerghost.log -Append
                    [Environment]::Exit(3)
                }   
            }

        }   

        $error[0] | Out-File D:\INMAC\${inmacVersion}_powerghost.log -Append      
 }
