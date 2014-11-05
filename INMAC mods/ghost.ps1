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
#"pdf is $pdf"| Out-File  D:\INMAC\${inmacVersion}_powerghost.log -Append
$outDir = $args[1] 
#"outDir is $outDir"| Out-File  D:\INMAC\${inmacVersion}_powerghost.log -Append
$pdfName = $args[2] 
#"pdfName is $pdfName"| Out-File  D:\INMAC\${inmacVersion}_powerghost.log -Append
$inmacVersion = $args[3]
#"inmac version is $inmacVersion" | Out-File  D:\INMAC\${inmacVersion}_powerghost.log -Append
$pdfPath = $pdf+$pdfName 
# Tool Paths #
$default = "D:\Program Files\gs\gs9.14\bin\gswin64c.exe"
$repair = "D:\Program Files\gs\gs\bin\gswin32c.exe" 
#"pdfPath is $pdfPath"| Out-File  D:\INMAC\${inmacVersion}_powerghost.log -Append

# functions #

function ConvertPDF([string]$version, [string]$param1, [string]$pdfPath1)
{
    $pinfo = New-Object System.Diagnostics.ProcessStartInfo
    $pinfo.FileName = $version
    $pinfo.RedirectStandardError = $true
    $pinfo.RedirectStandardOutput = $true
    $pinfo.UseShellExecute = $false
    $pinfo.Arguments = "-q -dNOPAUSE -sDEVICE=tiffg4 $param -r600 $pdfPath -c quit"
    $p = New-Object System.Diagnostics.Process
    $p.StartInfo = $pinfo
    $p.Start() #| Out-Null
    $timer = $p.WaitForExit(300000);
}

# MAIN #

    Get-Item $pdfPath |   ForEach-Object {   
  
        $tiff = $pdfPath -replace 'pdf$','tif'
        $tiffName = $pdfName -replace 'pdf$','tif'
        $outTif = $outDir+$tiffName
           
        #'Processing ' + $pdf | Out-File -Append "D:\INMAC\logs\${inmacVersion}_ghost.log"
        $workingTif = $outTif -replace '.tif$',''
        $param = "-sOutputFile=$workingTif-%06d.tif"
        #"workingTif is $workingTif"| Out-File  D:\INMAC\${inmacVersion}_powerghost.log -Append
        #Convert the pdf
        $pinfo = New-Object System.Diagnostics.ProcessStartInfo
        $pinfo.FileName = "D:\Program Files\gs\gs9.14\bin\gswin64c.exe"
        $pinfo.RedirectStandardError = $true
        $pinfo.RedirectStandardOutput = $true
        $pinfo.UseShellExecute = $false
        $pinfo.Arguments = "-q -dNOPAUSE -sDEVICE=tiffg4 $param -r600 $pdfPath -c quit"
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
                elseif($stdout.Contains("Error: /undefined in --.PDFexecform--"))
                {
                     "------------------------------------------------------------------" | Out-File D:\INMAC\${inmacVersion}_powerghost.log -Append
                    "$stderr" | Out-File D:\INMAC\${inmacVersion}_powerghost.log -Append
                    ConvertPDF $repair $param $pdfPath
                    
                    #[Environment]::Exit(4)
                }
                else 
                {
                    "------------------------------------------------------------------" | Out-File D:\INMAC\${inmacVersion}_powerghost.log -Append
                    "$stdout" | Out-File D:\INMAC\${inmacVersion}_powerghost.log -Append
                    "$stderr" | Out-File D:\INMAC\${inmacVersion}_powerghost.log -Append
                    [Environment]::Exit(3)
                }   
            }

        }   

        $error[0] | Out-File D:\INMAC\${inmacVersion}_powerghost.log -Append      
 }

