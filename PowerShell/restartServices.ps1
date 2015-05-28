<##############################################################################
# NAME: restartServices.ps1
# 
# AUTHOR:  Gregg Jenczyk, UMass (UITS)
# DATE:  2015/05/21
# EMAIL: gjenczyk@umassp.edu
# 
# DESCRIPTION:  This script restarts the services on a server based on the 
#               agent's status.  The list of servers to work is read from a csv.
#               Csv location is defined in csvPath. The services to work are 
#               defined in serviceToMatch. You need to set-executionpolicy bypass.
#               The script needs to be run as someone with the ability to restart 
#               services on the remote servers (admin access?)
#
# VERSION HISTORY
# 1.0 2015.05.21 Initial Version.
#
# TO ADD:
# 
# USEFUL SNIPPETS:
# "$($error[0]) $(Get-Date) " | Out-File -Append $runLog
##############################################################################>

## -- CONFIGURATION -- ##
$scriptName = 'restartServices.ps1'
$csvPath = '\\ssctxfs01\folderredir\gjenczyk.UMASSP.EDU\desktop\oemTest\'
$runLog = "${csvPath}restartServices_$(Get-Date -Format 'MMddyyHHmmss').log"
$emailData = "${csvPath}restartServices_$(Get-Date -Format 'MMddyyHHmmss').csv"
$formattedData = "${csvPath}restartServices_$(Get-Date -Format 'MMddyyHHmmss')_F.csv"
$serviceToMatch = @("ImageNow Retention", "ImageNow Monitor") #"Oracleagent12g"
$notify = @("gjenczyk@umassp.edu")
$started = @()
$failed = @()

## -- FUNCTIONS -- ##

<# Do this to get the newest csv in a directory.  Is it necessary? idk #>
function get-csv ([string]$pathToFile)
{
    $csvToUse = ''
    $csvCurrent = ''
    Get-ChildItem -path $pathToFile -filter '*.csv' | ForEach-Object {
        $csvCurrent = $_
        if (!$csvToUse)
        {
            $csvToUse = $csvCurrent
        }
        elseif ($csvCurrent.LastWriteTime -gt $csvToUse.LastWriteTime)
        {
            $csvToUse = $csvCurrent
        }
    }
    "Using: $csvToUse" | Out-File -Append $runLog
    return $csvToUse
}# end get-csv

<# this one wraps the email config in a neat package #>
function send-mail{
    #values that can be passed to sendmail. this list is extensible
    param([string[]]$to, [string]$f, [string]$s, [string]$m, [string[]]$a, [string]$flag)
    #SMTP server name
    $smtpServer = "69.16.78.38"
    #default parameter check
    if (!$m){
        $m = "This email was sent with no body."
    }
    if (!$to) {
        $to = "gjenczyk@umassp.edu" #be sure to change this to whoever when the time comes!
        $m += "`n`nThis message was generated with no recipient specified!"
    }
    if (!$f){ 
        $f = "gjenczyk@umassp.edu" #be sure to change this to whoever when the time comes!
    }
    if (!$s){
      $s = "Here's an email from me"
    }
    #Default Send-MailMessage parameters     
    $mailParams = @{
        To = $to
        From = $f
        Subject = $s
        Body = $m
        SmtpServer = $smtpServer
    }
    #Checking for non-standard parameters
    if ($a){
        #usage for $a: -a "attachment","2ndAttachment","etc"
        $mailParams += @{Attachment = $a}
     }
    if ($flag){
        if ($flag -eq "BodyAsHtml")
        {
            $mailParams += @{BodyAsHtml = $true}
        }
    }  
   Send-MailMessage @mailParams 
}# end send-mail

<#this one builds the search criteria for the services#>
function build-matchRule ([string[]] $matchService)
{
    $rule = ''
    for ($i = 0; $i -lt $matchService.length; $i++)
    {
        $rule += '$_.Name -match' + " `"$($matchService[$i])`""
        if ($i -lt $matchService.length-1)
        {
            $rule += ' -or '
        }
    }
    return $rule
}# end build-matchRule

<#this function puts html formatting on the email#>
function format-email ([string] $raw, $formatted)
{
    $a = "<style>"
    $a = $a + "TABLE{border-width: 1px;border-style: solid;border-color: black;border-collapse: collapse;}"
    $a = $a + "TH{border-width: 1px;padding: 5px;border-style: solid;border-color: black;background-color: gainsboro;}"
    $a = $a + "TD{border-width: 1px;padding: 5px;border-style: solid;border-color: black;text-align: center;}"
    $a = $a + "</style>"

    Import-Csv $raw -Header "Server","Service","Started?","Message" | ConvertTo-HTML -head $a | Out-File -FilePath $formatted
}# end format-email

<#function to clean up the leavings of the process#>
function clean-up
{
    foreach ($file in $args)
    {
        if(Test-Path $file)
        {
            Remove-Item $file
        }
    }
}

# - MAIN - #
try
{
    "Starting $scriptName @ $(Get-Date)" | Out-File -Append $runLog
    # Get the latest csv from the directory
    $workingCsv = get-csv $csvPath
    $matchRule = build-matchRule $serviceToMatch

    #load the csv an process each row
    import-csv -path ${csvPath}${workingCsv} -header "Server" | ForEach-Object {
        $thisServer = $_.Server
        "Working ${thisServer}" | Out-File -Append $runLog
        Get-Service -ComputerName $thisServer | Where-Object (& {[scriptblock]::Create($matchRule)} $matchRule) | ForEach-Object {
            #$_ | Format-List | Out-File -Append $runLog
            $serviceName = $_.Name
            $serviceStatus = $_.Status
            #splitting this out so we have a record of services that weren't actually stopped
            if ($serviceStatus -match "Stopped")
            {
                "$serviceName is $serviceStatus.  Attempting to restart @ $(Get-Date)..." | Out-File -Append $runLog
                Set-Service -ComputerName $thisServer -Name $serviceName -Status Running
                if(!$?)
                {
                    "Failed to start $serviceName  @ $(Get-Date)" | Out-File -Append $runLog
                    $error[0] | Format-List -Force | Out-File -Append $runLog 
                    "$thisServer,$serviceName,Failed,$($error[0])" | Out-File -Append $emailData
                }
                else
                {
                    "Started $serviceName  @ $(Get-Date)" | Out-File -Append $runLog
                    "$thisServer,$serviceName,Started," | Out-File -Append $emailData
                }
            }# end of processing stopped services
        }# end of checking on 
    }
    
    #format the email
    format-email $emailData $formattedData
    $body = (Get-Content $formattedData)
    
    #notify whoever
    send-mail -t $notify -a $runLog -s "Service Restart Report" -m $body -flag "BodyAsHtml"
}
catch
{
    $error[0] | Format-List -Force | Out-File -Append $runLog 
    send-mail -t $notify -a $runLog -s "Fatal Error in restartServices.ps1" -b "Check attached log for details"
}
finally
{
    #clean up the files
    clean-up $runLog $emailData $formattedData
}