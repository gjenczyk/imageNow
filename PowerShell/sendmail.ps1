# #############################################################################
# NAME: sendmail.ps1
# 
# AUTHOR:  Gregg Jenczyk, UMass (UITS)
# DATE:  2013/12/11
# EMAIL: gjenczyk@umassp.edu
# 
# COMMENT:  This scriptlet sends emails. Check the param section for options
#            
# VERSION HISTORY
# 1.0 2013.12.11 Initial Version.
#
# TO ADD:
# Need to make sure it can handle multiple attachments 
#
# USEFUL SNIPPETS:
# "$(Get-Date) " | Out-File -Append $runLog
# #############################################################################

 function sendmail{
    #values that can be passed to sendmail. this list is extensible
    param([string[]]$to, [string]$f, [string]$s, [string]$m, [string[]]$a, [string]$flag)

     #SMTP server name
     $smtpServer = "69.16.78.38"

     #default parameter check
     if (!$m){
     $m = "This email was sent with no body."
     }

     if (!$to) {
     $to = "gjenczyk@umassp.edu" #BE SURE TO CHANGE THIS TO DI.CORE.whatever when the time comes!
     $m += "`n`nThis message was generated with no recipient specified!"
     }

     if (!$f){ 
     $f = "Document.Imaging.Support@umassp.edu"
     }

     if (!$s){
      $s = "A message from Document Imaging"
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

}