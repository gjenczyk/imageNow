Param(
  [string]$FromIn,
  [string]$ToIn,
  [string]$SmtpServerIn,
  [string]$SubjectIn,
  [string]$nonHTMLemailBody
  )

  
. "\\boisnas215c1.umasscs.net\diimages67tst\script\PowerShell\powermail.ps1"

write-host $FromIn
  write-host $ToIn
  write-host $SmtpServerIn
  write-host $SubjectIn
  write-host $nonHTMLemailBody


$image = @{
     image = "\\boisnas215c1.umasscs.net\diimages67tst\script\images\DIlogoforemail.gif"
}

$body = @'
<html>  
  <body>
    <img src="cid:image"><br>
  </body>  
</html>  

'@

$params = @{ 
    InlineAttachments = $image 
    #Attachments = 'D:\test.txt', 'D:\sendmail.txt' 
    Body = $body
    BodyAsHtml = $true 
    Subject = $SubjectIn
    From = $FromIn
    To = $ToIn
    SmtpServer = $SmtpServerIn
}


Send-MailMessage  @params

