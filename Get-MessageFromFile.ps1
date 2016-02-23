



<#
 .DESCRIPTION
 NAME: Get-MessageFromFile.ps1
 
 AUTHOR: jrv , DSS
 DATE  : 8/20/2011
 .SYNOPSIS
     List the headers of an email message save in a MSG file. 
 .EXAMPLE
     Get-MessageFromFile 'c:\scripts\message.msg'
     Get-MessageFromFile 'c:\scripts\message.msg' -verbose
 .LINK
     http;//www.designedsystemsonline.com/upload/Ger-MessageFromFile.ps1
#>




Function Get-MessageFromFile{
     Param(
          [Parameter(
               ValueFromPipeLine=$true,
               Mandatory=$true,
               Position=0,
               HelpMessage='Enter path to MSG file'
          )][string]$filename
     )
     process{
          if(-not(Test-Path $filename)){
               Write-Host 'File not found!' -ForegroundColor red -BackgroundColor white
               return #emptyhanded
          }
          $stmMSGFile=New-Object -com ADODB.Stream
          $stmMSGFile.Charset="ascii"
          $stmMSGFile.Type=2 #adTypeText
          $stmMSGFile.Open()
          $stmMSGFile.LoadFromFile( $filename)
          [void]$stmMSGFile.ReadText(-1) #adReadAll
          $iMsg=New-Object -com CDO.Message
          $iMsg.DataSource.OpenObject($stmMSGFile, '_Stream')    
          Foreach($f in $iMsg.Fields){
               if($f.Name -like '*mailheader*'){
                    New-Object PSObject -Property @{URN=$f.Name; Value=$f.UnderLyingValue}
               }else{
                    Write-Verbose ('Skipping field: ' + $f.Name)
               }
          }
     }
}


$DropDir = "D:\test\"
$Messages = Get-ChildItem -r -Path $DropDir -file -include *.eml

#if ($Messages -ne $Null)

$Messages | foreach {Get-MessageFromFile $_ | select URn, Value | where URN -eq "urn:schemas:mailheader:x-ms-has-attach" | select Value | if ($_ -eq "yes") {Write-host "It has an attachment"}} 




| Select Value |}

 if ($_.Value -eq $true) {write-Host " $Messages.fullname Has an attachment"}
    else 
    {Write-host "$_ Has no attachment"}


else {write-host "Theres nothng to do"}}




if ($_.Value -ne $Null) {Write-host "not equal Null"
}




{$_.Value}









$Message = Get-MessageFromFile ddac477601d11c2700000002.eml


$Message = Get-MessageFromFile ddac477601d11c2700000002.eml
$mailSender = $message | select  URN,Value | where URN -eq "urn:schemas:mailheader:x-sender" | Select $_.Value 
$MailSender = $mailsender.Value
$hasAttachment = $Message | select URn, Value | where URN -eq "urn:schemas:mailheader:x-ms-has-attach" | Select $_.Value 
$hasAttachment = $hasAttachment.Value


#$test = Get-MessageFromFile ddac477601d11c2700000002.eml
#$test2 = $test | select  @{n="URN";e={$_.URN.replace("urn:schemas:mailheader:","")}},Value
#$test3 = $test2 | where URN -eq "x-sender" | Select Value 
#$test3.value


#$MailSender = Get-MessageFromFile ddac477601d11c2700000002.eml 
#$mailsender2 = $MailSender | select  @{n="URN";e={$_.URN.replace("urn:schemas:mailheader:","")}},Value
#$mailSender3 = $mailsender2 | where URN -eq "x-sender" | Select Value 



#@{n="URN";e={$_.URN.replace("urn:schemas:mailheader:","")}},Value | where URN -eq "x-sender" | Select $_.Value 
#$MailSender.value


#$MailSender = Get-MessageFromFile ddac477601d11c2700000002.eml | select  URN,Value | where URN -eq "urn:schemas:mailheader:x-sender" | Select $_.Value 
#$MailSender = $mailsender.Value

 .\munpack.exe D:\Users\buskeyl\OneDrive\_Scripts\mail\ddac477601d11c2700000002.eml -C $Homedir


