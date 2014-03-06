add-pssnapin quest.activeroles.admanagement
Function Send-ReportCreatedAccounts {
$Date = (Get-Date)
$Date = $Date.AddDays(-5)
$Accounts =  Get-QADUser -CreatedAfter $Date
$Date = $Date.tostring("dd.M.yyyy")
$Users  = $Accounts | ?{ $_.CanonicalName -match 'Users/' }
$Users = $Users | Select-Object Description,SamAccountName,CreationDate | Sort-Object Description,SamAccountName
$Users | ft -AutoSize | Out-File c:\temp\info.txt
$Count = $Users.count
$Subject = "Количество созданных на этой неделе учетных записей пользователей"
$From = ""
$To = ""
$To1 = ""
$MailBody = "Создано учетных записей пользователей с <b>$Date : $Count</b>"
$Attachments = "c:\temp\info.txt"
Send-MailMessage -From $From -Subject $Subject -To $To  -Body $MailBody -BodyAsHtml -Encoding UTF8 -SmtpServer exch-cas02 -Attachments $Attachments
Send-MailMessage -From $From -Subject $Subject -To $To1  -Body $MailBody -BodyAsHtml -Encoding UTF8 -SmtpServer exch-cas02 -Attachments $Attachments
}
Send-ReportCreatedAccounts