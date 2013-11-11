function New-ExchDB {
<#
  .SYNOPSIS
  Function creating new Exchange database and set qouta
  .DESCRIPTION
  ----------
  .EXAMPLE
  .\New-ExchDB.ps1 -ExchSrv mb28 -DBType 1G
  .EXAMPLE
  .\New-ExchDB.ps1 -Ex 28 -Db 200M -N DB123-200M
  .PARAMETER ExchSrv
  Set target Exchange mailbox server
  .PARAMETER DBType
  Set target quota template for new database. Accept next value: "200M", "1G", "2G", "5G", "20G".
  .PARAMETER NewDBName
  Set name for new database. If not set, default value generated automatically
  #>
[CmdletBinding()]
param (
[Parameter(Mandatory=$True)] $ExchSrv,
[Parameter(Mandatory=$True)] $DBType,
[Parameter(Mandatory=$False)] $NewDBName ="Value"
)
BEGIN {
write-host -ForegroundColor DarkYellow "Looking for mailbox server"
[int]$Counter = 0
Get-MailboxServer | % {
If ($_.Name -match $ExchSrv) {
$ExchSrv = $_.Name
$Counter = $Counter + 1
}
}
write-host -ForegroundColor Green "Target mailbox server is" ; $ExchSrv
$DBArray = Get-MailboxDatabase
$DBArray.Count
If ($Counter -eq 0) {
write-host -ForegroundColor Red "Can`t find mailbox server.Please type corrent value"
#написать выход из функции
}
ElseIf ($Counter -gt 1) {
write-host -ForegroundColor Red "Faund more then one server"
}
If ($NewDBName -eq "Value") {
write-host "Generation DB name by counter mailbox DBs"
$DBArray = Get-MailboxDatabase
$NewDBName = "DB" + $DBArray.Count + "-"+ $DBType
$NewDBName
}
$NewDBName
}
PROCESS {
write-host -ForegroundColor Green "Create database folder"
New-Item -Name $NewDBName -ItemType directory -Path ("\\" + $ExchSrv + "\D$\")
write-host -ForegroundColor Green "Create database log files folder"
New-Item -Name ($NewDBName + "-LOGS") -ItemType directory -Path ("\\" + $ExchSrv + "\E$\")
write-host -ForegroundColor Green "Create database"
New-MailboxDatabase -EdbFilePath ("D:\" + $NewDBName + "\" + $NewDBName + ".edb") -LogFolderPath ("E:\" + $NewDBName + "-LOGS") -Server $ExchSrv -Name $NewDBName | Out-Null
Start-Sleep 5
write-host -ForegroundColor DarkYellow  "Set mailboxes quota"
Switch ($DBType) {
"200M" {
write-host -ForegroundColor Green "200M"
Set-MailboxDatabase -Identity $NewDBName -IssueWarningQuota 180MB -ProhibitSendQuota 190MB -ProhibitSendReceiveQuota 200MB -QuotaNotificationSchedule 'Вс.2:00-Вс.2:15, Пн.2:00-Пн.2:15, Вт.2:00-Вт.2:15, Ср.2:00-Ср.2:15, Чт.2:00-Чт.2:15, Пт.2:00-Пт.2:15, Сб.2:00-Сб.2:15' -OfflineAddressBook '\OKOI_OAB_EXCH2010' -MailboxRetention '7.00:00:00' -DeletedItemRetention '5.00:00:00'
}
"1G" {
write-host -ForegroundColor Green "1G"
Set-MailboxDatabase -Identity $NewDBName -IssueWarningQuota 922MB -ProhibitSendQuota 973MB -ProhibitSendReceiveQuota 1024MB -QuotaNotificationSchedule 'Вс.2:00-Вс.2:15, Пн.2:00-Пн.2:15, Вт.2:00-Вт.2:15, Ср.2:00-Ср.2:15, Чт.2:00-Чт.2:15, Пт.2:00-Пт.2:15, Сб.2:00-Сб.2:15' -OfflineAddressBook '\OKOI_OAB_EXCH2010' -MailboxRetention '7.00:00:00' -DeletedItemRetention '5.00:00:00'
}
"2G" {
write-host -ForegroundColor Green "2G"
Set-MailboxDatabase -Identity $NewDBName -IssueWarningQuota 1844MB -ProhibitSendQuota 1946MB -ProhibitSendReceiveQuota 2048MB -QuotaNotificationSchedule 'Вс.2:00-Вс.2:15, Пн.2:00-Пн.2:15, Вт.2:00-Вт.2:15, Ср.2:00-Ср.2:15, Чт.2:00-Чт.2:15, Пт.2:00-Пт.2:15, Сб.2:00-Сб.2:15' -OfflineAddressBook '\OKOI_OAB_EXCH2010' -MailboxRetention '7.00:00:00' -DeletedItemRetention '5.00:00:00'
}
"5G" {
write-host -ForegroundColor Green "5G"
Set-MailboxDatabase -Identity $NewDBName -IssueWarningQuota 4608MB -ProhibitSendQuota 4864MB -ProhibitSendReceiveQuota 5120MB -QuotaNotificationSchedule 'Вс.2:00-Вс.2:15, Пн.2:00-Пн.2:15, Вт.2:00-Вт.2:15, Ср.2:00-Ср.2:15, Чт.2:00-Чт.2:15, Пт.2:00-Пт.2:15, Сб.2:00-Сб.2:15' -OfflineAddressBook '\OKOI_OAB_EXCH2010' -MailboxRetention '7.00:00:00' -DeletedItemRetention '5.00:00:00' }
"20G" {
write-host -ForegroundColor Green "20G"
Set-MailboxDatabase -Identity $NewDBName -IssueWarningQuota 18440MB -ProhibitSendQuota 19460MB -ProhibitSendReceiveQuota 20480MB -QuotaNotificationSchedule 'Вс.2:00-Вс.2:15, Пн.2:00-Пн.2:15, Вт.2:00-Вт.2:15, Ср.2:00-Ср.2:15, Чт.2:00-Чт.2:15, Пт.2:00-Пт.2:15, Сб.2:00-Сб.2:15' -OfflineAddressBook '\OKOI_OAB_EXCH2010' -MailboxRetention '7.00:00:00' -DeletedItemRetention '5.00:00:00' }
default {
write-host -ForegroundColor Red "Quota not set"
}
}
Start-Sleep 2
write-host -ForegroundColor Green "Mount database"
Mount-Database -Identity $NewDBName
}
END {
write-host -ForegroundColor Green "Send infomail"
if ($ExchSrv -match "mb2") {
$City = "Сочи"
} else {
$City = "Москва"
}
$MailBody = "Новая почтовая база<br><p>Сервер: <b>$ExchSrv</b><br>База: <b>$NewDBName</b><br>Расположение: <b>$City</b></p>"
Send-MailMessage -From ($env:USERNAME + "@sochi2014.com") -Subject "Создана новая почтовая база $NewDBName" -To 'sysadmins@SOCHI2014.COM' -Body $MailBody -BodyAsHtml -Encoding Unicode -SmtpServer exch-cas02
write-host -ForegroundColor Green "Done"
Start-Sleep 2
}
}
