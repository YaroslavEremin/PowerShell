function New-ExchDB {
<#
  .SYNOPSIS
  Function creating new Exchange database and set qouta
  .DESCRIPTION
  ----------
  .EXAMPLE
  .\New-ExchDB.ps1 -ExchSrv mb28 -DBType 1G
  .EXAMPLE
  .\New-ExchDB.ps1 -ExchSrv 28 -DBType 200M -NewDBName DB123-200M
  .PARAMETER ExchSrv
  Set target Exchange mailbox server
  .PARAMETER DBType
  Set target quota template for new database. Accesses next value: "200M", "1G", "2G", "5G", "20G".
  .PARAMETER NewDBName
  Set name for new database. If not set that parameter, default value generation auvtomaticly
  #>
[CmdletBinding()]
param (
[Parameter(Mandatory=$True)] $ExchSrv,
[Parameter(Mandatory=$True)] $DBType,
[Parameter(Mandatory=$True)] $NewDBName
)
BEGIN {
write-host "Looking for mailbox server"
[int]$Counter = 0
Get-MailboxServer | % {
If ($_.Name -match $ExchSrv) {
$ExchSrv = $_.Name
$Counter = $Counter + 1
}
}
$ExchSrv
$Counter
$DBType
If ($Counter -eq 0) {
write-host "Can`t find mailbox server.Please type corrent value"
#написать выход из функции
}
ElseIf ($Counter -gt 1) {
write-host "Faund more then one server"
}
If ( -not($NewDBName -like "*") ) {
write-host "Generation DB name by counter mailbox DBs"
$DBArray = Get-MailboxDatabase
$NewDBName = "DB" + $DBArray.Count + $DBType
$NewDBName
}
$NewDBName
}
PROCESS {
write-host "Create database folder"
Invoke-Command -ScriptBlock {New-Item -Name $NewDBName -ItemType directory -Path D:\} -ComputerName $ExchSrv -ArgumentList $NewDBName
#New-Item -Name $NewDBName -ItemType directory -Path D:\
write-host "Create database log files folder"
Invoke-Command -ScriptBlock {New-Item -Name ($NewDBName + "-LOGS") -ItemType directory -Path E:\} -ComputerName $ExchSrv -ArgumentList $NewDBName
#New-Item -Name ($NewDBName + "-LOGS") -ItemType directory -Path E:\
write-host "Create database"
<#New-MailboxDatabase -EdbFilePath ("D:\" + $NewDBName + "\" + $NewDBName + ".edb") -LogFolderPath ("E:\" + $NewDBName + "-LOGS") -Server $SrvName -Name $NewDBName
Start-Sleep 5
write-host "Set mailboxes quota"
Switch ($DBType) {
"200M" {Set-MailboxDatabase -Identity $NewDBName -IssueWarningQuota 180MB -ProhibitSendQuota 190MB -ProhibitSendReceiveQuota 200MB`
-QuotaNotificationSchedule 'Вс.2:00-Вс.2:15, Пн.2:00-Пн.2:15, Вт.2:00-Вт.2:15, Ср.2:00-Ср.2:15, Чт.2:00-Чт.2:15, Пт.2:00-Пт.2:15, Сб.2:00-Сб.2:15'`
-OfflineAddressBook '\<Address_Book>' -MailboxRetention '7.00:00:00' -DeletedItemRetention '5.00:00:00'
}
"1G" { Set-MailboxDatabase -Identity $NewDBName -IssueWarningQuota 922MB -ProhibitSendQuota 973MB -ProhibitSendReceiveQuota 1024MB`
-QuotaNotificationSchedule 'Вс.2:00-Вс.2:15, Пн.2:00-Пн.2:15, Вт.2:00-Вт.2:15, Ср.2:00-Ср.2:15, Чт.2:00-Чт.2:15, Пт.2:00-Пт.2:15, Сб.2:00-Сб.2:15'`
-OfflineAddressBook '\<Address_Book>' -MailboxRetention '7.00:00:00' -DeletedItemRetention '5.00:00:00'
}
"2G" { Set-MailboxDatabase -Identity $NewDBName -IssueWarningQuota 1844MB -ProhibitSendQuota 1946MB -ProhibitSendReceiveQuota 2048MB`
-QuotaNotificationSchedule 'Вс.2:00-Вс.2:15, Пн.2:00-Пн.2:15, Вт.2:00-Вт.2:15, Ср.2:00-Ср.2:15, Чт.2:00-Чт.2:15, Пт.2:00-Пт.2:15, Сб.2:00-Сб.2:15'`
-OfflineAddressBook '\<Address_Book>' -MailboxRetention '7.00:00:00' -DeletedItemRetention '5.00:00:00'
}
"5G" { Set-MailboxDatabase -Identity $NewDBName -IssueWarningQuota 4608MB -ProhibitSendQuota 4864MB -ProhibitSendReceiveQuota 5120MB`
-QuotaNotificationSchedule 'Вс.2:00-Вс.2:15, Пн.2:00-Пн.2:15, Вт.2:00-Вт.2:15, Ср.2:00-Ср.2:15, Чт.2:00-Чт.2:15, Пт.2:00-Пт.2:15, Сб.2:00-Сб.2:15'`
-OfflineAddressBook '\<Address_Book>' -MailboxRetention '7.00:00:00' -DeletedItemRetention '5.00:00:00' }
"20G" { Set-MailboxDatabase -Identity $NewDBName -IssueWarningQuota 18440MB -ProhibitSendQuota 19460MB -ProhibitSendReceiveQuota 20480MB`
-QuotaNotificationSchedule 'Вс.2:00-Вс.2:15, Пн.2:00-Пн.2:15, Вт.2:00-Вт.2:15, Ср.2:00-Ср.2:15, Чт.2:00-Чт.2:15, Пт.2:00-Пт.2:15, Сб.2:00-Сб.2:15'`
-OfflineAddressBook '\<Address_Book>' -MailboxRetention '7.00:00:00' -DeletedItemRetention '5.00:00:00' }
}
Start-Sleep 2
write-host "Mount database"
Mount-Database -Identity $NewDBName#>
}
END {
write-host "Done"
Start-Sleep 2
}
}

New-ExchDB