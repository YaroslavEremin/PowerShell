#Создание контакта через Exchange
$m_a = <mail_address> | New-MailContact -Name $m_a –Alias $m_a -ExternalEmailAddress $m_a

#Вывод статистики по количеству mailboxes в базах на серверах Exchange 
(Get-MailboxDatabase) | where {$_.Server -like <Template>} |
Select-Object Server,Name,@{Name='Type';Expression={[regex]::replace($_.name,'^.*-','')}},`
@{Name="Counter";Expression={(@(Get-Mailbox -Database $_.name)).Count}} |
Sort -Property "Type","Counter" -Descending | Format-Table -AutoSize

Get-MailboxServer | %{ Get-Mailboxstatistics -Server $_.name } |
?{ $_.DisconnectDate -ne $null } |
Select DisplayName,@{n="StoreMailboxIdentity";e={$_.MailboxGuid}},Database

#Вывод расширенной статистики по  mailboxes в базе
Get-Mailbox -Database <DB_Name> | sort-object |
Select-Object name,alias,servername,ProhibitSendQuota,IssueWarningQuota,MaxReceiveSize,`
MaxSendSize,DisplayName,Database,PrimarySmtpAddress,ProhibitSendReceiveQuota,`
@{n="Size(KB)";e = {$MBXstat = Get-MailboxStatistics $_.name; $MBXstat.totalItemsize}},`
@{n="Items"; e = {$MBXstat = Get-MailboxStatistics $_.name ; $MBXstat.itemcount; $MBXstat.storageLimitStatus}}

#Информация о квотах, примененных к почтовым базам в Организации (запускать в Exchange 2010)
Get-MailboxDatabase -IncludePreExchange2010 |
Select-Object Name, ProhibitSendReceiveQuota, ProhibitSendQuota, IssueWarningQuota |
Sort -Property Name | Format-Table -AutoSize

#Выгрузка в PST-файл (запускать в Exchange 2010)
New-MailboxExportRequest -Mailbox <MAilbox_Alias> -FilePath C:\Temp\<MAilbox_Alias>.pst

#Настройка квот и некоторых других параметров почтовых баз (перед использованием команды необходимо внести изменение в параметр Identity, подставив реальные значения.
Set-MailboxDatabase -Identity DB<NN>-200M -IssueWarningQuota 180MB -ProhibitSendQuota 190MB -ProhibitSendReceiveQuota 200MB` -QuotaNotificationSchedule 'Вс.2:00-Вс.2:15, Пн.2:00-Пн.2:15, Вт.2:00-Вт.2:15, Ср.2:00-Ср.2:15, Чт.2:00-Чт.2:15, Пт.2:00-Пт.2:15, Сб.2:00-Сб.2:15' -OfflineAddressBook '\<Address_Book>' -MailboxRetention '7.00:00:00' -DeletedItemRetention '5.00:00:00'
Set-MailboxDatabase -Identity DB<NN>-1G -IssueWarningQuota 922MB -ProhibitSendQuota 973MB -ProhibitSendReceiveQuota 1024MB -QuotaNotificationSchedule 'Вс.2:00-Вс.2:15, Пн.2:00-Пн.2:15, Вт.2:00-Вт.2:15, Ср.2:00-Ср.2:15, Чт.2:00-Чт.2:15, Пт.2:00-Пт.2:15, Сб.2:00-Сб.2:15' -OfflineAddressBook '\<Address_Book>' -MailboxRetention '7.00:00:00' -DeletedItemRetention '5.00:00:00'
Set-MailboxDatabase -Identity DB<NN>-2G -IssueWarningQuota 1844MB -ProhibitSendQuota 1946MB -ProhibitSendReceiveQuota 2048MB -QuotaNotificationSchedule 'Вс.2:00-Вс.2:15, Пн.2:00-Пн.2:15, Вт.2:00-Вт.2:15, Ср.2:00-Ср.2:15, Чт.2:00-Чт.2:15, Пт.2:00-Пт.2:15, Сб.2:00-Сб.2:15' -OfflineAddressBook '\<Address_Book>' -MailboxRetention '7.00:00:00' -DeletedItemRetention '5.00:00:00'
Set-MailboxDatabase -Identity DB<NN>-5G -IssueWarningQuota 4608MB -ProhibitSendQuota 4864MB -ProhibitSendReceiveQuota 5120MB -QuotaNotificationSchedule 'Вс.2:00-Вс.2:15, Пн.2:00-Пн.2:15, Вт.2:00-Вт.2:15, Ср.2:00-Ср.2:15, Чт.2:00-Чт.2:15, Пт.2:00-Пт.2:15, Сб.2:00-Сб.2:15' -OfflineAddressBook '\<Address_Book>' -MailboxRetention '7.00:00:00' -DeletedItemRetention '5.00:00:00'
Set-MailboxDatabase -Identity DB<NN>-20G -IssueWarningQuota 18440MB -ProhibitSendQuota 19460MB -ProhibitSendReceiveQuota 20480MB -QuotaNotificationSchedule 'Вс.2:00-Вс.2:15, Пн.2:00-Пн.2:15, Вт.2:00-Вт.2:15, Ср.2:00-Ср.2:15, Чт.2:00-Чт.2:15, Пт.2:00-Пт.2:15, Сб.2:00-Сб.2:15' -OfflineAddressBook '\<Address_Book>' -MailboxRetention '7.00:00:00' -DeletedItemRetention '5.00:00:00'

# Очистить список завершенных выгрузок в PST-файл
Get-MailboxExportRequest -Status Completed | Remove-MailboxExportRequest

#Просмотр размера почтовых баз
Get-MailboxDatabase -Status | select ServerName,Name,DatabaseSize

#Получение списка почтовых ящиков с персонально установленными квотами
Get-Mailbox -Filter { UseDatabaseQuotaDefaults -eq $False } -ResultSize Unlimited |
Select-Object Name,ServerName,Database | Format-Table -AutoSize

#Получение списка писем с темами отправленных/полученных сотрудником за определенный временной интервал.
#Полученных
Get-ExchangeServer | where {$_.isHubTransportServer -eq $true -or $_.isMailboxServer -eq $true}  | %{
Get-MessageTrackingLog -Recipients:<mddressail_a> -EventID "RECEIVE" -Start "02/18/2013 0:00:00" -End "03/06/2013 23:59:00"} |
Select-Object Timestamp,Sender,@{l="Recipients";e={$_.Recipients -join " "}},MessageSubject |
Sort-Object -Property Timestamp  | Export-Csv C:\TEMP\report.txt -encoding unicode
#Отправленных
Get-ExchangeServer | where {$_.isHubTransportServer -eq $true -or $_.isMailboxServer -eq $true}  | %{
Get-MessageTrackingLog -Sender <mddressail_a> -EventID "SEND" -Start "02/18/2013 0:00:00" -End "03/06/2013 23:59:00" } |
Select-Object Timestamp,Sender,@{l="Recipients";e={$_.Recipients -join " "}},MessageSubject |
Sort-Object -Property Timestamp  | Export-Csv C:\TEMP\report.txt -encoding unicode

#Статистика по размерам ящиков в базах Exchange
$Bases = <DB_Name>,<DB_Name>,<DB_Name>
$( ForEach ( $Base in $Bases ) { Get-Mailbox -Database $Base } ) |
Select-Object Alias, Database, @{ Name = "Size"; Expression = { $Size = Get-MailboxStatistics $_.name ; $Size.totalItemsize}} |
Sort-Object -Property Size -Descending | Export-Csv -Path d:\stat.csv

#10 самых маленьких ящиков в базе Exchange
$Bases = <DB_Name>,<DB_Name>,<DB_Name>
$( ForEach ( $Base in $Bases ) { Get-Mailbox -Database $Base } ) |
Select-Object Alias, Database, @{ Name = "Size"; Expression = { $Size = Get-MailboxStatistics $_.name ; $Size.totalItemsize}} |
Sort-Object -Property Size | Select-Object -First 10

#Создать список рассылки
$group = "name"
$manager = "user4"
$Users = "user1","user2","user3"
$info = "INC0233611 До 30.03.2014"
New-DistributionGroup -Name $Group -DisplayName $Group -Members $Users -Alias $group -ManagedBy $manager -PrimarySmtpAddress ($group + "@sochi2014.com") -Notes $info -SamAccountName $group
Get-ADGroupMember $group | Select-Object name,SamAccountName |ft -AutoSize


#Создать ящики
Import-Csv d:\m.csv |
%{ Write-Host -ForegroundColor Green 'Создаем ящик' ; $_.mailbox
New-Mailbox -Alias $_.mailbox -Password (ConvertTo-SecureString <your_password> -AsPlainText -Force)`
-Database <DB_Name> -DisplayName $_.mailbox -FirstName $_.mailbox -LastName $_.mailbox`
-Name $_.mailbox -SamAccountName $_.mailbox -UserPrincipalName (@($_.mailbox + '@contoso.com'))
Write-Host -ForegroundColor Green 'Прописываем менеджера и комментарий' ; $_.username
Set-ADUser -Identity $_.mailbox -Manager $_.username -Replace @{'info'=<комментарий>}
}

#Создать ящики 2
$mailboxes = "name"
$SendAs = "user1","user2","user3"
$FullAccess = "user1","user2"
$Manager = "user4"
$info = "______"
$MailboxDB = "DB_name"
ForEach ( $mail in  $mailboxes ) {
Write-Host -ForegroundColor Green 'Создаем ящик' ; $mail
New-Mailbox -Alias $mail -Password (ConvertTo-SecureString '3JJf*n31v(I&' -AsPlainText -Force) -Database $MailboxDB -DisplayName $mail -FirstName $mail -LastName $mail -Name $mail -SamAccountName $mail -UserPrincipalName (@($mail + '@sochi-2014.ru'))
Write-Host -ForegroundColor Green 'Прописываем менеджера и комментарий'
Start-Sleep 3
Set-ADUser -Identity $mail -Manager $Manager -Replace @{'info' = $info}
ForEach ( $i in $SendAs ) {
Add-ADPermission -Identity $mail -User $i -Extendedrights "Send As"
}
ForEach ( $j in $FullAccess ) {
Add-MailboxPermission -Identity $mail -User $j -AccessRights 'FullAccess'
}
}

#Создать переговорки
 | %{
Write-Host -ForegroundColor Green 'Создаем ящик' ; $_
New-Mailbox -Alias $_ -Password (ConvertTo-SecureString <your_password> -AsPlainText -Force) -Database <DB_Name> -DisplayName $_ -FirstName $_ -LastName $_ -Name $_ -SamAccountName $SAM -UserPrincipalName (@($_ + '@sochi-2014.ru')) -Room
Write-Host -ForegroundColor Green 'Прописываем менеджера и комментарий'
Start-Sleep 3
Set-ADUser -Identity $SAM -Manager <User> -Replace @{'info'=<комментарий>}
Write-Host -ForegroundColor Green 'Прописываем политику'
Set-CalendarProcessing -Identity $SAM -ResourceDelegates <User> -RequestInPolicy <User> -AllRequestInPolicy $TRUE
Get-Mailbox $SAM
}



#Дать права на почтовый ящик
Import-Csv d:\m.csv | %{
Add-ADPermission $_.mailbox -User $_.username -Extendedrights "Send As"
Add-MailboxPermission -Identity $_.mailbox -User $_.username -AccessRights 'FullAccess'
}

#Статистика по размерам ящиков в базах Exchange
$Bases = @(<DB_Name>,<DB_Name>)
$( ForEach ( $Base in $Bases ) { Get-Mailbox -Database $Base } ) |
Select-Object Alias, DatabaseName, @{ Name = "Size"; Expression = { $Size = Get-MailboxStatistics $_.name ; $Size.totalItemsize}} |
Sort-Object -Property Size -Descending |
Export-Csv -Path d:\stat.csv

#Информация по перемещениям почтовых ящиков со статистикой
 Get-MoveRequest | Get-MoveRequestStatistics | Select-Object -Property Alias,TotalMailboxSize,PercentComplete,BytesTransferredPerMinute |
Sort-Object Alias | Format-Table -AutoSize
#Информация по перемещениям почтовых ящиков
Get-MoveRequest | Select-Object -Property Alias,Status,SourceDatabase,TargetDatabase,@{name = 'TargetServer'; Expression = {(Get-MailboxDatabase $_.TargetDatabase).Server}} |
Sort-Object -Property SourceDatabase | Format-Table -AutoSize
#Перезапусть провалившиеся перемещения
$Users = Get-MoveRequest -MoveStatus Failed
Get-MoveRequest -MoveStatus Failed | Remove-MoveRequest
$Users | %{ New-MoveRequest $_.Alias -TargetDatabase $_.TargetDatabase }