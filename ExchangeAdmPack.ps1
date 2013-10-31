#�������� �������� ����� Exchange
$m_a = <mail_address> | New-MailContact -Name $m_a �Alias $m_a -ExternalEmailAddress $m_a

#����� ���������� �� ���������� mailboxes � ����� �� �������� Exchange 
(Get-MailboxDatabase) | where {$_.Server -like <Template>} |
Select-Object Server,Name,@{Name='Type';Expression={[regex]::replace($_.name,'^.*-','')}},`
@{Name="Counter";Expression={(@(Get-Mailbox -Database $_.name)).Count}} |
Sort -Property "Type","Counter" -Descending | Format-Table -AutoSize

Get-MailboxServer | %{ Get-Mailboxstatistics -Server $_.name } |
?{ $_.DisconnectDate -ne $null } |
Select DisplayName,@{n="StoreMailboxIdentity";e={$_.MailboxGuid}},Database

#����� ����������� ���������� ��  mailboxes � ����
Get-Mailbox -Database <DB_Name> | sort-object |
Select-Object name,alias,servername,ProhibitSendQuota,IssueWarningQuota,MaxReceiveSize,`
MaxSendSize,DisplayName,Database,PrimarySmtpAddress,ProhibitSendReceiveQuota,`
@{n="Size(KB)";e = {$MBXstat = Get-MailboxStatistics $_.name; $MBXstat.totalItemsize}},`
@{n="Items"; e = {$MBXstat = Get-MailboxStatistics $_.name ; $MBXstat.itemcount; $MBXstat.storageLimitStatus}}

#���������� � ������, ����������� � �������� ����� � ����������� (��������� � Exchange 2010)
Get-MailboxDatabase -IncludePreExchange2010 |
Select-Object Name, ProhibitSendReceiveQuota, ProhibitSendQuota, IssueWarningQuota |
Sort -Property Name | Format-Table -AutoSize

#�������� � PST-���� (��������� � Exchange 2010)
New-MailboxExportRequest -Mailbox <MAilbox_Alias> -FilePath C:\Temp\<MAilbox_Alias>.pst

#��������� ���� � ��������� ������ ���������� �������� ��� (����� �������������� ������� ���������� ������ ��������� � �������� Identity, ��������� �������� ��������.
Set-MailboxDatabase -Identity DB<NN>-200M -IssueWarningQuota 180MB -ProhibitSendQuota 190MB -ProhibitSendReceiveQuota 200MB` -QuotaNotificationSchedule '��.2:00-��.2:15, ��.2:00-��.2:15, ��.2:00-��.2:15, ��.2:00-��.2:15, ��.2:00-��.2:15, ��.2:00-��.2:15, ��.2:00-��.2:15' -OfflineAddressBook '\<Address_Book>' -MailboxRetention '7.00:00:00' -DeletedItemRetention '5.00:00:00'
Set-MailboxDatabase -Identity DB<NN>-1G -IssueWarningQuota 922MB -ProhibitSendQuota 973MB -ProhibitSendReceiveQuota 1024MB -QuotaNotificationSchedule '��.2:00-��.2:15, ��.2:00-��.2:15, ��.2:00-��.2:15, ��.2:00-��.2:15, ��.2:00-��.2:15, ��.2:00-��.2:15, ��.2:00-��.2:15' -OfflineAddressBook '\<Address_Book>' -MailboxRetention '7.00:00:00' -DeletedItemRetention '5.00:00:00'
Set-MailboxDatabase -Identity DB<NN>-2G -IssueWarningQuota 1844MB -ProhibitSendQuota 1946MB -ProhibitSendReceiveQuota 2048MB -QuotaNotificationSchedule '��.2:00-��.2:15, ��.2:00-��.2:15, ��.2:00-��.2:15, ��.2:00-��.2:15, ��.2:00-��.2:15, ��.2:00-��.2:15, ��.2:00-��.2:15' -OfflineAddressBook '\<Address_Book>' -MailboxRetention '7.00:00:00' -DeletedItemRetention '5.00:00:00'
Set-MailboxDatabase -Identity DB<NN>-5G -IssueWarningQuota 4608MB -ProhibitSendQuota 4864MB -ProhibitSendReceiveQuota 5120MB -QuotaNotificationSchedule '��.2:00-��.2:15, ��.2:00-��.2:15, ��.2:00-��.2:15, ��.2:00-��.2:15, ��.2:00-��.2:15, ��.2:00-��.2:15, ��.2:00-��.2:15' -OfflineAddressBook '\<Address_Book>' -MailboxRetention '7.00:00:00' -DeletedItemRetention '5.00:00:00'
Set-MailboxDatabase -Identity DB<NN>-20G -IssueWarningQuota 18440MB -ProhibitSendQuota 19460MB -ProhibitSendReceiveQuota 20480MB -QuotaNotificationSchedule '��.2:00-��.2:15, ��.2:00-��.2:15, ��.2:00-��.2:15, ��.2:00-��.2:15, ��.2:00-��.2:15, ��.2:00-��.2:15, ��.2:00-��.2:15' -OfflineAddressBook '\<Address_Book>' -MailboxRetention '7.00:00:00' -DeletedItemRetention '5.00:00:00'

# �������� ������ ����������� �������� � PST-����
Get-MailboxExportRequest -Status Completed | Remove-MailboxExportRequest

#�������� ������� �������� ���
Get-MailboxDatabase -Status | select ServerName,Name,DatabaseSize

#��������� ������ �������� ������ � ����������� �������������� �������
Get-Mailbox -Filter { UseDatabaseQuotaDefaults -eq $False } -ResultSize Unlimited |
Select-Object Name,ServerName,Database | Format-Table -AutoSize

#��������� ������ ����� � ������ ������������/���������� ����������� �� ������������ ��������� ��������.
#����������
Get-ExchangeServer | where {$_.isHubTransportServer -eq $true -or $_.isMailboxServer -eq $true}  |
Get-MessageTrackingLog -Recipients:<mddressail_a> -EventID "RECEIVE" -Start "02/18/2013 0:00:00" -End "03/06/2013 23:59:00" |
Select-Object Timestamp,Sender,@{l="Recipients";e={$_.Recipients -join " "}},MessageSubject |
Sort-Object -Property Timestamp  | Export-Csv C:\TEMP\report.txt -encoding unicode
#������������
Get-ExchangeServer | where {$_.isHubTransportServer -eq $true -or $_.isMailboxServer -eq $true}  |
Get-MessageTrackingLog -Sender <mddressail_a> -EventID "SEND" -Start "02/18/2013 0:00:00" -End "03/06/2013 23:59:00" |
Select-Object Timestamp,Sender,@{l="Recipients";e={$_.Recipients -join " "}},MessageSubject |
Sort-Object -Property Timestamp  | Export-Csv C:\TEMP\report.txt -encoding unicode

#������� �����
Import-Csv d:\m.csv |
%{ Write-Host -ForegroundColor Green '������� ����' ; $_.mailbox
New-Mailbox -Alias $_.mailbox -Password (ConvertTo-SecureString <your_password> -AsPlainText -Force)`
-Database <DB_Name> -DisplayName $_.mailbox -FirstName $_.mailbox -LastName $_.mailbox`
-Name $_.mailbox -SamAccountName $_.mailbox -UserPrincipalName (@($_.mailbox + '@contoso.com'))
Write-Host -ForegroundColor Green '����������� ��������� � �����������' ; $_.username
Set-ADUser -Identity $_.mailbox -Manager $_.username -Replace @{'info'=<�����������>}
}

#���� ����� �� �������� ����
Import-Csv d:\m.csv | %{
Add-ADPermission $_.mailbox -User $_.username -Extendedrights "Send As"
Add-MailboxPermission -Identity $_.mailbox -User $_.username -AccessRights 'FullAccess'
}

#���������� �� �������� ������ � ����� Exchange
$Bases = @(<DB_Name>,<DB_Name>)
$( ForEach ( $Base in $Bases ) { Get-Mailbox -Database $Base } ) |
Select-Object Alias, DatabaseName, @{ Name = "Size"; Expression = { $Size = Get-MailboxStatistics $_.name ; $Size.totalItemsize}} |
Sort-Object -Property Size -Descending |
Export-Csv -Path d:\stat.csv