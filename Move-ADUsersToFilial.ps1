Function Move-ADUsersToFilial {
<#
  .SYNOPSIS
  ????
  .DESCRIPTION
  ????
  .EXAMPLE
  ????
  #>
$LogPath="D:\logs\" + (Get-Date).tostring("yyyyMMdd") + ".txt"
If( Test-Path $BackupFolder ) {
write-host -ForegroundColor Green "Logfile --> $LogPath"
$Bases = (Get-MailboxDatabase) | Where-Object {$_.Server -like 'EXCH-MB2*' -and -not($_.Name -like "EXT-*")  } | Select-Object Name,@{Name='Type';Expression={[regex]::replace($_.name,'^.*-','')}},@{Name="Mailboxes";Expression={(@(Get-Mailbox -Database $_.name)).Count}}
$Bases200M = $Bases | Where-Object {$_.Type -eq '200M'} | Sort-Object Mailboxes
$Bases1G = $Bases | Where-Object {$_.Type -eq '1G'} | Sort-Object Mailboxes
$Bases2G = $Bases | Where-Object {$_.Type -eq '2G'} | Sort-Object Mailboxes
$Bases5G = $Bases | Where-Object {$_.Type -eq '5G'} | Sort-Object Mailboxes
$Bases20G = $Bases | Where-Object {$_.Type -eq '20G'} | Sort-Object Mailboxes
$DB200M = $Bases200M[0].Name
write-host -ForegroundColor Green "$DB200M"
$DB200M | Out-File -Encoding unicode -FilePath $LogPath
$DB1G = $Bases1G[0].Name
write-host -ForegroundColor Green "$DB1G"
$DB1G | Out-File -Encoding unicode -FilePath $LogPath -Append
$DB2G = $Bases2G[0].Name
write-host -ForegroundColor Green "$DB2G"
$DB2G | Out-File -Encoding unicode -FilePath $LogPath -Append
$DB5G = $Bases5G[0].Name
write-host -ForegroundColor Green "$DB5G"
$DB5G | Out-File -Encoding unicode -FilePath $LogPath -Append
$DB20G = $Bases20G[0].Name
write-host -ForegroundColor Green "$DB20G"
$DB20G | Out-File -Encoding unicode -FilePath $LogPath -Append
$MovingUsers = Get-ADUser -SearchBase "OU=Moscow,DC=SOCHI-2014,DC=RU"-Filter {extFilial -like '*'} -Properties description,extFilial,distinguishedName | Select-Object -Property sAMAccountName,description,extFilial,distinguishedName
#$MovingUsers = Get-ADUser -Identity mfedosov -Properties description,extFilial,distinguishedName | Select-Object -Property sAMAccountName,description,extFilial,distinguishedName
$MovingUsers | Out-File -Encoding unicode -FilePath $LogPath -Append
$MovingUsers.sAMAccountName
Get-MoveRequest -MoveStatus Completed | Remove-MoveRequest
ForEach ( $User in $MovingUsers )
{
    $NewPath = [regex]::match($User.distinguishedName,'OU.*').Value
    $NewPath = [regex]::replace($NewPath,'Moscow','Sochi')
    ($User.sAMAccountName + "-->" + $NewPath)
    Out-File -InputObject ($User.sAMAccountName + "-->" + $NewPath) -Encoding unicode -FilePath $LogPath -Append
    #Get-ADUser -Identity $User.sAMAccountName | Move-ADObject -TargetPath $NewPath
    $MailboxDB = (Get-Mailbox -Identity $User.sAMAccountName).database
    Switch -RegEx ($MailboxDB)
    {
        "200M" {New-MoveRequest -Identity $User.sAMAccountName -TargetDatabase $DB200M}
        "1G" {New-MoveRequest -Identity $User.sAMAccountName -TargetDatabase $DB1G}
        "2G" {New-MoveRequest -Identity $User.sAMAccountName -TargetDatabase $DB2G}
        "5G" {New-MoveRequest -Identity $User.sAMAccountName -TargetDatabase $DB5G}
        "20G" {New-MoveRequest -Identity $User.sAMAccountName -TargetDatabase $DB20G}
    } 
}
$MoveRequests = Get-MoveRequest | Select-Object -Property Alias,Status,@{Name='SourceDB';Expression={$_.SourceDatabase}},@{Name="TargetDB";Expression={$_.TargetDatabase}} | Sort-Object Alias | Format-Table -AutoSize
$MoveRequests
$MoveRequests | Out-File -Encoding unicode -FilePath $LogPath -Append
} else {
}
write-host -ForegroundColor Red "Folder for logs does not exist"
}