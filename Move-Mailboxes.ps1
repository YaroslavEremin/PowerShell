Function Move-Mailboxes {
<#
  .SYNOPSIS
  This function move mailboxes 
  .DESCRIPTION
  ?
  .EXAMPLE
  "user1","user2","user3" | Move-Mailboxes ECXH-MB1
  .EXAMPLE
  Get-aduser -Filter {title -like "Manager"} | Move-Mailboxes -ServerMask ECXH-MB1 -NoAction
  #>
  [CmdLetBinding()]
  Param(
    [Parameter(Mandatory = $True, Position = 0, ValueFromPipeline = $False)]
    [String]$ServerMask,
     
    [Parameter(Mandatory = $True, Position = 1, ValueFromPipeline = $True)]
    [String]$UserName,

    [Parameter(Mandatory = $False, Position = 2, ValueFromPipeline = $False)]
    [switch]$NoAction
)
BEGIN {
    $Bases = (Get-MailboxDatabase) | Where-Object {$_.Server -like ($ServerMask + '*') -and -not($_.Name -like "EXT-*")  } | Select-Object Name,@{Name='Type';Expression={[regex]::replace($_.name,'^.*-','')}},@{Name="Mailboxes";Expression={(@(Get-Mailbox -Database $_.name)).Count}}
    $Bases200M = $Bases | Where-Object {$_.Type -eq '200M'} | Sort-Object Mailboxes
    $Bases1G = $Bases | Where-Object {$_.Type -eq '1G'} | Sort-Object Mailboxes
    $Bases2G = $Bases | Where-Object {$_.Type -eq '2G'} | Sort-Object Mailboxes
    $Bases5G = $Bases | Where-Object {$_.Type -eq '5G'} | Sort-Object Mailboxes
    $Bases20G = $Bases | Where-Object {$_.Type -eq '20G'} | Sort-Object Mailboxes
    #Задание целевых баз для перемещения, с минимальным количеством ящиков
    [int]$200Mi = 0
    [int]$1Gi = 0
    [int]$2Gi = 0
    [int]$5Gi = 0
    [int]$20Gi = 0
    $DB200M = $Bases200M[$200Mi].Name
    $200Mcount = $Bases200M[$200Mi].Mailboxes
    write-host -ForegroundColor Green $DB200M " - " $200Mcount
    $DB1G = $Bases1G[$1Gi].Name
    $1Gicount = $Bases1G[$1Gi].Mailboxes
    write-host -ForegroundColor Green "$DB1G" " - " $1Gicount
    $DB2G = $Bases2G[$2Gi].Name
    $2Gcount = $Bases2G[$2Gi].Mailboxes
    write-host -ForegroundColor Green "$DB2G" " - " $2Gcount
    $DB5G = $Bases5G[$5Gi].Name
    $5Gicount = $Bases5G[$5Gi].Mailboxes
    write-host -ForegroundColor Green "$DB5G" " - " $5Gicount
    $DB20G = $Bases20G[$20Gi].Name
    $20Gicount = $Bases20G[$20Gi].Mailboxes
    write-host -ForegroundColor Green "$DB20G" " - " $20Gicount
}
PROCESS {
        $User = Get-ADUser -Identity $UserName
        $MailboxDB = (Get-Mailbox -Identity $User.sAMAccountName).database
        Switch -RegEx ($MailboxDB) {
            "200M" {
                If (!($NoAction)){ New-MoveRequest -Identity $User.sAMAccountName -TargetDatabase $DB200M } else {
                write-host -ForegroundColor Yellow "Move "$User.sAMAccountName" to "$DB200M}
                $200Mcount++
                If ($200Mcount -eq 250) {
                    $200Mi++
                    $DB200M = $Bases200M[$200Mi].Name
                    $200Mcount = $Bases200M[$200Mi].Mailboxes
                }
            }
            "1G" {
                If (!($NoAction)){ New-MoveRequest -Identity $User.sAMAccountName -TargetDatabase $DB1G } else {
                write-host -ForegroundColor Yellow "Move "$User.sAMAccountName" to "$DB1G}
                $1Gcount++
                If ($1Gcount -eq 50) {
                    $1Gi++
                    $DB1G = $Bases1G[$1Gi].Name
                    $1Gicount = $Bases1G[$1Gi].Mailboxes
                 }
            }
            "2G" {
                If (!($NoAction)){ New-MoveRequest -Identity $User.sAMAccountName -TargetDatabase $DB2G } else {
                write-host -ForegroundColor Yellow "Move "$User.sAMAccountName" to "$DB2G}
                $2Gcount++
                If ($2Gcount -eq 25) {
                    $2Gi++
                    $DB2G = $Bases2G[$2Gi].Name
                    $2Gcount = $Bases2G[$2Gi].Mailboxes
                }
            }
            "5G" {
                If (!($NoAction)){ New-MoveRequest -Identity $User.sAMAccountName -TargetDatabase $DB5G } else {
                write-host -ForegroundColor Yellow "Move "$User.sAMAccountName" to "$DB5G}
                $5Gcount++
                If ($5Gcount -eq 10) {
                    $5Gi++
                    $DB5G = $Bases5G[$5Gi].Name
                    $5Gicount = $Bases5G[$5Gi].Mailboxes
                }
            }
            "20G" {
                If (!($NoAction)){  New-MoveRequest -Identity $User.sAMAccountName -TargetDatabase $DB20G } else {
                write-host -ForegroundColor Yellow "Move "$User.sAMAccountName" to "$DB20G}
                $20Gcount++
                If ($20Gcount -eq 3) {
                    $20Gi++
                    $DB20G = $Bases20G[$20Gi].Name
                    $20Gicount = $Bases20G[$20Gi].Mailboxes
                }
            }
        }
        If (!($NoAction)){
        $UserMoveInfo = Get-MoveRequest -Identity $User.sAMAccountName | Select-Object -Property Alias,Status,SourceDatabase,TargetDatabase,@{name = 'TargetServer'; Expression = {(Get-MailboxDatabase $_.TargetDatabase).Server}}
        $UserMoveInfo | Format-Table -AutoSize
        }
}
END {

}
}