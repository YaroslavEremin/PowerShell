﻿Function Move-ADUsersToFilial {
<#
  .SYNOPSIS
  ????
  .DESCRIPTION
  ????
  .EXAMPLE
  ????
  #>
  BEGIN {
    #Путь куда сохраняется лог
    $LogPath = "D:\logs"
    $LogPath = $LogPath + "\" + (Get-Date).tostring("yyyyMMdd") + ".txt"
    write-host -ForegroundColor Green "Logfile --> $LogPath"
    #Создание списка баз
    $Bases = (Get-MailboxDatabase) | Where-Object {$_.Server -like 'EXCH-MB2*' -and -not($_.Name -like "EXT-*")  } | Select-Object Name,@{Name='Type';Expression={[regex]::replace($_.name,'^.*-','')}},@{Name="Mailboxes";Expression={(@(Get-Mailbox -Database $_.name)).Count}}
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
    $200Mcount = $Bases2G[$200Mi].Mailboxes
    write-host -ForegroundColor Green "$DB200M"
    $DB200M | Out-File -Encoding unicode -FilePath $LogPath
    $DB1G = $Bases1G[$1Gi].Name
    $1Gicount = $Bases2G[$1Gi].Mailboxes
    write-host -ForegroundColor Green "$DB1G"
    $DB1G | Out-File -Encoding unicode -FilePath $LogPath -Append
    $DB2G = $Bases2G[$2Gi].Name
    $2Gcount = $Bases2G[$2Gi].Mailboxes
    write-host -ForegroundColor Green "$DB2G"
    $DB2G | Out-File -Encoding unicode -FilePath $LogPath -Append
    $DB5G = $Bases5G[$5Gi].Name
    $5Gicount = $Bases2G[$5Gi].Mailboxes
    write-host -ForegroundColor Green "$DB5G"
    $DB5G | Out-File -Encoding unicode -FilePath $LogPath -Append
    $DB20G = $Bases20G[$20Gi].Name
    $20Gicount = $Bases2G[$20Gi].Mailboxes
    write-host -ForegroundColor Green "$DB20G"
    $DB20G | Out-File -Encoding unicode -FilePath $LogPath -Append
    #Создание списка пользователей для перемещения
    $MovingUsers = Get-ADUser -SearchBase "OU=Moscow,DC=SOCHI-2014,DC=RU" -Filter {extFilial -like '*'} -Properties description,extFilial,distinguishedName | Select-Object -Property sAMAccountName,description,@{Name='OrganizationalUnit';Expression={[regex]::match($_.distinguishedName,'OU.*').Value}} | Sort-Object -Property sAMAccountName
    $MovingUsers
    $MovingUsers.sAMAccountName | Get-Mailbox | Select-Object alias,servername,database,@{ Name = "Size"; Expression = { $Size = Get-MailboxStatistics $_.name ; $Size.totalItemsize}} |Format-Table -AutoSize | Out-File -Encoding unicode -FilePath $LogPath -Append
    #Удаление завершенных перемещений ящиков
    Get-MoveRequest -MoveStatus Completed | Remove-MoveRequest
}
PROCESS {
    ForEach ( $User in $MovingUsers ) {
        #Текущая база данных почтового ящика
        $MailboxDB = (Get-Mailbox -Identity $User.sAMAccountName).database
        #Перемешение ящика в Сочинскую базу соответствующего типа
        Switch -RegEx ($MailboxDB) {
            "200M" {
                New-MoveRequest -Identity $User.sAMAccountName -TargetDatabase $DB200M
                $200Mcount++
                If ($200Mcount -eq 250) {
                    $200Mi++
                    $DB200M = $Bases200M[$200Mi].Name
                    $200Mcount = $Bases2G[$200Mi].Mailboxes
                }
            }
            "1G" {
                New-MoveRequest -Identity $User.sAMAccountName -TargetDatabase $DB1G
                 $1Gcount++
                 If ($1Gcount -eq 50) {
                    $1Gi++
                    $DB1G = $Bases2G[$1Gi].Name
                    $1Gicount = $Bases2G[$1Gi].Mailboxes
                 }
            }
            "2G" {
                New-MoveRequest -Identity $User.sAMAccountName -TargetDatabase $DB2G
                $2Gcount++
                If ($2Gcount -eq 25) {
                    $2Gi++
                    $DB2G = $Bases2G[$2Gi].Name
                    $2Gcount = $Bases2G[$2Gi].Mailboxes
                }
            }
            "5G" {
                New-MoveRequest -Identity $User.sAMAccountName -TargetDatabase $DB5G
                $5Gcount++
                If ($5Gcount -eq 10) {
                    $5Gi++
                    $DB5G = $Bases5G[$5Gi].Name
                    $5Gicount = $Bases2G[$5Gi].Mailboxes
                }
            }
            "20G" {
                New-MoveRequest -Identity $User.sAMAccountName -TargetDatabase $DB20G
                $20Gcount++
                If ($20Gcount -eq 3) {
                    $20Gi++
                    $DB20G = $Bases20G[$20Gi].Name
                    $20Gicount = $Bases2G[$20Gi].Mailboxes
                }
            }
        }
        #Задание целевого OU для перемещения
        $NewPath = [regex]::replace($User.OrganizationalUnit,'Moscow','Sochi')
        ($User.sAMAccountName + "-->" + $NewPath)
        Out-File -InputObject ($User.sAMAccountName + "-->" + $User.OrganizationalUnit) -Encoding unicode -FilePath $LogPath -Append
        Out-File -InputObject ($User.sAMAccountName + "-->" + $NewPath) -Encoding unicode -FilePath $LogPath -Append
        #Перемещение учетной записи в Сочинский OU
        Get-ADUser -Identity $User.sAMAccountName | Move-ADObject -TargetPath $NewPath
    }
    #Создание списка пользователей перемещение ящиков которых окончилось неудачей
    $FailedUsers = (Get-MoveRequest -MoveStatus Failed).Alias
    #Удаление завершившихся неудачей перемещений ящиков
    Get-MoveRequest -MoveStatus Failed | Remove-MoveRequest
    ForEach ( $User in $FailedUsers ) {
        #Перемешение ящика в Сочинскую базу соответствующего типа
        Switch -RegEx ($User.SourceDatabase) {
            "200M" {
                New-MoveRequest -Identity $User.Alias -TargetDatabase $DB200M
                $200Mcount++
                If ($200Mcount -eq 250) {
                    $200Mi++
                    $DB200M = $Bases200M[$200Mi].Name
                    $200Mcount = $Bases2G[$200Mi].Mailboxes
                }
            }
            "1G" {
                New-MoveRequest -Identity $User.Alias -TargetDatabase $DB1G
                 $1Gcount++
                 If ($1Gcount -eq 50) {
                    $1Gi++
                    $DB1G = $Bases2G[$1Gi].Name
                    $1Gicount = $Bases2G[$1Gi].Mailboxes
                 }
            }
            "2G" {
                New-MoveRequest -Identity $User.Alias -TargetDatabase $DB2G
                $2Gcount++
                If ($2Gcount -eq 25) {
                    $2Gi++
                    $DB2G = $Bases2G[$2Gi].Name
                    $2Gcount = $Bases2G[$2Gi].Mailboxes
                }
            }
            "5G" {
                New-MoveRequest -Identity $User.Alias -TargetDatabase $DB5G
                $5Gcount++
                If ($5Gcount -eq 10) {
                    $5Gi++
                    $DB5G = $Bases5G[$5Gi].Name
                    $5Gicount = $Bases2G[$5Gi].Mailboxes
                }
            }
            "20G" {
                New-MoveRequest -Identity $User.Alias -TargetDatabase $DB20G
                $20Gcount++
                If ($20Gcount -eq 3) {
                    $20Gi++
                    $DB20G = $Bases20G[$20Gi].Name
                    $20Gicount = $Bases2G[$20Gi].Mailboxes
                }
            }
        }
    }
}
END {
    $MoveRequests = Get-MoveRequest | Select-Object -Property Alias,@{Name='SourceDB';Expression={$_.SourceDatabase}},@{Name="TargetDB";Expression={$_.TargetDatabase}},Status | Sort-Object Alias | Format-Table -AutoSize
    $MoveRequests
    $MoveRequests | Out-File -Encoding unicode -FilePath $LogPath -Append
    $MovingUsers | % {Get-ADUser -Identity $_.sAMAccountName
    <##Отправка отчета по почте
    If ($MovingUsers.Count -ge 1 -or $FailedUsers.Count -ge 1) {
        $MailSubject = "Отчет по перемещеннию учетных записей пользователей  в Сочинский OU"
        #$MailBody = Get-Content -Path $LogPath
        $MailBody = Get-Content -Path "D:\logs\20131129.txt"
        $MailBody = [regex]::replace($MailBody,'$','<br>') | Out-String
        Send-MailMessage -From ([regex]::replace($env:USERNAME,'-adm','') + "@sochi2014.com") -Subject $MailSubject -To 'yeremin@SOCHI2014.COM' -Body $MailBody -Encoding Unicode -SmtpServer exch-cas02
        write-host -ForegroundColor Green "Done"
    }#>
}
}