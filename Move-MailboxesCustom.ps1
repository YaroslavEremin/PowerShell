

Function Move-MailboxesCustom {
<#
  .SYNOPSIS
  This function move mailboxes
  .DESCRIPTION
  ?
  .EXAMPLE
  Move-MailboxesCustom ECXH-MB1 1G
  .EXAMPLE
  Move-MailboxesCustom -ServerMask ECXH-MB1 -Type 1G
  #>
  [CmdLetBinding()]
  Param(
    [Parameter(Mandatory = $True, Position = 0, ValueFromPipeline = $True)]
    [String]$ServerMask,
     
    [Parameter(Mandatory = $True, Position = 1, ValueFromPipeline = $True)]
    [String]$Type
)
    #Путь куда сохраняется лог
    $LogPath = "D:\logs"
    $LogPath = $LogPath + "\Move-MailboxesCustom-$ServerMask-$Type-" + (Get-Date).tostring("yyyyMMdd") + ".txt"
    write-host -ForegroundColor Green "Logfile --> $LogPath"
    Switch -regex ($Type) {
        "200M" { $Limit = 250 }
        "1G" { $Limit = 50 }
        "2G" { $Limit = 25 }
        "5G" { $Limit = 10 }
        "20G" { $Limit = 3 }
    }
    #If (-not(Get-MoveRequest)) {
        $Bases = (Get-MailboxDatabase) | Where-Object {$_.Server -like ($ServerMask + "*") -and -not($_.Name -like "EXT-*") -and ($_.Name -like ("*" + $Type)) } | Select-Object Name,@{Name="Mailboxes";Expression={(@(Get-Mailbox -Database $_.name)).Count}} | Sort-Object -Property Mailboxes
        write-host -ForegroundColor Green "Current state"
        $Bases | Format-Table -AutoSize
        $Bases | Format-Table -AutoSize | Out-File -Encoding unicode -FilePath $LogPath -Append
        $CrowdedBases = $Bases | Where-Object { $_.Mailboxes -gt $Limit }
        $Bases = $Bases | Where-Object { $_.Mailboxes -lt $Limit }
        If ($CrowdedBases) {
            ForEach ($Base in $CrowdedBases) {
                $N = $Base.Mailboxes - $Limit
                $MovingMailboxes = Get-Mailbox -Database $Base.Name
                #$MovingMailboxes = $MovingMailboxes | Select-Object Alias,@{ Name = "Size"; Expression = { $Size = Get-MailboxStatistics $_.name ; $Size.totalItemsize}} | Sort-Object -Property Size
                $MovingMailboxes = $MovingMailboxes | Select-Object -First $N
                $MovingMailboxes.Alias | %{write-host -ForegroundColor Green $_}
                $MovingMailboxes.Alias | Out-File -Encoding unicode -FilePath $LogPath -Append
                ForEach ($Mailbox in $MovingMailboxes) {
                    $TargetDB = $Bases[$Bases.Count - 1]
                    write-host -ForegroundColor Green "Moving" $Mailbox.Alias "from crowded base" $Base.Name "to" $TargetDB.Name
                    New-MoveRequest -Identity $Mailbox.Alias -TargetDatabase $TargetDB.Name | Out-Null
                    $TargetDB.Mailboxes++
                    $Bases = $Bases | Where-Object { $_.Mailboxes -lt $Limit -and $_.Mailboxes -gt 0 } 
                }
            }
        }
        $line = "=========================================================="
        $line | Out-File -Encoding unicode -FilePath $LogPath -Append
        $line
        While ($Bases.Count -gt 1) {
            $MovingMailboxes = Get-Mailbox -Database $Bases[0].Name
            write-host -ForegroundColor Green "Next mailboxes will be moved:"
            $MovingMailboxes.Alias | %{write-host -ForegroundColor Green $_}
            $MovingMailboxes.Alias | Out-File -Encoding unicode -FilePath $LogPath -Append
            ForEach ($Mailbox in $MovingMailboxes) {
                If ($Bases.Count -gt 1) {
                try {
                $TargetDB = $Bases[$Bases.Count - 1]
                }
                catch {
                return
                }
                write-host -ForegroundColor DarkGreen "Moving" $Mailbox.Alias "to" $TargetDB.Name
                New-MoveRequest -Identity $Mailbox.Alias -TargetDatabase $TargetDB.Name  -BadItemLimit 100 -AcceptLargeDataLoss | Out-Null
                $TargetDB.Mailboxes++
                $Bases[0].Mailboxes--
                $Bases = $Bases | Where-Object { $_.Mailboxes -lt $Limit -and $_.Mailboxes -gt 0 }
                } 
            }
            $line
        }
        $Bases = (Get-MailboxDatabase) | Where-Object {$_.Server -like ($ServerMask + "*") -and -not($_.Name -like "EXT-*") -and ($_.Name -like ("*" + $Type)) } | Select-Object Name,@{Name="Mailboxes";Expression={(@(Get-Mailbox -Database $_.name)).Count}} | Sort-Object -Property Mailboxes
        write-host -ForegroundColor Green "New state"
        $Bases | Format-Table -AutoSize
        $Bases | Format-Table -AutoSize | Out-File -Encoding unicode -FilePath $LogPath -Append
        Get-MoveRequest | Select-Object -Property Alias,@{Name='SourceDB';Expression={$_.SourceDatabase}},@{Name="TargetDB";Expression={$_.TargetDatabase}},Status | Sort-Object Alias | Format-Table -AutoSize | Out-File -Encoding unicode -FilePath $LogPath -Append
}

