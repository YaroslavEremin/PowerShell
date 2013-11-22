Function Clear-Eventlogs {            
 Param(
  $Computername = $ENV:COMPUTERNAME,
  [array]$EventLogs = @("application","security"),
  $BackupFolder = "C:\BackupEventLogs\"
  )
 Foreach ( $i in $EventLogs ) {
 If(!( Test-Path $BackupFolder )) { New-Item $BackupFolder -Type Directory }
 $eventlog="c:\BackupEventLogs\$i" + (Get-Date).tostring("yyyyMMdd") + ".evt"
 (get-wmiobject win32_nteventlogfile -ComputerName $computername |
  Where {$_.logfilename -eq "$i"}).backupeventlog($eventlog)
 Clear-EventLog -LogName $i
 }# end Foreach
}#end function