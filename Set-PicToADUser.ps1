function Set-PicToADUser {
param(
[Parameter(Mandatory=$true)]$UserName,
[Parameter(Mandatory=$true)]$ExchSrv
)
$Session = New-PSSession -ConnectionUri ("http://" + $ExchSrv +"/Powershell/") -Authentication Kerberos -ConfigurationName Microsoft.Exchange -ErrorAction SilentlyContinue
Import-PSSession $Session -AllowClobber | Out-Null
write-host -ForegroundColor Green "Looking for mailbox server"
[int]$Counter = 0
Get-MailboxServer | % {
If ($_.Name -match $ExchSrv) {
$ExchSrv = $_.Name
$Counter = $Counter + 1
$ExchSrv
}
}
If ($Counter -eq 0) {
write-host -ForegroundColor Red "Can`t find mailbox server.Please type corrent value"
write-host -ForegroundColor Red "Use default falue"
$ExchSrv = <Exchange server>
$ExchSrv
}
ElseIf ($Counter -gt 1) {
write-host -ForegroundColor Red "Faund more then one server"
}
$UserDesktop = New-Object –com Shell.Application
$UserDesktopPath = ($UserDesktop.namespace(0x10)).Self.Path
write-host -ForegroundColor Green 'Path to picture file is:'
$path = $UserDesktopPath + '\' + $UserName + '.jpg'
$path
write-host -ForegroundColor Green 'Set picture'
Import-RecipientDataProperty -Identity $UserName -Picture -FileData ([Byte[]]$(Get-Content -path $path -Encoding Byte -ReadCount 0))
Remove-PSSession $Session
}
write-host  -ForegroundColor Green 'Требования к файлу изображения: формат - JPG, размер - меньше 10Кб, файл должен лежать в папке "c:\user_pics", имя файла должно совпадать с логином пользователя в AD'
write-host  -ForegroundColor Green 'Введите логин пользователя.'
Set-PicToADUser