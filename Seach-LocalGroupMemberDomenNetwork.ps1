function Seach-LocalGroupMemberDomenNetwork(){
param(
$GroupSID
)
function Ping ($Name){
$ping = new-object System.Net.NetworkInformation.Ping
if ($ping.send($Name).Status -eq "Success") {$True}
else {$False}
trap {Write-Verbose "Error Ping"; $False; continue}
}

[string[]]$Info
[string[]]$Computers

# Получам список компьютеров доменной сети
$Computers = Get-ADComputer -SearchBase "ou=aos,ou=servers,dc=sochi-2014,dc=ru" -Filter *  |
Select-Object name -ExpandProperty name
foreach ($Computer in $Computers){
# Проверяем доступен компьютер или нет
$Alive = Ping $Computer
if ($Alive -eq "True"){Write-Host "Seach $Computer" -BackgroundColor Blue
# Получаем имя группы
Trap {Write-Host "Error WMI $Computer";Continue}
$GroupName = Get-WmiObject win32_group -ComputerName $Computer |
Where-Object {$_.SID -eq '$GroupSID'} |
Select-Object name -ExpandProperty name
if ($GroupName){
# Получаем список членов локальной группы
Trap {Write-Host "Error ADSI $Computer";Continue}
$Users = ([ADSI]"WinNT://$Computer/$GroupName").psbase.invoke("Members") |
% {$_.GetType().InvokeMember("Name", 'GetProperty', $null, $_, $null)}
# Записываем результаты
$Info += $Users | % {$_ | Select-Object @{e={$Computer};n='Computer'},@{e={$_};n='Login'}}
}
}
}
# Вывод результатов
$Info
}