
#Экспорт списка пользователей из группы в TXT-файл
Get-ADGroupMember <GroupName> | Select-Object -Property sAMAccountName |
Sort -Property sAMAccountName | Out-File -FilePath d:\ADGroupMember.txt

#Экспорт списка пользователей из группы в CSV-файл
Get-ADGroupMember <GroupName>  | Select-Object -Property sAMAccountName |
Sort -Property sAMAccountName | Export-Csv -Delimiter ";" -Path d:\ADGroupMember.csv

#Вывод в файл списка email адресов пользователей членов группы
$Users = Get-ADGroupMember <GroupName>
$(foreach ($CurrentUser in $Users) {
Get-ADUser -Identity $CurrentUser -Properties mail |  Select-Object -Property mail | Sort -Property mail 
}) | Out-File -FilePath d:\ADGroupMember.txt

#Добавить членов группы в другую группу:
Add-ADGroupMember <GroupName1> -members (Get-ADGroupMember <GroupName2>) -PassThru

#Добавить членов подразделения в группу:
Get-ADUser -SearchBase "OU=Staff,DC=contoso,DC=com" -Filter * | % {Add-ADGroupMember -Identity <GroupName> -Members $_}

#Вывод в файл списка пользователе с отбором по параметру:
Get-ADUser -Filter {description -eq <your_param>} -Properties description,displayname,userPrincipalName |
Select-Object -Property description,displayname,userPrincipalName |
Export-Csv -Delimiter ";" -Path d:\Userlist.csv -encoding "unicode"

#Поиск пользователей у которых не заполнен атрибут Организация
Get-ADUser -SearchBase "OU=Staff,DC=contoso,DC=com" -Filter {-not (company -like "*")} -Properties description,sAMAccountName |
Sort-Object -Property description | Format-Table -AutoSize | Out-File -FilePath d:\Userlist.txt

#Поиск пользователей у которых стоит галочка "Сменить пароль при следующем входе"
Get-ADUser -Filter { pwdLastSet -eq 0 } -SearchBase "OU=Users,DC=contoso,DC=com" |
Select-Object -Property sAMAccountName | Sort-Object sAMAccountName | Out-File d:\users.txt 

#Поиск пользователей по списку
$UserDesktop = New-Object –com Shell.Application
$UserDesktopPath = ($UserDesktop.namespace(0x10)).Self.Path
$Users = Import-Csv -Delimiter ',' ($UserDesktopPath + '\Userlist.csv')
$(Foreach($CurrentUser in $Users){
$description = "*" + $CurrentUser.description  + "*"
Get-ADUser -Filter { description -like $description } -Properties description,sAMAccountName |
Select-Object -Property description,sAMAccountName
}) | Export-Csv -Delimiter ";" -Path ($UserDesktopPath + '\Users.csv') -Encoding unicode

#Добавление группы пользователей из списка в группу АД
$UserDesktop = New-Object –com Shell.Application
$UserDesktopPath = ($UserDesktop.namespace(0x10)).Self.Path
$Users = Import-Csv -Delimiter ";" ($UserDesktopPath + '\Users.csv')
$(Foreach($CurrentUser in $Users){
$sAMAccountName = $CurrentUser.sAMAccountName
Add-ADGroupMember <GroupName> -members $sAMAccountName -PassThru
}) 

#Поиск по списку, для проверки отключенна ли учетная запись (512 - Active, 514 - Disable)
$UserDesktop = New-Object –com Shell.Application
$UserDesktopPath = ($UserDesktop.namespace(0x10)).Self.Path
$Users = Import-Csv -Delimiter ";" ($UserDesktopPath + '\Users.csv')
$(Foreach($CurrentUser in $Users){
$description = "*" + $CurrentUser.description  + "*"
Get-ADUser -Filter { description -like $description } -Properties description,sAMAccountName,useraccountcontrol |
Select-Object -Property description,sAMAccountName,useraccountcontrol
}) | Out-File -FilePath d:\Users.txt -Encoding unicode 


#Поиск пользователей входящих в группу локальных администраторов на серверах Windows
$UserDesktop = New-Object –com Shell.Application
$UserDesktopPath = ($UserDesktop.namespace(0x10)).Self.Path
Get-ADComputer -Filter { operatingSystem -like "*server*" } -SearchBase "OU=Servers,DC=contoso,DC=com" -Properties name,operatingSystem,distinguishedName |
Select-Object -Property name,operatingSystem,distinguishedName | Export-Csv -Delimiter ";" -Path ($UserDesktopPath + '\Servers.csv') -Encoding unicode 
$Computers = Import-Csv -Delimiter ";" ($UserDesktopPath + '\Servers.csv')
function Get-Admins {
Foreach ($Computer in $Computers){
        $ComputerName = $Computer.Name
        $GroupName = Get-WmiObject win32_group -ComputerName $ComputerName | ? {$_.SID -eq 'S-1-5-32-544'} | select name -ExpandProperty name
        $LocalGroup =[ADSI]"WinNT://$ComputerName/$GroupName"
        $GroupMembers = @($LocalGroup.psbase.Invoke("Members"))
        $Members = $GroupMembers | foreach {$_.GetType().InvokeMember("Name", 'GetProperty', $null, $_, $null)}
        foreach ($Member in $Members){
        $obj = New-Object System.Object
        $obj | Add-Member -MemberType NoteProperty -Name "Computer" -Value $ComputerName
        $obj | Add-Member -MemberType NoteProperty -Name "AdminGroup" -Value $GroupName
        $obj | Add-Member -MemberType NoteProperty -Name "AdminGroupMembers" -Value $Member
        $obj | Add-Member -MemberType NoteProperty -Name "OperatingSystem" -Value $Computer.operatingSystem
        $obj | Add-Member -MemberType NoteProperty -Name "DistinguishedName" -Value $Computer.distinguishedName
        $obj
        }
 #   }
} 
}
Get-Admins | Export-Csv -Delimiter ";" -Path C:\AdminList.csv -Encoding unicode 

