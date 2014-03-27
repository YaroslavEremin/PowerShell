<#----------------------------------------------------------------------------------------------------------
Основные служебные функции скрипта
----------------------------------------------------------------------------------------------------------#>

Function Clear-FiredUsersMain {

    Set-ScriptParameters
    Write-Log "Sochi 2014---------< Clear-FiredUsers >---------@Yaroslav Eremin"
    Write-Log "-------------------------------------------yaroslav@itwonline.ru"
    $Session = Create-PSSessionToExchange $Script:ExchangeServer


    If ( !($global:Workflow) ) {

        $global:Workflow = @{}
        $global:Workflow["Search"] = @()
        $global:Workflow["Profile"] = @()
        $global:Workflow["CleanUP"] = @()
        $global:Workflow["MoveTo"] = @()
        $SAMs = find-users $Script:DaysBeforeCleanUp $Script:SearchBase

        ForEach ($SAM in $SAMs) {

            $global:Workflow["Search"] += $SAM

        }

    }

    Show-ScriptMenu
    Finalize-Script $Session


}

#Проверено
Function Set-ScriptParameters {

    #$ScriptName = $MyInvocation.MyCommand.Name
    $ScriptName = "Clear-FiredUsers.ps1"
    #$CurrentDirectory = Get-ScriptDirectory
    #$ConfigurationFileName = [regex]::Replace($ScriptName,'\..*$','.ini')
    #$PathToConfigurationFile = $CurrentDirectory + "\" + $ConfigurationFileName
    $PathToConfigurationFile = "D:\GitHub\PowerShell\Clear-FiredUsers\Clear-FiredUsers.ini"
    $ConfigurationFileContent = Get-IniContent $PathToConfigurationFile

    #Путь к лог файлу
    If ($ConfigurationFileContent["MAIN"]["path_to_log_file"]) {
    
        $Script:PathToLogFile = $ConfigurationFileContent["MAIN"]["path_to_log_file"] + "\" + $ScriptName + "-" + (Get-Date -format yyyy-MM-dd_HH-mm-ss) + ".log"

    } Else {

        $Script:PathToLogFile = $CurrentDirectory + "\" + $ScriptName + "-" + (Get-Date -format yyyy-MM-dd_HH-mm-ss) + ".log"

    }

    #Путь к временному файлу
    If ($ConfigurationFileContent["MAIN"]["path_to_temp_workflow_file"]) {
    
        $Script:PathToTempWorkflowFile = $ConfigurationFileContent["MAIN"]["path_to_temp_workflow_file"]

    } Else {

        $TempWorkflowFileName = [regex]::Replace($ScriptName,'\..*$','.tmp')
        $Script:PathToTempWorkflowFile = $PathToLogFile = $CurrentDirectory + "\" + $TempWorkflowFileName

    }

    #Имя Exchange сервера для выполнения удаленных команд
    If ($ConfigurationFileContent["MAIN"]["exchange_server"]) {
    
        $Script:ExchangeServer = $ConfigurationFileContent["MAIN"]["exchange_server"]

    } Else {
        
        $ExchangeServers = $(Get-ADGroupMember 'Exchange Servers' | Where-Object {$_.objectclass -eq "computer"} | ForEach-Object {get-adcomputer $_.name} | ForEach-Object {$_.DNSHostName})
        $Script:ExchangeServer = Get-Random $ExchangeServers

    }

    #Количество дней перед очисткой заблокированных пользователей
    If ($ConfigurationFileContent["SEARCH"]["days_before_clean_up"]) {
    
        $Script:DaysBeforeCleanUp = [int]$ConfigurationFileContent["SEARCH"]["days_before_clean_up"]

    } Else {

        [int]$Script:DaysBeforeCleanUp = 3

    }

    #В каких подразделениях искать заблокированных пользоватей
    If ($ConfigurationFileContent["SEARCH"]["search_base"]) {
    
        [array]$Script:SearchBase = $ConfigurationFileContent["SEARCH"]["search_base"] -split ";"

    } Else {

        $Script:SearchBase = @((Get-ADDomain).UsersContainer)

    }

    #Путь для экспорта почтовых ящиков в PST файлы
    If ($ConfigurationFileContent["EXPORT"]["path_to_PST_files"]) {
    
        $Script:PathToPSTFiles = $ConfigurationFileContent["EXPORT"]["path_to_PST_files"]

    } Else {

        $Script:PathToPSTFiles = $CurrentDirectory

    }

    #Путь к папке куда логируются входы пользователей
    If ($ConfigurationFileContent["PROFILES"]["path_to_users_logon"]) {
    
        $Script:PathToUsersLogon = $ConfigurationFileContent["PROFILES"]["path_to_users_logon"]

    } Else {

        $Script:PathToUsersLogon = $CurrentDirectory

    }

    #Шаблон для поиска компьютеров на которые заходил пользователь
    If ($ConfigurationFileContent["PROFILES"]["delete_profiles_search_wildcard"]) {
    
        $Script:DeleteProfilesSearchWildcard = $ConfigurationFileContent["PROFILES"]["delete_profiles_search_wildcard"]

    } Else {

        $Script:DeleteProfilesSearchWildcard = "*"

    }

    #Путь к папке куда логируются входы пользователей
    If ($ConfigurationFileContent["CLEANUP"]["path_to_back_up_files"]) {
    
        $Script:PathToBackUpFiles = $ConfigurationFileContent["CLEANUP"]["path_to_back_up_files"]

    } Else {

        $Script:PathToBackUpFiles = $CurrentDirectory + "\BackUp"

    }

    #Список очищаемых атрибутов
    If ($ConfigurationFileContent["CLEANUP"]["clearing_attributes"]) {
    
        [array]$Script:ClearingAttributes = $ConfigurationFileContent["CLEANUP"]["clearing_attributes"] -split ","

    } Else {

        $Script:ClearingAttributes = @("MemberOf")

    }

    #Список очищаемых атрибутов
    If ($ConfigurationFileContent["MOVETO"]["OU_for_disabled_users"]) {
    
        $Script:OUForDisabledUsers = $ConfigurationFileContent["MOVETO"]["OU_for_disabled_users"]

    } Else {

        $Script:OUForDisabledUsers = "OU=Disabled Users," + @((Get-ADDomain).DistinguishedName)

    }

}

Function Show-ScriptMenu {

    Param (
        
        [Parameter(Mandatory = $False, Position = 0)] $MenuPosition = "m",

        [Parameter(Mandatory = $False)] [switch]$UnexpectedInput = $False,

        [Parameter(Mandatory = $False)] [switch]$RuntimeError = $False

    )

    $ExitFlag = $False

    do {

        Clear-Host

        If ($RuntimeError) {

            Write-Host "Runtime Error"
            write-log $Error[0].Exception
            write-log $Error[0].InvocationInfo.Line
            Write-Host ""

        }
        Write-Host "Sochi 2014---------< Clear-FiredUsers >---------@Yaroslav Eremin"
        Write-Host "-------------------------------------------yaroslav@itwonline.ru"
    
        If ($UnexpectedInput) {
                
            Write-Host "Unexpected Input. Tipe right value" -ForegroundColor Yellow
            $UnexpectedInput = $False

        }

        Switch ($MenuPosition) {

            m {

                Write-Host "----------------------< Main menu >-----------------------------" -ForegroundColor Green
                Write-Host "1. Find users to clean up" -ForegroundColor Green
                Write-Host "2. Export mailboxes" -ForegroundColor Green
                Write-Host "3. Delete profiles" -ForegroundColor Green
                Write-Host "4. Clean up attributes" -ForegroundColor Green
                Write-Host "5. Move to OU for disabled users" -ForegroundColor Green
                Write-Host "6. Exit" -ForegroundColor Green
                $MenuPosition = Read-Host -Prompt "Select"

                If ($MenuPosition -notmatch "^[1-6,m]$") {

                    $MenuPosition = "m"
                    $UnexpectedInput = $True
                }

            }

            1 {

                $global:Workflow["Search"]
                Write-Host "------------< Main menu -> Find users to clean up >-------------" -ForegroundColor Green
                Write-Host "1. Add 10 users to mailbox export" -ForegroundColor Green
                Write-Host "2. Add all users to mailbox export" -ForegroundColor Green
                Write-Host "3. Back" -ForegroundColor Green
                $Input = Read-Host -Prompt "Select"

                Switch ($Input) {

                    1 {
                        
                        Export-MailboxesToPST ($global:Workflow["Search"] | select -First 10)

                    }

                    2 {

                        Export-MailboxesToPST ($global:Workflow["Search"])

                    }

                    3 {

                        $MenuPosition = "m"

                    }

                    default {

                        $UnexpectedInput = $True

                    }

                }

            }

            2 {

                Check-ExportState
                Write-Host "------------< Main menu -> Export mailboxes >-------------" -ForegroundColor Green
                Write-Host "1. Check state" -ForegroundColor Green
                Write-Host "2. Resume failed" -ForegroundColor Green
                Write-Host "3. Back" -ForegroundColor Green
                $Input = Read-Host -Prompt "Select"

                Switch ($Input) {

                    1 {
                        
                        Disable-MailboxWithCompletedExport

                    }

                    2 {

                        Restart-FailedMailboxExportRequests

                    }

                    3 {

                        $MenuPosition = "m"

                    }

                    default {

                        $UnexpectedInput = $True

                    }

                }
            }

            3 {

                $global:Workflow["Profile"]
                Write-Host "------------< Main menu -> Delete profiles >--------------" -ForegroundColor Green
                Write-Host "1. Delete profiles" -ForegroundColor Green
                Write-Host "2. Back" -ForegroundColor Green
                $Input = Read-Host -Prompt "Select"

                Switch ($Input) {

                    1 {
                        
                        Delete-Profiles $global:Workflow["Profile"]

                    }

                    2 {

                        $MenuPosition = "m"

                    }

                    default {

                        $UnexpectedInput = $True

                    }

                }

            }

            4 {

                $global:Workflow["CleanUP"]
                Write-Host "-----------< Main menu -> Clean up attributes >-----------" -ForegroundColor Green
                Write-Host "1. Clean up attributes" -ForegroundColor Green
                Write-Host "2. Back" -ForegroundColor Green
                $Input = Read-Host -Prompt "Select"

                Switch ($Input) {

                    1 {
                        
                        Clean-UsersAttributes $Script:PathToBackUpFiles $global:Workflow["CleanUP"] $Script:ClearingAttributes

                    }

                    2 {

                        $MenuPosition = "m"

                    }

                    default {

                        $UnexpectedInput = $True

                    }

                }

            }

            5 {

                $global:Workflow["MoveTo"]
                Write-Host "-------< Main menu -> Move to OU for disabled users >------" -ForegroundColor Green
                Write-Host "1. Move to OU for disabled users" -ForegroundColor Green
                Write-Host "2. Back" -ForegroundColor Green
                $Input = Read-Host -Prompt "Select"

                Switch ($Input) {

                    1 {
                        
                        Move-ADUsersToOUForDisabledUsers $Script:OUForDisabledUsers $global:Workflow["MoveTo"]

                    }

                    2 {

                        $MenuPosition = "m"

                    }

                    default {

                        $UnexpectedInput = $True

                    }

                }

            }


            6 {

                Write-Host "-----------------< Main menu -> Exit >-------------------" -ForegroundColor Green
                Write-Host "1. Exit" -ForegroundColor Green
                Write-Host "2. Back" -ForegroundColor Green
                $Input = Read-Host -Prompt "Select"

                Switch ($Input) {

                    1 {
                        
                        $ExitFlag = $True

                    }

                    2 {

                        $MenuPosition = "m"

                    }

                    default {

                        $UnexpectedInput = $True

                    }

                }

            }

        }

    } Until ($ExitFlag)

}

#Проверено
#Логирование действий
function Write-Log {

    Param ( 

        [Parameter( Mandatory = $True, Position = 0, ValueFromPipeline = $True )]
        [String]$Message

     )
        
        Write-Host $Message
        $Message = ( Get-date ).ToString( "T" ) + "  " + $Message
        Out-File $Script:PathToLogFile -InputObject $Message -Append

}


<#----------------------------------------------------------------------------------------------------------
Дополнительные служебные функции скрипта 
----------------------------------------------------------------------------------------------------------#>

#Проверено
Function Get-IniContent { 
    <# 
    .Synopsis 
        Gets the content of an INI file 
         
    .Description 
        Gets the content of an INI file and returns it as a hashtable 
         
    .Notes 
        Author    : Oliver Lipkau <oliver@lipkau.net> 
        Blog      : http://oliver.lipkau.net/blog/ 
        Date      : 2010/03/12 
        Version   : 1.0 
         
        #Requires -Version 2.0 
         
    .Inputs 
        System.String 
         
    .Outputs 
        System.Collections.Hashtable 
         
    .Parameter FilePath 
        Specifies the path to the input file. 
         
    .Example 
        $FileContent = Get-IniContent "C:\myinifile.ini" 
        ----------- 
        Description 
        Saves the content of the c:\myinifile.ini in a hashtable called $FileContent 
     
    .Example 
        $inifilepath | $FileContent = Get-IniContent 
        ----------- 
        Description 
        Gets the content of the ini file passed through the pipe into a hashtable called $FileContent 
     
    .Example 
        C:\PS>$FileContent = Get-IniContent "c:\settings.ini" 
        C:\PS>$FileContent["Section"]["Key"] 
        ----------- 
        Description 
        Returns the key "Key" of the section "Section" from the C:\settings.ini file 
         
    .Link 
        Out-IniFile 
    #> 
     
    [CmdletBinding()] 
    Param( 
        [ValidateNotNullOrEmpty()] 
        [ValidateScript({(Test-Path $_) -and ((Get-Item $_).Extension -eq ".ini")})] 
        [Parameter(ValueFromPipeline=$True,Mandatory=$True)] 
        [string]$FilePath 
    ) 
     
    Begin 
        {Write-Verbose "$($MyInvocation.MyCommand.Name):: Function started"} 
         
    Process 
    { 
        Write-Verbose "$($MyInvocation.MyCommand.Name):: Processing file: $Filepath" 
             
        $ini = @{} 
        switch -regex -file $FilePath 
        { 
            "^\[(.+)\]$" # Section 
            { 
                $section = $matches[1] 
                $ini[$section] = @{} 
                $CommentCount = 0 
            } 
            "^(;.*)$" # Comment 
            { 
                if (!($section)) 
                { 
                    $section = "No-Section" 
                    $ini[$section] = @{} 
                } 
                $value = $matches[1] 
                $CommentCount = $CommentCount + 1 
                $name = "Comment" + $CommentCount 
                $ini[$section][$name] = $value 
            }  
            "(.+?)\:(.*)" # Key 
            { 
                if (!($section)) 
                { 
                    $section = "No-Section" 
                    $ini[$section] = @{} 
                } 
                $name,$value = $matches[1..2] 
                $ini[$section][$name] = $value 
            } 
        } 
        Write-Verbose "$($MyInvocation.MyCommand.Name):: Finished Processing file: $path" 
        Return $ini 
    } 
         
    End 
        {Write-Verbose "$($MyInvocation.MyCommand.Name):: Function ended"} 
}

#Данная функция возвращает каталог, в котором хранится наш скрипт
#Таботает тупо, переписать
function Get-ScriptDirectory {

    $Invocation = ( Get-Variable MyInvocation -Scope 2 ).Value
    Split-Path $Invocation.MyCommand.Path
    
}
#Проверено
function Create-PSSessionToExchange {

    Param ( 

        [Parameter( Mandatory = $True, Position = 0)] [String]$ExchServer

     )

    write-log "Попытка создать PSSession до сервера $ExchServer"
    $ExchSession = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri ( "http://" + $ExchServer + "/Powershell/" ) -authentication Kerberos -ErrorAction SilentlyContinue
    Import-PSSession $ExchSession -AllowClobber | Out-Null
    if ( $ExchSession ) {

        write-log  "Удаленная PS-сессия создана";
        Return $ExchSession

    } else {

        write-log  "Удаленная PS-сессия не создана"
        exit

    }




}

#Проверено
#Закрытие удаленной сессии PS и завершение работы скрипта.
function Finalize-Script {

    Param ( 

        [Parameter( Mandatory = $True, Position = 0)] $ExchSession

     )

    if ( $ExchSession ) {

        Remove-PSSession $ExchSession

    }

    write-log  "Закрываем PSSession до сервера Exchange"
    write-log  "Скрипт завершил работу"
    exit

}


<#----------------------------------------------------------------------------------------------------------
Основные рабочие функции скрипта
----------------------------------------------------------------------------------------------------------#>


#Проверено
#Поиск пользователей в целевых OU
function find-users {

    Param ( 

        [Parameter( Mandatory = $True, Position = 0)] $Delay,
        [Parameter( Mandatory = $True, Position = 1)] [array]$OUForSearch

     )

    Trap {

        continue

    }
    $Delay = - $Delay
    $Date = (Get-Date)
    $Date = $Date.AddDays($Delay)
    $Users = $NULL

    ForEach ($OU in $OUForSearch) {

        $Users += Get-ADUser -SearchBase $OU -Filter * -Properties description,accountExpirationDate,modifyTimeStamp,useraccountcontrol
        
    }
    
    $Users = $Users | ?{ ($_.accountExpirationDate -lt $Date -and $_.accountExpirationDate -ne $NULL) -or ($_.modifyTimeStamp -lt $Date -and $_.useraccountcontrol -eq 514) }
    $Users = $Users | select  description,SamAccountName,accountExpirationDate,modifyTimeStamp | sort accountExpirationDate,modifyTimeStamp -Descending

    If ( $Users ) {

        Write-Log "Users found"
        Return $Users.SamAccountName

    } else {

        Write-Log "Users not found"
        Exit

    }

}

#Запуск экспорта почтовых ящиков
function Export-MailboxesToPST {

    Param (

        [Parameter(Mandatory = $TRUE, Position = 0)] [array]$Users

    )

    Trap {

        continue

    }

    ForEach ($User in $Users) {

        $UserMailbox = $Null
        $ADUser = Get-ADUser -Identity $User -Properties mail

        If ( $ADUser.mail ) {

            $UserMailbox = Get-Mailbox $( $ADUser.mail )  -ErrorAction SilentlyContinue

        }

        if ( $UserMailbox -and $UserMailbox.Database -match "200M") {
            
            write-log "$User is not important person. Just disable mailbox."
            Disable-Mailbox $User -Confirm:$False -ErrorAction SilentlyContinue
            $global:Workflow["Profile"] += @($User)
            
        } ElseIf ( $UserMailbox ) {
              
            write-log  ( "Начинаем экспорт почтового ящика " + $ADUser.SamAccountName + " в файл." )
            New-MailboxExportRequest -Mailbox $UserMailbox.Alias -FilePath ( "\\exch-mb27\e$\PST\" + $ADUser.sAMAccountName + ".pst" )3 | Out-Null
        
        } else {

            write-log ( "У пользователя " + $ADUser.SamAccountName + " нет почтового ящика" )
            Set-ADUser -Identity $ADUser.SamAccountName -EmailAddress $NULL
            $global:Workflow["Profile"] += @($User)
            
        }

        $global:Workflow["Search"] = $global:Workflow["Search"] | ?{ $_ -ne $User }
        
    }

}

function Disable-MailboxWithCompletedExport {

    Trap {

        continue

    }

    $CompletedExportRequests = (Get-MailboxExportRequest  | ? { $_.Status -match "Completed" })
    $CompletedExportRequests
    ForEach ( $CompletedExportRequest in $CompletedExportRequests ) {
        $SAM =$null
        $SAM = [regex]::Replace($CompletedExportRequest.FilePath,'^.*\\','')
        $SAM = [regex]::Replace($SAM,'\.pst','')
        write-log ( "Экспорт почтового ящика пользователя " + $SAM + " в PST завершен." )
        $CompletedExportRequest | Remove-MailboxExportRequest -WhatIf:$debugMode -Confirm:$False
        Start-Sleep 5
        write-log "Отключаем почтовый ящик"
        Disable-Mailbox $SAM -Confirm:$False -WhatIf:$debugMode -ErrorAction SilentlyContinue
        $global:Workflow["Profile"] += @($SAM)

    }

    (Get-MailboxExportRequest  | ? { $_.Status -match "Completed" }) | Remove-MailboxExportRequest -Confirm:$False

    If ($CompletedExportRequests) {

        $CompletedExportRequests | Select-Object -Property @{name = 'Name'; Expression = {[regex]::match($_.Mailbox,'[a-z,A-Z,\s]*$').Value}},SourceDatabase,Status |
        Sort-Object -Property Status,Identity | Format-Table -AutoSize
        Start-Sleep 7

    }

}

#Удаление профилей пользователя на найденных компьютерах
function Delete-Profiles {

    Param (

        [Parameter(Mandatory = $TRUE, Position = 0)] [array]$Users

    )

    ForEach ($User in $Users) {

        write-log "Начинаем поиск профилей пользователя: $User"
        $computers = @()
        $computers = Find-Profiles $Script:PathToUsersLogon $Script:DeleteProfilesSearchWildcard $User
        $global:Workflow["Profile"] = $global:Workflow["Profile"] | ?{ $_ -ne $User }
        $global:Workflow["CleanUP"] += @($User)

        If ($computers.Count -ne 0) {

            write-log  "Начинаем удаление профилей пользователя: $User"

            ForEach ($computer in $computers) {

                $Profile = $null

                Try {

                    $Profile = Get-WmiObject Win32_UserProfile -computer $computer -filter "localpath='C:\\Users\\$SAM'"

                } catch {
        
                     write-log $Error[0].Exception
                     write-log $Error[0].InvocationInfo.Line

                }
        
                If ($Profile) {

                    write-log "Profile on $computer is found. Removing"

                    If ( !($debugMode) ) {

                        Try {

                            $Profile.delete()

                        } Catch {

                            write-log $Error[0].Exception
                            write-log $Error[0].InvocationInfo.Line

                        }

                    }

                } else {

                    write-log "Profile on $computer is not found"

                }

            }

        } Else {

            write-log  "$User profiles not found"

        }

    }

    Start-Sleep 7

}

#Удаление атрибутов
Function Clean-UsersAttributes {

    Param (

        [Parameter(Mandatory = $True, Position = 0)] $Path,
        [Parameter(Mandatory = $True, Position = 1)] [array]$Users,
        [Parameter(Mandatory = $True, Position = 2)] [array]$Attributes

    )

    ForEach ($User in $Users) {
        $User = [string]$User
        $ADUser = Get-ADUser -Identity $User -Properties $Attributes -ErrorAction SilentlyContinue
        $FilePath = $Path + "\" + $User + ".xml"
        ($ADUser | ConvertTo-XML –NoTypeInformation).Save($FilePath)
        write-log "Атрибуты пользователя $User выгружены в $FilePath"

        ForEach ($Attribute in $Attributes) {

            If ($Attribute -eq "MemberOf") {

                ForEach ( $Group in $ADUser.MemberOf ) {

                    if ( !( $Group -match "DV-users|MS-dax|Domain Users" ) ) {

                        write-log "Пользователь $User удаляется из группы $Group"
                        Remove-AdGroupMember -Identity $Group -members $ADUser.SamAccountName -Confirm:$false -ErrorAction SilentlyContinue
    
                    }

                }

            } Else {

                write-log "Removing $Attribute of user $User"
                Set-ADUser -Identity $User -Clear $Attribute -ErrorAction SilentlyContinue

            }

        }

        $global:Workflow["CleanUP"] = $global:Workflow["CleanUP"] | ?{ $_ -ne $User }
        $global:Workflow["MoveTo"] += @($User)

    }
  
    Start-Sleep 7
}

function Move-ADUsersToOUForDisabledUsers {

    Param (

        [Parameter(Mandatory = $TRUE, Position = 0)] $DestinationOU,
        [Parameter(Mandatory = $TRUE, Position = 1)] [array]$Users

    )

    ForEach ($User in $Users) {
        
        $User = [string]$User
        $ADUser = $null

        If ($User) {

            $ADUser = Get-ADUser -Identity $User -ErrorAction SilentlyContinue

        }

        If ( $ADUser -and ($ADUser.DistinguishedName -notmatch $DestinationOU) ) {

            $MoveToOU = [regex]::replace($ADUser.DistinguishedName,',DC=.*$','')
            $MoveToOU = [regex]::match($MoveToOU,'[^,]*$').Value
            $UserDestinationOU = $MoveToOU + "," + $DestinationOU
            write-log  "Переносим учетную запись $User в $UserDestinationOU"
            $ADUser | Move-ADObject -TargetPath $UserDestinationOU -ErrorAction SilentlyContinue
            Start-Sleep 3
            write-log "Блокируем учетную запись пользователя: $User"
            Disable-ADAccount $User -ErrorAction SilentlyContinue
                    
        }

        $global:Workflow["MoveTo"] = $global:Workflow["CleanUP"] | ?{ $_ -ne $User }

    }
    
    Start-Sleep 7
}


<#----------------------------------------------------------------------------------------------------------
Дополнительные рабочие функции скрипта
----------------------------------------------------------------------------------------------------------#>



function Check-ExportState {

    Trap {

        continue

    }
        
    Write-Host ( ( Get-date ).ToString( "T" ) + " Waiting of export completion" ) -ForegroundColor Green
    $ExportState = Get-MailboxExportRequest
    $ExportState | Get-MailboxExportRequestStatistics | Select-Object -Property SourceAlias,SourceDatabase,PercentComplete,Status,BytesTransferredPerMinute |
    Sort-Object -Property PercentComplete -Descending | Format-Table -AutoSize -Property SourceAlias,SourceDatabase,`
        @{Label = "% Complete" ; Expression = { $_.PercentComplete } },Status,`
        @{Label = "MB/Minute" ; Expression = { [regex]::Replace($_.BytesTransferredPerMinute,' \(.*','') } }
               
}

#Поиск на какие компьютеры заходил пользователь
function find-profiles {

    Param (

        [Parameter(Mandatory = $True, Position = 0)] $PathToSearch,
        [Parameter(Mandatory = $True, Position = 1)] $SearchWildcard,
        [Parameter(Mandatory = $True, Position = 2)] $SearchUser

    )

    Trap {

        continue

    }

    $Computers = @(Get-ChildItem $PathToSearch -Filter ( $SearchUser + '@' + $SearchWildcard ) | %{[regex]::Replace( [regex]::replace( $_,'.*@','' ),'\.txt$','' )})
    return $Computers

}

function Restart-FailedMailboxExportRequests {

    Trap {

        continue

    }

    $FaiedExportRequests = Get-MailboxExportRequest -Status Failed

    ForEach ( $FaiedExportRequest in $FaiedExportRequests ) {

        write-log ( "Экспорт почтового ящика " + $FaiedExportRequest.Mailbox + " в PST завершился ошибкой" )
        $FaiedExportRequest | Resume-MailboxExportRequest -WhatIf:$debugMode
        write-log  "Возобновляем экспорт почтового ящика в PST"

    }

}

#Start script
Clear-FiredUsersMain