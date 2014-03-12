Param ( 

    [switch]$debugMode = $false

 )

#Поиск пользователей в целевых OU, отключенных заданное количество дней назад
function find-users {

    $Date = (Get-Date)
    $Date = $Date.AddDays(-4)
    $Users = $NULL

    Try {

        $Users = get-aduser -SearchBase "OU=Sochi,DC=SOCHI-2014,DC=RU" -Filter * -Properties * | ? { $_.useraccountcontrol -eq "514" -and $_.whenChanged -lt $Date }

    } Catch {

        write-log $Error[0].Exception
        write-log $Error[0].InvocationInfo.Line

    }

    Try {

        $Users += get-aduser -SearchBase "OU=Moscow,DC=SOCHI-2014,DC=RU" -Filter * -Properties  * | ? { $_.useraccountcontrol -eq "514" -and $_.whenChanged -lt $Date }

    } Catch {

        write-log $Error[0].Exception
        write-log $Error[0].InvocationInfo.Line

    }

    If ( $Users ) {

        Write-Log "Users found"
        Return $Users

    } else {

        Write-Log "Users not found"
        Exit

    }

}

#Поиск на какие компьютеры заходил пользователь
function find-profiles {

    $computers = Get-ChildItem \\msk-support\ActiveUsers -Filter ( $ADUser.SamAccountName + '@wts*' ) | %{[regex]::Replace( [regex]::replace( $_,'.*@','' ),'\.txt$','' )}
    return $computers

}

#Удаление профилей пользователя на найденных компьютерах
function delete-profiles {

    write-log  ("Начинаем удаление профилей пользователя: " + $ADUser.SamAccountName)
    $computers = @()
    $computers = find-profiles
    write-log "Trying to search user profile on $computers"
    $SAM = $ADUser.SamAccountName

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

            If ( -not($debugMode) ) {

                $Profile.delete()

            }

        } else {

            write-log "Profile on $computer is not found"

        }

    }

}

#Логирование действий
function Write-Log {

    Param ( 

        [Parameter( Mandatory = $True, Position = 0, ValueFromPipeline = $True )]
        [String]$Message

     )
        
        Write-Host $Message
        $Message = ( Get-date ).ToString( "T" ) + "  " + $Message
        Out-File $PathToLogFile -InputObject $Message -Append

}

#Данная функция возвращает каталог, в котором хранится наш скрипт
function Get-ScriptDirectory {

    $Invocation = ( Get-Variable MyInvocation -Scope 1 ).Value
    Split-Path $Invocation.MyCommand.Path
    
}

#Запуск экспорта почтовых ящиков отключенных пользователей
function Export-MailboxesToPST {

    Get-MailboxExportRequest -Status Completed | Remove-MailboxExportRequest -Confirm:$False | Out-Null
    Get-MailboxExportRequest -Status Failed | Remove-MailboxExportRequest -Confirm:$False | Out-Null
    
    ForEach ($ADUser in $ADUsers) {

        $UserMailbox = $Null

        If ( $ADUser.mail ) {

            Try {

                $UserMailbox = get-mailbox $( $ADUser.mail )

            } Catch {

                write-log $Error[0].Exception
                write-log $Error[0].InvocationInfo.Line

            }

        }

        if ( $UserMailbox ) {
     
            $PSTFilePath = "\\exch-mb27\e$\PST\" + $ADUser.sAMAccountName+".pst"
            write-log  ( "Начинаем экспорт почтового ящика " + $ADUser.SamAccountName + " в файл " + $PSTFilePath )
            $ExportRequests += New-MailboxExportRequest -Mailbox $UserMailbox.Alias -FilePath $PSTFilePath -WhatIf:$debugMode

        } else {

            write-log ( "У пользователя " + $ADUser.SamAccountName + " нет почтового ящика" )
            Set-ADUser -Identity $ADUser.SamAccountName -EmailAddress $NULL
            
        }

    }

}

#Перемещение выгруженных в PST файлы почтовых ящиков на файловый ресурс архива.
function Move-ArchivePSTFiles {

    $Flag = $True

    do {
        
        $ExportRequests = Get-MailboxExportRequest
        $FaiedExportRequests = $ExportRequests  | ? { $_.Status -eq "Failed" }

        ForEach ($FaiedExportRequest in $FaiedExportRequests) {

            write-log ( "Экспорт почтового ящика " + $FaiedExportRequest.Mailbox + " в PST завершился ошибкой" )
            $FaiedExportRequest | Resume-MailboxExportRequest -WhatIf:$debugMode
            write-log  "Возобновляем экспорт почтового ящика в PST"

        }

        $ExportRequestsWithNotEnoughRights = $NULL
        $ExportRequestsWithNotEnoughRights = $ExportRequests  | ? { $_.Status -eq "Completed" -and ( Get-Item $PSTFilePath ).Length -eq 271360 }

        ForEach ($ExportRequestWithNotEnoughRights  in $ExportRequestsWithNotEnoughRights ) {

            write-log ( "Экспорт почтового ящика " + $ExportRequestWithNotEnoughRights.Mailbox + " в PST завершился ошибкой. Размер выходного файла равен 265Kb. Необходимо предоставить учетной записи, от имени которой запускается скрипт, доступ уровня Full Access к данному почтовому ящику." )
            $CurrentUser = [Environment]::UserName
            Add-MailboxPermission -Identity $mail -User $j -AccessRights 'FullAccess' | Out-Null
            Start-Sleep 15
            $PSTFilePath = $NULL
            $PSTFilePath = $ExportRequestWithNotEnoughRights.FilePath
            $UserMailbox = $ExportRequestWithNotEnoughRights.Mailbox
            write-log "Удаляем файл экспорта $PSTFilePath"
            Remove-Item -Path $PSTFilePath -Force -Confirm:$False -WhatIf:$debugMode
            write-log "Удаляем запрос экспорта"
            $ExportRequestWithNotEnoughRights | Remove-MailboxExportRequest -WhatIf:$debugMode -Confirm:$False
            write-log "Пересоздаем запрос экспорта"
            New-MailboxExportRequest -Mailbox $UserMailbox.Alias -FilePath $PSTFilePath  -WhatIf:$debugMode -Confirm:$False
        }

        $CompletedExportRequests = $NULL
        $CompletedExportRequests = $ExportRequests  | ? { $_.Status -eq "Completed" }

        ForEach ($CompletedExportRequest in $CompletedExportRequests) {

            write-log ( "Экспорт почтового ящика " + $CompletedExportRequest.Mailbox + " в PST завершен." )
            $PSTFilePath = $NULL
            $PSTFilePath = $CompletedExportRequest.FilePath
            $NewPSTFilePath = [regex]::replace( $PSTFilePath, "\\exch-mb27\e$\PST\", "\\file-06\d$\PST_backup\Удаленные с сервера\" )
            write-log "Переносим выгруженный файл в $NewPSTFilePath"
            move-item -LiteralPath $PSTFilePath -Destination $NewPSTFilePath -Force -Confirm:$False -WhatIf:$debugMode
            write-log "Отключаем почтовый ящик"
            $CompletedExportRequest | Remove-MailboxExportRequest -WhatIf:$debugMode -Confirm:$False
            Disable-Mailbox $UserMailbox.Alias -Confirm:$False -WhatIf:$debugMode

        }

        Start-Sleep 15
        $ExportRequests = $NULL
        $FaiedExportRequests = $NULL
        $ExportRequests = Get-MailboxExportRequest
        $FaiedExportRequests = $ExportRequests | ? { $_.Status -eq "Failed" }
        
        If ( $ExportRequests.Count -eq $FaiedExportRequests.Count ) {

            $Flag =$False

            ForEach ($FaiedExportRequest in $FaiedExportRequests) {

                write-log ( "Экспорт почтового ящика " + $FaiedExportRequest.Mailbox + " в PST завершился ошибкой" )

            }

        } ElseIf ( -not($ExportRequests) ) {

            $Flag =$False
            write-log "Все экспорты завершены"

        }

    } until ( $Flag )

}

#Закрытие удаленной сессии PS и завершение работы скрипта.
function Finalize-Script {

    if ( $Session )   {

        Remove-PSSession $Session

    }

    write-log  "Закрываем PSSession до сервера Exchange"
    write-log  "Скрипт завершил работу"
    exit

}

Function Block-User {

    $GroupLogPath = Join-Path ( $CurrentDirectory ) "\$( $ADUser.SamAccountName ).xml"
    ( $ADUser.MemberOf | ConvertTo-XML –NoTypeInformation ).Save( $GroupLogPath )
    write-log ( "Список групп пользователя " + $ADUser.SamAccountName + " выгружен в $GroupLogPath" )
    $ADUser | fl | Out-File -FilePath ( [regex]::Replace( $GroupLogPath,'xml$','txt' ) )
    write-log "Выгружены атрибуты учетной записи пользователя $Username в .txt файл"

    foreach ( $Group in $ADUser.MemberOf ) {

        if ( !( $Group -match "DV-users|MS-dax|Domain Users" ) ) {

            write-log ( "Пользователь " + $ADUser.SamAccountName + " удаляется из группы $Group")
            get-adgroup $Group | Remove-AdGroupMember -member $ADUser.DistinguishedName -Confirm:$false -WhatIf:$debugMode
    
        }

    }
  
    write-log ( "Блокируем учетную запись пользователя " + $ADUser.SamAccountName )
    Disable-ADAccount $ADUser -WhatIf:$debugMode
    write-log ( "Удаляем информацию о телефонах пользователя " + $ADUser.SamAccountName )
    $AdUser | Set-ADUser -MobilePhone $Null -OfficePhone $Null -HomePhone $Null -WhatIf:$debugMode
  
    If ( $AdUser.DistinguishedName -match "OU=Sochi" ) {
    
        $MoveToOU = "Sochi"
        
    } else {
    
        $MoveToOU = "Moscow"

    }
  
    $DestinationOU = "OU="+$MoveToOU+",OU=Disabled Users,DC=SOCHI-2014,DC=RU"
    write-log  "Переносим учетную запись $Username в $DestinationOU"
    Move-ADObject $ADUser -TargetPath $DestinationOU -WhatIf:$debugMode
  
}

$CurrentDirectory = Get-ScriptDirectory
$ErrorActionPreferenceState = $ErrorActionPreference
$ErrorActionPreference = "Stop"
import-module activedirectory
$PathToLogFile = $CurrentDirectory + "\Clear-DisabledUsers_" + (Get-Date -format yyyy-MM-dd_HH-mm-ss) + ".txt"
#Создаем удаленную сессию для импорта командлетов с сервера Exchange
$Session = $null
$ExchServer = 'exch-cas02-n1'
#Пытаемся создать удаленную PS-сессию до сервера
write-log "Попытка создать PSSession до сервера $ExchServer"
$Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri ( "http://" + $ExchServer + "/Powershell/" ) -authentication Kerberos -ErrorAction SilentlyContinue

if ( $Session ) {

    write-log  "Удаленная PS-сессия создана";$Session

} else {

    write-log  "Удаленная PS-сессия не создана"
    exit

}

Import-PSSession $Session -AllowClobber | Out-Null
$ADUsers = $NULL
$ADUsers = Find-Users
$ADUsers | ft -AutoSize SamAccountName,Description,Title
Start-Sleep 10
Export-MailboxesToPST
Move-ArchivePSTFiles

ForEach ( $ADUser in $ADUsers ) {

    delete-profiles
    Block-User
        
}

$ErrorActionPreference = $ErrorActionPreferenceState
Finalize-Script